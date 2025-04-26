import os
import json
import base64

import pandas as pd
import streamlit as st
import gspread
from google.oauth2 import service_account
import holidays  # make sure holidays is in your requirements.txt

from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1-Jkuwl9e1FBY6le08_KA3k7v9J3kfDvSYxw7oOJDtPQ/"
    "edit?gid=1789939189"
)
RAW_TAB_NAME = "raw_orders"
LOCAL_TZ = "America/Toronto"
# -----------------------------------------------------------------------------


def get_worksheet():
    """Fetch the raw_orders sheet via a base64‐encoded service‐account key."""
    b64 = os.getenv("GOOGLE_SERVICE_KEY_B64")
    if not b64:
        st.error("🚨 Missing GOOGLE_SERVICE_KEY_B64 in Secrets")
        st.stop()

    info = json.loads(base64.b64decode(b64).decode("utf-8"))
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.readonly",
        ],
    )
    client = gspread.authorize(creds)
    return client.open_by_url(SHEET_URL).worksheet(RAW_TAB_NAME)


def load_data():
    """Pull the CSV into a DataFrame, cast types, compute Outstanding Qty & Bucket."""
    ws = get_worksheet()
    df = pd.DataFrame(ws.get_all_records())

    # 1) Cast to proper dtypes
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0) - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # 2) Business‐day “tomorrow” in ON (skip Sat, Sun, ON stat holidays)
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)

    def next_open_day(d: pd.Timestamp) -> pd.Timestamp:
        candidate = d + pd.Timedelta(days=1)
        while candidate.weekday() >= 5 or candidate in ca_holidays:
            candidate += pd.Timedelta(days=1)
        return candidate

    tomorrow = next_open_day(today)

    # 3) Bucket logic
    conds = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0) & (df["Quantity Fulfilled/Received"] > 0),
    ]
    labels = ["Overdue", "Due Tomorrow", "Partially Shipped"]

    df["Bucket"] = pd.NA
    for c, lbl in zip(conds, labels):
        df.loc[c, "Bucket"] = lbl

    return df


def main():
    st.set_page_config(page_title="Orderdesk Dashboard", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")

    df = load_data()

    # ─── KPIs ────────────────────────────────────────────────────────────────
    k1, k2, k3 = st.columns(3)
    k1.metric("Overdue", f"{int((df.Bucket == 'Overdue').sum())}")
    k2.metric("Due Tomorrow", f"{int((df.Bucket == 'Due Tomorrow').sum())}")
    k3.metric("Partially Shipped", f"{int((df.Bucket == 'Partially Shipped').sum())}")

    # ─── Sidebar Filters ──────────────────────────────────────────────────────
    with st.sidebar:
        st.header("Filters")
        custs = st.multiselect("Customer", sorted(df["Name"].unique()), default=None)
        rush   = st.checkbox("Rush orders only")
        if custs:
            df = df[df["Name"].isin(custs)]
        if rush and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # ─── Tabs & Grids ──────────────────────────────────────────────────────────
    for bucket_name in ["Overdue", "Due Tomorrow", "Partially Shipped"]:
        tab = st.tab(bucket_name) if False else st  # workaround for this example
    # Actually create tabs:
    tab_overdue, tab_due, tab_partial = st.tabs(
        ["Overdue", "Due Tomorrow", "Partially Shipped"]
    )
    mapping = {
        "Overdue": tab_overdue,
        "Due Tomorrow": tab_due,
        "Partially Shipped": tab_partial,
    }

    for bname, tab in mapping.items():
        sub = df[df["Bucket"] == bname]
        if sub.empty:
            tab.info(f"No {bname.lower()} orders 🎉")
            continue

        # — build summary (one row per order) —
        summary = (
            sub.groupby(["Document Number", "Name", "Ship Date"], as_index=False)
               .agg({
                   "Outstanding Qty": "sum",
                   "Quantity Fulfilled/Received": "sum",
               })
               .rename(columns={
                   "Document Number": "Order #",
                   "Name": "Customer",
                   "Outstanding Qty": "Outstanding",
                   "Quantity Fulfilled/Received": "Shipped",
               })
               .sort_values("Ship Date")
        )

        # — flag BOM items on any overdue order —
        def flag_bom(order_no):
            block = sub[sub["Document Number"] == order_no]
            if "Item Type" not in block:
                return ""
            mask = (
                (block["Item Type"] == "Assembly/Bill of Materials")
                & (block["Outstanding Qty"] > 0)
            )
            return "⚠️" if mask.any() else ""

        summary["BOM Flag"] = summary["Order #"].apply(flag_bom)

        # — pack detail rows into a hidden column —
        detail_map = {}
        for ord_no in summary["Order #"]:
            detail = sub[sub["Document Number"] == ord_no][
                ["Item", "Item Type", "Quantity", "Quantity Fulfilled/Received", "Outstanding Qty", "Memo"]
            ].rename(columns={
                "Quantity": "Qty Ordered",
                "Quantity Fulfilled/Received": "Qty Shipped",
            })
            detail_map[ord_no] = detail.to_dict("records")
        summary["lineItems"] = summary["Order #"].map(detail_map)

        # — configure the Ag-Grid —
        gb = GridOptionsBuilder.from_dataframe(summary)
        gb.configure_default_column(flex=1, min_column_width=80, sortable=True)
        gb.configure_column("Ship Date", 
                            type=["dateColumnFilter","customDateTimeFormat"],
                            custom_format_string="yyyy-MM-dd")
        gb.configure_column("lineItems", hide=True)

        gb.configure_grid_options(
            masterDetail=True,
            detailCellRenderer='agDetailCellRenderer',
            detailCellRendererParams={
                "detailGridOptions": {
                    "defaultColDef": {"sortable": True, "flex": 1, "minWidth": 80},
                    "columnDefs": [
                        {"field": "Item"},
                        {"field": "Item Type"},
                        {"field": "Qty Ordered"},
                        {"field": "Qty Shipped"},
                        {"field": "Outstanding Qty"},
                        {"field": "Memo", "flex": 1},
                    ],
                },
                "getDetailRowData": JsCode("""
                    function(params) {
                      params.successCallback(params.data.lineItems);
                    }
                """),
            },
        )

        tab.markdown("**▶ Click the arrow at left to expand line‐items**")
        AgGrid(
            summary,
            gridOptions=gb.build(),
            theme="streamlit",
            enable_enterprise_modules=False,
            fit_columns_on_grid_load=True,
        )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
