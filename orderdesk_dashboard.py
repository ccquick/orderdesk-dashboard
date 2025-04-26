import os
import json
import base64
import pandas as pd
import streamlit as st
import gspread
from google.oauth2 import service_account
import holidays  # pip install holidays
from st_aggrid import AgGrid
from st_aggrid.grid_options_builder import GridOptionsBuilder
from st_aggrid.shared import JsCode

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
    b64 = os.getenv("GOOGLE_SERVICE_KEY_B64")
    if not b64:
        st.error("🚨 Missing GOOGLE_SERVICE_KEY_B64 in Secrets")
        st.stop()

    info = json.loads(base64.b64decode(b64).decode())
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
    ws = get_worksheet()
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    # if your new column is named "Type", rename it
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # 1) cast types & compute Outstanding
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0) - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # 2) business-day “tomorrow” in ON (skip weekends + holidays)
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)

    def next_open(d: pd.Timestamp) -> pd.Timestamp:
        n = d + pd.Timedelta(days=1)
        while n.weekday() >= 5 or n in ca_holidays:
            n += pd.Timedelta(days=1)
        return n

    tomorrow = next_open(today)

    # 3) bucket each row
    conditions = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0) & (df["Quantity Fulfilled/Received"] > 0),
    ]
    choices = ["Overdue", "Due Tomorrow", "Partially Shipped"]
    df["Bucket"] = pd.Series(pd.NA, index=df.index)
    for cond, lbl in zip(conditions, choices):
        df.loc[cond, "Bucket"] = lbl

    return df


def main():
    st.set_page_config(page_title="Orderdesk Dashboard", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")

    df = load_data()

    # ─── KPIs & FILTERS ────────────────────────────────────────────────────────
    c1, c2, c3 = st.columns(3)
    for col, lbl in zip((c1, c2, c3), ["Overdue", "Due Tomorrow", "Partially Shipped"]):
        col.metric(lbl, int((df["Bucket"] == lbl).sum()))

    with st.sidebar:
        st.header("Filters")
        custs = sorted(df["Name"].unique())
        sel_cust = st.multiselect("Customer", custs)
        rush_only = st.checkbox("Rush orders only")
    if sel_cust:
        df = df[df["Name"].isin(sel_cust)]
    if rush_only and "Rush Order" in df.columns:
        df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # ─── TABS ─────────────────────────────────────────────────────────────────
    tabs = st.tabs(["Overdue", "Due Tomorrow", "Partially Shipped"])
    for bucket, tab in zip(["Overdue", "Due Tomorrow", "Partially Shipped"], tabs):
        sub = df[df["Bucket"] == bucket]
        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # build summary table
        summary = (
            sub.groupby(["Document Number", "Name", "Ship Date"], as_index=False)
            .agg({
                "Outstanding Qty": "sum",
                "Quantity Fulfilled/Received": "sum",
            })
            .rename(columns={
                "Document Number": "Order #",
                "Name": "Customer",
                "Ship Date": "Ship Date",
                "Outstanding Qty": "Outstanding",
                "Quantity Fulfilled/Received": "Shipped",
            })
            .sort_values("Ship Date")
        )

        # ――――――――――――――――――――――――――――――――――――――
        # **FIX**: convert datetime → string so AgGrid accepts it
        summary["Ship Date"] = summary["Ship Date"].astype(str)
        # ――――――――――――――――――――――――――――――――――――――

        # flag any order with outstanding BOM lines
        bom_orders = set(
            sub.loc[
                (sub["Item Type"] == "Assembly/Bill of Materials") &
                (sub["Outstanding Qty"] > 0),
                "Document Number"
            ]
        )
        summary["BOM Flag"] = summary["Order #"].apply(lambda o: "⚠️" if o in bom_orders else "")

        # assemble master/detail payload
        records = []
        for _, r in summary.iterrows():
            rec = r.to_dict()
            details = sub[sub["Document Number"] == r["Order #"]]
            rec["lineItems"] = details[
                ["Item", "Item Type", "Quantity", "Quantity Fulfilled/Received", "Outstanding Qty", "Memo"]
            ].rename(columns={
                "Quantity": "Qty Ordered",
                "Quantity Fulfilled/Received": "Qty Shipped",
                "Outstanding Qty": "Outstanding",
            }).to_dict("records")
            records.append(rec)

        # build AgGrid options
        gb = GridOptionsBuilder.from_dataframe(summary.drop(columns=["BOM Flag"]))
        gb.configure_column("BOM Flag", header_name="BOM", width=80)
        gb.configure_grid_options(
            masterDetail=True,
            detailCellRendererParams={
                "detailGridOptions": {
                    "columnDefs": [
                        {"field": "Item", "minWidth": 200},
                        {"field": "Item Type", "minWidth": 180},
                        {"field": "Qty Ordered"},
                        {"field": "Qty Shipped"},
                        {"field": "Outstanding"},
                        {"field": "Memo", "minWidth": 250},
                    ]
                },
                "getDetailRowData": JsCode(
                    "function(params) { params.successCallback(params.data.lineItems); }"
                ),
            },
        )
        grid_opts = gb.build()

        tab.markdown("▶️ Click the arrow on the left to expand line-items")
        AgGrid(
            records,
            gridOptions=grid_opts,
            allow_unsafe_jscode=True,
            fit_columns_on_grid_load=True,
        )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
