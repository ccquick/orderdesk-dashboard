import os
import json
import base64

import pandas as pd
import streamlit as st
import gspread
from google.oauth2 import service_account
import holidays  # pip install holidays
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

# -----------------------------------------------------------------------------
# CONFIGURATION (edit just this block)
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
    ws = get_worksheet()
    df = pd.DataFrame(ws.get_all_records())

    # cast types
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0) - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # business-day “tomorrow” (skip Sat/Sun and ON stat holidays)
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)

    def next_open_day(d):
        nd = d + pd.Timedelta(days=1)
        while nd.weekday() >= 5 or nd in ca_holidays:
            nd += pd.Timedelta(days=1)
        return nd

    tomorrow = next_open_day(today)

    # bucket flags
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

    # ─── KPIs ──────────────────────────────────────────────────────────────
    k1, k2, k3 = st.columns(3)
    for col, label in zip((k1, k2, k3), ["Overdue", "Due Tomorrow", "Partially Shipped"]):
        col.metric(label, int((df["Bucket"] == label).sum()))

    # ─── SIDEBAR FILTERS ───────────────────────────────────────────────────
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect("Customer", sorted(df["Name"].unique()), default=[])
        rush_only = st.checkbox("Rush orders only")
    if customers:
        df = df[df["Name"].isin(customers)]
    if rush_only and "Rush Order" in df.columns:
        df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # ─── TABS ───────────────────────────────────────────────────────────────
    tabs = st.tabs(["Overdue", "Due Tomorrow", "Partially Shipped"])
    for bucket, tab in zip(["Overdue","Due Tomorrow","Partially Shipped"], tabs):
        with tab:
            sub = df[df["Bucket"] == bucket]
            if sub.empty:
                st.info(f"No {bucket.lower()} orders 🎉")
                continue

            # ─ Summary: one row per order ────────────────────────────────
            summary = (
                sub.groupby(
                    ["Document Number", "Name", "Ship Date"], as_index=False
                )
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

            # attach line-items as a list for each order
            detail_map = {}
            for order in summary["Order #"]:
                rows = sub[sub["Document Number"] == order]
                detail_map[order] = rows[[
                    "Item",
                    "Item Type",
                    "Quantity",
                    "Quantity Fulfilled/Received",
                    "Outstanding Qty",
                    "Memo"
                ]].rename(columns={
                    "Quantity": "Qty Ordered",
                    "Quantity Fulfilled/Received": "Qty Shipped",
                }).to_dict("records")

            summary["details"] = summary["Order #"].map(detail_map)

            # ─ Build Ag-Grid master/detail ───────────────────────────────
            gb = GridOptionsBuilder.from_dataframe(summary)
            gb.configure_column("details", hide=True)
            gb.configure_grid_options(
                masterDetail=True,
                detailRowAutoHeight=True,
                detailCellRendererParams={
                    "detailGridOptions": {
                        "columnDefs": [
                            {"field":"Item","headerName":"Item","flex":1},
                            {"field":"Item Type","headerName":"Item Type","flex":1},
                            {"field":"Qty Ordered","headerName":"Qty Ordered","flex":1},
                            {"field":"Qty Shipped","headerName":"Qty Shipped","flex":1},
                            {"field":"Outstanding","headerName":"Outstanding","flex":1},
                            {"field":"Memo","headerName":"Memo","flex":2},
                        ],
                        "defaultColDef": {"sortable":True,"resizable":True}
                    },
                    "getDetailRowData": JsCode("""
                        function(params) {
                          params.successCallback(params.data.details);
                        }
                    """),
                }
            )
            gridOptions = gb.build()

            AgGrid(
                summary,
                gridOptions=gridOptions,
                enable_enterprise_modules=True,   # masterDetail is an enterprise feature
                allow_unsafe_jscode=True,         # to allow our JS callback
                fit_columns_on_grid_load=True,
                theme="material-dark",
                key=f"aggrid_{bucket}"
            )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
