import os
import json
import base64

import pandas as pd
import streamlit as st
import gspread
from google.oauth2 import service_account
import holidays                   # pip install holidays
from st_aggrid import AgGrid      # pip install st-aggrid
from st_aggrid import GridOptionsBuilder, JsCode

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

    # If your new column came in as "Type", rename it to "Item Type"
    if "Type" in df.columns and "Item Type" not in df.columns:
        df.rename(columns={"Type": "Item Type"}, inplace=True)

    # cast types
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0) - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # compute business‐day “tomorrow” (skip weekends + ON stat holidays)
    ca_hols = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)

    def next_open(d):
        nd = d + pd.Timedelta(days=1)
        while nd.weekday() >= 5 or nd in ca_hols:
            nd += pd.Timedelta(days=1)
        return nd

    tomorrow = next_open(today)

    # bucket logic
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

    # ─── KPIs ───────────────────────────────────────────────────
    c1, c2, c3 = st.columns(3)
    for col, lbl in zip((c1, c2, c3), ["Overdue", "Due Tomorrow", "Partially Shipped"]):
        col.metric(lbl, f"{int((df['Bucket']==lbl).sum())}")

    # ─── SIDEBAR FILTERS ───────────────────────────────────────
    with st.sidebar:
        st.header("Filters")
        custs = st.multiselect("Customer", sorted(df["Name"].unique()), default=[])
        rush  = st.checkbox("Rush orders only")
    if custs:
        df = df[df["Name"].isin(custs)]
    if rush and "Rush Order" in df.columns:
        df = df[df["Rush Order"].str.capitalize()=="Yes"]

    # ─── TABS ──────────────────────────────────────────────────
    tabs = st.tabs(["Overdue","Due Tomorrow","Partially Shipped"])
    for bucket, tab in zip(["Overdue","Due Tomorrow","Partially Shipped"], tabs):
        with tab:
            sub = df[df["Bucket"]==bucket]
            if sub.empty:
                st.info(f"No {bucket.lower()} orders 🎉")
                continue

            # build summary (one‐row‐per‐order)
            summary = (
                sub.groupby(
                    ["Document Number","Name","Ship Date"], as_index=False
                )
                .agg({
                    "Outstanding Qty":"sum",
                    "Quantity Fulfilled/Received":"sum"
                })
                .rename(columns={
                    "Document Number":"Order #",
                    "Name":"Customer",
                    "Ship Date":"Ship Date",
                    "Outstanding Qty":"Outstanding",
                    "Quantity Fulfilled/Received":"Shipped",
                })
                .sort_values("Ship Date")
            )

            # build mapping order→list-of-detail‐dicts
            detail_map = {}
            for ord_no in summary["Order #"]:
                d = sub[sub["Document Number"]==ord_no]
                detail_map[ord_no] = d[[
                    "Item",
                    "Item Type",              # now present
                    "Quantity",
                    "Quantity Fulfilled/Received",
                    "Outstanding Qty",
                    "Memo",
                ]].rename(columns={
                    "Quantity":"Qty Ordered",
                    "Quantity Fulfilled/Received":"Qty Shipped",
                    "Outstanding Qty":"Outstanding",
                }).to_dict("records")

            summary["details"] = summary["Order #"].map(detail_map)

            # configure AgGrid master/detail
            gb = GridOptionsBuilder.from_dataframe(summary)
            gb.configure_column("details", hide=True)
            gb.configure_grid_options(
                masterDetail=True,
                detailRowAutoHeight=True,
                detailCellRendererParams={
                    "detailGridOptions": {
                        "columnDefs":[
                            {"field":"Item", "headerName":"Item","flex":1},
                            {"field":"Item Type","headerName":"Item Type","flex":1},
                            {"field":"Qty Ordered","headerName":"Qty Ordered","flex":1},
                            {"field":"Qty Shipped","headerName":"Qty Shipped","flex":1},
                            {"field":"Outstanding","headerName":"Outstanding","flex":1},
                            {"field":"Memo","headerName":"Memo","flex":2},
                        ],
                        "defaultColDef":{"sortable":True,"resizable":True}
                    },
                    "getDetailRowData": JsCode("""
                        function(params) {
                          params.successCallback(params.data.details);
                        }
                    """),
                },
            )
            grid_opts = gb.build()

            AgGrid(
                summary,
                gridOptions=grid_opts,
                enable_enterprise_modules=True,  # masterDetail needs enterprise
                allow_unsafe_jscode=True,        # for our JS callback
                fit_columns_on_grid_load=True,
                theme="material-dark",
                key=bucket
            )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__=="__main__":
    main()
