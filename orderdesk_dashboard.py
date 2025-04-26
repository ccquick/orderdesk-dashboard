import os
import json
import base64

import pandas as pd
import streamlit as st
import gspread
from google.oauth2 import service_account
import holidays

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

def load_data() -> pd.DataFrame:
    ws = get_worksheet()
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    # cast
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0)
        - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # business‐day “tomorrow” logic
    on_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    def next_open(d):
        n = d + pd.Timedelta(days=1)
        while n.weekday() >= 5 or n in on_holidays:
            n += pd.Timedelta(days=1)
        return n
    tomorrow = next_open(today)

    # bucket
    conds = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0)
          & (df["Quantity Fulfilled/Received"] > 0),
    ]
    labels = ["Overdue", "Due Tomorrow", "Partially Shipped"]
    df["Bucket"] = pd.NA
    for c, l in zip(conds, labels):
        df.loc[c, "Bucket"] = l

    return df

def main():
    st.set_page_config(page_title="Orderdesk Dashboard", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")

    df = load_data()

    # ─── KPIs ─────────────────
    k1, k2, k3 = st.columns(3)
    for col, lab in zip((k1, k2, k3), ["Overdue", "Due Tomorrow", "Partially Shipped"]):
        col.metric(lab, int((df["Bucket"] == lab).sum()))

    # ─── FILTERS ─────────────
    with st.sidebar:
        st.header("Filters")
        custs = st.multiselect("Customer", sorted(df["Name"].unique()), default=None)
        rush  = st.checkbox("Rush orders only")
        if custs:
            df = df[df["Name"].isin(custs)]
        if rush and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize()=="Yes"]

    # ─── TABS ────────────────
    # ← pass the list *positionally*, not as `labels=…`
    tabs = st.tabs(["Overdue", "Due Tomorrow", "Partially Shipped"])

    for bucket, tab in zip(["Overdue","Due Tomorrow","Partially Shipped"], tabs):
        sub = df[df["Bucket"] == bucket]
        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # build summary + detail payload
        summary = (
            sub
            .groupby(["Document Number","Name","Ship Date"], as_index=False)
            .agg({
                "Outstanding Qty":"sum",
                "Quantity Fulfilled/Received":"sum",
            })
            .rename(columns={
                "Document Number":"Order #",
                "Name":"Customer",
                "Outstanding Qty":"Outstanding",
                "Quantity Fulfilled/Received":"Shipped",
            })
            .sort_values("Ship Date")
        )
        # BOM flag
        def has_bom(o):
            detail = sub[sub["Document Number"]==o]
            return "⚠️" if (
                (detail["Item Type"]=="Assembly/Bill of Materials")
                & (detail["Outstanding Qty"]>0)
            ).any() else ""
        summary["BOM Flag"] = summary["Order #"].apply(has_bom)

        # attach line‐items
        lookup = {}
        for o in summary["Order #"]:
            dt = sub[sub["Document Number"]==o]
            lookup[o] = dt[[
                "Item","Item Type",
                "Quantity","Quantity Fulfilled/Received",
                "Outstanding Qty","Memo"
            ]].rename(columns={
                "Quantity":"Qty Ordered",
                "Quantity Fulfilled/Received":"Qty Shipped",
            }).to_dict("records")
        summary["items"] = summary["Order #"].map(lookup)

        # master/detail via AgGrid
        detail_renderer = JsCode("""
        function(params) {
          const eGui = document.createElement('div');
          new agGrid.Grid(eGui, {
            columnDefs:[
              {field:'Item'},{field:'Item Type'},{field:'Qty Ordered'},
              {field:'Qty Shipped'},{field:'Outstanding Qty'},{field:'Memo',flex:1},
            ],
            defaultColDef:{flex:1,minWidth:100,sortable:true},
            rowData: params.data.items,
            domLayout:'autoHeight'
          });
          return eGui;
        }
        """)
        gb = GridOptionsBuilder.from_dataframe(summary)
        gb.configure_default_column(flex=1,min_column_width=80,sortable=True)
        gb.configure_column("Ship Date",
                             type=["dateColumnFilter","customDateTimeFormat"],
                             custom_format_string="yyyy-MM-dd")
        gb.configure_column("items", hide=True)
        gb.configure_grid_options(
            masterDetail=True,
            detailCellRenderer=detail_renderer,
            detailRowAutoHeight=True,
            getRowNodeId="data=>data['Order #']"
        )
        gb.configure_pagination(paginationAutoPageSize=True)
        grid_opts = gb.build()

        tab.markdown("**▶ Click the arrow to expand an order’s line-items**")
        AgGrid(
            summary,
            gridOptions=grid_opts,
            theme="streamlit",
            enable_enterprise_modules=False,
            fit_columns_on_grid_load=True,
        )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")

if __name__=="__main__":
    main()
