import os
import json
import base64
import pandas as pd
import streamlit as st
import gspread
from google.oauth2 import service_account
import holidays  # pip install holidays
from st_aggrid import AgGrid, GridOptionsBuilder

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
    sh = client.open_by_url(SHEET_URL)
    return sh.worksheet(RAW_TAB_NAME)


def load_data():
    ws = get_worksheet()
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    # cast types
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0)
        - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # business-day "tomorrow"
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    def next_open_day(d):
        c = d + pd.Timedelta(days=1)
        while c.weekday() >= 5 or c in ca_holidays:
            c += pd.Timedelta(days=1)
        return c
    tomorrow = next_open_day(today)

    # bucket
    conditions = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0) & (df["Quantity Fulfilled/Received"] > 0),
    ]
    choices = ["Overdue", "Due Tomorrow", "Partially Shipped"]
    df["Bucket"] = pd.Series(pd.NA, index=df.index)
    for cond, label in zip(conditions, choices):
        df.loc[cond, "Bucket"] = label

    return df


def main():
    st.set_page_config(page_title="Orderdesk Dashboard", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")

    df = load_data()

    # KPIs
    k1, k2, k3 = st.columns(3)
    k1.metric("Overdue", int((df["Bucket"] == "Overdue").sum()))
    k2.metric("Due Tomorrow", int((df["Bucket"] == "Due Tomorrow").sum()))
    k3.metric("Partially Shipped", int((df["Bucket"] == "Partially Shipped").sum()))

    # filters
    with st.sidebar:
        st.header("Filters")
        cust = st.multiselect("Customer", sorted(df["Name"].unique()), default=None)
        rush = st.checkbox("Rush orders only")
        if cust:
            df = df[df["Name"].isin(cust)]
        if rush and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # tabs
    tab_over, tab_due, tab_part = st.tabs([
        "Overdue", "Due Tomorrow", "Partially Shipped"
    ])
    for bucket, tab in zip(["Overdue","Due Tomorrow","Partially Shipped"],
                           [tab_over,tab_due,tab_part]):
        sub = df[df["Bucket"] == bucket]
        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # AG-Grid grouping
        gb = GridOptionsBuilder.from_dataframe(sub)
        # group by Document Number
        gb.configure_column("Document Number", rowGroup=True, hide=True)
        # show first Ship Date & Name per group
        gb.configure_column("Ship Date", aggFunc="first", headerName="Ship Date")
        gb.configure_column("Name", aggFunc="first", headerName="Customer")
        # aggregate qty columns
        gb.configure_column("Outstanding Qty", aggFunc="sum", headerName="Outstanding")
        gb.configure_column(
            "Quantity Fulfilled/Received", aggFunc="sum", headerName="Shipped"
        )
        # detail columns
        gb.configure_column("Item", headerName="Line Item")
        gb.configure_column("Memo", headerName="Memo")

        gb.configure_grid_options(
            groupDefaultExpanded=0,  # start collapsed
            autoGroupColumnDef={
                "headerName":"Order #",
                "cellRendererParams": {"suppressCount": True}
            },
            defaultColDef={"flex":1, "sortable":True, "filter":True},
        )
        grid_opts = gb.build()

        tab.markdown(
            "Click the ▸ to expand each order's line-items"
        )
        AgGrid(
            sub,
            gridOptions=grid_opts,
            theme="material-dark",
            fit_columns_on_grid_load=True,
            enable_enterprise_modules=False,
        )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")

if __name__ == "__main__":
    main()
