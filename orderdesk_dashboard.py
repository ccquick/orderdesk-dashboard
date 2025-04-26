import os
import json
import base64
import pandas as pd
import streamlit as st
import gspread
from google.oauth2 import service_account

# -----------------------------------------------------------------------------
# CONFIGURATION (edit just this block)
# -----------------------------------------------------------------------------
SHEET_URL = "https://docs.google.com/spreadsheets/d/1-Jkuwl9e1FBY6le08_KA3k7v9J3kfDvSYxw7oOJDtPQ/edit?gid=1789939189"  # 👈 paste your sheet link
RAW_TAB_NAME = "raw_orders"
LOCAL_TZ = "America/Toronto"

# -----------------------------------------------------------------------------
# Helper functions
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
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0) - df["Quantity Fulfilled/Received"].fillna(0)
    )
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    tomorrow = today + pd.Timedelta(days=1)
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

    # KPI metrics
    kpi_cols = st.columns(3)
    for col, label in zip(kpi_cols, ["Overdue", "Due Tomorrow", "Partially Shipped"]):
        count = int((df["Bucket"] == label).sum())
        col.metric(label, f"{count}")

    # Sidebar Filters
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect(
            "Customer", sorted(df["Name"].unique().tolist()), default=None
        )
        rush_only = st.checkbox("Rush orders only")
        if customers:
            df = df[df["Name"].isin(customers)]
        if rush_only and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # Tabs per bucket
    tab_overdue, tab_due, tab_part = st.tabs(["Overdue", "Due Tomorrow", "Partially Shipped"])
    tab_map = {"Overdue": tab_overdue, "Due Tomorrow": tab_due, "Partially Shipped": tab_part}

    for bucket, tab in tab_map.items():
        sub = df[df["Bucket"] == bucket]
        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # Aggregate to one row per order
        agg = (
            sub.groupby(["Document Number", "Name", "Ship Date"], as_index=False)
            .agg({
                "Outstanding Qty": "sum",
                "Quantity Fulfilled/Received": "sum",
            })
        )

        # For each order, show an expander with line-level details
        for _, order in agg.sort_values("Ship Date").iterrows():
            header = (
                f"📄 {order['Document Number']}  •  {order['Name']}  •  "
                f"{order['Ship Date'].date()}  •  Outstanding: {order['Outstanding Qty']}"
            )
            with tab.expander(header, expanded=False):
                lines = sub[sub['Document Number'] == order['Document Number']][
                    ['Quantity', 'Item', 'Memo']
                ]
                tab.dataframe(lines.reset_index(drop=True), use_container_width=True)

    st.caption(
        "Data auto-refreshes hourly from NetSuite saved search ➜ Google Sheet ➜ Streamlit"
    )

if __name__ == "__main__":
    main()
