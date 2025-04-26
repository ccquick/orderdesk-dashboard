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
SHEET_URL = "https://docs.google.com/spreadsheets/d/1-Jkuwl9e1FBY6le08_KA3k7v9J3kfDvSYxw7oOJDtPQ/edit?gid=1789939189"  # 👈 your sheet link
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

    # cast types
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce").dt.tz_localize(None)
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0)
        - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # compute today/tomorrow
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    tomorrow = today + pd.Timedelta(days=1)

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
    # metrics
    kpi_cols = st.columns(3)
    for col, label in zip(kpi_cols, ["Overdue", "Due Tomorrow", "Partially Shipped"]):
        count = int((df["Bucket"] == label).sum())
        col.metric(label, f"{count}")

    # sidebar filters
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect(
            "Customer", sorted(df["Name"].unique()), default=None
        )
        rush_only = st.checkbox("Rush orders only")
        if customers:
            df = df[df["Name"].isin(customers)]
        if rush_only and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # tabs
    tabs = st.tabs(["Overdue", "Due Tomorrow", "Partially Shipped"])
    tab_map = dict(zip(["Overdue", "Due Tomorrow", "Partially Shipped"], tabs))

    for bucket, tab in tab_map.items():
        sub = df[df["Bucket"] == bucket]
        with tab:
            if sub.empty:
                st.info(f"No {bucket.lower()} orders 🎉")
                continue

            # summary table
            summary = (
                sub.groupby(["Document Number", "Name", "Ship Date"], as_index=False)
                .agg({
                    "Outstanding Qty": "sum",
                    "Quantity Fulfilled/Received": "sum",
                })
            )
            st.dataframe(
                summary.sort_values("Ship Date"),
                use_container_width=True,
            )

            # line‐item expanders
            for _, order in summary.iterrows():
                order_no = order["Document Number"]
                cust = order["Name"]
                ship_dt = pd.to_datetime(order["Ship Date"]).date()
                label = f"Order {order_no} — {cust} ({ship_dt})"
                with st.expander(label, expanded=False):
                    lines = sub[sub["Document Number"] == order_no][
                        ["Quantity", "Item", "Memo"]
                    ]
                    st.dataframe(lines, use_container_width=True)

    st.caption(
        "Data auto‑refreshes hourly from NetSuite saved search ➜ Google Sheet ➜ Streamlit"
    )

if __name__ == "__main__":
    main()
