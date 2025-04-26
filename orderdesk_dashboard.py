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
# 1. Share your Google Sheet (the one the Apps Script fills) with the
#    service account email in your JSON key.  Example: xyz@project.iam.gserviceaccount.com
# 2. Add the path of that JSON key file to your Streamlit secrets or
#    set an environment variable GOOGLE_SERVICE_KEY.
# 3. Set SHEET_URL below to your sheet URL.
# -----------------------------------------------------------------------------

SHEET_URL = "https://docs.google.com/spreadsheets/d/1-Jkuwl9e1FBY6le08_KA3k7v9J3kfDvSYxw7oOJDtPQ/edit?gid=1789939189#gid=1789939189"  # 👈 paste your sheet link
RAW_TAB_NAME = "raw_orders"
LOCAL_TZ = "America/Toronto"

# -----------------------------------------------------------------------------
# Helper functions
# -----------------------------------------------------------------------------

def get_worksheet():
    # 1) Grab the Base64-encoded JSON from Secrets
    b64 = os.getenv("GOOGLE_SERVICE_KEY_B64")
    if not b64:
        st.error("🚨 Missing GOOGLE_SERVICE_KEY_B64 in Secrets")
        st.stop()

    # 2) Decode it back to a JSON dict
    info = json.loads(base64.b64decode(b64).decode("utf-8"))

    # 3) Build credentials from that dict
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.readonly",
        ],
    )

    # 4) Authorize & open the sheet
    client = gspread.authorize(creds)
    sh = client.open_by_url(SHEET_URL)
    return sh.worksheet(RAW_TAB_NAME)

def load_data():
    # 1) get the worksheet and pull it into a DataFrame
    ws = get_worksheet()
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    # 2) cast types
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0)
        - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # 3) normalize today/tomorrow as tz-naive to match Ship Date dtype
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    tomorrow = today + pd.Timedelta(days=1)

    # 4) bucket logic
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

    kpi_cols = st.columns(3)
    for col, label in zip(kpi_cols, ["Overdue", "Due Tomorrow", "Partially Shipped"]):
        count = int((df["Bucket"] == label).sum())
        col.metric(label, f"{count}")

    # Filters
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
    tab_overdue, tab_due_tomorrow, tab_partial = st.tabs(
        ["Overdue", "Due Tomorrow", "Partially Shipped"]
    )

    tab_map = {
        "Overdue": tab_overdue,
        "Due Tomorrow": tab_due_tomorrow,
        "Partially Shipped": tab_partial,
    }

    for bucket, tab in tab_map.items():
        sub = df[df["Bucket"] == bucket]
        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
        else:
            tab.dataframe(
                sub[
                    [
                        "Document Number",
                        "Name",
                        "Ship Date",
                        "Outstanding Qty",
                        "Quantity Fulfilled/Received",
                        "Item",
                        "Memo",
                    ]
                ].sort_values("Ship Date"),
                use_container_width=True,
            )

    st.caption("Data auto‑refreshes hourly from NetSuite saved search ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
