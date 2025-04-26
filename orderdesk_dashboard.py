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
SHEET_URL = "https://docs.google.com/spreadsheets/d/1-Jkuwl9e1FBY6le08_KA3k7v9J3kfDvSYxw7oOJDtPQ/edit?gid=1789939189"
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
    df = pd.DataFrame(ws.get_all_records())

    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce").dt.tz_localize(None)
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    ).fillna(0)
    df["Outstanding Qty"] = df["Quantity"] - df["Quantity Fulfilled/Received"]

    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    tomorrow = today + pd.Timedelta(days=1)

    conditions = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0) & (df["Quantity Fulfilled/Received"] > 0),
    ]
    labels = ["Overdue", "Due Tomorrow", "Partially Shipped"]
    df["Bucket"] = pd.NA
    for cond, label in zip(conditions, labels):
        df.loc[cond, "Bucket"] = label

    return df


def main():
    st.set_page_config(page_title="Orderdesk Shipment Status", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")

    df = load_data()

    # KPI metrics
    kpis = st.columns(3)
    for col, label in zip(kpis, ["Overdue", "Due Tomorrow", "Partially Shipped"]):
        col.metric(label, int((df["Bucket"] == label).sum()))

    # Sidebar filters
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect("Customer", sorted(df["Name"].unique()), default=None)
        rush = st.checkbox("Rush orders only")
        if customers:
            df = df[df["Name"].isin(customers)]
        if rush and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.strip().str.lower() == "yes"]

    tabs = st.tabs(["Overdue", "Due Tomorrow", "Partially Shipped"])
    for bucket, tab in zip(labels, tabs):
        sub = df[df["Bucket"] == bucket]
        with tab:
            if sub.empty:
                st.info(f"No {bucket.lower()} orders 🎉")
            else:
                # group by order header
                grouped = sub.groupby([
                    "Document Number", "Name", "Ship Date"
                ], as_index=False)
                for _, order in grouped:
                    doc = order["Document Number"].iloc[0]
                    name = order["Name"].iloc[0]
                    ship = order["Ship Date"].iloc[0].date()
                    total_out = order["Outstanding Qty"].sum()
                    total_rec = order["Quantity Fulfilled/Received"].sum()

                    exp_label = f"{doc} | {name} | {ship} | Out: {int(total_out)} | Rec: {int(total_rec)}"
                    with st.expander(exp_label):
                        st.dataframe(
                            order[["Quantity", "Item", "Memo"]]
                            .rename(columns={
                                "Quantity": "Qty",
                                "Item": "Item Code",
                                "Memo": "Description"
                            })
                            .reset_index(drop=True),
                            use_container_width=True
                        )

    st.caption("Data auto‑refreshes hourly from NetSuite saved search ➜ Google Sheet ➜ Streamlit")

if __name__ == "__main__":
    main()
