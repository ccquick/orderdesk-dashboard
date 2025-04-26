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
#    service account email in your JSON key.
# 2. Add the Base64‑encoded JSON key to Streamlit secrets under
#    GOOGLE_SERVICE_KEY_B64 (without newlines).
# 3. Set SHEET_URL below to your sheet URL.
# -----------------------------------------------------------------------------

SHEET_URL = "https://docs.google.com/spreadsheets/d/1-Jkuwl9e1FBY6le08_KA3k7v9J3kfDvSYxw7oOJDtPQ/edit"
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
    sheet = client.open_by_url(SHEET_URL)
    return sheet.worksheet(RAW_TAB_NAME)


def load_data():
    ws = get_worksheet()
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    # cast types
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce").dt.tz_localize(None)
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    ).fillna(0).astype(int)
    df["Outstanding Qty"] = df["Quantity"] - df["Quantity Fulfilled/Received"]

    # dates for bucketing
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
    k1, k2, k3 = st.columns(3)
    k1.metric("Overdue", int((df["Bucket"] == "Overdue").sum()))
    k2.metric("Due Tomorrow", int((df["Bucket"] == "Due Tomorrow").sum()))
    k3.metric("Partially Shipped", int((df["Bucket"] == "Partially Shipped").sum()))

    # sidebar filters
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect("Customer", sorted(df["Name"].unique()), default=None)
        rush = st.checkbox("Rush orders only")
        if customers:
            df = df[df["Name"].isin(customers)]
        if rush and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.lower() == "yes"]

    # tabs by bucket
    tabs = st.tabs(["Overdue", "Due Tomorrow", "Partially Shipped"])
    for bucket, tab in zip(["Overdue", "Due Tomorrow", "Partially Shipped"], tabs):
        with tab:
            subset = df[df["Bucket"] == bucket]
            if subset.empty:
                st.info(f"No {bucket.lower()} orders 🎉")
                continue

            # summary + expanders inline
            # group to unique orders
            summary = (
                subset
                .groupby(["Document Number","Name","Ship Date"], as_index=False)
                .agg({
                    "Outstanding Qty": "sum",
                    "Quantity Fulfilled/Received": "sum",
                })
                .sort_values("Ship Date")
                .reset_index(drop=True)
            )

            for _, row in summary.iterrows():
                doc = row["Document Number"]
                name = row["Name"]
                date = row["Ship Date"].strftime("%Y-%m-%d")
                outq = int(row["Outstanding Qty"])
                recq = int(row["Quantity Fulfilled/Received"])
                label = f"Order {doc} — {name} ({date}) | Outstanding: {outq} / {recq}"

                with st.expander(label):
                    lines = subset[subset["Document Number"] == doc]
                    st.dataframe(
                        lines[["Quantity","Item","Memo"]]
                        .reset_index(drop=True),
                        use_container_width=True
                    )

    st.caption("Data auto‑refreshes hourly from NetSuite → Google Sheet → Streamlit")


if __name__ == "__main__":
    main()
