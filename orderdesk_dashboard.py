import os
import json
import base64

import pandas as pd
pd.set_option("display.max_colwidth", None)

import streamlit as st
import gspread
from google.oauth2 import service_account
import holidays  # pip install holidays

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
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    # rename any new 'Type' → 'Item Type'
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # cast columns
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0)
        - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # business‐day math for ON (skip weekends & holidays)
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    def next_open_day(d):
        c = d + pd.Timedelta(1, "D")
        while c.weekday() >= 5 or c in ca_holidays:
            c += pd.Timedelta(1, "D")
        return c
    tomorrow = next_open_day(today)

    # bucket
    conds = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0) & (df["Quantity Fulfilled/Received"] > 0),
    ]
    labs = ["Overdue", "Due Tomorrow", "Partially Shipped"]
    df["Bucket"] = pd.NA
    for c, l in zip(conds, labs):
        df.loc[c, "Bucket"] = l

    return df


def main():
    st.set_page_config(page_title="Orderdesk Dashboard", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")

    df = load_data()

    # ─── KPIs ─────────────────────────────────────────────────────────────────
    c1, c2 = st.columns(2)
    c1.metric("Overdue", int((df["Bucket"] == "Overdue").sum() + (df["Status"] == "Pending Billing/Partially Fulfilled").sum()))
    c2.metric("Due Tomorrow", int((df["Bucket"] == "Due Tomorrow").sum()))

    # ─── FILTERS ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect("Customer", sorted(df["Name"].unique()), default=None)
        rush_only = st.checkbox("Rush orders only")
        if customers:
            df = df[df["Name"].isin(customers)]
        if rush_only and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # ─── TABS ─────────────────────────────────────────────────────────────────
    tab_overdue, tab_due = st.tabs(["Overdue", "Due Tomorrow"])
    tabs = {"Overdue": tab_overdue, "Due Tomorrow": tab_due}

    # which orders need chemical flag
    chem_orders = {
        o for o in df.loc[
            (df["Item Type"] == "Assembly/Bill of Materials")
            & (df["Outstanding Qty"] > 0),
            "Document Number"
        ].unique()
    }

    for bucket, tab in tabs.items():
        sub = df[df["Bucket"] == bucket]
        # but Overdue tab also pulls in the partially-shipped
        if bucket == "Overdue":
            sub = pd.concat([
                sub,
                df[df["Status"] == "Pending Billing/Partially Fulfilled"]
            ], ignore_index=True)

        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # build summary
        summary = (
            sub.groupby(["Document Number", "Name", "Ship Date", "Status"], as_index=False)
            .agg({
                "Outstanding Qty": "sum",
                "Quantity Fulfilled/Received": "sum",
                "Order Delay Comments": lambda x: "\n".join(x.dropna().unique()),
            })
            .rename(columns={
                "Document Number": "Order #",
                "Name": "Customer",
                "Ship Date": "Ship Date",
                "Outstanding Qty": "Outstanding",
                "Quantity Fulfilled/Received": "Shipped",
                "Order Delay Comments": "Delay Comments",
            })
            .sort_values("Ship Date")
        )
        summary["Chemical Order Flag"] = summary["Order #"].apply(
            lambda o: "⚠️" if o in chem_orders else ""
        )

        # styling function
        def _row_style(row):
            color = (
                "#fff3cd"
                if row["Status"] == "Pending Billing/Partially Fulfilled"
                else "#f8d7da"
            )
            return [f"background-color: {color}"] * len(row)

        styler = (summary
                  .style
                  .apply(_row_style, axis=1)
                  .set_properties(subset=["Order #", "Customer", "Ship Date",
                                          "Outstanding", "Shipped",
                                          "Chemical Order Flag", "Delay Comments", "Status"],
                                  **{"text-align": "left"})
                 )

        tab.dataframe(styler, use_container_width=True)

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
