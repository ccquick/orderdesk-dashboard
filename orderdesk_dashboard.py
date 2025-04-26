import os
import json
import base64

import pandas as pd
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

    # if your sheet named the new column "Type", rename it so the rest of the code can use "Item Type"
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # 1) Cast core columns
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0)
        - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # 2) Compute today & next Ontario business‐day
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)

    def next_open_day(date: pd.Timestamp) -> pd.Timestamp:
        d = date + pd.Timedelta(days=1)
        while d.weekday() >= 5 or d in ca_holidays:
            d += pd.Timedelta(days=1)
        return d

    tomorrow = next_open_day(today)

    # 3) Bucket logic
    conditions = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0)
        & (df["Quantity Fulfilled/Received"] > 0),
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

    # ─── KPIs (now counting UNIQUE orders) ────────────────────────────────────────
    c1, c2, c3 = st.columns(3)

    # treat "Pending Billing/Partially Fulfilled" as also overdue
    overdue_set = set(
        df.loc[
            (df["Bucket"] == "Overdue")
            | (df["Status"] == "Pending Billing/Partially Fulfilled"),
            "Document Number",
        ].unique()
    )
    due_tomorrow_count = df.loc[
        df["Bucket"] == "Due Tomorrow",
        "Document Number",
    ].nunique()
    partial_count = df.loc[
        df["Bucket"] == "Partially Shipped",
        "Document Number",
    ].nunique()

    c1.metric("Overdue", len(overdue_set))
    c2.metric("Due Tomorrow", int(due_tomorrow_count))
    c3.metric("Partially Shipped", int(partial_count))

    # ─── FILTERS ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect("Customer", sorted(df["Name"].unique()))
        rush_only = st.checkbox("Rush orders only")
        if customers:
            df = df[df["Name"].isin(customers)]
        if rush_only and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # ─── TABS ─────────────────────────────────────────────────────────────────
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
            continue

        # Summary table (one row per order)
        summary = (
            sub.groupby(["Document Number", "Name", "Ship Date"], as_index=False)
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
        tab.dataframe(summary, use_container_width=True)

        # Drill-down selector
        labels = summary.apply(
            lambda r: (
                f"Order {r['Order #']} — {r['Customer']} "
                f"({r['Ship Date'].date()}) | Out: {r['Outstanding']}"
            ),
            axis=1,
        ).tolist()

        sel = tab.selectbox(
            "Show line-items for…",
            ["— choose an order —"] + labels,
            key=bucket,
        )
        if sel != "— choose an order —":
            order_no = int(sel.split()[1])
            detail = sub[sub["Document Number"] == order_no]

            with tab.expander("▶ Full line-item details", expanded=True):
                tab.table(
                    detail[[
                        "Item",
                        "Item Type",
                        "Quantity",
                        "Quantity Fulfilled/Received",
                        "Outstanding Qty",
                        "Memo",
                    ]]
                    .rename(columns={
                        "Quantity": "Qty Ordered",
                        "Quantity Fulfilled/Received": "Qty Shipped",
                        "Outstanding Qty": "Outstanding",
                    })
                )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
