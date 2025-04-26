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
    sh = client.open_by_url(SHEET_URL)
    return sh.worksheet(RAW_TAB_NAME)


def load_data():
    ws = get_worksheet()
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    # lump in your "Type" → "Item Type" rename if needed
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # core casts
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0) - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # Ontario business-day logic
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)

    def next_open_day(d: pd.Timestamp) -> pd.Timestamp:
        c = d + pd.Timedelta(1, "D")
        while c.weekday() >= 5 or c in ca_holidays:
            c += pd.Timedelta(1, "D")
        return c

    tomorrow = next_open_day(today)

    # bucket it
    conds = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0)
        & (df["Quantity Fulfilled/Received"] > 0),
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

    # KPI counts must be distinct orders, not line items
    overdue_orders = df.loc[df["Bucket"] == "Overdue", "Document Number"].nunique()
    partial_orders = df.loc[
        df["Status"] == "Pending Billing/Partially Fulfilled",
        "Document Number",
    ].nunique()
    due_orders = df.loc[df["Bucket"] == "Due Tomorrow", "Document Number"].nunique()

    c1, c2 = st.columns(2)
    c1.metric("Overdue", int(overdue_orders + partial_orders))
    c2.metric("Due Tomorrow", int(due_orders))

    # sidebar filters
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect("Customer", sorted(df["Name"].unique()))
        rush_only = st.checkbox("Rush orders only")
        if customers:
            df = df[df["Name"].isin(customers)]
        if rush_only and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # tabs
    tab_overdue, tab_due = st.tabs(["Overdue", "Due Tomorrow"])
    tabs = {"Overdue": tab_overdue, "Due Tomorrow": tab_due}

    # precompute your chemical-order flag
    chem_orders = set(
        df.loc[
            (df["Item Type"] == "Assembly/Bill of Materials")
            & (df["Outstanding Qty"] > 0),
            "Document Number",
        ].unique()
    )

    for bucket, tab in tabs.items():
        sub = df[df["Bucket"] == bucket]
        if bucket == "Overdue":
            # tack on the partials into that tab
            sub = pd.concat(
                [sub, df[df["Status"] == "Pending Billing/Partially Fulfilled"]],
                ignore_index=True,
            )

        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        summary = (
            sub.groupby(
                ["Document Number", "Name", "Ship Date", "Status"], as_index=False
            )
            .agg(
                {
                    "Quantity": "sum",
                    "Quantity Fulfilled/Received": "sum",
                    "Order Delay Comments": lambda x: "\n".join(x.dropna().unique()),
                }
            )
            .rename(
                columns={
                    "Document Number": "Order #",
                    "Name": "Customer",
                    "Ship Date": "Ship Date",
                    "Quantity": "Qty Ordered",
                    "Quantity Fulfilled/Received": "Qty Shipped",
                    "Order Delay Comments": "Delay Comments",
                }
            )
            .sort_values("Ship Date")
        )
        summary["Chemical Order Flag"] = summary["Order #"].apply(
            lambda o: "⚠️" if o in chem_orders else ""
        )

        # style it, hide that first index column
        styler = summary.style.hide_index()

        if bucket == "Overdue":
            def _row_style(r):
                color = (
                    "#fff3cd"
                    if r["Status"] == "Pending Billing/Partially Fulfilled"
                    else "#f8d7da"
                )
                return [f"background-color: {color}"] * len(r)

            styler = styler.apply(_row_style, axis=1).set_properties(
                **{"text-align": "left"}
            )

        # **<-- here we switch to .write(styler) so hide_index() actually takes effect**
        tab.write(styler, use_container_width=True)

        # drill-down dropdown
        labels = summary.apply(
            lambda r: f"Order {r['Order #']} — {r['Customer']} ({r['Ship Date'].date()})",
            axis=1,
        ).tolist()

        sel = tab.selectbox(
            "Show line-items for…", ["— choose an order —"] + labels, key=bucket
        )
        if sel != "— choose an order —":
            order_no = int(sel.split()[1])
            detail = df[df["Document Number"] == order_no]
            with tab.expander("▶ Full line-item details", expanded=True):
                tab.table(
                    detail[
                        [
                            "Item",
                            "Item Type",
                            "Qty Ordered",
                            "Qty Shipped",
                            "Outstanding Qty",
                            "Delay Comments",
                        ]
                    ]
                )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
