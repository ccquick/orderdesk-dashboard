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
    "1-Jkuwl9e1FBY6le08_KA3k7v9J3kfDvSYxw7oOJDtPQ"
    "/edit#gid=1789939189"
)
RAW_TAB_NAME = "raw_orders"
LOCAL_TZ = "America/Toronto"
# -----------------------------------------------------------------------------

def get_worksheet():
    """
    Authenticate with Google Sheets and return the raw_orders worksheet.
    """
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
    """
    Pulls data from Google Sheets, renames & types columns,
    calculates Outstanding Qty and business-day buckets.
    """
    ws = get_worksheet()
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    # rename "Type" → "Item Type" for consistency
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # cast core columns
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0)
        - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # compute 'today' and next Ontario open day
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    def next_open_day(d):
        c = d + pd.Timedelta(days=1)
        while c.weekday() >= 5 or c in ca_holidays:
            c += pd.Timedelta(days=1)
        return c
    tomorrow = next_open_day(today)

    # assign buckets
    conds = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0) & (df["Quantity Fulfilled/Received"] > 0),
    ]
    labs = ["Overdue", "Due Tomorrow", "Partially Shipped"]
    df["Bucket"] = pd.NA
    for cond, lab in zip(conds, labs):
        df.loc[cond, "Bucket"] = lab
    return df


def main():
    st.set_page_config(page_title="Orderdesk Dashboard", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")

    df = load_data()

    # ─── KPIs ──────────────────────────────────────────────────────────
    overdue_count = int(
        (df["Bucket"] == "Overdue").sum()
        + (df["Status"] == "Pending Billing/Partially Fulfilled").sum()
    )
    due_count = int((df["Bucket"] == "Due Tomorrow").sum())
# ─── KPIs ─────────────────────────────────────────────────────────────────
c1, c2 = st.columns(2)

# count unique overdue orders (including partials)
overdue_orders = set(df.loc[df["Bucket"] == "Overdue", "Document Number"])
partial_orders = set(df.loc[df["Status"] == "Pending Billing/Partially Fulfilled", "Document Number"])
overdue_count = len(overdue_orders.union(partial_orders))
c1.metric("Overdue", overdue_count)

# count unique due‐tomorrow orders
due_tomorrow_count = df.loc[df["Bucket"] == "Due Tomorrow", "Document Number"].nunique()
c2.metric("Due Tomorrow", due_tomorrow_count)

    # ─── FILTERS ──────────────────────────────────────────────────────
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect("Customer", sorted(df["Name"].unique()))
        rush_only = st.checkbox("Rush orders only")
        if customers:
            df = df[df["Name"].isin(customers)]
        if rush_only and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # ─── TABS ──────────────────────────────────────────────────────────
    tab_overdue, tab_due = st.tabs(["Overdue", "Due Tomorrow"])
    tabs = {"Overdue": tab_overdue, "Due Tomorrow": tab_due}

    # precompute chemical flag orders
    chem_orders = set(
        df.loc[
            (df["Item Type"] == "Assembly/Bill of Materials")
            & (df["Outstanding Qty"] > 0),
            "Document Number",
        ].unique()
    )

    for bucket, tab in tabs.items():
        sub = df[df["Bucket"] == bucket]
        # include partials in Overdue
        if bucket == "Overdue":
            partial = df[df["Status"] == "Pending Billing/Partially Fulfilled"]
            sub = pd.concat([sub, partial], ignore_index=True)

        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # build summary (one row per order)
        summary = (
            sub.groupby(
                ["Document Number", "Name", "Ship Date", "Status"],
                as_index=False,
            )
            .agg({
                "Order Delay Comments": lambda x: "\n".join(x.dropna().unique()),
            })
            .rename(columns={
                "Document Number": "Order #",
                "Name": "Customer",
                "Ship Date": "Ship Date",
                "Order Delay Comments": "Delay Comments",
            })
            .sort_values("Ship Date")
            # reset index so the dataframe index column disappears
            .reset_index(drop=True)
        )
        summary["Chemical Order Flag"] = summary["Order #"].apply(
            lambda o: "⚠️" if o in chem_orders else ""
        )

        display_cols = [
            "Order #",
            "Customer",
            "Ship Date",
            "Status",
            "Delay Comments",
            "Chemical Order Flag",
        ]

        if bucket == "Overdue":
            # style rows: red for overdue, yellow for partial
            def style_row(r):
                bg = (
                    "#fff3cd"
                    if r["Status"] == "Pending Billing/Partially Fulfilled"
                    else "#f8d7da"
                )
                return [f"background-color: {bg}"] * len(r)

            styled = summary[display_cols].style.apply(style_row, axis=1)
            tab.dataframe(
                styled,
                use_container_width=True,
                hide_index=True,   # remove that extra index column
            )
        else:
            tab.dataframe(
                summary[display_cols],
                use_container_width=True,
                hide_index=True,
            )

        # drill-down dropdown
        labels = summary.apply(
            lambda r: f"Order {r['Order #']} — {r['Customer']} ({r['Ship Date'].date()})",
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
                    detail[
                        [
                            "Item",
                            "Item Type",
                            "Quantity",
                            "Quantity Fulfilled/Received",
                            "Outstanding Qty",
                            "Memo",
                        ]
                    ].rename(
                        columns={
                            "Quantity": "Qty Ordered",
                            "Quantity Fulfilled/Received": "Qty Shipped",
                            "Outstanding Qty": "Outstanding",
                        }
                    )
                )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
