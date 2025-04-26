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
SHEET_ID    = "1-Jkuwl9e1FBY6le08_KA3k7v9J3kfDvSYxw7oOJDtPQ"
RAW_TAB     = "raw_orders"
LOCAL_TZ    = "America/Toronto"
# -----------------------------------------------------------------------------

def get_worksheet():
    b64 = os.getenv("GOOGLE_SERVICE_KEY_B64")
    if not b64:
        st.error("🚨 Missing GOOGLE_SERVICE_KEY_B64")
        st.stop()
    info = json.loads(base64.b64decode(b64).decode())
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.readonly",
        ],
    )
    client = gspread.authorize(creds)
    sh = client.open_by_key(SHEET_ID)
    return sh.worksheet(RAW_TAB)


def load_data():
    ws = get_worksheet()
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    # normalize any "Type" → "Item Type"
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # cast core cols
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0) - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # Ontario business‐day logic
    ca_hols = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    def next_open(d):
        c = d + pd.Timedelta(days=1)
        while c.weekday() >= 5 or c in ca_hols:
            c += pd.Timedelta(days=1)
        return c
    tomorrow = next_open(today)

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

    # KPIs: count distinct orders
    overdue_cnt = df.loc[df.Bucket=="Overdue", "Document Number"].nunique()
    partial_cnt = df.loc[df.Status=="Pending Billing/Partially Fulfilled", 
                         "Document Number"].nunique()
    due_cnt     = df.loc[df.Bucket=="Due Tomorrow", "Document Number"].nunique()

    col1, col2 = st.columns(2)
    col1.metric("Overdue", overdue_cnt + partial_cnt)
    col2.metric("Due Tomorrow", due_cnt)

    # sidebar filters
    with st.sidebar:
        st.header("Filters")
        custs = st.multiselect("Customer", sorted(df["Name"].unique()))
        rush = st.checkbox("Rush orders only")
    if custs:
        df = df[df["Name"].isin(custs)]
    if rush and "Rush Order" in df.columns:
        df = df[df["Rush Order"].str.capitalize()=="Yes"]

    # tabs
    tab_ovd, tab_due = st.tabs(["Overdue", "Due Tomorrow"])
    tabs = {"Overdue": tab_ovd, "Due Tomorrow": tab_due}

    # precompute chemical‐flag orders
    chem_set = set(
        df.loc[
            (df["Item Type"]=="Assembly/Bill of Materials") &
            (df["Outstanding Qty"]>0),
            "Document Number"
        ]
    )

    # function to calc business days late
    def biz_days_late(ship):
        if pd.isna(ship):
            return ""
        days = 0
        d = ship
        ca = holidays.CA(prov="ON")
        now = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
        while d < now:
            d += pd.Timedelta(days=1)
            if d.weekday()<5 and d not in ca:
                days += 1
        return days

    for bucket, tab in tabs.items():
        sub = df[df["Bucket"]==bucket].copy()
        if bucket=="Overdue":
            extra = df[df["Status"]=="Pending Billing/Partially Fulfilled"]
            sub = pd.concat([sub, extra], ignore_index=True)

        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # build summary
        summary = (
            sub.groupby(
                ["Document Number","Name","Ship Date","Status"], as_index=False
            )
            .agg({
                "Order Delay Comments": lambda x: "\n".join(x.dropna().unique())
            })
            .rename(columns={
                "Document Number":"Order #",
                "Name":"Customer",
                "Ship Date":"Ship Date",
                "Order Delay Comments":"Delay Comments",
            })
            .sort_values("Ship Date")
        )
        # chemical flag + days late
        summary["Chemical Order Flag"] = summary["Order #"].map(
            lambda o: "⚠️" if o in chem_set else ""
        )
        summary["Days Late"] = summary["Ship Date"].map(biz_days_late)

        # columns to display
        disp_cols = [
            "Order #",
            "Customer",
            "Ship Date",
            "Status",
            "Delay Comments",
            "Chemical Order Flag",
            "Days Late",
        ]

        # show styled only for Overdue
        if bucket=="Overdue":
            def row_color(r):
                return (
                    ["background-color:#fff3cd"]*len(r)
                    if r.Status=="Pending Billing/Partially Fulfilled"
                    else ["background-color:#f8d7da"]*len(r)
                )
            styler = (
                summary[disp_cols]
                .style
                .apply(row_color, axis=1)
            )
            tab.write(styler)

        else:
            tab.dataframe(summary[disp_cols], hide_index=True, use_container_width=True)

        # drilldown dropdown
        labels = summary.apply(
            lambda r: f"Order {r['Order #']} — {r['Customer']} ({r['Ship Date'].date()})",
            axis=1
        ).tolist()
        sel = tab.selectbox(
            "Show line-items for…",
            ["— choose an order —"] + labels,
            key=bucket,
        )
        if sel!="— choose an order —":
            order_no = int(sel.split()[1])
            detail = sub[sub["Document Number"]==order_no]
            with tab.expander("▶ Full line-item details", expanded=True):
                tab.table(
                    detail[[
                        "Item",
                        "Item Type",
                        "Quantity",
                        "Quantity Fulfilled/Received",
                        "Outstanding Qty",
                        "Memo",
                    ]].rename(columns={
                        "Quantity":"Qty Ordered",
                        "Quantity Fulfilled/Received":"Qty Shipped",
                        "Outstanding Qty":"Outstanding",
                    })
                )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__=="__main__":
    main()
