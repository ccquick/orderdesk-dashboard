import os
import json
import base64

import pandas as pd
pd.set_option("display.max_colwidth", None)

import numpy as np
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
    df = pd.DataFrame(ws.get_all_records())

    # rename Type → Item Type
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # core casts
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = df["Quantity"].fillna(0) - df["Quantity Fulfilled/Received"].fillna(0)

    # Ontario business-day holiday list
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)

    def next_open_day(d):
        c = d + pd.Timedelta(days=1)
        while c.weekday() >= 5 or c in ca_holidays:
            c += pd.Timedelta(days=1)
        return c

    tomorrow = next_open_day(today)

    # buckets
    conds = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0) & (df["Quantity Fulfilled/Received"] > 0),
    ]
    labs = ["Overdue", "Due Tomorrow", "Partially Shipped"]
    df["Bucket"] = pd.NA
    for c, l in zip(conds, labs):
        df.loc[c, "Bucket"] = l

    return df, ca_holidays, today


def main():
    st.set_page_config(page_title="Orderdesk Dashboard", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")

    df, ca_holidays, today = load_data()

    # KPI counts by unique Document Number
    overdue_ids = set(df.loc[df["Bucket"]=="Overdue", "Document Number"])
    partial_ids = set(df.loc[df["Status"]=="Pending Billing/Partially Fulfilled", "Document Number"])
    due_ids     = set(df.loc[df["Bucket"]=="Due Tomorrow", "Document Number"])

    c1, c2 = st.columns(2)
    c1.metric("Overdue", len(overdue_ids | partial_ids))
    c2.metric("Due Tomorrow", len(due_ids))

    # sidebar filters
    with st.sidebar:
        st.header("Filters")
        cust = st.multiselect("Customer", sorted(df["Name"].unique()))
        rush = st.checkbox("Rush orders only")
        if cust:
            df = df[df["Name"].isin(cust)]
        if rush and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize()=="Yes"]

    # tabs
    tab_over, tab_due = st.tabs(["Overdue", "Due Tomorrow"])
    tabs = {"Overdue": tab_over, "Due Tomorrow": tab_due}

    # which orders have BOM lines
    chem_orders = set(
        df.loc[
            (df["Item Type"]=="Assembly/Bill of Materials") &
            (df["Outstanding Qty"]>0),
            "Document Number"
        ].unique()
    )

    # for busday_count
    hols = [d.isoformat() for d in ca_holidays]

    for bucket, tab in tabs.items():
        sub = df[df["Bucket"]==bucket].copy()
        if bucket=="Overdue":
            extra = df[df["Status"]=="Pending Billing/Partially Fulfilled"]
            sub = pd.concat([sub, extra], ignore_index=True)

        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # build summary (one row per order)
        summary = (
            sub.groupby(["Document Number","Name","Ship Date","Status"], as_index=False)
            .agg({
                "Order Delay Comments": lambda x: "\n".join(x.dropna().unique())
            })
            .rename(columns={
                "Document Number":"Order #",
                "Name":"Customer",
                "Ship Date":"Ship Date",
                "Order Delay Comments":"Delay Comments"
            })
            .sort_values("Ship Date")
        )

        # days late
        summary["Days Late"] = summary["Ship Date"].apply(
            lambda d: np.busday_count(
                d.date().isoformat(),
                today.date().isoformat(),
                weekmask="Mon Tue Wed Thu Fri",
                holidays=hols,
            )
        )
        # chemical-order flag
        summary["Chemical Order Flag"] = summary["Order #"].apply(
            lambda o: "⚠️" if o in chem_orders else ""
        )

        # drop index so no extra column
        summary.reset_index(drop=True, inplace=True)

        # color only the Overdue tab
        if bucket=="Overdue":
            def _row_style(r):
                bg = "#fff3cd" if r.Status=="Pending Billing/Partially Fulfilled" else "#f8d7da"
                return [f"background-color: {bg}"]*len(r)

            styled = summary.style.apply(_row_style, axis=1)
            tab.write(styled)  # .write will render the styler without showing index
        else:
            tab.dataframe(summary, use_container_width=True, hide_index=True)

        # drill-down dropdown
        labels = summary["Order #"].astype(str) + " — " + summary["Customer"] + " (" + summary["Ship Date"].dt.date.astype(str) + ")"
        sel = tab.selectbox("Show line-items for…", ["— choose an order —"] + labels.tolist(), key=bucket)
        if sel != "— choose an order —":
            order_no = int(sel.split()[0])
            detail = sub[sub["Document Number"]==order_no]
            with tab.expander("▶ Full line-item details", expanded=True):
                tab.table(
                    detail[[
                        "Item","Item Type",
                        "Quantity","Quantity Fulfilled/Received","Outstanding Qty","Memo"
                    ]].rename(columns={
                        "Quantity":"Qty Ordered",
                        "Quantity Fulfilled/Received":"Qty Shipped",
                        "Outstanding Qty":"Outstanding"
                    })
                )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__=="__main__":
    main()
