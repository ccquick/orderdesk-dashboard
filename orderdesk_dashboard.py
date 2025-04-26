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

    # rename if needed
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

    # Ontario business-day logic
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    def next_open_day(d):
        c = d + pd.Timedelta(1, "D")
        while c.weekday() >= 5 or c in ca_holidays:
            c += pd.Timedelta(1, "D")
        return c
    tomorrow = next_open_day(today)

    # bucket assignment
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

    # KPI counts by unique orders
    overdue_orders = df.loc[df["Bucket"]=="Overdue", "Document Number"].nunique()
    pbf_orders     = df.loc[df["Status"]=="Pending Billing/Partially Fulfilled", "Document Number"].nunique()
    due_orders     = df.loc[df["Bucket"]=="Due Tomorrow", "Document Number"].nunique()

    c1, c2 = st.columns(2)
    c1.metric("Overdue", int(overdue_orders + pbf_orders))
    c2.metric("Due Tomorrow", int(due_orders))

    # Sidebar filters
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect("Customer", sorted(df["Name"].unique()))
        rush_only  = st.checkbox("Rush orders only")
        if customers:
            df = df[df["Name"].isin(customers)]
        if rush_only and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize()=="Yes"]

    # Tabs
    tab_overdue, tab_due = st.tabs(["Overdue", "Due Tomorrow"])
    tabs = {"Overdue": tab_overdue, "Due Tomorrow": tab_due}

    # Pre‐compute chemical (BOM) flag
    chem_orders = set(
        df.loc[
            (df["Item Type"]=="Assembly/Bill of Materials") &
            (df["Outstanding Qty"]>0),
            "Document Number"
        ].unique()
    )

    for bucket, tab in tabs.items():
        sub = df[df["Bucket"]==bucket]
        if bucket=="Overdue":
            sub = pd.concat([
                sub,
                df[df["Status"]=="Pending Billing/Partially Fulfilled"]
            ], ignore_index=True)

        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # roll up to one row per order
        summary = (
            sub.groupby(
                ["Document Number","Name","Ship Date","Status"],
                as_index=False
            )
            .agg({
                "Outstanding Qty":"sum",
                "Quantity Fulfilled/Received":"sum",
                "Order Delay Comments": lambda s: "\n".join(s.dropna().unique()),
            })
            .rename(columns={
                "Document Number":"Order #",
                "Name":"Customer",
                "Ship Date":"Ship Date",
                "Outstanding Qty":"Outstanding",
                "Quantity Fulfilled/Received":"Shipped",
                "Order Delay Comments":"Delay Comments",
            })
            .sort_values("Ship Date")
        )
        summary["Chemical Order Flag"] = summary["Order #"].apply(
            lambda o: "⚠️" if o in chem_orders else ""
        )

        if bucket=="Overdue":
            # color partially fulfilled yellow, others red
            def _row_style(r):
                return [
                    "background-color: #fff3cd"
                    if r["Status"]=="Pending Billing/Partially Fulfilled"
                    else "background-color: #f8d7da"
                ] * len(r)

            styler = (
                summary.style
                       .apply(_row_style, axis=1)
                       .set_properties(**{"text-align":"left"})
            )
            tab.dataframe(styler, hide_index=True, use_container_width=True)

        else:  # Due Tomorrow
            tab.dataframe(summary, hide_index=True, use_container_width=True)

        # drill-down dropdown
        labels = summary.apply(
            lambda r: (
                f"Order {r['Order #']} — {r['Customer']} "
                f"({r['Ship Date'].date()}) | Out: {r['Outstanding']}"
            ), axis=1
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
                        "Item","Item Type",
                        "Quantity","Quantity Fulfilled/Received",
                        "Outstanding Qty","Order Delay Comments"
                    ]].rename(columns={
                        "Quantity":"Qty Ordered",
                        "Quantity Fulfilled/Received":"Qty Shipped",
                        "Outstanding Qty":"Outstanding",
                        "Order Delay Comments":"Delay Comments"
                    })
                )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__=="__main__":
    main()
