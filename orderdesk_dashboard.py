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
    df = pd.DataFrame(ws.get_all_records())

    # rename if needed
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # cast
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0)
        - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # Ontario business-day
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    def next_open_day(d):
        c = d + pd.Timedelta(days=1)
        while c.weekday() >= 5 or c in ca_holidays:
            c += pd.Timedelta(days=1)
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

    # KPI: count distinct orders
    overdue_cnt = df.loc[df["Bucket"]=="Overdue","Document Number"].nunique()
    partial_cnt = df.loc[df["Status"]=="Pending Billing/Partially Fulfilled","Document Number"].nunique()
    due_cnt     = df.loc[df["Bucket"]=="Due Tomorrow","Document Number"].nunique()

    c1, c2 = st.columns(2)
    c1.metric("Overdue", int(overdue_cnt + partial_cnt))
    c2.metric("Due Tomorrow", int(due_cnt))

    # Sidebar
    with st.sidebar:
        st.header("Filters")
        custs = st.multiselect("Customer", sorted(df["Name"].unique()))
        rush  = st.checkbox("Rush orders only")
        if custs:
            df = df[df["Name"].isin(custs)]
        if rush and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize()=="Yes"]

    # Tabs
    tab_over, tab_due = st.tabs(["Overdue","Due Tomorrow"])
    tabs = {"Overdue":tab_over, "Due Tomorrow":tab_due}

    # which orders need the chemical‐flag
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

        # roll‐up per order
        summary = (
            sub.groupby(
                ["Document Number","Name","Ship Date","Status"], as_index=False
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

        # drop Outstanding & Shipped from the displayed grid
        display = summary.drop(columns=["Outstanding","Shipped"])

        if bucket=="Overdue":
            def row_style(r):
                bg = "#fff3cd" if r["Status"]=="Pending Billing/Partially Fulfilled" else "#f8d7da"
                return [f"background-color: {bg}"]*len(r)

            styler = (
                display.style
                       .apply(row_style, axis=1)
                       .set_properties(**{"text-align":"left"})
            )
            tab.dataframe(styler, hide_index=True, use_container_width=True)
        else:
            tab.dataframe(display, hide_index=True, use_container_width=True)

        # drill‐down dropdown
        labels = summary.apply(
            lambda r: (
                f"Order {r['Order #']} — {r['Customer']} "
                f"({r['Ship Date'].date()}) | Out: {r['Outstanding']}"
            ), axis=1
        ).tolist()
        sel = tab.selectbox(
            "Show line-items for…",
            ["— choose an order —"] + labels,
            key=bucket
        )
        if sel!="— choose an order —":
            no = int(sel.split()[1])
            detail = sub[sub["Document Number"]==no]
            with tab.expander("▶ Full line-item details", expanded=True):
                tab.table(
                    detail[[
                        "Item","Item Type",
                        "Quantity","Quantity Fulfilled/Received",
                        "Outstanding Qty","Order Delay Comments"
                    ]]
                    .rename(columns={
                        "Quantity":"Qty Ordered",
                        "Quantity Fulfilled/Received":"Qty Shipped",
                        "Outstanding Qty":"Outstanding",
                        "Order Delay Comments":"Delay Comments"
                    })
                )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__=="__main__":
    main()
