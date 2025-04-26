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
    # open_by_url sometimes strips off the “#gid=…” so this works reliably:
    ss = client.open_by_key(SHEET_URL.split("/d/")[1].split("/")[0])
    return ss.worksheet(RAW_TAB_NAME)


def load_data():
    ws = get_worksheet()
    data = ws.get_all_records()
    df = pd.DataFrame(data)

    # normalize any newly-named “Type” column
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # core casts
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0)
        - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # what is “today” and “tomorrow” in our Ontario business calendar?
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)

    def next_open_day(d: pd.Timestamp) -> pd.Timestamp:
        c = d + pd.Timedelta(days=1)
        while c.weekday() >= 5 or c in ca_holidays:
            c += pd.Timedelta(days=1)
        return c

    tomorrow = next_open_day(today)

    # bucket each line
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

    # ─── KPIs ─────────────────────────────────────────────────────────────────
    c1, c2 = st.columns(2)
    # count distinct orders for each
    overdue_orders = df.loc[df.Bucket=="Overdue", "Document Number"].nunique()
    due_orders     = df.loc[df.Bucket=="Due Tomorrow", "Document Number"].nunique()
    # plus any “Pending Billing/Partially Fulfilled” lines also live under Overdue
    pbf_orders     = df.loc[df.Status=="Pending Billing/Partially Fulfilled", "Document Number"].nunique()
    c1.metric("Overdue", overdue_orders + pbf_orders)
    c2.metric("Due Tomorrow",   due_orders)

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
    tab_overdue, tab_due = st.tabs(["Overdue", "Due Tomorrow"])
    tabs = {"Overdue": tab_overdue, "Due Tomorrow": tab_due}

    # precompute chemical flag orders
    chem_orders = {
        o for o in df.loc[
            (df["Item Type"]=="Assembly/Bill of Materials") &
            (df["Outstanding Qty"]>0),
            "Document Number"
        ].unique()
    }

    for bucket, tab in tabs.items():
        sub = df[df["Bucket"]==bucket]
        if bucket=="Overdue":
            # include partially-billed/fulfilled as “overdue”
            extra = df[df["Status"]=="Pending Billing/Partially Fulfilled"]
            sub = pd.concat([sub, extra], ignore_index=True)

        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # roll up one row per order
        summary = (
            sub.groupby(
                ["Document Number","Name","Ship Date","Status"], as_index=False
            )
            .agg({
                "Outstanding Qty":"sum",
                "Quantity Fulfilled/Received":"sum",
                "Order Delay Comments": lambda x:"\n".join(x.dropna().unique()),
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
        summary["Chemical Order Flag"] = summary["Order #"].map(
            lambda o: "⚠️" if o in chem_orders else ""
        )
        # add days-late col
        ca_hols = holidays.CA(prov="ON")
        def business_days_late(ship:pd.Timestamp):
            if ship is pd.NaT: return ""
            days=0; d=ship
            while d < pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None):
                d+=pd.Timedelta(days=1)
                if d.weekday()<5 and d not in ca_hols:
                    days+=1
            return days
        summary["Days Late"] = summary["Ship Date"].map(business_days_late)

        # render
        cols = ["Order #","Customer","Ship Date","Status","Delay Comments","Chemical Order Flag","Days Late"]
        if bucket=="Overdue":
            # style overdue tab with reds & yellows, and hide that index
            def _row_style(r):
                bg = "#fff3cd" if r.Status=="Pending Billing/Partially Fulfilled" else "#f8d7da"
                return [f"background-color: {bg}"]*len(r)

            styled = (
                summary[cols]
                .style
                .apply(_row_style, axis=1)
                .set_table_styles([
                    {"selector":"th.row_heading","props":"display:none;"},
                    {"selector":"td.row_heading","props":"display:none;"},
                ])
                .set_properties(**{"text-align":"left"})
            )
            tab.write(styled)

        else:
            # Due Tomorrow: no color, but still hide pandas index
            styled = (
                summary[cols]
                .style
                .set_table_styles([
                    {"selector":"th.row_heading","props":"display:none;"},
                    {"selector":"td.row_heading","props":"display:none;"},
                ])
            )
            tab.write(styled)

        # drill-down dropdown
        labels = summary.apply(
            lambda r: f"Order {r['Order #']} — {r['Customer']} ({r['Ship Date'].date()})",
            axis=1
        ).tolist()
        sel = tab.selectbox("Show line-items for…", ["— choose an order —"]+labels, key=bucket)
        if sel!="— choose an order —":
            ord_no = int(sel.split()[1])
            detail = sub[sub["Document Number"]==ord_no]
            with tab.expander("▶ Full line-item details", expanded=True):
                tab.table(
                    detail[[
                        "Item","Item Type","Quantity","Quantity Fulfilled/Received",
                        "Outstanding Qty","Memo"
                    ]]
                    .rename(columns={
                        "Quantity":"Qty Ordered",
                        "Quantity Fulfilled/Received":"Qty Shipped",
                        "Outstanding Qty":"Outstanding",
                    })
                )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__=="__main__":
    main()
