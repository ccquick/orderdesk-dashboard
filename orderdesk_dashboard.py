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
    # make sure this ID exactly matches your sheet’s ID:
    "1-Jkuwl9e1FBY6le08_KA3k7v9J3kfDvSYxw7oOJDtPQ"
    "/edit#gid=1789939189"
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

    # rename “Type” → “Item Type” if needed
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # cast numeric / datetime columns
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0)
        - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # figure out today & next Ontario business-day
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    def next_open_day(d):
        c = d + pd.Timedelta(days=1)
        while c.weekday() >= 5 or c in ca_holidays:
            c += pd.Timedelta(days=1)
        return c
    tomorrow = next_open_day(today)

    # bucket logic
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
    c1.metric(
        "Overdue",
        int(
            (df["Bucket"] == "Overdue").sum()
            + (df["Status"] == "Pending Billing/Partially Fulfilled").sum()
        ),
    )
    c2.metric("Due Tomorrow", int((df["Bucket"] == "Due Tomorrow").sum()))

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

    # which orders need the “chemical” flag?
    chem_orders = {
        o
        for o in df.loc[
            (df["Item Type"] == "Assembly/Bill of Materials")
            & (df["Outstanding Qty"] > 0),
            "Document Number",
        ].unique()
    }

    for bucket, tab in tabs.items():
        sub = df[df["Bucket"] == bucket]
        if bucket == "Overdue":
            extra = df[df["Status"] == "Pending Billing/Partially Fulfilled"]
            sub = pd.concat([sub, extra], ignore_index=True)

        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # build the one-row-per-order summary
        summary = (
            sub.groupby(
                ["Document Number", "Name", "Ship Date", "Status"], as_index=False
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
            # color overdue vs partial
            def _row_style(r):
                bg = "#fff3cd" if r["Status"] == "Pending Billing/Partially Fulfilled" else "#f8d7da"
                return [f"background-color: {bg}"] * len(r)

            styler = (
                summary[display_cols]
                .style
                .hide(axis="index")       # <-- hides that extra index column
                .apply(_row_style, axis=1)
                .set_properties(**{"text-align": "left"})
            )
            tab.dataframe(styler, use_container_width=True)
        else:
            tab.dataframe(summary[display_cols], use_container_width=True)

        # drill-down dropdown
        order_labels = summary.apply(
            lambda r: f"Order {r['Order #']} — {r['Customer']} ({r['Ship Date'].date()})",
            axis=1,
        ).tolist()

        sel = tab.selectbox(
            "Show line-items for…",
            ["— choose an order —"] + order_labels,
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
                    ].rename(columns={
                        "Quantity": "Qty Ordered",
                        "Quantity Fulfilled/Received": "Qty Shipped",
                        "Outstanding Qty": "Outstanding",
                    })
                )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
