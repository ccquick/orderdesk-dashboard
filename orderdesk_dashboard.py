import os
import json
import base64

import pandas as pd
pd.set_option("display.max_colwidth", None)
from pandas.tseries.offsets import CustomBusinessDay

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

# ─────────────────────────────────────────────────────────────────────────────
# DATA ACCESS
# ─────────────────────────────────────────────────────────────────────────────

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
    """Read Google Sheet → tidy types → add helper columns."""
    ws = get_worksheet()
    df = pd.DataFrame(ws.get_all_records())

    # column‑name harmonisation (older search called it “Type”)
    if "Type" in df.columns and "Item Type" not in df.columns:
        df = df.rename(columns={"Type": "Item Type"})

    # ── basic typing ────────────────────────────────────────────
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Quantity Fulfilled/Received"] = pd.to_numeric(
        df["Quantity Fulfilled/Received"], errors="coerce"
    )
    df["Outstanding Qty"] = (
        df["Quantity"].fillna(0) - df["Quantity Fulfilled/Received"].fillna(0)
    )

    # ── business‑day helpers ───────────────────────────────────
    ca_holidays = holidays.CA(prov="ON")
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    bd = CustomBusinessDay(
        weekmask="Mon Tue Wed Thu Fri",
        holidays=list(ca_holidays.keys()),
    )

    def next_open_day(d):
        c = d + pd.Timedelta(days=1)
        while c.weekday() >= 5 or c in ca_holidays:
            c += pd.Timedelta(days=1)
        return c

    tomorrow = next_open_day(today)

    # ── bucket calculation ─────────────────────────────────────
    conds = [
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] <= today),
        (df["Outstanding Qty"] > 0) & (df["Ship Date"] == tomorrow),
        (df["Outstanding Qty"] > 0) & (df["Quantity Fulfilled/Received"] > 0),
    ]
    labs = ["Overdue", "Due Tomorrow", "Partially Shipped"]
    df["Bucket"] = pd.NA
    for c, l in zip(conds, labs):
        df.loc[c, "Bucket"] = l

    # ── business‑day count of lateness ─────────────────────────
    def calc_days_overdue(ship_dt):
        if pd.isna(ship_dt) or ship_dt >= today:
            return 0
        drange = pd.date_range(
            start=ship_dt.normalize(),
            end=today - pd.Timedelta(days=1),
            freq=bd,
        )
        return len(drange)

    df["Days Overdue"] = df["Ship Date"].apply(calc_days_overdue)

    # ── NEW: flag drop‑shipments (non‑blank Purchase Order) ────
    if "Purchase Order" in df.columns:
        df["Is Drop Ship"] = df["Purchase Order"].astype(str).str.strip().ne("")
    else:
        df["Is Drop Ship"] = False

    return df


# ─────────────────────────────────────────────────────────────────────────────
#   STREAMLIT APP
# ─────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="Orderdesk Dashboard", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")

    df = load_data()

    # ── KPI cards ───────────────────────────────────────────────
    overdue_orders = df.loc[
        (~df["Is Drop Ship"])  # exclude drop‑shipments
        & (
            (df["Bucket"] == "Overdue")
            | (df["Status"] == "Pending Billing/Partially Fulfilled")
        ),
        "Document Number",
    ].nunique()

    due_orders = df.loc[
        (~df["Is Drop Ship"]) & (df["Bucket"] == "Due Tomorrow"),
        "Document Number",
    ].nunique()

    drop_orders = df.loc[df["Is Drop Ship"], "Document Number"].nunique()

    c1, c2, c3 = st.columns(3)
    c1.metric("Overdue", int(overdue_orders))
    c2.metric("Due Tomorrow", int(due_orders))
    c3.metric("Drop Shipments", int(drop_orders))

    # ── sidebar filters ─────────────────────────────────────────
    with st.sidebar:
        st.header("Filters")
        customers = st.multiselect("Customer", sorted(df["Name"].unique()))
        rush_only = st.checkbox("Rush orders only")
        if customers:
            df = df[df["Name"].isin(customers)]
        if rush_only and "Rush Order" in df.columns:
            df = df[df["Rush Order"].str.capitalize() == "Yes"]

    # ── create tabs ────────────────────────────────────────────
    tab_overdue, tab_due, tab_drop = st.tabs(
        ["Overdue", "Due Tomorrow", "Drop Shipments"]
    )
    tabs = {
        "Overdue": tab_overdue,
        "Due Tomorrow": tab_due,
        "Drop Shipments": tab_drop,
    }

    for bucket, tab in tabs.items():
        # -------------------------------------------------------
        # Subset data for this tab
        # -------------------------------------------------------
        if bucket == "Overdue":
            sub = df[
                (~df["Is Drop Ship"])  # exclude drop‑shipments
                & (
                    (df["Bucket"] == "Overdue")
                    | (
                        df["Status"] == "Pending Billing/Partially Fulfilled"
                    )
                )
            ]
        elif bucket == "Due Tomorrow":
            sub = df[(~df["Is Drop Ship"]) & (df["Bucket"] == "Due Tomorrow")]
        else:  # Drop Shipments
            sub = df[df["Is Drop Ship"]]

        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # -------------------------------------------------------
        # Order‑level summary
        # -------------------------------------------------------
        summary = (
            sub.groupby(
                ["Document Number", "Name", "Ship Date", "Status"],
                as_index=False,
            )
            .agg({
                "Order Delay Comments": lambda x: "\n".join(x.dropna().unique()),
                "Days Overdue": "max",
            })
            .rename(
                columns={
                    "Document Number": "Order #",
                    "Name": "Customer",
                    "Ship Date": "Ship Date",
                    "Order Delay Comments": "Delay Comments",
                    "Days Overdue": "Days Late",
                }
            )
            .sort_values("Ship Date")
        )

        summary["Ship Date"] = summary["Ship Date"].dt.date  # drop time

        # Chemical order flag (Assembly/BOM with outstanding qty)
        def has_active_bom(o):
            mask = (
                (sub["Document Number"] == o)
                & (sub["Item Type"] == "Assembly/Bill of Materials")
                & (sub["Outstanding Qty"] > 0)
            )
            return mask.any()

        summary["Chemical Order Flag"] = summary["Order #"].apply(
            lambda o: "⚠️" if has_active_bom(o) else ""
        )

        cols = [
            "Order #",
            "Customer",
            "Ship Date",
            "Status",
            "Delay Comments",
            "Chemical Order Flag",
        ]
        if bucket == "Overdue":
            cols.append("Days Late")

        # -------------------------------------------------------
        # Styling
        # -------------------------------------------------------
        def row_color(r):
            if bucket == "Overdue" or bucket == "Drop Shipments":
                bg = (
                    "#fff3cd"
                    if r["Status"] == "Pending Billing/Partially Fulfilled"
                    else "#f8d7da"
                )
            else:
                bg = "#ffffff"
            return [f"background-color: {bg}"] * len(r)

        styler = (
            summary[cols]
            .style
            .apply(row_color, axis=1)
            .set_properties(**{"text-align": "left"})
        )
        tab.dataframe(styler, use_container_width=True, hide_index=True)

        # -------------------------------------------------------
        # Line‑item drill‑down
        # -------------------------------------------------------
        order_labels = summary.apply(
            lambda r: f"Order {r['Order #']} — {r['Customer']} ({r['Ship Date']})",
            axis=1,
        ).tolist()
        sel = tab.selectbox(
            "Show line-items for…",
            ["— choose an order —"] + order_labels,
            key=bucket,
        )
        if sel != "— choose an order —":
            on = int(sel.split()[1])
            detail = sub[sub["Document Number"] == on]
            detail = detail.drop_duplicates(subset=["Item"])  # dedupe items
            with tab.expander("▶ Full line-item details", expanded=True):
                detail_display = detail[
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
                for c in ["Qty Ordered", "Qty Shipped", "Outstanding"]:
                    detail_display[c] = detail_display[c].apply(
                        lambda x: int(x) if float(x).is_integer() else x
                    )

                def _highlight(row):
                    return [
                        "background-color: #fff3cd" if row["Outstanding"] > 0 else ""
                        for _ in row
                    ]

                styled = (
                    detail_display.style.apply(_highlight, axis=1)
                    .set_properties(**{"text-align": "left"})
                )
                tab.dataframe(styled, use_container_width=True, hide_index=True)

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
