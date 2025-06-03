import os
import json
import base64
import datetime

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
PO_LINKS_TAB_NAME = "po_links"  # New tab for PO links data
LOCAL_TZ = "America/Toronto"
# -----------------------------------------------------------------------------

# ─────────────────────────────────────────────────────────────────────────────
# DATA ACCESS
# ─────────────────────────────────────────────────────────────────────────────

def get_worksheet(tab_name):
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
    sheet = client.open_by_url(SHEET_URL).worksheet(tab_name)
    return sheet, client.open_by_url(SHEET_URL)


def get_last_updated_time():
    try:
        # Get the file metadata from Drive API
        _, spreadsheet = get_worksheet(RAW_TAB_NAME)
        
        # Format the timestamp in local timezone
        local_tz = datetime.datetime.now().astimezone().tzinfo
        current_time = datetime.datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
        
        return current_time
    except Exception as e:
        return "Unknown"


def load_data():
    """Read Google Sheet → tidy types → add helper columns."""
    ws, _ = get_worksheet(RAW_TAB_NAME)
    df = pd.DataFrame(ws.get_all_records())

    # column‑name harmonisation (older search called it "Type")
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


def load_po_data():
    """Read PO links from Google Sheet."""
    try:
        ws, _ = get_worksheet(PO_LINKS_TAB_NAME)
        df = pd.DataFrame(ws.get_all_records())
        
        # Clean up column names and extract order numbers without "#"
        if "Sales Order" in df.columns:
            # Extract numeric part from "Sales Order #43244"
            df["Sales Order Number"] = df["Sales Order"].str.extract(r'#?(\d+)').astype(int)
        
        # Convert ETA to datetime
        if "ETA" in df.columns:
            df["ETA"] = pd.to_datetime(df["ETA"], errors="coerce")
            
        return df
    except Exception as e:
        st.warning(f"Unable to load PO links data: {str(e)}")
        return pd.DataFrame()  # Return empty DataFrame on error


def get_po_status_summary(order_num, po_data):
    """Generate PO status summary for an order."""
    if po_data.empty:
        return ""
        
    # Filter PO data for this order
    pos = po_data[po_data["Sales Order Number"] == order_num]
    if pos.empty:
        return ""
    
    received = sum(pos["Status"].str.lower() == "received")
    total = len(pos)
    
    if received == total:
        return "✅ All POs Received"
    elif received > 0:
        return f"🔄 {received}/{total} POs Received"
    else:
        # Find closest ETA
        etas = pos["ETA"].dropna()
        if etas.empty:
            return "⏳ Awaiting POs (No ETA)"
        
        today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
        closest_eta = min(etas)
        days_away = (closest_eta - today).days
        
        if days_away <= 0:
            return "⚠️ PO ETA has passed"
        else:
            # Include the PO number if there's only one
            if len(pos) == 1:
                po_num = pos["Linked PO"].iloc[0]
                return f"⏳ PO {po_num} due in {days_away}d"
            else:
                return f"⏳ Next PO due in {days_away}d"


def display_color_legend():
    """Display a legend explaining the color coding used in the dashboard."""
    st.markdown("### Color Legend")
    
    # Create a 2x2 grid for the color legends
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div style="display: flex; align-items: center; margin-bottom: 8px;">
            <div style="background-color: #f8d7da; width: 20px; height: 20px; margin-right: 8px;"></div>
            <span>Overdue Orders</span>
        </div>
        <div style="display: flex; align-items: center; margin-bottom: 8px;">
            <div style="background-color: #fff3cd; width: 20px; height: 20px; margin-right: 8px;"></div>
            <span>Partially Fulfilled Orders</span>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div style="display: flex; align-items: center; margin-bottom: 8px;">
            <div style="background-color: #cfe2ff; width: 20px; height: 20px; margin-right: 8px;"></div>
            <span>Rush Orders</span>
        </div>
        <div style="display: flex; align-items: center; margin-bottom: 8px;">
            <span style="margin-right: 8px;">🧪</span>
            <span>Chemical Orders</span>
        </div>
        """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
#   STREAMLIT APP
# ─────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="Orderdesk Dashboard", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")
    
    last_updated = get_last_updated_time()
    
    df = load_data()
    po_data = load_po_data()

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

    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
    c1.metric("Overdue", int(overdue_orders))
    c2.metric("Due Tomorrow", int(due_orders))
    c3.metric("Drop Shipments", int(drop_orders))
    c4.markdown(f"<div style='text-align: right; padding-top: 20px; color: #888;'>Last updated: {last_updated}</div>", unsafe_allow_html=True)
    
    # Display color legend
    with st.expander("Show Color Legend", expanded=False):
        display_color_legend()

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
                ["Document Number", "Name", "Ship Date", "Status", "Rush Order"],
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

        # Add PO status information
        if not po_data.empty:
            summary["PO Status"] = summary["Order #"].apply(
                lambda o: get_po_status_summary(o, po_data)
            )
            
            # Create a priority field for sorting
            def get_priority(row):
                # Rush orders are highest priority
                if "Rush Order" in row and row["Rush Order"].capitalize() == "Yes":
                    priority = 0
                # Orders with "PO ETA has passed" are next highest priority
                elif "PO Status" in row and "PO ETA has passed" in str(row["PO Status"]):
                    priority = 1
                # Orders with no PO dependency are next
                elif "PO Status" in row and row["PO Status"] == "":
                    priority = 2
                # Orders with "All POs Received" are next
                elif "PO Status" in row and "All POs Received" in str(row["PO Status"]):
                    priority = 3
                # Orders with some POs received are next
                elif "PO Status" in row and "POs Received" in str(row["PO Status"]):
                    priority = 4
                # Orders waiting on POs with no ETA are next
                elif "PO Status" in row and "No ETA" in str(row["PO Status"]):
                    priority = 5
                # Orders waiting on POs with future ETAs are lowest priority
                else:
                    priority = 6
                return priority
                
            summary["Priority"] = summary.apply(get_priority, axis=1)
            
            # Sort by priority (highest first), then by Days Late (highest first)
            if "Days Late" in summary.columns:
                summary = summary.sort_values(["Days Late", "Priority"], ascending=[False, True])
            else:
                summary = summary.sort_values(["Priority", "Ship Date"], ascending=[True, True])

        # Chemical order flag (Assembly/BOM with outstanding qty)
        def has_active_bom(o):
            mask = (
                (sub["Document Number"] == o)
                & (sub["Item Type"] == "Assembly/Bill of Materials")
                & (sub["Outstanding Qty"] > 0)
            )
            return mask.any()

        summary["Chemical Order Flag"] = summary["Order #"].apply(
            lambda o: "🧪" if has_active_bom(o) else ""
        )

        cols = [
            # 'Days Late' first if present
        ]
        if bucket == "Overdue" and "Days Late" in summary.columns:
            cols.append("Days Late")
        cols += [
            "Order #",
            "Customer",
            "Ship Date",
            "Chemical Order Flag",
            "Status",
            "PO Status",  # New column for PO status
            "Delay Comments",
            "Rush Order",
        ]
        # Remove 'Days Late' from the end if present
        if bucket == "Overdue" and "Days Late" in cols[1:]:
            cols.remove("Days Late")
        # Create a list of columns to show (excluding "Rush Order" which we only use for styling)
        display_cols = [col for col in cols if col != "Rush Order"]

        # -------------------------------------------------------
        # Styling
        # -------------------------------------------------------
        def row_color(r):
            # Check if it's a rush order
            is_rush = r["Rush Order"].capitalize() == "Yes" if "Rush Order" in r else False
            if is_rush:
                # Highlight rush orders in blue
                bg = "#cfe2ff"  # light blue color
            elif bucket == "Overdue" or bucket == "Drop Shipments":
                bg = (
                    "#fff3cd"
                    if r["Status"] == "Pending Billing/Partially Fulfilled"
                    else "#f8d7da"
                )
            else:
                bg = "#ffffff"
            return [f"background-color: {bg}"] * len(r)

        column_config = {
            "Days Late": st.column_config.NumberColumn("Days Late", width="small"),
            "Order #": st.column_config.TextColumn("Order #", width="small"),
            "Customer": st.column_config.TextColumn("Customer", width="medium"),
            "Ship Date": st.column_config.TextColumn("Ship Date", width="small"),
            "Chemical Order Flag": st.column_config.TextColumn("Chem", width="small"),
            "Status": st.column_config.TextColumn("Status", width="medium"),
            "PO Status": st.column_config.TextColumn("PO Status", width="large"),
            "Delay Comments": st.column_config.TextColumn("Delay Comments", width="large"),
        }

        if bucket == "Overdue":
            # Define what counts as actionable
            actionable_mask = (
                summary["PO Status"].isin(["", "✅ All POs Received"]) |
                summary["PO Status"].str.contains("POs Received", na=False)
            )
            actionable_orders = summary[actionable_mask]
            waiting_on_po_orders = summary[~actionable_mask]

            # Show main actionable table (with your current styling)
            tab.dataframe(
                actionable_orders[display_cols]
                    .style
                    .apply(row_color, axis=1)
                    .set_properties(**{"text-align": "left"}),
                use_container_width=True,
                hide_index=True,
                column_config=column_config
            )

            # Show secondary table for waiting on PO (neutral style)
            if not waiting_on_po_orders.empty:
                tab.markdown("#### Orders Waiting on PO (not actionable)")
                tab.dataframe(
                    waiting_on_po_orders[display_cols]
                        .style
                        .set_properties(**{"background-color": "#f5f5f5", "text-align": "left"}),
                    use_container_width=True,
                    hide_index=True,
                    column_config=column_config
                )
        else:
            tab.dataframe(
                summary[display_cols]
                    .style
                    .apply(row_color, axis=1)
                    .set_properties(**{"text-align": "left"}),
                use_container_width=True,
                hide_index=True,
                column_config=column_config
            )

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
            
            # First show order line items
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
            
            # Then show linked PO information if available
            if not po_data.empty:
                po_info = po_data[po_data["Sales Order Number"] == on]
                if not po_info.empty:
                    with tab.expander("📦 Purchase Order Information", expanded=True):
                        po_display = po_info[["Linked PO", "Vendor", "ETA", "Status", "ControlChem Pickup"]]
                        tab.dataframe(po_display, use_container_width=True, hide_index=True)

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
