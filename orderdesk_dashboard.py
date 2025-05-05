import os
import json
import base64
from datetime import datetime, timedelta

import pandas as pd
pd.set_option("display.max_colwidth", None)
from pandas.tseries.offsets import CustomBusinessDay

import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
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

# Custom theme colors
THEME = {
    "primary": "#0083B8",
    "warning": "#FFC107",
    "danger": "#DC3545",
    "success": "#28A745",
    "info": "#17A2B8",
    "light": "#F8F9FA",
    "dark": "#343A40",
}

# Custom CSS for styling
CUSTOM_CSS = """
<style>
    .main {
        background-color: #F8F9FA;
    }
    .main-header {
        font-size: 1.8rem;
        color: #343A40;
        margin-bottom: 1rem;
    }
    .kpi-card {
        background-color: white;
        border-radius: 5px;
        padding: 1rem;
        box-shadow: 0 0.125rem 0.25rem rgba(0, 0, 0, 0.075);
        text-align: center;
    }
    .kpi-value {
        font-size: 2rem;
        font-weight: bold;
    }
    .kpi-label {
        font-size: 1rem;
        color: #6C757D;
    }
    .data-table {
        margin-top: 1rem;
    }
    .stTabs > div > div:first-child {
        background-color: #F8F9FA;
        border-radius: 5px 5px 0 0;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    .stTabs [aria-selected="true"] {
        color: white !important;
        background-color: #0083B8 !important;
        border-radius: 5px 5px 0 0;
    }
    .overdue-row {
        background-color: #F8D7DA;
    }
    .partial-row {
        background-color: #FFF3CD;
    }
    .chemical-flag {
        color: #DC3545;
        font-weight: bold;
    }
    .info-box {
        background-color: #D1ECF1;
        color: #0C5460;
        padding: 1rem;
        border-radius: 5px;
        margin-bottom: 1rem;
    }
    .warning-box {
        background-color: #FFF3CD;
        color: #856404;
        padding: 1rem;
        border-radius: 5px;
        margin-bottom: 1rem;
    }
    .filter-container {
        background-color: white;
        padding: 1rem;
        border-radius: 5px;
        box-shadow: 0 0.125rem 0.25rem rgba(0, 0, 0, 0.075);
        margin-bottom: 1rem;
    }
    .stat-container {
        padding: 1rem;
        background-color: white;
        border-radius: 5px;
        box-shadow: 0 0.125rem 0.25rem rgba(0, 0, 0, 0.075);
        margin-top: 1rem;
    }
    .footer {
        margin-top: 2rem;
        text-align: center;
        color: #6C757D;
        font-size: 0.8rem;
    }
    .search-container {
        margin-bottom: 1rem;
    }
</style>
"""

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
    st.session_state.data_loading = True
    
    ws = get_worksheet()
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

    # ── flag drop‑shipments (non‑blank Purchase Order) ────
    if "Purchase Order" in df.columns:
        df["Is Drop Ship"] = df["Purchase Order"].astype(str).str.strip().ne("")
    else:
        df["Is Drop Ship"] = False
    
    # -- Check for Rush Orders --
    if "Rush Order" in df.columns:
        df["Rush Order"] = df["Rush Order"].astype(str).str.capitalize() == "Yes"
    else:
        df["Rush Order"] = False
    
    st.session_state.data_loading = False
    return df


# ─────────────────────────────────────────────────────────────────────────────
# VISUALIZATION HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def create_order_status_chart(df):
    """Create a pie chart showing order status breakdown."""
    status_counts = df["Status"].value_counts().reset_index()
    status_counts.columns = ["Status", "Count"]
    
    fig = px.pie(
        status_counts, 
        values="Count", 
        names="Status",
        color_discrete_sequence=[THEME["primary"], THEME["warning"], THEME["danger"], THEME["success"], THEME["info"]],
        hole=0.4
    )
    
    fig.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
        showlegend=True,
        height=300
    )
    
    return fig


def create_chemical_orders_chart(df):
    """Create a bar chart showing chemical vs non-chemical orders."""
    # Identify chemical orders (those with Assembly/BOM items)
    chemical_orders = set()
    for order_num in df["Document Number"].unique():
        mask = (
            (df["Document Number"] == order_num)
            & (df["Item Type"] == "Assembly/Bill of Materials")
            & (df["Outstanding Qty"] > 0)
        )
        if mask.any():
            chemical_orders.add(order_num)
    
    # Count orders
    total_orders = df["Document Number"].nunique()
    chemical_count = len(chemical_orders)
    non_chemical_count = total_orders - chemical_count
    
    data = pd.DataFrame({
        "Order Type": ["Chemical Orders", "Non-Chemical Orders"],
        "Count": [chemical_count, non_chemical_count]
    })
    
    fig = px.bar(
        data,
        x="Order Type",
        y="Count",
        color="Order Type",
        color_discrete_map={
            "Chemical Orders": THEME["danger"],
            "Non-Chemical Orders": THEME["primary"]
        },
        text="Count"
    )
    
    fig.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        showlegend=False,
        height=300,
        xaxis_title="",
        yaxis_title="Number of Orders"
    )
    
    fig.update_traces(textposition="outside")
    
    return fig


def create_days_overdue_chart(df):
    """Create a histogram of days overdue for late shipments."""
    overdue_df = df[df["Bucket"] == "Overdue"].copy()
    
    if overdue_df.empty:
        return None
    
    # Group by order and get max days overdue
    order_days = overdue_df.groupby("Document Number")["Days Overdue"].max().reset_index()
    
    fig = px.histogram(
        order_days,
        x="Days Overdue",
        nbins=20,
        color_discrete_sequence=[THEME["danger"]]
    )
    
    fig.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        xaxis_title="Business Days Overdue",
        yaxis_title="Number of Orders",
        height=300
    )
    
    return fig


def create_shipment_timeline(df):
    """Create a timeline of upcoming shipments."""
    today = pd.Timestamp.now(tz=LOCAL_TZ).normalize().tz_localize(None)
    
    # Filter to only upcoming shipments with outstanding quantity
    upcoming_df = df[
        (df["Outstanding Qty"] > 0) & 
        (df["Ship Date"] >= today)
    ].copy()
    
    if upcoming_df.empty:
        return None
    
    # Group by ship date and count orders
    ship_dates = upcoming_df.groupby(upcoming_df["Ship Date"].dt.date)["Document Number"].nunique().reset_index()
    ship_dates.columns = ["Ship Date", "Order Count"]
    
    # Sort by date
    ship_dates = ship_dates.sort_values("Ship Date")
    
    # Only show next 10 days with shipments
    ship_dates = ship_dates.head(10)
    
    fig = px.bar(
        ship_dates,
        x="Ship Date",
        y="Order Count",
        color_discrete_sequence=[THEME["primary"]],
        text="Order Count"
    )
    
    fig.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        xaxis_title="Ship Date",
        yaxis_title="Number of Orders",
        height=300
    )
    
    fig.update_traces(textposition="outside")
    
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# UI COMPONENTS
# ─────────────────────────────────────────────────────────────────────────────

def render_kpi_cards(df):
    """Render KPI summary cards at the top of the dashboard."""
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
    
    rush_orders = df.loc[
        df.get("Rush Order", False) == True,
        "Document Number"
    ].nunique()

    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(
            f"""
            <div class="kpi-card" style="border-left: 5px solid {THEME['danger']}">
                <div class="kpi-value">{int(overdue_orders)}</div>
                <div class="kpi-label">Overdue Shipments</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    
    with col2:
        st.markdown(
            f"""
            <div class="kpi-card" style="border-left: 5px solid {THEME['warning']}">
                <div class="kpi-value">{int(due_orders)}</div>
                <div class="kpi-label">Due Tomorrow</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    
    with col3:
        st.markdown(
            f"""
            <div class="kpi-card" style="border-left: 5px solid {THEME['info']}">
                <div class="kpi-value">{int(drop_orders)}</div>
                <div class="kpi-label">Drop Shipments</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    
    with col4:
        st.markdown(
            f"""
            <div class="kpi-card" style="border-left: 5px solid {THEME['primary']}">
                <div class="kpi-value">{int(rush_orders)}</div>
                <div class="kpi-label">Rush Orders</div>
            </div>
            """,
            unsafe_allow_html=True
        )


def render_filters(df):
    """Render filter controls in the sidebar."""
    st.sidebar.markdown(
        f"""
        <div class="main-header">
            🔍 Filters
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Customer filter
    customers = st.sidebar.multiselect(
        "Customer", 
        options=sorted(df["Name"].unique()),
        key="customer_filter"
    )
    
    # Status filter
    statuses = st.sidebar.multiselect(
        "Status",
        options=sorted(df["Status"].unique()),
        key="status_filter"
    )
    
    # Rush order filter
    rush_only = st.sidebar.checkbox(
        "Rush orders only", 
        key="rush_filter"
    )
    
    # Chemical order filter
    chemical_only = st.sidebar.checkbox(
        "Chemical orders only",
        key="chemical_filter"
    )
    
    # Search box
    search_term = st.sidebar.text_input(
        "Search (Order #, Item, or Memo)",
        key="search_filter"
    )
    
    # Apply filters
    filtered_df = df.copy()
    
    if customers:
        filtered_df = filtered_df[filtered_df["Name"].isin(customers)]
    
    if statuses:
        filtered_df = filtered_df[filtered_df["Status"].isin(statuses)]
    
    if rush_only and "Rush Order" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["Rush Order"] == True]
    
    if chemical_only:
        # Keep only orders with at least one Assembly/BOM item
        chemical_orders = set()
        for order_num in filtered_df["Document Number"].unique():
            mask = (
                (filtered_df["Document Number"] == order_num)
                & (filtered_df["Item Type"] == "Assembly/Bill of Materials")
            )
            if mask.any():
                chemical_orders.add(order_num)
        
        if chemical_orders:
            filtered_df = filtered_df[filtered_df["Document Number"].isin(chemical_orders)]
        else:
            filtered_df = filtered_df.head(0)  # Empty DataFrame with same structure
    
    if search_term:
        search_term = search_term.lower()
        mask = (
            filtered_df["Document Number"].astype(str).str.lower().str.contains(search_term)
            | filtered_df["Item"].astype(str).str.lower().str.contains(search_term)
            | filtered_df["Memo"].astype(str).str.lower().str.contains(search_term)
        )
        filtered_df = filtered_df[mask]
    
    # Last refresh time
    st.sidebar.markdown("---")
    now = datetime.now().strftime("%b %d, %Y %I:%M %p")
    st.sidebar.markdown(
        f"""
        <div class="footer">
            Last refreshed: {now}
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Add refresh button
    if st.sidebar.button("🔄 Refresh Data"):
        st.cache_data.clear()
        st.experimental_rerun()
    
    return filtered_df


def render_dashboard_tabs(df):
    """Render the main dashboard tabs."""
    # Create tabs
    tab_overview, tab_overdue, tab_due, tab_drop = st.tabs([
        "📊 Overview", 
        "⚠️ Overdue", 
        "🔜 Due Tomorrow", 
        "🚚 Drop Shipments"
    ])
    
    # Overview Tab
    with tab_overview:
        st.markdown(
            """
            <div class="main-header">
                📊 Shipment Overview
            </div>
            """,
            unsafe_allow_html=True
        )
        
        # Charts
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(
                """
                <div class="stat-container">
                    <h3>Order Status</h3>
                """,
                unsafe_allow_html=True
            )
            status_chart = create_order_status_chart(df)
            st.plotly_chart(status_chart, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)
        
        with col2:
            st.markdown(
                """
                <div class="stat-container">
                    <h3>Chemical vs Non-Chemical Orders</h3>
                """,
                unsafe_allow_html=True
            )
            chemical_chart = create_chemical_orders_chart(df)
            st.plotly_chart(chemical_chart, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)
        
        col3, col4 = st.columns(2)
        
        with col3:
            st.markdown(
                """
                <div class="stat-container">
                    <h3>Days Overdue Distribution</h3>
                """,
                unsafe_allow_html=True
            )
            overdue_chart = create_days_overdue_chart(df)
            if overdue_chart:
                st.plotly_chart(overdue_chart, use_container_width=True)
            else:
                st.info("No overdue orders to display.")
            st.markdown("</div>", unsafe_allow_html=True)
        
        with col4:
            st.markdown(
                """
                <div class="stat-container">
                    <h3>Upcoming Shipments</h3>
                """,
                unsafe_allow_html=True
            )
            timeline_chart = create_shipment_timeline(df)
            if timeline_chart:
                st.plotly_chart(timeline_chart, use_container_width=True)
            else:
                st.info("No upcoming shipments to display.")
            st.markdown("</div>", unsafe_allow_html=True)
        
        # All Orders Table
        st.markdown(
            """
            <div class="main-header" style="margin-top: 20px;">
                📋 All Orders
            </div>
            """,
            unsafe_allow_html=True
        )
        
        # Group by order for summary view
        order_summary = df.groupby(
            ["Document Number", "Name", "Ship Date", "Status"],
            as_index=False
        ).agg({
            "Outstanding Qty": "sum",
            "Quantity": "sum",
            "Quantity Fulfilled/Received": "sum",
            "Order Delay Comments": lambda x: "\n".join(x.dropna().unique()),
        }).rename(
            columns={
                "Document Number": "Order #",
                "Name": "Customer",
                "Order Delay Comments": "Comments"
            }
        )
        
        # Format date column
        order_summary["Ship Date"] = order_summary["Ship Date"].dt.date
        
        # Add Chemical Flag
        def has_chemicals(order_num):
            mask = (
                (df["Document Number"] == order_num) & 
                (df["Item Type"] == "Assembly/Bill of Materials") &
                (df["Outstanding Qty"] > 0)
            )
            return "⚠️" if mask.any() else ""
        
        order_summary["Chemical"] = order_summary["Order #"].apply(has_chemicals)
        
        # Order columns
        display_cols = [
            "Order #", "Customer", "Ship Date", "Status", 
            "Outstanding Qty", "Quantity", "Comments", "Chemical"
        ]
        
        # Display the table
        st.dataframe(
            order_summary[display_cols],
            use_container_width=True,
            height=400
        )
    
    # Overdue Tab
    with tab_overdue:
        st.markdown(
            """
            <div class="main-header">
                ⚠️ Overdue Shipments
            </div>
            """,
            unsafe_allow_html=True
        )
        
        # Filter for overdue shipments
        overdue_df = df.loc[
            (~df["Is Drop Ship"]) & 
            (
                (df["Bucket"] == "Overdue") |
                (df["Status"] == "Pending Billing/Partially Fulfilled")
            )
        ]
        
        if overdue_df.empty:
            st.success("No overdue shipments! 🎉")
        else:
            st.markdown(
                """
                <div class="warning-box">
                    <h3>⚠️ Attention Required</h3>
                    <p>The following shipments are past their expected ship date. Chemical shipments are marked with ⚠️.</p>
                </div>
                """,
                unsafe_allow_html=True
            )
            
            # Group by order
            overdue_summary = overdue_df.groupby(
                ["Document Number", "Name", "Ship Date", "Status", "Days Overdue"],
                as_index=False
            ).agg({
                "Order Delay Comments": lambda x: "\n".join(x.dropna().unique()),
            }).rename(
                columns={
                    "Document Number": "Order #",
                    "Name": "Customer",
                    "Order Delay Comments": "Delay Reason",
                    "Days Overdue": "Days Late"
                }
            )
            
            # Format date
            overdue_summary["Ship Date"] = overdue_summary["Ship Date"].dt.date
            
            # Add Chemical Flag
            overdue_summary["Chemical"] = overdue_summary["Order #"].apply(
                lambda o: has_chemicals(o)
            )
            
            # Sort by days overdue (most overdue first)
            overdue_summary = overdue_summary.sort_values(
                "Days Late", 
                ascending=False
            )
            
            # Display columns
            display_cols = [
                "Order #", "Customer", "Ship Date", "Status",
                "Days Late", "Delay Reason", "Chemical"
            ]
            
            # Define styling function
            def style_overdue_table(df):
                return df.style.apply(
                    lambda row: [
                        "background-color: #FFF3CD" if row["Status"] == "Pending Billing/Partially Fulfilled"
                        else "background-color: #F8D7DA" 
                        for _ in row
                    ], 
                    axis=1
                )
            
            # Display table
            st.dataframe(
                style_overdue_table(overdue_summary[display_cols]),
                use_container_width=True
            )
            
            # Order details expander
            selected_order = st.selectbox(
                "Show line-items for order:",
                ["-- Select an order --"] + overdue_summary["Order #"].tolist(),
                key="overdue_order_select"
            )
            
            if selected_order != "-- Select an order --":
                with st.expander("Order Details", expanded=True):
                    # Get line items for this order
                    order_items = overdue_df[
                        overdue_df["Document Number"] == selected_order
                    ].drop_duplicates(subset=["Item"])
                    
                    # Format for display
                    display_items = order_items[[
                        "Item", "Item Type", "Quantity", 
                        "Quantity Fulfilled/Received", "Outstanding Qty", "Memo"
                    ]].rename(columns={
                        "Quantity": "Ordered",
                        "Quantity Fulfilled/Received": "Fulfilled",
                        "Outstanding Qty": "Outstanding"
                    })
                    
                    # Style the table
                    def style_items(df):
                        return df.style.apply(
                            lambda row: [
                                "background-color: #FFF3CD" if row["Outstanding"] > 0
                                else "" for _ in row
                            ],
                            axis=1
                        )
                    
                    st.dataframe(
                        style_items(display_items),
                        use_container_width=True
                    )
    
    # Due Tomorrow Tab
    with tab_due:
        st.markdown(
            """
            <div class="main-header">
                🔜 Shipments Due Tomorrow
            </div>
            """,
            unsafe_allow_html=True
        )
        
        # Filter for tomorrow's shipments
        due_tomorrow_df = df.loc[
            (~df["Is Drop Ship"]) & (df["Bucket"] == "Due Tomorrow")
        ]
        
        if due_tomorrow_df.empty:
            st.success("No shipments due tomorrow! 🎉")
        else:
            st.markdown(
                """
                <div class="info-box">
                    <h3>🔔 Coming Up Tomorrow</h3>
                    <p>The following shipments are scheduled for tomorrow. Chemical shipments are marked with ⚠️.</p>
                </div>
                """,
                unsafe_allow_html=True
            )
            
            # Group by order
            tomorrow_summary = due_tomorrow_df.groupby(
                ["Document Number", "Name", "Ship Date", "Status"],
                as_index=False
            ).agg({
                "Order Delay Comments": lambda x: "\n".join(x.dropna().unique()),
            }).rename(
                columns={
                    "Document Number": "Order #",
                    "Name": "Customer",
                    "Order Delay Comments": "Comments"
                }
            )
            
            # Format date
            tomorrow_summary["Ship Date"] = tomorrow_summary["Ship Date"].dt.date
            
            # Add Chemical Flag
            tomorrow_summary["Chemical"] = tomorrow_summary["Order #"].apply(
                lambda o: has_chemicals(o)
            )
            
            # Sort by customer
            tomorrow_summary = tomorrow_summary.sort_values("Customer")
            
            # Display columns
            display_cols = [
                "Order #", "Customer", "Status", "Comments", "Chemical"
            ]
            
            # Display table
            st.dataframe(
                tomorrow_summary[display_cols],
                use_container_width=True
            )
            
            # Order details expander
            selected_order = st.selectbox(
                "Show line-items for order:",
                ["-- Select an order --"] + tomorrow_summary["Order #"].tolist(),
                key="tomorrow_order_select"
            )
            
            if selected_order != "-- Select an order --":
                with st.expander("Order Details", expanded=True):
                    # Get line items for this order
                    order_items = due_tomorrow_df[
                        due_tomorrow_df["Document Number"] == selected_order
                    ].drop_duplicates(subset=["Item"])
                    
                    # Format for display
                    display_items = order_items[[
                        "Item", "Item Type", "Quantity", 
                        "Quantity Fulfilled/Received", "Outstanding Qty", "Memo"
                    ]].rename(columns={
                        "Quantity": "Ordered",
                        "Quantity Fulfilled/Received": "Fulfilled",
                        "Outstanding Qty": "Outstanding"
                    })
                    
                    # Style the table
                    def style_items(df):
                        return df.style.apply(
                            lambda row: [
                                "background-color: #FFF3CD" if row["Outstanding"] > 0
                                else "" for _ in row
                            ],
                            axis=1
                        )
                    
                    st.dataframe(
                        style_items(display_items),
                        use_container_width=True
                    )
    
    # Drop Shipments Tab
    with tab_drop:
        st.markdown(
            """
            <div class="main-header">
                🚚 Drop Shipments
            </div>
            """,
            unsafe_allow_html=True
        )
        
        # Filter for drop shipments
        drop_ship_df = df.loc[df["Is Drop Ship"]]
        
        if drop_ship_df.empty:
            st.success("No active drop shipments! 🎉")
        else:
            st.markdown(
                """
                <div class="info-box">
                    <h3>🚚 Vendor Direct Shipments</h3>
                    <p>The following orders are being shipped directly from vendors to customers. Chemical shipments are marked with ⚠️.</p>
                </div>
                """,
                unsafe_allow_html=True
            )
            
            # Group by order
            drop_summary = drop_ship_df.groupby(
                ["Document Number", "Name", "Ship Date", "Status", "Purchase Order"],
                as_index=False
            ).agg({
                "Order Delay Comments": lambda x: "\n".join(x.dropna().unique()),
                "Days Overdue": "max"
            }).rename(
                columns={
                    "Document Number": "Order #",
                    "Name": "Customer",
                    "Order Delay Comments": "Comments",
                    "Purchase Order": "PO Number",
                    "Days Overdue": "Days Late"
                }
            )
            
            # Format date
            drop_summary["Ship Date"] = drop_summary["Ship Date"].dt.date
            
            # Add Chemical Flag
            drop_summary["Chemical"] = drop_summary["Order #"].apply(
                lambda o: has_chemicals(o)
            )
            
            # Sort by days overdue (most overdue first)
            drop_summary = drop_summary.sort_values(
                "Days Late", 
                ascending=False
            )
            
            # Display columns
            display_cols = [
                "Order #", "Customer", "Ship Date", "Status",
                "PO Number", "Days Late", "Comments", "Chemical"
            ]
            
            # Define styling function
            def style_drop_table(df):
                return df.style.apply(
                    lambda row: [
                        "background-color: #FFF3CD" if row["Status"] == "Pending Billing/Partially Fulfilled"
                        else "background-color: #F8D7DA" if row["Days Late"] > 0
                        else "" 
                        for _ in row
                    ], 
                    axis=1
                )
            
            # Display table
            st.dataframe(
                style_drop_table(drop_summary[display_cols]),
                use_container_width=True
            )
            
            # Order details expander
            selected_order = st.selectbox(
                "Show line-items for order:",
                ["-- Select an order --"] + drop_summary["Order #"].tolist(),
                key="drop_order_select"
            )
            
            if selected_order != "-- Select an order --":
                with st.expander("Order Details", expanded=True):
                    # Get line items for this order
                    order_items = drop_ship_df[
                        drop_ship_df["Document Number"] == selected_order
                    ].drop_duplicates(subset=["Item"])
                    
                    # Format for display
                    display_items = order_items[[
                        "Item", "Item Type", "Quantity", 
                        "Quantity Fulfilled/Received", "Outstanding Qty", "Memo"
                    ]].rename(columns={
                        "Quantity": "Ordered",
                        "Quantity Fulfilled/Received": "Fulfilled",
                        "Outstanding Qty": "Outstanding"
                    })
                    
                    # Style the table
                    def style_items(df):
                        return df.style.apply(
                            lambda row: [
                                "background-color: #FFF3CD" if row["Outstanding"] > 0
                                else "" for _ in row
                            ],
                            axis=1
                        )
                    
                    st.dataframe(
                        style_items(display_items),
                        use_container_width=True
                    )


# ─────────────────────────────────────────────────────────────────────────────
#   STREAMLIT APP
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # Initialize session state
    if "data_loading" not in st.session_state:
        st.session_state.data_loading = False
    
    # Page config
    st.set_page_config(
        page_title="Orderdesk Dashboard",
        page_icon="📦",
        layout="wide"
    )
    
    # Apply custom CSS
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    
    # Page header
    st.markdown(
        """
        <div class="main-header">
            📦 Orderdesk Shipment Status Dashboard
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Load data with caching
    try:
        with st.spinner("Loading data..."):
            data = st.cache_data(ttl=3600)(load_data)()
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        st.stop()
    
    # Render KPI cards
    render_kpi_cards(data)
    
    # Apply filters
    filtered_data = render_filters(data)
    
    # Render dashboard tabs
    render_dashboard_tabs(filtered_data)
    
    # Footer
    st.markdown(
        """
        <div class="footer">
            Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit
        </div>
        """,
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
