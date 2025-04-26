import os
import json
import base64

import pandas as pd
import streamlit as st
import gspread
from google.oauth2 import service_account
import holidays

from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

# … your existing imports & CONFIGURATION above …

def main():
    st.set_page_config(page_title="Orderdesk Dashboard", layout="wide")
    st.title("📦 Orderdesk Shipment Status Dashboard")

    df = load_data()

    # … KPI & sidebar filters untouched …

    tabs = st.tabs(["Overdue", "Due Tomorrow", "Partially Shipped"])
    for bucket, tab in zip(["Overdue","Due Tomorrow","Partially Shipped"], tabs):
        sub = df[df["Bucket"] == bucket]
        if sub.empty:
            tab.info(f"No {bucket.lower()} orders 🎉")
            continue

        # build summary
        summary = (
            sub.groupby(["Document Number","Name","Ship Date"], as_index=False)
               .agg({"Outstanding Qty":"sum","Quantity Fulfilled/Received":"sum"})
               .rename(columns={
                   "Document Number":"Order #",
                   "Name":"Customer",
                   "Outstanding Qty":"Outstanding",
                   "Quantity Fulfilled/Received":"Shipped",
               })
               .sort_values("Ship Date")
        )

        # flag any BOM items outstanding
        def has_bom(o):
            d = sub[sub["Document Number"] == o]
            if "Item Type" not in d: return ""
            m = (d["Item Type"].eq("Assembly/Bill of Materials") & d["Outstanding Qty"].gt(0))
            return "⚠️" if m.any() else ""
        summary["BOM Flag"] = summary["Order #"].apply(has_bom)

        # pack the line-items into a list-of-dicts column
        detail_lookup = {}
        for o in summary["Order #"]:
            det = sub[sub["Document Number"] == o]
            detail_lookup[o] = det[[
                "Item","Item Type","Quantity","Quantity Fulfilled/Received","Outstanding Qty","Memo"
            ]].rename(columns={
                "Quantity":"Qty Ordered",
                "Quantity Fulfilled/Received":"Qty Shipped",
            }).to_dict("records")
        summary["lineItems"] = summary["Order #"].map(detail_lookup)

        # build a gridOptions with masterDetail = true
        gb = GridOptionsBuilder.from_dataframe(summary)
        gb.configure_default_column(flex=1, min_column_width=80, sortable=True)
        gb.configure_column("Ship Date",
                            type=["dateColumnFilter","customDateTimeFormat"],
                            custom_format_string="yyyy-MM-dd")
        # hide our payload column
        gb.configure_column("lineItems", hide=True)

        # tell ag-grid to use its built-in detail renderer
        gb.configure_grid_options(
            masterDetail=True,
            detailCellRenderer='agDetailCellRenderer',
            detailCellRendererParams={
                # how to pull the child rows
                "getDetailRowData": JsCode("""
                    function(params) {
                        params.successCallback(params.data.lineItems);
                    }
                """),
                # what the child grid should look like
                "detailGridOptions": {
                    "columnDefs": [
                        {"field":"Item"},
                        {"field":"Item Type"},
                        {"field":"Qty Ordered"},
                        {"field":"Qty Shipped"},
                        {"field":"Outstanding Qty"},
                        {"field":"Memo", "flex":1}
                    ],
                    "defaultColDef": {"sortable":True, "flex":1, "minWidth":80}
                }
            }
        )
        gridOptions = gb.build()

        tab.markdown("**▶ Click the arrow on a row to expand its line-items**")
        AgGrid(
            summary,
            gridOptions=gridOptions,
            theme="streamlit",
            enable_enterprise_modules=False,
            fit_columns_on_grid_load=True,
        )

    st.caption("Data auto-refreshes hourly from NetSuite ➜ Google Sheet ➜ Streamlit")


if __name__ == "__main__":
    main()
