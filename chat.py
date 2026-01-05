import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Page configuration
st.set_page_config(
    page_title="Amazon PO Working Analysis",
    page_icon="📦",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main {
        background: linear-gradient(to bottom right, #EBF4FF, #E0E7FF);
    }
    .stAlert {
        background-color: #EBF4FF;
    }
    </style>
""", unsafe_allow_html=True)

# Title
st.title("📦 Amazon Po Working Analysis Dashboard")
st.markdown("Upload your files and analyze inventory, sales, and RIS data")
st.divider()

# Initialize session state
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'business_pivot' not in st.session_state:
    st.session_state.business_pivot = None

# Sidebar for file uploads
with st.sidebar:
    st.header("📁 Upload Files")
    
    business_report_file = st.file_uploader(
        "Business Report CSV", 
        type=['csv'],
        help="BusinessReport.csv"
    )
    
    pm_file = st.file_uploader(
        "Purchase Master (PM.xlsx)", 
        type=['xlsx', 'xls'],
        help="Contains ASIN, SKU, Brand information"
    )
    
    inventory_file = st.file_uploader(
        "Inventory CSV", 
        type=['csv'],
        help="Current stock levels from Amazon"
    )
    
    ris_file = st.file_uploader(
        "RIS Data (processed_ris_data.xlsx)", 
        type=['xlsx', 'xls'],
        help="Regional Inventory Storage data"
    )
    
    state_fc_file = st.file_uploader(
        "State FC Cluster (Excel)", 
        type=['xlsx', 'xls'],
        help="Fulfillment center to state mapping"
    )
    
    st.divider()
    
    days = st.number_input(
        "Number of Days for Analysis",
        min_value=1,
        max_value=365,
        value=90,
        help="Used to calculate DRR and DOC"
    )
    
    st.divider()
    
    process_button = st.button("🔄 Process Data", type="primary", use_container_width=True)

# Main processing logic
if process_button:
    if not all([business_report_file, pm_file, inventory_file, ris_file, state_fc_file]):
        st.error("⚠️ Please upload all required files!")
    else:
        try:
            with st.spinner("Processing data... Please wait..."):
                # Load Business Report
                business_report = pd.read_csv(business_report_file)
                
                # Clean and prepare data
                business_report["Total Order Items"] = (
                    business_report["Total Order Items"]
                    .astype(str)
                    .str.replace(",", "", regex=False)
                    .astype(float)
                )
                
                business_report["Total Order Items - B2B"] = (
                    business_report["Total Order Items - B2B"]
                    .astype(str)
                    .str.replace(",", "", regex=False)
                    .astype(float)
                )
                
                # Create Business Pivot
                business_pivot = pd.pivot_table(
                    business_report,
                    index=["SKU", "(Child) ASIN"],
                    values=["Total Order Items", "Total Order Items - B2B"],
                    aggfunc="sum"
                ).reset_index()
                
                business_pivot["Total Sales"] = (
                    business_pivot["Total Order Items"] + 
                    business_pivot["Total Order Items - B2B"]
                )
                
                # Load PM file
                pm = pd.read_excel(pm_file)
                
                # Map PM data
                vendor_sku_map = pm.set_index("ASIN")["Vendor SKU Codes"].to_dict()
                brand_map = pm.set_index("ASIN")["Brand"].to_dict()
                brand_manager_map = pm.set_index("ASIN")["Brand Manager"].to_dict()
                
                business_pivot["Vendor SKU Codes"] = business_pivot["(Child) ASIN"].map(vendor_sku_map)
                business_pivot["Brand"] = business_pivot["(Child) ASIN"].map(brand_map)
                business_pivot["Brand Manager"] = business_pivot["(Child) ASIN"].map(brand_manager_map)
                
                # Load Inventory
                inventory = pd.read_csv(inventory_file)
                
                inventory["afn-fulfillable-quantity"] = pd.to_numeric(
                    inventory["afn-fulfillable-quantity"], errors="coerce"
                ).fillna(0)
                
                inventory["afn-reserved-quantity"] = pd.to_numeric(
                    inventory["afn-reserved-quantity"], errors="coerce"
                ).fillna(0)
                
                inventory_pivot = pd.pivot_table(
                    inventory,
                    index="asin",
                    values=["afn-fulfillable-quantity", "afn-reserved-quantity"],
                    aggfunc="sum"
                ).reset_index()
                
                inventory_pivot["Total Stock"] = (
                    inventory_pivot["afn-fulfillable-quantity"] +
                    inventory_pivot["afn-reserved-quantity"]
                )
                
                # Map inventory data
                afn_fulfillable_lookup = inventory_pivot.set_index("asin")["afn-fulfillable-quantity"].to_dict()
                afn_reserved_lookup = inventory_pivot.set_index("asin")["afn-reserved-quantity"].to_dict()
                stock_lookup = inventory_pivot.set_index("asin")["Total Stock"].to_dict()
                
                business_pivot["afn-fulfillable-quantity"] = business_pivot["(Child) ASIN"].map(afn_fulfillable_lookup)
                business_pivot["afn-reserved-quantity"] = business_pivot["(Child) ASIN"].map(afn_reserved_lookup)
                business_pivot["Total Stock"] = business_pivot["(Child) ASIN"].map(stock_lookup)
                
                # Calculate DRR and DOC
                business_pivot["DRR"] = business_pivot["Total Sales"] / days
                business_pivot["DRR"] = business_pivot["DRR"].replace(0, 0.0001)
                business_pivot["DOC"] = business_pivot["Total Stock"] / business_pivot["DRR"]
                business_pivot["DRR"] = business_pivot["DRR"].round(2)
                business_pivot["DOC"] = business_pivot["DOC"].round(1)
                
                # Load RIS Data
                ris_data = pd.read_excel(ris_file)
                
                ris_data["Shipped Quantity"] = pd.to_numeric(
                    ris_data["Shipped Quantity"], errors="coerce"
                ).fillna(0)
                
                asin_fc_ris_pivot = pd.pivot_table(
                    ris_data,
                    index=["ASIN", "FC Cluster"],
                    columns="RIS Status",
                    values="Shipped Quantity",
                    aggfunc="sum",
                    fill_value=0
                ).reset_index()
                
                # RIS High Cluster (sorted by RIS descending)
                if "RIS" in asin_fc_ris_pivot.columns:
                    ris_high = asin_fc_ris_pivot.sort_values("RIS", ascending=True)
                    ris_high_cluster_map = ris_high.set_index("ASIN")["FC Cluster"].to_dict()
                    ris_qty_map = ris_high.set_index("ASIN")["RIS"].to_dict()
                    
                    business_pivot["RIS High Cluster"] = business_pivot["(Child) ASIN"].map(ris_high_cluster_map)
                    business_pivot["RIS Qty"] = business_pivot["(Child) ASIN"].map(ris_qty_map)
                    business_pivot["RIS Qty"] = business_pivot["RIS Qty"].fillna(0)
                    business_pivot["RIS High Cluster"] = business_pivot["RIS High Cluster"].fillna("")
                
                # RIS Low Cluster (sorted by Low RIS descending)
                if "Low RIS" in asin_fc_ris_pivot.columns:
                    ris_low = asin_fc_ris_pivot.sort_values("Low RIS", ascending=True)
                    ris_low_cluster_map = ris_low.set_index("ASIN")["FC Cluster"].to_dict()
                    ris_low_qty_map = ris_low.set_index("ASIN")["Low RIS"].to_dict()
                    
                    business_pivot["RIS Low Cluster"] = business_pivot["(Child) ASIN"].map(ris_low_cluster_map)
                    business_pivot["RIS Low Qty"] = business_pivot["(Child) ASIN"].map(ris_low_qty_map)
                    business_pivot["RIS Low Qty"] = business_pivot["RIS Low Qty"].fillna(0)
                    business_pivot["RIS Low Cluster"] = business_pivot["RIS Low Cluster"].fillna("")
                
                # Load State FC mapping
                state_fc = pd.read_excel(state_fc_file, sheet_name="Sheet1")
                ris_state_map = state_fc.set_index("Cluster")["State"].to_dict()
                
                business_pivot["RIS State"] = business_pivot["RIS High Cluster"].map(ris_state_map)
                business_pivot["RIS State"] = business_pivot["RIS State"].fillna("")
                
                business_pivot["RIS Low State"] = business_pivot["RIS Low Cluster"].map(ris_state_map)
                business_pivot["RIS Low State"] = business_pivot["RIS Low State"].fillna("")
                
                # Create PO State
                business_pivot["PO State"] = business_pivot["DOC"].apply(
                    lambda x: "Create A PO" if x <= 7 else "We have Stock"
                )
                
                # Reorder columns
                column_order = [
                    "SKU", "(Child) ASIN", "Vendor SKU Codes", "Brand", "Brand Manager",
                    "Total Order Items", "Total Order Items - B2B", "Total Sales",
                    "afn-fulfillable-quantity", "afn-reserved-quantity", "Total Stock",
                    "DRR", "DOC", "RIS High Cluster", "RIS Qty", "RIS State",
                    "RIS Low Cluster", "RIS Low Qty", "RIS Low State", "PO State"
                ]
                
                business_pivot = business_pivot[column_order]
                
                st.session_state.business_pivot = business_pivot
                st.session_state.processed = True
                
                st.success("✅ Data processed successfully!")
                st.rerun()
                
        except Exception as e:
            st.error(f"❌ Error processing data: {str(e)}")
            st.exception(e)

# Display results
if st.session_state.processed and st.session_state.business_pivot is not None:
    df = st.session_state.business_pivot
    
    # Summary metrics
    st.header("📊 Summary Metrics")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Products", len(df))
    
    with col2:
        needs_po = len(df[df["PO State"] == "Create A PO"])
        st.metric("Need Purchase Order", needs_po, delta=f"{(needs_po/len(df)*100):.1f}%")
    
    with col3:
        has_stock = len(df[df["PO State"] == "We have Stock"])
        st.metric("Has Adequate Stock", has_stock, delta=f"{(has_stock/len(df)*100):.1f}%")
    
    with col4:
        avg_doc = df["DOC"].mean()
        st.metric("Avg Days of Coverage", f"{avg_doc:.1f}")
    
    st.divider()
    
    # Tabs for different views
    tab1, tab2, tab3 = st.tabs(["📋 All Products", "⚠️ Low Stock Alert", "🗺️ RIS Analysis"])
    
    with tab1:
        st.subheader("All Products Data")
        
        # Filters
        col1, col2, col3 = st.columns(3)
        
        with col1:
            brands = ["All"] + sorted(df["Brand"].dropna().unique().tolist())
            selected_brand = st.selectbox("Filter by Brand", brands)
        
        with col2:
            managers = ["All"] + sorted(df["Brand Manager"].dropna().unique().tolist())
            selected_manager = st.selectbox("Filter by Brand Manager", managers)
        
        with col3:
            po_states = ["All", "Create A PO", "We have Stock"]
            selected_po = st.selectbox("Filter by PO State", po_states)
        
        # Apply filters
        filtered_df = df.copy()
        
        if selected_brand != "All":
            filtered_df = filtered_df[filtered_df["Brand"] == selected_brand]
        
        if selected_manager != "All":
            filtered_df = filtered_df[filtered_df["Brand Manager"] == selected_manager]
        
        if selected_po != "All":
            filtered_df = filtered_df[filtered_df["PO State"] == selected_po]
        
        st.dataframe(filtered_df, use_container_width=True, height=400)
        
        # Download button for All Products
        st.divider()
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            all_products_output = BytesIO()
            with pd.ExcelWriter(all_products_output, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, sheet_name='All Products', index=False)
            all_products_output.seek(0)
            
            st.download_button(
                label="📥 Download All Products Report (Excel)",
                data=all_products_output,
                file_name=f"all_products_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
    
    with tab2:
        st.subheader("⚠️ Products Requiring Purchase Orders (DOC ≤ 7)")
        
        low_stock = df[df["PO State"] == "Create A PO"].sort_values("DOC")
        
        if len(low_stock) > 0:
            st.warning(f"Found {len(low_stock)} products that need purchase orders!")
            
            # Show critical items (DOC = 0)
            critical = low_stock[low_stock["DOC"] == 0]
            if len(critical) > 0:
                st.error(f"🚨 {len(critical)} products have ZERO stock!")
                st.dataframe(
                    critical[["SKU", "(Child) ASIN", "Brand", "Total Stock", "DRR", "DOC", "RIS State"]],
                    use_container_width=True
                )
            
            st.divider()
            st.write("All Low Stock Items:")
            st.dataframe(low_stock, use_container_width=True, height=400)
            
            # Download button for Low Stock
            st.divider()
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                low_stock_output = BytesIO()
                with pd.ExcelWriter(low_stock_output, engine='openpyxl') as writer:
                    low_stock.to_excel(writer, sheet_name='Low Stock Items', index=False)
                    if len(critical) > 0:
                        critical.to_excel(writer, sheet_name='Zero Stock Critical', index=False)
                low_stock_output.seek(0)
                
                st.download_button(
                    label="📥 Download Low Stock Report (Excel)",
                    data=low_stock_output,
                    file_name=f"low_stock_report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
        else:
            st.success("✅ All products have adequate stock!")
    
    with tab3:
        st.subheader("🗺️ RIS (Regional Inventory Storage) Analysis")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Top RIS Clusters (Highest RIS Quantity)**")
            ris_high_summary = df.groupby("RIS High Cluster")["RIS Qty"].sum().sort_values(ascending=False).head(10)
            st.dataframe(ris_high_summary, use_container_width=True)
        
        with col2:
            st.write("**Top Non-RIS Clusters (Highest Non-RIS Quantity)**")
            ris_low_summary = df.groupby("RIS Low Cluster")["RIS Low Qty"].sum().sort_values(ascending=False).head(10)
            st.dataframe(ris_low_summary, use_container_width=True)
        
        st.divider()
        
        st.write("**RIS by State**")
        state_summary = df.groupby("RIS State").agg({
            "RIS Qty": "sum",
            "(Child) ASIN": "count"
        }).sort_values("RIS Qty", ascending=False)
        state_summary.columns = ["Total RIS Quantity", "Number of Products"]
        st.dataframe(state_summary, use_container_width=True)
        
        st.divider()
        
        st.write("**Detailed RIS Data by Product**")
        ris_detailed = df[df["RIS High Cluster"] != ""][["SKU", "(Child) ASIN", "Brand", "Brand Manager", "RIS High Cluster", "RIS Qty", "RIS State", "RIS Low Cluster", "RIS Low Qty", "RIS Low State"]]
        st.dataframe(ris_detailed, use_container_width=True, height=300)
        
        # Download button for RIS Analysis
        st.divider()
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            ris_output = BytesIO()
            with pd.ExcelWriter(ris_output, engine='openpyxl') as writer:
                ris_detailed.to_excel(writer, sheet_name='RIS Detailed', index=False)
                
                ris_cluster_summary = df.groupby("RIS High Cluster").agg({
                    "RIS Qty": "sum",
                    "(Child) ASIN": "count"
                }).reset_index()
                ris_cluster_summary.columns = ["Cluster", "Total RIS Qty", "Product Count"]
                ris_cluster_summary.to_excel(writer, sheet_name='RIS Cluster Summary', index=False)
                
                state_summary_export = df.groupby("RIS State").agg({
                    "RIS Qty": "sum",
                    "(Child) ASIN": "count"
                }).reset_index()
                state_summary_export.columns = ["State", "Total RIS Qty", "Product Count"]
                state_summary_export.to_excel(writer, sheet_name='RIS State Summary', index=False)
            ris_output.seek(0)
            
            st.download_button(
                label="📥 Download RIS Analysis Report (Excel)",
                data=ris_output,
                file_name=f"ris_analysis_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
    
else:
    # Welcome screen
    st.info("👈 Please upload all required files in the sidebar and click 'Process Data' to begin analysis.")
    
    st.markdown("""
    ### 📝 Analysis Overview
    
    This application will:
    
    1. **Calculate Key Metrics:**
       - DRR (Daily Run Rate) = Total Sales / Number of Days
       - DOC (Days of Coverage) = Total Stock / DRR
       - Identify products needing purchase orders (DOC ≤ 7)
    
    2. **RIS Analysis:**
       - Identify highest RIS clusters
       - Map regional inventory distribution
       - Analyze Non-RIS patterns
    
    3. **Generate Reports:**
       - Complete inventory analysis
       - Low stock alerts
       - Regional distribution insights
    
    ### 📂 Required Files:
    - Business Report CSV (3-month sales data)
    - Purchase Master Excel (PM.xlsx with ASIN, SKU, Brand info)
    - Inventory CSV (current stock levels)
    - RIS Data Excel (regional inventory storage)
    - State FC Cluster Excel (fulfillment center mapping)
    """)


