import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import matplotlib.pyplot as plt
import seaborn as sns
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# Set Streamlit Page Configuration
st.set_page_config(page_title="Sales Month-Wise Analytics Dashboard", layout="wide", initial_sidebar_state="expanded")

st.title("📊 Monthly Sales & Insights Comparison Dashboard")

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS & DATA CLEANING
# -----------------------------------------------------------------------------
@st.cache_data
def load_all_sheets(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    all_sheets = xl.sheet_names
    
    monthly_data = {}
    pending_df = pd.DataFrame()

    for sheet in all_sheets:
        clean_sheet_name = sheet.strip().upper()
        
        # Handle Pending Dispatch Sheet
        if "PENDING DISPATCH" in clean_sheet_name:
            df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
            header_idx = None
            for idx, row in df_raw.iterrows():
                row_str = " ".join(row.dropna().astype(str)).upper()
                if "PARTY NAME" in row_str or "DISCRIPTION" in row_str or "PO QTY" in row_str:
                    header_idx = idx
                    break
            if header_idx is not None:
                pending_df = pd.read_excel(xl, sheet_name=sheet, skiprows=header_idx)
                pending_df.columns = [str(c).strip().upper() for c in pending_df.columns]
        else:
            # Monthly Sales Sheets
            df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
            header_idx = None
            for idx, row in df_raw.iterrows():
                row_str = " ".join(row.dropna().astype(str)).upper()
                if "PARTY NAME" in row_str or "PO QTY" in row_str or "SELLER NAME" in row_str:
                    header_idx = idx
                    break
            
            if header_idx is not None:
                df = pd.read_excel(xl, sheet_name=sheet, skiprows=header_idx)
                df.columns = [str(c).strip().upper() for c in df.columns]
                
                # Column mapping for consistency across historical sheets
                rename_map = {
                    'S. NO.': 'S_NO', 'SR NO': 'S_NO', 'SR NO ': 'S_NO',
                    'DO .NO.': 'DO NO', 'DO NO ': 'DO NO',
                    'THIKNESS': 'THICKNESS', 'SIZE (MM)': 'SIZE', 'SIZE ()': 'SIZE',
                    'DISCRIPTION': 'ITEM', 'REMARKS': 'REMARK', 'MOB NO.': 'MOBILE NO'
                }
                df.rename(columns=rename_map, inplace=True)
                monthly_data[clean_sheet_name] = process_dataframe(df, clean_sheet_name)

    # Clean Pending Dispatch Data
    if not pending_df.empty:
        rename_map = {
            'S. NO.': 'S_NO', 'SR NO': 'S_NO', 'DO .NO.': 'DO NO', 'DO NO ': 'DO NO',
            'THIKNESS': 'THICKNESS', 'SIZE (MM)': 'SIZE', 'DISCRIPTION': 'ITEM',
            'REMARKS': 'REMARK', 'MOB NO.': 'MOBILE NO'
        }
        pending_df.rename(columns=rename_map, inplace=True)
        pending_df = process_dataframe(pending_df, "PENDING DISPATCH")

    return monthly_data, pending_df

def process_dataframe(df, sheet_name):
    if df.empty:
        return df

    df = df.dropna(how='all').copy()

    # Clean numeric fields
    num_cols = ["PO QTY (MT)", "PER TON", "DISP.QTY", "PENDING"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Revenue calculation
    if "PO QTY (MT)" in df.columns and "PER TON" in df.columns:
        df["TOTAL REVENUE"] = df["PO QTY (MT)"] * df["PER TON"]
    else:
        df["TOTAL REVENUE"] = 0

    # Clean text columns
    str_cols = ["PO NO", "DO NO", "PARTY NAME", "BROKER", "SECTOR", "PLACE", "SELLER NAME", "ITEM", "GRADE", "STATUS", "SIZE"]
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()

    # Parse Width and Length from SIZE
    def parse_size(size_str):
        if pd.isna(size_str):
            return None, None
        match = re.search(r'(\d+)\s*[X\*x]\s*(\d+)', str(size_str))
        if match:
            return float(match.group(1)), float(match.group(2))
        return None, None

    if "SIZE" in df.columns:
        parsed_sizes = df["SIZE"].apply(parse_size)
        df["WIDTH"] = [p[0] for p in parsed_sizes]
        df["LENGTH"] = [p[1] for p in parsed_sizes]

    df["MONTH_SHEET"] = sheet_name
    return df

# PDF Generation Function
def generate_pdf_report(df_summary, title="Monthly Sales Report"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{title}</b>", styles['Title']))
    story.append(Spacer(1, 15))

    # Keep relevant columns for PDF export
    export_cols = [c for c in ["PO NO", "DO NO", "PARTY NAME", "SELLER NAME", "ITEM", "PO QTY (MT)", "PER TON", "DISP.QTY", "PENDING", "STATUS"] if c in df_summary.columns]
    pdf_df = df_summary[export_cols].head(50)

    data = [pdf_df.columns.tolist()] + pdf_df.astype(str).values.tolist()
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f9f9f9')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))

    story.append(table)
    doc.build(story)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# SIDEBAR CONTROLS (MONTH SELECTION)
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Data Source & Month Selector")
uploaded_file = st.sidebar.file_uploader("Upload Excel File (Sales Data.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    monthly_data, pending_df = load_all_sheets(uploaded_file)
    available_months = list(monthly_data.keys())

    st.sidebar.markdown("---")
    st.sidebar.header("🗓️ Select Month for Analysis")

    # Primary Month Selection
    selected_month = st.sidebar.selectbox("Select Primary Month:", available_months, index=len(available_months)-1 if available_months else 0)

    # Checkbox for Comparing Two Months
    chk_compare_months = st.sidebar.checkbox("Compare Between Two Months")
    compare_month = None
    if chk_compare_months:
        compare_options = [m for m in available_months if m != selected_month]
        compare_month = st.sidebar.selectbox("Select Second Month to Compare:", compare_options)

    # Checkboxes for Additional Views
    chk_overall_comp = st.sidebar.checkbox("Overall Period Analysis (All Months Combined)")
    chk_pending_dispatch = st.sidebar.checkbox("Show Pending Dispatch Section", value=True)

    # Get Selected Data
    if chk_overall_comp:
        primary_df = pd.concat(monthly_data.values(), ignore_index=True) if monthly_data else pd.DataFrame()
        st.info("Displaying Combined Overall Analytics Across All Available Months.")
    else:
        primary_df = monthly_data.get(selected_month, pd.DataFrame())

    # -----------------------------------------------------------------------------
    # KPI SECTION
    # -----------------------------------------------------------------------------
    title_prefix = "All Months Combined" if chk_overall_comp else f"Month: {selected_month}"
    st.markdown(f"## 📌 Key Performance Indicators (KPIs) - {title_prefix}")

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total POs", primary_df["PO NO"].nunique() if "PO NO" in primary_df else 0)
        st.metric("Total DOs", primary_df["DO NO"].nunique() if "DO NO" in primary_df else 0)
    with col2:
        st.metric("Parties", primary_df["PARTY NAME"].nunique() if "PARTY NAME" in primary_df else 0)
        st.metric("Brokers", primary_df["BROKER"].nunique() if "BROKER" in primary_df else 0)
    with col3:
        st.metric("Sellers", primary_df["SELLER NAME"].nunique() if "SELLER NAME" in primary_df else 0)
        st.metric("Sectors", primary_df["SECTOR"].nunique() if "SECTOR" in primary_df else 0)
    with col4:
        st.metric("Items", primary_df["ITEM"].nunique() if "ITEM" in primary_df else 0)
        st.metric("Thicknesses", primary_df["THICKNESS"].nunique() if "THICKNESS" in primary_df else 0)
    with col5:
        st.metric("Grades", primary_df["GRADE"].nunique() if "GRADE" in primary_df else 0)
        st.metric("Places", primary_df["PLACE"].nunique() if "PLACE" in primary_df else 0)

    st.markdown("---")
    
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    sum_po_qty = primary_df["PO QTY (MT)"].sum() if "PO QTY (MT)" in primary_df else 0
    sum_disp_qty = primary_df["DISP.QTY"].sum() if "DISP.QTY" in primary_df else 0
    sum_pending_qty = primary_df["PENDING"].sum() if "PENDING" in primary_df else 0
    total_rev = primary_df["TOTAL REVENUE"].sum() if "TOTAL REVENUE" in primary_df else 0

    kpi1.metric("Total Order Qty (MT)", f"{sum_po_qty:,.2f}")
    kpi2.metric("Dispatched Qty (MT)", f"{sum_disp_qty:,.2f}")
    kpi3.metric("Pending Qty (MT)", f"{sum_pending_qty:,.2f}")
    kpi4.metric("Total Revenue (₹)", f"₹{total_rev:,.2f}")

    # -----------------------------------------------------------------------------
    # MONTH COMPARISON SECTION
    # -----------------------------------------------------------------------------
    if chk_compare_months and compare_month:
        comp_df = monthly_data.get(compare_month, pd.DataFrame())
        st.markdown("---")
        st.markdown(f"### ⚖️ Month Comparison: **{selected_month}** vs **{compare_month}**")

        c_po_qty = comp_df["PO QTY (MT)"].sum() if "PO QTY (MT)" in comp_df else 0
        c_disp_qty = comp_df["DISP.QTY"].sum() if "DISP.QTY" in comp_df else 0
        c_pending_qty = comp_df["PENDING"].sum() if "PENDING" in comp_df else 0
        c_rev = comp_df["TOTAL REVENUE"].sum() if "TOTAL REVENUE" in comp_df else 0

        cc1, cc2, cc3, cc4 = st.columns(4)
        cc1.metric(f"PO Qty ({compare_month})", f"{c_po_qty:,.2f}", delta=f"{sum_po_qty - c_po_qty:,.2f}")
        cc2.metric(f"Dispatched ({compare_month})", f"{c_disp_qty:,.2f}", delta=f"{sum_disp_qty - c_disp_qty:,.2f}")
        cc3.metric(f"Pending ({compare_month})", f"{c_pending_qty:,.2f}", delta=f"{sum_pending_qty - c_pending_qty:,.2f}")
        cc4.metric(f"Revenue ({compare_month})", f"₹{c_rev:,.2f}", delta=f"₹{total_rev - c_rev:,.2f}")

        # Comparative Visualization
        comp_summary = pd.DataFrame({
            "Month": [selected_month, compare_month],
            "PO Qty (MT)": [sum_po_qty, c_po_qty],
            "Dispatched Qty (MT)": [sum_disp_qty, c_disp_qty],
            "Pending Qty (MT)": [sum_pending_qty, c_pending_qty]
        }).melt(id_vars="Month", var_name="Metric", value_name="Tons (MT)")

        fig, ax = plt.subplots(figsize=(8, 3.5))
        sns.barplot(data=comp_summary, x="Metric", y="Tons (MT)", hue="Month", ax=ax, palette="Set2")
        ax.set_title("Volume Comparison Between Selected Months")
        st.pyplot(fig)

    # -----------------------------------------------------------------------------
    # DETAILED INSIGHTS & TABS
    # -----------------------------------------------------------------------------
    st.markdown("---")
    st.header("📈 Detailed Graphical Insights")

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "👤 Party Wise Summary", 
        "👨‍💼 Salesperson Wise", 
        "📐 Size Wise (Width x Length)", 
        "🤝 Broker vs Direct", 
        "📊 Order Status"
    ])

    with tab1:
        st.subheader(f"Party Wise Summary ({selected_month})")
        if not primary_df.empty and "PARTY NAME" in primary_df:
            party_summary = primary_df.groupby("PARTY NAME").agg(
                Total_Orders=("PO NO", "nunique"),
                Total_PO_Qty=("PO QTY (MT)", "sum"),
                Total_Dispatched=("DISP.QTY", "sum"),
                Total_Pending=("PENDING", "sum")
            ).reset_index().sort_values(by="Total_PO_Qty", ascending=False)

            st.dataframe(party_summary, use_container_width=True)

            fig, ax = plt.subplots(figsize=(10, 4))
            sns.barplot(data=party_summary.head(10), x="Total_PO_Qty", y="PARTY NAME", ax=ax, palette="viridis")
            ax.set_title(f"Top 10 Parties by Order Quantity - {selected_month}")
            st.pyplot(fig)

    with tab2:
        st.subheader(f"Salesperson Wise Performance ({selected_month})")
        if not primary_df.empty and "SELLER NAME" in primary_df:
            seller_summary = primary_df.groupby(["SELLER NAME", "STATUS"]).agg(
                PO_Qty=("PO QTY (MT)", "sum"),
                Disp_Qty=("DISP.QTY", "sum"),
                Pending_Qty=("PENDING", "sum")
            ).reset_index()

            st.dataframe(seller_summary, use_container_width=True)

            fig, ax = plt.subplots(figsize=(10, 4))
            sns.barplot(data=primary_df, x="SELLER NAME", y="PO QTY (MT)", hue="STATUS", estimator=sum, ci=None, ax=ax)
            plt.xticks(rotation=45)
            ax.set_title("Salesperson Volume by Order Status")
            st.pyplot(fig)

    with tab3:
        st.subheader("Width & Length Distribution")
        if "WIDTH" in primary_df and "LENGTH" in primary_df and not primary_df.empty:
            fig, ax = plt.subplots(figsize=(8, 4))
            sns.scatterplot(data=primary_df, x="WIDTH", y="LENGTH", size="PO QTY (MT)", hue="ITEM", alpha=0.7, ax=ax)
            ax.set_title("Material Dimensions Scatterplot (Width vs Length)")
            st.pyplot(fig)

    with tab4:
        st.subheader("Broker vs Direct Orders")
        if "BROKER" in primary_df and not primary_df.empty:
            broker_summary = primary_df.groupby("BROKER")["PO QTY (MT)"].sum().reset_index()
            fig, ax = plt.subplots(figsize=(6, 4))
            ax.pie(broker_summary["PO QTY (MT)"], labels=broker_summary["BROKER"], autopct='%1.1f%%', startangle=90)
            ax.set_title("Direct vs Broker Order Quantity Share")
            st.pyplot(fig)

    with tab5:
        st.subheader("Status Breakdown")
        if "STATUS" in primary_df and not primary_df.empty:
            status_counts = primary_df["STATUS"].value_counts().reset_index()
            status_counts.columns = ["STATUS", "COUNT"]
            st.dataframe(status_counts, use_container_width=True)

    # -----------------------------------------------------------------------------
    # EXCEL TABLE DATA & PDF DOWNLOAD
    # -----------------------------------------------------------------------------
    st.markdown("---")
    st.header(f"📋 Data Table - {selected_month}")
    st.dataframe(primary_df, use_container_width=True)

    st.markdown("### 📥 Download PDF Summary")
    pdf_buf = generate_pdf_report(primary_df, title=f"Sales Summary - {selected_month}")
    st.download_button(
        label=f"📄 Download PDF Report ({selected_month})",
        data=pdf_buf,
        file_name=f"Sales_Report_{selected_month}.pdf",
        mime="application/pdf"
    )

    # -----------------------------------------------------------------------------
    # PENDING DISPATCH SECTION (PLACED AT THE END)
    # -----------------------------------------------------------------------------
    if chk_pending_dispatch and not pending_df.empty:
        st.markdown("<br><hr style='border:2px solid red;'><br>", unsafe_allow_clause=True)
        st.header("🔴 PENDING DISPATCH REPORT & DASHBOARD")

        p_col1, p_col2, p_col3, p_col4 = st.columns(4)
        p_col1.metric("Pending Orders Count", len(pending_df))
        p_col2.metric("Total Pending Qty (MT)", f"{pending_df['PENDING'].sum():,.2f}")
        p_col3.metric("Total Revenue (₹)", f"₹{pending_df['TOTAL REVENUE'].sum():,.2f}")
        p_col4.metric("Affected Parties", pending_df["PARTY NAME"].nunique() if "PARTY NAME" in pending_df else 0)

        p_tab1, p_tab2 = st.tabs(["📊 Pending Graphs", "📋 Pending Data Table"])

        with p_tab1:
            if "PARTY NAME" in pending_df and "PENDING" in pending_df:
                fig, ax = plt.subplots(figsize=(10, 4))
                p_party = pending_df.groupby("PARTY NAME")["PENDING"].sum().nlargest(10).reset_index()
                sns.barplot(data=p_party, x="PENDING", y="PARTY NAME", ax=ax, palette="Reds_r")
                ax.set_title("Top 10 Parties with Highest Pending Dispatch Qty")
                st.pyplot(fig)

        with p_tab2:
            st.dataframe(pending_df, use_container_width=True)

else:
    st.info("👈 Please upload your `Sales Data.xlsx` workbook in the left sidebar to start.")
