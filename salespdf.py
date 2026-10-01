import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import matplotlib.pyplot as plt
import seaborn as sns
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Set Streamlit Page Configuration
st.set_page_config(page_title="Sales Data & Insights Dashboard", layout="wide", initial_sidebar_state="expanded")

st.title("📊 Executive Sales Data & Insights Dashboard")

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS & DATA CLEANING
# -----------------------------------------------------------------------------
@st.cache_data
def load_all_sheets(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    all_sheets = xl.sheet_names
    
    monthly_dfs = []
    pending_df = pd.DataFrame()

    standard_cols = [
        "S_NO", "PO NO", "DO NO", "PO DATE", "PARTY NAME", "BROKER", "SECTOR", 
        "PLACE", "SELLER NAME", "ITEM", "THICKNESS", "SIZE", "GRADE", 
        "PO QTY (MT)", "PER TON", "INV NO.", "DATE", "DISP.QTY", "PENDING", 
        "PAYMENT", "DISP. TH.", "STATUS", "REMARK", "MOBILE NO"
    ]

    for sheet in all_sheets:
        clean_sheet_name = sheet.strip().upper()
        
        # Parse Monthly Sales Sheets
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
            df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
            header_idx = None
            for idx, row in df_raw.iterrows():
                row_str = " ".join(row.dropna().astype(str)).upper()
                if "PARTY NAME" in row_str or "PO QTY" in row_str or "SELLER NAME" in row_str:
                    header_idx = idx
                    break
            
            if header_idx is not None:
                df = pd.read_excel(xl, sheet_name=sheet, skiprows=header_idx)
                # Standardize Column Names
                df.columns = [str(c).strip().upper() for c in df.columns]
                
                # Column mapping for inconsistencies
                rename_map = {
                    'S. NO.': 'S_NO', 'SR NO': 'S_NO', 'SR NO ': 'S_NO',
                    'DO .NO.': 'DO NO', 'DO NO ': 'DO NO',
                    'THIKNESS': 'THICKNESS', 'SIZE (MM)': 'SIZE', 'SIZE ()': 'SIZE',
                    'DISCRIPTION': 'ITEM', 'REMARKS': 'REMARK', 'MOB NO.': 'MOBILE NO'
                }
                df.rename(columns=rename_map, inplace=True)
                df['MONTH_SHEET'] = clean_sheet_name
                monthly_dfs.append(df)

    # Combine Monthly Data
    if monthly_dfs:
        combined_df = pd.concat(monthly_dfs, ignore_index=True)
    else:
        combined_df = pd.DataFrame()

    return combined_df, pending_df

def process_dataframe(df):
    if df.empty:
        return df

    # Drop fully empty rows
    df = df.dropna(how='all').copy()

    # Clean numeric fields
    num_cols = ["PO QTY (MT)", "PER TON", "DISP.QTY", "PENDING"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Calculate Revenue
    if "PO QTY (MT)" in df.columns and "PER TON" in df.columns:
        df["TOTAL REVENUE"] = df["PO QTY (MT)"] * df["PER TON"]
    else:
        df["TOTAL REVENUE"] = 0

    # Clean String Fields
    str_cols = ["PO NO", "DO NO", "PARTY NAME", "BROKER", "SECTOR", "PLACE", "SELLER NAME", "ITEM", "GRADE", "STATUS", "SIZE"]
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()

    # Parse Dates
    date_cols = ["PO DATE", "DATE"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format="%d.%m.%Y", errors='coerce').dt.date

    # Extract Width and Length from SIZE (e.g., 1500X6300)
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

    return df

# PDF Generation Function
def generate_pdf_report(df_summary, title="Sales Performance Report"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{title}</b>", styles['Title']))
    story.append(Spacer(1, 15))

    # Convert DataFrame to ReportLab Table
    data = [df_summary.columns.tolist()] + df_summary.astype(str).values.tolist()
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f2f2f2')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))

    story.append(table)
    doc.build(story)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# SIDEBAR CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Data Source & Options")
uploaded_file = st.sidebar.file_uploader("Upload Excel File (Sales Data.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    raw_df, pending_raw_df = load_all_sheets(uploaded_file)
    df = process_dataframe(raw_df)
    pending_df = process_dataframe(pending_raw_df)

    st.sidebar.markdown("---")
    st.sidebar.header("🗓️ Filter & Date Options")

    # Get available dispatch dates
    available_dates = sorted([d for d in df["DATE"].dropna().unique() if pd.notna(d)])
    
    selected_date = st.sidebar.selectbox("Select Target Date for Report:", available_dates)

    # Checkboxes as requested
    chk_compare_dates = st.sidebar.checkbox("Compare Data Between Two Dates")
    compare_date = None
    if chk_compare_dates:
        compare_date = st.sidebar.selectbox("Select Second Date to Compare:", [d for d in available_dates if d != selected_date])

    chk_overall_comp = st.sidebar.checkbox("Overall Period Analysis (All Time)")
    chk_pending_dispatch = st.sidebar.checkbox("Show Pending Dispatch Tab / Section", value=True)

    # Filtered Dataframes
    if chk_overall_comp:
        filtered_df = df.copy()
        st.info("Displaying Overall Period Analytics across all sheets.")
    else:
        filtered_df = df[df["DATE"] == selected_date] if selected_date else df.copy()

    # -----------------------------------------------------------------------------
    # MAIN DASHBOARD: KPIS & METRICS
    # -----------------------------------------------------------------------------
    st.markdown(f"## 📌 Key Performance Indicators (KPIs)")
    if selected_date and not chk_overall_comp:
        st.subheader(f"Data for Date: {selected_date}")

    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("Total POs", filtered_df["PO NO"].nunique() if "PO NO" in filtered_df else 0)
        st.metric("Total DOs", filtered_df["DO NO"].nunique() if "DO NO" in filtered_df else 0)
    with col2:
        st.metric("Parties", filtered_df["PARTY NAME"].nunique() if "PARTY NAME" in filtered_df else 0)
        st.metric("Brokers", filtered_df["BROKER"].nunique() if "BROKER" in filtered_df else 0)
    with col3:
        st.metric("Sellers", filtered_df["SELLER NAME"].nunique() if "SELLER NAME" in filtered_df else 0)
        st.metric("Sectors", filtered_df["SECTOR"].nunique() if "SECTOR" in filtered_df else 0)
    with col4:
        st.metric("Items", filtered_df["ITEM"].nunique() if "ITEM" in filtered_df else 0)
        st.metric("Thicknesses", filtered_df["THICKNESS"].nunique() if "THICKNESS" in filtered_df else 0)
    with col5:
        st.metric("Grades", filtered_df["GRADE"].nunique() if "GRADE" in filtered_df else 0)
        st.metric("Places", filtered_df["PLACE"].nunique() if "PLACE" in filtered_df else 0)

    st.markdown("---")
    
    # Financial & Aggregate Metrics
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    sum_po_qty = filtered_df["PO QTY (MT)"].sum()
    sum_disp_qty = filtered_df["DISP.QTY"].sum()
    sum_pending_qty = filtered_df["PENDING"].sum()
    total_rev = filtered_df["TOTAL REVENUE"].sum()

    kpi_col1.metric("Total Order Qty (MT)", f"{sum_po_qty:,.2f}")
    kpi_col2.metric("Dispatched Qty (MT)", f"{sum_disp_qty:,.2f}")
    kpi_col3.metric("Pending Qty (MT)", f"{sum_pending_qty:,.2f}")
    kpi_col4.metric("Total Revenue (₹)", f"₹{total_rev:,.2f}")

    # DATE COMPARISON SECTION
    if chk_compare_dates and compare_date:
        st.markdown("---")
        st.markdown(f"### ⚖️ Comparison: {selected_date} vs {compare_date}")
        df_comp = df[df["DATE"] == compare_date]

        comp_c1, comp_c2, comp_c3, comp_c4 = st.columns(4)
        c_po_qty = df_comp["PO QTY (MT)"].sum()
        c_disp_qty = df_comp["DISP.QTY"].sum()
        c_pending_qty = df_comp["PENDING"].sum()
        c_rev = df_comp["TOTAL REVENUE"].sum()

        comp_c1.metric(f"PO Qty ({compare_date})", f"{c_po_qty:,.2f}", delta=f"{sum_po_qty - c_po_qty:,.2f}")
        comp_c2.metric(f"Dispatched Qty ({compare_date})", f"{c_disp_qty:,.2f}", delta=f"{sum_disp_qty - c_disp_qty:,.2f}")
        comp_c3.metric(f"Pending Qty ({compare_date})", f"{c_pending_qty:,.2f}", delta=f"{sum_pending_qty - c_pending_qty:,.2f}")
        comp_c4.metric(f"Revenue ({compare_date})", f"₹{c_rev:,.2f}", delta=f"₹{total_rev - c_rev:,.2f}")

    # -----------------------------------------------------------------------------
    # DETAILED INSIGHTS & RELATIONAL ANALYTICS
    # -----------------------------------------------------------------------------
    st.markdown("---")
    st.header("📈 Detailed Graphical Analysis & Relationships")

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "👤 Party & Last Order Stats", 
        "👨‍💼 Salesperson Performance", 
        "📐 Dimension Wise (Width x Length)", 
        "🤝 Broker vs Direct", 
        "📊 Order & Status Summary"
    ])

    with tab1:
        st.subheader("Party-Wise Order Summary & Last Order Date")
        party_summary = df.groupby("PARTY NAME").agg(
            Total_Orders=("PO NO", "nunique"),
            Total_PO_Qty=("PO QTY (MT)", "sum"),
            Total_Dispatched=("DISP.QTY", "sum"),
            Total_Pending=("PENDING", "sum"),
            Last_Order_Date=("PO DATE", "max")
        ).reset_index().sort_values(by="Total_PO_Qty", ascending=False)

        st.dataframe(party_summary, use_container_width=True)

        fig, ax = plt.subplots(figsize=(10, 4))
        top_parties = party_summary.head(10)
        sns.barplot(data=top_parties, x="Total_PO_Qty", y="PARTY NAME", ax=ax, palette="viridis")
        ax.set_title("Top 10 Parties by Sales Order Qty (MT)")
        st.pyplot(fig)

    with tab2:
        st.subheader("Salesperson Wise Breakdown (Dispatched, Pending & Item Status)")
        seller_summary = filtered_df.groupby(["SELLER NAME", "STATUS"]).agg(
            PO_Qty=("PO QTY (MT)", "sum"),
            Disp_Qty=("DISP.QTY", "sum"),
            Pending_Qty=("PENDING", "sum")
        ).reset_index()

        st.dataframe(seller_summary, use_container_width=True)

        fig, ax = plt.subplots(figsize=(10, 4))
        sns.barplot(data=filtered_df, x="SELLER NAME", y="PO QTY (MT)", hue="STATUS", estimator=sum, ci=None, ax=ax)
        plt.xticks(rotation=45)
        ax.set_title("Salesperson Performance by Order Status")
        st.pyplot(fig)

    with tab3:
        st.subheader("Sales Breakdown by Material Size (Width & Length)")
        if "WIDTH" in filtered_df.columns and "LENGTH" in filtered_df.columns:
            fig, ax = plt.subplots(figsize=(8, 4))
            sns.scatterplot(data=filtered_df, x="WIDTH", y="LENGTH", size="PO QTY (MT)", hue="ITEM", alpha=0.7, ax=ax)
            ax.set_title("Material Dimensions (Width vs Length vs Qty)")
            st.pyplot(fig)

    with tab4:
        st.subheader("Order Distribution: Broker vs Direct Sales")
        broker_summary = filtered_df.groupby("BROKER")["PO QTY (MT)"].sum().reset_index()
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.pie(broker_summary["PO QTY (MT)"], labels=broker_summary["BROKER"], autopct='%1.1f%%', startangle=90)
        ax.set_title("Direct vs Broker Order Quantity Share")
        st.pyplot(fig)

    with tab5:
        st.subheader("Item Status Overview")
        status_counts = filtered_df["STATUS"].value_counts().reset_index()
        status_counts.columns = ["STATUS", "COUNT"]
        st.dataframe(status_counts, use_container_width=True)

    # -----------------------------------------------------------------------------
    # EXCEL TABLE DATA & PDF DOWNLOAD
    # -----------------------------------------------------------------------------
    st.markdown("---")
    st.header("📋 Detailed Excel Sheet Data")
    st.dataframe(filtered_df, use_container_width=True)

    # Download Section
    st.markdown("### 📥 Download Report")
    pdf_buffer = generate_pdf_report(filtered_df.head(50), title=f"Sales Report - {selected_date}")
    st.download_button(
        label="📄 Download Detailed PDF Summary Report",
        data=pdf_buffer,
        file_name=f"Sales_Report_{selected_date}.pdf",
        mime="application/pdf"
    )

    # -----------------------------------------------------------------------------
    # PENDING DISPATCH SECTION (PLACED AT THE END SEPARATELY)
    # -----------------------------------------------------------------------------
    if chk_pending_dispatch and not pending_df.empty:
        st.markdown("<br><hr style='border:2px solid red;'><br>", unsafe_allow_clause=True)
        st.header("🔴 PENDING DISPATCH REPORT & DASHBOARD")

        p_col1, p_col2, p_col3, p_col4 = st.columns(4)
        p_col1.metric("Pending Orders Count", len(pending_df))
        p_col2.metric("Total Pending Qty (MT)", f"{pending_df['PENDING'].sum():,.2f}")
        p_col3.metric("Total Pending Value (₹)", f"₹{pending_df['TOTAL REVENUE'].sum():,.2f}")
        p_col4.metric("Affected Parties", pending_df["PARTY NAME"].nunique() if "PARTY NAME" in pending_df else 0)

        p_tab1, p_tab2 = st.tabs(["📊 Pending Analytics & Graphs", "📋 Pending Data Table"])

        with p_tab1:
            fig, ax = plt.subplots(figsize=(10, 4))
            p_party = pending_df.groupby("PARTY NAME")["PENDING"].sum().nlargest(10).reset_index()
            sns.barplot(data=p_party, x="PENDING", y="PARTY NAME", ax=ax, palette="magma")
            ax.set_title("Top 10 Parties with Highest Pending Dispatch Qty")
            st.pyplot(fig)

        with p_tab2:
            st.dataframe(pending_df, use_container_width=True)

else:
    st.info("👈 Please upload your `Sales Data.xlsx` workbook in the sidebar to view the dashboard.")
