import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
import seaborn as sns
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Sales Analytics & PDF Exporter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Monthly & Overall Sales Analytics Dashboard")

# -----------------------------------------------------------------------------
# DATA PROCESSING & CLEANING
# -----------------------------------------------------------------------------
def process_dataframe(df, sheet_name, is_pending_sheet=False):
    """Standardizes column names, removes summary/total rows, and calculates totals."""
    if df.empty:
        return df

    df = df.dropna(how='all').copy()

    # Column name cleaning & mapping
    rename_map = {
        'S. NO.': 'S_NO', 'SR NO': 'S_NO', 'SR NO ': 'S_NO', 'S.NO': 'S_NO',
        'DO .NO.': 'DO NO', 'DO NO ': 'DO NO', 'D.O. NO.': 'DO NO',
        'THIKNESS': 'THICKNESS', 'SIZE (MM)': 'SIZE', 'SIZE ()': 'SIZE',
        'DISCRIPTION': 'ITEM', 'DESCRIPTION': 'ITEM', 'ITEM NAME': 'ITEM',
        'REMARKS': 'REMARK', 'MOB NO.': 'MOBILE NO', 'MOBILE': 'MOBILE NO'
    }
    df.rename(columns=rename_map, inplace=True)

    # Exclude bottom total/summary rows from calculations
    if "PARTY NAME" in df.columns:
        df = df[~df["PARTY NAME"].astype(str).str.upper().str.contains("TOTAL|SUM|GRAND TOTAL|AVERAGE", na=False)]
    if "S_NO" in df.columns:
        df = df[df["S_NO"].notna()]

    # Clean numeric fields
    num_cols = ["PO QTY (MT)", "PER TON", "DISP.QTY", "PENDING"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Compute Revenue: PO QTY * PER TON (or PENDING * PER TON for pending sheets)
    qty_col = "PENDING" if (is_pending_sheet and "PENDING" in df.columns) else "PO QTY (MT)"
    if qty_col in df.columns and "PER TON" in df.columns:
        df["TOTAL REVENUE"] = df[qty_col] * df["PER TON"]
    elif "PO QTY (MT)" in df.columns and "PER TON" in df.columns:
        df["TOTAL REVENUE"] = df["PO QTY (MT)"] * df["PER TON"]
    else:
        df["TOTAL REVENUE"] = 0.0

    # Format text fields
    str_cols = ["PO NO", "DO NO", "PARTY NAME", "BROKER", "SECTOR", "PLACE", "SELLER NAME", "ITEM", "GRADE", "STATUS", "SIZE"]
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()

    df["MONTH_SHEET"] = sheet_name
    return df

@st.cache_data
def load_all_sheets(uploaded_file):
    """Loads all monthly tabs and Pending Dispatch tab from workbook independently."""
    xl = pd.ExcelFile(uploaded_file)
    all_sheets = xl.sheet_names
    
    monthly_data = {}
    pending_df = pd.DataFrame()

    for sheet in all_sheets:
        clean_sheet_name = sheet.strip().upper()
        
        # Pending Dispatch Sheet Detection
        if "PENDING DISPATCH" in clean_sheet_name or "PENDING" in clean_sheet_name:
            df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
            header_idx = None
            for idx, row in df_raw.iterrows():
                row_str = " ".join(row.dropna().astype(str)).upper()
                if "PARTY NAME" in row_str or "DISCRIPTION" in row_str or "PO QTY" in row_str or "PENDING" in row_str:
                    header_idx = idx
                    break
            if header_idx is not None:
                pending_df = pd.read_excel(xl, sheet_name=sheet, skiprows=header_idx)
                pending_df.columns = [str(c).strip().upper() for c in pending_df.columns]
                pending_df = process_dataframe(pending_df, "PENDING DISPATCH", is_pending_sheet=True)
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
                monthly_data[clean_sheet_name] = process_dataframe(df, clean_sheet_name)

    return monthly_data, pending_df

# -----------------------------------------------------------------------------
# ADVANCED PDF GENERATOR (KPIs + CHART + TABLES)
# -----------------------------------------------------------------------------
def generate_full_pdf_report(month_name, metrics, chart_buf, detail_df, party_df, seller_df):
    """Generates an executive PDF report containing KPIs, Chart images, and detailed Data Tables."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20
    )
    styles = getSampleStyleSheet()
    story = []

    # Title
    story.append(Paragraph(f"<b>Executive Sales Report - {month_name}</b>", styles['Title']))
    story.append(Spacer(1, 10))

    # 1. KPI Summary Block
    kpi_data = [
        ["Total POs", "Total DOs", "No. of Parties", "Total PO Qty (MT)", "Total Revenue (₹)", "Dispatched Qty (MT)", "Pending Qty (MT)"],
        [
            f"{metrics['po_count']}",
            f"{metrics['do_count']}",
            f"{metrics['parties_count']}",
            f"{metrics['po_qty']:,.2f}",
            f"₹{metrics['revenue']:,.2f}",
            f"{metrics['disp_qty']:,.2f}",
            f"{metrics['pending_qty']:,.2f}"
        ]
    ]
    kpi_table = Table(kpi_data, colWidths=[100]*7)
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#D9E1F2')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 15))

    # 2. Add Chart Image
    if chart_buf is not None:
        chart_buf.seek(0)
        img = RLImage(chart_buf, width=700, height=220)
        story.append(img)
        story.append(Spacer(1, 15))

    # Helper function for rendering tables
    def build_section_table(df_data, title, cols_limit=10):
        story.append(Paragraph(f"<b>{title}</b>", styles['Heading2']))
        story.append(Spacer(1, 5))
        
        display_df = df_data.iloc[:, :cols_limit].copy()
        for c in display_df.select_dtypes(include=[np.number]).columns:
            display_df[c] = display_df[c].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "0.00")

        data = [display_df.columns.tolist()] + display_df.astype(str).values.tolist()
        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2F5597')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#F2F2F2')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        return t

    # 3. Add Tables
    if not party_df.empty:
        story.append(build_section_table(party_df, "Party-Wise Breakdown"))
        story.append(Spacer(1, 10))

    if not seller_df.empty:
        story.append(build_section_table(seller_df, "Salesperson-Wise Breakdown"))
        story.append(Spacer(1, 10))

    if not detail_df.empty:
        story.append(build_section_table(detail_df, "Detailed Excel Sheet Data"))

    doc.build(story)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# MAIN APP FLOW & CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Data Source")
uploaded_file = st.sidebar.file_uploader("Upload Sales Data (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    monthly_data, pending_df = load_all_sheets(uploaded_file)
    available_months = list(monthly_data.keys())
    all_months_df = pd.concat(monthly_data.values(), ignore_index=True) if monthly_data else pd.DataFrame()

    st.sidebar.markdown("---")
    st.sidebar.header("🗓️ Month & Comparison")

    selected_month = st.sidebar.selectbox("Select Primary Month:", available_months, index=len(available_months)-1 if available_months else 0)
    
    # Month Comparison Checkbox
    enable_comparison = st.sidebar.checkbox("Compare Between Two Months")
    comp_month = None
    if enable_comparison and len(available_months) > 1:
        comp_month = st.sidebar.selectbox("Select Month to Compare With:", [m for m in available_months if m != selected_month])

    chk_pending_dispatch = st.sidebar.checkbox("Show Pending Dispatch Section", value=True)

    # Load Active Month DataFrame
    raw_df = monthly_data.get(selected_month, pd.DataFrame()).copy()

    # -------------------------------------------------------------------------
    # GLOBAL FILTERS SECTION
    # -------------------------------------------------------------------------
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Dynamic Data Filters")

    filter_party = st.sidebar.multiselect("Filter by Party Name:", sorted(raw_df["PARTY NAME"].dropna().unique()) if "PARTY NAME" in raw_df else [])
    filter_seller = st.sidebar.multiselect("Filter by Sales Person:", sorted(raw_df["SELLER NAME"].dropna().unique()) if "SELLER NAME" in raw_df else [])
    filter_status = st.sidebar.multiselect("Filter by Status:", sorted(raw_df["STATUS"].dropna().unique()) if "STATUS" in raw_df else [])
    filter_item = st.sidebar.multiselect("Filter by Item Name:", sorted(raw_df["ITEM"].dropna().unique()) if "ITEM" in raw_df else [])
    filter_grade = st.sidebar.multiselect("Filter by Grade:", sorted(raw_df["GRADE"].dropna().unique()) if "GRADE" in raw_df else [])

    # Apply Filters
    filtered_df = raw_df.copy()
    if filter_party:
        filtered_df = filtered_df[filtered_df["PARTY NAME"].isin(filter_party)]
    if filter_seller:
        filtered_df = filtered_df[filtered_df["SELLER NAME"].isin(filter_seller)]
    if filter_status:
        filtered_df = filtered_df[filtered_df["STATUS"].isin(filter_status)]
    if filter_item:
        filtered_df = filtered_df[filtered_df["ITEM"].isin(filter_item)]
    if filter_grade:
        filtered_df = filtered_df[filtered_df["GRADE"].isin(filter_grade)]

    # -------------------------------------------------------------------------
    # 1. KPI SECTION
    # -------------------------------------------------------------------------
    st.markdown(f"## 📌 Key Metrics Summary - {selected_month}")

    metrics = {
        'po_count': filtered_df["PO NO"].nunique() if "PO NO" in filtered_df else 0,
        'do_count': filtered_df["DO NO"].nunique() if "DO NO" in filtered_df else 0,
        'parties_count': filtered_df["PARTY NAME"].nunique() if "PARTY NAME" in filtered_df else 0,
        'po_qty': filtered_df["PO QTY (MT)"].sum() if "PO QTY (MT)" in filtered_df else 0,
        'revenue': filtered_df["TOTAL REVENUE"].sum() if "TOTAL REVENUE" in filtered_df else 0,
        'disp_qty': filtered_df["DISP.QTY"].sum() if "DISP.QTY" in filtered_df else 0,
        'pending_qty': filtered_df["PENDING"].sum() if "PENDING" in filtered_df else 0
    }

    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    k1.metric("Overall POs", f"{metrics['po_count']}")
    k2.metric("Overall DOs", f"{metrics['do_count']}")
    k3.metric("Parties", f"{metrics['parties_count']}")
    k4.metric("PO Qty (MT)", f"{metrics['po_qty']:,.2f}")
    k5.metric("Amount (₹)", f"₹{metrics['revenue']:,.2f}")
    k6.metric("Dispatched Qty", f"{metrics['disp_qty']:,.2f}")
    k7.metric("Pending Qty", f"{metrics['pending_qty']:,.2f}")

    # Comparison KPIs Block
    if enable_comparison and comp_month and comp_month in monthly_data:
        st.markdown(f"### 🔄 Comparison with {comp_month}")
        c_df = monthly_data[comp_month]
        c_po_qty = c_df["PO QTY (MT)"].sum() if "PO QTY (MT)" in c_df else 0
        c_disp_qty = c_df["DISP.QTY"].sum() if "DISP.QTY" in c_df else 0
        c_pending_qty = c_df["PENDING"].sum() if "PENDING" in c_df else 0
        c_rev = c_df["TOTAL REVENUE"].sum() if "TOTAL REVENUE" in c_df else 0

        cm1, cm2, cm3, cm4 = st.columns(4)
        cm1.metric(f"PO Qty ({comp_month})", f"{c_po_qty:,.2f}", f"{(metrics['po_qty'] - c_po_qty):,.2f}")
        cm2.metric(f"Dispatched ({comp_month})", f"{c_disp_qty:,.2f}", f"{(metrics['disp_qty'] - c_disp_qty):,.2f}")
        cm3.metric(f"Pending ({comp_month})", f"{c_pending_qty:,.2f}", f"{(metrics['pending_qty'] - c_pending_qty):,.2f}")
        cm4.metric(f"Amount ({comp_month})", f"₹{c_rev:,.2f}", f"₹{(metrics['revenue'] - c_rev):,.2f}")

    st.markdown("---")

    # -------------------------------------------------------------------------
    # 2. GRAPH SECTION
    # -------------------------------------------------------------------------
    st.markdown("## 📊 Sales Analytics Visualizations")

    fig, ax = plt.subplots(1, 2, figsize=(14, 4))
    
    # Chart 1: Dispatched vs Pending Qty
    status_summary = pd.DataFrame({
        'Status': ['Dispatched Qty', 'Pending Qty'],
        'MT Volume': [metrics['disp_qty'], metrics['pending_qty']]
    })
    sns.barplot(data=status_summary, x='Status', y='MT Volume', ax=ax[0], palette=['#2ca02c', '#d62728'])
    ax[0].set_title("Dispatched vs Pending Qty (MT)")
    ax[0].set_ylabel("Metric Tons (MT)")

    # Chart 2: Top 5 Parties Volume
    if "PARTY NAME" in filtered_df and not filtered_df.empty:
        top_parties = filtered_df.groupby("PARTY NAME")["PO QTY (MT)"].sum().nlargest(5).reset_index()
        sns.barplot(data=top_parties, x="PO QTY (MT)", y="PARTY NAME", ax=ax[1], palette="Blues_r")
        ax[1].set_title("Top 5 Parties by Order Volume (MT)")
        ax[1].set_xlabel("PO Qty (MT)")

    plt.tight_layout()
    st.pyplot(fig)

    # Save Chart Buffer for PDF Exporter
    chart_buf = io.BytesIO()
    fig.savefig(chart_buf, format="png", dpi=150, bbox_inches='tight')

    st.markdown("---")

    # -------------------------------------------------------------------------
    # 3. DETAILED EXCEL SECTION
    # -------------------------------------------------------------------------
    st.markdown(f"## 📋 Detailed Excel Sheet Data - {selected_month}")
    st.dataframe(filtered_df, use_container_width=True)

    st.markdown("---")

    # -------------------------------------------------------------------------
    # 4. PARTY-WISE DETAILED SECTION
    # -------------------------------------------------------------------------
    st.markdown("## 🏢 Party Wise Detail")
    party_cols = ["PARTY NAME", "PO QTY (MT)", "PENDING", "DISP.QTY", "SELLER NAME", "BROKER", "STATUS"]
    avail_party_cols = [c for c in party_cols if c in filtered_df.columns]
    
    if avail_party_cols:
        party_detail = filtered_df.groupby(["PARTY NAME"] + [c for c in ["SELLER NAME", "BROKER", "STATUS"] if c in filtered_df.columns]).agg({
            "PO QTY (MT)": "sum" if "PO QTY (MT)" in filtered_df else "count",
            "PENDING": "sum" if "PENDING" in filtered_df else "count",
            "DISP.QTY": "sum" if "DISP.QTY" in filtered_df else "count"
        }).reset_index()
        
        # Rename for presentation
        party_detail.rename(columns={
            "PO QTY (MT)": "Ordered Qty (MT)",
            "PENDING": "Pending Qty (MT)",
            "DISP.QTY": "Dispatched Qty (MT)",
            "SELLER NAME": "Sales Person",
            "PARTY NAME": "Party Name"
        }, inplace=True)

        st.dataframe(party_detail, use_container_width=True)
    else:
        party_detail = pd.DataFrame()

    # -------------------------------------------------------------------------
    # 5. SALESPERSON-WISE DETAILED SECTION
    # -------------------------------------------------------------------------
    st.markdown("## 👨‍💼 Sales Person Wise Report")
    if "SELLER NAME" in filtered_df.columns:
        seller_detail = filtered_df.groupby(["SELLER NAME"] + ([ "STATUS" ] if "STATUS" in filtered_df.columns else [])).agg({
            "PO QTY (MT)": "sum" if "PO QTY (MT)" in filtered_df else "count",
            "DISP.QTY": "sum" if "DISP.QTY" in filtered_df else "count",
            "PENDING": "sum" if "PENDING" in filtered_df else "count"
        }).reset_index()

        seller_detail.rename(columns={
            "SELLER NAME": "Sales Person Name",
            "PO QTY (MT)": "Ordered Qty (MT)",
            "DISP.QTY": "Dispatched Qty (MT)",
            "PENDING": "Pending Qty (MT)"
        }, inplace=True)

        st.dataframe(seller_detail, use_container_width=True)
    else:
        seller_detail = pd.DataFrame()

    # -------------------------------------------------------------------------
    # 6. DOWNLOAD PDF SECTION (KPIS + CHART + DETAILS)
    # -------------------------------------------------------------------------
    st.markdown("---")
    st.markdown("## 📥 Download Complete PDF Report")
    
    full_pdf_buf = generate_full_pdf_report(
        month_name=selected_month,
        metrics=metrics,
        chart_buf=chart_buf,
        detail_df=filtered_df,
        party_df=party_detail,
        seller_df=seller_detail
    )

    st.download_button(
        label=f"📄 Download Complete {selected_month} PDF Report (KPIs, Graphs & Detailed Sections)",
        data=full_pdf_buf,
        file_name=f"Full_Sales_Report_{selected_month}.pdf",
        mime="application/pdf"
    )

    # -------------------------------------------------------------------------
    # 7. OPTIONAL PENDING DISPATCH SECTION
    # -------------------------------------------------------------------------
    if chk_pending_dispatch and not pending_df.empty:
        st.markdown("<br><hr style='border:2px solid red;'><br>", unsafe_allow_html=True)
        st.header("🔴 PENDING DISPATCH REPORT & DASHBOARD")

        p_col1, p_col2, p_col3, p_col4 = st.columns(4)
        actual_pending_qty = pending_df["PENDING"].sum() if "PENDING" in pending_df else pending_df["PO QTY (MT)"].sum()
        
        p_col1.metric("Pending Orders Count", f"{len(pending_df)}")
        p_col2.metric("Total Pending Qty (MT)", f"{actual_pending_qty:,.2f}")
        p_col3.metric("Total Revenue (₹)", f"₹{pending_df['TOTAL REVENUE'].sum():,.2f}")
        p_col4.metric("Affected Parties", f"{pending_df['PARTY NAME'].nunique() if 'PARTY NAME' in pending_df else 0}")

        st.dataframe(pending_df, use_container_width=True)

else:
    st.info("👈 Upload your `Sales Data.xlsx` workbook in the left sidebar to generate the interactive dashboard and PDF download.")
