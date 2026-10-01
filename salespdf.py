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
st.set_page_config(
    page_title="Sales Analytics & PDF Exporter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Monthly & Overall Sales Analytics Dashboard")

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS & DATA PROCESSING
# -----------------------------------------------------------------------------
def process_dataframe(df, sheet_name, is_pending_sheet=False):
    """Clean, standardize column names, filter summary rows, and compute revenue."""
    if df.empty:
        return df

    df = df.dropna(how='all').copy()

    # Column name cleaning & standardization
    rename_map = {
        'S. NO.': 'S_NO', 'SR NO': 'S_NO', 'SR NO ': 'S_NO', 'S.NO': 'S_NO',
        'DO .NO.': 'DO NO', 'DO NO ': 'DO NO', 'D.O. NO.': 'DO NO',
        'THIKNESS': 'THICKNESS', 'SIZE (MM)': 'SIZE', 'SIZE ()': 'SIZE',
        'DISCRIPTION': 'ITEM', 'DESCRIPTION': 'ITEM',
        'REMARKS': 'REMARK', 'MOB NO.': 'MOBILE NO', 'MOBILE': 'MOBILE NO'
    }
    df.rename(columns=rename_map, inplace=True)

    # Filter out bottom total/summary rows from Excel sheets
    if "PARTY NAME" in df.columns:
        df = df[~df["PARTY NAME"].astype(str).str.upper().str.contains("TOTAL|SUM|GRAND TOTAL|AVERAGE", na=False)]
    if "S_NO" in df.columns:
        df = df[df["S_NO"].notna()]

    # Clean numerical fields
    num_cols = ["PO QTY (MT)", "PER TON", "DISP.QTY", "PENDING"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Compute Revenue based on Sheet Type
    qty_col = "PENDING" if (is_pending_sheet and "PENDING" in df.columns) else "PO QTY (MT)"
    if qty_col in df.columns and "PER TON" in df.columns:
        df["TOTAL REVENUE"] = df[qty_col] * df["PER TON"]
    elif "PO QTY (MT)" in df.columns and "PER TON" in df.columns:
        df["TOTAL REVENUE"] = df["PO QTY (MT)"] * df["PER TON"]
    else:
        df["TOTAL REVENUE"] = 0.0

    # Format text columns
    str_cols = ["PO NO", "DO NO", "PARTY NAME", "BROKER", "SECTOR", "PLACE", "SELLER NAME", "ITEM", "GRADE", "STATUS", "SIZE"]
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()

    df["MONTH_SHEET"] = sheet_name
    return df

@st.cache_data
def load_all_sheets(uploaded_file):
    """Load and parse monthly tabs and Pending Dispatch tab independently."""
    xl = pd.ExcelFile(uploaded_file)
    all_sheets = xl.sheet_names
    
    monthly_data = {}
    pending_df = pd.DataFrame()

    for sheet in all_sheets:
        clean_sheet_name = sheet.strip().upper()
        
        # Detect and Process Pending Dispatch Sheet
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
            # Process Monthly Sales Sheets
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
# PDF REPORT GENERATOR
# -----------------------------------------------------------------------------
def generate_pdf_report(df_summary, title="Sales Data Report"):
    """Generate a clean PDF document directly matching the current view table."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=15,
        leftMargin=15,
        topMargin=20,
        bottomMargin=20
    )
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{title}</b>", styles['Title']))
    story.append(Spacer(1, 10))

    # Clean columns for PDF table rendering
    cols_to_include = [c for c in df_summary.columns if c not in ["MONTH_SHEET", "MOBILE NO"]]
    pdf_df = df_summary[cols_to_include].copy()

    for c in pdf_df.select_dtypes(include=[np.number]).columns:
        pdf_df[c] = pdf_df[c].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "0.00")

    data = [pdf_df.columns.tolist()] + pdf_df.astype(str).values.tolist()
    
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f9f9f9')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))

    story.append(table)
    doc.build(story)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# MAIN APP & SIDEBAR CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Data Source & Controls")
uploaded_file = st.sidebar.file_uploader("Upload Sales Data Workbook (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    monthly_data, pending_df = load_all_sheets(uploaded_file)
    available_months = list(monthly_data.keys())
    all_months_df = pd.concat(monthly_data.values(), ignore_index=True) if monthly_data else pd.DataFrame()

    st.sidebar.markdown("---")
    st.sidebar.header("🗓️ Dashboard Views")

    view_mode = st.sidebar.radio(
        "Select Report View:",
        ["Monthly Detail View", "Overall Party-Wise Summary", "Overall Salesperson Summary"]
    )

    # -------------------------------------------------------------------------
    # VIEW 1: MONTHLY DETAIL VIEW
    # -------------------------------------------------------------------------
    if view_mode == "Monthly Detail View":
        selected_month = st.sidebar.selectbox(
            "Select Primary Month:",
            available_months,
            index=len(available_months)-1 if available_months else 0
        )

        compare_months = st.sidebar.checkbox("Compare Between Two Months")
        second_month = None
        if compare_months and len(available_months) > 1:
            second_month = st.sidebar.selectbox("Select Secondary Month for Comparison:", available_months, index=0)

        chk_pending_dispatch = st.sidebar.checkbox("Show Pending Dispatch Section", value=True)

        raw_df = monthly_data.get(selected_month, pd.DataFrame())

        st.markdown(f"## 📌 Key Metrics - {selected_month}")
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        
        sum_po_qty = raw_df["PO QTY (MT)"].sum() if "PO QTY (MT)" in raw_df else 0
        sum_disp_qty = raw_df["DISP.QTY"].sum() if "DISP.QTY" in raw_df else 0
        sum_pending_qty = raw_df["PENDING"].sum() if "PENDING" in raw_df else 0
        total_rev = raw_df["TOTAL REVENUE"].sum() if "TOTAL REVENUE" in raw_df else 0

        kpi1.metric("Total Order Qty (MT)", f"{sum_po_qty:,.2f}")
        kpi2.metric("Dispatched Qty (MT)", f"{sum_disp_qty:,.2f}")
        kpi3.metric("Pending Qty (MT)", f"{sum_pending_qty:,.2f}")
        kpi4.metric("Total Revenue (₹)", f"₹{total_rev:,.2f}")

        if second_month and second_month in monthly_data:
            st.markdown(f"### 🔄 Comparison with {second_month}")
            df2 = monthly_data[second_month]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(f"Total Qty ({second_month})", f"{df2['PO QTY (MT)'].sum():,.2f}", f"{(sum_po_qty - df2['PO QTY (MT)'].sum()):,.2f}")
            c2.metric(f"Dispatched ({second_month})", f"{df2['DISP.QTY'].sum():,.2f}", f"{(sum_disp_qty - df2['DISP.QTY'].sum()):,.2f}")
            c3.metric(f"Pending ({second_month})", f"{df2['PENDING'].sum():,.2f}", f"{(sum_pending_qty - df2['PENDING'].sum()):,.2f}")
            c4.metric(f"Revenue ({second_month})", f"₹{df2['TOTAL REVENUE'].sum():,.2f}", f"₹{(total_rev - df2['TOTAL REVENUE'].sum()):,.2f}")

        st.markdown("---")
        st.header(f"📋 Sales Report Table - {selected_month}")
        st.dataframe(raw_df, use_container_width=True)

        # Download PDF for current monthly view
        pdf_buf = generate_pdf_report(raw_df, title=f"Sales Report - {selected_month}")
        st.download_button(
            label=f"📄 Download {selected_month} Report PDF",
            data=pdf_buf,
            file_name=f"Sales_Report_{selected_month}.pdf",
            mime="application/pdf"
        )

        # PENDING DISPATCH SECTION
        if chk_pending_dispatch and not pending_df.empty:
            st.markdown("<br><hr style='border:2px solid red;'><br>", unsafe_allow_html=True)
            st.header("🔴 PENDING DISPATCH REPORT & DASHBOARD")

            p_col1, p_col2, p_col3, p_col4 = st.columns(4)
            
            actual_pending_qty = pending_df["PENDING"].sum() if "PENDING" in pending_df else pending_df["PO QTY (MT)"].sum()
            actual_pending_orders = len(pending_df)
            actual_revenue = pending_df["TOTAL REVENUE"].sum()
            affected_parties = pending_df["PARTY NAME"].nunique() if "PARTY NAME" in pending_df else 0

            p_col1.metric("Pending Orders Count", f"{actual_pending_orders}")
            p_col2.metric("Total Pending Qty (MT)", f"{actual_pending_qty:,.2f}")
            p_col3.metric("Total Revenue (₹)", f"₹{actual_revenue:,.2f}")
            p_col4.metric("Affected Parties", f"{affected_parties}")

            st.dataframe(pending_df, use_container_width=True)

            pending_pdf_buf = generate_pdf_report(pending_df, title="Pending Dispatch Report")
            st.download_button(
                label="📄 Download Pending Dispatch PDF Report",
                data=pending_pdf_buf,
                file_name="Pending_Dispatch_Report.pdf",
                mime="application/pdf"
            )

    # -------------------------------------------------------------------------
    # VIEW 2: OVERALL PARTY-WISE SUMMARY
    # -------------------------------------------------------------------------
    elif view_mode == "Overall Party-Wise Summary":
        st.header("🏢 Overall Party-Wise Sales Summary")
        if "PARTY NAME" in all_months_df.columns:
            party_summary = all_months_df.groupby("PARTY NAME").agg(
                Total_Orders=("PO NO", "nunique"),
                Total_PO_Qty_MT=("PO QTY (MT)", "sum"),
                Total_Dispatched_MT=("DISP.QTY", "sum"),
                Total_Pending_MT=("PENDING", "sum"),
                Total_Revenue_INR=("TOTAL REVENUE", "sum")
            ).reset_index().sort_values(by="Total_PO_Qty_MT", ascending=False)

            st.dataframe(party_summary, use_container_width=True)

            # Chart Top 10 Parties
            st.subheader("📊 Top 10 Parties by Order Volume (MT)")
            fig, ax = plt.subplots(figsize=(10, 4))
            top_10 = party_summary.head(10)
            sns.barplot(data=top_10, x="Total_PO_Qty_MT", y="PARTY NAME", palette="Blues_r", ax=ax)
            ax.set_title("Top 10 Parties by Volume")
            ax.set_xlabel("Quantity (MT)")
            st.pyplot(fig)

            party_pdf_buf = generate_pdf_report(party_summary, title="Overall Party-Wise Sales Summary")
            st.download_button(
                label="📄 Download Party-Wise Summary PDF",
                data=party_pdf_buf,
                file_name="Party_Wise_Sales_Summary.pdf",
                mime="application/pdf"
            )

    # -------------------------------------------------------------------------
    # VIEW 3: OVERALL SALESPERSON SUMMARY
    # -------------------------------------------------------------------------
    elif view_mode == "Overall Salesperson Summary":
        st.header("👨‍💼 Overall Salesperson Performance Summary")
        if "SELLER NAME" in all_months_df.columns:
            seller_summary = all_months_df.groupby("SELLER NAME").agg(
                Total_Orders=("PO NO", "nunique"),
                Total_Parties=("PARTY NAME", "nunique"),
                Total_PO_Qty_MT=("PO QTY (MT)", "sum"),
                Total_Dispatched_MT=("DISP.QTY", "sum"),
                Total_Pending_MT=("PENDING", "sum"),
                Total_Revenue_INR=("TOTAL REVENUE", "sum")
            ).reset_index().sort_values(by="Total_PO_Qty_MT", ascending=False)

            st.dataframe(seller_summary, use_container_width=True)

            # Chart Performance
            st.subheader("📈 Salesperson Volume Breakdown (MT)")
            fig, ax = plt.subplots(figsize=(10, 4))
            sns.barplot(data=seller_summary, x="Total_PO_Qty_MT", y="SELLER NAME", palette="Greens_r", ax=ax)
            ax.set_title("Salesperson Volume Contribution")
            ax.set_xlabel("Quantity (MT)")
            st.pyplot(fig)

            seller_pdf_buf = generate_pdf_report(seller_summary, title="Overall Salesperson Performance Summary")
            st.download_button(
                label="📄 Download Salesperson Summary PDF",
                data=seller_pdf_buf,
                file_name="Salesperson_Performance_Summary.pdf",
                mime="application/pdf"
            )
else:
    st.info("👈 Please upload your `Sales Data.xlsx` workbook in the left sidebar to generate the analytics and PDF reports.")
