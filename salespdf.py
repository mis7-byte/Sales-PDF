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
st.set_page_config(page_title="Sales Analytics & PDF Exporter", layout="wide", initial_sidebar_state="expanded")

st.title("📊 Monthly & Overall Sales Analytics Dashboard")

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

    # Parse Width and Length from SIZE column
    def parse_size(size_str):
        if pd.isna(size_str) or str(size_str).strip() == "":
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

# PDF Generation Helper
def generate_pdf_report(df_summary, title="Sales Data Report"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=15, leftMargin=15, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{title}</b>", styles['Title']))
    story.append(Spacer(1, 10))

    cols_to_include = [c for c in df_summary.columns if c not in ["MONTH_SHEET", "S_NO", "MOBILE NO"]]
    pdf_df = df_summary[cols_to_include].copy()

    # Format numbers nicely
    for c in pdf_df.select_dtypes(include=[np.number]).columns:
        pdf_df[c] = pdf_df[c].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "0.00")

    data = [pdf_df.columns.tolist()] + pdf_df.astype(str).values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 6.5),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f9f9f9')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))

    story.append(table)
    doc.build(story)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# SIDEBAR CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Data Source & Month Selector")
uploaded_file = st.sidebar.file_uploader("Upload Excel File (Sales Data.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    monthly_data, pending_df = load_all_sheets(uploaded_file)
    available_months = list(monthly_data.keys())

    # Combine all month sheets for overall calculations
    all_months_df = pd.concat(monthly_data.values(), ignore_index=True) if monthly_data else pd.DataFrame()

    st.sidebar.markdown("---")
    st.sidebar.header("🗓️ Select View Mode")

    view_mode = st.sidebar.radio(
        "Select Dashboard View:",
        ["Monthly Detail View", "Overall Party-Wise Summary", "Overall Salesperson Summary"]
    )

    if view_mode == "Monthly Detail View":
        selected_month = st.sidebar.selectbox("Select Primary Month:", available_months, index=len(available_months)-1 if available_months else 0)

        # Secondary Month Comparison
        chk_compare_months = st.sidebar.checkbox("Compare Between Two Months")
        compare_month = None
        if chk_compare_months:
            compare_options = [m for m in available_months if m != selected_month]
            compare_month = st.sidebar.selectbox("Select Second Month to Compare:", compare_options)

        chk_pending_dispatch = st.sidebar.checkbox("Show Pending Dispatch Section", value=True)

        raw_df = monthly_data.get(selected_month, pd.DataFrame())

        # Dynamic Filters
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 Dynamic Data Filters")

        filtered_df = raw_df.copy()

        if "STATUS" in raw_df.columns:
            status_list = ["ALL"] + sorted([s for s in raw_df["STATUS"].dropna().unique() if str(s).strip() not in ["NAN", ""]])
            selected_status = st.sidebar.selectbox("Filter by Order Status:", status_list)
            if selected_status != "ALL":
                filtered_df = filtered_df[filtered_df["STATUS"] == selected_status]

        if "BROKER" in raw_df.columns:
            broker_options = ["ALL", "DIRECT ORDERS ONLY", "BROKER ORDERS ONLY"]
            selected_broker_type = st.sidebar.selectbox("Filter Broker vs Direct:", broker_options)
            if selected_broker_type == "DIRECT ORDERS ONLY":
                filtered_df = filtered_df[filtered_df["BROKER"].isin(["DIRECT", "NONE", "NAN", "", "N/A"])]
            elif selected_broker_type == "BROKER ORDERS ONLY":
                filtered_df = filtered_df[~filtered_df["BROKER"].isin(["DIRECT", "NONE", "NAN", "", "N/A"])]

        # Key Metrics
        st.markdown(f"## 📌 Key Metrics - {selected_month}")

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        sum_po_qty = filtered_df["PO QTY (MT)"].sum() if "PO QTY (MT)" in filtered_df else 0
        sum_disp_qty = filtered_df["DISP.QTY"].sum() if "DISP.QTY" in filtered_df else 0
        sum_pending_qty = filtered_df["PENDING"].sum() if "PENDING" in filtered_df else 0
        total_rev = filtered_df["TOTAL REVENUE"].sum() if "TOTAL REVENUE" in filtered_df else 0

        kpi1.metric("Total Order Qty (MT)", f"{sum_po_qty:,.2f}")
        kpi2.metric("Dispatched Qty (MT)", f"{sum_disp_qty:,.2f}")
        kpi3.metric("Pending Qty (MT)", f"{sum_pending_qty:,.2f}")
        kpi4.metric("Total Revenue (₹)", f"₹{total_rev:,.2f}")

        # Comparison Section
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

        # Data View & Export
        st.markdown("---")
        st.header(f"📋 Data View - {selected_month}")
        st.dataframe(filtered_df, use_container_width=True)

        st.markdown("### 📥 Download PDF Report")
        pdf_buf = generate_pdf_report(filtered_df, title=f"Sales Report - {selected_month}")
        st.download_button(
            label=f"📄 Download PDF Report ({selected_month})",
            data=pdf_buf,
            file_name=f"Sales_Report_{selected_month}.pdf",
            mime="application/pdf"
        )

        if chk_pending_dispatch and not pending_df.empty:
            st.markdown("<br><hr style='border:2px solid red;'><br>", unsafe_allow_html=True)
            st.header("🔴 PENDING DISPATCH REPORT & DASHBOARD")

            p_col1, p_col2, p_col3, p_col4 = st.columns(4)
            p_col1.metric("Pending Orders Count", len(pending_df))
            p_col2.metric("Total Pending Qty (MT)", f"{pending_df['PENDING'].sum():,.2f}")
            p_col3.metric("Total Revenue (₹)", f"₹{pending_df['TOTAL REVENUE'].sum():,.2f}")
            p_col4.metric("Affected Parties", pending_df["PARTY NAME"].nunique() if "PARTY NAME" in pending_df else 0)

            st.dataframe(pending_df, use_container_width=True)

            pending_pdf_buf = generate_pdf_report(pending_df, title="Pending Dispatch Report")
            st.download_button(
                label="📄 Download Pending Dispatch PDF",
                data=pending_pdf_buf,
                file_name="Pending_Dispatch_Report.pdf",
                mime="application/pdf"
            )

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

            p_col1, p_col2, p_col3 = st.columns(3)
            p_col1.metric("Total Unique Parties", len(party_summary))
            p_col2.metric("Top Party Volume (MT)", f"{party_summary['Total_PO_Qty_MT'].max():,.2f}")
            p_col3.metric("Total Overall Revenue (₹)", f"₹{party_summary['Total_Revenue_INR'].sum():,.2f}")

            st.markdown("---")
            st.subheader("📊 Top 10 Parties by Order Quantity")
            fig, ax = plt.subplots(figsize=(10, 4))
            sns.barplot(data=party_summary.head(10), x="Total_PO_Qty_MT", y="PARTY NAME", ax=ax, palette="Blues_r")
            ax.set_xlabel("Total PO Quantity (MT)")
            ax.set_ylabel("Party Name")
            st.pyplot(fig)

            st.markdown("---")
            st.subheader("📋 Party Summary Table")
            st.dataframe(party_summary, use_container_width=True)

            st.markdown("### 📥 Download Party Summary PDF")
            party_pdf = generate_pdf_report(party_summary, title="Overall Party-Wise Sales Summary")
            st.download_button(
                label="📄 Download Overall Party Summary PDF",
                data=party_pdf,
                file_name="Overall_Party_Summary.pdf",
                mime="application/pdf"
            )

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

            s_col1, s_col2, s_col3 = st.columns(3)
            s_col1.metric("Active Salespersons", len(seller_summary))
            s_col2.metric("Top Salesperson Volume (MT)", f"{seller_summary['Total_PO_Qty_MT'].max():,.2f}")
            s_col3.metric("Total Overall Revenue (₹)", f"₹{seller_summary['Total_Revenue_INR'].sum():,.2f}")

            st.markdown("---")
            st.subheader("📊 Salesperson Performance Breakdown")
            fig, ax = plt.subplots(figsize=(10, 4))
            sns.barplot(data=seller_summary, x="SELLER NAME", y="Total_PO_Qty_MT", ax=ax, palette="viridis")
            plt.xticks(rotation=45)
            ax.set_ylabel("Total PO Quantity (MT)")
            st.pyplot(fig)

            st.markdown("---")
            st.subheader("📋 Salesperson Summary Table")
            st.dataframe(seller_summary, use_container_width=True)

            st.markdown("### 📥 Download Salesperson Summary PDF")
            seller_pdf = generate_pdf_report(seller_summary, title="Overall Salesperson Performance Summary")
            st.download_button(
                label="📄 Download Overall Salesperson Summary PDF",
                data=seller_pdf,
                file_name="Overall_Salesperson_Summary.pdf",
                mime="application/pdf"
            )

else:
    st.info("👈 Please upload your `Sales Data.xlsx` workbook in the left sidebar to get started.")
