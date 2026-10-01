import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import re
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="Sales & Dispatch Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Sales & Dispatch Analytics Dashboard")

# ---------------------------------------------------------
# Flexible Excel Header & Column Normalization
# ---------------------------------------------------------
EXPECTED_COLUMNS = {
    'S_NO': ['s_no', 's.no', 'sr no', 'serial no', 'sno', 's_no.'],
    'PO NO': ['po no', 'po.no', 'po number', 'pono', 'po_no'],
    'DO NO': ['do .no.', 'do no', 'do.no', 'do_no', 'dono'],
    'PO DATE': ['po date', 'podate', 'po_date', 'order date'],
    'PARTY NAME': ['party name', 'party_name', 'customer name', 'party'],
    'BROKER': ['broker'],
    'SECTOR': ['sector'],
    'PLACE': ['place', 'city', 'location'],
    'SELLER NAME': ['seller name', 'seller_name', 'sales person', 'salesperson', 'sales executive'],
    'ITEM': ['item', 'item name', 'item_name', 'product'],
    'THICKNESS': ['thikness', 'thickness', 'thk', 'thk.'],
    'SIZE': ['size'],
    'GRADE': ['grade'],
    'PO QTY (MT)': ['po qty (mt)', 'po qty', 'po_qty', 'ordered qty', 'order qty'],
    'PER TON': ['per ton', 'rate', 'price/ton', 'rate per ton'],
    'INV NO': ['inv no.', 'inv no', 'invoice no', 'inv_no'],
    'DATE': ['date', 'disp date', 'dispatch date', 'inv date'],
    'DISP.QTY': ['disp.qty', 'disp qty', 'dispatched qty', 'disp_qty'],
    'PENDING': ['pending', 'pending qty', 'bal qty'],
    'PAYMENT': ['payment', 'payment mode', 'terms'],
    'DISP. TH.': ['disp. th.', 'disp th', 'dispatch through', 'disp_th'],
    'STATUS': ['status', 'order status'],
    'REMARK': ['remark', 'remarks'],
    'MOBILE NO': ['mobile no', 'mobile', 'phone']
}

def load_and_clean_sheet(file_bytes, sheet_name):
    # Detect header row index within first 10 rows
    df_raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)
    
    header_row_idx = 0
    max_matches = 0
    
    for row_idx in range(min(10, len(df_raw))):
        row_values = df_raw.iloc[row_idx].astype(str).str.strip().str.lower().tolist()
        matches = 0
        for std_col, aliases in EXPECTED_COLUMNS.items():
            if any(alias in row_values for alias in aliases):
                matches += 1
        if matches > max_matches:
            max_matches = matches
            header_row_idx = row_idx
            
    # Read sheet starting from detected header index
    df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Standardize column headers
    renamed_cols = {}
    for col in df.columns:
        col_lower = str(col).strip().lower()
        matched = False
        for std_col, aliases in EXPECTED_COLUMNS.items():
            if col_lower in aliases or any(a in col_lower for a in aliases):
                renamed_cols[col] = std_col
                matched = True
                break
        if not matched:
            renamed_cols[col] = col

    df.rename(columns=renamed_cols, inplace=True)
    
    # Ensure all expected columns exist
    for std_col in EXPECTED_COLUMNS.keys():
        if std_col not in df.columns:
            df[std_col] = np.nan

    # Datetime conversions
    df['PO DATE'] = pd.to_datetime(df['PO DATE'], errors='coerce')
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    
    # Numeric sanitization
    numeric_cols = ['PO QTY (MT)', 'PER TON', 'DISP.QTY', 'PENDING']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
    df['AMOUNT'] = df['PO QTY (MT)'] * df['PER TON']
    
    # Extract Width from Size string (e.g., 1500X6300 -> 1500)
    def parse_width(size_val):
        if pd.isna(size_val):
            return "N/A"
        match = re.search(r'(\d+)\s*[xX*]\s*(\d+)', str(size_val))
        if match:
            return match.group(1)
        num = re.findall(r'\d+', str(size_val))
        return num[0] if num else str(size_val)

    df['WIDTH'] = df['SIZE'].apply(parse_width)
    
    # String conversions & null handling
    str_cols = ['PARTY NAME', 'SELLER NAME', 'ITEM', 'STATUS', 'REMARK', 'BROKER', 'SECTOR', 'PLACE', 'PO NO', 'DO NO']
    for c in str_cols:
        df[c] = df[c].fillna('Unknown').astype(str).str.strip()

    return df

# ---------------------------------------------------------
# PDF Generator Tool
# ---------------------------------------------------------
def generate_pdf_report(month_name, kpis, tables_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#1E3A8A'), spaceAfter=12
    )
    section_style = ParagraphStyle(
        'DocSection', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#1E40AF'), spaceBefore=10, spaceAfter=6
    )
    
    story.append(Paragraph(f"<b>Sales & Dispatch Summary Report - {month_name}</b>", title_style))
    story.append(Spacer(1, 10))
    
    story.append(Paragraph("Key Performance Indicators (KPIs)", section_style))
    kpi_data = [
        ["Metric", "Value"],
        ["Total PO Count", str(kpis['total_po'])],
        ["Total DO Count", str(kpis['total_do'])],
        ["Number of Parties", str(kpis['num_parties'])],
        ["Total PO Quantity (MT)", f"{kpis['total_po_qty']:,.2f}"],
        ["Total PO Amount (₹)", f"₹{kpis['total_amount']:,.2f}"],
        ["Dispatched Quantity (MT)", f"{kpis['dispatched_qty']:,.2f}"],
        ["Pending Quantity (MT)", f"{kpis['pending_qty']:,.2f}"]
    ]
    t_kpi = Table(kpi_data, colWidths=[220, 220])
    t_kpi.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#3B82F6')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#F3F4F6')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 15))
    
    for title, df in tables_dict.items():
        if df is not None and not df.empty:
            story.append(Paragraph(f"<b>{title}</b>", section_style))
            sub_df = df.head(15).reset_index()
            table_data = [sub_df.columns.tolist()] + sub_df.values.tolist()
            table_data = [[str(cell)[:25] for cell in row] for row in table_data]
            
            t_data = Table(table_data)
            t_data.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4B5563')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
            ]))
            story.append(t_data)
            story.append(Spacer(1, 12))
            
    doc.build(story)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# Sidebar Navigation & Upload
# ---------------------------------------------------------
st.sidebar.title("📌 Navigation & Controls")
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("👈 Please upload an Excel workbook from the sidebar to view metrics.")
    st.stop()

xl = pd.ExcelFile(uploaded_file)
sheet_names = xl.sheet_names

selected_sheet = st.sidebar.selectbox("Select Month / Sheet", sheet_names)

section = st.sidebar.radio("Go to Section", [
    "📈 KPI Overview",
    "👤 Sales Person Analytics",
    "📅 Month Wise & Date Filter",
    "🏢 Party Wise Report",
    "📐 Material & Width Breakdown"
])

df = load_and_clean_sheet(uploaded_file, selected_sheet)

def calculate_kpis(data):
    total_po = data['PO NO'].replace('Unknown', np.nan).dropna().nunique()
    total_do = data['DO NO'].replace('Unknown', np.nan).dropna().nunique()
    num_parties = data['PARTY NAME'].replace('Unknown', np.nan).dropna().nunique()
    total_po_qty = data['PO QTY (MT)'].sum()
    total_amount = data['AMOUNT'].sum()
    dispatched_qty = data['DISP.QTY'].sum()
    pending_qty = data['PENDING'].sum()
    
    return {
        'total_po': total_po,
        'total_do': total_do,
        'num_parties': num_parties,
        'total_po_qty': total_po_qty,
        'total_amount': total_amount,
        'dispatched_qty': dispatched_qty,
        'pending_qty': pending_qty
    }

# =========================================================
# SECTION 1: KPI OVERVIEW
# =========================================================
if section == "📈 KPI Overview":
    st.header(f"KPI Overview — {selected_sheet}")
    
    kpis = calculate_kpis(df)
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Overall PO Count", f"{kpis['total_po']:,}")
    c2.metric("Overall DO Count", f"{kpis['total_do']:,}")
    c3.metric("Number of Parties", f"{kpis['num_parties']:,}")
    c4.metric("Total PO Qty (MT)", f"{kpis['total_po_qty']:,.2f}")
    
    c5, c6, c7 = st.columns(3)
    c5.metric("Total Amount (PO Qty × Rate)", f"₹{kpis['total_amount']:,.2f}")
    c6.metric("Sum of Dispatched Qty (MT)", f"{kpis['dispatched_qty']:,.2f}")
    c7.metric("Sum of Pending Qty (MT)", f"{kpis['pending_qty']:,.2f}")
    
    st.markdown("---")
    st.subheader("Dispatched vs Pending Overview")
    
    fig_kpi = go.Figure(data=[go.Pie(
        labels=['Dispatched Qty', 'Pending Qty'],
        values=[kpis['dispatched_qty'], kpis['pending_qty']],
        hole=.4,
        marker_colors=['#10B981', '#EF4444']
    )])
    st.plotly_chart(fig_kpi, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📊 Data Table View")
    st.dataframe(df, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📄 Export Report")
    
    sp_summary = df.groupby('SELLER NAME').agg(
        Ordered=('PO QTY (MT)', 'sum'),
        Dispatched=('DISP.QTY', 'sum'),
        Pending=('PENDING', 'sum')
    ).reset_index()
    
    pdf_buffer = generate_pdf_report(selected_sheet, kpis, {"Sales Executive Performance": sp_summary})
    st.download_button(
        label="📥 Download Complete PDF Summary Report",
        data=pdf_buffer,
        file_name=f"Sales_Report_{selected_sheet}.pdf",
        mime="application/pdf"
    )

# =========================================================
# SECTION 2: SALES PERSON ANALYTICS
# =========================================================
elif section == "👤 Sales Person Analytics":
    st.header("Sales Person Analytics")
    
    df['IS_CANCELLED'] = df['STATUS'].str.lower().str.contains('cancel') | df['REMARK'].str.lower().str.contains('cancel')
    df['CANCELLED_QTY'] = np.where(df['IS_CANCELLED'], df['PO QTY (MT)'], 0)
    
    st.subheader("1. Item-wise Breakdown per Sales Executive")
    sp_item_grp = df.groupby(['SELLER NAME', 'ITEM']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        Dispatched=('DISP.QTY', 'sum'),
        Pending=('PENDING', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum')
    ).reset_index()
    st.dataframe(sp_item_grp, use_container_width=True)
    
    st.markdown("---")
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.subheader("2. Cancelled Orders per Sales Person")
        sp_cancelled = df[df['IS_CANCELLED']].groupby('SELLER NAME').agg(
            CancelledOrders=('PO NO', 'nunique'),
            CancelledQty=('PO QTY (MT)', 'sum')
        ).reset_index()
        st.dataframe(sp_cancelled, use_container_width=True)
        
    with col_b:
        st.subheader("3. Pending Orders per Sales Person")
        sp_pending = df[df['PENDING'] > 0].groupby('SELLER NAME').agg(
            PendingOrders=('PO NO', 'nunique'),
            PendingQty=('PENDING', 'sum')
        ).reset_index()
        st.dataframe(sp_pending, use_container_width=True)
        
    st.markdown("---")
    col_c, col_d = st.columns(2)
    
    with col_c:
        st.subheader("4. Dispatched Qty per Sales Person")
        sp_disp = df.groupby('SELLER NAME')['DISP.QTY'].sum().reset_index()
        fig_disp = px.bar(sp_disp, x='SELLER NAME', y='DISP.QTY', title="Dispatched Qty by Seller", color_discrete_sequence=['#10B981'])
        st.plotly_chart(fig_disp, use_container_width=True)
        
    with col_d:
        st.subheader("5. Order Received Qty per Sales Person")
        sp_rec = df.groupby('SELLER NAME')['PO QTY (MT)'].sum().reset_index()
        fig_rec = px.bar(sp_rec, x='SELLER NAME', y='PO QTY (MT)', title="Received Order Qty by Seller", color_discrete_sequence=['#3B82F6'])
        st.plotly_chart(fig_rec, use_container_width=True)

    st.markdown("---")
    st.subheader("6. Short Closed Orders (Remarks containing 'SC')")
    sc_df = df[df['REMARK'].str.lower().str.contains(r'\bsc\b|short close', na=False)]
    if not sc_df.empty:
        sc_summary = sc_df.groupby(['SELLER NAME', 'PARTY NAME', 'PO NO', 'REMARK']).agg(
            ShortClosedQty=('PENDING', 'sum')
        ).reset_index()
        st.dataframe(sc_summary, use_container_width=True)
    else:
        st.info("No Short Closed ('SC') orders detected in Remarks for this sheet.")

# =========================================================
# SECTION 3: MONTH WISE & DATE FILTER
# =========================================================
elif section == "📅 Month Wise & Date Filter":
    st.header(f"Date Analytics — {selected_sheet}")
    
    mode = st.radio("Select View Mode", ["Single Date Filter", "Compare Two Dates"])
    valid_dates = sorted(df['DATE'].dropna().dt.date.unique())
    
    if not valid_dates:
        st.warning("No valid Dispatch Dates found in this sheet.")
        st.stop()
        
    if mode == "Single Date Filter":
        selected_date = st.selectbox("Select Date", valid_dates)
        filtered_df = df[df['DATE'].dt.date == selected_date]
        st.subheader(f"Details for Date: {selected_date}")
        
        d_kpis = calculate_kpis(filtered_df)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("POs Active", d_kpis['total_po'])
        c2.metric("Dispatched Qty", f"{d_kpis['dispatched_qty']:,.2f} MT")
        c3.metric("Pending Qty", f"{d_kpis['pending_qty']:,.2f} MT")
        c4.metric("Total Amount", f"₹{d_kpis['total_amount']:,.2f}")
        
        st.dataframe(filtered_df, use_container_width=True)
        
    else:
        st.subheader("Compare Performance Between Two Dates")
        col1, col2 = st.columns(2)
        with col1:
            date1 = st.selectbox("Select First Date", valid_dates, index=0)
        with col2:
            date2 = st.selectbox("Select Second Date", valid_dates, index=min(1, len(valid_dates)-1))
            
        df1 = df[df['DATE'].dt.date == date1]
        df2 = df[df['DATE'].dt.date == date2]
        
        kpi1 = calculate_kpis(df1)
        kpi2 = calculate_kpis(df2)
        
        comp_df = pd.DataFrame({
            "Metric": ["Dispatched Qty (MT)", "Pending Qty (MT)", "Total Amount (₹)", "Unique Parties"],
            f"Date: {date1}": [kpi1['dispatched_qty'], kpi1['pending_qty'], kpi1['total_amount'], kpi1['num_parties']],
            f"Date: {date2}": [kpi2['dispatched_qty'], kpi2['pending_qty'], kpi2['total_amount'], kpi2['num_parties']],
            "Difference": [
                kpi2['dispatched_qty'] - kpi1['dispatched_qty'],
                kpi2['pending_qty'] - kpi1['pending_qty'],
                kpi2['total_amount'] - kpi1['total_amount'],
                kpi2['num_parties'] - kpi1['num_parties']
            ]
        })
        
        st.table(comp_df)
        
        fig_comp = go.Figure(data=[
            go.Bar(name=str(date1), x=["Dispatched Qty", "Pending Qty"], y=[kpi1['dispatched_qty'], kpi1['pending_qty']]),
            go.Bar(name=str(date2), x=["Dispatched Qty", "Pending Qty"], y=[kpi2['dispatched_qty'], kpi2['pending_qty']])
        ])
        fig_comp.update_layout(barmode='group', title="Comparison of Quantities (MT)")
        st.plotly_chart(fig_comp, use_container_width=True)

# =========================================================
# SECTION 4: PARTY WISE REPORT
# =========================================================
elif section == "🏢 Party Wise Report":
    st.header("Party Wise Analytics")
    
    party_grp = df.groupby(['PARTY NAME', 'SELLER NAME']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        PendingQty=('PENDING', 'sum')
    ).reset_index()
    
    st.subheader("Party Wise Ordered, Dispatched, and Pending Qty (with Sales Person)")
    
    all_parties = ["All"] + sorted(party_grp['PARTY NAME'].unique().tolist())
    selected_party = st.selectbox("Filter by Party Name", all_parties)
    
    if selected_party != "All":
        party_grp = party_grp[party_grp['PARTY NAME'] == selected_party]
        
    st.dataframe(party_grp, use_container_width=True)
    
    fig_party = px.bar(
        party_grp.head(20), 
        x='PARTY NAME', 
        y=['OrderedQty', 'DispatchedQty', 'PendingQty'],
        title="Top Parties - Ordered vs Dispatched vs Pending",
        barmode='group'
    )
    st.plotly_chart(fig_party, use_container_width=True)

# =========================================================
# SECTION 5: MATERIAL & WIDTH BREAKDOWN
# =========================================================
elif section == "📐 Material & Width Breakdown":
    st.header("Material & Width Breakdown")
    st.info("💡 Size is parsed strictly on **Width** (e.g. 1500 extracted from 1500X6300).")
    
    df['IS_CANCELLED'] = df['STATUS'].str.lower().str.contains('cancel') | df['REMARK'].str.lower().str.contains('cancel')
    df['CANCELLED_QTY'] = np.where(df['IS_CANCELLED'], df['PO QTY (MT)'], 0)
    
    width_grp = df.groupby(['ITEM', 'WIDTH']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        PendingQty=('PENDING', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum')
    ).reset_index()
    
    st.dataframe(width_grp, use_container_width=True)
    
    fig_width = px.bar(
        width_grp, 
        x='WIDTH', 
        y=['OrderedQty', 'DispatchedQty', 'PendingQty'], 
        color='ITEM',
        title="Qty Breakdown by Width and Item",
        barmode='group'
    )
    st.plotly_chart(fig_width, use_container_width=True)
