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
            
    df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=header_row_idx)
    df.columns = [str(c).strip() for c in df.columns]
    
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
    
    for std_col in EXPECTED_COLUMNS.keys():
        if std_col not in df.columns:
            df[std_col] = np.nan

    # Datetime conversions (Day-first + dot replacement)
    if 'PO DATE' in df.columns:
        po_date_clean = df['PO DATE'].astype(str).str.replace('.', '/', regex=False)
        df['PO DATE'] = pd.to_datetime(po_date_clean, dayfirst=True, errors='coerce')

    if 'DATE' in df.columns:
        disp_date_clean = df['DATE'].astype(str).str.replace('.', '/', regex=False)
        df['DATE'] = pd.to_datetime(disp_date_clean, dayfirst=True, errors='coerce')
    
    df['PO_DATE_STR'] = df['PO DATE'].dt.strftime('%d/%m/%Y').fillna('N/A')
    
    numeric_cols = ['PO QTY (MT)', 'PER TON', 'DISP.QTY', 'PENDING']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
    df['AMOUNT'] = df['PO QTY (MT)'] * df['PER TON']
    
    def parse_width(size_val):
        if pd.isna(size_val):
            return "N/A"
        match = re.search(r'(\d+)\s*[xX*]\s*(\d+)', str(size_val))
        if match:
            return match.group(1)
        num = re.findall(r'\d+', str(size_val))
        return num[0] if num else str(size_val)

    df['WIDTH'] = df['SIZE'].apply(parse_width)
    
    str_cols = ['PARTY NAME', 'SELLER NAME', 'ITEM', 'STATUS', 'REMARK', 'BROKER', 'SECTOR', 'PLACE', 'PO NO', 'DO NO']
    for c in str_cols:
        df[c] = df[c].fillna('Unknown').astype(str).str.strip()

    return df

# ---------------------------------------------------------
# PDF Report Generator
# ---------------------------------------------------------
def generate_pdf_report(section_title, sheet_name, kpis, tables_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#1E3A8A'), spaceAfter=10
    )
    section_style = ParagraphStyle(
        'DocSection', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#1E40AF'), spaceBefore=10, spaceAfter=6
    )
    
    story.append(Paragraph(f"<b>{section_title} Report — {sheet_name}</b>", title_style))
    story.append(Spacer(1, 10))
    
    if kpis:
        story.append(Paragraph("Key Performance Indicators (KPIs)", section_style))
        kpi_data = [["Metric", "Value"]]
        for k, v in kpis.items():
            if isinstance(v, (int, np.integer)):
                val_str = f"{v:,}"
            elif isinstance(v, (float, np.floating)):
                val_str = f"₹{v:,.2f}" if "Amount" in k else f"{v:,.2f}"
            else:
                val_str = str(v)
            kpi_data.append([k, val_str])
            
        t_kpi = Table(kpi_data, colWidths=[220, 220])
        t_kpi.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#3B82F6')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 5),
            ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#F3F4F6')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
        ]))
        story.append(t_kpi)
        story.append(Spacer(1, 15))
    
    for title, df in tables_dict.items():
        if df is not None and not df.empty:
            story.append(Paragraph(f"<b>{title}</b>", section_style))
            sub_df = df.head(20).reset_index(drop=True)
            table_data = [sub_df.columns.tolist()] + sub_df.values.tolist()
            table_data = [[str(cell)[:25] for cell in row] for row in table_data]
            
            col_width = 480 / max(len(sub_df.columns), 1)
            t_data = Table(table_data, colWidths=[col_width]*len(sub_df.columns))
            t_data.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4B5563')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('FONTSIZE', (0,0), (-1,-1), 7),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
            ]))
            story.append(t_data)
            story.append(Spacer(1, 12))
            
    doc.build(story)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# Sidebar Controls & File Upload
# ---------------------------------------------------------
st.sidebar.title("📌 Navigation & Controls")
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("👈 Please upload an Excel workbook from the left sidebar to start.")
    st.stop()

xl = pd.ExcelFile(uploaded_file)
sheet_names = xl.sheet_names

# Main Navigation Section
section = st.sidebar.radio("Go to Section", [
    "📅 Month Wise & Date Filter",
    "📊 All Sales & Dispatch Analytics",
    "🚚 Pending Dispatch"
])

def calculate_kpis(data):
    total_po = data['PO NO'].replace('Unknown', np.nan).dropna().nunique()
    total_do = data['DO NO'].replace('Unknown', np.nan).dropna().nunique()
    num_parties = data['PARTY NAME'].replace('Unknown', np.nan).dropna().nunique()
    total_po_qty = data['PO QTY (MT)'].sum()
    total_amount = data['AMOUNT'].sum()
    dispatched_qty = data['DISP.QTY'].sum()
    pending_qty = data['PENDING'].sum()
    
    return {
        'Overall PO Count': total_po,
        'Overall DO Count': total_do,
        'Number of Parties': num_parties,
        'Total PO Quantity (MT)': total_po_qty,
        'Total PO Amount': total_amount,
        'Dispatched Qty (MT)': dispatched_qty,
        'Pending Qty (MT)': pending_qty
    }

# =========================================================
# SECTION 1: MONTH WISE & DATE FILTER
# =========================================================
if section == "📅 Month Wise & Date Filter":
    selected_sheet = st.sidebar.selectbox("Select Month / Sheet", sheet_names)
    df = load_and_clean_sheet(uploaded_file, selected_sheet)
    
    st.header(f"Date & Month Analytics — {selected_sheet}")
    
    mode = st.radio("Select View Mode", ["Single Date Filter", "Compare Month-Wise / Date-Wise"])
    
    po_dates_df = df.dropna(subset=['PO DATE']).copy()
    po_dates_df['PO_DATE_ONLY'] = po_dates_df['PO DATE'].dt.date
    unique_dates = sorted(po_dates_df['PO_DATE_ONLY'].unique())
    
    if not unique_dates:
        st.warning("No valid PO Dates found in this sheet.")
        st.stop()

    if mode == "Single Date Filter":
        date_options = [d.strftime('%d/%m/%Y') for d in unique_dates]
        selected_date_str = st.selectbox("Select PO Date (DD/MM/YYYY)", date_options)
        
        selected_date = pd.to_datetime(selected_date_str, format='%d/%m/%Y').date()
        filtered_df = df[df['PO DATE'].dt.date == selected_date]
        
        st.subheader(f"Details for PO Date: {selected_date_str}")
        d_kpis = calculate_kpis(filtered_df)
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("POs Received", d_kpis['Overall PO Count'])
        c2.metric("Total Ordered Qty", f"{d_kpis['Total PO Quantity (MT)']:,.2f} MT")
        c3.metric("Dispatched Qty", f"{d_kpis['Dispatched Qty (MT)']:,.2f} MT")
        c4.metric("Total Amount", f"₹{d_kpis['Total PO Amount']:,.2f}")
        
        st.dataframe(filtered_df, use_container_width=True)
        
        st.markdown("---")
        pdf_buf = generate_pdf_report(
            f"PO Date {selected_date_str}", 
            selected_sheet, 
            d_kpis, 
            {"PO Details": filtered_df[['PO NO', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'PO QTY (MT)', 'DISP.QTY', 'PENDING']]}
        )
        st.download_button("📥 Download Single Date Report PDF", data=pdf_buf, file_name=f"Date_Report_{selected_date_str.replace('/', '-')}.pdf", mime="application/pdf")

    else:
        st.subheader("Compare Performance")
        comp_type = st.radio("Comparison Mode", ["Date Wise", "Month Wise"], horizontal=True)
        
        if comp_type == "Date Wise":
            date_options = [d.strftime('%d/%m/%Y') for d in unique_dates]
            col1, col2 = st.columns(2)
            with col1:
                date1_str = st.selectbox("First PO Date (DD/MM/YYYY)", date_options, index=0)
            with col2:
                date2_str = st.selectbox("Second PO Date (DD/MM/YYYY)", date_options, index=min(1, len(date_options)-1))
                
            d1 = pd.to_datetime(date1_str, format='%d/%m/%Y').date()
            d2 = pd.to_datetime(date2_str, format='%d/%m/%Y').date()
            
            df1 = df[df['PO DATE'].dt.date == d1]
            df2 = df[df['PO DATE'].dt.date == d2]
            
            label1, label2 = date1_str, date2_str
            
        else:
            col1, col2 = st.columns(2)
            with col1:
                sheet1 = st.selectbox("First Month Sheet", sheet_names, index=0, key="m1")
            with col2:
                sheet2 = st.selectbox("Second Month Sheet", sheet_names, index=min(1, len(sheet_names)-1), key="m2")
                
            df1 = load_and_clean_sheet(uploaded_file, sheet1)
            df2 = load_and_clean_sheet(uploaded_file, sheet2)
            
            label1, label2 = sheet1, sheet2
            
        kpi1 = calculate_kpis(df1)
        kpi2 = calculate_kpis(df2)
        
        comp_df = pd.DataFrame({
            "Metric": ["Total PO Count", "Ordered Qty (MT)", "Dispatched Qty (MT)", "Pending Qty (MT)", "Total Amount (₹)", "Parties Count"],
            f"{label1}": [kpi1['Overall PO Count'], kpi1['Total PO Quantity (MT)'], kpi1['Dispatched Qty (MT)'], kpi1['Pending Qty (MT)'], kpi1['Total PO Amount'], kpi1['Number of Parties']],
            f"{label2}": [kpi2['Overall PO Count'], kpi2['Total PO Quantity (MT)'], kpi2['Dispatched Qty (MT)'], kpi2['Pending Qty (MT)'], kpi2['Total PO Amount'], kpi2['Number of Parties']],
            "Difference": [
                kpi2['Overall PO Count'] - kpi1['Overall PO Count'],
                kpi2['Total PO Quantity (MT)'] - kpi1['Total PO Quantity (MT)'],
                kpi2['Dispatched Qty (MT)'] - kpi1['Dispatched Qty (MT)'],
                kpi2['Pending Qty (MT)'] - kpi1['Pending Qty (MT)'],
                kpi2['Total PO Amount'] - kpi1['Total PO Amount'],
                kpi2['Number of Parties'] - kpi1['Number of Parties']
            ]
        })
        
        st.table(comp_df)
        
        fig_comp = go.Figure(data=[
            go.Bar(name=str(label1), x=["Ordered Qty", "Dispatched Qty", "Pending Qty"], y=[kpi1['Total PO Quantity (MT)'], kpi1['Dispatched Qty (MT)'], kpi1['Pending Qty (MT)']]),
            go.Bar(name=str(label2), x=["Ordered Qty", "Dispatched Qty", "Pending Qty"], y=[kpi2['Total PO Quantity (MT)'], kpi2['Dispatched Qty (MT)'], kpi2['Pending Qty (MT)']])
        ])
        fig_comp.update_layout(barmode='group', title=f"Comparison: {label1} vs {label2}")
        st.plotly_chart(fig_comp, use_container_width=True)
        
        st.markdown("---")
        pdf_buf = generate_pdf_report(f"Comparison {label1} vs {label2}", selected_sheet, None, {"Comparison Summary": comp_df})
        st.download_button("📥 Download Comparison Report PDF", data=pdf_buf, file_name=f"Comparison_Report_{label1}_vs_{label2}.pdf", mime="application/pdf")

# =========================================================
# SECTION 2: ALL SALES & DISPATCH ANALYTICS (MERGED SECTION)
# =========================================================
elif section == "📊 All Sales & Dispatch Analytics":
    selected_sheet = st.sidebar.selectbox("Select Month / Sheet", sheet_names)
    df = load_and_clean_sheet(uploaded_file, selected_sheet)
    
    st.header(f"All Sales & Dispatch Analytics — {selected_sheet}")
    
    # ---------------- 1. KPI OVERVIEW ----------------
    st.subheader("1. Key Performance Indicators (KPIs)")
    kpis = calculate_kpis(df)
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Overall PO Count", f"{kpis['Overall PO Count']:,}")
    c2.metric("Overall DO Count", f"{kpis['Overall DO Count']:,}")
    c3.metric("Number of Parties", f"{kpis['Number of Parties']:,}")
    c4.metric("Total PO Qty (MT)", f"{kpis['Total PO Quantity (MT)']:,.2f}")
    
    c5, c6, c7 = st.columns(3)
    c5.metric("Total Amount (PO Qty × Rate)", f"₹{kpis['Total PO Amount']:,.2f}")
    c6.metric("Sum of Dispatched Qty (MT)", f"{kpis['Dispatched Qty (MT)']:,.2f}")
    c7.metric("Sum of Pending Qty (MT)", f"{kpis['Pending Qty (MT)']:,.2f}")
    
    fig_kpi = go.Figure(data=[go.Pie(
        labels=['Dispatched Qty', 'Pending Qty'],
        values=[kpis['Dispatched Qty (MT)'], kpis['Pending Qty (MT)']],
        hole=.4,
        marker_colors=['#10B981', '#EF4444']
    )])
    st.plotly_chart(fig_kpi, use_container_width=True)
    
    st.markdown("---")
    
    # ---------------- 2. SALES PERSON ANALYTICS ----------------
    st.subheader("2. Sales Executive Analytics")
    
    df['IS_CANCELLED'] = df['STATUS'].str.lower().str.contains('cancel') | df['REMARK'].str.lower().str.contains('cancel')
    df['CANCELLED_QTY'] = np.where(df['IS_CANCELLED'], df['PO QTY (MT)'], 0)
    
    sp_item_grp = df.groupby(['SELLER NAME', 'ITEM']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        Dispatched=('DISP.QTY', 'sum'),
        Pending=('PENDING', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum')
    ).reset_index()
    st.write("**Sales Executive Item-Wise Breakdown**")
    st.dataframe(sp_item_grp, use_container_width=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.write("**Cancelled Orders per Sales Person**")
        sp_cancelled = df[df['IS_CANCELLED']].groupby('SELLER NAME').agg(
            CancelledOrders=('PO NO', 'nunique'),
            CancelledQty=('PO QTY (MT)', 'sum')
        ).reset_index()
        st.dataframe(sp_cancelled, use_container_width=True)
        
    with col_b:
        st.write("**Pending Orders per Sales Person**")
        sp_pending = df[df['PENDING'] > 0].groupby('SELLER NAME').agg(
            PendingOrders=('PO NO', 'nunique'),
            PendingQty=('PENDING', 'sum')
        ).reset_index()
        st.dataframe(sp_pending, use_container_width=True)
        
    col_c, col_d = st.columns(2)
    with col_c:
        sp_disp = df.groupby('SELLER NAME')['DISP.QTY'].sum().reset_index()
        fig_disp = px.bar(sp_disp, x='SELLER NAME', y='DISP.QTY', title="Dispatched Qty by Sales Person", color_discrete_sequence=['#10B981'])
        st.plotly_chart(fig_disp, use_container_width=True)
        
    with col_d:
        sp_rec = df.groupby('SELLER NAME')['PO QTY (MT)'].sum().reset_index()
        fig_rec = px.bar(sp_rec, x='SELLER NAME', y='PO QTY (MT)', title="Order Received Qty by Sales Person", color_discrete_sequence=['#3B82F6'])
        st.plotly_chart(fig_rec, use_container_width=True)

    st.write("**Short Closed Orders (Remarks with 'SC')**")
    sc_df = df[df['REMARK'].str.lower().str.contains(r'\bsc\b|short close', na=False)]
    if not sc_df.empty:
        sc_summary = sc_df.groupby(['SELLER NAME', 'PARTY NAME', 'PO NO', 'REMARK']).agg(ShortClosedQty=('PENDING', 'sum')).reset_index()
        st.dataframe(sc_summary, use_container_width=True)
    else:
        sc_summary = pd.DataFrame()
        st.info("No Short Closed ('SC') orders found.")

    st.markdown("---")
    
    # ---------------- 3. PARTY WISE ANALYTICS ----------------
    st.subheader("3. Party Wise Analytics")
    
    party_grp = df.groupby(['PARTY NAME', 'SELLER NAME']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        PendingQty=('PENDING', 'sum')
    ).reset_index()
    
    all_parties = ["All"] + sorted(party_grp['PARTY NAME'].unique().tolist())
    selected_party = st.selectbox("Filter by Party Name", all_parties)
    
    filtered_party_grp = party_grp if selected_party == "All" else party_grp[party_grp['PARTY NAME'] == selected_party]
    st.dataframe(filtered_party_grp, use_container_width=True)
    
    fig_party = px.bar(
        filtered_party_grp.head(20), 
        x='PARTY NAME', 
        y=['OrderedQty', 'DispatchedQty', 'PendingQty'],
        title="Top Parties - Ordered vs Dispatched vs Pending",
        barmode='group'
    )
    st.plotly_chart(fig_party, use_container_width=True)

    st.markdown("---")
    
    # ---------------- 4. MATERIAL & WIDTH BREAKDOWN ----------------
    st.subheader("4. Material & Width Breakdown")
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

    st.markdown("---")
    st.subheader("📄 Export Master Analytics PDF")
    
    tables_to_pdf = {
        "Sales Executive Performance": sp_item_grp,
        "Party Wise Summary": party_grp.head(20),
        "Material & Width Breakdown": width_grp
    }
    
    pdf_buf = generate_pdf_report("Complete Sales & Dispatch Analytics", selected_sheet, kpis, tables_to_pdf)
    st.download_button("📥 Download Full Analytics PDF Report", data=pdf_buf, file_name=f"Master_Analytics_{selected_sheet}.pdf", mime="application/pdf")

# =========================================================
# SECTION 3: PENDING DISPATCH (STANDALONE SECTION)
# =========================================================
elif section == "🚚 Pending Dispatch":
    st.header("🚚 Pending Dispatch Standalone Report")
    
    # Auto-detect sheet named 'PendingDispatch' or similar
    pd_sheet_candidates = [s for s in sheet_names if 'pending' in s.lower() and 'dispatch' in s.lower()]
    target_pd_sheet = pd_sheet_candidates[0] if pd_sheet_candidates else (sheet_names[0] if sheet_names else None)
    
    pd_sheet = st.sidebar.selectbox("Select Pending Dispatch Sheet", sheet_names, index=sheet_names.index(target_pd_sheet) if target_pd_sheet in sheet_names else 0)
    
    df_pd = load_and_clean_sheet(uploaded_file, pd_sheet)
    
    # Calculate KPIs for Pending Dispatch
    pd_kpis = {
        'Total Pending Orders': df_pd[df_pd['PENDING'] > 0]['PO NO'].replace('Unknown', np.nan).dropna().nunique(),
        'Total Parties': df_pd[df_pd['PENDING'] > 0]['PARTY NAME'].replace('Unknown', np.nan).dropna().nunique(),
        'Total Pending Qty (MT)': df_pd['PENDING'].sum(),
        'Pending Amount (Est. ₹)': (df_pd['PENDING'] * df_pd['PER TON']).sum()
    }
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Pending Orders", f"{pd_kpis['Total Pending Orders']:,}")
    c2.metric("Parties Impacted", f"{pd_kpis['Total Parties']:,}")
    c3.metric("Total Pending Qty", f"{pd_kpis['Total Pending Qty (MT)']:,.2f} MT")
    c4.metric("Est. Pending Value", f"₹{pd_kpis['Pending Amount (Est. ₹)']:,.2f}")
    
    st.markdown("---")
    st.subheader("Pending Quantity Breakdown by Sales Person & Item")
    
    pd_summary = df_pd[df_pd['PENDING'] > 0].groupby(['SELLER NAME', 'ITEM', 'PARTY NAME']).agg(
        PendingQty=('PENDING', 'sum'),
        OrderedQty=('PO QTY (MT)', 'sum')
    ).reset_index()
    
    fig_pd = px.bar(
        pd_summary.groupby('SELLER NAME')['PendingQty'].sum().reset_index(),
        x='SELLER NAME',
        y='PendingQty',
        title="Pending Dispatch Qty (MT) by Sales Person",
        color_discrete_sequence=['#EF4444']
    )
    st.plotly_chart(fig_pd, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📋 Pending Dispatch Detailed Data Table")
    
    pending_details_df = df_pd[df_pd['PENDING'] > 0][['PO NO', 'PO_DATE_STR', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'SIZE', 'PO QTY (MT)', 'DISP.QTY', 'PENDING', 'REMARK']]
    st.dataframe(pending_details_df, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📄 Download Pending Dispatch Report")
    
    pdf_buf = generate_pdf_report(
        "Pending Dispatch Summary",
        pd_sheet,
        pd_kpis,
        {"Pending Dispatch Details": pending_details_df}
    )
    st.download_button(
        "📥 Download Pending Dispatch Report PDF",
        data=pdf_buf,
        file_name=f"Pending_Dispatch_Report_{pd_sheet}.pdf",
        mime="application/pdf"
    )
