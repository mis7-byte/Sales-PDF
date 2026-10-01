import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import re

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="Sales & Dispatch Dashboard",
    page_icon="📊",
    layout="wide"
)

# Advanced CSS rules to ensure tables expand FULLY (no scrollbars) when printing to PDF
st.markdown("""
    <style>
    @media print {
        /* Hide sidebar, buttons, headers, and footers during print */
        section[data-testid="stSidebar"], .stButton, header, footer, iframe {
            display: none !important;
        }
        .main .block-container {
            max-width: 100% !important;
            padding: 0 !important;
            margin: 0 !important;
        }
        body {
            background-color: white !important;
        }
        
        /* EXPAND DATAFRAMES & TABLES TO FULL HEIGHT FOR PDF PRINTING */
        div[data-testid="stDataFrame"], 
        div[data-testid="stTable"],
        div[data-testid="element-container"],
        .stDataFrame div {
            height: auto !important;
            max-height: none !important;
            overflow: visible !important;
        }
        
        /* Force scrollable table body containers to show all rows */
        div[data-testid="stDataFrame"] > div > div {
            max-height: none !important;
            height: auto !important;
            overflow: visible !important;
        }
        
        /* Table page break protection */
        tr, td, th {
            page-break-inside: avoid !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

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

    # Datetime conversions (Day-first + dot replacement e.g. 01.09.2026 -> 01/09/2026)
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
    
    str_cols = ['PARTY NAME', 'SELLER NAME', 'ITEM', 'STATUS', 'REMARK', 'BROKER', 'SECTOR', 'PLACE', 'PO NO', 'DO NO']
    for c in str_cols:
        df[c] = df[c].fillna('Unknown').astype(str).str.strip()

    # Cancelled order identification
    df['IS_CANCELLED'] = df['STATUS'].str.lower().str.contains('cancel') | df['REMARK'].str.lower().str.contains('cancel')
    df['CANCELLED_QTY'] = np.where(df['IS_CANCELLED'], df['PO QTY (MT)'], 0.0)
    df['ACTIVE_PENDING_QTY'] = np.where(df['IS_CANCELLED'], 0.0, df['PENDING'])

    # Parse Width from dimensions
    def parse_width(size_val):
        if pd.isna(size_val):
            return "N/A"
        match = re.search(r'(\d+)\s*[xX*]\s*(\d+)', str(size_val))
        if match:
            return match.group(1)
        num = re.findall(r'\d+', str(size_val))
        return num[0] if num else str(size_val)

    df['WIDTH'] = df['SIZE'].apply(parse_width)

    return df

# ---------------------------------------------------------
# Sidebar Controls & Navigation
# ---------------------------------------------------------
st.sidebar.title("📌 Navigation & Controls")
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("👈 Please upload an Excel workbook from the left sidebar to view dashboard metrics.")
    st.stop()

xl = pd.ExcelFile(uploaded_file)
sheet_names = xl.sheet_names

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
    cancelled_qty = data['CANCELLED_QTY'].sum()
    pending_qty = data['ACTIVE_PENDING_QTY'].sum()
    
    return {
        'Overall PO Count': total_po,
        'Overall DO Count': total_do,
        'Number of Parties': num_parties,
        'Total PO Quantity (MT)': total_po_qty,
        'Total PO Amount': total_amount,
        'Dispatched Qty (MT)': dispatched_qty,
        'Cancelled Qty (MT)': cancelled_qty,
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
            "Metric": ["Total PO Count", "Ordered Qty (MT)", "Dispatched Qty (MT)", "Cancelled Qty (MT)", "Pending Qty (MT)", "Total Amount (₹)", "Parties Count"],
            f"{label1}": [kpi1['Overall PO Count'], kpi1['Total PO Quantity (MT)'], kpi1['Dispatched Qty (MT)'], kpi1['Cancelled Qty (MT)'], kpi1['Pending Qty (MT)'], kpi1['Total PO Amount'], kpi1['Number of Parties']],
            f"{label2}": [kpi2['Overall PO Count'], kpi2['Total PO Quantity (MT)'], kpi2['Dispatched Qty (MT)'], kpi2['Cancelled Qty (MT)'], kpi2['Pending Qty (MT)'], kpi2['Total PO Amount'], kpi2['Number of Parties']],
            "Difference": [
                kpi2['Overall PO Count'] - kpi1['Overall PO Count'],
                kpi2['Total PO Quantity (MT)'] - kpi1['Total PO Quantity (MT)'],
                kpi2['Dispatched Qty (MT)'] - kpi1['Dispatched Qty (MT)'],
                kpi2['Cancelled Qty (MT)'] - kpi1['Cancelled Qty (MT)'],
                kpi2['Pending Qty (MT)'] - kpi1['Pending Qty (MT)'],
                kpi2['Total PO Amount'] - kpi1['Total PO Amount'],
                kpi2['Number of Parties'] - kpi1['Number of Parties']
            ]
        })
        
        st.table(comp_df)
        
        fig_comp = go.Figure(data=[
            go.Bar(
                name=str(label1), 
                x=["Ordered Qty", "Dispatched Qty", "Cancelled Qty", "Pending Qty"], 
                y=[kpi1['Total PO Quantity (MT)'], kpi1['Dispatched Qty (MT)'], kpi1['Cancelled Qty (MT)'], kpi1['Pending Qty (MT)']], 
                text=[f"{kpi1['Total PO Quantity (MT)']:,.1f}", f"{kpi1['Dispatched Qty (MT)']:,.1f}", f"{kpi1['Cancelled Qty (MT)']:,.1f}", f"{kpi1['Pending Qty (MT)']:,.1f}"], 
                textposition='outside'
            ),
            go.Bar(
                name=str(label2), 
                x=["Ordered Qty", "Dispatched Qty", "Cancelled Qty", "Pending Qty"], 
                y=[kpi2['Total PO Quantity (MT)'], kpi2['Dispatched Qty (MT)'], kpi2['Cancelled Qty (MT)'], kpi2['Pending Qty (MT)']], 
                text=[f"{kpi2['Total PO Quantity (MT)']:,.1f}", f"{kpi2['Dispatched Qty (MT)']:,.1f}", f"{kpi2['Cancelled Qty (MT)']:,.1f}", f"{kpi2['Pending Qty (MT)']:,.1f}"], 
                textposition='outside'
            )
        ])
        fig_comp.update_layout(barmode='group', title=f"Comparison: {label1} vs {label2}")
        st.plotly_chart(fig_comp, use_container_width=True)

# =========================================================
# SECTION 2: ALL SALES & DISPATCH ANALYTICS
# =========================================================
elif section == "📊 All Sales & Dispatch Analytics":
    selected_sheet = st.sidebar.selectbox("Select Month / Sheet", sheet_names)
    df = load_and_clean_sheet(uploaded_file, selected_sheet)
    
    col_head, col_btn = st.columns([3, 1])
    with col_head:
        st.header(f"All Sales & Dispatch Analytics — {selected_sheet}")
    with col_btn:
        st.components.v1.html("""
            <button onclick="window.parent.print()" style="background-color: #1E40AF; color: white; border: none; padding: 10px 16px; font-size: 14px; border-radius: 6px; cursor: pointer; margin-top: 10px; width: 100%;">
                🖨️ Print Full Screen to PDF
            </button>
        """, height=50)

    # ---------------- 1. KPI OVERVIEW ----------------
    st.subheader("1. Key Performance Indicators (KPIs)")
    kpis = calculate_kpis(df)
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Overall PO Count", f"{kpis['Overall PO Count']:,}")
    c2.metric("Overall DO Count", f"{kpis['Overall DO Count']:,}")
    c3.metric("Number of Parties", f"{kpis['Number of Parties']:,}")
    c4.metric("Total PO Qty (MT)", f"{kpis['Total PO Quantity (MT)']:,.2f}")
    
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Total Amount (PO Qty × Rate)", f"₹{kpis['Total PO Amount']:,.2f}")
    c6.metric("Sum of Dispatched Qty (MT)", f"{kpis['Dispatched Qty (MT)']:,.2f}")
    c7.metric("Sum of Cancelled Qty (MT)", f"{kpis['Cancelled Qty (MT)']:,.2f}")
    c8.metric("Sum of Active Pending Qty (MT)", f"{kpis['Pending Qty (MT)']:,.2f}")
    
    # Donut Chart with Data Labels
    fig_kpi = go.Figure(data=[go.Pie(
        labels=['Dispatched Qty', 'Cancelled Qty', 'Pending Qty'],
        values=[kpis['Dispatched Qty (MT)'], kpis['Cancelled Qty (MT)'], kpis['Pending Qty (MT)']],
        hole=.4,
        textinfo='label+value+percent',
        texttemplate='%{label}<br>%{value:,.2f} MT (%{percent})',
        marker_colors=['#10B981', '#F59E0B', '#EF4444']
    )])
    fig_kpi.update_layout(title="Overall Status Breakdown (with Data Labels)")
    st.plotly_chart(fig_kpi, use_container_width=True)
    
    st.markdown("---")
    
    # ---------------- 2. SALES PERSON ANALYTICS ----------------
    st.subheader("2. Sales Executive Analytics")
    
    sp_item_grp = df.groupby(['SELLER NAME', 'ITEM']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        Dispatched=('DISP.QTY', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum'),
        Pending=('ACTIVE_PENDING_QTY', 'sum')
    ).reset_index()
    st.write("**Sales Executive Item-Wise Breakdown**")
    st.dataframe(sp_item_grp, use_container_width=True, height=len(sp_item_grp) * 38 + 40)
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.write("**Cancelled Orders per Sales Person**")
        sp_cancelled = df[df['IS_CANCELLED']].groupby('SELLER NAME').agg(
            CancelledOrders=('PO NO', 'nunique'),
            CancelledQty=('CANCELLED_QTY', 'sum')
        ).reset_index()
        st.dataframe(sp_cancelled, use_container_width=True)
        
    with col_b:
        st.write("**Pending Orders per Sales Person**")
        sp_pending = df[df['ACTIVE_PENDING_QTY'] > 0].groupby('SELLER NAME').agg(
            PendingOrders=('PO NO', 'nunique'),
            PendingQty=('ACTIVE_PENDING_QTY', 'sum')
        ).reset_index()
        st.dataframe(sp_pending, use_container_width=True)
        
    col_c, col_d = st.columns(2)
    with col_c:
        sp_disp = df.groupby('SELLER NAME')['DISP.QTY'].sum().reset_index()
        fig_disp = px.bar(
            sp_disp, 
            x='SELLER NAME', 
            y='DISP.QTY', 
            title="Dispatched Qty by Sales Person", 
            text_auto=',.1f',
            color_discrete_sequence=['#10B981']
        )
        fig_disp.update_traces(textposition='outside')
        st.plotly_chart(fig_disp, use_container_width=True)
        
    with col_d:
        sp_rec = df.groupby('SELLER NAME')['PO QTY (MT)'].sum().reset_index()
        fig_rec = px.bar(
            sp_rec, 
            x='SELLER NAME', 
            y='PO QTY (MT)', 
            title="Order Received Qty by Sales Person", 
            text_auto=',.1f',
            color_discrete_sequence=['#3B82F6']
        )
        fig_rec.update_traces(textposition='outside')
        st.plotly_chart(fig_rec, use_container_width=True)

    st.write("**Short Closed Orders (Remarks with 'SC')**")
    sc_df = df[df['REMARK'].str.lower().str.contains(r'\bsc\b|short close', na=False)]
    if not sc_df.empty:
        sc_summary = sc_df.groupby(['SELLER NAME', 'PARTY NAME', 'PO NO', 'REMARK']).agg(ShortClosedQty=('PENDING', 'sum')).reset_index()
        st.dataframe(sc_summary, use_container_width=True)
    else:
        st.info("No Short Closed ('SC') orders found.")

    st.markdown("---")
    
    # ---------------- 3. PARTY WISE ANALYTICS ----------------
    st.subheader("3. Party Wise Analytics")
    
    party_grp = df.groupby(['PARTY NAME', 'SELLER NAME']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum'),
        PendingQty=('ACTIVE_PENDING_QTY', 'sum')
    ).reset_index()
    
    all_parties = ["All"] + sorted(party_grp['PARTY NAME'].unique().tolist())
    selected_party = st.selectbox("Filter by Party Name", all_parties)
    
    filtered_party_grp = party_grp if selected_party == "All" else party_grp[party_grp['PARTY NAME'] == selected_party]
    st.dataframe(filtered_party_grp, use_container_width=True, height=min(len(filtered_party_grp) * 38 + 40, 1000))
    
    fig_party = px.bar(
        filtered_party_grp.head(15), 
        x='PARTY NAME', 
        y=['OrderedQty', 'DispatchedQty', 'CancelledQty', 'PendingQty'],
        title="Top Parties - Ordered vs Dispatched vs Cancelled vs Pending",
        barmode='group',
        text_auto=',.1f'
    )
    fig_party.update_traces(textposition='outside')
    st.plotly_chart(fig_party, use_container_width=True)

    st.markdown("---")
    
    # ---------------- 4. MATERIAL & WIDTH BREAKDOWN ----------------
    st.subheader("4. Material & Width Breakdown")
    width_grp = df.groupby(['ITEM', 'WIDTH']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum'),
        PendingQty=('ACTIVE_PENDING_QTY', 'sum')
    ).reset_index()
    st.dataframe(width_grp, use_container_width=True, height=len(width_grp) * 38 + 40)
    
    fig_width = px.bar(
        width_grp, 
        x='WIDTH', 
        y=['OrderedQty', 'DispatchedQty', 'CancelledQty', 'PendingQty'], 
        color='ITEM',
        title="Qty Breakdown by Width and Item",
        barmode='group',
        text_auto=',.1f'
    )
    fig_width.update_traces(textposition='outside')
    st.plotly_chart(fig_width, use_container_width=True)

# =========================================================
# SECTION 3: PENDING DISPATCH
# =========================================================
elif section == "🚚 Pending Dispatch":
    st.header("🚚 Pending Dispatch Standalone Report")
    
    pd_sheet_candidates = [s for s in sheet_names if 'pending' in s.lower() and 'dispatch' in s.lower()]
    target_pd_sheet = pd_sheet_candidates[0] if pd_sheet_candidates else (sheet_names[0] if sheet_names else None)
    
    pd_sheet = st.sidebar.selectbox("Select Pending Dispatch Sheet", sheet_names, index=sheet_names.index(target_pd_sheet) if target_pd_sheet in sheet_names else 0)
    
    df_pd = load_and_clean_sheet(uploaded_file, pd_sheet)
    active_pd = df_pd[(df_pd['ACTIVE_PENDING_QTY'] > 0) & (~df_pd['IS_CANCELLED'])]
    
    pd_kpis = {
        'Total Pending Orders': active_pd['PO NO'].replace('Unknown', np.nan).dropna().nunique(),
        'Total Parties': active_pd['PARTY NAME'].replace('Unknown', np.nan).dropna().nunique(),
        'Total Pending Qty (MT)': active_pd['ACTIVE_PENDING_QTY'].sum(),
        'Pending Amount (Est. ₹)': (active_pd['ACTIVE_PENDING_QTY'] * active_pd['PER TON']).sum()
    }
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Pending Orders", f"{pd_kpis['Total Pending Orders']:,}")
    c2.metric("Parties Impacted", f"{pd_kpis['Total Parties']:,}")
    c3.metric("Total Pending Qty", f"{pd_kpis['Total Pending Qty (MT)']:,.2f} MT")
    c4.metric("Est. Pending Value", f"₹{pd_kpis['Pending Amount (Est. ₹)']:,.2f}")
    
    st.markdown("---")
    st.subheader("Pending Quantity Breakdown by Sales Person")
    
    fig_pd = px.bar(
        active_pd.groupby('SELLER NAME')['ACTIVE_PENDING_QTY'].sum().reset_index(),
        x='SELLER NAME',
        y='ACTIVE_PENDING_QTY',
        title="Pending Dispatch Qty (MT) by Sales Person",
        text_auto=',.1f',
        color_discrete_sequence=['#EF4444']
    )
    fig_pd.update_traces(textposition='outside')
    st.plotly_chart(fig_pd, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📋 Pending Dispatch Detailed Data Table")
    
    pending_details_df = active_pd[['PO NO', 'PO_DATE_STR', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'SIZE', 'PO QTY (MT)', 'DISP.QTY', 'ACTIVE_PENDING_QTY', 'REMARK']]
    pending_details_df.rename(columns={'ACTIVE_PENDING_QTY': 'PENDING QTY'}, inplace=True)
    st.dataframe(pending_details_df, use_container_width=True, height=len(pending_details_df) * 38 + 40)
