import streamlit as st
import pandas as pd
import plotly.express as px
import re

# Set page configuration as first Streamlit call
st.set_page_config(page_title="Sales & Dispatch Analytics", layout="wide")

# ---------------------------------------------------------
# 1. COLUMN MAPPING & CLEANING UTILITIES
# ---------------------------------------------------------
COLUMN_MAP = {
    'S_NO': 'S_NO', 'SR NO': 'S_NO', 'S.NO': 'S_NO',
    'PO NO': 'PO_NO', 'PO NO.': 'PO_NO', 'PO NUMBER': 'PO_NO',
    'DO .NO.': 'DO_NO', 'DO NO': 'DO_NO', 'DO NO.': 'DO_NO',
    'PO DATE': 'PO_DATE', 'ORDER DATE': 'PO_DATE',
    'PARTY NAME': 'PARTY_NAME', 'CUSTOMER NAME': 'PARTY_NAME',
    'BROKER': 'BROKER',
    'SECTOR': 'SECTOR',
    'PLACE': 'PLACE', 'CITY': 'PLACE',
    'SELLER NAME': 'SELLER_NAME', 'SALES PERSON': 'SELLER_NAME',
    'ITEM': 'ITEM', 'DISCRIPTION': 'ITEM', 'DESCRIPTION': 'ITEM',
    'THIKNESS': 'THICKNESS', 'THICKNESS': 'THICKNESS',
    'SIZE': 'SIZE', 'SIZE (MM)': 'SIZE',
    'GRADE': 'GRADE',
    'PO QTY (MT)': 'PO_QTY', 'PO QTY': 'PO_QTY',
    'PER TON': 'PER_TON', 'RATE PER TON': 'PER_TON', 'RATE': 'PER_TON',
    'INV NO.': 'INV_NO', 'INV NO': 'INV_NO', 'INVOICE NO': 'INV_NO',
    'DATE': 'DISPATCH_DATE', 'DISPATCH DATE': 'DISPATCH_DATE',
    'DISP.QTY': 'DISP_QTY', 'DISPATCH QTY': 'DISP_QTY', 'DISPATCHED QTY': 'DISP_QTY',
    'PENDING': 'PENDING_QTY', 'PENDING QTY': 'PENDING_QTY',
    'PAYMENT': 'PAYMENT_TERMS', 'PAYMENT MODE': 'PAYMENT_TERMS',
    'DISP. TH.': 'DISPATCH_THROUGH', 'DISPATCH THROUGH': 'DISPATCH_THROUGH',
    'STATUS': 'STATUS',
    'REMARK': 'REMARK', 'REMARKS': 'REMARK',
    'MOBILE NO': 'MOBILE_NO', 'MOB NO.': 'MOBILE_NO', 'MOBILE': 'MOBILE_NO'
}

EXPECTED_COLUMNS = [
    'S_NO', 'PO_NO', 'DO_NO', 'PO_DATE', 'PARTY_NAME', 'BROKER', 'SECTOR', 
    'PLACE', 'SELLER_NAME', 'ITEM', 'THICKNESS', 'SIZE', 'GRADE', 'PO_QTY', 
    'PER_TON', 'INV_NO', 'DISPATCH_DATE', 'DISP_QTY', 'PENDING_QTY', 
    'PAYMENT_TERMS', 'DISPATCH_THROUGH', 'STATUS', 'REMARK', 'MOBILE_NO'
]

def parse_size(size_val):
    if pd.isna(size_val):
        return None, None
    size_str = str(size_val).upper().replace(" ", "").replace("*", "X")
    match = re.search(r'(\d+(?:\.\d+)?)\s*X\s*(\d+(?:\.\d+)?)', size_str)
    if match:
        try:
            return float(match.group(1)), float(match.group(2))
        except ValueError:
            return None, None
    return None, None

def clean_data(df):
    renamed_cols = {}
    for col in df.columns:
        clean_col = str(col).strip().upper()
        renamed_cols[col] = COLUMN_MAP.get(clean_col, clean_col)
    df = df.rename(columns=renamed_cols)

    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df['PO_DATE'] = pd.to_datetime(df['PO_DATE'], errors='coerce')
    df['DISPATCH_DATE'] = pd.to_datetime(df['DISPATCH_DATE'], errors='coerce')

    num_cols = ['PO_QTY', 'PER_TON', 'DISP_QTY', 'PENDING_QTY', 'THICKNESS']
    for col in num_cols:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '').str.strip(), errors='coerce').fillna(0)

    text_cols = ['PARTY_NAME', 'BROKER', 'SECTOR', 'PLACE', 'SELLER_NAME', 'ITEM', 'GRADE', 'STATUS', 'PO_NO', 'DO_NO']
    for col in text_cols:
        df[col] = df[col].astype(str).fillna("Unknown").str.strip()

    df['TOTAL_REVENUE'] = df['PO_QTY'] * df['PER_TON']
    sizes = df['SIZE'].apply(parse_size)
    df['WIDTH'] = [s[0] for s in sizes]
    df['LENGTH'] = [s[1] for s in sizes]

    return df

@st.cache_data
def load_excel_data(uploaded_file):
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names

    monthly_dfs = []
    pending_df = pd.DataFrame()

    for sheet in sheet_names:
        df_sheet = excel_file.parse(sheet)
        if sheet.strip().upper() == "PENDING DISPATCH":
            pending_df = clean_data(df_sheet)
        else:
            cleaned = clean_data(df_sheet)
            cleaned['SHEET_NAME'] = sheet
            monthly_dfs.append(cleaned)

    combined_monthly = pd.concat(monthly_dfs, ignore_index=True) if monthly_dfs else pd.DataFrame()
    return combined_monthly, pending_df

# ---------------------------------------------------------
# 2. RENDER KPI DASHBOARD & CHARTS
# ---------------------------------------------------------
def render_kpis_and_charts(df, title_prefix=""):
    if df.empty:
        st.warning(f"No records found for {title_prefix}.")
        return

    st.subheader(f"📊 {title_prefix} KPI Overview")
    
    kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)
    kpi_col1.metric("Unique POs", f"{df['PO_NO'].nunique():,}")
    kpi_col2.metric("Unique DOs", f"{df['DO_NO'].nunique():,}")
    kpi_col3.metric("Parties", f"{df['PARTY_NAME'].nunique():,}")
    kpi_col4.metric("Brokers", f"{df['BROKER'].nunique():,}")
    kpi_col5.metric("Sales Persons", f"{df['SELLER_NAME'].nunique():,}")

    kpi_col6, kpi_col7, kpi_col8, kpi_col9, kpi_col10 = st.columns(5)
    kpi_col6.metric("Total PO Qty (MT)", f"{df['PO_QTY'].sum():,.2f}")
    kpi_col7.metric("Dispatched Qty (MT)", f"{df['DISP_QTY'].sum():,.2f}")
    kpi_col8.metric("Pending Qty (MT)", f"{df['PENDING_QTY'].sum():,.2f}")
    kpi_col9.metric("Total Revenue (₹)", f"₹{df['TOTAL_REVENUE'].sum():,.2f}")
    kpi_col10.metric("Sectors Served", f"{df['SECTOR'].nunique():,}")

    st.divider()

    st.subheader("📈 Visual Analytics & Insights")

    chart_c1, chart_c2 = st.columns(2)
    with chart_c1:
        seller_summary = df.groupby('SELLER_NAME')[['PO_QTY', 'DISP_QTY', 'PENDING_QTY']].sum().reset_index()
        fig_seller = px.bar(
            seller_summary,
            x='SELLER_NAME', y=['PO_QTY', 'DISP_QTY', 'PENDING_QTY'],
            barmode='group', title="Salesperson Performance (Ordered vs Dispatched vs Pending)",
            labels={'value': 'Quantity (MT)', 'SELLER_NAME': 'Sales Person'}
        )
        st.plotly_chart(fig_seller, use_container_width=True)

    with chart_c2:
        party_summary = df.groupby('PARTY_NAME')['PENDING_QTY'].sum().nlargest(10).reset_index()
        fig_party_pending = px.bar(
            party_summary,
            x='PENDING_QTY', y='PARTY_NAME', orientation='h',
            title="Top 10 Parties by Pending Quantity",
            labels={'PENDING_QTY': 'Pending Qty (MT)', 'PARTY_NAME': 'Party Name'}
        )
        st.plotly_chart(fig_party_pending, use_container_width=True)

    chart_c3, chart_c4 = st.columns(2)
    with chart_c3:
        fig_status = px.pie(
            df, names='STATUS', title="Order Status Distribution",
            hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel
        )
        st.plotly_chart(fig_status, use_container_width=True)

    with chart_c4:
        valid_sizes = df[df['WIDTH'].notnull() & df['LENGTH'].notnull()]
        if not valid_sizes.empty:
            fig_size = px.scatter(
                valid_sizes,
                x='WIDTH', y='LENGTH', size='PO_QTY', color='SELLER_NAME',
                hover_data=['PARTY_NAME', 'ITEM', 'GRADE'],
                title="Size Analysis (Width vs Length vs Order Volume)",
                labels={'WIDTH': 'Width (mm)', 'LENGTH': 'Length (mm)'}
            )
            st.plotly_chart(fig_size, use_container_width=True)
        else:
            st.info("No valid Width/Length size data available for scatter plot.")

    st.subheader("📋 Detailed Data View")
    st.dataframe(df.drop(columns=['WIDTH', 'LENGTH'], errors='ignore'), use_container_width=True)


# ---------------------------------------------------------
# 3. STREAMLIT APPLICATION LAYOUT
# ---------------------------------------------------------
st.title("🏭 Sales, Dispatch & Order Tracking Analytics Dashboard")

uploaded_file = st.sidebar.file_uploader("Upload Sales Excel Workbook", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_monthly, df_pending = load_excel_data(uploaded_file)

        main_tabs = st.tabs(["📅 Date Analytics & Comparison", "📜 Party Ordering History", "⏳ PENDING DISPATCH Sheet"])

        with main_tabs[0]:
            st.sidebar.header("Date Filter Settings")
            mode = st.sidebar.radio("Analysis Mode", ["Single Date Analysis", "Between Two Dates Comparison"])

            valid_dates = df_monthly['DISPATCH_DATE'].dropna()
            
            if valid_dates.empty:
                st.warning("No valid dispatch dates found in the uploaded monthly sheets.")
            else:
                min_date = valid_dates.min().date()
                max_date = valid_dates.max().date()

                if mode == "Single Date Analysis":
                    selected_date = st.sidebar.date_input("Select Dispatch Date", value=max_date, min_value=min_date, max_value=max_date)
                    filtered_df = df_monthly[df_monthly['DISPATCH_DATE'].dt.date == selected_date]
                    render_kpis_and_charts(filtered_df, title_prefix=f"Date: {selected_date}")

                else:
                    col_d1, col_d2 = st.sidebar.columns(2)
                    date1 = col_d1.date_input("Start Date", value=min_date, min_value=min_date, max_value=max_date)
                    date2 = col_d2.date_input("End Date", value=max_date, min_value=min_date, max_value=max_date)

                    df_range = df_monthly[(df_monthly['DISPATCH_DATE'].dt.date >= date1) & (df_monthly['DISPATCH_DATE'].dt.date <= date2)]
                    
                    st.header(f"Range Analysis: {date1} to {date2}")
                    render_kpis_and_charts(df_range, title_prefix=f"Period ({date1} to {date2})")

                    st.divider()
                    st.subheader("🔄 Direct Day-vs-Day Comparison")
                    
                    df_d1 = df_monthly[df_monthly['DISPATCH_DATE'].dt.date == date1]
                    df_d2 = df_monthly[df_monthly['DISPATCH_DATE'].dt.date == date2]

                    comp_col1, comp_col2 = st.columns(2)
                    with comp_col1:
                        st.markdown(f"### 📅 Baseline: {date1}")
                        st.metric("Total Dispatch Qty", f"{df_d1['DISP_QTY'].sum():,.2f} MT")
                        st.metric("Total Revenue", f"₹{df_d1['TOTAL_REVENUE'].sum():,.2f}")
                        st.metric("Active Parties", df_d1['PARTY_NAME'].nunique())
                    
                    with comp_col2:
                        diff_qty = df_d2['DISP_QTY'].sum() - df_d1['DISP_QTY'].sum()
                        diff_rev = df_d2['TOTAL_REVENUE'].sum() - df_d1['TOTAL_REVENUE'].sum()
                        
                        st.markdown(f"### 📅 Target: {date2}")
                        st.metric("Total Dispatch Qty", f"{df_d2['DISP_QTY'].sum():,.2f} MT", delta=f"{diff_qty:,.2f} MT")
                        st.metric("Total Revenue", f"₹{df_d2['TOTAL_REVENUE'].sum():,.2f}", delta=f"₹{diff_rev:,.2f}")
                        st.metric("Active Parties", df_d2['PARTY_NAME'].nunique(), delta=df_d2['PARTY_NAME'].nunique() - df_d1['PARTY_NAME'].nunique())

        with main_tabs[1]:
            st.header("🏢 Party Ordering Lifecycle & Recency Analytics")
            
            if not df_monthly.empty:
                party_summary = df_monthly.groupby('PARTY_NAME').agg(
                    First_Order_Date=('PO_DATE', 'min'),
                    Last_Order_Date=('PO_DATE', 'max'),
                    Total_Orders=('PO_NO', 'nunique'),
                    Total_PO_Qty=('PO_QTY', 'sum'),
                    Total_Dispatched_Qty=('DISP_QTY', 'sum'),
                    Total_Pending_Qty=('PENDING_QTY', 'sum'),
                    Total_Spend=('TOTAL_REVENUE', 'sum')
                ).reset_index().sort_values(by='Last_Order_Date', ascending=False)

                st.dataframe(party_summary, use_container_width=True)

                selected_party = st.selectbox("Drill-down by Party Name", options=df_monthly['PARTY_NAME'].unique())
                if selected_party:
                    party_df = df_monthly[df_monthly['PARTY_NAME'] == selected_party]
                    st.subheader(f"Order Details for {selected_party}")
                    st.dataframe(party_df[['PO_NO', 'PO_DATE', 'SELLER_NAME', 'BROKER', 'ITEM', 'SIZE', 'PO_QTY', 'DISP_QTY', 'PENDING_QTY', 'STATUS']], use_container_width=True)
            else:
                st.info("No data available in monthly sheets.")

        with main_tabs[2]:
            st.header("⏳ Dedicated Pending Dispatch Report")
            if not df_pending.empty:
                render_kpis_and_charts(df_pending, title_prefix="PENDING DISPATCH SHEET")
            else:
                st.info("No 'PENDING DISPATCH' sheet found in the uploaded file or the sheet contains no records.")

    except Exception as e:
        st.error(f"An error occurred while processing the file: {str(e)}")

else:
    st.info("👈 Upload an Excel workbook using the sidebar to generate reports.")
