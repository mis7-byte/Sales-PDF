import io
import re
import warnings
import pandas as pd
import plotly.express as px
import streamlit as st

# Suppress openpyxl warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(page_title="Sales & Dispatch Analytics", layout="wide")

# ---------------------------------------------------------
# 1. COLUMN MAPPING & CLEANING UTILITIES
# ---------------------------------------------------------
COLUMN_MAP = {
    'S_NO': 'S_NO', 'SR NO': 'S_NO', 'S.NO': 'S_NO', 'S. NO.': 'S_NO',
    'PO NO': 'PO_NO', 'PO NO.': 'PO_NO', 'PO NUMBER': 'PO_NO',
    'DO .NO.': 'DO_NO', 'DO NO': 'DO_NO', 'DO NO.': 'DO_NO', 'DO .NO': 'DO_NO',
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
    'PO QTY (MT)': 'PO_QTY', 'PO QTY': 'PO_QTY', 'QTY': 'PO_QTY', 'QUANTITY': 'PO_QTY',
    'PER TON': 'PER_TON', 'RATE PER TON': 'PER_TON', 'RATE': 'PER_TON',
    'INV NO.': 'INV_NO', 'INV NO': 'INV_NO', 'INVOICE NO': 'INV_NO',
    'DATE': 'DISPATCH_DATE', 'DISPATCH DATE': 'DISPATCH_DATE',
    'DISP.QTY': 'DISP_QTY', 'DISPATCH QTY': 'DISP_QTY', 'DISPATCHED QTY': 'DISP_QTY',
    'PENDING': 'PENDING_QTY', 'PENDING QTY': 'PENDING_QTY',
    'PAYMENT': 'PAYMENT_TERMS', 'PAYMENT MODE': 'PAYMENT_TERMS',
    'DISP. TH.': 'DISPATCH_THROUGH', 'DISPATCH THROUGH': 'DISPATCH_THROUGH',
    'STATUS': 'STATUS',
    'REMARK': 'REMARK', 'REMARKS': 'REMARK',
    'MOBILE NO': 'MOBILE_NO', 'MOB NO.': 'MOBILE_NO', 'MOB NO': 'MOBILE_NO', 'MOBILE': 'MOBILE_NO',
    'PCS': 'PIECES', 'NO OF PCS': 'PIECES'
}

EXPECTED_COLUMNS = [
    'S_NO', 'PO_NO', 'DO_NO', 'PO_DATE', 'PARTY_NAME', 'BROKER', 'SECTOR', 
    'PLACE', 'SELLER_NAME', 'ITEM', 'THICKNESS', 'SIZE', 'GRADE', 'PO_QTY', 
    'PER_TON', 'INV_NO', 'DISPATCH_DATE', 'DISP_QTY', 'PENDING_QTY', 
    'PAYMENT_TERMS', 'DISPATCH_THROUGH', 'STATUS', 'REMARK', 'MOBILE_NO'
]

def parse_size(size_val):
    """Extracts Width and Length from strings like 1500X6300 or 1500*6300."""
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

def safe_convert_date(series):
    """Converts date series safely, coercing invalid years (<2000 or >2099) to NaT."""
    clean_series = series.astype(str).str.strip()
    parsed_dates = pd.to_datetime(clean_series, errors='coerce', format='mixed')
    
    years = parsed_dates.dt.year
    valid_mask = (years >= 2000) & (years <= 2099)
    valid_mask = valid_mask.fillna(False)
    
    return parsed_dates.where(valid_mask, pd.NaT)

def clean_numeric(series):
    """Strips text units ('MT', 'PCS', 'TONS'), commas, and converts to float."""
    clean_s = series.astype(str).str.upper()
    clean_s = clean_s.str.replace(r'[^\d\.\-]', '', regex=True).str.strip()
    return pd.to_numeric(clean_s, errors='coerce').fillna(0)

def clean_data(df):
    """Standardizes column mapping, cleans units, handles text/dates, and derives metrics."""
    renamed_cols = {}
    for col in df.columns:
        clean_col = str(col).strip().upper()
        renamed_cols[col] = COLUMN_MAP.get(clean_col, clean_col)
    df = df.rename(columns=renamed_cols)

    # Ensure missing columns exist
    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = None

    # Safe Date parsing
    df['PO_DATE'] = safe_convert_date(df['PO_DATE'])
    df['DISPATCH_DATE'] = safe_convert_date(df['DISPATCH_DATE'])

    # Clean numbers with embedded units
    num_cols = ['PO_QTY', 'PER_TON', 'DISP_QTY', 'PENDING_QTY', 'THICKNESS']
    for col in num_cols:
        df[col] = clean_numeric(df[col])

    # Clean text columns
    text_cols = [
        'PARTY_NAME', 'BROKER', 'SECTOR', 'PLACE', 'SELLER_NAME', 
        'ITEM', 'GRADE', 'STATUS', 'PO_NO', 'DO_NO', 'PAYMENT_TERMS', 
        'DISPATCH_THROUGH', 'REMARK', 'MOBILE_NO'
    ]
    for col in text_cols:
        df[col] = df[col].astype(str).replace(['nan', 'None', 'NAT', 'N/A', ''], 'Unknown').str.strip()

    # Derived Calculations
    df['TOTAL_REVENUE'] = df['PO_QTY'] * df['PER_TON']
    sizes = df['SIZE'].apply(parse_size)
    df['WIDTH'] = [s[0] for s in sizes]
    df['LENGTH'] = [s[1] for s in sizes]

    return df

@st.cache_data(show_spinner=False)
def load_excel_data(file_bytes):
    """Memory-efficient loader reading Excel bytes directly."""
    file_obj = io.BytesIO(file_bytes)
    excel_file = pd.ExcelFile(file_obj, engine='openpyxl')
    sheet_names = excel_file.sheet_names

    monthly_dfs = []
    pending_df = pd.DataFrame()

    # Matches monthly sheets (e.g. "SEP 2026", "AUG 2026", "APRIL - 2023")
    month_regex = re.compile(
        r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|JANUARY|FEBRUARY|MARCH|APRIL|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)',
        re.IGNORECASE
    )

    for sheet in sheet_names:
        sheet_clean = sheet.strip().upper()
        df_sheet = excel_file.parse(sheet)

        if df_sheet.empty:
            continue

        if "PENDING DISPATCH" in sheet_clean or "PENDING" in sheet_clean:
            pending_df = clean_data(df_sheet)
        elif month_regex.search(sheet_clean):
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
    
    # Primary Metrics
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Unique POs", f"{df['PO_NO'].nunique():,}")
    m2.metric("Unique DOs", f"{df['DO_NO'].nunique():,}")
    m3.metric("Parties", f"{df['PARTY_NAME'].nunique():,}")
    m4.metric("Brokers", f"{df['BROKER'].nunique():,}")
    m5.metric("Sales Persons", f"{df['SELLER_NAME'].nunique():,}")

    m6, m7, m8, m9, m10 = st.columns(5)
    m6.metric("Total PO Qty (MT)", f"{df['PO_QTY'].sum():,.2f}")
    m7.metric("Dispatched Qty (MT)", f"{df['DISP_QTY'].sum():,.2f}")
    m8.metric("Pending Qty (MT)", f"{df['PENDING_QTY'].sum():,.2f}")
    m9.metric("Total Revenue (₹)", f"₹{df['TOTAL_REVENUE'].sum():,.2f}")
    m10.metric("Sectors / Cities", f"{df['SECTOR'].nunique()} / {df['PLACE'].nunique()}")

    st.divider()

    st.subheader("📈 Visual Analytics & Insights")

    # Row 1: Salesperson Performance & Top Pending Parties
    c1, c2 = st.columns(2)
    with c1:
        seller_summary = df.groupby('SELLER_NAME')[['PO_QTY', 'DISP_QTY', 'PENDING_QTY']].sum().reset_index()
        fig_seller = px.bar(
            seller_summary,
            x='SELLER_NAME', y=['PO_QTY', 'DISP_QTY', 'PENDING_QTY'],
            barmode='group', title="Salesperson Performance (Ordered vs Dispatched vs Pending)",
            labels={'value': 'Quantity (MT)', 'SELLER_NAME': 'Sales Person'}
        )
        st.plotly_chart(fig_seller, use_container_width=True)

    with c2:
        party_pending = df.groupby('PARTY_NAME')['PENDING_QTY'].sum().nlargest(10).reset_index()
        fig_party_pending = px.bar(
            party_pending,
            x='PENDING_QTY', y='PARTY_NAME', orientation='h',
            title="Top 10 Parties by Pending Quantity",
            labels={'PENDING_QTY': 'Pending Qty (MT)', 'PARTY_NAME': 'Party Name'}
        )
        fig_party_pending.update_layout(yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig_party_pending, use_container_width=True)

    # Row 2: Status Breakdown & Size Matrix
    c3, c4 = st.columns(2)
    with c3:
        status_summary = df.groupby(['SELLER_NAME', 'STATUS'])['PO_QTY'].sum().reset_index()
        fig_status = px.bar(
            status_summary,
            x='SELLER_NAME', y='PO_QTY', color='STATUS',
            title="Order Status Breakdown by Salesperson (Cancelled / OK / Pending)",
            labels={'PO_QTY': 'Quantity (MT)', 'SELLER_NAME': 'Sales Person'}
        )
        st.plotly_chart(fig_status, use_container_width=True)

    with c4:
        valid_sizes = df[df['WIDTH'].notnull() & df['LENGTH'].notnull()]
        if not valid_sizes.empty:
            fig_size = px.scatter(
                valid_sizes,
                x='WIDTH', y='LENGTH', size='PO_QTY', color='SELLER_NAME',
                hover_data=['PARTY_NAME', 'ITEM', 'GRADE'],
                title="Size Analysis (Width vs Length vs Order Qty)",
                labels={'WIDTH': 'Width (mm)', 'LENGTH': 'Length (mm)', 'PO_QTY': 'PO Qty'}
            )
            st.plotly_chart(fig_size, use_container_width=True)
        else:
            st.info("No valid Width/Length dimension strings found.")

    # Row 3: Broker Channel vs Direct Sales
    c5, c6 = st.columns(2)
    with c5:
        broker_summary = df.groupby('BROKER')['PO_QTY'].sum().nlargest(8).reset_index()
        fig_broker = px.pie(
            broker_summary, values='PO_QTY', names='BROKER',
            title="Top Brokers Channel Share (PO Quantity)",
            hole=0.4
        )
        st.plotly_chart(fig_broker, use_container_width=True)

    with c6:
        sector_summary = df.groupby('SECTOR')['TOTAL_REVENUE'].sum().reset_index()
        fig_sector = px.pie(
            sector_summary, values='TOTAL_REVENUE', names='SECTOR',
            title="Sector-wise Revenue Contribution",
            hole=0.4
        )
        st.plotly_chart(fig_sector, use_container_width=True)

    st.subheader("📋 Detailed Records Table")
    display_cols = [c for c in EXPECTED_COLUMNS if c in df.columns] + ['TOTAL_REVENUE', 'WIDTH', 'LENGTH']
    st.dataframe(df[display_cols], use_container_width=True)


# ---------------------------------------------------------
# 3. STREAMLIT APPLICATION LAYOUT
# ---------------------------------------------------------
st.title("🏭 Sales, Dispatch & Order Tracking Analytics Dashboard")

uploaded_file = st.sidebar.file_uploader("Upload Sales Excel Workbook", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        with st.spinner("Processing Excel workbook (parsing sheets & normalizing data)..."):
            file_bytes = uploaded_file.getvalue()
            df_monthly, df_pending = load_excel_data(file_bytes)

        main_tabs = st.tabs([
            "📅 Date Analytics & Comparison", 
            "📜 Party Ordering History & Recency", 
            "⏳ PENDING DISPATCH Sheet"
        ])

        # ---------------------------------------------------------
        # TAB 1: DATE ANALYTICS & DAY-VS-DAY COMPARISON
        # ---------------------------------------------------------
        with main_tabs[0]:
            if df_monthly.empty:
                st.warning("No monthly sales data found in the uploaded file.")
            else:
                st.sidebar.header("Date Filter Settings")
                mode = st.sidebar.radio("Analysis Mode", ["Single Date Analysis", "Between Two Dates Comparison"])

                valid_dates = df_monthly['DISPATCH_DATE'].dropna()

                if valid_dates.empty:
                    st.warning("No valid dispatch dates found in monthly sheets.")
                else:
                    min_date = valid_dates.min().date()
                    max_date = valid_dates.max().date()

                    if mode == "Single Date Analysis":
                        selected_date = st.sidebar.date_input(
                            "Select Dispatch Date", 
                            value=max_date, 
                            min_value=min_date, 
                            max_value=max_date
                        )
                        filtered_df = df_monthly[df_monthly['DISPATCH_DATE'].dt.date == selected_date]
                        render_kpis_and_charts(filtered_df, title_prefix=f"Date: {selected_date}")

                    else:
                        col_d1, col_d2 = st.sidebar.columns(2)
                        date1 = col_d1.date_input("Start / Baseline Date", value=min_date, min_value=min_date, max_value=max_date)
                        date2 = col_d2.date_input("End / Target Date", value=max_date, min_value=min_date, max_value=max_date)

                        df_range = df_monthly[
                            (df_monthly['DISPATCH_DATE'].dt.date >= date1) & 
                            (df_monthly['DISPATCH_DATE'].dt.date <= date2)
                        ]
                        
                        st.header(f"Range Analysis: {date1} to {date2}")
                        render_kpis_and_charts(df_range, title_prefix=f"Period ({date1} to {date2})")

                        st.divider()
                        st.subheader("🔄 Direct Day-vs-Day Baseline Comparison")
                        
                        df_d1 = df_monthly[df_monthly['DISPATCH_DATE'].dt.date == date1]
                        df_d2 = df_monthly[df_monthly['DISPATCH_DATE'].dt.date == date2]

                        comp_col1, comp_col2 = st.columns(2)
                        with comp_col1:
                            st.markdown(f"### 📅 Baseline: {date1}")
                            st.metric("Total PO Qty", f"{df_d1['PO_QTY'].sum():,.2f} MT")
                            st.metric("Total Dispatch Qty", f"{df_d1['DISP_QTY'].sum():,.2f} MT")
                            st.metric("Total Revenue", f"₹{df_d1['TOTAL_REVENUE'].sum():,.2f}")
                            st.metric("Active Parties", df_d1['PARTY_NAME'].nunique())
                        
                        with comp_col2:
                            diff_po = df_d2['PO_QTY'].sum() - df_d1['PO_QTY'].sum()
                            diff_disp = df_d2['DISP_QTY'].sum() - df_d1['DISP_QTY'].sum()
                            diff_rev = df_d2['TOTAL_REVENUE'].sum() - df_d1['TOTAL_REVENUE'].sum()
                            diff_party = df_d2['PARTY_NAME'].nunique() - df_d1['PARTY_NAME'].nunique()
                            
                            st.markdown(f"### 📅 Target: {date2}")
                            st.metric("Total PO Qty", f"{df_d2['PO_QTY'].sum():,.2f} MT", delta=f"{diff_po:,.2f} MT")
                            st.metric("Total Dispatch Qty", f"{df_d2['DISP_QTY'].sum():,.2f} MT", delta=f"{diff_disp:,.2f} MT")
                            st.metric("Total Revenue", f"₹{df_d2['TOTAL_REVENUE'].sum():,.2f}", delta=f"₹{diff_rev:,.2f}")
                            st.metric("Active Parties", df_d2['PARTY_NAME'].nunique(), delta=diff_party)

        # ---------------------------------------------------------
        # TAB 2: PARTY ORDERING LIFECYCLE & RECENCY ANALYTICS
        # ---------------------------------------------------------
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

                # Format Date Columns for Display
                party_summary['First_Order_Date'] = party_summary['First_Order_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')
                party_summary['Last_Order_Date'] = party_summary['Last_Order_Date'].dt.strftime('%Y-%m-%d').fillna('N/A')

                st.subheader("Over-the-Period Party Summary Table")
                st.dataframe(party_summary, use_container_width=True)

                st.divider()
                st.subheader("🔍 Deep-Dive Party Drilldown")
                selected_party = st.selectbox("Select Party Name", options=sorted(df_monthly['PARTY_NAME'].unique()))
                if selected_party:
                    party_df = df_monthly[df_monthly['PARTY_NAME'] == selected_party]
                    
                    p1, p2, p3, p4 = st.columns(4)
                    p1.metric("Total Orders Placed", party_df['PO_NO'].nunique())
                    p2.metric("Total Ordered Qty", f"{party_df['PO_QTY'].sum():,.2f} MT")
                    p3.metric("Total Dispatched Qty", f"{party_df['DISP_QTY'].sum():,.2f} MT")
                    p4.metric("Total Pending Qty", f"{party_df['PENDING_QTY'].sum():,.2f} MT")

                    st.markdown(f"**Item & Order Lifecycle Details for `{selected_party}`:**")
                    st.dataframe(
                        party_df[['PO_NO', 'PO_DATE', 'SELLER_NAME', 'BROKER', 'ITEM', 'THICKNESS', 'SIZE', 'PO_QTY', 'DISP_QTY', 'PENDING_QTY', 'STATUS', 'REMARK']], 
                        use_container_width=True
                    )
            else:
                st.info("No data available to construct Party ordering history.")

        # ---------------------------------------------------------
        # TAB 3: DEDICATED PENDING DISPATCH REPORT
        # ---------------------------------------------------------
        with main_tabs[2]:
            st.header("⏳ Dedicated Pending Dispatch Report")
            if not df_pending.empty:
                render_kpis_and_charts(df_pending, title_prefix="PENDING DISPATCH SHEET")
            else:
                st.info("No 'PENDING DISPATCH' sheet found in the uploaded file, or the sheet contains no records.")

    except Exception as e:
        st.error(f"An error occurred while parsing the workbook: {str(e)}")

else:
    st.info("👈 Upload your Sales Excel workbook using the sidebar to generate the analytics dashboard.")
