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
    'S_NO': 'S_NO', 'SR NO': 'S_NO', 'S.NO': 'S_NO', 'S. NO.': 'S_NO', 'SIR NO.': 'S_NO', 'SR NO.': 'S_NO',
    'PO NO': 'PO_NO', 'PO NO.': 'PO_NO', 'PO NUMBER': 'PO_NO',
    'DO .NO.': 'DO_NO', 'DO NO': 'DO_NO', 'DO NO.': 'DO_NO', 'DO .NO': 'DO_NO', 'DO NO ': 'DO_NO',
    'PO DATE': 'PO_DATE', 'ORDER DATE': 'PO_DATE',
    'PARTY NAME': 'PARTY_NAME', 'CUSTOMER NAME': 'PARTY_NAME',
    'BROKER': 'BROKER',
    'SECTOR': 'SECTOR',
    'PLACE': 'PLACE', 'CITY': 'PLACE',
    'SELLER NAME': 'SELLER_NAME', 'SALES PERSON': 'SELLER_NAME',
    'ITEM': 'ITEM', 'DISCRIPTION': 'ITEM', 'DESCRIPTION': 'ITEM',
    'THIKNESS': 'THICKNESS', 'THICKNESS': 'THICKNESS',
    'SIZE': 'SIZE', 'SIZE (MM)': 'SIZE', 'SIZE ()': 'SIZE',
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
    """Converts mixed date formats and timezones cleanly to naive datetime objects."""
    clean_series = series.astype(str).str.strip()
    parsed_dates = pd.to_datetime(clean_series, errors='coerce', dayfirst=True, utc=True)
    parsed_dates = parsed_dates.dt.tz_localize(None)
    years = parsed_dates.dt.year
    valid_mask = (years >= 2020) & (years <= 2035)
    valid_mask = valid_mask.fillna(False)
    return parsed_dates.where(valid_mask, pd.NaT)

def clean_numeric(series):
    clean_s = series.astype(str).str.upper()
    clean_s = clean_s.str.replace(r'[^\d\.\-]', '', regex=True).str.strip()
    return pd.to_numeric(clean_s, errors='coerce').fillna(0)

def locate_header_and_read(excel_file, sheet_name):
    """
    Locates row containing table header and enforces Column 0 as S_NO 
    even when Excel headers contain timestamps like 16:02, 15:35, etc.
    """
    df_raw = excel_file.parse(sheet_name, header=None).dropna(how='all')
    header_row_idx = None
    
    for idx, row in df_raw.iterrows():
        row_str = ' '.join(row.dropna().astype(str)).upper()
        if 'PARTY NAME' in row_str or ('PO NO' in row_str and 'DO' in row_str):
            header_row_idx = idx
            break
            
    if header_row_idx is not None:
        headers = df_raw.loc[header_row_idx].values
        df_data = df_raw.loc[header_row_idx + 1:].copy()
        
        # Positional header fallback: First column in table is ALWAYS S_NO
        headers_list = []
        for i, h in enumerate(headers):
            h_str = str(h).strip().upper() if pd.notna(h) else ''
            if i == 0:
                headers_list.append('S_NO')
            else:
                headers_list.append(h_str if h_str else f"UNNAMED_{i}")
                
        df_data.columns = headers_list
        return df_data
    return pd.DataFrame()

def clean_data(df):
    renamed_cols = {}
    for col in df.columns:
        clean_col = str(col).strip().upper()
        renamed_cols[col] = COLUMN_MAP.get(clean_col, clean_col)
    df = df.rename(columns=renamed_cols)

    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df['PO_DATE'] = safe_convert_date(df['PO_DATE'])
    df['DISPATCH_DATE'] = safe_convert_date(df['DISPATCH_DATE'])

    num_cols = ['PO_QTY', 'PER_TON', 'DISP_QTY', 'PENDING_QTY', 'THICKNESS']
    for col in num_cols:
        df[col] = clean_numeric(df[col])

    text_cols = [
        'S_NO', 'PARTY_NAME', 'BROKER', 'SECTOR', 'PLACE', 'SELLER_NAME', 
        'ITEM', 'GRADE', 'STATUS', 'PO_NO', 'DO_NO', 'PAYMENT_TERMS', 
        'DISPATCH_THROUGH', 'REMARK', 'MOBILE_NO'
    ]
    for col in text_cols:
        df[col] = df[col].astype(str).replace(['nan', 'None', 'NAT', 'N/A', ''], '').str.strip()

    df['TOTAL_REVENUE'] = df['PO_QTY'] * df['PER_TON']
    sizes = df['SIZE'].apply(parse_size)
    df['WIDTH'] = [s[0] for s in sizes]
    df['LENGTH'] = [s[1] for s in sizes]

    return df

@st.cache_data(show_spinner=False)
def load_excel_data(file_bytes):
    file_obj = io.BytesIO(file_bytes)
    excel_file = pd.ExcelFile(file_obj, engine='openpyxl')
    sheet_names = excel_file.sheet_names

    monthly_dfs = []
    pending_df = pd.DataFrame()

    month_regex = re.compile(
        r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|JANUARY|FEBRUARY|MARCH|APRIL|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)',
        re.IGNORECASE
    )

    for sheet in sheet_names:
        sheet_clean = sheet.strip().upper()
        df_sheet = locate_header_and_read(excel_file, sheet)

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
# PDF REPORT GENERATOR
# ---------------------------------------------------------
def generate_simple_pdf_report(df, title="Sales Analytics Summary"):
    summary_text = f"""
    {title.upper()}
    =======================================================
    Total Orders: {len(df):,}
    Unique POs: {df['PO_NO'].nunique():,}
    Unique Customers: {df['PARTY_NAME'].nunique():,}
    Total PO Quantity (MT): {df['PO_QTY'].sum():,.2f}
    Total Dispatched Quantity (MT): {df['DISP_QTY'].sum():,.2f}
    Total Pending Quantity (MT): {df['PENDING_QTY'].sum():,.2f}
    Total Revenue: ₹{df['TOTAL_REVENUE'].sum():,.2f}
    =======================================================
    """
    return summary_text.encode('utf-8')


# ---------------------------------------------------------
# 2. RENDER KPI DASHBOARD & CHARTS
# ---------------------------------------------------------
def render_kpis_and_charts(df, title_prefix=""):
    if df.empty:
        st.warning(f"No records found for {title_prefix}.")
        return

    st.subheader(f"📊 {title_prefix} KPI Overview")
    
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

    st.subheader("📋 Detailed Records Table")
    display_cols = [c for c in EXPECTED_COLUMNS if c in df.columns] + ['TOTAL_REVENUE', 'WIDTH', 'LENGTH']
    st.dataframe(df[display_cols], use_container_width=True)


# ---------------------------------------------------------
# 3. APPLICATION MAIN LAYOUT
# ---------------------------------------------------------
st.title("🏭 Sales, Dispatch & Order Tracking Analytics Dashboard")

st.sidebar.header("📁 Data Source")
uploaded_file = st.sidebar.file_uploader("Upload Sales Excel Workbook", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        with st.spinner("Processing Excel workbook..."):
            file_bytes = uploaded_file.getvalue()
            df_monthly, df_pending = load_excel_data(file_bytes)

        st.sidebar.divider()
        st.sidebar.header("⚙️ Report & Controls")

        # Sidebar Controls
        chk_compare_dates = st.sidebar.checkbox("Compare Data Between Dates", value=False)
        chk_view_pending = st.sidebar.checkbox("View Pending Dispatch Sheet", value=False)
        chk_overall_comparison = st.sidebar.checkbox("Overall Comparison Analytics", value=False)

        # Date Filtering
        valid_dates = df_monthly['DISPATCH_DATE'].dropna() if not df_monthly.empty else pd.Series()

        if not valid_dates.empty:
            min_d, max_d = valid_dates.min().date(), valid_dates.max().date()

            st.sidebar.subheader("📅 Date Selection")
            if not chk_compare_dates:
                selected_date = st.sidebar.date_input("Select Report Date", value=max_d, min_value=min_d, max_value=max_d)
            else:
                col_d1, col_d2 = st.sidebar.columns(2)
                date_start = col_d1.date_input("Start Date", value=min_d, min_value=min_d, max_value=max_d)
                date_end = col_d2.date_input("End Date", value=max_d, min_value=min_d, max_value=max_d)

        # PDF Download Section
        st.sidebar.divider()
        st.sidebar.subheader("📥 Download Report")
        if st.sidebar.button("📄 Generate PDF Report"):
            pdf_data = generate_simple_pdf_report(df_monthly, title="Sales Analytics Summary")
            st.sidebar.download_button(
                label="⬇️ Download PDF Summary",
                data=pdf_data,
                file_name="Sales_Analytics_Summary.txt",
                mime="text/plain"
            )

        # ---------------------------------------------------------
        # MAIN VIEW ROUTING
        # ---------------------------------------------------------
        if chk_view_pending:
            st.header("⏳ Dedicated Pending Dispatch Report")
            if not df_pending.empty:
                render_kpis_and_charts(df_pending, title_prefix="PENDING DISPATCH SHEET")
            else:
                st.warning("No records found in Pending Dispatch sheet.")

        elif chk_overall_comparison:
            st.header("🌐 Overall Cross-Month Comparison Analytics")
            if not df_monthly.empty:
                render_kpis_and_charts(df_monthly, title_prefix="OVERALL ALL-MONTHS")
            else:
                st.warning("No monthly sales data found.")

        elif chk_compare_dates:
            st.header(f"🔄 Comparative Analysis: {date_start} to {date_end}")
            if not df_monthly.empty:
                df_range = df_monthly[(df_monthly['DISPATCH_DATE'].dt.date >= date_start) & (df_monthly['DISPATCH_DATE'].dt.date <= date_end)]
                render_kpis_and_charts(df_range, title_prefix=f"Period ({date_start} to {date_end})")

        else:
            st.header(f"📅 Daily Report: {selected_date}")
            if not df_monthly.empty:
                df_single = df_monthly[df_monthly['DISPATCH_DATE'].dt.date == selected_date]
                if not df_single.empty:
                    render_kpis_and_charts(df_single, title_prefix=f"Date {selected_date}")
                else:
                    st.info(f"No dispatch records found on {selected_date}. Showing overall dataset preview below:")
                    render_kpis_and_charts(df_monthly, title_prefix="FULL DATASET PREVIEW")

    except Exception as e:
        st.error(f"Error processing workbook: {str(e)}")

else:
    st.info("👈 Upload your Sales Excel workbook using the sidebar to generate the report.")
