import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import re
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether
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
    'ITEM': ['item', 'item name', 'item_name', 'product', 'discription'],
    'THICKNESS': ['thikness', 'thickness', 'thk', 'thk.'],
    'SIZE': ['size', 'size (mm)'],
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

    numeric_cols = ['PO QTY (MT)', 'PER TON', 'DISP.QTY', 'PENDING']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    if 'PARTY NAME' in df.columns:
        is_total_word = df['PARTY NAME'].astype(str).str.lower().str.contains('total|sum', na=False)
        is_empty_party = df['PARTY NAME'].isna() | (df['PARTY NAME'].astype(str).str.strip() == '') | (df['PARTY NAME'].astype(str).str.lower() == 'nan')
        is_empty_do = df['DO NO'].isna() | (df['DO NO'].astype(str).str.strip() == '') | (df['DO NO'].astype(str).str.lower() == 'nan')
        
        df = df[~(is_total_word | (is_empty_party & is_empty_do))].copy()

    if 'PO DATE' in df.columns:
        po_date_clean = df['PO DATE'].astype(str).str.replace('.', '/', regex=False)
        df['PO DATE'] = pd.to_datetime(po_date_clean, dayfirst=True, errors='coerce')

    if 'DATE' in df.columns:
        disp_date_clean = df['DATE'].astype(str).str.replace('.', '/', regex=False)
        df['DATE'] = pd.to_datetime(disp_date_clean, dayfirst=True, errors='coerce')
    
    df['PO_DATE_STR'] = df['PO DATE'].dt.strftime('%d/%m/%Y').fillna('N/A')
    df['AMOUNT'] = df['PO QTY (MT)'] * df['PER TON']
    
    str_cols = ['PARTY NAME', 'SELLER NAME', 'ITEM', 'STATUS', 'REMARK', 'BROKER', 'SECTOR', 'PLACE', 'PO NO', 'DO NO', 'THICKNESS']
    for c in str_cols:
        df[c] = df[c].fillna('N/A').astype(str).str.strip()

    df['THICKNESS_MM'] = df['THICKNESS'].astype(str).str.replace(r'(?i)\s*mm', '', regex=True).str.strip()
    df['IS_CANCELLED'] = df['STATUS'].str.lower().str.contains('cancel') | df['REMARK'].str.lower().str.contains('cancel')
    df['CANCELLED_QTY'] = np.where(df['IS_CANCELLED'], df['PO QTY (MT)'], 0.0)
    df['ACTIVE_PENDING_QTY'] = np.where(df['IS_CANCELLED'], 0.0, df['PENDING'])

    def parse_width(size_val):
        if pd.isna(size_val) or str(size_val).strip() in ['', 'nan', 'N/A']:
            return "N/A"
        match = re.search(r'(\d+)\s*[xX*]\s*(\d+)', str(size_val))
        if match:
            return match.group(1)
        num = re.findall(r'\d+', str(size_val))
        return num[0] if num else str(size_val)

    df['WIDTH_MM'] = df['SIZE'].apply(parse_width)

    return df

# ---------------------------------------------------------
# Matplotlib Chart Helpers for PDF Export
# ---------------------------------------------------------
def make_pie_chart_bytes(labels, values, title):
    fig, ax = plt.subplots(figsize=(6, 3.2), dpi=200)
    valid_vals = [v if v > 0 else 0 for v in values]
    
    if sum(valid_vals) == 0:
        plt.close(fig)
        return None
        
    wedges, texts, autotexts = ax.pie(
        valid_vals, 
        labels=labels, 
        autopct="%1.1f%%", 
        startangle=90, 
        colors=['#10B981', '#F59E0B', '#EF4444']
    )
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_weight('bold')
        autotext.set_fontsize(8)
        
    ax.set_title(title, fontsize=10, fontweight='bold', color='#1E3A8A', pad=10)
    plt.tight_layout()
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf

def make_bar_chart_bytes(df_data, x_col, y_cols, title, color='#2563EB'):
    if df_data is None or df_data.empty:
        return None
        
    fig, ax = plt.subplots(figsize=(7, 3.5), dpi=200)
    
    if isinstance(y_cols, list):
        x = np.arange(len(df_data[x_col]))
        width = 0.8 / len(y_cols)
        for idx, col in enumerate(y_cols):
            rects = ax.bar(x + idx*width, df_data[col], width, label=col)
            ax.bar_label(rects, fmt='%.1f', padding=2, fontsize=5)
        ax.set_xticks(x + width*(len(y_cols)-1)/2)
        ax.set_xticklabels(df_data[x_col].astype(str), rotation=35, ha='right', fontsize=6)
        ax.legend(fontsize=6, loc='upper right')
    else:
        rects = ax.bar(df_data[x_col].astype(str), df_data[y_cols], color=color, width=0.45)
        ax.bar_label(rects, fmt='%.1f', padding=2, fontsize=6.5, fontweight='bold')
        plt.xticks(rotation=35, ha='right', fontsize=6.5)

    ax.set_title(title, fontsize=10, fontweight='bold', color='#1E3A8A', pad=10)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    plt.tight_layout()
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf

def make_comparison_bar_chart_bytes(label1, label2, kpi1, kpi2, title):
    fig, ax = plt.subplots(figsize=(7, 3.5), dpi=200)
    categories = ["Ordered Qty", "Dispatched Qty", "Cancelled Qty", "Pending Qty"]
    y1 = [kpi1['Total PO Quantity (MT)'], kpi1['Dispatched Qty (MT)'], kpi1['Cancelled Qty (MT)'], kpi1['Pending Qty (MT)']]
    y2 = [kpi2['Total PO Quantity (MT)'], kpi2['Dispatched Qty (MT)'], kpi2['Cancelled Qty (MT)'], kpi2['Pending Qty (MT)']]
    
    x = np.arange(len(categories))
    width = 0.35
    
    rects1 = ax.bar(x - width/2, y1, width, label=str(label1), color='#2563EB')
    rects2 = ax.bar(x + width/2, y2, width, label=str(label2), color='#10B981')
    
    ax.bar_label(rects1, fmt='%.1f', padding=2, fontsize=5)
    ax.bar_label(rects2, fmt='%.1f', padding=2, fontsize=5)
    
    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=7)
    ax.set_title(title, fontsize=10, fontweight='bold', color='#1E3A8A', pad=10)
    ax.legend(fontsize=7, loc='upper right')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    plt.tight_layout()
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf

# ---------------------------------------------------------
# Dynamic PDF Generator
# ---------------------------------------------------------
def generate_exact_screen_pdf(sheet_name, kpis, chart_buffers, tables_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
    )
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#1E3A8A'), spaceAfter=8
    )
    section_style = ParagraphStyle(
        'DocSection', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor('#1E40AF'), spaceBefore=14, spaceAfter=6
    )
    cell_style = ParagraphStyle(
        'TableCell', parent=styles['Normal'], fontSize=7, leading=8.5, alignment=0
    )
    cell_header = ParagraphStyle(
        'HeaderCell', parent=styles['Normal'], fontSize=7.5, leading=9, textColor=colors.whitesmoke, fontName="Helvetica-Bold", alignment=0
    )
    
    story.append(Paragraph(f"<b>📊 Sales & Dispatch Report — {sheet_name}</b>", title_style))
    story.append(Spacer(1, 6))
    
    if kpis:
        story.append(Paragraph("<b>1. Key Performance Indicators (KPIs)</b>", section_style))
        kpi_items = []
        for k, v in kpis.items():
            if isinstance(v, float):
                val_str = f"{v:,.2f}"
            elif isinstance(v, int):
                val_str = f"{v:,}"
            else:
                val_str = str(v)
            kpi_items.append((k, val_str))
        
        kpi_matrix = []
        for i in range(0, len(kpi_items), 4):
            chunk = kpi_items[i:i+4]
            row_titles = [Paragraph(f"<b>{item[0]}</b>", cell_header) for item in chunk]
            row_vals = [Paragraph(f"<b>{item[1]}</b>", cell_style) for item in chunk]
            
            while len(row_titles) < 4:
                row_titles.append(Paragraph("", cell_header))
                row_vals.append(Paragraph("", cell_style))
                
            kpi_matrix.append(row_titles)
            kpi_matrix.append(row_vals)
            
        t_kpi = Table(kpi_matrix, colWidths=[135]*4)
        t_kpi.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2563EB')),
            ('BACKGROUND', (0,2), (-1,2), colors.HexColor('#2563EB')),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
        ]))
        story.append(t_kpi)
        story.append(Spacer(1, 10))

    if chart_buffers:
        story.append(Paragraph("<b>2. Visual Analytics</b>", section_style))
        for fig_title, buf in chart_buffers.items():
            if buf is not None:
                story.append(KeepTogether([
                    Paragraph(f"<b>{fig_title}</b>", ParagraphStyle('SubHead', parent=styles['Normal'], fontSize=9, fontName="Helvetica-Bold", textColor=colors.HexColor('#1E40AF'))),
                    Spacer(1, 3),
                    Image(buf, width=520, height=230),
                    Spacer(1, 8)
                ]))

    for title, df_table in tables_dict.items():
        if df_table is not None and not df_table.empty:
            story.append(Paragraph(f"<b>{title}</b>", section_style))
            
            sub_df = df_table.copy().reset_index(drop=True)
            cols = sub_df.columns.tolist()
            
            table_data = [[Paragraph(f"<b>{col}</b>", cell_header) for col in cols]]
            for row in sub_df.values.tolist():
                formatted_row = []
                for val in row:
                    val_str = f"{val:,.2f}" if isinstance(val, (float, np.floating)) else str(val)
                    formatted_row.append(Paragraph(val_str, cell_style))
                table_data.append(formatted_row)
            
            col_width = 540 / max(len(cols), 1)
            t_data = Table(table_data, colWidths=[col_width]*len(cols), repeatRows=1)
            t_data.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F9FAFB')]),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
            ]))
            
            story.append(t_data)
            story.append(Spacer(1, 10))
            
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

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
    total_po = data['PO NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique()
    total_do = data['DO NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique()
    num_parties = data['PARTY NAME'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique()
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
        
        display_cols = [c for c in ['PO NO', 'DO NO', 'PO_DATE_STR', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'THICKNESS', 'SIZE', 'PO QTY (MT)', 'PER TON', 'DISP.QTY', 'PENDING', 'STATUS', 'REMARK'] if c in filtered_df.columns]
        st.dataframe(filtered_df[display_cols], use_container_width=True)

        st.markdown("---")
        st.subheader("📥 Download Date-Wise PDF Report")
        
        pdf_table_cols = [c for c in ['PO NO', 'DO NO', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'PO QTY (MT)', 'DISP.QTY', 'PENDING', 'STATUS'] if c in filtered_df.columns]
        date_pdf_bytes = generate_exact_screen_pdf(
            f"{selected_sheet} - {selected_date_str}",
            d_kpis,
            {},
            {f"Orders for Date {selected_date_str}": filtered_df[pdf_table_cols]}
        )
        st.download_button(
            f"📥 Download Report for {selected_date_str}",
            data=date_pdf_bytes,
            file_name=f"Report_{selected_sheet}_{selected_date_str.replace('/', '-')}.pdf",
            mime="application/pdf"
        )

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

        st.markdown("---")
        st.subheader("📥 Download Comparison PDF Report")
        
        comp_chart_buf = make_comparison_bar_chart_bytes(
            label1, label2, kpi1, kpi2, f"Comparison: {label1} vs {label2}"
        )
        
        comp_pdf_bytes = generate_exact_screen_pdf(
            f"Comparison ({label1} vs {label2})",
            {f"Summary": f"Comparing {label1} and {label2}"},
            {"Comparison Chart": comp_chart_buf},
            {"Comparison Metrics Table": comp_df}
        )
        st.download_button(
            "📥 Download Comparison PDF Report",
            data=comp_pdf_bytes,
            file_name=f"Comparison_Report_{str(label1).replace('/', '-')}_vs_{str(label2).replace('/', '-')}.pdf",
            mime="application/pdf"
        )

# =========================================================
# SECTION 2: ALL SALES & DISPATCH ANALYTICS
# =========================================================
elif section == "📊 All Sales & Dispatch Analytics":
    selected_sheet = st.sidebar.selectbox("Select Month / Sheet", sheet_names)
    df = load_and_clean_sheet(uploaded_file, selected_sheet)
    
    st.header(f"All Sales & Dispatch Analytics — {selected_sheet}")
    
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
    
    fig_kpi = go.Figure(data=[go.Pie(
        labels=['Dispatched Qty', 'Cancelled Qty', 'Pending Qty'],
        values=[kpis['Dispatched Qty (MT)'], kpis['Cancelled Qty (MT)'], kpis['Pending Qty (MT)']],
        hole=.4,
        textinfo='label+value+percent',
        texttemplate='%{label}<br>%{value:,.2f} MT (%{percent})',
        marker_colors=['#10B981', '#F59E0B', '#EF4444']
    )])
    fig_kpi.update_layout(title="Overall Status Breakdown")
    st.plotly_chart(fig_kpi, use_container_width=True)
    
    st.markdown("---")
    
    st.subheader("2. Sales Executive Analytics")
    sp_item_grp = df.groupby(['SELLER NAME', 'ITEM']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        Dispatched=('DISP.QTY', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum'),
        Pending=('ACTIVE_PENDING_QTY', 'sum')
    ).reset_index()
    st.write("**Sales Executive Item-Wise Breakdown**")
    st.dataframe(sp_item_grp, use_container_width=True)
    
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
        sc_summary = pd.DataFrame()
        st.info("No Short Closed ('SC') orders found.")

    st.markdown("---")
    
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
    st.dataframe(filtered_party_grp, use_container_width=True)
    
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
    
    st.subheader("4. Detailed Item, Thickness & Width Summary")
    item_spec_grp = df.groupby(['ITEM', 'THICKNESS_MM', 'WIDTH_MM']).agg(
        OrderedQty=('PO QTY (MT)', 'sum'),
        DispatchedQty=('DISP.QTY', 'sum'),
        CancelledQty=('CANCELLED_QTY', 'sum'),
        PendingQty=('ACTIVE_PENDING_QTY', 'sum')
    ).reset_index()
    
    item_spec_grp.rename(columns={
        'ITEM': 'Item Name',
        'THICKNESS_MM': 'Thickness (mm)',
        'WIDTH_MM': 'Width (mm)',
        'OrderedQty': 'Ordered Qty (MT)',
        'DispatchedQty': 'Dispatched Qty (MT)',
        'CancelledQty': 'Cancelled Qty (MT)',
        'PendingQty': 'Pending Qty (MT)'
    }, inplace=True)

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        all_items_list = sorted(item_spec_grp['Item Name'].unique().tolist())
        selected_items = st.multiselect("Filter Item Name(s)", all_items_list, default=[])
    with col_f2:
        all_thickness_list = sorted(item_spec_grp['Thickness (mm)'].unique().tolist(), key=lambda x: str(x))
        selected_thicknesses = st.multiselect("Filter Thickness (mm)", all_thickness_list, default=[])
    with col_f3:
        all_width_list = sorted(item_spec_grp['Width (mm)'].unique().tolist(), key=lambda x: str(x))
        selected_widths = st.multiselect("Filter Width (mm)", all_width_list, default=[])

    display_spec_table = item_spec_grp.copy()
    if selected_items:
        display_spec_table = display_spec_table[display_spec_table['Item Name'].isin(selected_items)]
    if selected_thicknesses:
        display_spec_table = display_spec_table[display_spec_table['Thickness (mm)'].isin(selected_thicknesses)]
    if selected_widths:
        display_spec_table = display_spec_table[display_spec_table['Width (mm)'].isin(selected_widths)]

    st.dataframe(display_spec_table, use_container_width=True)

    st.markdown("---")
    st.subheader("📄 Download Dashboard PDF Report")
    
    chart_bufs = {
        "Overall Status Breakdown": make_pie_chart_bytes(
            ['Dispatched Qty', 'Cancelled Qty', 'Pending Qty'],
            [kpis['Dispatched Qty (MT)'], kpis['Cancelled Qty (MT)'], kpis['Pending Qty (MT)']],
            "Overall Status Breakdown"
        ),
        "Dispatched Qty by Sales Person": make_bar_chart_bytes(
            sp_disp, 'SELLER NAME', 'DISP.QTY', "Dispatched Qty by Sales Person", color='#10B981'
        ),
        "Order Received Qty by Sales Person": make_bar_chart_bytes(
            sp_rec, 'SELLER NAME', 'PO QTY (MT)', "Order Received Qty by Sales Person", color='#3B82F6'
        ),
        "Top Parties Breakdown": make_bar_chart_bytes(
            filtered_party_grp.head(10), 'PARTY NAME', ['OrderedQty', 'DispatchedQty', 'PendingQty'], "Top Parties Breakdown"
        )
    }
    
    tables_to_pdf = {
        "Sales Executive Performance Breakdown": sp_item_grp,
        "Cancelled Orders per Sales Person": sp_cancelled,
        "Pending Orders per Sales Person": sp_pending,
        "Party Wise Summary": party_grp,
        "Filtered Item, Thickness & Width Summary": display_spec_table
    }
    
    pdf_bytes = generate_exact_screen_pdf(selected_sheet, kpis, chart_bufs, tables_to_pdf)
    if pdf_bytes:
        st.download_button(
            "📥 Download Complete PDF Report", 
            data=pdf_bytes, 
            file_name=f"Dashboard_Report_{selected_sheet}.pdf", 
            mime="application/pdf"
        )

# =========================================================
# SECTION 3: PENDING DISPATCH
# =========================================================
elif section == "🚚 Pending Dispatch":
    st.header("🚚 Pending Dispatch Standalone Report")
    
    pd_sheet_candidates = [s for s in sheet_names if 'pending' in s.lower() and 'dispatch' in s.lower()]
    target_pd_sheet = pd_sheet_candidates[0] if pd_sheet_candidates else (sheet_names[0] if sheet_names else None)
    
    pd_sheet = st.sidebar.selectbox("Select Pending Dispatch Sheet", sheet_names, index=sheet_names.index(target_pd_sheet) if target_pd_sheet in sheet_names else 0)
    
    df_pd = load_and_clean_sheet(uploaded_file, pd_sheet)
    active_pd = df_pd[(df_pd['ACTIVE_PENDING_QTY'] > 0) & (~df_pd['IS_CANCELLED'])].copy()
    
    pd_kpis = {
        'Delivery Orders (DOs)': active_pd['DO NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique(),
        'Purchase Orders (POs)': active_pd['PO NO'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique(),
        'Parties Impacted': active_pd['PARTY NAME'].replace(['N/A', 'Unknown', 'nan'], np.nan).dropna().nunique(),
        'Pending Line Items': len(active_pd),
        'Total Pending Qty (MT)': active_pd['ACTIVE_PENDING_QTY'].sum(),
        'Est. Pending Value (₹)': (active_pd['ACTIVE_PENDING_QTY'] * active_pd['PER TON']).sum()
    }
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Pending DOs Count", f"{pd_kpis['Delivery Orders (DOs)']:,}")
    c2.metric("Pending POs Count", f"{pd_kpis['Purchase Orders (POs)']:,}")
    c3.metric("Parties Impacted", f"{pd_kpis['Parties Impacted']:,}")

    c4, c5, c6 = st.columns(3)
    c4.metric("Pending Line Items", f"{pd_kpis['Pending Line Items']:,}")
    c5.metric("Total Pending Qty", f"{pd_kpis['Total Pending Qty (MT)']:,.2f} MT")
    c6.metric("Est. Pending Value", f"₹{pd_kpis['Est. Pending Value (₹)']:,.2f}")
    
    st.markdown("---")
    st.subheader("Pending Quantity Breakdown by Sales Person")
    
    sp_pd_df = active_pd.groupby('SELLER NAME')['ACTIVE_PENDING_QTY'].sum().reset_index()
    
    fig_pd = px.bar(
        sp_pd_df,
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
    
    pending_details_df = active_pd[['PO NO', 'DO NO', 'PO_DATE_STR', 'PARTY NAME', 'SELLER NAME', 'ITEM', 'THICKNESS_MM', 'SIZE', 'PO QTY (MT)', 'DISP.QTY', 'ACTIVE_PENDING_QTY', 'REMARK']].copy()
    pending_details_df.rename(columns={'THICKNESS_MM': 'THICKNESS (mm)', 'ACTIVE_PENDING_QTY': 'PENDING QTY', 'PO_DATE_STR': 'PO DATE'}, inplace=True)
    st.dataframe(pending_details_df, use_container_width=True)

    st.markdown("---")
    st.subheader("📥 Export Pending Dispatch PDF")
    
    pd_chart_bufs = {
        "Pending Quantity Breakdown": make_bar_chart_bytes(
            sp_pd_df, 'SELLER NAME', 'ACTIVE_PENDING_QTY', "Pending Dispatch Qty by Sales Person", color='#EF4444'
        )
    }
    
    pdf_bytes = generate_exact_screen_pdf(
        pd_sheet,
        pd_kpis,
        pd_chart_bufs,
        {"Pending Dispatch Details": pending_details_df}
    )
    if pdf_bytes:
        st.download_button(
            "📥 Download Pending Dispatch PDF",
            data=pdf_bytes,
            file_name=f"Pending_Dispatch_{pd_sheet}.pdf",
            mime="application/pdf"
        )
