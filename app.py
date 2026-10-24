import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(
    page_title="ANM FASHION - Accounts Payable & Yarn Report",
    page_icon="🧵",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { background-color: #0e1117; }
    .company-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #0f172a 100%);
        padding: 24px;
        border-radius: 12px;
        border-bottom: 4px solid #f59e0b;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
    }
    .company-title {
        font-size: 32px !important;
        font-weight: 800 !important;
        letter-spacing: 1.5px;
        color: #ffffff;
        margin: 0;
    }
    .company-subtitle {
        font-size: 18px !important;
        color: #f59e0b;
        font-weight: 600;
        margin-top: 5px;
    }
    .kpi-card {
        background: #1e293b;
        padding: 18px;
        border-radius: 10px;
        border-left: 5px solid #3b82f6;
        box-shadow: 0 2px 10px rgba(0,0,0,0.2);
    }
    .kpi-card-gold { border-left-color: #f59e0b !important; }
    .kpi-card-green { border-left-color: #10b981 !important; }
    .kpi-card-red { border-left-color: #ef4444 !important; }
    .kpi-title {
        color: #94a3b8;
        font-size: 13px;
        font-weight: 600;
        text-transform: uppercase;
    }
    .kpi-value {
        color: #ffffff;
        font-size: 24px;
        font-weight: 700;
        margin-top: 5px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="company-header">
        <div class="company-title">ANM FASHION</div>
        <div class="company-subtitle">🧵 YARN PAYABLE REPORT & ACCOUNTS DASHBOARD</div>
    </div>
""", unsafe_allow_html=True)

st.sidebar.header("⚙️ Settings & Controls")
header_row = st.sidebar.number_input("Header Row Index (0 = 1st line, 2 = 3rd line)", min_value=0, max_value=10, value=2, step=1)

uploaded_file = st.sidebar.file_uploader("Upload Excel File (.xlsx, .xls)", type=["xlsx", "xls"])

def parse_month_year(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip().upper()
    try:
        dt = pd.to_datetime(val_str, errors='coerce')
        if not pd.isna(dt):
            return dt.strftime('%b-%Y').upper()
    except:
        pass
        
    months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    for idx, m in enumerate(months):
        if m in val_str:
            import re
            years = re.findall(r'\d{4}|\d{2}', val_str)
            if years:
                yr = years[-1]
                if len(yr) == 2:
                    yr = "20" + yr
                return f"{m}-{yr}"
    return val_str

def sort_months_chronologically(month_list):
    def get_sort_key(m_str):
        try:
            return datetime.strptime(str(m_str), "%b-%Y")
        except:
            try:
                return datetime.strptime(str(m_str), "%B-%Y")
            except:
                return datetime(2099, 12, 31)
    return sorted(month_list, key=get_sort_key)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, header=header_row)
        df.columns = [str(c).strip() for c in df.columns]
        
        st.sidebar.success("✅ File Loaded Successfully!")
        
        cols_lower = {c.lower(): c for c in df.columns}
        
        default_party = next((v for k, v in cols_lower.items() if 'party' in k or 'supplier' in k), df.columns[0])
        default_month = next((v for k, v in cols_lower.items() if 'month' in k or 'date' in k), df.columns[1] if len(df.columns) > 1 else df.columns[0])
        
        default_invoice = next((v for k, v in cols_lower.items() if k == 'invoice' or 'total amount' in k), df.columns[-3] if len(df.columns)>=3 else df.columns[-1])
        default_paid = next((v for k, v in cols_lower.items() if k == 'paid'), None)
        default_balance = next((v for k, v in cols_lower.items() if k == 'balance'), None)
        default_qty = next((v for k, v in cols_lower.items() if k == 'qty' or 'quantity' in k or 'bags' in k), None)

        st.sidebar.subheader("📌 Column Mapping")
        party_col = st.sidebar.selectbox("Party Name Column", df.columns, index=list(df.columns).index(default_party))
        month_col = st.sidebar.selectbox("Month Column", df.columns, index=list(df.columns).index(default_month))
        
        invoice_col = st.sidebar.selectbox("Invoice Amount Column", df.columns, index=list(df.columns).index(default_invoice))
        
        paid_options = [None] + list(df.columns)
        paid_index = (list(df.columns).index(default_paid) + 1) if default_paid and default_paid in df.columns else 0
        paid_col = st.sidebar.selectbox("Paid Amount Column", paid_options, index=paid_index)

        balance_options = [None] + list(df.columns)
        balance_index = (list(df.columns).index(default_balance) + 1) if default_balance and default_balance in df.columns else 0
        balance_col = st.sidebar.selectbox("Balance Amount Column (Target for Summary)", balance_options, index=balance_index)

        qty_options = [None] + list(df.columns)
        qty_index = (list(df.columns).index(default_qty) + 1) if default_qty and default_qty in df.columns else 0
        qty_col = st.sidebar.selectbox("Yarn Quantity Column (Optional)", qty_options, index=qty_index)

        # Data Cleaning
        df_clean = df.copy()
        df_clean[party_col] = df_clean[party_col].fillna('').astype(str).str.strip()
        df_clean = df_clean[~df_clean[party_col].isin(['None', 'nan', '', 'TOTAL', 'GRAND TOTAL', '0'])]
        
        df_clean[invoice_col] = pd.to_numeric(df_clean[invoice_col], errors='coerce').fillna(0)
        
        if paid_col:
            df_clean[paid_col] = pd.to_numeric(df_clean[paid_col], errors='coerce').fillna(0)
            
        if balance_col:
            df_clean[balance_col] = pd.to_numeric(df_clean[balance_col], errors='coerce').fillna(0)
        else:
            df_clean['CALC_BALANCE'] = df_clean[invoice_col] - (df_clean[paid_col] if paid_col else 0)
            balance_col = 'CALC_BALANCE'

        if qty_col:
            df_clean[qty_col] = pd.to_numeric(df_clean[qty_col], errors='coerce').fillna
