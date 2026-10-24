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

st.sidebar.header("⚙️ Settings & Filters")
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
        default_amount = next((v for k, v in cols_lower.items() if 'invoice amount' in k or 'total' in k or 'amount' in k), df.columns[-1])
        default_qty = next((v for k, v in cols_lower.items() if 'qty' in k or 'quantity' in k or 'bags' in k), None)

        st.sidebar.subheader("📌 Column Mapping")
        party_col = st.sidebar.selectbox("Party Name Column", df.columns, index=list(df.columns).index(default_party))
        month_col = st.sidebar.selectbox("Month Column", df.columns, index=list(df.columns).index(default_month))
        amount_col = st.sidebar.selectbox("Invoice Amount Column", df.columns, index=list(df.columns).index(default_amount))
        
        qty_options = [None] + list(df.columns)
        qty_index = (list(df.columns).index(default_qty) + 1) if default_qty and default_qty in df.columns else 0
        qty_col = st.sidebar.selectbox("Yarn Quantity Column (Optional)", qty_options, index=qty_index)

        df_clean = df.copy()
        df_clean[party_col] = df_clean[party_col].fillna('').astype(str).str.strip()
        df_clean = df_clean[~df_clean[party_col].isin(['None', 'nan', '', 'TOTAL', 'GRAND TOTAL', '0'])]
        
        df_clean[amount_col] = pd.to_numeric(df_clean[amount_col], errors='coerce').fillna(0)
        if qty_col:
            df_clean[qty_col] = pd.to_numeric(df_clean[qty_col], errors='coerce').fillna(0)

        df_clean['STD_MONTH'] = df_clean[month_col].apply(parse_month_year)

        st.sidebar.markdown("---")
        st.sidebar.subheader("🔍 Search & Filter Parties")
        unique_parties = sorted([str(p) for p in df_clean[party_col].unique() if str(p).strip()])
        selected_parties = st.sidebar.multiselect("Select Parties to Display", unique_parties, default=unique_parties)

        if selected_parties:
            df_clean = df_clean[df_clean[party_col].isin(selected_parties)]

        tot_inv = df_clean[amount_col].sum()
        tot_qty = df_clean[qty_col].sum() if qty_col else 0

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"""
                <div class="kpi-card kpi-card-gold">
                    <div class="kpi-title">TOTAL YARN PAYABLE</div>
                    <div class="kpi-value">PKR {tot_inv:,.0f}</div>
                </div>
            """, unsafe_allow_html=True)
        
        with c2:
            st.markdown(f"""
                <div class="kpi-card kpi-card-green">
                    <div class="kpi-title">TOTAL YARN RECEIVED</div>
                    <div class="kpi-value">{tot_qty:,.0f} Bags/Kg</div>
                </div>
            """, unsafe_allow_html=True)

        with c3:
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">TOTAL RECORDS / PARTIES</div>
                    <div class="kpi-value">{len(df_clean)} Rows ({len(selected_parties)} Parties)</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        tab_pivot, tab_details = st.tabs(["📊 Month-wise Pivot Summary", "🧾 Invoice-wise Detailed Report"])

        with tab_pivot:
            st.subheader("Party-wise & Chronological Month-wise Matrix Pivot")
            
            pivot = pd.pivot_table(
                df_clean,
                index=party_col,
                columns='STD_MONTH',
                values=amount_col,
                aggfunc='sum',
                fill_value=0
            )

            sorted_months = sort_months_chronologically(pivot.columns)
            pivot = pivot[sorted_months]

            pivot['TOTAL'] = pivot.sum(axis=1)
            total_row = pivot.sum(axis=0)
            total_row.name = 'GRAND TOTAL'
            pivot_with_total = pd.concat([pivot, pd.DataFrame(total_row).T])

            format_func = lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) and x > 0 else "-"
            if hasattr(pivot_with_total, 'map'):
                formatted_matrix = pivot_with_total.map(format_func)
            else:
                formatted_matrix = pivot_with_total.applymap(format_func)

            st.dataframe(formatted_matrix, use_container_width=True, height=520)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                pivot_with_total.to_excel(writer, sheet_name='Yarn_Payable_Summary')
            output.seek(0)

            st.download_button(
                label="📥 Download Professional Excel Report (.xlsx)",
                data=output,
                file_name="ANM_Fashion_Yarn_Payable_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab_details:
            st.subheader("Filtered Invoice Line-Items")
            st.dataframe(df_clean, use_container_width=True)

    except Exception as e:
        st.error(f"Error Processing File: {e}")
else:
    st.info("👈 ANM FASHION Dashboard run karne ke liye sidebar se apni Excel sheet (.xlsx) upload karein.")
