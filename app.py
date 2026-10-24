import streamlit as st
import pandas as pd
from datetime import datetime
import io
import streamlit.components.v1 as components

st.set_page_config(
    page_title="ANM FASHION - Accounts Payable & Yarn Report",
    page_icon="🧵",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Dark Theme & Print-Optimized CSS
st.markdown("""
    <style>
    .main { background-color: #0e1117; color: #ffffff; }
    .stApp { background-color: #0e1117; }
    
    [data-testid="stSidebar"] {
        background-color: #111827 !important;
        border-right: 1px solid #1f2937;
    }

    h1, h2, h3, h4, h5, h6, span, p, label {
        color: #ffffff !important;
    }

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
        color: #ffffff !important;
        margin: 0;
    }
    .company-subtitle {
        font-size: 18px !important;
        color: #f59e0b !important;
        font-weight: 600;
        margin-top: 5px;
    }

    .kpi-card {
        background: #1e293b;
        padding: 18px;
        border-radius: 10px;
        border: 1px solid #334155;
        border-left: 5px solid #3b82f6;
        box-shadow: 0 2px 10px rgba(0,0,0,0.2);
    }
    .kpi-card-gold { border-left-color: #f59e0b !important; }
    .kpi-card-green { border-left-color: #10b981 !important; }
    .kpi-card-red { border-left-color: #ef4444 !important; }
    
    .kpi-title {
        color: #94a3b8 !important;
        font-size: 13px;
        font-weight: 600;
        text-transform: uppercase;
    }
    .kpi-value {
        color: #ffffff !important;
        font-size: 24px;
        font-weight: 700;
        margin-top: 5px;
    }

    /* Print Settings: Sirf Table aur content print ho, baqi sab hide ho jaye */
    @media print {
        [data-testid="stSidebar"] { display: none !important; }
        .stButton { display: none !important; }
        .company-header { display: none !important; }
        header { display: none !important; }
        footer { display: none !important; }
        body, html, [data-testid="stAppViewContainer"] {
            background-color: #ffffff !important;
            color: #000000 !important;
        }
        .stTabs { color: #000000 !important; }
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
        balance_col = st.sidebar.selectbox("Balance Amount Column", balance_options, index=balance_index)

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
            df_clean[qty_col] = pd.to_numeric(df_clean[qty_col], errors='coerce').fillna(0)

        df_clean['STD_MONTH'] = df_clean[month_col].apply(parse_month_year)

        # Summary View Filter Selection
        st.sidebar.markdown("---")
        st.sidebar.subheader("🎯 Summary View Filter")
        summary_view_option = st.sidebar.radio(
            "Select View & Summary Metric:",
            ["3. Balance Outstanding (Default)", "1. Total Outstanding (Invoice)", "2. Paid Cheque (Paid)"]
        )

        if "1. Total Outstanding" in summary_view_option:
            pivot_value_col = invoice_col
            view_label = "Total Invoice Amount"
        elif "2. Paid Cheque" in summary_view_option and paid_col:
            pivot_value_col = paid_col
            view_label = "Paid Cheque Amount"
        else:
            pivot_value_col = balance_col
            view_label = "Balance Outstanding Amount"

        # Party Search Filter
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔍 Search & Filter Parties")
        unique_parties = sorted([str(p) for p in df_clean[party_col].unique() if str(p).strip()])
        selected_parties = st.sidebar.multiselect("Select Parties to Display", unique_parties, default=unique_parties)

        if selected_parties:
            df_clean = df_clean[df_clean[party_col].isin(selected_parties)]

        tot_inv = df_clean[invoice_col].sum()
        tot_paid_amt = df_clean[paid_col].sum() if paid_col else 0
        tot_balance = df_clean[balance_col].sum()

        # Top Executive KPI Cards
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
                <div class="kpi-card kpi-card-gold">
                    <div class="kpi-title">TOTAL INVOICE AMOUNT</div>
                    <div class="kpi-value">PKR {tot_inv:,.0f}</div>
                </div>
            """, unsafe_allow_html=True)

        with c2:
            st.markdown(f"""
                <div class="kpi-card kpi-card-green">
                    <div class="kpi-title">TOTAL PAID CHEQUES</div>
                    <div class="kpi-value">PKR {tot_paid_amt:,.0f}</div>
                </div>
            """, unsafe_allow_html=True)

        with c3:
            st.markdown(f"""
                <div class="kpi-card kpi-card-red">
                    <div class="kpi-title">TOTAL BALANCE PAYABLE</div>
                    <div class="kpi-value">PKR {tot_balance:,.0f}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with c4:
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">ACTIVE PARTIES</div>
                    <div class="kpi-value">{len(selected_parties)} Parties</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Tabs for Summary and Details
        tab_pivot, tab_details = st.tabs(["📊 Month-wise Summary Pivot", "🧾 Detailed Invoice Ledger"])

        with tab_pivot:
            st.subheader(f"Party-wise & Month-wise Summary ({view_label})")
            
            pivot = pd.pivot_table(
                df_clean,
                index=party_col,
                columns='STD_MONTH',
                values=pivot_value_col,
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

            # Excel Download & Print Buttons side-by-side at the bottom
            col_dl, col_pr = st.columns(2)
            
            with col_dl:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    pivot_with_total.to_excel(writer, sheet_name='Summary_Report')
                output.seek(0)

                st.download_button(
                    label="📥 Download Professional Excel Report (.xlsx)",
                    data=output,
                    file_name=f"ANM_Fashion_{view_label.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            with col_pr:
                if st.button("🖨️ Print / Save Summary Report as PDF", use_container_width=True, key="print_summary_btn"):
                    components.html("<script>window.parent.print();</script>", height=0, width=0)

        with tab_details:
            st.subheader("Filtered Invoice Line-Items & Detailed Ledger")
            
            st.dataframe(df_clean, use_container_width=True, height=520)

            # Excel Download & Print Buttons side-by-side at the bottom for Details
            col_dl_d, col_pr_d = st.columns(2)
            
            with col_dl_d:
                output_details = io.BytesIO()
                with pd.ExcelWriter(output_details, engine='openpyxl') as writer:
                    df_clean.to_excel(writer, index=False, sheet_name='Detailed_Ledger')
                output_details.seek(0)

                st.download_button(
                    label="📥 Download Detailed Ledger Excel (.xlsx)",
                    data=output_details,
                    file_name="ANM_Fashion_Detailed_Ledger.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            with col_pr_d:
                if st.button("🖨️ Print / Save Detailed Ledger as PDF", use_container_width=True, key="print_detail_btn"):
                    components.html("<script>window.parent.print();</script>", height=0, width=0)

    except Exception as e:
        st.error(f"Error Processing File: {e}")
else:
    st.info("👈 ANM FASHION Dashboard run karne ke liye sidebar se apni Excel sheet (.xlsx) upload karein.")
