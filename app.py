import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime

# Page Config
st.set_page_config(
    page_title="Accounts Payable & Yarn Receiving Dashboard",
    page_icon="🧵",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
    <style>
    .main .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    .stMetric {
        background-color: #1e222d;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #29b6f6;
    }
    .metric-card-2 { border-left-color: #66bb6a !important; }
    .metric-card-3 { border-left-color: #ef5350 !important; }
    .metric-card-4 { border-left-color: #ab47bc !important; }
    </style>
""", unsafe_allow_javascript=True)

st.title("🧵 Accounts Payable & Yarn Receiving Dashboard")
st.caption("Automated Textile Ledger, Party-wise Matrix Pivot & Payable Analytics")

# Sidebar Controls
st.sidebar.header("⚙️ File & Column Settings")
header_row = st.sidebar.number_input("Header Row Index (0 = 1st line, 2 = 3rd line)", min_value=0, max_value=10, value=2, step=1)

uploaded_file = st.sidebar.file_uploader("Upload Excel File (.xlsx, .xls)", type=["xlsx", "xls"])

def parse_month_year(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip().upper()
    
    # Common date format parsings
    try:
        dt = pd.to_datetime(val_str, errors='coerce')
        if not pd.isna(dt):
            return dt.strftime('%b-%Y').upper()
    except:
        pass
        
    # Manual string mapping if JAN-2026 / JAN-26 / 2026-01
    months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    for idx, m in enumerate(months):
        if m in val_str:
            # find year digits
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
        # Load excel data
        df = pd.read_excel(uploaded_file, header=header_row)
        df.columns = [str(c).strip() for c in df.columns]
        
        st.sidebar.success("✅ File Loaded Successfully!")
        
        # Column Matching Diagnostics
        cols_lower = {c.lower(): c for c in df.columns}
        
        default_party = next((v for k, v in cols_lower.items() if 'party' in k or 'supplier' in k), df.columns[0])
        default_month = next((v for k, v in cols_lower.items() if 'month' in k or 'date' in k), df.columns[1] if len(df.columns) > 1 else df.columns[0])
        default_amount = next((v for k, v in cols_lower.items() if 'invoice amount' in k or 'total' in k or 'amount' in k), df.columns[-1])
        default_qty = next((v for k, v in cols_lower.items() if 'qty' in k or 'quantity' in k or 'bags' in k), None)
        default_paid = next((v for k, v in cols_lower.items() if 'paid' in k or 'payment' in k), None)
        default_balance = next((v for k, v in cols_lower.items() if 'balance' in k or 'outstanding' in k), None)

        st.sidebar.subheader("📌 Column Mapping")
        party_col = st.sidebar.selectbox("Party Name Column", df.columns, index=list(df.columns).index(default_party))
        month_col = st.sidebar.selectbox("Month Column", df.columns, index=list(df.columns).index(default_month))
        amount_col = st.sidebar.selectbox("Invoice Amount Column", df.columns, index=list(df.columns).index(default_amount))
        qty_col = st.sidebar.selectbox("Yarn Quantity Column (Optional)", [None] + list(df.columns), index=(list(df.columns).index(default_qty) + 1) if default_qty else 0)
        paid_col = st.sidebar.selectbox("Paid Amount Column (Optional)", [None] + list(df.columns), index=(list(df.columns).index(default_paid) + 1) if default_paid else 0)
        balance_col = st.sidebar.selectbox("Balance Column (Optional)", [None] + list(df.columns), index=(list(df.columns).index(default_balance) + 1) if default_balance else 0)

        # Data Cleaning
        df_clean = df.copy()
        df_clean[party_col] = df_clean[party_col].astype(str).str.strip()
        df_clean = df_clean[~df_clean[party_col].isin(['None', 'nan', '', 'TOTAL', 'GRAND TOTAL'])]
        
        # Clean numeric fields
        df_clean[amount_col] = pd.to_numeric(df_clean[amount_col], errors='coerce').fillna(0)
        if qty_col:
            df_clean[qty_col] = pd.to_numeric(df_clean[qty_col], errors='coerce').fillna(0)
        if paid_col:
            df_clean[paid_col] = pd.to_numeric(df_clean[paid_col], errors='coerce').fillna(0)
        if balance_col:
            df_clean[balance_col] = pd.to_numeric(df_clean[balance_col], errors='coerce').fillna(0)

        # Format Month Standard
        df_clean['STD_MONTH'] = df_clean[month_col].apply(parse_month_year)

        # Executive Metrics
        tot_inv = df_clean[amount_col].sum()
        tot_paid = df_clean[paid_col].sum() if paid_col else 0
        tot_bal = df_clean[balance_col].sum() if balance_col else (tot_inv - tot_paid)
        tot_qty = df_clean[qty_col].sum() if qty_col else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Payable Amount", f"PKR {tot_inv:,.0f}")
        m2.metric("Total Paid", f"PKR {tot_paid:,.0f}")
        m3.metric("Outstanding Balance", f"PKR {tot_bal:,.0f}")
        m4.metric("Total Yarn Quantity Received", f"{tot_qty:,.0f} Bags/Kg" if qty_col else f"{len(df_clean)} Invoices")

        st.markdown("---")

        # Tabs Navigation
        tab_pivot, tab_details, tab_charts = st.tabs(["📅 Month-wise Pivot Summary", "🧾 Invoice-wise Detailed Report", "📈 Party Analytics"])

        with tab_pivot:
            st.subheader("Party-wise & Chronological Month-wise Matrix Pivot")
            
            val_metric = st.radio("Select Matrix Metric:", ["Invoice Amount", "Yarn Quantity", "Paid Amount"], horizontal=True)
            
            selected_val_col = amount_col
            if val_metric == "Yarn Quantity" and qty_col:
                selected_val_col = qty_col
            elif val_metric == "Paid Amount" and paid_col:
                selected_val_col = paid_col

            # Pivot Table Generation
            pivot = pd.pivot_table(
                df_clean,
                index=party_col,
                columns='STD_MONTH',
                values=selected_val_col,
                aggfunc='sum',
                fill_value=0
            )

            # Sort Months Chronologically
            sorted_months = sort_months_chronologically(pivot.columns)
            pivot = pivot[sorted_months]

            # Add Total Row & Column
            pivot['TOTAL'] = pivot.sum(axis=1)
            total_row = pivot.sum(axis=0)
            total_row.name = 'GRAND TOTAL'
            pivot_with_total = pd.concat([pivot, pd.DataFrame(total_row).T])

            # Formatting Matrix Values
            formatted_matrix = pivot_with_total.map(lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) and x > 0 else "-")

            st.dataframe(formatted_matrix, use_container_width=True, height=500)

            # Download CSV Button
            csv_data = pivot_with_total.to_csv().encode('utf-8')
            st.download_button("📥 Download Summary Matrix (CSV)", data=csv_data, file_name=f"Party_Month_Summary_{val_metric.replace(' ','_')}.csv", mime="text/csv")

        with tab_details:
            st.subheader("Filtered Invoice Line-Items")
            st.dataframe(df_clean, use_container_width=True)

        with tab_charts:
            st.subheader("Visual Analytics")
            c1, c2 = st.columns(2)
            
            with c1:
                top_parties = df_clean.groupby(party_col)[amount_col].sum().nlargest(10).reset_index()
                fig_bar = px.bar(top_parties, x=party_col, y=amount_col, title="Top 10 Suppliers by Billing Amount (PKR)", color=amount_col, color_continuous_scale="Viridis")
                st.plotly_chart(fig_bar, use_container_width=True)

            with c2:
                m_trend = df_clean.groupby('STD_MONTH')[amount_col].sum().reset_index()
                # Sort trend
                m_trend['SORT_KEY'] = m_trend['STD_MONTH'].apply(lambda x: datetime.strptime(str(x), "%b-%Y") if parse_month_year(x) else datetime(2099,1,1))
                m_trend = m_trend.sort_values('SORT_KEY')
                fig_line = px.line(m_trend, x='STD_MONTH', y=amount_col, title="Monthly Billing Trend", markers=True)
                st.plotly_chart(fig_line, use_container_width=True)

    except Exception as e:
        st.error(f"Error Processing File: {e}")
        st.info("💡 Tip: Try adjusting the 'Header Row Index' in the sidebar if columns are not recognized properly.")
else:
    st.info("👈 Dashboard run karne ke liye sidebar se apni Excel sheet (.xlsx) upload karein.")
