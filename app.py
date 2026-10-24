import streamlit as st
import pandas as pd
from datetime import datetime

# Page Config
st.set_page_config(
    page_title="Accounts Payable & Yarn Receiving Dashboard",
    page_icon="🧵",
    layout="wide"
)

st.title("🧵 Accounts Payable & Yarn Receiving Dashboard")
st.caption("Automated Textile Ledger & Chronological Month-wise Matrix Pivot Summary")

# Sidebar Controls
st.sidebar.header("⚙️ File & Column Settings")
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
        qty_col = st.sidebar.selectbox("Yarn Quantity Column (Optional)", [None] + list(df.columns), index=(list(df.columns).index(default_qty) + 1) if default_qty else 0)

        # Data Cleaning
        df_clean = df.copy()
        df_clean[party_col] = df_clean[party_col].astype(str).str.strip()
        df_clean = df_clean[~df_clean[party_col].isin(['None', 'nan', '', 'TOTAL', 'GRAND TOTAL'])]
        
        df_clean[amount_col] = pd.to_numeric(df_clean[amount_col], errors='coerce').fillna(0)
        if qty_col:
            df_clean[qty_col] = pd.to_numeric(df_clean[qty_col], errors='coerce').fillna(0)

        df_clean['STD_MONTH'] = df_clean[month_col].apply(parse_month_year)

        # KPI Summary Cards
        tot_inv = df_clean[amount_col].sum()
        tot_qty = df_clean[qty_col].sum() if qty_col else 0

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Payable Amount", f"PKR {tot_inv:,.0f}")
        m2.metric("Total Yarn Quantity Received", f"{tot_qty:,.0f} Bags/Kg" if qty_col else "N/A")
        m3.metric("Total Invoices / Transactions", f"{len(df_clean)} Records")

        st.markdown("---")

        tab_pivot, tab_details = st.tabs(["📅 Month-wise Pivot Summary", "🧾 Invoice-wise Detailed Report"])

        with tab_pivot:
            st.subheader("Party-wise & Chronological Month-wise Matrix Pivot")
            
            # Pivot Table Generation
            pivot = pd.pivot_table(
                df_clean,
                index=party_col,
                columns='STD_MONTH',
                values=amount_col,
                aggfunc='sum',
                fill_value=0
            )

            # Sort Months Chronologically (JAN to DEC)
            sorted_months = sort_months_chronologically(pivot.columns)
            pivot = pivot[sorted_months]

            # Add Total Column & Row
            pivot['TOTAL'] = pivot.sum(axis=1)
            total_row = pivot.sum(axis=0)
            total_row.name = 'GRAND TOTAL'
            pivot_with_total = pd.concat([pivot, pd.DataFrame(total_row).T])

            # Formatting Matrix Values (Commas and Dash for Zeros)
            format_func = lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) and x > 0 else "-"
            if hasattr(pivot_with_total, 'map'):
                formatted_matrix = pivot_with_total.map(format_func)
            else:
                formatted_matrix = pivot_with_total.applymap(format_func)

            st.dataframe(formatted_matrix, use_container_width=True, height=500)

            # Download CSV Button
            csv_data = pivot_with_total.to_csv().encode('utf-8')
            st.download_button("📥 Download Pivot Matrix CSV", data=csv_data, file_name="Monthly_Payable_Summary.csv", mime="text/csv")

        with tab_details:
            st.subheader("Filtered Invoice Line-Items")
            st.dataframe(df_clean, use_container_width=True)

    except Exception as e:
        st.error(f"Error Processing File: {e}")
else:
    st.info("👈 Dashboard run karne ke liye sidebar se apni Excel sheet (.xlsx) upload karein.")
