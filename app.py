import streamlit as st
import pandas as pd

st.set_page_config(page_title="Accounts Payable Automation Tool", layout="wide")

st.title("🧵 Accounts Payable & Yarn Receiving Dashboard")
st.write("Apni Excel file yahan upload karein taake automated reports generate ho sakein.")

uploaded_file = st.file_uploader("Excel Sheet Upload Karein (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success("File kamyabi se load ho gayi hai!")
    
    tab1, tab2, tab3 = st.tabs(["📅 Month-wise Summary", "🧾 Invoice-wise Report", "📊 Raw Data View"])
    
    with tab1:
        st.subheader("Party-wise & Month-wise Payable Summary")
        try:
            # Map column names flexibly
            cols = {str(c).strip().upper(): c for c in df.columns}
            party_col = cols.get('PARTY NAME', cols.get('PARTY', None))
            month_col = cols.get('MONTH', None)
            
            if party_col and month_col:
                summary_df = df.groupby([party_col, month_col]).sum(numeric_only=True).reset_index()
                st.dataframe(summary_df, use_container_width=True)
            else:
                st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.error(f"Summary Error: {e}")
            
    with tab2:
        st.subheader("Invoice-wise Detailed Report")
        st.dataframe(df, use_container_width=True)
        
    with tab3:
        st.subheader("Complete Raw Data")
        st.dataframe(df, use_container_width=True)
else:
    st.info("Kripya pehle apni Excel file upload karein.")
