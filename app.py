import streamlit as st
import pandas as pd

st.set_page_config(page_title="Accounts Payable Automation Tool", layout="wide")

st.title("🧵 Accounts Payable & Yarn Receiving Dashboard")
st.write("Apni Excel file yahan upload karein taake automated reports generate ho sakein.")

# Sidebar for advanced settings
st.sidebar.header("⚙️ Settings")
header_row = st.sidebar.number_input("Header Row Number (0 for 1st row, 1 for 2nd row)", min_value=0, max_value=5, value=0, step=1)

uploaded_file = st.file_uploader("Excel Sheet Upload Karein (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read Excel with dynamic header row
        df = pd.read_excel(uploaded_file, header=header_row)
        
        # Clean column names (remove extra spaces)
        df.columns = [str(col).strip() for col in df.columns]
        
        st.success("File kamyabi se load ho gayi hai!")
        
        tab1, tab2, tab3 = st.tabs(["📅 Month-wise Summary", "🧾 Invoice-wise Report", "📊 Raw Data View"])
        
        with tab1:
            st.subheader("Party-wise & Month-wise Payable Summary")
            try:
                # Find matching columns case-insensitively
                cols_map = {c.upper(): c for c in df.columns}
                party_col = cols_map.get('PARTY NAME', cols_map.get('PARTY', None))
                month_col = cols_map.get('MONTH', None)
                
                if party_col and month_col:
                    summary_df = df.groupby([party_col, month_col]).sum(numeric_only=True).reset_index()
                    st.dataframe(summary_df, use_container_width=True)
                else:
                    st.warning("Columns 'PARTY NAME' ya 'MONTH' nahi mil rahe. Neeche Raw Data View mein check karein.")
                    st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"Summary Error: {e}")
                
        with tab2:
            st.subheader("Invoice-wise Detailed Report")
            st.dataframe(df, use_container_width=True)
            
        with tab3:
            st.subheader("Complete Raw Data")
            st.dataframe(df, use_container_width=True)
            
    except Exception as e:
        st.error(f"File parhne mein masla aaya: {e}")
else:
    st.info("Kripya pehle apni Excel file upload karein.")
