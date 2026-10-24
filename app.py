import streamlit as st
import pandas as pd

st.set_page_config(page_title="Accounts Payable Automation Tool", layout="wide")

st.title("🧵 Accounts Payable & Yarn Receiving Dashboard")
st.write("Apni Excel file yahan upload karein taake automated reports generate ho sakein.")

# Sidebar for settings
st.sidebar.header("⚙️ Settings")
header_row = st.sidebar.number_input("Header Row Number (0 for 1st row, 1 for 2nd row)", min_value=0, max_value=5, value=2, step=1)

uploaded_file = st.file_uploader("Excel Sheet Upload Karein (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read Excel with dynamic header row
        df = pd.read_excel(uploaded_file, header=header_row)
        
        # Clean column names
        df.columns = [str(col).strip() for col in df.columns]
        
        st.success("File kamyabi se load ho gayi hai!")
        
        tab1, tab2, tab3 = st.tabs(["📅 Month-wise Pivot Summary", "🧾 Invoice-wise Report", "📊 Raw Data View"])
        
        with tab1:
            st.subheader("Party-wise & Month-wise Matrix Summary (Pivot Format)")
            try:
                cols_map = {c.upper(): c for c in df.columns}
                party_col = cols_map.get('PARTY NAME', cols_map.get('PARTY', None))
                month_col = cols_map.get('MONTH', None)
                amount_col = cols_map.get('INVOICE AMOUNT', cols_map.get('AMOUNT', None))
                
                if party_col and month_col and amount_col:
                    # Convert amount to numeric safely
                    df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce').fillna(0)
                    
                    # Create Pivot Table (Parties as rows, Months as columns)
                    pivot_df = df.pivot_table(
                        index=party_col, 
                        columns=month_col, 
                        values=amount_col, 
                        aggfunc='sum', 
                        fill_value=0
                    )
                    
                    # Format numbers with commas and dash for zeros (using map or applymap depending on pandas version)
                    format_func = lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) and x > 0 else "-"
                    if hasattr(pivot_df, 'map'):
                        formatted_pivot = pivot_df.map(format_func)
                    else:
                        formatted_pivot = pivot_df.applymap(format_func)
                    
                    st.dataframe(formatted_pivot, use_container_width=True)
                    
                    # Download button for summary report
                    csv = pivot_df.to_csv().encode('utf-8')
                    st.download_button("📥 Download Pivot Summary CSV", data=csv, file_name="Monthly_Pivot_Summary.csv", mime="text/csv")
                else:
                    st.warning("Columns 'PARTY NAME', 'MONTH' ya 'INVOICE AMOUNT/AMOUNT' theek se nahi mil rahe. 'Raw Data View' tab mein columns check karein.")
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
