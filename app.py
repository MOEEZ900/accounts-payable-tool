import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Page Configuration
st.set_page_config(
    page_title="ANM FASHION - Accounts Payable & Yarn Report",
    page_icon="🧵",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Dark Theme UI with Advanced Print-Optimized CSS
st.markdown(
    """
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

    /* PERFECT PRINT & PDF CSS FIX - Hiding Tabs, Buttons & Sidebar */
    @media print {
        @page {
            size: A4 landscape;
            margin: 10mm;
        }

        [data-testid="stSidebar"], .stButton, header, footer, .company-header, .kpi-card, [data-testid="stTabs"], .stTabs, div[data-baseweb="tab-list"] {
            display: none !important;
        }

        body, html, [data-testid="stAppViewContainer"], .main, .stApp {
            background-color: #ffffff !important;
            color: #000000 !important;
        }

        h1, h2, h3, h4, h5, h6, span, p, label, div {
            color: #000000 !important;
        }

        [data-testid="stDataFrame"] {
            background-color: #ffffff !important;
            height: auto !important;
            max-height: none !important;
            overflow: visible !important;
        }

        div[data-testid="stDataFrame"] * {
            background-color: #ffffff !important;
            color: #000000 !important;
            border-color: #cccccc !important;
        }

        table, th, td {
            border: 1px solid #000000 !important;
            color: #000000 !important;
            background-color: #ffffff !important;
            font-size: 11px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Header Section
st.markdown(
    """
    <div class="company-header">
        <div class="company-title">ANM FASHION</div>
        <div class="company-subtitle">🧵 YARN PAYABLE REPORT & ACCOUNTS DASHBOARD</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Sidebar Controls
st.sidebar.header("⚙️ Settings & Controls")
header_row = st.sidebar.number_input(
    "Header Row Index (0 = 1st line, 2 = 3rd line)",
    min_value=0,
    max_value=10,
    value=2,
    step=1,
)

uploaded_file = st.sidebar.file_uploader(
    "Upload Excel File (.xlsx, .xls)", type=["xlsx", "xls"]
)


def parse_month_year(val):
  if pd.isna(val):
    return None
  val_str = str(val).strip().upper()
  try:
    dt = pd.to_datetime(val_str, errors="coerce")
    if not pd.isna(dt):
      return dt.strftime("%b-%Y").upper()
  except Exception:
    pass

  months = [
      "JAN",
      "FEB",
      "MAR",
      "APR",
      "MAY",
      "JUN",
      "JUL",
      "AUG",
      "SEP",
      "OCT",
      "NOV",
      "DEC",
  ]
  for m in months:
    if m in val_str:
      years = re.findall(r"\d{4}|\d{2}", val_str)
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
    except Exception:
      try:
        return datetime.strptime(str(m_str), "%B-%Y")
      except Exception:
        return datetime(2099, 12, 31)

  return sorted(month_list, key=get_sort_key)


if uploaded_file is not None:
  try:
    df = pd.read_excel(uploaded_file, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    # Remove Unnamed Columns Safely
    valid_cols = [c for c in df.columns if "Unnamed" not in str(c)]
    df = df[valid_cols]

    st.sidebar.success("✅ File Loaded Successfully!")

    cols_lower = {c.lower(): c for c in df.columns}

    default_party = df.columns[0]
    for k, v in cols_lower.items():
      if "party" in k or "supplier" in k:
        default_party = v
        break

    default_month = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    for k, v in cols_lower.items():
      if "month" in k or "date" in k:
        default_month = v
        break

    default_invoice = (
        df.columns[-3] if len(df.columns) >= 3 else df.columns[-1]
    )
    for k, v in cols_lower.items():
      if k == "invoice" or "total amount" in k:
        default_invoice = v
        break

    default_paid = None
    for k, v in cols_lower.items():
      if k == "paid":
        default_paid = v
        break

    default_balance = None
    for k, v in cols_lower.items():
      if k == "balance":
        default_balance = v
        break

    default_qty = None
    for k, v in cols_lower.items():
      if k == "qty" or "quantity" in k or "bags" in k:
        default_qty = v
        break

    st.sidebar.subheader("📌 Column Mapping")
    party_col = st.sidebar.selectbox(
        "Party Name Column",
        df.columns,
        index=list(df.columns).index(default_party),
    )
    month_col = st.sidebar.selectbox(
        "Month Column",
        df.columns,
        index=list(df.columns).index(default_month),
    )

    invoice_col = st.sidebar.selectbox(
        "Invoice Amount Column",
        df.columns,
        index=list(df.columns).index(default_invoice),
    )

    paid_options = [None] + list(df.columns)
    paid_index = (
        (list(df.columns).index(default_paid) + 1)
        if default_paid and default_paid in df.columns
        else 0
    )
    paid_col = st.sidebar.selectbox(
        "Paid Amount Column", paid_options, index=paid_index
    )

    balance_options = [None] + list(df.columns)
    balance_index = (
        (list(df.columns).index(default_balance) + 1)
        if default_balance and default_balance in df.columns
        else 0
    )
    balance_col = st.sidebar.selectbox(
        "Balance Amount Column", balance_options, index=balance_index
    )

    qty_options = [None] + list(df.columns)
    qty_index = (
        (list(df.columns).index(default_qty) + 1)
        if default_qty and default_qty in df.columns
        else 0
    )
    qty_col = st.sidebar.selectbox(
        "Yarn Quantity Column (Optional)", qty_options, index=qty_index
    )

    # Initialize Session State for Data Persistence during edits
    if "df_edited" not in st.session_state:
      st.session_state.df_edited = df.copy()

    # Data Cleaning Base
    df_clean = st.session_state.df_edited.copy()
    df_clean[party_col] = df_clean[party_col].fillna("").astype(str).str.strip()
    df_clean = df_clean[
        ~df_clean[party_col].isin(
            ["None", "nan", "", "TOTAL", "GRAND TOTAL", "0"]
        )
    ]

    df_clean[invoice_col] = pd.to_numeric(
        df_clean[invoice_col], errors="coerce"
    ).fillna(0)

    if paid_col:
      df_clean[paid_col] = pd.to_numeric(
          df_clean[paid_col], errors="coerce"
      ).fillna(0)

    if balance_col:
      df_clean[balance_col] = pd.to_numeric(
          df_clean[balance_col], errors="coerce"
      ).fillna(0)
    else:
      df_clean["CALC_BALANCE"] = df_clean[invoice_col] - (
          df_clean[paid_col] if paid_col else 0
      )
      balance_col = "CALC_BALANCE"

    if qty_col:
      df_clean[qty_col] = pd.to_numeric(
          df_clean[qty_col], errors="coerce"
      ).fillna(0)

    df_clean["STD_MONTH"] = df_clean[month_col].apply(parse_month_year)

    # Summary View Filter Selection
    st.sidebar.markdown("---")
    st.sidebar.subheader("🎯 Summary View Filter")
    summary_view_option = st.sidebar.radio(
        "Select View & Summary Metric:",
        [
            "3. Balance Outstanding (Default)",
            "1. Total Outstanding (Invoice)",
            "2. Paid Cheque (Paid)",
        ],
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

    # Option to Hide Zero/Blank Month Columns
    st.sidebar.markdown("---")
    st.sidebar.subheader("👁️ Display Options")
    hide_zero_cols = st.sidebar.checkbox(
        "Hide Blank/Zero Month Columns", value=True
    )

    # NEW OPTION: Add or Less Columns in Detailed Ledger
    st.sidebar.markdown("---")
    st.sidebar.subheader("🛠️ Ledger Columns Control")
    all_available_cols = list(df_clean.columns)
    if "STD_MONTH" in all_available_cols:
      all_available_cols.remove("STD_MONTH")

    default_selected_cols = [
        c
        for c in [
            party_col,
            month_col,
            invoice_col,
            paid_col,
            balance_col,
            qty_col,
        ]
        if c and c in all_available_cols
    ]
    # Agar kuch default select na ho toh saare columns dikhayein
    if not default_selected_cols:
      default_selected_cols = all_available_cols

    selected_ledger_cols = st.sidebar.multiselect(
        "Select Columns to Display in Ledger",
        options=all_available_cols,
        default=default_selected_cols,
    )

    # Party Search Filter
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Search & Filter Parties")
    unique_parties = sorted(
        [str(p) for p in df_clean[party_col].unique() if str(p).strip()]
    )
    selected_parties = st.sidebar.multiselect(
        "Select Parties to Display", unique_parties, default=unique_parties
    )

    if selected_parties:
      df_clean = df_clean[df_clean[party_col].isin(selected_parties)]

    tot_inv = df_clean[invoice_col].sum()
    tot_paid_amt = df_clean[paid_col].sum() if paid_col else 0
    tot_balance = df_clean[balance_col].sum()

    # Top Executive KPI Cards
    c1, c2, c3, c4 = st.columns(4)
    with c1:
      st.markdown(
          '<div class="kpi-card kpi-card-gold"><div'
          ' class="kpi-title">TOTAL INVOICE AMOUNT</div><div'
          ' class="kpi-value">PKR {:,.0f}</div></div>'.format(tot_inv),
          unsafe_allow_html=True,
      )

    with c2:
      st.markdown(
          '<div class="kpi-card kpi-card-green"><div'
          ' class="kpi-title">TOTAL PAID CHEQUES</div><div'
          ' class="kpi-value">PKR {:,.0f}</div></div>'.format(tot_paid_amt),
          unsafe_allow_html=True,
      )

    with c3:
      st.markdown(
          '<div class="kpi-card kpi-card-red"><div'
          ' class="kpi-title">TOTAL BALANCE PAYABLE</div><div'
          ' class="kpi-value">PKR {:,.0f}</div></div>'.format(tot_balance),
          unsafe_allow_html=True,
      )

    with c4:
      st.markdown(
          '<div class="kpi-card"><div class="kpi-title">ACTIVE PARTIES</div><div'
          ' class="kpi-value">{} Parties</div></div>'.format(
              len(selected_parties)
          ),
          unsafe_allow_html=True,
      )

    st.markdown("<br>", unsafe_allow_html=True)

    # Tabs for Summary and Details
    tab_pivot, tab_details = st.tabs(
        ["📊 Month-wise Summary Pivot", "🧾 Detailed Invoice Ledger & Editing"]
    )

    with tab_pivot:
      st.subheader(f"Party-wise & Month-wise Summary ({view_label})")

      pivot = pd.pivot_table(
          df_clean,
          index=party_col,
          columns="STD_MONTH",
          values=pivot_value_col,
          aggfunc="sum",
          fill_value=0,
      )

      sorted_months = sort_months_chronologically(pivot.columns)
      pivot = pivot[sorted_months]

      if hide_zero_cols:
        cols_to_keep = [
            c for c in pivot.columns if c == "TOTAL" or pivot[c].sum() > 0
        ]
        pivot = pivot[cols_to_keep]

      pivot["TOTAL"] = pivot.sum(axis=1)
      total_row = pivot.sum(axis=0)
      total_row.name = 'GRAND TOTAL'
      pivot_with_total = pd.concat([pivot, pd.DataFrame(total_row).T])

      format_func = (
          lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) and x > 0 else "-"
      )
      if hasattr(pivot_with_total, "map"):
        formatted_matrix = pivot_with_total.map(format_func)
      else:
        formatted_matrix = pivot_with_total.applymap(format_func)

      st.markdown("### ANM FASHION - " + view_label)
      st.write(
          formatted_matrix.to_html(classes="styled-table"),
          unsafe_allow_html=True,
      )

      col_dl, col_pr = st.columns(2)

      with col_dl:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          pivot_with_total.to_excel(writer, sheet_name="Summary_Report")
        output.seek(0)

        st.download_button(
            label="📥 Download Professional Excel Report (.xlsx)",
            data=output,
            file_name="ANM_Fashion_Summary_Report.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

      with col_pr:
        if st.button(
            "🖨️ Print / Save Summary Report as PDF",
            use_container_width=True,
            key="print_summary_btn",
        ):
          components.html(
              "<script>window.parent.print();</script>", height=0, width=0
          )

    with tab_details:
      st.subheader(
          "🧾 Detailed Invoice Ledger & Direct Editing (Cheque Paid Update)"
      )
      st.info(
          "💡 Aap sidebar mein 'Ledger Columns Control' se koi bhi column add ya"
          " remove kar sakte hain, aur table mein direct amount edit kar"
          " sakte hain!"
      )

      col_filter_p, col_empty = st.columns([2, 2])
      with col_filter_p:
        print_party_target = st.selectbox(
            "Select Specific Party for Invoice-wise Detail / Statement Print:",
            ["All Filtered Parties"] + unique_parties,
            key="edit_party_select",
        )

      if print_party_target != "All Filtered Parties":
        subset_df = df_clean[df_clean[party_col] == print_party_target]
      else:
        subset_df = df_clean

      # Apply selected columns filter for display & editing
      if selected_ledger_cols:
        display_subset = subset_df[
            [c for c in selected_ledger_cols if c in subset_df.columns]
        ]
      else:
        display_subset = subset_df

      # Interactive Editable Dataframe for Live Updates
      edited_subset = st.data_editor(
          display_subset,
          use_container_width=True,
          height=450,
          key="data_editor_view",
      )

      # Update session state with edited values back into the main dataframe
      if not edited_subset.empty:
        common_indices = edited_subset.index.intersection(
            st.session_state.df_edited.index
        )
        for col in edited_subset.columns:
          if col in st.session_state.df_edited.columns:
            st.session_state.df_edited.loc[common_indices, col] = (
                edited_subset.loc[common_indices, col]
            )

      # Automatically recalculate balance after editing if paid and invoice columns exist
      if paid_col and invoice_col and balance_col:
        st.session_state.df_edited[balance_col] = pd.to_numeric(
            st.session_state.df_edited[invoice_col], errors="coerce"
        ).fillna(0) - pd.to_numeric(
            st.session_state.df_edited[paid_col], errors="coerce"
        ).fillna(
            0
        )

      col_dl_d, col_pr_d = st.columns(2)

      file_suffix = (
          print_party_target.replace(" ", "_")
          if print_party_target != "All Filtered Parties"
          else "All"
      )

      with col_dl_d:
        output_details = io.BytesIO()
        with pd.ExcelWriter(output_details, engine="openpyxl") as writer:
          subset_df.to_excel(writer, index=False, sheet_name="Detailed_Ledger")
        output_details.seek(0)

        st.download_button(
            label="📥 Download Detailed Ledger Excel (.xlsx)",
            data=output_details,
            file_name=f"ANM_Fashion_Detailed_Ledger_{file_suffix}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

      with col_pr_d:
        if st.button(
            "🖨️ Print / Save Invoice-wise Detail as PDF",
            use_container_width=True,
            key="print_detail_btn",
        ):
          components.html(
              "<script>window.parent.print();</script>", height=0, width=0
          )

  except Exception as e:
    st.error(f"Error Processing File: {e}")
else:
  st.info(
      "👈 ANM FASHION Dashboard run karne ke liye sidebar se apni Excel sheet"
      " upload karein."
  )
