import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime

# --- 1. Setup and Connect to Google Sheets ---
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
credentials = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=scopes
)
client = gspread.authorize(credentials)

# Open the Google Sheet using the exact URL (Keep your URL here)
SHEET_URL = "https://docs.google.com/spreadsheets/d/1kBFKNBkSNLJS4-qJrZ1QfhRGMkyGJCEft0T89JreTXw/edit?usp=sharing" 
sheet = client.open_by_url(SHEET_URL).sheet1

# --- 2. Build the User Interface (Tabs) ---
st.set_page_config(page_title="IEEE Treasury App", page_icon="💸", layout="wide")
st.title("💸 IEEE Treasury Management")

# Create two tabs
tab1, tab2 = st.tabs(["📝 Submit Claim", "📊 Admin: Reconciliation & Export"])

# ==========================================
# TAB 1: CLAIM SUBMISSION (For Committee)
# ==========================================
with tab1:
    st.markdown("Submit event expenses or income. This automatically updates the Master Ledger.")
    with st.form("claim_form", clear_on_submit=True):
        date = st.date_input("Date of Transaction")
        payee = st.text_input("Your Name (Claimant / Payer)")
        description = st.text_input("Description (e.g., Paper plates for Welcoming Night)")
        
        category = st.selectbox("Category", ["Event Expense", "Merch Sales", "Operations", "Sponsorship", "Other"])
        txn_type = st.radio("Transaction Type", ["Expense (Claim)", "Income (Revenue)"])
        amount = st.number_input("Amount (RM)", min_value=0.0, format="%.2f")
        
        submit = st.form_submit_button("Submit Transaction")

        if submit:
            if not payee or not description or amount <= 0:
                st.error("Please fill in all details and ensure the amount is greater than 0.")
            else:
                with st.spinner("Updating Master Ledger..."):
                    records = sheet.get_all_records()
                    if len(records) > 0:
                        last_record = records[-1]
                        last_balance_str = str(last_record.get('Running Balance', '0')).replace('RM', '').replace(',', '').strip()
                        last_balance = float(last_balance_str) if last_balance_str else 0.0
                        
                        last_id = last_record.get('Transaction ID', 'TRX-0000')
                        try:
                            id_num = int(last_id.split('-')[1])
                            new_id = f"TRX-{id_num + 1:04d}"
                        except:
                            new_id = f"TRX-{len(records) + 1:04d}"
                    else:
                        last_balance = 0.0
                        new_id = "TRX-0001"

                    is_income = "Income" in txn_type
                    income_val = amount if is_income else ""
                    expense_val = amount if not is_income else ""
                    new_balance = (last_balance + amount) if is_income else (last_balance - amount)

                    row_data = [
                        new_id, str(date), description, category,
                        f"{income_val:.2f}" if income_val else "",
                        f"{expense_val:.2f}" if expense_val else "",
                        "Pending Link", payee,
                        "Cleared" if is_income else "Pending PV Submission",
                        "", f"RM {new_balance:,.2f}"
                    ]
                    
                    sheet.append_row(row_data)
                    st.success(f"✅ Successfully submitted! Tracked as {new_id}")


# ==========================================
# TAB 2: MONTHLY RECONCILIATION & EXPORT
# ==========================================
with tab2:
    st.header("Monthly Ledger Generator")
    
    # 1. Fetch current data into a Pandas DataFrame
    records = sheet.get_all_records()
    if records:
        df = pd.DataFrame(records)
        df['Date'] = pd.to_datetime(df['Date'])
        
        # 2. Select Month for processing
        months_available = df['Date'].dt.strftime('%B %Y').unique()
        selected_month = st.selectbox("Select Month to Process", months_available)
        
        # Filter data for selected month
        month_df = df[df['Date'].dt.strftime('%B %Y') == selected_month]
        
        st.subheader(f"Transactions for {selected_month}")
        st.dataframe(month_df[['Date', 'Description', 'Income', 'Expense', 'Internal Status', 'Running Balance']])
        
        st.divider()
        
        # 3. SA Reconciliation Module
        st.subheader("⚖️ SA Reconciliation")
        col1, col2 = st.columns(2)
        
        with col1:
            sa_reported_balance = st.number_input("Enter SA's Reported Closing Balance (RM)", format="%.2f", step=10.0)
            
            # Clean income/expense strings to calculate internal balance
            month_df['Clean_Income'] = pd.to_numeric(month_df['Income'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            month_df['Clean_Expense'] = pd.to_numeric(month_df['Expense'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            total_income = month_df['Clean_Income'].sum()
            total_expense = month_df['Clean_Expense'].sum()
            
            # Get internal closing balance (last row of the filtered month)
            internal_closing_str = str(month_df.iloc[-1]['Running Balance']).replace('RM', '').replace(',', '').strip()
            internal_closing = float(internal_closing_str)
            
            variance = internal_closing - sa_reported_balance
            
        with col2:
            st.metric("Internal System Closing Balance", f"RM {internal_closing:,.2f}")
            if sa_reported_balance > 0:
                if variance == 0:
                    st.success("✅ Balances match perfectly! You are cleared to submit the ledger.")
                else:
                    st.error(f"⚠️ Variance Detected: RM {variance:,.2f}")
                    st.warning("This is likely due to the following pending Payment Vouchers:")
                    pending_df = month_df[month_df['Internal Status'].str.contains("Pending", na=False, case=False)]
                    st.dataframe(pending_df[['Description', 'Expense', 'Payee/Payer']])
        
        # 4. Generate SA-Formatted CSV
        st.divider()
        st.subheader("📥 Export SA Monthly Ledger")
        st.markdown("This generates the CSV formatted exactly how the Students' Association requires it.")
        
        # Map master columns to SA columns
        export_df = pd.DataFrame()
        export_df['Date'] = month_df['Date'].dt.strftime('%Y-%m-%d')
        export_df['Details'] = month_df['Description']
        export_df['Income (RM)'] = month_df['Income']
        export_df['Expenses (RM)'] = month_df['Expense']
        export_df['OR No'] = "" # To be filled manually if needed
        export_df['CS-PV No'] = month_df['Notes'] # Assuming PVs are stored in Notes
        export_df['Balance (RM)'] = month_df['Running Balance']
        
        csv_export = export_df.to_csv(index=False).encode('utf-8')
        
        st.download_button(
            label=f"Download {selected_month} Ledger (CSV)",
            data=csv_export,
            file_name=f"Monthly_Ledger_{selected_month.replace(' ', '_')}.csv",
            mime='text/csv',
            type="primary"
        )
    else:
        st.info("No records found in the Master Ledger yet.")