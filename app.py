import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime

# --- 1. Setup and Connect to Google Sheets ---
scopes = ["https://www.googleapis.com/auth/spreadsheets"]

# We load the JSON credentials securely from Streamlit's secrets manager
credentials = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=scopes
)
client = gspread.authorize(credentials)

# Open the Google Sheet
# Open the Google Sheet using the exact URL
SHEET_URL = "https://docs.google.com/spreadsheets/d/1kBFKNBkSNLJS4-qJrZ1QfhRGMkyGJCEft0T89JreTXw/edit?usp=sharing" 
sheet = client.open_by_url(SHEET_URL).sheet1

# --- 2. Build the User Interface ---
st.set_page_config(page_title="IEEE Claim Submission", page_icon="💸")
st.title("💸 IEEE Treasury Claim Form")
st.markdown("Submit your event expenses or income. This automatically updates the Master Ledger.")

with st.form("claim_form", clear_on_submit=True):
    date = st.date_input("Date of Transaction")
    payee = st.text_input("Your Name (Claimant / Payer)")
    description = st.text_input("Description (e.g., Paper plates for Welcoming Night)")
    
    category = st.selectbox("Category", ["Event Expense", "Merch Sales", "Operations", "Sponsorship", "Other"])
    txn_type = st.radio("Transaction Type", ["Expense (Claim)", "Income (Revenue)"])
    amount = st.number_input("Amount (RM)", min_value=0.0, format="%.2f")
    
    submit = st.form_submit_button("Submit Transaction")

    # --- 3. Process the Submission ---
    if submit:
        if not payee or not description or amount <= 0:
            st.error("Please fill in all details and ensure the amount is greater than 0.")
        else:
            with st.spinner("Updating Master Ledger..."):
                # Fetch existing records to calculate the new ID and Balance
                records = sheet.get_all_records()
                
                if len(records) > 0:
                    last_record = records[-1]
                    # Clean the string to do math (remove 'RM ' and commas)
                    last_balance_str = str(last_record.get('Running Balance', '0')).replace('RM', '').replace(',', '').strip()
                    last_balance = float(last_balance_str) if last_balance_str else 0.0
                        
                    # Auto-generate next Transaction ID
                    last_id = last_record.get('Transaction ID', 'TRX-0000')
                    try:
                        id_num = int(last_id.split('-')[1])
                        new_id = f"TRX-{id_num + 1:04d}"
                    except:
                        new_id = f"TRX-{len(records) + 1:04d}"
                else:
                    last_balance = 0.0
                    new_id = "TRX-0001"

                # Calculate new balance
                is_income = "Income" in txn_type
                income_val = amount if is_income else ""
                expense_val = amount if not is_income else ""
                new_balance = (last_balance + amount) if is_income else (last_balance - amount)

                # Format the row exactly matching your CSV columns
                row_data = [
                    new_id,
                    str(date),
                    description,
                    category,
                    f"{income_val:.2f}" if income_val else "",
                    f"{expense_val:.2f}" if expense_val else "",
                    "Pending Link", # We will add Google Drive upload in the next phase
                    payee,
                    "Cleared" if is_income else "Pending PV Submission",
                    "", # Notes
                    f"RM {new_balance:,.2f}"
                ]
                
                # Append to the Google Sheet
                sheet.append_row(row_data)
                st.success(f"✅ Successfully submitted! Tracked as {new_id}")