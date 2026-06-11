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
    
    records = sheet.get_all_records()
    if records:
        df = pd.DataFrame(records)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        valid_dates = df.dropna(subset=['Date'])
        months_available = valid_dates['Date'].dt.strftime('%B %Y').unique()
        
        if len(months_available) > 0:
            selected_month = st.selectbox("Select Month to Process", months_available)
            month_df = valid_dates[valid_dates['Date'].dt.strftime('%B %Y') == selected_month].copy()
            
            st.subheader(f"Transactions for {selected_month}")
            st.dataframe(month_df[['Date', 'Description', 'Income', 'Expense', 'Internal Status', 'Running Balance']])
            
            st.divider()
            st.subheader("⚖️ SA Reconciliation")
            col1, col2 = st.columns(2)
            
            with col1:
                sa_reported_balance = st.number_input("Enter SA's Reported Closing Balance (RM)", format="%.2f", step=10.0)
                month_df['Clean_Income'] = pd.to_numeric(month_df['Income'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                month_df['Clean_Expense'] = pd.to_numeric(month_df['Expense'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                
                internal_closing_str = str(month_df.iloc[-1]['Running Balance']).replace('RM', '').replace(',', '').strip()
                internal_closing = float(internal_closing_str)
                variance = internal_closing - sa_reported_balance
                
            with col2:
                st.metric("Internal System Closing Balance", f"RM {internal_closing:,.2f}")
                if sa_reported_balance > 0:
                    if variance == 0:
                        st.success("✅ Balances match perfectly!")
                    else:
                        st.error(f"⚠️ Variance Detected: RM {variance:,.2f}")
            
            st.divider()
            st.subheader("📤 Publish Formatted Sheet")
            
            # Form Inputs for Signatures and Logos
            logo_ieee_url = st.text_input("IEEE Logo URL (ending in .png/.jpg)", value="https://github.com/SeanTai123/ieee-treasurer/blob/main/ieee.png?raw=true")
            logo_sa_url = st.text_input("SA Logo URL (ending in .png/.jpg)", value="https://github.com/SeanTai123/ieee-treasurer/blob/main/SA.png?raw=true")
            
            col_sig1, col_sig2 = st.columns(2)
            with col_sig1:
                prep_name = st.text_input("Prepared by (Name)", value="Wong Yan Jie")
                prep_email = st.text_input("Prepared by (Email)", value="efyyw6@nottingham.edu.my")
                prep_sig_url = st.text_input("Preparer Signature URL (Optional)", value="")
            with col_sig2:
                ver_name = st.text_input("Verified by (Name)", value="Marcus Yew Min Jie")
                ver_email = st.text_input("Verified by (Email)", value="ecymy2@nottingham.edu.my")
                ver_sig_url = st.text_input("Verifier Signature URL (Optional)", value="")
                
            sig_date = st.date_input("Signature Date").strftime('%d/%m/%Y')

            if st.button(f"Generate '{selected_month}' Tab"):
                with st.spinner("Building SA layout, injecting images, and formatting..."):
                    try:
                        # 1. Calculate Opening Balance
                        first_txn_idx = month_df.index[0]
                        if first_txn_idx > 0:
                            prev_balance_str = str(df.iloc[first_txn_idx - 1]['Running Balance']).replace('RM', '').replace(',', '').strip()
                            opening_balance = float(prev_balance_str)
                        else:
                            first_inc = month_df.iloc[0]['Clean_Income']
                            first_exp = month_df.iloc[0]['Clean_Expense']
                            first_bal = float(str(month_df.iloc[0]['Running Balance']).replace('RM', '').replace(',', '').strip())
                            opening_balance = first_bal - first_inc + first_exp

                        # 2. Prepare Image Formulas
                        # The '1' in the formula tells Google Sheets to resize the image to fit inside the merged cell
                        ieee_formula = f'=IMAGE("{logo_ieee_url}", 1)' if logo_ieee_url else ""
                        sa_formula = f'=IMAGE("{logo_sa_url}", 1)' if logo_sa_url else ""
                        prep_sig_formula = f'=IMAGE("{prep_sig_url}", 1)' if prep_sig_url else ""
                        ver_sig_formula = f'=IMAGE("{ver_sig_url}", 1)' if ver_sig_url else ""

                        # 3. Construct the Layout (Matching the screenshot)
                        sa_layout = [
                            [ieee_formula, "", "", "", "", sa_formula, ""], # Row 1 (Logos)
                            ["", "", "", "", "", "", ""],                   # Row 2
                            ["", "", "", "", "", "", ""],                   # Row 3
                            ["IEEE UNM Student Branch", "", "", "", "", "", ""], # Row 4
                            [f"Monthly Ledger for {selected_month}", "", "", "", "", "", ""], # Row 5
                            ["", "", "", "", "", "", ""],                   # Row 6
                            ["Date", "Details", "Income", "Expenses", "OR No", "CS-PV No", "Balance"], # Row 7
                            ["", "", "RM", "RM", "", "", "RM"],             # Row 8
                            [month_df.iloc[0]['Date'].strftime('%Y-%m-%d'), "Opening Balance", "", "", "", "", f"{opening_balance:.2f}"] # Row 9
                        ]

                        # Add transactions
                        for _, row in month_df.iterrows():
                            inc = f"{row['Clean_Income']:.2f}" if row['Clean_Income'] > 0 else ""
                            exp = f"{row['Clean_Expense']:.2f}" if row['Clean_Expense'] > 0 else ""
                            bal = f"{float(str(row['Running Balance']).replace('RM', '').replace(',', '').strip()):.2f}"
                            sa_layout.append([
                                row['Date'].strftime('%Y-%m-%d'),
                                row['Description'],
                                inc, exp, "", row['Notes'], bal
                            ])

                        # Add bottom summary and signatures
                        sa_layout.extend([
                            ["", "", "", "", "", "", ""],
                            ["", "", "", "", "", "TOTAL: ", f"{internal_closing:.2f}"],
                            ["", "", "", "", "", "", ""],
                            # Signature Image Row (If URLs are provided)
                            [prep_sig_formula, "", "", ver_sig_formula, "", "", ""],
                            [f"Prepared by: {prep_name}", "", "", f"Verified by: {ver_name}", "", "", ""],
                            [f"Email Username: {prep_email}", "", "", f"Email Username: {ver_email}", "", "", ""],
                            [f"Date: {sig_date}", "", "", f"Date: {sig_date}", "", "", ""]
                        ])

                        # 4. Create or Overwrite the Worksheet
                        sheet_title = f"Ledger {selected_month}"
                        try:
                            target_sheet = client.open_by_url(SHEET_URL).worksheet(sheet_title)
                            target_sheet.clear()
                        except gspread.exceptions.WorksheetNotFound:
                            target_sheet = client.open_by_url(SHEET_URL).add_worksheet(title=sheet_title, rows=len(sa_layout)+10, cols=8)

                        # Push data
                        target_sheet.update(range_name='A1', values=sa_layout, value_input_option='USER_ENTERED')
                        
                        # 5. Apply Visual Formatting 
                        # Merge cells for logos (A1:B3 for IEEE, F1:G3 for SA)
                        target_sheet.merge_cells('A1:B3')
                        target_sheet.merge_cells('F1:G3')
                        
                        # Merge cells for signatures so they have space
                        sig_row = len(sa_layout) - 3 # Row index where signatures are
                        target_sheet.merge_cells(f'A{sig_row}:B{sig_row+1}') 
                        target_sheet.merge_cells(f'D{sig_row}:E{sig_row+1}')

                        # Bold the Headers and the TOTAL line
                        target_sheet.format('A4:A5', {'textFormat': {'bold': True}})
                        target_sheet.format('A7:G8', {'textFormat': {'bold': True}, 'horizontalAlignment': 'CENTER'})
                        total_row = len(sa_layout) - 5
                        target_sheet.format(f'F{total_row}:G{total_row}', {'textFormat': {'bold': True}})
                        
                        # Set column widths (Approximate, you can tweak these later in Sheets)
                        # gspread standard column resize requires a slightly more complex batch update, 
                        # but standard text length usually auto-stretches decently.

                        st.success(f"✅ Successfully formatted and published '{sheet_title}' with logos!")
                        st.markdown(f"[Click here to check your Google Sheet]({SHEET_URL})")

                    except Exception as e:
                        st.error(f"An error occurred: {e}")
        else:
            st.info("No valid month data found. Submit a transaction first.")
    else:
        st.info("No records found in the Master Ledger.")