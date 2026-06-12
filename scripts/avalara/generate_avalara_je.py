import pandas as pd
import numpy as np
import datetime
import calendar
import argparse
import os

STATE_ABBREV = {
    'AL': 'Alabama', 'AK': 'Alaska', 'AZ': 'Arizona', 'AR': 'Arkansas',
    'CA': 'California', 'CO': 'Colorado', 'CT': 'Connecticut', 'DE': 'Delaware',
    'FL': 'Florida', 'GA': 'Georgia', 'HI': 'Hawaii', 'ID': 'Idaho',
    'IL': 'Illinois', 'IN': 'Indiana', 'IA': 'Iowa', 'KS': 'Kansas',
    'KY': 'Kentucky', 'LA': 'Louisiana', 'ME': 'Maine', 'MD': 'Maryland',
    'MA': 'Massachusetts', 'MI': 'Michigan', 'MN': 'Minnesota', 'MS': 'Mississippi',
    'MO': 'Missouri', 'MT': 'Montana', 'NE': 'Nebraska', 'NV': 'Nevada',
    'NH': 'New Hampshire', 'NJ': 'New Jersey', 'NM': 'New Mexico', 'NY': 'New York',
    'NC': 'North Carolina', 'ND': 'North Dakota', 'OH': 'Ohio', 'OK': 'Oklahoma',
    'OR': 'Oregon', 'PA': 'Pennsylvania', 'RI': 'Rhode Island', 'SC': 'South Carolina',
    'SD': 'South Dakota', 'TN': 'Tennessee', 'TX': 'Texas', 'UT': 'Utah',
    'VT': 'Vermont', 'VA': 'Virginia', 'WA': 'Washington', 'WV': 'West Virginia',
    'WI': 'Wisconsin', 'WY': 'Wyoming', 'DC': 'Washington D.C.'
}

def load_vendor_map(csv_path):
    df_vendors = pd.read_csv(csv_path)
    # create mapping from State -> Vendor Name
    return dict(zip(df_vendors['State'], df_vendors['Vendor Name']))

def get_next_period_info(period_str):
    """
    Given a period like '202511', return:
    - the next month abbreviation and year (e.g. 'Dec 2025')
    - the last day of that next month as 'MM/DD/YYYY'
    """
    period_str = str(period_str).strip()
    if len(period_str) == 6 and period_str.isdigit():
        year = int(period_str[:4])
        month = int(period_str[4:])
    else:
        # Fallback to current date if we can't parse
        print(f"Warning: Could not parse Period '{period_str}'. Defaulting to current date.")
        now = datetime.datetime.now()
        year, month = now.year, now.month
    
    # Calculate next month
    if month == 12:
        next_month = 1
        next_year = year + 1
    else:
        next_month = month + 1
        next_year = year
        
    next_month_abbr = calendar.month_abbr[next_month]
    memo_str = f"{next_month_abbr} {next_year}"
    
    last_day = calendar.monthrange(next_year, next_month)[1]
    date_str = f"{next_month:02d}/{last_day:02d}/{next_year}"
    
    return memo_str, date_str

def safe_sum(series):
    return pd.to_numeric(series, errors='coerce').fillna(0).sum()

def format_amount(val):
    if val == 0:
        return ''
    return round(val, 2)

def generate_je(avalara_path, vendor_csv_path, output_path):
    # Load vendor map
    vendor_map = load_vendor_map(vendor_csv_path)
    
    # Read Avalara Report
    # Find header row dynamically
    df_raw = pd.read_excel(avalara_path, header=None)
    header_idx = -1
    for i, row in df_raw.iterrows():
        row_vals = [str(x) for x in row.values]
        if 'State' in row_vals and 'Amount Due To Avalara' in row_vals:
            header_idx = i
            break
            
    if header_idx == -1:
        raise ValueError("Could not find the header row containing 'State' and 'Amount Due To Avalara'.")
        
    df = df_raw.iloc[header_idx+1:].copy()
    df.columns = df_raw.iloc[header_idx].values
    df = df.dropna(subset=['State'])
    
    # Extract Period to get Memo and Date
    period_values = df['Period'].dropna().unique()
    if len(period_values) > 0:
        period_str = str(period_values[0])
    else:
        # default
        period_str = "000000" 
        
    next_period_memo, je_date = get_next_period_info(period_str)
    
    je_memo = f"{next_period_memo} - Avalara State Sales Tax Remittance"
    subsidiary = "Acme Holdings : Acme, Inc."
    
    cols_to_check = ['Amount Due To Avalara', 'Prior Period Vendor Discount', 'Prior Period Vendor Rounding']
    # Ensure they exist in df
    for c in cols_to_check:
        if c not in df.columns:
            df[c] = 0.0
            
    # Group by State
    df['State_Acronym'] = df['State'].astype(str).str.extract(r'([A-Z]{2})') # extracts NY from NY(US)
    
    csv_rows = []
    
    total_due = safe_sum(df['Amount Due To Avalara'])
    
    # 1. Cash Line (Credit)
    if total_due != 0:
        csv_rows.append({
            'Account': '111070 Cash and Cash Equivalents : Chase Checking x0001',
            'Debit': '',
            'Credit': format_amount(total_due),
            'Line Memo': je_memo,
            'Name': '',
            'Subsidiary': subsidiary,
            'Department': '',
            'Journal Entry Memo': je_memo,
            'Date': je_date
        })
        
    # Process each state
    grouped = df.groupby('State_Acronym')
    
    for state_acronym, group in grouped:
        st_due = safe_sum(group['Amount Due To Avalara'])
        st_prior_discount = safe_sum(group['Prior Period Vendor Discount'])
        st_prior_rounding = safe_sum(group['Prior Period Vendor Rounding'])
        
        if st_due == 0 and st_prior_discount == 0 and st_prior_rounding == 0:
            continue
            
        full_state_name = STATE_ABBREV.get(state_acronym, state_acronym)
        vendor_name = vendor_map.get(full_state_name, '')
        
        # Sales Tax Expense Lines
        def add_expense_line(amount, description):
            if amount == 0: return
            
            # If negative in Avalara, it's a Credit in JE (reduces expense/adjustment)
            # If positive, it's a Debit in JE
            debit = amount if amount > 0 else 0
            credit = -amount if amount < 0 else 0
            
            # The Avalara report amounts are negative for discounts, so they are credits
            csv_rows.append({
                'Account': '651210 Other Business Expenses : Sales Tax Expense',
                'Debit': format_amount(debit),
                'Credit': format_amount(credit),
                'Line Memo': f"{next_period_memo} - {state_acronym} {description}",
                'Name': vendor_name,
                'Subsidiary': subsidiary,
                'Department': 'General & Administrative : GA',
                'Journal Entry Memo': je_memo,
                'Date': je_date
            })
            
        add_expense_line(st_prior_discount, "prior period discount")
        add_expense_line(st_prior_rounding, "prior period rounding")
        
        # Sales Tax Payable Line
        # Payable Debit = Amount Due - Adjustments
        total_adjustments = st_prior_discount + st_prior_rounding
        payable_debit = st_due - total_adjustments
        
        if payable_debit != 0:
            # Note: Liability debit reduces payable. Liability credit increases payable.
            # Sales tax payable is typically debited when amount is paid.
            # If payable_debit is negative, it should be a credit.
            deb = payable_debit if payable_debit > 0 else 0
            cred = -payable_debit if payable_debit < 0 else 0
            
            csv_rows.append({
                'Account': '235100 Sales Taxes Payable : Sales Taxes Payable - Avatax',
                'Debit': format_amount(deb),
                'Credit': format_amount(cred),
                'Line Memo': f"{next_period_memo} - {state_acronym} Sales Tax",
                'Name': vendor_name,
                'Subsidiary': subsidiary,
                'Department': '',
                'Journal Entry Memo': je_memo,
                'Date': je_date
            })
            
    # Create output dataframe
    df_out = pd.DataFrame(csv_rows)
    # Reorder columns to ensure exact match
    expected_cols = ['Account', 'Debit', 'Credit', 'Line Memo', 'Name', 'Subsidiary', 'Department', 'Journal Entry Memo', 'Date']
    df_out = df_out[expected_cols]
    
    df_out.to_csv(output_path, index=False)
    print(f"Successfully generated Journal Entry CSV: {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Avalara JE CSV")
    parser.add_argument("avalara_report", help="Path to the Avalara raw report (.xlsx)")
    parser.add_argument("vendor_map", help="Path to State Department of Taxation Name.csv")
    parser.add_argument("output_path", help="Path to save the generated JE CSV")
    args = parser.parse_args()
    
    generate_je(args.avalara_report, args.vendor_map, args.output_path)
