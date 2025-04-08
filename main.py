import argparse
import datetime
from math import trunc
import os
import pandas as pd
import calendar
import re

def detect_header_row(raw_df):
    expected_keywords = {"items", "invoice", "amount"}
    for i, row in raw_df.iterrows():
        row_values = [str(cell).lower() for cell in row if pd.notna(cell)]
        for cell in row_values:
            for keyword in expected_keywords:
                if keyword in cell:
                    return i
    return None

def parse_month_cols(df):
    def is_month_column(col, idx):
        try:
            # Try parsing known month-year patterns by normalizing strings first
            if isinstance(col, str):
                clean_col = col.replace(" ", "-").replace("/", "-")
                # Handle compact formats like 'May2024' or 'May24'
                compact_match = re.fullmatch(r'([A-Za-z]{3,9})(\d{2,4})', clean_col)
                if compact_match:
                    month_part = compact_match.group(1)
                    year_part = compact_match.group(2)
                    clean_col = f"{month_part}-{year_part}"
                full_date_match = re.fullmatch(r'\d{1,2}[-/\s]?(?:\d{1,2}|[A-Za-z]{3,9})[-/\s]?\d{2,4}', col.strip())                
                if full_date_match:
                    dt = pd.to_datetime(col.strip(), dayfirst=True)
                else:
                    dt = pd.to_datetime("01-" + clean_col, dayfirst=True)
            else:
                dt = pd.to_datetime(col, dayfirst=True)
            normalized = datetime.datetime(dt.year, dt.month, 1)
            df.columns.values[idx] = normalized
            return True
        except Exception:
            return False

    month_indices = [i for i, col in enumerate(df.columns) if is_month_column(col, i)]
    if not month_indices:
        raise ValueError("No month-formatted columns found in header.")
    
    return month_indices[0], month_indices[-1]

def read_excel_file(file_path: str, use_xls=False):
    try:
        engine = "xlrd" if use_xls else None
        raw_df = pd.read_excel(file_path, header=None, dtype=str, engine=engine)
        # Detect header row
        header_row = detect_header_row(raw_df)
        if header_row is None:
            print("Error: Could not detect header row.")
            return

        # Load actual DataFrame with detected header
        df = pd.read_excel(file_path, header=header_row, engine=None)

        # Clean up (drop fully empty rows, reset index)
        df.dropna(how='all', inplace=True)
        df.dropna(subset=["Items"], inplace=True)
        df.reset_index(drop=True, inplace=True)
        return df
    except Exception as e:
        print(f"Error while processing the Excel file: {e}")
        raise e


def main():
    parser = argparse.ArgumentParser(description="DoubleEntryPro - Clean and process Excel accounting files.")
    parser.add_argument('--path', type=str, required=True, help='Path to the Excel file (.xls or .xlsx)')
    args = parser.parse_args()

    # Normalize the path
    file_path = os.path.abspath(args.path)

    # Validate file exists
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    # Read the file raw (no header, all as strings)
    file_type = file_path.split(".")[-1]
    if file_type == "csv":
        df = pd.read_csv(file_path)
        print("CSV file loaded successfully:")
    elif file_type == "xlsx": 
        df = read_excel_file(file_path)
        print("Excel file loaded successfully:")
    elif file_type == "xls":
        df = read_excel_file(file_path, use_xls=True)
        print("Excel file loaded successfully:")

    
    else:
        print(f"Invalid input file type:   {file_type}.\nPlease provide a .xls, .xlsx or .csv file.")
        return
    
    start_idx, end_idx = parse_month_cols(df)
    
    while True:
        target = input("Please enter the month and year to process (MMM-YY):\n").strip()
        try:
            target_datetime = pd.to_datetime("01-" + target)
        except Exception:
            print(f"Sorry, {target} is not a valid month. Please use format MMM-YY, MMM-YYYY or MMMYY, MMMYYYY.")
            continue
        if target_datetime in df.columns:
            break
        else:
            print(f"Sorry, the input document only has amortizations from {df.columns[start_idx].strftime('%b %y')} to {df.columns[end_idx].strftime('%b %y')}.\nPlease enter a month within that range.")
    
    
    filtered = df[df[target_datetime].notna() & (df[target_datetime] != 0)]
    
    prepay_ledger_code = input("Please enter your Prepayments Ledger Code:\n").upper()
    
    year = target_datetime.year
    month = target_datetime.month
    last_day = calendar.monthrange(year, month)[1]
    last_day_of_month = datetime.datetime(year, month, last_day)
    date = last_day_of_month.strftime('%d/%m/%Y')  # e.g., '31 May 24'
    
    rows = []
    for idx, val in zip(filtered.index, filtered[target_datetime]):
        item = "Prepayment amortization for " + df.at[idx, "Items"].title()
        reference = int(df.at[idx, "Invoice number"])
        amount = abs(round(val, 2))
        exp_ledger_code = input(f"Please enter the expense ledger code for\t{df.at[idx, 'Items'].title()}:\n").upper()
        # Add positive and negative entries
        rows.append({"Date": date, "Description": item, "Reference": reference, "Account": exp_ledger_code, "Amount": amount})
        rows.append({"Date": date, "Description": item, "Reference": reference, "Account": prepay_ledger_code, "Amount": -amount})

    # Ensure the outputs directory exists
    output_dir = os.path.join(os.path.dirname(__file__), "outputs")
    os.makedirs(output_dir, exist_ok=True)

    # Determine the next output index
    existing_files = [f for f in os.listdir(output_dir) if f.endswith(".csv") and f[:-4].isdigit()]
    existing_indices = sorted([int(f[:-4]) for f in existing_files])
    out_idx = existing_indices[-1] + 1 if existing_indices else 0

    output_df = pd.DataFrame(rows)
    output_df.to_csv(f"{output_dir}/{out_idx}.csv", index=False)
    print(f"Entries written to {output_dir}/{out_idx}.csv")


if __name__ == "__main__":
    main()
