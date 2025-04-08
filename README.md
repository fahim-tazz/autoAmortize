# DoubleEntryPro

AutoAmortize is a Python script for processing amortization schedules from Excel or CSV files and generating double-entry accounting records automatically.

## Setup Instructions

1. **Clone the Repository**

```bash
git clone https://github.com/fahim-tazz/autoAmortize.git
cd autoAmortize
```

2. **Set Up the Environment**

   We recommend using a virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
```

3. **Install Dependencies**

```bash
pip install -r requirements.txt
```

## Running the Script

Run the script with an Excel or CSV file path:

```bash
python main.py --path path/to/your/file.xlsx
```

- The script supports `.xlsx`, `.xls`, and `.csv` formats.
- Month columns must be reasonably formatted or recognizable (e.g., `May24`, `15/05/2024`, `May-2024`, etc.). US-style formats `Month-Date-Year` are not supported.

## Output

- Processed accounting entries are saved in the `outputs/` folder.
- Files are automatically numbered (`0.csv`, `1.csv`, etc.) to avoid overwrites.

## Customization

You can modify the script to:

- Adjust column keywords (e.g., `"Items"`, `"Invoice number"`, etc.)
- Tweak output format as needed for your accounting platform
- Hardcode ledger codes for Prepayments (and Expenses, if appropriate for your use case).

Created by [@fahim-tazz](https://github.com/fahim-tazz)

---

## Technical Report

### Scalability

This script is scalable to any number of months in the amortization schedule. It dynamically identifies and extracts month columns without assuming a fixed number or position. The month extraction logic ensures that even with extended schedules (e.g., 24 or 36 months), the script continues to work seamlessly without manual adjustment.

### Flexibility and Robustness

The script is designed to handle a wide range of real-world inconsistencies:

- **Excel File Structure:** It handles Excel files with variable table positions — including extra rows above or below the table, and empty rows within the data.
- **Flexible Month Parsing:** Whether reading from CSV or Excel, the script supports various month formats like `May24`, `May-24`, `May2024`, `May 2024`, `May-2024`, `15/05/2024`, and `15 May 2024`. It automatically normalizes all month references to the first day of the month for consistent processing.
- **Input Flexibility:** When prompting the user for the target month, the script accepts a wide range of formats and intelligently converts them into a valid datetime object that matches our normalized Dataframe.
- **Robust Column Matching:** Even if headers are in string form or native datetime objects (e.g., Excel cell formats), the script converts and matches them reliably.

### Correctness

We assume that the user is using a central Prepayments Ledger for all Prepayment Credit entries, and an itemized Expense ledger for each type of expense for Expense Debit entries.

To ensure the accounting entries are correct:

- The user is prompted once for the centralized **Prepayments Ledger Code**, which is reused across all entries.
- For each unique expense type (e.g., Web Hosting, Insurance), the script prompts the user individually for the corresponding **Expense Ledger Code** (e.g. Insurance Expense Ledger).
- This separation ensures accuracy in accounting records since the amortization schedule does not include ledger codes explicitly.

While prompting for expense ledger codes per-item may be slightly tedious for the user, this design guarantees that all accounting entries are sent to the right account in the user's system based on user-confirmed inputs, especially since the expense ledger codes are not provided in the amortization schedule.
