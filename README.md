# MyBCA E-statement Processing Script
This Python script is designed to process bank statements in PDF format, extract transaction data, calculate balances, and save the data into an Excel spreadsheet. It utilizes various libraries such as pandas, numpy, tabula, tqdm, and openpyxl.

## Features
1. PDF Parsing: Uses the tabula library to extract tabular data from PDF files.
2. Transaction Extraction: Extracts transaction data from bank statements, including dates, descriptions, amounts, and transaction types.
3. Balance Calculation: Calculates the balance for each transaction based on the previous balance and transaction amount.
4. Output to Excel: Saves the processed transaction data into an Excel spreadsheet, with each statement represented as a separate sheet.
5. Sheet Reordering: Reorders the sheets in the Excel file based on the period (year and month) of each statement.

## Usage
1. Ensure you have Python installed on your system.
2. Install the required libraries by running:
```bash
  pip install pipenv
  pipenv shell
  pipenv install
```
3. Place your bank statement PDF files in a folder named statements.
Run the script bank_statement_processing.py.

## Additional Notes
The script assumes the contents of ```statements``` folder are e-statements downloaded from myBCA application
