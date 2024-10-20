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
3. Create a folder named ```statements```
4. Place your bank statement PDF files in the ```statements``` folder
5. Run the script ```python main.py```

## Additional Notes
The script assumes the contents of ```statements``` folder are e-statements downloaded from myBCA application

## Output Example
![Screenshot 2024-10-18 003737](https://github.com/user-attachments/assets/f39e8c65-d919-47fd-a9bd-ec1c2a7d85ed)
