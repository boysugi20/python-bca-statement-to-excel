from tabula import read_pdf
import pandas as pd
import numpy as np
import os
from tqdm import tqdm

def is_currency(value):
    if pd.isna(value) or value == '':
        return False
    try:
        float(str(value).replace(',', ''))
    except ValueError:
        return False
    else:
        return True


def clean_numeric_columns(dataframe, columns):

    for column in columns:
        dataframe[column] = dataframe[column].str.replace(',', '')
        dataframe[column] = pd.to_numeric(dataframe[column], errors='coerce')
        dataframe[column] = dataframe[column].astype('float')

    return dataframe


def union_source(dataframes):

    dfs = []
    for temp_df in dataframes:

        # Split DB into new column
        temp_df[['amount', 'type']] = temp_df[4].str.extract(r'([\d,]+(?:\.\d+)?)\s*(DB|CR)?')
        temp_df = temp_df.drop(temp_df.columns[4], axis=1)

        if(len(temp_df.columns) == 7):
            # Name column and reorder
            temp_df.columns = ['date', 'desc', 'detail', 'branch', 'balance', 'amount', 'type']
            temp_df = temp_df[['date', 'desc', 'detail', 'branch', 'amount', 'type', 'balance']]

            dfs.append(temp_df)

    df = pd.concat(dfs, ignore_index=True)
    df = df.fillna(value=np.nan)
                
    return df


def insert_shifted_column(dataframe):

    # Add new columns with shifted values for comparison
    dataframe['prev_date'] = dataframe['date'].shift(1)
    dataframe['prev_desc'] = dataframe['desc'].shift(1)
    dataframe['prev_detail'] = dataframe['detail'].shift(1)
    dataframe['prev_branch'] = dataframe['branch'].shift(1)
    dataframe['prev_amount'] = dataframe['amount'].shift(1)
    dataframe['prev_transaction_type'] = dataframe['type'].shift(1)
    dataframe['prev_balance'] = dataframe['balance'].shift(1)

    dataframe = dataframe.fillna(value=np.nan)

    return dataframe


def extract_transactions(dataframe):

    transactions = []
    details = []
    descs = []
    temp = {}

    for index, row in dataframe.iterrows():

        if (row['desc'] == 'DR KOREKSI BUNGA') or (row['desc'] == 'BUNGA'):
            transaction = {
                "date": temp['date'],
                "desc": ' | '.join(descs) if descs else '',
                "detail": ' | '.join(details) if details else '',
                "branch": temp['branch'],
                "amount": temp['amount'],
                "transaction_type": temp['transaction_type'] if temp['transaction_type'] == 'DB' else 'CR',
                "balance": temp['balance']
            }
            transactions.append(transaction)
            break

        # New Transaction
        if(not pd.isna(row['amount'])) and (pd.isna(row['prev_amount'])):

            # Save previous transaction
            if temp:
                transaction = {
                    "date": temp['date'],
                    "desc": ' | '.join(descs) if descs else '',
                    "detail": ' | '.join(details) if details else '',
                    "branch": temp['branch'],
                    "amount": temp['amount'],
                    "transaction_type": temp['transaction_type'] if temp['transaction_type'] == 'DB' else 'CR',
                    "balance": temp['balance']
                }
                transactions.append(transaction)
                details = []
                descs = []
                temp = {}

            temp = {
                'date': row['date'],
                'branch': row['branch'],
                'amount': row['amount'],
                'transaction_type': row['type'],
                'balance': row['balance']
            }

        if (not pd.isna(row['desc'])):
            descs.append(row['desc'])
        if (not pd.isna(row['detail'])):
            details.append(row['detail'])

    transaction_dataframe = pd.DataFrame(transactions)

    return transaction_dataframe


def save_to_excel(dataframe, output_filename):

    if os.path.isfile(output_filename):
        writer = pd.ExcelWriter(output_filename, engine="openpyxl", mode='a', if_sheet_exists='replace')
    else:
        writer = pd.ExcelWriter(output_filename, engine="openpyxl")

    transaction_dataframe.to_excel(writer, sheet_name=periode, index=False)

    workbook = writer.book
    worksheet = writer.sheets[periode]

    # Format column into currency IDR
    for cell in worksheet['E']:
        cell.number_format = '_-Rp* #,##0.00_-;[Red]-Rp* #,##0.00_-;_-Rp* "-"_-;_-@_-'
    for cell in worksheet['G']:
        cell.number_format = '_-Rp* #,##0.00_-;[Red]-Rp* #,##0.00_-;_-Rp* "-"_-;_-@_-'

    # Autofit column width
    for column_cells in worksheet.columns:
        max_length = max(len(str(cell.value)) for cell in column_cells)
        if column_cells[0].column in ['E', 'G']:
            worksheet.column_dimensions[column_cells[0].column_letter].width = max_length + 50
        else:
            worksheet.column_dimensions[column_cells[0].column_letter].width = max_length + 2

    writer.close()

    return


statements = "statements"
pbar = tqdm(os.listdir(statements))
for filename in pbar:
    file_path = os.path.join(statements, filename)
    pbar.set_description("Processing %s" % filename)

    if os.path.isfile(file_path):

        # Get header information
        header_dataframe = read_pdf(file_path, area = (95, 324, 155, 570), pages='1', pandas_options={'header': None, 'dtype': str}, force_subprocess=True)[0]
        periode = header_dataframe.loc[header_dataframe[0] == 'PERIODE', 2].values[0]
        periode = ' '.join(reversed(periode.split()))
        no_rekening = header_dataframe.loc[header_dataframe[0] == 'NO. REKENING', 2].values[0]

        dataframes = read_pdf(file_path, area = (251, 25, 805, 577), columns=[86, 184, 300, 340, 467], pages='all', pandas_options={'header': None, 'dtype': str}, force_subprocess=True)

        df = union_source(dataframes)
        df = clean_numeric_columns(df, ['amount', 'balance'])
        df = insert_shifted_column(df)

        transaction_dataframe = extract_transactions(df)

        save_to_excel(transaction_dataframe, f'{no_rekening}.xlsx')
