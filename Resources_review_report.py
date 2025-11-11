# Import libraries
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook

# Create dataframes from csv files
poline_df = pd.read_csv('Collection review - PO lines.csv')
collection_df = pd.read_csv('Electronic resources review collections.csv', dtype={'Electronic Collection Id': str})
portfolio_df = pd.read_csv('Electronic resources review individual subscriptions.csv',
                           dtype={'Electronic Collection Id': str, 'Portfolio Id': str})
expenditure_df = pd.read_csv('Collection review expenditure.csv')

merged_df = poline_df.merge(
    collection_df[['PO Line Reference', 'Electronic Collection Public Name', 'Electronic Collection Id', 
                   'License Name', 'Electronic Collection Additional PO Lines', 'Electronic Collection Linked To CZ']],
    on='PO Line Reference',
    how='left'  # Keeps all rows from poline_df and adds matching data from collection_df
)

temp_df = merged_df.merge(
    portfolio_df[['PO Line Reference', 'Electronic Collection Public Name', 
                  'Electronic Collection Id',
                  'Portfolio License Name', 'License Name', 'Portfolio Additional PO Lines', 'Portfolio Linked To CZ']],
    on='PO Line Reference',
    how='left',  # Keep all rows from merged_df
    suffixes=('', '_portfolio')  # Temporarily rename conflicting columns
)

# Fill missing values in existing columns using portfolio data
for col in ['Electronic Collection Public Name', 'Electronic Collection Id', 'License Name']:
    temp_df[col] = temp_df[col].combine_first(temp_df[f"{col}_portfolio"])

# Drop extra columns from portfolio_df
final_df = temp_df.drop(columns=[f"{col}_portfolio" for col in ['Electronic Collection Public Name', 
                                                                'License Name']])

# Replace NaN with empty string for searching
final_df[['Electronic Collection Additional PO Lines', 'Portfolio Additional PO Lines']] = \
    final_df[['Electronic Collection Additional PO Lines', 'Portfolio Additional PO Lines']].fillna('')

# Add new column with default value
final_df["Is additional PO"] = ""

# Function to find matches and copy data
def copy_matching_data(row):
    po_ref = row['PO Line Reference']
    
    # Find rows where PO Line Reference appears in either of the two columns
    mask = final_df['Electronic Collection Additional PO Lines'].str.contains(fr'\b{po_ref}\b', na=False) | \
           final_df['Portfolio Additional PO Lines'].str.contains(fr'\b{po_ref}\b', na=False)
    
    matching_rows = final_df[mask]
    
    if not matching_rows.empty:
        # Copy data from the first matching row
        for col in [
            'Electronic Collection Public Name', 'Electronic Collection Id', 'License Name',
            'Electronic Collection Additional PO Lines', 'Portfolio License Name', 'Portfolio Additional PO Lines'
        ]:
            row[col] = matching_rows[col].iloc[0]  # Take the first match
        
        # Mark as additional PO
        row["Is additional PO"] = "Y"
    
    return row

# Apply function row-wise
final_df = final_df.apply(copy_matching_data, axis=1)

df = final_df.merge(
    expenditure_df[['PO Line Reference', 'Transaction Expenditure Amount', 'Transaction Date Fiscal Year']],
    on='PO Line Reference',
    how='left'  # Keeps all rows from poline_df and adds matching data from collection_df
)

df = df.reindex(columns=['PO Line Reference', 'PO Line Title', 'Status (Active)', 'Order Line Type', 'Vendor Name', 
                         'Net Price', 'Currency',  'Transaction Expenditure Amount','Transaction Date Fiscal Year', 
                         'Reporting Code Description - 1st', 'Electronic Collection Public Name',
                         'Electronic Collection Id', 'License Name','Electronic Collection Additional PO Lines', 
                         'Electronic Collection Linked To CZ', 'Portfolio License Name', 'Portfolio Additional PO Lines', 
                         'Portfolio Linked To CZ','Is additional PO' ])

df.to_excel('Resource review.xlsx', index=False) 

df.rename(columns={
    'PO Line Reference': 'Purchase Order Number',
    'PO Line Title': 'Electronic Resource Title',
    'Transaction Expenditure Amount': 'Expenditure Amount (AUD)',
    'Transaction Date Fiscal Year': 'Fiscal Year',
    'Reporting Code Description - 1st': 'Reporting Code',
    'Electronic Collection Public Name': 'Electronic Collection',
}, inplace=True)

df[['COUNTER stats', 'SUSHI in Alma', 'Able to provide stats?', 'Usage',
    'Amount of content overlap with other resources',
    'Relevance to disciplinary and campus needs',
    'Required for accreditation', 'User interface functionality']] = None

# Save DataFrame to Excel
df.to_excel('Resource review.xlsx', index=False)

# Load the workbook and select the active sheet
wb = load_workbook('Resource review.xlsx')
ws = wb.active

# Hide specific columns (e.g., columns C and E)
columns_to_hide = ['A', 'C', 'L', 'M', 'N','P', 'Q', 'S']  # Adjust as needed

for col in columns_to_hide:
    ws.column_dimensions[col].hidden = True

# Save the modified workbook
wb.save('Resource review.xlsx')
