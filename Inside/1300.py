import pandas as pd
from da import df

# Step 1: Get unique rows based on Invoice No and Item No (keep first occurrence)
unique_df = df.drop_duplicates(subset=['Invoice No', 'Item No'], keep='first')

# Step 2: Group by Invoice No and aggregate
# Keep first values of other columns and sum Gross Amount
result = unique_df.groupby('Invoice No', as_index=False).agg({
    'Invoice Date': 'first',
    'Cardholder': 'first',
    'Place': 'first',
    'Gross Amount': 'sum',
    'Account No': 'first',
    'Merchant': 'first',
    'Sales Date': 'first',
    # Add more columns as needed
})

result['Gross Amount'] = result['Gross Amount'].round(2)*-1  

# Step 4: Add "Posting type" column with "AP" for all rows
result['Posting type'] = 'AP'
result['Account Code'] = '1300'

# Step 5: Add Description column (Invoice No + Invoice Date)
result['Description'] = result['Invoice No'].astype(str) + ' ' + result['Invoice Date'].astype(str)

print("Final Result:")
print(result)
print(f"\nTotal Gross Amount: {result['Gross Amount'].sum()}")