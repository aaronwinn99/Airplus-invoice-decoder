import pandas as pd
from da import df

tax = df.groupby('Invoice No', as_index=False).agg({
    'Invoice Date': 'first',
    'Place': 'first',
    'Tax': 'sum',
    'Sales Date': 'first',
    # Add more columns as needed
})

tax['Posting type'] = 'GL'
tax['Account Code'] = '1910'

# Add Description column (Invoice No + Invoice Date)
tax['Description'] = tax['Invoice No'].astype(str) + ' ' + tax['Invoice Date'].astype(str)

print("Final Result:")
print(tax)