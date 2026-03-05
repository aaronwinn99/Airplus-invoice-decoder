import pandas as pd
from da import df

# ========== 1300.py Logic ==========
unique_df = df.drop_duplicates(subset=['Invoice No', 'Item No'], keep='first')

result_1300 = unique_df.groupby('Invoice No', as_index=False).agg({
    'Invoice Date': 'first',
    'Gross Amount': 'sum',
    'Sales Date': 'first',
})

result_1300['Gross Amount'] = result_1300['Gross Amount'].round(2)*-1  
result_1300['Posting type'] = 'AP'
result_1300['Account Code'] = '1300'
result_1300['Description'] = result_1300['Invoice No'].astype(str) + ' ' + result_1300['Invoice Date'].astype(str)


# ========== 1910.py Logic ==========
result_1910 = df.groupby('Invoice No', as_index=False).agg({
    'Invoice Date': 'first',
    'Tax': 'sum',
    'Sales Date': 'first',
})

result_1910['Posting type'] = 'GL'
result_1910['Account Code'] = '1910'
result_1910['Description'] = result_1910['Invoice No'].astype(str) + ' ' + result_1910['Invoice Date'].astype(str)

# Rename Tax column to match a consistent naming
result_1910.rename(columns={'Tax': 'Amount'}, inplace=True)


# ========== REST.py Logic ==========
df_rest = df[['Invoice No', 'Invoice Date', 'Place', 'Sales Date', 'Account Code', 'Cost Centre', 'Project No', 'Activity Code', 'Type', 'Service line2', 'Item No', 'Name', 'Travel Date', 'Net Amount (SC)', 'VAT Rate']].copy()
df_rest['Posting type'] = 'GL'
df_rest['Account Code'] = df_rest['Account Code'].astype('Int64') 
df_rest['Activity Code'] = df_rest['Activity Code'].astype('Int64')

# Rename Net Amount (SC) to Amount
df_rest.rename(columns={'Net Amount (SC)': 'Amount'}, inplace=True)

# Project No formatting
mask = df_rest['Project No'].str.len() == 16
mask1 = df_rest['Project No'].str.len() == 17
mask2 = df_rest['Project No'] == 'NO'
df_rest.loc[mask, 'Project No'] = df_rest.loc[mask, 'Project No'].str[:12] + '-0' + df_rest.loc[mask, 'Project No'].str[12:]
df_rest.loc[mask1, 'Project No'] = df_rest.loc[mask1, 'Project No'].str[:14] + '-' + df_rest.loc[mask1, 'Project No'].str[14:]
df_rest.loc[mask2, 'Project No'] = ''

# Account Code logic
mask_na = df_rest['Account Code'].isna()
mask_650 = df_rest['Project No'].str.endswith('650').fillna(False)
df_rest.loc[mask_na & mask_650, 'Account Code'] = 4300
df_rest.loc[mask_na & ~mask_650, 'Account Code'] = 4301

# Activity Code logic
activity_codes = []
last_activity = None

for idx, row in df_rest.iterrows():
    if pd.notna(row['Type']):
        if row['Type'] == 'FL':
            activity_codes.append(100)
            last_activity = 100
        elif row['Type'] == 'DO':
            activity_codes.append(101)
            last_activity = 101
        elif row['Type'] == 'SO':
            if pd.notna(row['Service line2']):
                if 'Flug' in str(row['Service line2']):
                    activity_codes.append(100)
                    last_activity = 100
                elif 'Bahn' in str(row['Service line2']):
                    activity_codes.append(101)
                    last_activity = 101
                else:
                    activity_codes.append(last_activity)
            else:
                activity_codes.append(last_activity)
        else:
            activity_codes.append(None)
    else:
        activity_codes.append(last_activity)

df_rest['Activity Code'] = activity_codes

# Description column
df_rest['Description'] = (
    df_rest['Invoice No'].astype(str) + ' POS' + 
    df_rest['Item No'].astype(str) + ' ' + 
    df_rest['Name'].fillna('') + ' ' + 
    df_rest['Travel Date'].fillna('')
).str.strip()

# Tax Code logic based on VAT rate
tax_codes = []
for idx, row in df_rest.iterrows():
    vat_rate = row['VAT Rate']
    if vat_rate == 19:
        tax_codes.append('SP')
    elif vat_rate == 7:
        tax_codes.append('PL')
    else:
        tax_codes.append('ZE')

df_rest['Tax Code'] = tax_codes

df_rest = df_rest.drop(columns=['Service line2', 'Item No', 'Name', 'Travel Date', 'Place', 'Type', 'VAT Rate'])


# ========== Combine All Tables ==========
# Rename amount columns for consistency  
result_1300.rename(columns={'Gross Amount': 'Amount'}, inplace=True)
result_1910.rename(columns={'Tax': 'Amount'}, inplace=True)
# df_rest already has 'Amount' column (renamed from Net Amount (SC))
df_rest_combined = df_rest.copy()

# Add Cur_amount column (copy of Amount) to all dataframes
result_1300['Cur_amount'] = result_1300['Amount']
result_1910['Cur_amount'] = result_1910['Amount']
df_rest_combined['Cur_amount'] = df_rest_combined['Amount']

# Add missing columns to align dataframes
for col in df_rest_combined.columns:
    if col not in result_1300.columns:
        result_1300[col] = None
    if col not in result_1910.columns:
        result_1910[col] = None

for col in result_1300.columns:
    if col not in df_rest_combined.columns:
        df_rest_combined[col] = None
        
for col in result_1910.columns:
    if col not in df_rest_combined.columns:
        df_rest_combined[col] = None

# Concatenate all dataframes vertically using append
combined_df = pd.concat([result_1300, result_1910, df_rest_combined], 
                        ignore_index=True, 
                        sort=False)

# Reorder columns in desired sequence
desired_columns = ['Account Code', 'Cost Centre', 'Project No', 'Activity Code', 
                   'Invoice No', 'Invoice Date', 'Cur_amount', 'Amount', 
                   'Description', 'Tax Code', 'Posting type']

combined_df = combined_df[desired_columns]

# Transform data types
combined_df['Account Code'] = pd.to_numeric(combined_df['Account Code'], errors='coerce').astype('Int64')  # Convert to integer
combined_df['Cost Centre'] = combined_df['Cost Centre'].astype('string')  # Convert to text
combined_df['Invoice Date'] = pd.to_datetime(combined_df['Invoice Date'], format='%d.%m.%Y')  # Convert to date

# ========== Export to Excel ==========
import os
excel_path = os.path.join(os.path.dirname(__file__), '..', 'Combined_Output.xlsx')

with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    combined_df.to_excel(writer, sheet_name='All Data', index=False)

print(f"Total rows: {len(combined_df)}")
print(f"\nCombined Data (first 5 rows - Account 1300):")
print(combined_df.head(5))
print(f"\nCombined Data (rows 3-7 - Account 1910):")
print(combined_df.iloc[2:7])
print(f"\nCombined Data (rows 8-13 - GL Details):")
print(combined_df.iloc[7:13])
print(f"\nCombined Data (last 5 rows):")
print(combined_df.tail(5))
print(f"\n✅ Excel file saved to: {excel_path}")
print(f"Total Shape: {combined_df.shape}")
