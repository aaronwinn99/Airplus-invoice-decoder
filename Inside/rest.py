import pandas as pd
from da import df

# First, keep Type and Merchant/Service columns temporarily to determine flight vs train
df_full = df.copy()

df = df[['Invoice No', 'Invoice Date', 'Place', 'Tax', 'Sales Date', 'Place', 'Sales Date', 'Account Code', 'Cost Centre', 'Project No', 'Activity Code', 'Type', 'Merchant', 'Service line2', 'Item No', 'Name', 'Travel Date']].copy()
df['Posting type'] = 'GL'
df['Account Code'] = df['Account Code'].astype('Int64') 
df['Activity Code'] = df['Activity Code'].astype('Int64')

# Add '-' after 11th character if Project No has 16 characters
mask = df['Project No'].str.len() == 16
mask1 = df['Project No'].str.len() == 17
mask2 = (df['Project No'].str.len() < 16) & (df['Project No'].notna() & (df['Project No'] != 'NO'))
df.loc[mask, 'Project No'] = df.loc[mask, 'Project No'].str[:12] + '-0' + df.loc[mask, 'Project No'].str[12:]
df.loc[mask1, 'Project No'] = df.loc[mask1, 'Project No'].str[:14] + '-' + df.loc[mask1, 'Project No'].str[14:]
df.loc[mask2, 'Project No'] = df.loc[mask2, 'Project No'] + ' PROJECT NUMBER INC'

# Replace NA Account Code based on Project No ending
mask_na = df['Account Code'].isna()  # Find NA values
mask_650 = df['Project No'].str.endswith('650').fillna(False)  # Project No ends with 650

# If NA and ends with 650 → 4300
df.loc[mask_na & mask_650, 'Account Code'] = 4300

# If NA and doesn't end with 650 → 4301
df.loc[mask_na & ~mask_650, 'Account Code'] = 4301

# Set Activity Code based on flight vs train
# FL = Flight → 100, DO = Train → 101
# SO = Service → Check Service line2 for "SE Flug" (100) or "SE Bahn" (101)
activity_codes = []
last_activity = None

for idx, row in df.iterrows():
    if pd.notna(row['Type']):
        if row['Type'] == 'FL':
            activity_codes.append(100)
            last_activity = 100
        elif row['Type'] == 'DO':
            activity_codes.append(101)
            last_activity = 101
        elif row['Type'] == 'SO':
            # Check Service line2 for flight or train service
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
                # No Service line2, inherit from previous
                activity_codes.append(last_activity)
        else:
            activity_codes.append(None)
    else:  # Detail row without type, inherit from previous
        activity_codes.append(last_activity)

df['Activity Code'] = activity_codes

# Create Description column: Invoice No + POS + Item No + Traveller + Travel Date
df['Description'] = (
    df['Invoice No'].astype(str) + ' POS' + 
    df['Item No'].astype(str) + ' ' + 
    df['Name'].fillna('') + ' ' + 
    df['Travel Date'].fillna('')
).str.strip()

# Drop Type, Merchant, Service line2, Item No, Name, Travel Date columns as they're no longer needed
df = df.drop(columns=['Type', 'Merchant', 'Service line2', 'Item No', 'Name', 'Travel Date', 'Place', 'Tax', 'Sales Date'])

print(df)
print(df.columns)