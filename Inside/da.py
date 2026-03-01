import pandas as pd
import os


# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
# Go up one level to parent directory
parent_dir = os.path.dirname(script_dir)
csv_path = os.path.join(parent_dir, 'Test.csv')

# Read CSV with correct delimiter and encoding
df = pd.read_csv(csv_path, sep=';', encoding='ISO-8859-15')

# Remove completely empty columns (those that are all NaN)
df = df.dropna(axis=1, how='all')

# Remove trailing empty columns
df = df.loc[:, (df.astype(str) != '').any()]

# Replace comma decimal separators with dots for numeric columns
numeric_cols = ['Gross Amount', 'Net Amount (SC)', 'Tax(SC)', 'Gross Amount (BC)', 'VAT Rate', 'Fees (Tax)']
for col in numeric_cols:
    if col in df.columns:
        df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
        df[col] = pd.to_numeric(df[col], errors='coerce')

df.columns = df.columns.str.replace('Tax(SC)', 'Tax', regex=False)
df.columns = df.columns.str.replace('Order No', 'Activity Code', regex=False)
df.columns = df.columns.str.replace('Action No', 'Account Code', regex=False)


# Display cleaned data
    # print("Shape:", df.shape)
    # print("\nColumns:", df.columns.tolist())
    # print("\nFirst few rows:\n")
    # print(df.head())
    # print("\nData types:")
    # print(df.dtypes)

# Optional: Save cleaned data as CSV
# save_path = os.path.join(parent_dir, 'Cleaned_Test.csv')
# df.to_csv(save_path, sep=';', encoding='ISO-8859-15', index=False)
# print(f"\nCleaned data saved to: {save_path}")
