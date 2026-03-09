import pandas as pd
import os


# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
# Go up one level to parent directory
parent_dir = os.path.dirname(script_dir)
csv_path = os.path.join(parent_dir, 'Ehotel.Re.csv')
csv_path1 = os.path.join(parent_dir, 'Aus.csv')

# Read CSV with correct delimiter and encoding
df = pd.read_csv(csv_path, sep=';', encoding='ISO-8859-15')
df1 = pd.read_csv(csv_path1, sep=';', encoding='ISO-8859-15')

# German to English column name translation dictionary
german_to_english = {
    'Kartennummer': 'Card Number',
    'Karteninhaber-Name': 'Cardholder Name',
    'Karteninhaber-Stadt': 'Cardholder City',
    'Rechnungsnummer': 'Invoice Number',
    'Rechnungsdatum': 'Invoice Date',
    'Bruttobetrag': 'Gross Amount',
    'Positionsnummer': 'Position Number',
    'Leistungsart': 'Service Type',
    'Dokumentennummer': 'Document Number',
    'Name': 'Name',
    'Routing': 'Routing',
    'Leistungserbringer': 'Service Provider',
    'Verkaufsdatum': 'Sale Date',
    'Reisedatum': 'Travel Date',
    'Klasse': 'Class',
    'Airline Code': 'Airline Code',
    'AirlineCode': 'Airline Code',
    'Verkaufswährung': 'Sales Currency',
    'Netto(AbrW)': 'Net (Billing Currency)',
    'MwSt(AbrW)': 'VAT (Billing Currency)',
    'Abrechnungswährung': 'Billing Currency',
    'Brutto(AW)': 'Gross (Billing Currency)',
    'Details': 'Details',
    'Personal-ID': 'Personal ID',
    'Dienststelle': 'Department',
    'Kostenstelle': 'Cost Center',
    'Abrechnungseinheit': 'Billing Unit',
    'Internes Konto': 'Internal Account',
    'Bearbeitungsdatum': 'Processing Date',
    'Projektnummer': 'Project Number',
    'Auftragsnummer': 'Order Number',
    'Aktionsnummer': 'Action Number',
    'Reiseziel': 'Destination',
    'Kundenreferenz': 'Customer Reference',
    'Nullrechnungsnummer': 'Zero Invoice Number',
    'IATA-Nummer': 'IATA Number',
    'MwSt-Satz(%)': 'VAT Rate (%)',
    'Geb.-Zeichen': 'Fee Mark',
    'Leistungscode': 'Service Code',
    'DOM-Kennzeichen': 'DOM Mark',
    'Fälligkeitstag': 'Due Date',
    'Zusatzversicherung': 'Additional Insurance',
    'Leistungsbeschreibung1': 'Service Description 1',
    'Leistungsbeschreibung2': 'Service Description 2',
    'Leistungsbeschreibung3': 'Service Description 3',
    'Gebühren': 'Fees',
    'A.I.D.A Nummer': 'A.I.D.A Number',
    'MwSt-Typ': 'VAT Type',
    'Umsatzsteuernummer': 'Sales Tax Number',
    'Leistungserbringer-Adresse': 'Service Provider Address',
    'Abrechnungsmodell': 'Billing Model',
    'Steuerverfahren': 'Tax Procedure',
    'Hotel-Land': 'Hotel Country',
    'Netto(VerkW)': 'Net (Sales Currency)',
    'MwSt(VerkW)': 'VAT (Sales Currency)',
    'Brutto(VW)': 'Gross (Sales Currency)',
    'Original-Hotel-Invoice-URLs': 'Original Hotel Invoice URLs',
    'Beleg': 'Receipt',
    'Belegdatum': 'Receipt Date',
    'Zahlbelege': 'Payment Receipts',
    'Abrechnungsschritt': 'Billing Step',
    'Bezahlart': 'Payment Method'
}

# Rename columns to English
df = df.rename(columns=german_to_english)
df1 = df1.rename(columns=german_to_english)

# Remove completely empty columns (those that are all NaN)
df = df.dropna(axis=1, how='all')
df1 = df1.dropna(axis=1, how='all')

# Remove trailing empty columns
df = df.loc[:, (df.astype(str) != '').any()]
df1 = df1.loc[:, (df1.astype(str) != '').any()]

# Convert VAT column to numeric for proper summing
df['VAT (Billing Currency)'] = df['VAT (Billing Currency)'].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
df1['VAT (Billing Currency)'] = df1['VAT (Billing Currency)'].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
df['Net (Billing Currency)'] = df['Net (Billing Currency)'].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
df1['Net (Billing Currency)'] = df1['Net (Billing Currency)'].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
df['Gross Amount'] = df['Gross Amount'].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
df1['Gross Amount'] = df1['Gross Amount'].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
df1['Invoice Date'] = pd.to_datetime(df1['Invoice Date'], errors='coerce')
df['VAT Rate (%)'] = df['VAT Rate (%)'].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
df1['VAT Rate (%)'] = df1['VAT Rate (%)'].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
df['Gross (Billing Currency)'] = df['Gross (Billing Currency)'].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
df1['Gross (Billing Currency)'] = df1['Gross (Billing Currency)'].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
# Save to Excel
# save_path = os.path.join(parent_dir, 'Cleaned_Test.xlsx')
# save_path1 = os.path.join(parent_dir, 'Cleaned_Aus.xlsx')
# df.to_excel(save_path, index=False)
# df1.to_excel(save_path1, index=False)
# print(f"\nCleaned data saved to: {save_path}")
# print(f"\nCleaned data saved to: {save_path1}")
# print("\nColumns:", df.columns.tolist())
# print("\nFirst few rows:\n")
# print(df.head())
# print("\nData types:")
# print(df.dtypes)

# print("\nShape:", df1.shape)
# print("\nColumns:", df1.columns.tolist())
# print("\nFirst few rows:\n")
# print(df1.head())
# print("\nData types:")
# print(df1.dtypes)


# save_path = os.path.join(parent_dir, 'Cleaned_Rechu.xlsx')
# save_path1 = os.path.join(parent_dir, 'Cleaned_Aus.xlsx')
# df.to_excel(save_path, index=False)
# df1.to_excel(save_path1, index=False)
# print(f"\nCleaned data saved to: {save_path}")
# print(f"\nCleaned data saved to: {save_path1}")


# work on the Rechun first

# total = df.groupby('Invoice Number', as_index=False).agg({
#     'Invoice Date': 'first',
#     'Gross Amount': 'first',
# })

# total['Posting type'] = 'AP'
# total['Account Code'] = '1300'

# print("Total Result:")
# print(total)

# tax = df.groupby('Invoice Number', as_index=False).agg({
#     'Invoice Date': 'first',
#     'VAT (Billing Currency)': 'sum',
# })
# tax['Posting type'] = 'GL'
# tax['Account Code'] = '1910'

# tax['Description'] = 'tax charged for ' + tax['Invoice Number'].astype(str) + ' ' + tax['Invoice Date'].astype(str)

# print("Tax Result:")
# print(tax)

# rest = df[['Invoice Number', 'Cost Center', 'Project Number', 'Net (Billing Currency)', 'Invoice Date', 'VAT Rate (%)']].copy()
# rest['Posting type'] = 'GL'
# rest['Account Code'] = rest.apply(lambda row: 4301 if str(row['Project Number']).endswith('650') else (4300 if pd.notna(row['Project Number']) and str(row['Project Number']).strip() != '' else 6300), axis=1)

# mask = rest['Project Number'].str.len() == 16
# mask1 = rest['Project Number'].str.len() == 17
# rest.loc[mask, 'Project Number'] = rest.loc[mask, 'Project Number'].str[:12] + '-0' + rest.loc[mask, 'Project Number'].str[12:]
# rest.loc[mask1, 'Project Number'] = rest.loc[mask1, 'Project Number'].str[:14] + '-' + rest.loc[mask1, 'Project Number'].str[14:]

# tax_codes = []
# for idx, row in rest.iterrows():
#     vat_rate = row['VAT Rate (%)']
#     if vat_rate == 19:
#         tax_codes.append('SP')
#     elif vat_rate == 7:
#         tax_codes.append('PL')
#     else:
#         tax_codes.append('ZE')

# rest['Tax Code'] = tax_codes
# rest['Description'] = 'posting for ' + rest['Invoice Number'].astype(str) + ' ' + rest['Invoice Date'].astype(str)

# print("Rest Result:")
# print(rest)

# work on the Ausleng


total = df1.groupby('Invoice Number', as_index=False).agg({
    'Invoice Date': 'first',
    'Gross (Billing Currency)': 'sum',
    'Receipt': 'first',
    'Payment Receipts': 'first',
    'Service Code': 'first',

})

total['Posting type'] = 'AP'
total['Account Code'] = '1300'
total['Invoice Number'] = total['Receipt'].astype(str)+' '+total['Payment Receipts'].astype(str)+' '+total['Service Code'].astype(str)
total['Description'] = total['Receipt'].astype(str) + ' Total'

print("Total Result:")
print(total)

tax = df1.groupby('Invoice Number', as_index=False).agg({
    'Receipt Date': 'first',
    'VAT (Billing Currency)': 'sum',
    'Receipt': 'first',
    'Payment Receipts': 'first',
    'Service Code': 'first',
})
tax['Posting type'] = 'GL'
tax['Account Code'] = '1910'

tax['Invoice Number'] = tax['Receipt'].astype(str)+' '+tax['Payment Receipts'].astype(str)+' '+tax['Service Code'].astype(str)

tax['Description'] = tax['Receipt'].astype(str) + ' Tax liability '

print("Tax Result:")
print(tax)

rest = df1[['Invoice Number', 'Cost Center', 'Net (Billing Currency)', 'Invoice Date', 'VAT Rate (%)', 'Receipt', 'Payment Receipts', 'Service Code', 'Position Number', 'Name', 'Service Description 1'] + (['Project Number'] if 'Project Number' in df1.columns else [])].copy()
rest['Invoice Number'] = rest['Receipt'].astype(str)+' '+rest['Payment Receipts'].astype(str)+' '+rest['Service Code'].astype(str)
rest['Posting type'] = 'GL'

# Set Account Code based on Project Number if it exists
if 'Project Number' in rest.columns:
    rest['Account Code'] = rest.apply(lambda row: 4301 if str(row['Project Number']).endswith('650') else (4300 if pd.notna(row['Project Number']) and str(row['Project Number']).strip() != '' else 6300), axis=1)
else:
    rest['Account Code'] = 6300

# Process Project Number if it exists
if 'Project Number' in rest.columns:
    mask = rest['Project Number'].str.len() == 16
    mask1 = rest['Project Number'].str.len() == 17
    rest.loc[mask, 'Project Number'] = rest.loc[mask, 'Project Number'].str[:12] + '-0' + rest.loc[mask, 'Project Number'].str[12:]
    rest.loc[mask1, 'Project Number'] = rest.loc[mask1, 'Project Number'].str[:14] + '-' + rest.loc[mask1, 'Project Number'].str[14:]

tax_codes = []
for idx, row in rest.iterrows():
    vat_rate = row['VAT Rate (%)']
    if vat_rate == 19:
        tax_codes.append('SP')
    elif vat_rate == 7:
        tax_codes.append('PL')
    else:
        tax_codes.append('ZE')

rest['Tax Code'] = tax_codes
rest['Description'] = rest['Invoice Number'].astype(str) + ' POS ' + rest['Position Number'].astype(str) + ' ' + rest['Name'].astype(str) + ' ' + rest['Service Description 1'].astype(str)

print("Rest Result:")
print(rest)