import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Invoice Processor", layout="wide")
st.title("📊 Airplus Journal Generator")
st.markdown("Upload CSV files and process them to generate GL accounting entries")

def process_invoice_data(df):
    """Process invoice data through the transformation logic"""
    
    # Replace 'NO' in Project No with blank (case-insensitive)
    df = df.copy()
    df['Project No'] = df['Project No'].apply(lambda x: '' if (isinstance(x, str) and x.strip().upper() == 'NO') else x)
    
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
        'Tax(SC)': 'sum',
        'Sales Date': 'first',
    })
    
    result_1910['Posting type'] = 'GL'
    result_1910['Account Code'] = '1910'
    result_1910['Description'] = result_1910['Invoice No'].astype(str) + ' ' + result_1910['Invoice Date'].astype(str)
    result_1910.rename(columns={'Tax(SC)': 'Amount'}, inplace=True)
    
    # ========== REST.py Logic ==========
    df_rest = df[['Invoice No', 'Invoice Date', 'Place', 'Sales Date', 'Accounting Unit', 'Cost Centre', 'Project No', 'Type', 'Service line2', 'Item No', 'Name', 'Travel Date', 'Net Amount (SC)', 'VAT Rate']].copy()
    
    df_rest.rename(columns={'Accounting Unit': 'Account Code'}, inplace=True)
    df_rest['Posting type'] = 'GL'
    # Keep Account Code as string to preserve original values
    # Don't convert to numeric yet - we need original values for the logic below
    
    # Rename Net Amount (SC) to Amount
    df_rest.rename(columns={'Net Amount (SC)': 'Amount'}, inplace=True)
    
    # Initialize Activity Code column (will be populated below)
    df_rest['Activity Code'] = None
    
    # Project No formatting
    mask = df_rest['Project No'].str.len() == 16
    mask1 = df_rest['Project No'].str.len() == 17
    df_rest.loc[mask, 'Project No'] = df_rest.loc[mask, 'Project No'].str[:12] + '-0' + df_rest.loc[mask, 'Project No'].str[12:]
    df_rest.loc[mask1, 'Project No'] = df_rest.loc[mask1, 'Project No'].str[:14] + '-' + df_rest.loc[mask1, 'Project No'].str[14:]
    
    # Account Code logic - simple approach
    mask_na = df_rest['Account Code'].isna()  # Find NA values
    mask_650 = df_rest['Project No'].str.endswith('650')  # Project No ends with 650
    
    # If NA and ends with 650 → 4300
    df_rest.loc[mask_na & mask_650, 'Account Code'] = 4300
    
    # If NA and doesn't end with 650 → 4301
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
    
    # Convert Activity Code to Int64 where not None
    df_rest['Activity Code'] = pd.to_numeric(df_rest['Activity Code'], errors='coerce').astype('Int64')
    
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
    result_1300.rename(columns={'Gross Amount': 'Amount'}, inplace=True)
    df_rest_combined = df_rest.copy()
    
    # Add Cur_amount column (copy of Amount) to all dataframes
    result_1300['Cur_amount'] = result_1300['Amount']
    result_1910['Cur_amount'] = result_1910['Amount']
    df_rest_combined['Cur_amount'] = df_rest_combined['Amount']
    
    # Ensure all dataframes have the same columns with consistent dtypes
    all_cols = set(result_1300.columns) | set(result_1910.columns) | set(df_rest_combined.columns)
    for df_part in [result_1300, result_1910, df_rest_combined]:
        for col in all_cols:
            if col not in df_part.columns:
                df_part[col] = pd.NA
    
    # Concatenate all dataframes vertically
    combined_df = pd.concat([result_1300, result_1910, df_rest_combined], 
                            ignore_index=True, 
                            sort=False)
    
    # Reorder columns in desired sequence
    desired_columns = ['Account Code', 'Cost Centre', 'Project No', 'Activity Code', 
                       'Invoice No', 'Invoice Date', 'Cur_amount', 'Amount', 
                       'Description', 'Tax Code', 'Posting type']
    
    combined_df = combined_df[desired_columns]
    
    # Transform data types
    # Keep Account Code as string to preserve original values
    combined_df['Account Code'] = combined_df['Account Code'].astype('string')
    combined_df['Cost Centre'] = combined_df['Cost Centre'].astype('string')
    combined_df['Invoice Date'] = pd.to_datetime(combined_df['Invoice Date'], format='%d.%m.%Y')
    
    return combined_df


def load_and_clean_csv(uploaded_file):
    """Load CSV file and convert numeric values"""
    df = pd.read_csv(uploaded_file, sep=';', encoding='ISO-8859-15')
    
    # Convert comma decimals to dots for numeric columns
    numeric_columns = [
        'Gross Amount', 'Net Amount (SC)', 'Tax(SC)', 'Gross Amount (BC)',
        'VAT Rate'
    ]
    
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce')
    
    return df


# Sidebar for file upload
st.sidebar.header("📁 Upload Files")
uploaded_files = st.sidebar.file_uploader(
    "Choose CSV files to process",
    type=['csv'],
    accept_multiple_files=True
)

if uploaded_files:
    st.sidebar.success(f"✅ {len(uploaded_files)} file(s) uploaded")
    
    # Main processing area
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.subheader("Processing Options")
    
    with col2:
        process_button = st.button("🚀 Process Files", width='stretch')
    
    if process_button:
        try:
            with st.spinner('Loading and processing files...'):
                # Load all CSV files
                all_dataframes = []
                for uploaded_file in uploaded_files:
                    df = load_and_clean_csv(uploaded_file)
                    all_dataframes.append(df)
                
                # Combine all dataframes if multiple files
                if len(all_dataframes) > 1:
                    combined_input_df = pd.concat(all_dataframes, ignore_index=True)
                else:
                    combined_input_df = all_dataframes[0]
                
                st.info(f"📊 Loaded {len(combined_input_df)} total rows from {len(uploaded_files)} file(s)")
                
                # Process the data
                result_df = process_invoice_data(combined_input_df)
                
                st.success(f"✅ Processing complete! Generated {len(result_df)} GL entries")
                
                # Display preview
                st.subheader("📋 Preview of Processed Data")
                st.dataframe(result_df.head(10), width='stretch')
                
                # Generate Excel file
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    result_df.to_excel(writer, sheet_name='All Data', index=False)
                excel_buffer.seek(0)
                
                # Download button
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_buffer.getvalue(),
                    file_name="Combined_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch'
                )
                
                # Display statistics
                st.subheader("📈 Processing Statistics")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    ap_count = len(result_df[result_df['Posting type'] == 'AP'])
                    st.metric("AP Entries (1300)", ap_count)
                
                with col2:
                    tax_count = len(result_df[result_df['Posting type'] == 'GL'].groupby('Account Code').get_group(1910) if 1910 in result_df[result_df['Posting type'] == 'GL']['Account Code'].values else [])
                    st.metric("Tax Entries (1910)", tax_count)
                
                with col3:
                    gl_count = len(result_df[(result_df['Posting type'] == 'GL') & (result_df['Account Code'] != 1910)])
                    st.metric("GL Details", gl_count)
                
                with col4:
                    st.metric("Total Entries", len(result_df))
                
        except Exception as e:
            st.error(f"❌ Error processing files: {str(e)}")
            st.error("Please ensure your CSV files have the correct structure and encoding")

else:
    st.info("👈 Upload CSV files using the sidebar to get started")
    st.markdown("""
    ## How it works:
    1. **Upload** one or more CSV files using the file uploader in the sidebar
    2. **Click** the "Process Files" button to transform the data
    3. **Download** the processed Excel file with GL entries
    
    ### Expected CSV Columns:
    - Invoice No, Invoice Date, Gross Amount, Tax(SC), Net Amount (SC)
    - Accounting Unit, Cost Centre, Project No
    - Type, Service line2, Item No, Name, Travel Date, VAT Rate
    - Plus other supporting columns (Place, Sales Date, etc.)
    """)
