import streamlit as st
import pandas as pd
import io
import os

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================

st.set_page_config(
    page_title="Invoice Processor",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# NAVIGATION
# ============================================================================

page = st.sidebar.radio(
    "📑 Select file",
    ["🛫 Airplus", "🏨 Ehotel - Rechnung", "🏨 Ehotel - Auslagen"],
    label_visibility="visible"
)

# ============================================================================
# AIRPLUS PAGE 
# ============================================================================

if page == "🛫 Airplus":
    st.title("📊 Airplus Journal Generator")
    st.markdown("Upload CSV files and process them to generate GL accounting entries")

    def process_invoice_data(df):
        """Process invoice data through the transformation logic"""
        
        # Replace 'NO' in Project No with blank (case-insensitive)
        df = df.copy()
        df.columns = df.columns.str.replace('Order No', 'Activity Code', regex=False)
        df.columns = df.columns.str.replace('Action No', 'Account Code', regex=False)
        df['Project No'] = df['Project No'].apply(lambda x: '' if (isinstance(x, str) and x.strip().upper() == 'NO') else x)
        
        # ========== 1300.py Logic ==========
        unique_df = df.drop_duplicates(subset=['Invoice No', 'Item No'], keep='first')
        
        result_1300 = unique_df.groupby('Invoice No', as_index=False).agg({
            'Invoice Date': 'first',
            'Gross Amount': 'sum',
            'Sales Date': 'first',
        })
        
        result_1300['Invoice No'] = result_1300['Invoice No'].str.replace(' ', '', regex=False)
        result_1300['Gross Amount'] = result_1300['Gross Amount'].round(2)*-1  
        result_1300['Posting type'] = 'AP'
        result_1300['Account Code'] = 1300
        result_1300['Description'] = result_1300['Invoice No'].astype(str) + ' ' + result_1300['Invoice Date'].astype(str)
        result_1300['SUPID'] = 41297
        
        # ========== 1910.py Logic ==========
        result_1910 = df.groupby('Invoice No', as_index=False).agg({
            'Invoice Date': 'first',
            'Tax(SC)': 'sum',
            'Sales Date': 'first',
        })
        
        result_1910['Invoice No'] = result_1910['Invoice No'].str.replace(' ', '', regex=False)
        result_1910['Posting type'] = 'GL'
        result_1910['Account Code'] = 1910
        result_1910['Description'] = result_1910['Invoice No'].astype(str) + ' ' + result_1910['Invoice Date'].astype(str)
        result_1910.rename(columns={'Tax(SC)': 'Amount'}, inplace=True)
        result_1910['SUPID'] = 41297
        
        # ========== REST.py Logic ==========
        df_rest = df[['Invoice No', 'Invoice Date', 'Place', 'Sales Date', 'Account Code', 'Cost Centre', 'Project No', 'Type', 'Service line2', 'Item No', 'Name', 'Travel Date', 'Net Amount (SC)', 'VAT Rate', 'Routing']].copy()
        df_rest['Invoice No'] = df_rest['Invoice No'].str.replace(' ', '', regex=False)
        
        df_rest['Posting type'] = 'GL'
        df_rest['SUPID'] = 41297
        
        # Track originally NA values BEFORE any conversion
        mask_na = df_rest['Account Code'].isna()
        
        # Convert to Int64 exactly like total.py does
        df_rest['Account Code'] = df_rest['Account Code'].astype('Int64')
        
        # Rename Net Amount (SC) to Amount
        df_rest.rename(columns={'Net Amount (SC)': 'Amount'}, inplace=True)
        
        # Initialize Activity Code column (will be populated below)
        df_rest['Activity Code'] = None
        
        # Project No formatting
        mask = df_rest['Project No'].str.len() == 16
        mask1 = df_rest['Project No'].str.len() == 17
        mask2 = df_rest['Project No'] == 'NO'
        df_rest.loc[mask, 'Project No'] = df_rest.loc[mask, 'Project No'].str[:12] + '-0' + df_rest.loc[mask, 'Project No'].str[12:]
        df_rest.loc[mask1, 'Project No'] = df_rest.loc[mask1, 'Project No'].str[:14] + '-' + df_rest.loc[mask1, 'Project No'].str[14:]
        df_rest.loc[mask2, 'Project No'] = ''
        
        # Account Code logic - only apply to originally NA values
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
        
        # Convert Activity Code to Int64 where not None
        df_rest['Activity Code'] = pd.to_numeric(df_rest['Activity Code'], errors='coerce').astype('Int64')
        
        # Description column
        df_rest['Description'] = (
            df_rest['Invoice No'].astype(str) + ' POS' + 
            df_rest['Item No'].astype(str) + ' ' + 
            df_rest['Name'].fillna('') + ' ' + 
            df_rest['Travel Date'].fillna('') + ' ' +
            df_rest['Routing'].fillna('')
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
                           'Description', 'Tax Code', 'Posting type', 'SUPID']
        
        combined_df = combined_df[desired_columns]
        
        # Transform data types
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

# ============================================================================
# EHOTEL RECHNUNG PAGE
# ============================================================================

elif page == "🏨 Ehotel - Rechnung":
    st.title("📊 Ehotel Rechnung Journal Generator")
    st.markdown("Upload Rechnung CSV files and generate GL accounting entries")

    def process_ehotel_rechuntype(df):
        """Process Ehotel Rechnung invoice data"""
        
        df = df.copy()
        
        # Convert numeric columns with comma decimals
        numeric_cols = ['VAT (Billing Currency)', 'Net (Billing Currency)', 'Gross Amount', 'Gross (Billing Currency)', 'VAT Rate (%)']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
        
        # Convert date columns
        date_cols = ['Invoice Date', 'Receipt Date']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], format='%d.%m.%Y', errors='coerce')
        
        # ========== TOTAL (AP - 1300) ==========
        total = df.groupby('Invoice Number', as_index=False).agg({
            'Invoice Date': 'first',
            'Gross (Billing Currency)': 'sum',
        })
        total['Posting type'] = 'AP'
        total['Account Code'] = 1300

        total['Amount'] = total['Gross (Billing Currency)'] * -1
        total['Cur_amount'] = total['Amount']
        total['Description'] = total['Invoice Number'].astype(str) + ' Total'
        total['SUPID'] = 41253
        
        # ========== TAX (GL - 1910) ==========
        tax = df.groupby('Invoice Number', as_index=False).agg({
            'Invoice Date': 'first',
            'VAT (Billing Currency)': 'sum',
        })
        tax['Posting type'] = 'GL'
        tax['Account Code'] = 1910
        tax['Amount'] = tax['VAT (Billing Currency)']
        tax['Cur_amount'] = tax['Amount']
        tax['Description'] = tax['Invoice Number'].astype(str) + ' Tax liability '
        tax['SUPID'] = 41253
        # ========== DETAILS (GL - 4300/4301) ==========
        cols_to_use = ['Invoice Number', 'Cost Center', 'Net (Billing Currency)', 'Invoice Date'] + \
                      (['Position Number'] if 'Position Number' in df.columns else []) + \
                      (['Name'] if 'Name' in df.columns else []) + \
                      (['Service Description 1'] if 'Service Description 1' in df.columns else []) + \
                      (['VAT Rate (%)'] if 'VAT Rate (%)' in df.columns else []) + \
                      (['Project Number'] if 'Project Number' in df.columns else [])
        rest = df[cols_to_use].copy()
        rest['Posting type'] = 'GL'
        rest['Amount'] = rest['Net (Billing Currency)']
        rest['Cur_amount'] = rest['Amount']
        rest['SUPID'] = 41253
        
        # Account Code logic
        if 'Project Number' in rest.columns:
            rest['Account Code'] = rest.apply(
                lambda row: 4301 if str(row['Project Number']).endswith('650') 
                else (4300 if pd.notna(row['Project Number']) and str(row['Project Number']).strip() != '' else 6300),
                axis=1
            )
            rest['Project No'] = rest['Project Number']
            
            # Format Project Number with dash - handle NaN values
            rest['Project No'] = rest['Project No'].astype(str)
            # Replace 'nan' strings with empty string
            rest.loc[rest['Project No'] == 'nan', 'Project No'] = ''
            mask = rest['Project No'].str.len() == 16
            mask1 = rest['Project No'].str.len() == 17
            mask2 = rest['Project No'] == 'NO'
            rest.loc[mask, 'Project No'] = rest.loc[mask, 'Project No'].str[:12] + '-0' + rest.loc[mask, 'Project No'].str[12:]
            rest.loc[mask1, 'Project No'] = rest.loc[mask1, 'Project No'].str[:14] + '-' + rest.loc[mask1, 'Project No'].str[14:]
            rest.loc[mask2, 'Project No'] = ''
        else:
            rest['Account Code'] = 6300
            rest['Project No'] = ''
        
        rest['Cost Centre'] = rest['Cost Center']
        rest['Activity Code'] = 110
        
        # Tax codes based on VAT rate (only if column exists)
        if 'VAT Rate (%)' in rest.columns:
            tax_codes = []
            for vat_rate in rest['VAT Rate (%)']:
                if vat_rate == 19:
                    tax_codes.append('SP')
                elif vat_rate == 7:
                    tax_codes.append('PL')
                else:
                    tax_codes.append('ZE')
            rest['Tax Code'] = tax_codes
        else:
            rest['Tax Code'] = 'ZE'
        
        # Build description with optional columns
        description_parts = [rest['Invoice Number'].astype(str)]
        if 'Position Number' in rest.columns:
            description_parts.append(' POS ' + rest['Position Number'].astype(str))
        if 'Name' in rest.columns:
            description_parts.append(' ' + rest['Name'].astype(str))
        if 'Service Description 1' in rest.columns:
            description_parts.append(' ' + rest['Service Description 1'].astype(str))
        
        rest['Description'] = description_parts[0]
        for part in description_parts[1:]:
            rest['Description'] = rest['Description'] + part
        rest['Cur_amount'] = rest['Amount']
        rest.rename(columns={'Invoice Number': 'Invoice No'}, inplace=True)
        
        # Add missing columns to total and tax dataframes
        total['Cost Centre'] = pd.NA
        total['Project No'] = pd.NA
        total['Activity Code'] = pd.NA
        total['Tax Code'] = pd.NA
        total['Cur_amount'] = total['Amount']
        total.rename(columns={'Invoice Number': 'Invoice No'}, inplace=True)
        
        tax['Cost Centre'] = pd.NA
        tax['Project No'] = pd.NA
        tax['Activity Code'] = pd.NA
        tax['Tax Code'] = pd.NA
        tax['Cur_amount'] = tax['Amount']
        tax.rename(columns={'Invoice Number': 'Invoice No'}, inplace=True)
        
        # Combine all
        combined = pd.concat([total, tax, rest], ignore_index=True)
        
        # Reorder columns to match Airplus format
        desired_columns = ['Account Code', 'Cost Centre', 'Project No', 'Activity Code', 
                           'Invoice No', 'Invoice Date', 'Cur_amount', 'Amount', 
                           'Description', 'Tax Code', 'Posting type', 'SUPID']
        combined = combined[desired_columns]
        return combined


    def load_ehotel_csv(uploaded_file):
        """Load Ehotel CSV with German to English translation"""
        df = pd.read_csv(uploaded_file, sep=';', encoding='ISO-8859-15')
        
        german_to_english = {
            'Rechnungsnummer': 'Invoice Number',
            'Rechnungsdatum': 'Invoice Date',
            'Bruttobetrag': 'Gross Amount',
            'Kostenstelle': 'Cost Center',
            'Projektnummer': 'Project Number',
            'Netto(AbrW)': 'Net (Billing Currency)',
            'MwSt(AbrW)': 'VAT (Billing Currency)',
            'Brutto(AW)': 'Gross (Billing Currency)',
            'MwSt-Satz(%)': 'VAT Rate (%)',
            'Positionsnummer': 'Position Number',
            'Name': 'Name',
            'Leistungsbeschreibung1': 'Service Description 1',
        }
        
        df = df.rename(columns=german_to_english)
        return df


    st.sidebar.header("📁 Upload Files")
    uploaded_files = st.sidebar.file_uploader(
        "Choose Rechnung CSV files",
        type=['csv'],
        accept_multiple_files=True,
        key="ehotel_rech"
    )

    if uploaded_files:
        st.sidebar.success(f"✅ {len(uploaded_files)} file(s) uploaded")
        
        col1, col2 = st.columns([3, 1])
        with col2:
            process_button = st.button("🚀 Process Files", width='stretch', key="rech_process")
        
        if process_button:
            try:
                with st.spinner('Loading and processing files...'):
                    all_dfs = [load_ehotel_csv(f) for f in uploaded_files]
                    combined_df = pd.concat(all_dfs, ignore_index=True) if len(all_dfs) > 1 else all_dfs[0]
                    
                    st.info(f"📊 Loaded {len(combined_df)} rows from {len(uploaded_files)} file(s)")
                    
                    result_df = process_ehotel_rechuntype(combined_df)
                    
                    st.success(f"✅ Processing complete! Generated {len(result_df)} GL entries")
                    st.subheader("📋 Preview")
                    st.dataframe(result_df.head(10), width='stretch')
                    
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        result_df.to_excel(writer, sheet_name='Journal', index=False)
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Excel File",
                        data=excel_buffer.getvalue(),
                        file_name="Ehotel_Rechnung_Journal.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.subheader("📈 Statistics")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("AP Entries", len(result_df[result_df['Posting type'] == 'AP']))
                    with col2:
                        st.metric("Tax Entries", len(result_df[result_df['Account Code'] == 1910]))
                    with col3:
                        st.metric("Detail Entries", len(result_df[(result_df['Posting type'] == 'GL') & (result_df['Account Code'] != 1910)]))
                    with col4:
                        st.metric("Total Entries", len(result_df))
                    
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
    else:
        st.info("👈 Upload Rechnung CSV files to begin")

# ============================================================================
# EHOTEL AUSLAGEN PAGE
# ============================================================================

elif page == "🏨 Ehotel - Auslagen":
    st.title("📊 Ehotel Auslagen Journal Generator")
    st.markdown("Upload Auslagen CSV files and generate GL accounting entries")

    def process_ehotel_ausclen(df):
        """Process Ehotel Auslagen invoice data"""
        
        df = df.copy()
        
        # Convert numeric columns with comma decimals
        numeric_cols = ['VAT (Billing Currency)', 'Net (Billing Currency)', 'Gross Amount', 'Gross (Billing Currency)', 'VAT Rate (%)']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '.').apply(pd.to_numeric, errors='coerce')
        
        # Convert date columns
        date_cols = ['Invoice Date', 'Receipt Date']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], format='%d.%m.%Y', errors='coerce')
        
        # ========== TOTAL (AP - 1300) ==========
        total = df.groupby('Invoice Number', as_index=False).agg({
            'Receipt Date': 'first',
            'Gross (Billing Currency)': 'sum',
            'Receipt': 'first',
            'Payment Receipts': 'first',
            'Service Code': 'first',
            **({'Position Number': 'first'} if 'Position Number' in df.columns else {}),
            **({'Name': 'first'} if 'Name' in df.columns else {}),
            **({'Service Description 1': 'first'} if 'Service Description 1' in df.columns else {}),
        })
        total['Posting type'] = 'AP'
        total['Account Code'] = 1300
        total['Amount'] = total['Gross (Billing Currency)'] * -1
        total['Cur_amount'] = total['Amount']
        total['Invoice Number'] = total['Receipt'].astype(str) + ' ' + total['Payment Receipts'].astype(str) + ' ' + total['Service Code'].astype(str)
        total['Description'] = total['Invoice Number'].astype(str) + ' POS ' + total['Position Number'].astype(str) + ' ' + total['Name'].astype(str) + ' ' + total['Service Description 1'].astype(str)
        total['Cost Centre'] = pd.NA
        total['Project No'] = pd.NA
        total['Activity Code'] = pd.NA
        total['Tax Code'] = pd.NA
        total['SUPID'] = 41253
        
        # ========== TAX (GL - 1910) ==========
        tax = df.groupby('Invoice Number', as_index=False).agg({
            'Receipt Date': 'first',
            'VAT (Billing Currency)': 'sum',
            'Receipt': 'first',
            'Payment Receipts': 'first',
            'Service Code': 'first',
            **({'Position Number': 'first'} if 'Position Number' in df.columns else {}),
            **({'Name': 'first'} if 'Name' in df.columns else {}),
            **({'Service Description 1': 'first'} if 'Service Description 1' in df.columns else {}),
        })
        tax['Posting type'] = 'GL'
        tax['Account Code'] = 1910
        tax['Amount'] = tax['VAT (Billing Currency)']
        tax['Cur_amount'] = tax['Amount']
        tax['Invoice Number'] = tax['Receipt'].astype(str) + ' ' + tax['Payment Receipts'].astype(str) + ' ' + tax['Service Code'].astype(str)
        tax['Description'] = tax['Invoice Number'].astype(str) + ' POS ' + tax['Position Number'].astype(str) + ' ' + tax['Name'].astype(str) + ' ' + tax['Service Description 1'].astype(str)
        tax['Cost Centre'] = pd.NA
        tax['Project No'] = pd.NA
        tax['Activity Code'] = pd.NA
        tax['Tax Code'] = pd.NA
        tax['SUPID'] = 41253
        # ========== DETAILS (GL - 4300/4301) ==========
        cols_to_use = ['Invoice Number', 'Cost Center', 'Net (Billing Currency)', 'Receipt Date', 'Receipt', 'Payment Receipts', 'Service Code'] + \
                      (['Position Number'] if 'Position Number' in df.columns else []) + \
                      (['Name'] if 'Name' in df.columns else []) + \
                      (['Service Description 1'] if 'Service Description 1' in df.columns else []) + \
                      (['VAT Rate (%)'] if 'VAT Rate (%)' in df.columns else []) + \
                      (['Project Number'] if 'Project Number' in df.columns else [])
        rest = df[cols_to_use].copy()
        rest['Posting type'] = 'GL'
        rest['Amount'] = rest['Net (Billing Currency)']
        rest['Cur_amount'] = rest['Amount']
        rest['SUPID'] = 41253
        
        # Account Code logic
        if 'Project Number' in rest.columns:
            rest['Account Code'] = rest.apply(
                lambda row: 4301 if str(row['Project Number']).endswith('650') 
                else (4300 if pd.notna(row['Project Number']) and str(row['Project Number']).strip() != '' else 6300),
                axis=1
            )
            rest['Project No'] = rest['Project Number']
            
            # Format Project Number with dash - handle NaN values
            rest['Project No'] = rest['Project No'].astype(str)
            # Replace 'nan' strings with empty string
            rest.loc[rest['Project No'] == 'nan', 'Project No'] = ''
            mask = rest['Project No'].str.len() == 16
            mask1 = rest['Project No'].str.len() == 17
            mask2 = rest['Project No'] == 'NO'
            rest.loc[mask, 'Project No'] = rest.loc[mask, 'Project No'].str[:12] + '-0' + rest.loc[mask, 'Project No'].str[12:]
            rest.loc[mask1, 'Project No'] = rest.loc[mask1, 'Project No'].str[:14] + '-' + rest.loc[mask1, 'Project No'].str[14:]
            rest.loc[mask2, 'Project No'] = ''
        else:
            rest['Account Code'] = 6300
            rest['Project No'] = ''
        
        rest['Invoice Number'] = rest['Receipt'].astype(str) + ' ' + rest['Payment Receipts'].astype(str) + ' ' + rest['Service Code'].astype(str)
        
        # Build description with optional columns
        description_parts = [rest['Invoice Number'].astype(str)]
        if 'Position Number' in rest.columns:
            description_parts.append(' POS ' + rest['Position Number'].astype(str))
        if 'Name' in rest.columns:
            description_parts.append(' ' + rest['Name'].astype(str))
        if 'Service Description 1' in rest.columns:
            description_parts.append(' ' + rest['Service Description 1'].astype(str))
        
        rest['Description'] = description_parts[0]
        for part in description_parts[1:]:
            rest['Description'] = rest['Description'] + part
        rest['Cost Centre'] = rest['Cost Center']
        rest['Activity Code'] = 110
        
        # Tax codes based on VAT rate (only if column exists)
        if 'VAT Rate (%)' in rest.columns:
            tax_codes = []
            for vat_rate in rest['VAT Rate (%)']:
                if vat_rate == 19:
                    tax_codes.append('SP')
                elif vat_rate == 7:
                    tax_codes.append('PL')
                else:
                    tax_codes.append('ZE')
            rest['Tax Code'] = tax_codes
        else:
            rest['Tax Code'] = 'ZE'
        rest.rename(columns={'Invoice Number': 'Invoice No'}, inplace=True)
        
        # Rename all dataframes for consistency
        total.rename(columns={'Invoice Number': 'Invoice No'}, inplace=True)
        tax.rename(columns={'Invoice Number': 'Invoice No'}, inplace=True)
        
        # Combine all
        combined = pd.concat([total, tax, rest], ignore_index=True)
        
        # Reorder columns to match Airplus format
        desired_columns = ['Account Code', 'Cost Centre', 'Project No', 'Activity Code', 
                           'Invoice No', 'Receipt Date', 'Cur_amount', 'Amount', 
                           'Description', 'Tax Code', 'Posting type', 'SUPID']
        combined = combined[desired_columns]
        return combined


    def load_ehotel_csv(uploaded_file):
        """Load Ehotel CSV with German to English translation"""
        df = pd.read_csv(uploaded_file, sep=';', encoding='ISO-8859-15')
        
        german_to_english = {
            'Rechnungsnummer': 'Invoice Number',
            'Rechnungsdatum': 'Invoice Date',
            'Bruttobetrag': 'Gross Amount',
            'Kostenstelle': 'Cost Center',
            'Projektnummer': 'Project Number',
            'Netto(AbrW)': 'Net (Billing Currency)',
            'MwSt(AbrW)': 'VAT (Billing Currency)',
            'Brutto(AW)': 'Gross (Billing Currency)',
            'Beleg': 'Receipt',
            'Belegdatum': 'Receipt Date',
            'Zahlbelege': 'Payment Receipts',
            'Leistungscode': 'Service Code',
            'MwSt-Satz(%)': 'VAT Rate (%)',
            'Positionsnummer': 'Position Number',
            'Name': 'Name',
            'Leistungsbeschreibung1': 'Service Description 1',
        }
        
        df = df.rename(columns=german_to_english)
        return df


    st.sidebar.header("📁 Upload Files")
    uploaded_files = st.sidebar.file_uploader(
        "Choose Auslagen CSV files",
        type=['csv'],
        accept_multiple_files=True,
        key="ehotel_aus"
    )

    if uploaded_files:
        st.sidebar.success(f"✅ {len(uploaded_files)} file(s) uploaded")
        
        col1, col2 = st.columns([3, 1])
        with col2:
            process_button = st.button("🚀 Process Files", width='stretch', key="aus_process")
        
        if process_button:
            try:
                with st.spinner('Loading and processing files...'):
                    all_dfs = [load_ehotel_csv(f) for f in uploaded_files]
                    combined_df = pd.concat(all_dfs, ignore_index=True) if len(all_dfs) > 1 else all_dfs[0]
                    
                    st.info(f"📊 Loaded {len(combined_df)} rows from {len(uploaded_files)} file(s)")
                    
                    result_df = process_ehotel_ausclen(combined_df)
                    
                    st.success(f"✅ Processing complete! Generated {len(result_df)} GL entries")
                    st.subheader("📋 Preview")
                    st.dataframe(result_df.head(10), width='stretch')
                    
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        result_df.to_excel(writer, sheet_name='Journal', index=False)
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Excel File",
                        data=excel_buffer.getvalue(),
                        file_name="Ehotel_Auslagen_Journal.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.subheader("📈 Statistics")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("AP Entries", len(result_df[result_df['Posting type'] == 'AP']))
                    with col2:
                        st.metric("Tax Entries", len(result_df[result_df['Account Code'] == 1910]))
                    with col3:
                        st.metric("Detail Entries", len(result_df[(result_df['Posting type'] == 'GL') & (result_df['Account Code'] != 1910)]))
                    with col4:
                        st.metric("Total Entries", len(result_df))
                    
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
    else:
        st.info("👈 Upload Ausclen CSV files to begin")

