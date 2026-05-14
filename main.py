import streamlit as st
import pandas as pd
from io import BytesIO
import time

st.set_page_config(layout="wide")

st.title("Reconciliation Processing App")

st.subheader("Upload Files")

uploaded_file1 = st.file_uploader(
    "Upload Sales Reversal File (Excel)",
    type=["xlsx"]
)
st.caption(
    "Required Sheet: Sales Reversal | "
    "Header can be located anywhere within first 10 rows | "
    "Required Columns: B:K,M,O,P,R | "
    "Also requires Aging sheet for tracker mapping"
)

uploaded_file2 = st.file_uploader(
    "Upload Reversal Report File (Excel)",
    type=["xlsx"]
)
st.caption(
    "Header can be located anywhere within first 10 rows | "
    "Required Columns: Orderlocn, Hubname, Cust no, CUST NAME, invoiceno, "
    "Invoice Date, Cr Invoice Total, orderno, period from, period to, "
    "ainvoiceno, a Invoice Dt, New Invoice Total, Rev Remarks"
)

uploaded_file3 = st.file_uploader(
    "Upload Rebilled Invoice File (Excel)",
    type=["xlsx"]
)
st.caption(
    "Header can be located anywhere within first 10 rows | "
    "Required Columns: SoLocn, Hub, CustNo, Customer Name, InvNo, "
    "Old invoice Date, Amount, Rebilled Invoice, Date, Amount, Sub Category"
)

uploaded_file4 = st.file_uploader(
    "Upload Ageing CSV File (CSV)",
    type=["csv"]
)
st.caption(
    "Header can be located anywhere within first 10 rows"
)

uploaded_file5 = st.file_uploader(
    "Upload Mapping File (Excel)",
    type=["xlsx"]
)
st.caption(
    "Required Sheets: Mapping, SA List | "
    "Mapping Sheet Columns: Old So Code, Branch | "
    "SA List Columns: Customer Code, SA"
)

run = st.button("Run")

log_container = st.container()

def date_convert(x):
    if pd.isna(x):
        return pd.NaT

    if isinstance(x, pd.Timestamp):
        return x.normalize()

    if isinstance(x, (int, float)) and not isinstance(x, bool):
        if x > 20000:
            return pd.to_datetime(
                x,
                origin="1899-12-30",
                unit="D"
            ).normalize()

    dt = pd.to_datetime(x, errors="coerce")

    if pd.isna(dt):
        return pd.NaT

    return dt.normalize()

def find_header_row_excel(file, sheet_name, required_columns, max_rows=10):
    try:
        preview = pd.read_excel(
            file,
            sheet_name=sheet_name,
            header=None,
            nrows=max_rows
        )

        for i in range(max_rows):
            row_values = preview.iloc[i].astype(str).str.strip().tolist()

            if all(col in row_values for col in required_columns):
                return i

        return None

    except Exception as e:
        st.error(f"Error finding header row in Excel file: {e}")
        st.stop()

def find_header_row_csv(file, required_columns, max_rows=10):
    try:
        preview = pd.read_csv(
            file,
            header=None,
            encoding="latin1",
            nrows=max_rows
        )

        for i in range(max_rows):
            row_values = preview.iloc[i].astype(str).str.strip().tolist()

            if all(col in row_values for col in required_columns):
                return i

        return None

    except Exception as e:
        st.error(f"Error finding header row in CSV file: {e}")
        st.stop()

if run:

    with log_container:

        status_text = st.empty()
        progress_bar = st.progress(0)

        if not all([
            uploaded_file1,
            uploaded_file2,
            uploaded_file3,
            uploaded_file4,
            uploaded_file5
        ]):
            status_text.warning("Please upload all required files.")
            st.stop()

        status_text.info("Initializing reconciliation process...")
        progress_bar.progress(5)
        time.sleep(0.2)

        # ---------------- SALES REVERSAL ---------------- #

        sales_required_columns = [
            'Old inv Dt',
            'New invoice date',
            'Pr from',
            'Pr to'
        ]

        sheet_name_sales = "Sales Reversal"

        try:
            xl_check = pd.ExcelFile(uploaded_file1)

            if sheet_name_sales not in xl_check.sheet_names:
                st.error(f"Required sheet not found: {sheet_name_sales}")
                st.stop()

        except Exception as e:
            st.error(f"Error validating Sales Reversal file: {e}")
            st.stop()

        status_text.info("Searching header row in Sales Reversal sheet...")
        progress_bar.progress(10)
        time.sleep(0.2)

        header_row_sales = find_header_row_excel(
            uploaded_file1,
            sheet_name_sales,
            sales_required_columns
        )

        if header_row_sales is None:
            st.error(
                "Could not find required headers within first 10 rows "
                "in Sales Reversal sheet."
            )
            st.stop()

        try:
            status_text.info("Reading Sales Reversal data...")
            progress_bar.progress(15)

            df = pd.read_excel(
                uploaded_file1,
                sheet_name=sheet_name_sales,
                header=header_row_sales,
                usecols="B:K,M,O,P,R"
            )

        except Exception as e:
            st.error(f"Error reading Sales Reversal file: {e}")
            st.stop()

        try:
            df = df.rename(columns={df.columns[10]: "New Invoice number"})
        except Exception as e:
            st.error(f"Error renaming Sales Reversal columns: {e}")
            st.stop()

        try:
            status_text.info("Converting Sales Reversal dates...")
            progress_bar.progress(20)

            for col in ['Old inv Dt', 'New invoice date', 'Pr from', 'Pr to']:
                if col in df.columns:
                    df[col] = df[col].apply(date_convert)

        except Exception as e:
            st.error(f"Error during Sales Reversal date conversion: {e}")
            st.stop()

        time.sleep(0.2)

        # ---------------- REVERSAL REPORT ---------------- #

        reversal_required_columns = [
            'Orderlocn',
            'Hubname',
            'Cust no',
            'CUST NAME',
            'invoiceno'
        ]

        status_text.info("Searching header row in Reversal Report...")
        progress_bar.progress(25)
        time.sleep(0.2)

        header_row_reversal = find_header_row_excel(
            uploaded_file2,
            0,
            reversal_required_columns
        )

        if header_row_reversal is None:
            st.error(
                "Could not find required headers within first 10 rows "
                "in Reversal Report file."
            )
            st.stop()

        try:
            status_text.info("Reading Reversal Report data...")
            progress_bar.progress(30)

            df2 = pd.read_excel(
                uploaded_file2,
                header=header_row_reversal,
                usecols=[
                    'Orderlocn',
                    'Hubname',
                    'Cust no',
                    'CUST NAME',
                    'invoiceno',
                    'Invoice Date',
                    'Cr Invoice Total',
                    'orderno',
                    'period from',
                    'period to',
                    'ainvoiceno',
                    'a Invoice Dt',
                    'New Invoice Total',
                    'Rev Remarks'
                ]
            )

        except Exception as e:
            st.error(f"Error reading Reversal Report file: {e}")
            st.stop()

        try:
            status_text.info("Processing Reversal Report dates...")
            progress_bar.progress(35)

            for col in [
                'Invoice Date',
                'a Invoice Dt',
                'period from',
                'period to'
            ]:
                df2[col] = df2[col].apply(date_convert)

        except Exception as e:
            st.error(f"Error during Reversal Report date conversion: {e}")
            st.stop()

        try:
            df2.columns = df.columns
            df = pd.concat([df, df2], ignore_index=True)

        except Exception as e:
            st.error(f"Error combining Reversal Report data: {e}")
            st.stop()

        time.sleep(0.2)

        # ---------------- REBILLED INVOICE ---------------- #

        rebilled_required_columns = [
            'SoLocn',
            'Hub',
            'CustNo',
            'Customer Name',
            'InvNo'
        ]

        status_text.info("Searching header row in Rebilled Invoice file...")
        progress_bar.progress(40)
        time.sleep(0.2)

        header_row_rebilled = find_header_row_excel(
            uploaded_file3,
            0,
            rebilled_required_columns
        )

        if header_row_rebilled is None:
            st.error(
                "Could not find required headers within first 10 rows "
                "in Rebilled Invoice file."
            )
            st.stop()

        try:
            status_text.info("Reading Rebilled Invoice data...")
            progress_bar.progress(45)

            df3 = pd.read_excel(
                uploaded_file3,
                header=header_row_rebilled,
                usecols=[
                    'SoLocn',
                    'Hub',
                    'CustNo',
                    'Customer Name',
                    'InvNo',
                    'Old invoice Date',
                    '   Amount  ',
                    'Rebilled Invoice',
                    'Date',
                    '  Amount ',
                    'Sub Category'
                ]
            )

        except Exception as e:
            st.error(f"Error reading Rebilled Invoice file: {e}")
            st.stop()

        try:
            status_text.info("Processing Rebilled Invoice dates...")
            progress_bar.progress(50)

            df3['Old invoice Date'] = df3['Old invoice Date'].apply(date_convert)
            df3['Date'] = df3['Date'].apply(date_convert)

        except Exception as e:
            st.error(f"Error converting Rebilled Invoice dates: {e}")
            st.stop()

        try:
            df3.insert(7, "c", None)
            df3.insert(7, "b", None)
            df3.insert(7, "a", None)

            df3.columns = df.columns

            df = pd.concat([df, df3], ignore_index=True)

        except Exception as e:
            st.error(f"Error preparing Rebilled Invoice data: {e}")
            st.stop()

        # ---------------- STATUS LOGIC ---------------- #

        try:
            status_text.info("Generating status flags...")
            progress_bar.progress(55)

            df4 = df.sort_values('Old inv Dt')

            df4['a'] = df4['New Invoice number'].astype(str).str.strip()

            first_rows = df4.drop_duplicates(
                subset='a',
                keep='first'
            )

            df4 = df4.drop(columns='a')

            df['Status'] = 'Y'

            df.loc[first_rows.index, 'Status'] = 'N'

        except Exception as e:
            st.error(f"Error generating status flags: {e}")
            st.stop()

        time.sleep(0.2)

        # ---------------- AGEING CSV ---------------- #

        ageing_required_columns = [
            'location_no',
            'ORD_LOCN',
            'INVOICE_NO'
        ]

        status_text.info("Searching header row in Ageing CSV...")
        progress_bar.progress(60)
        time.sleep(0.2)

        header_row_ageing = find_header_row_csv(
            uploaded_file4,
            ageing_required_columns
        )

        if header_row_ageing is None:
            st.error(
                "Could not find required headers within first 10 rows "
                "in Ageing CSV file."
            )
            st.stop()

        try:
            status_text.info("Reading Ageing CSV data...")
            progress_bar.progress(65)

            df5 = pd.read_csv(
                uploaded_file4,
                header=header_row_ageing,
                encoding='latin1'
            )

        except Exception as e:
            st.error(f"Error reading Ageing CSV file: {e}")
            st.stop()

        try:
            status_text.info("Cleaning Ageing data...")
            progress_bar.progress(70)

            for col in [
                'location_no',
                'ORD_LOCN',
                'INVOICE_NO',
                'hub',
                'Pay_Term_Desc'
            ]:
                if col in df5.columns:
                    df5[col] = df5[col].astype(str).str.strip()

            df5['ORD_LOCN'] = df5['ORD_LOCN'].str.upper()

        except Exception as e:
            st.error(f"Error cleaning Ageing data: {e}")
            st.stop()

        try:

            def removingd(r):
                if r == '15TO30':
                    return '30D'
                elif r == '30TO45':
                    return '45D'
                elif r == '45TO60':
                    return '60D'
                elif r == '60TO90':
                    return '90D'
                elif r == 'LESS15':
                    return '15D'
                elif r == 'ABOVE90':
                    return '90D'
                elif r == 'ADV':
                    return '0D'
                else:
                    return r

            status_text.info("Transforming payment term buckets...")
            progress_bar.progress(75)

            df5['Pay_Term_Desc'] = (
                df5['Pay_Term_Desc']
                .astype(str)
                .apply(removingd)
            )

            df5['Pay_Term_Desc'] = df5['Pay_Term_Desc'].str[:-1]

            df5['Pay_Term_Desc'] = (
                df5['Pay_Term_Desc']
                .replace(['', 'na'], '0')
                .astype(int)
            )

        except Exception as e:
            st.error(f"Error processing payment terms: {e}")
            st.stop()

        time.sleep(0.2)

        # ---------------- MAPPING ---------------- #

        try:
            mapping_excel = pd.ExcelFile(uploaded_file5)

            required_sheets = ["Mapping", "SA List"]

            for sheet in required_sheets:
                if sheet not in mapping_excel.sheet_names:
                    st.error(f"Required sheet not found: {sheet}")
                    st.stop()

        except Exception as e:
            st.error(f"Error validating Mapping file: {e}")
            st.stop()

        try:
            status_text.info("Reading branch mapping data...")
            progress_bar.progress(80)

            df6 = pd.read_excel(
                uploaded_file5,
                sheet_name="Mapping",
                usecols=['Old So Code', 'Branch']
            )

            df6['Old So Code'] = df6['Old So Code'].astype(str).str.strip()
            df6['Branch'] = df6['Branch'].astype(str).str.strip()

        except Exception as e:
            st.error(f"Error reading Mapping sheet: {e}")
            st.stop()

        try:
            status_text.info("Applying branch mappings...")
            progress_bar.progress(82)

            df5 = pd.merge(
                df5,
                df6,
                left_on='ORD_LOCN',
                right_on='Old So Code',
                how="left"
            )

            df5 = df5.drop(columns=['Old So Code'])

        except Exception as e:
            st.error(f"Error merging branch mapping: {e}")
            st.stop()

        try:
            status_text.info("Reading SA mapping data...")
            progress_bar.progress(85)

            df7 = pd.read_excel(
                uploaded_file5,
                sheet_name="SA List",
                usecols=['Customer Code', 'SA']
            )

        except Exception as e:
            st.error(f"Error reading SA List sheet: {e}")
            st.stop()

        try:
            status_text.info("Applying SA mappings...")
            progress_bar.progress(88)

            df5 = pd.merge(
                df5,
                df7,
                left_on='Cust_no',
                right_on='Customer Code',
                how="left"
            )

            df5 = df5.drop(columns=['Customer Code'])

            df5['SA'] = (
                df5['SA']
                .fillna('NSA')
                .replace('', 'NSA')
            )

            df5 = df5.rename(columns={'SA': 'A/C Type'})

        except Exception as e:
            st.error(f"Error merging SA mapping: {e}")
            st.stop()

        time.sleep(0.2)

        # ---------------- AGING SHEET ---------------- #

        try:
            status_text.info("Reading Aging tracker sheet...")
            progress_bar.progress(90)

            df8 = pd.read_excel(
                uploaded_file1,
                sheet_name="Aging ",
                usecols=[
                    'ORD_LOCN',
                    'Cust_no',
                    'INVOICE_NO',
                    'Recoverable / Not recoverable for tracker'
                ],
                header=header_row_sales
            )

        except Exception as e:
            st.error(f"Error reading Aging tracker sheet: {e}")
            st.stop()

        try:
            status_text.info("Generating reconciliation keys...")
            progress_bar.progress(93)

            df8['Key'] = (
                df8['Cust_no'].astype(str)
                + df8['ORD_LOCN']
                + df8['INVOICE_NO']
            )

            df5['Key'] = (
                df5['Cust_no'].astype(str)
                + df5['ORD_LOCN']
                + df5['INVOICE_NO']
            )

        except Exception as e:
            st.error(f"Error generating reconciliation keys: {e}")
            st.stop()

        try:
            status_text.info("Applying recoverability tracker...")
            progress_bar.progress(95)

            df5 = pd.merge(
                df5,
                df8[['Key', 'Recoverable / Not recoverable for tracker']],
                on='Key',
                how="left"
            )

            df5['Recoverable / Not recoverable for tracker'] = (
                df5['Recoverable / Not recoverable for tracker']
                .fillna('Recoverable')
                .replace('', 'Recoverable')
            )

            df5 = df5.sort_values('DOC_DATE')

        except Exception as e:
            st.error(f"Error applying recoverability tracker: {e}")
            st.stop()

        # ---------------- OUTPUT ---------------- #

        try:
            status_text.info("Generating final output workbook...")
            progress_bar.progress(98)

            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine='openpyxl'
            ) as writer:

                df.to_excel(
                    writer,
                    sheet_name='Sales Reversal',
                    index=False
                )

                df5.to_excel(
                    writer,
                    sheet_name='Ageing',
                    index=False
                )

        except Exception as e:
            st.error(f"Error generating output workbook: {e}")
            st.stop()

        progress_bar.progress(100)

        status_text.success(
            "Reconciliation completed successfully. "
            "Output file is ready for download."
        )

        st.download_button(
            label="Download Output File (Excel)",
            data=output.getvalue(),
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---------------- DOCUMENTATION ---------------- #

with st.expander("What This Tool Does"):

    st.write("""
    This reconciliation tool combines Sales Reversal, Reversal Report,
    Rebilled Invoice, Ageing, and Mapping data into a consolidated output.

    It validates invoice movement, rebilling linkage, payment terms,
    branch mapping, SA classification, and recoverability tracking.

    Final output provides a processed reconciliation-ready dataset
    for operational and finance review.
    """)

with st.expander("How to Use"):

    st.write("""
    1. Upload all required files
    2. Click Run
    3. Wait for processing completion
    4. Download the generated output file
    """)

with st.expander("Output Details"):

    st.write("""
    Output File Contains:

    • Sales Reversal Sheet
        - Combined reversal records
        - Rebilled invoice records
        - Status identification

    • Ageing Sheet
        - Cleaned ageing data
        - Branch mapping
        - SA classification
        - Recoverability tracker
        - Payment term normalization

    The report helps identify invoice reversals,
    rebilling linkage, recoverability status,
    and ageing alignment.
    """)

with st.expander("Financial Logic"):

    st.write("""
    • Invoice records from multiple sources are consolidated
      into a single reconciliation structure.

    • Earliest occurrence of each New Invoice Number
      is marked separately using Status logic.

    • Payment term buckets are normalized into standard day values.

    • Branch mappings are applied using operational location codes.

    • Customer accounts are classified into SA / NSA categories.

    • Recoverability tracking identifies whether ageing balances
      are marked recoverable or not recoverable.
    """)
