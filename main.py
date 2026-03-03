import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")

st.title("Reconciliation Processing App")

def date_convert(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, pd.Timestamp):
        return x.normalize()
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        if x > 20000:
            return pd.to_datetime(x, origin="1899-12-30", unit="D").normalize()
    dt = pd.to_datetime(x, errors="coerce")
    if pd.isna(dt):
        return pd.NaT
    return dt.normalize()

uploaded_file1 = st.file_uploader("Upload Sales Reversal File (Excel)", type=["xlsx"])
st.caption("Sheet: Sales Reversal | Header Row Number Required | Columns: B:K,M,O,P,R")

uploaded_file2 = st.file_uploader("Upload Reversal Report File (Excel)", type=["xlsx"])
st.caption("Columns: Orderlocn, Hubname, Cust no, CUST NAME, invoiceno, Invoice Date, Cr Invoice Total, orderno, period from, period to, ainvoiceno, a Invoice Dt, New Invoice Total, Rev Remarks")

uploaded_file3 = st.file_uploader("Upload Rebilled Invoice File (Excel)", type=["xlsx"])
st.caption("Columns: SoLocn, Hub, CustNo, Customer Name, InvNo, Old invoice Date,    Amount  , Rebilled Invoice, Date,   Amount , Sub Category")

uploaded_file4 = st.file_uploader("Upload Ageing CSV File (CSV)", type=["csv"])
st.caption("Header Row Number Required")

uploaded_file5 = st.file_uploader("Upload Mapping File (Excel)", type=["xlsx"])
st.caption("Sheets: Mapping (Old So Code, Branch) | SA List (Customer Code, SA)")

header_row_sales = st.number_input("Sales Reversal Header Row", min_value=0, value=1)
sheet_name_sales = st.text_input("Sales Reversal Sheet Name", value="Sales Reversal")

header_row_ageing = st.number_input("Ageing CSV Header Row", min_value=0, value=2)

run = st.button("Run")

if run:
    if uploaded_file1 and uploaded_file2 and uploaded_file3 and uploaded_file4 and uploaded_file5:

        st.write("Reading Sales Reversal...")
        df = pd.read_excel(uploaded_file1, sheet_name=sheet_name_sales, header=header_row_sales, usecols="B:K,M,O,P,R")
        df = df.rename(columns={df.columns[10]: "New Invoice number"})

        for col in ['Old inv Dt','New invoice date','Pr from','Pr to']:
            if col in df.columns:
                df[col] = df[col].apply(date_convert)

        st.write("Reading Reversal Report...")
        df2 = pd.read_excel(uploaded_file2, usecols=['Orderlocn','Hubname','Cust no','CUST NAME','invoiceno',
                                                    'Invoice Date','Cr Invoice Total','orderno','period from',
                                                    'period to','ainvoiceno','a Invoice Dt',
                                                    'New Invoice Total','Rev Remarks'])

        for col in ['Invoice Date','a Invoice Dt','period from','period to']:
            df2[col] = df2[col].apply(date_convert)

        df2.columns = df.columns
        df = pd.concat([df, df2], ignore_index=True)

        st.write("Reading Rebilled Invoice...")
        df3 = pd.read_excel(uploaded_file3, usecols=['SoLocn','Hub','CustNo','Customer Name','InvNo',
                                                    'Old invoice Date','   Amount  ','Rebilled Invoice',
                                                    'Date','  Amount ','Sub Category'])

        df3['Old invoice Date'] = df3['Old invoice Date'].apply(date_convert)
        df3['Date'] = df3['Date'].apply(date_convert)

        df3.insert(7,"c",None)
        df3.insert(7,"b",None)
        df3.insert(7,"a",None)

        df3.columns = df.columns
        df = pd.concat([df, df3], ignore_index=True)

        df4 = df.sort_values('Old inv Dt')
        df4['a'] = df4['New Invoice number'].str.strip()

        first_rows = df4.drop_duplicates(subset='a', keep='first')
        df4 = df4.drop(columns='a')

        df['Status'] = 'Y'
        df.loc[first_rows.index, 'Status'] = 'N'

        st.write("Reading Ageing CSV...")
        df5 = pd.read_csv(uploaded_file4, header=header_row_ageing, encoding='latin1')

        for col in ['location_no','ORD_LOCN','INVOICE_NO','hub','Pay_Term_Desc']:
            if col in df5.columns:
                df5[col] = df5[col].astype(str).str.strip()

        df5['ORD_LOCN'] = df5['ORD_LOCN'].str.upper()

        def removingd(r):
            if r=='15TO30': return '30D'
            elif r=='30TO45': return '45D'
            elif r=='45TO60': return '60D'
            elif r=='60TO90': return '90D'
            elif r=='LESS15': return '15D'
            elif r=='ABOVE90': return '90D'
            elif r=='ADV': return '0D'
            else: return r

        df5['Pay_Term_Desc'] = df5['Pay_Term_Desc'].astype(str).apply(removingd)
        df5['Pay_Term_Desc'] = df5['Pay_Term_Desc'].str[:-1]
        df5['Pay_Term_Desc'] = df5['Pay_Term_Desc'].replace(['','na'],'0').astype(int)

        st.write("Reading Mapping...")
        df6 = pd.read_excel(uploaded_file5, sheet_name="Mapping", usecols=['Old So Code','Branch'])
        df6['Old So Code'] = df6['Old So Code'].str.strip()
        df6['Branch'] = df6['Branch'].str.strip()

        df5 = pd.merge(df5, df6, left_on='ORD_LOCN', right_on='Old So Code', how="left")
        df5 = df5.drop(columns=['Old So Code'])

        df7 = pd.read_excel(uploaded_file5, sheet_name="SA List", usecols=['Customer Code','SA'])
        df5 = pd.merge(df5, df7, left_on='Cust_no', right_on='Customer Code', how="left")
        df5 = df5.drop(columns=['Customer Code'])
        df5['SA'] = df5['SA'].fillna('NSA').replace('','NSA')
        df5 = df5.rename(columns={'SA':'A/C Type'})

        st.write("Reading Aging Sheet...")
        df8 = pd.read_excel(uploaded_file1, sheet_name="Aging ", usecols=['ORD_LOCN','Cust_no','INVOICE_NO',
                                                                         'Recoverable / Not recoverable for tracker'],
                            header=header_row_sales)

        df8['Key'] = df8['Cust_no'].astype(str)+df8['ORD_LOCN']+df8['INVOICE_NO']
        df5['Key'] = df5['Cust_no'].astype(str)+df5['ORD_LOCN']+df5['INVOICE_NO']

        df5 = pd.merge(df5, df8[['Key','Recoverable / Not recoverable for tracker']],
                       on='Key', how="left")

        df5['Recoverable / Not recoverable for tracker'] = df5['Recoverable / Not recoverable for tracker']\
                                                            .fillna('Recoverable').replace('','Recoverable')

        df5 = df5.sort_values('DOC_DATE')

        st.write("Generating Output...")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Sales Reversal', index=False)
            df5.to_excel(writer, sheet_name='Ageing', index=False)

        st.download_button(
            label="Download Output File (Excel)",
            data=output.getvalue(),
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Please upload all required files.")
