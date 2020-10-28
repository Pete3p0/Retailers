# Import libraries
import streamlit as st
import numpy as np
import pandas as pd
import base64
from io import BytesIO
import datetime as dt
import locale
locale.setlocale( locale.LC_ALL, 'en_ZA.UTF-8' )

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
    df.to_excel(writer, sheet_name='Sheet1',index=False)
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def get_table_download_link(df):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(df)
    b64 = base64.b64encode(val)
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="Sales.xlsx">Download Excel file</a>' # decode b'abc' => abc

st.title('Retailer Sales Reports')

Date_End = st.date_input("Week ending: ")
Date_Start = Date_End - dt.timedelta(days=7)

if Date_End.day < 10:
    Day = '0'+str(Date_End.day)
else:
    Day = str(Date_End.day)
Month = Date_End.month
Year = str(Date_End.year)
Short_Date_Dict = {1:'Jan', 2:'Feb', 3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}

option = st.selectbox(
    'Please select a retailer:',
    ('Please select','Bradlows/Russels','Checkers', 'Musica','Takealot','TFG'))
st.write('You selected:', option)

st.write("")
st.markdown("Please ensure data is in the **_first sheet_** of your Excel Workbook")

map_file = st.file_uploader('Retailer Map', type='xlsx')
if map_file:
    df_map = pd.read_excel(map_file)

data_file = st.file_uploader('Weekly Sales Data',type='xlsx')
if data_file:
    df_data = pd.read_excel(data_file)

# Bradlows/Russels
if option == 'Bradlows/Russels':
    try:
        # Get retailers map
        df_br_retailers_map = df_map
        df_br_retailers_map = df_br_retailers_map.rename(columns={'Article number':'SKU No. B&R'})
        df_br_retailers_map = df_br_retailers_map[['SKU No. B&R','Product Code','RSP']]

        # Get retailer data
        df_br_data = df_data
        df_br_data.columns = df_br_data.iloc[1]
        df_br_data = df_br_data.iloc[2:]

        # Fill sales qty
        df_br_data['Sales Qty*'].fillna(0,inplace=True)

        # Drop result rows
        df_br_data.drop(df_br_data[df_br_data['Article'] == 'Result'].index, inplace = True) 
        df_br_data.drop(df_br_data[df_br_data['Site'] == 'Result'].index, inplace = True) 
        df_br_data.drop(df_br_data[df_br_data['Cluster'] == 'Overall Result'].index, inplace = True) 

        # Get SKU No. column
        df_br_data['SKU No. B&R'] = df_br_data['Article'].astype(float)

        # Site columns
        df_br_data['Store Name'] = df_br_data['Site'] + ' - ' + df_br_data['Site Name'] 

        # Consolidate
        df_br_data_new = df_br_data[['Cluster','SKU No. B&R','Description','Store Name','Sales Qty*','Valuated Stock Qty(Total)']]

        # Merge with retailer map
        df_br_data_merged = df_br_data_new.merge(df_br_retailers_map, how='left', on='SKU No. B&R',indicator=True)

        # Find missing data
        missing_model_br = df_br_data_merged['Product Code'].isnull()
        df_br_missing_model = df_br_data_merged[missing_model_br]
        df_missing = df_br_missing_model[['SKU No. B&R','Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)
        st.write(" ")

        missing_rsp_br = df_br_data_merged['RSP'].isnull()
        df_br_missing_rsp = df_br_data_merged[missing_rsp_br]
        df_missing_2 = df_br_missing_rsp[['SKU No. B&R','Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)
        
    except:
        st.write('File not selected yet')

    try:
        # Set date columns
        df_br_data_merged['Start Date'] = Date_Start

        # Total amount column
        df_br_data_merged['Total Amt'] = df_br_data_merged['Sales Qty*'] * df_br_data_merged['RSP']

        # Tidy columns
        df_br_data_merged['Forecast Group'] = 'Bradlows/Russels'
        df_br_data_merged['Store Name']= df_br_data_merged['Store Name'].str.title() 

        # Rename columns
        df_br_data_merged = df_br_data_merged.rename(columns={'Sales Qty*': 'Sales Qty'})
        df_br_data_merged = df_br_data_merged.rename(columns={'SKU No. B&R': 'SKU No.'})
        df_br_data_merged = df_br_data_merged.rename(columns={'Valuated Stock Qty(Total)': 'SOH Qty'})

        # Don't change these headings. Rather change the ones above
        final_df_br = df_br_data_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]

        # Show final df
        total = final_df_br['Total Amt'].sum()
        st.write('The total sales for the week are: ',locale.currency( total, grouping=True))
        final_df_br

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_br), unsafe_allow_html=True)

    except:
        st.write('Check data')

# Checkers

elif option == 'Checkers':
    try:
        # Get retailers data
        df_checkers_retailers_map = df_map

        # Get retailer data
        df_checkers_data = df_data
        df_checkers_data.columns = df_checkers_data.iloc[2]
        df_checkers_data = df_checkers_data.iloc[3:]

        # Rename columns
        df_checkers_data = df_checkers_data.rename(columns={'Item Code': 'Article'})
        
        # Merge with Sony Range
        df_checkers_merged = df_checkers_data.merge(df_checkers_retailers_map, how='left', on='Article')
        
        # Find missing data
        missing_model_checkers = df_checkers_merged['SMD Code'].isnull()
        df_checkers_missing_model = df_checkers_merged[missing_model_checkers]
        df_missing = df_checkers_missing_model[['Article','Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)
        st.write(" ")

        missing_rsp_checkers = df_checkers_merged['RSP'].isnull()
        df_checkers_missing_rsp = df_checkers_merged[missing_rsp_checkers]
        df_missing_2 = df_checkers_missing_rsp[['Article','Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.write('File not selected yet')

    try:
        # Add columns for dates
        df_checkers_merged['Start Date'] = Date_Start

        # Add Total Amount column
        Units_Sold = 'Units :'+ Day +' '+ Short_Date_Dict[Month] + ' ' + Year
        df_checkers_merged['Total Amt'] = df_checkers_merged[Units_Sold] * df_checkers_merged['RSP']

        # Add column for retailer and SOH
        df_checkers_merged['Forecast Group'] = 'Checkers'
        df_checkers_merged['SOH Qty'] = 0

        # Rename columns
        df_checkers_merged = df_checkers_merged.rename(columns={'Article': 'SKU No.'})
        df_checkers_merged = df_checkers_merged.rename(columns={Units_Sold: 'Sales Qty'})
        df_checkers_merged = df_checkers_merged.rename(columns={'SMD Code': 'Product Code'})
        df_checkers_merged = df_checkers_merged.rename(columns={'Branch': 'Store Name'})

        # Final df. Don't change these headings. Rather change the ones above
        final_df_checkers_sales = df_checkers_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]

        # Show final df
        total = final_df_checkers_sales['Total Amt'].sum()
        st.write('The total sales for the week are: ',locale.currency( total, grouping=True))
        final_df_checkers_sales

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_checkers_sales), unsafe_allow_html=True)
    except:
        st.write('Check data')


# Musica
elif option == 'Musica':
    try:
        # Get retailers map
        df_musica_retailers_map = df_map
        df_retailers_map_musica_final = df_musica_retailers_map[['Musica Code','SMD code','RSP','SMD Desc', 'Grouping']]

        # Get retailer data
        df_musica_data = df_data
        df_musica_data = df_musica_data.rename(columns={'SKU No.': 'Musica Code'})
        df_musica_data = df_musica_data.rename(columns={'Sales.Qty': 'Sales Qty'})  
        #Merge with retailer map
        df_musica_merged = df_musica_data.merge(df_retailers_map_musica_final, how='left', on='Musica Code')    
        # Find missing data
        missing_model = df_musica_merged['SMD code'].isnull()
        df_musica_missing_model = df_musica_merged[missing_model]
        df_missing = df_musica_missing_model[['Musica Code','Title Desc']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp = df_musica_merged['RSP'].isnull()
        df_musica_missing_rsp = df_musica_merged[missing_rsp]
        df_missing_2 = df_musica_missing_rsp[['Musica Code','Title Desc']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.write('File not selected yet')
    try:
        # Set date columns
        df_musica_merged['Start Date'] = Date_Start
   
        # Total amount column
        df_musica_merged['Total Amt'] = df_musica_merged['Sales Qty'] * df_musica_merged['RSP']

        # Add retailer column
        df_musica_merged['Forecast Group'] = 'Musica'

        # Rename columns
        df_musica_merged = df_musica_merged.rename(columns={'Musica Code': 'SKU No.'})
        df_musica_merged = df_musica_merged.rename(columns={'SMD code': 'Product Code'})

        # Don't change these headings. Rather change the ones above
        final_df_musica = df_musica_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]

        # Show final df
        total = final_df_musica['Total Amt'].sum()
        st.write('The total sales for the week are: ',locale.currency( total, grouping=True))
        final_df_musica

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_musica), unsafe_allow_html=True)
    except:
        st.write('Check data')

# Takealot
elif option == 'Takealot':
    try:
        # Get retailers map
        df_takealot_retailers_map = df_map
        df_retailers_map_takealot_final = df_takealot_retailers_map[['idProduct','Description','Manufacturer','SMD Code','RSP']]
        # Get retailer data
        df_takealot_data = df_data
        df_takealot_data = df_takealot_data.iloc[1:]
        #Merge with retailer map
        df_takealot_merged = df_takealot_data.merge(df_retailers_map_takealot_final, how='left', on='idProduct')    
        # Find missing data
        missing_model = df_takealot_merged['SMD Code'].isnull()
        df_takealot_missing_model = df_takealot_merged[missing_model]
        df_missing = df_takealot_missing_model[['idProduct','Supplier Code']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp = df_takealot_merged['RSP'].isnull()
        df_takealot_missing_rsp = df_takealot_merged[missing_rsp]
        df_missing_2 = df_takealot_missing_rsp[['idProduct','Supplier Code']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.write('File not selected yet')
    try:
        # Set date columns
        df_takealot_merged['Start Date'] = Date_Start

        # Total amount column
        df_takealot_merged['Total Amt'] = df_takealot_merged['Units Sold Qty'] * df_takealot_merged['RSP']

        # Add retailer and store column
        df_takealot_merged['Forecast Group'] = 'Takealot'
        df_takealot_merged['Store Name'] = ''

        # Rename columns
        df_takealot_merged = df_takealot_merged.rename(columns={'idProduct': 'SKU No.','Units Sold Qty' :'Sales Qty','Total SOH':'SOH Qty','SMD Code':'Product Code' })

        # Don't change these headings. Rather change the ones above
        final_df_takealot = df_takealot_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]

        # Show final df
        total = final_df_takealot['Total Amt'].sum()
        st.write('The total sales for the week are: ',locale.currency( total, grouping=True))
        final_df_takealot

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_takealot), unsafe_allow_html=True)
    except:
        st.write('Check data')


# TFG
elif option == 'TFG':
    try:
        # Get retailers map
        df_tfg_retailers_map = df_map
        df_retailers_map_tfg_final = df_tfg_retailers_map[['Article Code','Code','DES','RSP', 'Grouping']]
        # Get retailer data
        df_tfg_data = df_data
        # Apply the split string method on the Style code to get the SKU No. out
        df_tfg_data['Article Code'] = df_tfg_data['Style'].str.split(' ').str[0]
        # Convert to float
        df_tfg_data['Article Code'] = df_tfg_data['Article Code'].astype(float)
        # Merge with retailer map 
        df_tfg_merged = df_tfg_data.merge(df_retailers_map_tfg_final, how='left', on='Article Code')

        # Find missing data
        missing_model_tfg = df_tfg_merged['Code'].isnull()
        df_tfg_missing_model = df_tfg_merged[missing_model_tfg]
        df_missing = df_tfg_missing_model[['Article Code','Style']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_tfg = df_tfg_merged['RSP'].isnull()
        df_tfg_missing_rsp = df_tfg_merged[missing_rsp_tfg] 
        df_missing_2 = df_tfg_missing_rsp[['Article Code','Style']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.write('File not selected yet')

    try:
        # Set date columns
        df_tfg_merged['Start Date'] = Date_Start

        # Total amount column
        df_tfg_merged['Total Amt'] = df_tfg_merged['Sls (U)'] * df_tfg_merged['RSP']

        # Add retailer and store column
        df_tfg_merged['Forecast Group'] = 'TFG'
        df_tfg_merged['Store Name'] = ''

        # Rename columns
        df_tfg_merged = df_tfg_merged.rename(columns={'Article Code': 'SKU No.','Sls (U)' :'Sales Qty', 'CSOH Incl IT (U)':'SOH Qty', 'Code' : 'Product Code' })

        # Don't change these headings. Rather change the ones above
        df_tfg_merged = df_tfg_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]

        # Show final df
        total = df_tfg_merged['Total Amt'].sum()
        st.write('The total sales for the week are: ',locale.currency( total, grouping=True))
        df_tfg_merged

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(df_tfg_merged), unsafe_allow_html=True)
    except:
        st.write('Check data')
