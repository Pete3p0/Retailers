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
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download='+option+'_'+Year+str(Month)+Day+".xlsx"'>Download Excel file</a>' # decode b'abc' => abc

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
    ('Please select','Bradlows/Russels','Clicks','Checkers','Incredible Connection','Makro', 'Musica','Takealot','TFG'))
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
        st.markdown("**Retailer map column headings:** Article number, Product Code & RSP")
        st.markdown("**Retailer data column headings:** Cluster, Article, Description, Site, Site Name, Valuated Stock Qty(Total), Sales Qty*")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

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
        st.markdown("**Retailer map column headings:** Article, SMD Code & RSP")
        st.markdown("**Retailer data column headings:** Item Code, Description, "+Units_Sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

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

# Clicks

elif option == 'Clicks':
    try:
        # Get retailers map
        df_clicks_retailers_map = df_map
        df_retailers_map_clicks_final = df_clicks_retailers_map[['Clicks Product Number','SMD CODE','SMD DESC','RSP']]

        # Get retailer data
        df_clicks_data = df_data
        df_clicks_data.columns = df_clicks_data.iloc[3]
        df_clicks_data = df_clicks_data.iloc[5:]

        # Drop result rows
        df_clicks_data.drop(df_clicks_data[df_clicks_data['Product Status'] == 'Sum:'].index, inplace = True) 

        # Merge with retailer map 
        df_clicks_merged = df_clicks_data.merge(df_retailers_map_clicks_final, how='left', on='Clicks Product Number')

        # Find missing data
        missing_model_clicks = df_clicks_merged['SMD CODE'].isnull()
        df_clicks_missing_model = df_clicks_merged[missing_model_clicks]
        df_missing = df_clicks_missing_model[['Clicks Product Number','Product Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_clicks = df_clicks_merged['RSP'].isnull()
        df_clicks_missing_rsp = df_clicks_merged[missing_rsp_clicks] 
        df_missing_2 = df_clicks_missing_rsp[['Clicks Product Number','Product Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)
    except:
        st.markdown("**Retailer map column headings:** Clicks Product Number,SMD CODE,SMD DESC,RSP")
        st.markdown("**Retailer data column headings:** Store Description, Clicks Product Number, Product Description, Store Stock Qty, Sales Qty LW TY")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")   

    try:
        # Set date columns
        df_clicks_merged['Start Date'] = Date_Start

        # Total amount column
        df_clicks_merged['Total Amt'] = df_clicks_merged['Sales Qty LW TY'] * df_clicks_merged['RSP']

        # Add retailer column
        df_clicks_merged['Forecast Group'] = 'Clicks'

        # Rename columns
        df_clicks_merged = df_clicks_merged.rename(columns={'Clicks Product Number': 'SKU No.'})
        df_clicks_merged = df_clicks_merged.rename(columns={'SMD CODE': 'Product Code'})
        df_clicks_merged = df_clicks_merged.rename(columns={'Store Description': 'Store Name'})
        df_clicks_merged = df_clicks_merged.rename(columns={'Store Stock Qty': 'SOH Qty'})
        df_clicks_merged = df_clicks_merged.rename(columns={'Sales Qty LW TY': 'Sales Qty'})

        # Don't change these headings. Rather change the ones above
        final_df_clicks = df_clicks_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        
        # Show final df
        total = final_df_clicks['Total Amt'].sum()
        st.write('The total sales for the week are: ',locale.currency( total, grouping=True))
        final_df_clicks

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_clicks), unsafe_allow_html=True)
    
    except:
        st.write('Check data')   

# Incredible Connection

elif option == 'Incredible Connection':
    try:
        Units_Sold = ('Qty Sold '+ str(Month) + '.' + Year)

        # Get retailers map
        df_ic_retailers_map = df_map
        

        # Get previous week
        ic_data_prev = st.file_uploader('Previous week', type='xlsx')
        if ic_data_prev:
            df_ic_data_prev = pd.read_excel(ic_data_prev)
        df_ic_data_prev['Lookup'] = df_ic_data_prev['Article'].astype(str) + df_ic_data_prev['Site']
        df_ic_data_prev = df_ic_data_prev.rename(columns={Units_Sold: 'Prev Sales'})
        df_ic_data_prev_final = df_ic_data_prev[['Lookup','Prev Sales']]

        # Get current week
        df_ic_data = df_data
        df_ic_data['Lookup'] = df_ic_data['Article'].astype(str) + df_ic_data['Site']

        # Rename columns
        df_ic_retailers_map = df_ic_retailers_map.rename(columns={'RRP': 'RSP'})

        # Merge with retailer map and previous week
        df_ic_data_merge_curr = df_ic_data.merge(df_ic_data_prev_final, how='left', on='Lookup')
        df_ic_merged = df_ic_data_merge_curr.merge(df_ic_retailers_map, how='left', on='Article')

        missing_model_ic = df_ic_merged['SMD Code'].isnull()
        df_ic_missing_model = df_ic_merged[missing_model_ic]
        df_missing = df_ic_missing_model[['Article','Article Name']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_ic = df_ic_merged['RSP'].isnull()
        df_ic_missing_rsp = df_ic_merged[missing_rsp_ic]
        df_missing_2 = df_ic_missing_rsp[['Article','Article Name']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Article, SMD Code & RRP")
        st.markdown("**Retailer data column headings:** Article, Article Name, Site, Site Name, Total SOH Qty & "+Units_Sold)
        st.markdown("Column headings are **case sensitive**")

    try:
        # Set date columns
        df_ic_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_ic_merged['Sales Qty'] = df_ic_merged[Units_Sold] - df_ic_merged['Prev Sales']
        df_ic_merged['Total Amt'] = df_ic_merged['Sales Qty'] * df_ic_merged['RSP']

        # Add column for retailer and SOH
        df_ic_merged['Forecast Group'] = 'Incredible Connection'

        # Rename columns
        df_ic_merged = df_ic_merged.rename(columns={'Article': 'SKU No.'})
        df_ic_merged = df_ic_merged.rename(columns={'Total SOH Qty': 'SOH Qty'})
        df_ic_merged = df_ic_merged.rename(columns={'SMD Code': 'Product Code'})
        df_ic_merged = df_ic_merged.rename(columns={'Site Name': 'Store Name'})

        # Final df. Don't change these headings. Rather change the ones above
        final_df_ic_sales = df_ic_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]

        # Show final df
        total = final_df_ic_sales['Total Amt'].sum()
        st.write('The total sales for the week are: ',locale.currency( total, grouping=True))
        final_df_ic_sales

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_ic_sales), unsafe_allow_html=True)

    except:
        st.write('Check data')


# Makro

elif option == 'Makro':
    Week = st.text_input("Enter week number: ")
    weekly_sales = Week+'-'+Year
    makro_stores = st.file_uploader('Stores', type='xlsx')
    if makro_stores:
        df_makro_stores = pd.read_excel(makro_stores)
   
    try:
        # Get retailers map
        df_makro_retailers_map = df_map
        df_retailers_map_makro_final = df_makro_retailers_map[['Article','SMD Product Code','SMD Description','RSP']]

        # Get retailer data
        df_makro_data = df_data

        # Merge with retailer map 
        df_makro_merged = df_makro_data.merge(df_retailers_map_makro_final, how='left', on='Article')

        # Merge with stores
        df_makro_merged = df_makro_merged.merge(df_makro_stores, how='left', on='Site')
        
        # Find missing data
        missing_model_makro = df_makro_merged['SMD Product Code'].isnull()
        df_makro_missing_model = df_makro_merged[missing_model_makro]
        df_missing = df_makro_missing_model[['Article','Article Desc']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_makro = df_makro_merged['RSP'].isnull()
        df_makro_missing_rsp = df_makro_merged[missing_rsp_makro]  
        df_missing_2 = df_makro_missing_rsp[['Article','Article Desc']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.write('File not selected yet')

    try:
        # Set date columns
        df_makro_merged['Start Date'] = Date_Start

        # Total amount column
        df_makro_merged['Total Amt'] = df_makro_merged[weekly_sales] * df_makro_merged['RSP']
        
        # Add retailer column
        df_makro_merged['Forecast Group'] = 'Makro'

        # Rename columns
        df_makro_merged = df_makro_merged.rename(columns={'Article': 'SKU No.'})
        df_makro_merged = df_makro_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_makro_merged = df_makro_merged.rename(columns={'SOH': 'SOH Qty'})
        df_makro_merged = df_makro_merged.rename(columns={weekly_sales: 'Sales Qty'})

        # Don't change these headings. Rather change the ones above
        final_df_makro = df_makro_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]

        # Show final df
        total = final_df_makro['Total Amt'].sum()
        st.write('The total sales for the week are: ',locale.currency( total, grouping=True))
        final_df_makro

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_makro), unsafe_allow_html=True)

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

else:
    st.write('File not selected yet')
