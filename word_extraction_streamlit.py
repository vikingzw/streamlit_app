from urllib import response
from xml.etree.ElementInclude import include
import streamlit as st
import docx2txt
# import docx
import pandas as pd
import os
import numpy as np
from word_extraction_functions_traps import get_traps_data
from word_extraction_functions_stamps import get_stamps_data
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

def to_excel(df):
    # https://stackoverflow.com/questions/67564627/how-to-download-excel-file-in-python-and-streamlit
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def main():
    ############################### Preliminaries ####################################
    st.set_page_config(
        page_title='Summarize Reports', 
        page_icon=None, layout='centered', 
        initial_sidebar_state="auto", 
        menu_items=None
        )

    st.title('Summarize Reports')

    ###################################################################################
    st.write(""" ## Select test type""")
    test_type = st.selectbox(' ',('Traps','Stamps'),index=0)


    ###################################################################################
    st.write("""
    ***
    """)
  
    files = st.file_uploader('Upload files',type=['docx'],accept_multiple_files=True)


    if len(files)>0:
        if test_type == 'Traps':
            try:
                df = get_traps_data(files)
            finally:
                st.write(""" ### Reload the page when switching from Traps to Stamps or vice versa. """)
            start_date = df['Sampling Date'].iloc[0]
            end_date = df['Sampling Date'].iloc[-1]
        else:
            try:
                df = get_stamps_data(files)
            finally:
                st.write(""" ### Reload the page when switching from Traps to Stamps or vice versa. """)

            start_date = df['Date of extraction'].iloc[0]
            end_date = df['Date of extraction'].iloc[-1]

        # st.write(df)
        fname = '{}_data_{}_{}.xlsx'.format(test_type,start_date,end_date)

        df_xlsx = to_excel(df)
        
        st.download_button(
            label = 'Download Excel File',
            data = df_xlsx,
            file_name = fname)

        # st.write('Reload the page when switching from Traps to Stamps or vice versa.')



if __name__ == '__main__':
    main()