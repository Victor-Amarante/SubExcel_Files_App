import streamlit as st
import pandas as pd
import os

import base64
from io import StringIO, BytesIO

def generate_excel_download_link(df):
    # Credit Excel: https://discuss.streamlit.io/t/how-to-add-a-download-excel-csv-function-to-a-button/4474/5
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

st.set_page_config(page_title='Subdivisão Planilhas')

st.title("Subdivisão Planilhas 🗃️")
st.subheader('Importe uma planilha no formato Excel')

uploaded_file = st.file_uploader('Escolha um arquivo do tipo XLSX', type='xlsx')

if uploaded_file:
    st.markdown('---')
    df = pd.read_excel(uploaded_file, engine = 'openpyxl')
    st.dataframe(df)

    st.markdown('---')

    chunk_size = st.number_input("Tamanho de cada partição (linhas)", min_value=1, value=300)
    separate_button = st.button('Subdividir a planilha geral')
    
    if separate_button:
        n = 300
        list_df = [df[i:i+n] for i in range(0, df.shape[0], n)]
        # Write each smaller DataFrame to a separate Excel file
        for i, df in enumerate(list_df):
            st.write(f"Partição {i+1}:")
            generate_excel_download_link(df)
