import streamlit as st
# import camelot.io as camelot
import camelot
from tabulate import tabulate
import pandas as pd
from PyPDF2 import PdfFileReader
from pathlib import Path
import os
import numpy as np
import json
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import xlsxwriter


def get_pdf_content_lines(pdf_file_path):
    pdf_lines = []
    with open(pdf_file_path, 'rb') as f:
        pdf_reader = PdfFileReader(f)
        for page in pdf_reader.pages:
            for split_line in page.extractText().splitlines():
                pdf_lines.append(split_line)
        return pdf_lines
# def get_pdf_content_lines(pdf_file):
#     pdf_lines = []
#     pdf_reader = PdfFileReader(pdf_file)
#     for page in pdf_reader.pages:
#         for split_line in page.extractText().splitlines():
#             pdf_lines.append(split_line)
#     return pdf_lines


class NpEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        if isinstance(obj, np.floating):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        return super(NpEncoder, self).default(obj)


def to_excel(df_con):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_con.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.close()
    processed_data = output.getvalue()
    return processed_data

#@st.cache_data
def convert(file):
    if not file:
        return None

    #save_folder = str(Path.home())
    save_folder = os.getcwd()
    save_path = Path(save_folder, file.name)
    print("save_path: ", save_path)
    with open(save_path, mode='wb') as w:
        w.write(file.getvalue())
    if save_path.exists():
        print(f'File {file.name} is successfully saved!')
    file_url = str(save_path)
    print("Saved file locally at file_url: ", file_url)

    # Extract tables
    tables = camelot.read_pdf(file_url, flavor='stream', pages='all')
    print("Total tables extracted:", tables.n)
    print("\n")

    # Get line-by-line array for proper sequencing of salesman data
    pdf_lines = get_pdf_content_lines(file_url)
    print("Processing PDF Lines:")
    last_salesman = ""
    salesman_dict = {}
    for line in pdf_lines:
        print(line)
        line_splits = line.split(' ')
        if len(line_splits) < 2:
            continue

        po_number = line_splits[0]
        if line_splits[1] == "Salesman:":
            last_salesman = po_number
            continue
        salesman_dict[po_number] = last_salesman

    print(salesman_dict)
    print("\n")

    # Data Cleanup
    master_df = None
    # print the first table as Pandas DataFrame
    for table in tables:
        #print("\n")
        #print(tabulate(table.df, headers="keys"))
        if table.df is not None:
            if master_df is None:
                master_df = table.df
            else:
                master_df = pd.concat([master_df, table.df], ignore_index=True)
                # Remove duplicate tables that get included when camelot gets confused about the boundaries of multiple tables on one page.
                master_df = master_df.drop_duplicates(keep='first')

    #print(tabulate(master_df, headers="keys"))
    #print("\n")

    if master_df is not None:
        # Set Dataframe headers
        headerLabels = master_df.loc[0]
        print("headerLabels:\n", headerLabels)
        print("\n")
        master_df.columns = master_df.loc[0]

        # Drop all header rows
        master_df = master_df[master_df[headerLabels[0]] != headerLabels[0]]

        # Drop all title, subtitle and total rows
        master_df = master_df[master_df[headerLabels[0]] != ""]

        # Drop salesman rows
        master_df = master_df[master_df[headerLabels[0]] != "Salesman:"]
        # Add Salesman data column and values
        master_df = master_df.assign(Salesman='')
        for index, row in master_df.iterrows():
            master_df.at[index, 'Salesman'] = salesman_dict[row.iloc[0]]

        print(tabulate(master_df, headers="keys"))

        #Change all int type columns to float type
        print(master_df.dtypes)
        for c in master_df.columns:
            if master_df[c].dtype == int:
                master_df[c] = master_df[c].astype(float)


        binary_data = to_excel(master_df)
        st.download_button(
            label="Download Xlsx File",
            data=binary_data,
            file_name='pdf_tables.xlsx',
        )


# Streamlit UI

st.title("PDF Tables to Xlsx")

st.header("Input")

if 'input_file_exists' not in st.session_state:
    st.session_state.input_file_exists = False
input_file = st.file_uploader(label='Choose a PDF', key='input_file', type=["pdf"])
if input_file is not None:
    st.session_state.input_file_exists = True

df = convert(input_file)

#if st.button('Convert', disabled=not st.session_state.input_file_exists):
#    st.write(df)



# if :
#     st.table(data=output_file)
#     st.download_button(
#         label="Download Xlsx",
#         data=output_file,
#         file_name='pdf_tables.xlsx',
#         mime='application/xlsx',
#     )
