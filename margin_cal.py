import streamlit as st
import pandas as pd
import plotly.express as px
import webbrowser as wb
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import os
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from io import BytesIO
import calendar
from datetime import datetime, date, time
current_datetime = datetime.now()
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import math

st.set_page_config(page_title="Margin Calulator", page_icon="🧮", layout="wide")
st.write("🏢 Goh Office Supplies")
st.title("Margin Calculator🧮")
st.markdown("_________________________________________________________________")

def excel_file(name):
    file = st.file_uploader(name,type=['xlsx'])
    if not file:
        st.stop()
    df_cart = pd.read_excel(file)

    #rows_to_take = second_column[second_column.iloc[:, 0] == "Marketplace Order No."].index[0]
    #df_cart = df_cart.iloc[rows_to_take:]
    #df_cart.columns = df_cart.iloc[0]
    #df_cart = df_cart.drop([rows_to_take])

    #df_cart
    return df_cart

def dfs_to_excel(df_list, sheet_list, name, current_datetime):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for dataframe, sheet in zip(df_list, sheet_list):
            dataframe.to_excel(writer, sheet_name=sheet, index=False)
    output.seek(0)
    "#"
    st.download_button(
                    label=f"Export Data",
                    data=output,
                    file_name=f"Auto-bill_{name}_{current_datetime}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return output

df_oc = excel_file('OC Sales Order Enquiry')
"#"
"Note: Make sure that 'Model' and 'Cost' is in first sheet of the file"
df_cost = excel_file('Cost Excel File')

"________________________________________________________"
"RESULTS: "
df_merge = pd.merge(df_oc , df_cost, on='Model', how='left')
df_merge['New Margin'] = df_merge['Order Income By Item'] - (df_merge['Cost']*df_merge['Quantity'])
df_merge

"#"
"Missing Model Cost"
df_missing = df_merge[df_merge['Cost'].isna()]
df_missing = df_missing['Model'].to_frame()
df_missing = df_missing.drop_duplicates(subset='Model', keep='first')
df_missing = df_missing.reset_index()
df_missing = df_missing.drop(['index'], axis=1)
df_missing 

dfs_to_excel([df_merge, df_missing ], ['Sales Order Details', 'Missing Model Cost'], 'Margin', current_datetime)

