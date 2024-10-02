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
st.markdown("#")

def excel_file(name):
    file = st.file_uploader(name,type=['xlsx'])
    if not file:
        st.stop()
    df_cart = pd.read_excel(file)

    #rows_to_take = second_column[second_column.iloc[:, 0] == "Marketplace Order No."].index[0]
    #df_cart = df_cart.iloc[rows_to_take:]
    #df_cart.columns = df_cart.iloc[0]
    #df_cart = df_cart.drop([rows_to_take])

    df_cart
    "#"
    return df_cart

df_oc = excel_file('OC Sales Order Enquiry')
df_cost = excel_file('Cost Excel File')

df_merge = pd.merge(df_oc , df_cost, on='Model', how='left')
df_merge
