#!/usr/bin/env python
# coding=utf-8

# ------------------------------------------------------------------------------
# Connect to Google Drive API | Exercise 1 Solution
# ------------------------------------------------------------------------------
# jose maria sosa


import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]

file_name = "credentials.json"
creds = ServiceAccountCredentials.from_json_keyfile_name(file_name, scope)
client = gspread.authorize(creds)

book_name = 'sales'
sheet_name = 'sales1'
sheet = client.open(book_name).worksheet(sheet_name)

# ------------------------------------------------------------------------------
# 1. Print the Total_Orders.

col = sheet.col_values(2)

print(col)


# ------------------------------------------------------------------------------
# 2. Print the sales and orders only of May 2019.

row = sheet.row_values(7)

print(row)


# ------------------------------------------------------------------------------
# 3. Print the table of sales as a Pandas DataFrame.

data = sheet.get_all_records()
data = pd.DataFrame(data)

print(data)


# ------------------------------------------------------------------------------
# 4. Calculate the Forecast for september using the average of the last 3 month.

forecast_sales = sum(data.loc[5:8, 'Sales'].tolist())/3
forecast_orders = sum(data.loc[5:8, 'Total_Orders'].tolist())/3

newRow = ['2019-09', forecast_sales, forecast_orders]
sheet.append_row(newRow)

# Import the sheet again with the new changes.
sheet = client.open(book_name).worksheet(sheet_name)


# ------------------------------------------------------------------------------
# 5. Calculate the Average Ticket Price.

data = sheet.get_all_records()
data = pd.DataFrame(data)

data['avg_tix'] = data['Sales'] / data['Total_Orders']

newCol = data['avg_tix'].tolist()
newCol.insert(0, 'Avg_Tix')

cell_list = sheet.range('D1:D11')

for ix, cell in enumerate(cell_list):
    cell.value = newCol[ix]

# Update in batch
sheet.update_cells(cell_list)

