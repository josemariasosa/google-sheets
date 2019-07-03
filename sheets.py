#!/usr/bin/env python
# coding=utf-8

# ------------------------------------------------------------------------------
# Connect to Google Drive API
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

sheet = client.open("google-sheet-test").sheet1
# sheet = client.open("google-sheet-test").worksheet(sheet_name)

data = sheet.get_all_records()

data = pd.DataFrame(data)

print(data)

row = sheet.row_values(3)

col = sheet.col_values(3)

cell = sheet.cell(1,2).value

print(row)
print(col)
print(cell)

# # insert
# insertRow = [5, "Ricardo", "Amarillo"]
# sheet.insert_row(insertRow, 4)

# # delete
# sheet.delete_row(4)

# # update
# sheet.update_cell(2,2, "Changed!")
