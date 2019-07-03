#!/usr/bin/env python
# coding=utf-8

# ------------------------------------------------------------------------------
# Connect to Google Drive API Class
# ------------------------------------------------------------------------------
# jose maria sosa

import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials


class ConnectGoogleSheet(object):

    """ Connect to a Google Sheet to download, modify and upload data.
    """

    def __init__(self, book_name, sheet_name):

        self.book_name = book_name  
        self.sheet_name = sheet_name

        # Connect to the spread sheet.
        self.connectSheet() 

    # --------------------------------------------------------------------------

    def connectSheet(self):

        # Google Drive Credentials.
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive"
        ]
        file_name = "credentials.json"
        creds = ServiceAccountCredentials.from_json_keyfile_name(file_name,
                                                                 scope)
        client = gspread.authorize(creds)

        # Open the Spread Sheet.
        self.sheet = client.open(self.book_name).worksheet(self.sheet_name)

        return None

    # --------------------------------------------------------------------------

    def numberToLetters(self, q):

        """ Use this function to select a range of cells ('A1:B4' for example)
        """

        q = q - 1
        result = ''

        while q >= 0:
            remain = q % 26
            result = chr(remain+65) + result;
            q = q//26 - 1

        return result

    # --------------------------------------------------------------------------

    def updateAllTable(self, results):

        """ Get and format the data from Google Analytics.
        """

        # Reshape the old Spreadsheet.
        num_lines, num_columns = results.shape
        self.sheet.resize(rows=num_lines+1)

        # Select the list of cells.
        cells = 'A2:' + self.numberToLetters(num_columns) + str(num_lines + 1)
        cell_list = self.sheet.range(cells)

        # Modifying the values in the range.
        for cell in cell_list:
            val = results.iloc[cell.row-2,cell.col-1]
            cell.value = val

        # Update in batch. Setting up parameters.
        limit = len(cell_list)
        batch_size = 20000 # The size was arbitrarily selected, but it works.
        range_init = 0
        range_end = batch_size

        while limit > 0:
            new_list = cell_list[range_init:range_end]
            self.sheet.update_cells(new_list)

            range_init = range_init + batch_size
            range_end = range_end + batch_size
            limit = limit - batch_size

        return None

    # --------------------------------------------------------------------------
    
    def getAllTable(self):

        table = self.sheet.get_all_values()
        headers = table.pop(0)

        table = pd.DataFrame(table, columns=headers)

        return table

    # --------------------------------------------------------------------------
# ------------------------------------------------------------------------------


def main():

    sheet = ConnectGoogleSheet("google-sheet-test", "Sheet1")

    print(sheet.getAllTable())

# ------------------------------------------------------------------------------


if __name__ == '__main__':
    main()
