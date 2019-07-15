# google-sheets

This repository could eventually help you to: download, modify and upload data from/to a Google Spread Sheet using the gspread library in Python.

## Resources

As an example, we are going to use the following spread sheet from Google Drive: https://docs.google.com/spreadsheets/d/19v5Zlc9FrHV5avziwdAsSBlki_m-xnwohgnwZAJLkoQ/edit?usp=sharing.

## Steps to connect with Google Drive and Sheets.

### 1. Generate credentials.

1. Visit: [Google Cloud Console](https://console.cloud.google.com/).
2. Generate a new project. I named mine: **google-sheet-test**.
3. Open the project in the Cloud Console.
4. Click on API overview, or on API Services. On the top of the screen, inside the search bar, look for **Google Drive API**, click on it. Click on the Enable button.
5. We will need to generate credentials for this API, so, click on **Generate Credentials**.
6. Which API are you using? Select: **Google Drive API**. Where will you be calling the API from? Select: **Web Server**. What data will you be accessing? Select: **Application data**. The last question, Are you planning to use this API with App Engine or Compute Engine? Select: No. Click on the blue button: What credentials do I need?.
7. In the section this section, Add credentials to your project, generate a **Service account name**. And as a role, select Project Editor. Finally, make sure Key type JSON is selected, and click continue.
8. A JSON file is downloaded with our credentials.
9. Now, we will go back to the project in the Cloud Console and enable another API. Search for **Google Sheets API** and click on the Enable button.
10. At last, move the json file to the project folder (google-sheet) and change the name of the file to `credentials.json`.

### 2. Give access to the Google Sheet.

1. From the `credentials.json` file, copy the **client_email**.
2. Open the Spread Sheet, for this [example go here](https://docs.google.com/spreadsheets/d/19v5Zlc9FrHV5avziwdAsSBlki_m-xnwohgnwZAJLkoQ/edit?usp=sharing).
3. Click on Share button. Add the **client_email** to the People list. Make sure the option **Can edit** is selected. And click the Send button.

### 3. Create a new Python 3 project.

1. Using the terminal, create a new virtual env. And install the following libraries `gspread`, `oauth2client` and `pandas`: 

```bash
$ virtualenv -p python3 venv

$ source venv/bin/activate

$ pip install gspread oauth2client pandas
```

2. Create a new python file, we named it `sheets.py`. And import the libraries.

```python
import gspread
from oauth2client.service_account import ServiceAccountCredentials
```

3. Create, inside the python file, a list called **scope** using the following values. And set up the credentials using the json file. Lastly, generate a client using `gspread`.

```python
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]

file_name = "credentials.json"
creds = ServiceAccountCredentials.from_json_keyfile_name(file_name, scope)

client = gspread.authorize(creds)
```

4. Open the sheet of the Google Spread Sheet using the name of the book and the sheet we want. Use Pandas to display the data.

```python
sheet = client.open("google-sheet-test").sheet1
# sheet = client.open("google-sheet-test").worksheet(sheet_name)

data = sheet.get_all_records()

data = pd.DataFrame(data)

print(data)
```

Instead of the attribute `sheet1` you can also use `worksheet` and give the name of the sheet as an argument.

5. If only the row values are needed. The given number is the row position, being the position 1 the headers. For column is very similar. Also, we can even get the value inside of a position in the spread sheet.

```python
row = sheet.row_values(3)

col = sheet.col_values(3)

cell = sheet.cell(1,2).value

print(row)
print(col)
print(cell)
```

6. If we want to insert, delete or update information to the spread sheet.

```python
# insert
insertRow = [5, "Ricardo", "Amarillo"]
sheet.insert_row(insertRow, 4)

# delete
sheet.delete_row(4)

# update
sheet.update_cell(2,2, "Changed!")
```

### 4. Declaring a class in Python to modify large tables.

In the file, called `connectGoogleSheet.py`, we can find a class named **ConnectGoogleSheet** with 2 main methods:

1. getAllTable() - Import as a Pandas object all the data in the table from the spread sheet.
2. updateAllTable(new_table) - Upload all the new table to the spread sheet.

**Warning!** The updateAllTable method will delete all current data in the spread sheet and overwrite all the new information, so it must be used with care.


