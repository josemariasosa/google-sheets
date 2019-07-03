# google-sheets

This repository could eventually help you to: download, modify and upload data from/to a Google Spread Sheet using the gspread library in Python.

## Resources

As an example, we are going to use the following spread sheet from Google Drive: https://docs.google.com/spreadsheets/d/19v5Zlc9FrHV5avziwdAsSBlki_m-xnwohgnwZAJLkoQ/edit?usp=sharing.

## Steps

### 1. Generate credentials.

1. Visit: [Google Cloud Console](https://console.cloud.google.com/).
2. Generate a new project. I named mine: **google-sheet-test**.
3. Open the project in the Cloud Console.
4. Click on API overview, or on API Services.
5. On the top of the screen, inside the search bar, look for **Google Drive API**, click on it.
6. Click on the Enable button.
7. We will need to generate credentials for this API, so, click on **Generate Credentials**.
8. Which API are you using? Select: **Google Drive API**.
9. Where will you be calling the API from? Select: **Web Server**.
10. What data will you be accessing? Select: **Application data**. The last question, Are you planning to use this API with App Engine or Compute Engine? Select: No.


