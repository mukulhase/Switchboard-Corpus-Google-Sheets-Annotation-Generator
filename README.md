# Switchboard-Corpus-Google-Sheets-Annotation-Generator
A script to generate a Google Sheet containing all dialogues and extra fields with data validation for the purpose of bulk collaborative annotation.

## Usage
1. Create project on Google API console and then create google service account. Described here: “Manage Google Spreadsheets with Python and Gspread” by Alexander Molochko https://link.medium.com/kFmu56NXPZ

2. Download keyfile JSON for that service account

3. Create target spreadsheet from your personal (sheets.google.com) account and share with service account with edit access.

4. Get id of the Google sheet from the url (the part after `document/` and before `/edit`

5. Install pip dependencies. `pip install -r requirements.txt`

6. Run create.py (uses python 3) with the appropriate arguments. (`python create.py -h` for help)
