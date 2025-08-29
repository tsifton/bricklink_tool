import os
from google.oauth2.service_account import Credentials
import gspread

ORDERS_DIR = "orders"
WANTED_LISTS_DIR = "wanted_lists"
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")

GOOGLE_SHEET_NAME = "Minifig Profit Tool"
CONFIG_TAB_NAME = "Config"
LEFTOVERS_TAB_NAME = "Leftover Inventory"

def load_google_sheet():
    # Load or create the main Google Sheet for the tool.
    creds = Credentials.from_service_account_file(
        CREDENTIALS_FILE,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
    )
    client = gspread.authorize(creds)
    try:
        return client.open(GOOGLE_SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        return client.create(GOOGLE_SHEET_NAME)

def get_or_create_worksheet(sheet, name, rows=100, cols=20):
    # Get a worksheet by name, or create it if it doesn't exist.
    try:
        return sheet.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        return sheet.add_worksheet(title=name, rows=str(rows), cols=str(cols))

def get_config_value(sheet, label, cell):
    # Get a configuration value from the config worksheet, prompting the user if missing.
    ws = get_or_create_worksheet(sheet, CONFIG_TAB_NAME)
    val = ws.acell(cell).value
    if not val or not val.strip():
        num = float(input(f"Enter a value for '{label}': ").strip())
        ws.update(values=[[num]], range_name=cell)
        return num
    return float(val)
