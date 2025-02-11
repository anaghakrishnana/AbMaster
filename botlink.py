import os
import tkinter as tk
from tkinter import ttk, messagebox
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Get script directory
script_dir = os.path.dirname(os.path.abspath(__file__))

# Define the path to the credentials.txt file
credentials_file_path = os.path.join(script_dir, "credential_location.txt")

try:
    # Read the credential file path from the text file
    with open(credentials_file_path, "r") as file:
        file_path = file.read().strip()

    if not file_path:
        raise ValueError("The credentials.txt file is empty. Please provide a valid file path.")

    print(f"File path set to: {file_path}")

    # Load credentials for Google Sheets API
    credentials = Credentials.from_service_account_file(
        file_path,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    
    gc = gspread.authorize(credentials)

    # Open Google Spreadsheet by ID
    spreadsheet_id = '1zh9n4FcZCqL3w3GK59WLxBz9IYeVx9hDDdzH-Koawdg'
    sheet = gc.open_by_key(spreadsheet_id)

    # Select the first worksheet
    worksheet = sheet.worksheet('Panel:Tim_TEST_PANEL') # You can also use `sheet.worksheet("Sheet Name")` for specific sheets

    # Get all values from the worksheet
    data = worksheet.get_all_values()

    # Convert to pandas DataFrame
    df = pd.DataFrame(data)

    # Set first row as column headers
    df.columns = df.iloc[0]  # Make first row the header
    df = df[1:]  # Remove the first row from the data

    # Define output file paths
    csv_output_path = os.path.join(script_dir, "exported_data.csv")
    excel_output_path = os.path.join(script_dir, "exported_data.xlsx")

    # Export DataFrame to CSV
    df.to_csv(csv_output_path, index=False)

    # Export DataFrame to Excel
    df.to_excel(excel_output_path, index=False, engine="openpyxl")

    print(f"Export completed: {csv_output_path} and {excel_output_path}")

except FileNotFoundError:
    print(f"Error: The file 'credentials.txt' was not found in the directory {script_dir}.")
except PermissionError:
    print(f"Error: Permission denied while trying to read 'credentials.txt' in the directory {script_dir}.")
except ValueError as ve:
    print(f"Error: {ve}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
