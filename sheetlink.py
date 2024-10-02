############this is a sheetlink all of  our code handling populating items to and  from the 
import os
import tkinter as tk
from tkinter import ttk
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Load the credentials from the downloaded JSON file
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join("/Users/westtn/Downloads/licinewgui-a3841a5e622d.json")
credentials = Credentials.from_service_account_file(file_path, scopes=['https://www.googleapis.com/auth/spreadsheets'])

# Authorize and open the desired spreadsheet by its ID
gc = gspread.authorize(credentials)
spreadsheet_id = '1zh9n4FcZCqL3w3GK59WLxBz9IYeVx9hDDdzH-Koawdg'
sheet = gc.open_by_key(spreadsheet_id)


def fetch_mouse_inventory():
    gc = gspread.authorize(credentials)
    spreadsheet_id = '1zh9n4FcZCqL3w3GK59WLxBz9IYeVx9hDDdzH-Koawdg'
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet('Antibodies [Mouse]') # take the  current third worksheet
    return worksheet.get_all_values()
def fetch_mouse_numbers():
    gc = gspread.authorize(credentials)
    spreadsheet_id = '1zh9n4FcZCqL3w3GK59WLxBz9IYeVx9hDDdzH-Koawdg'
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet('Antibodies [Mouse]') # take the  current third worksheet
    print(worksheet.row_values(1)[1:] )
    return worksheet.row_values(1)[1:]  
def fetch_human_inventory():
    gc = gspread.authorize(credentials)
    spreadsheet_id = '1zh9n4FcZCqL3w3GK59WLxBz9IYeVx9hDDdzH-Koawdg'
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet('Antibodies [Human]') # take the  current third worksheet
    return worksheet.get_all_values()
def fetch_human_numbers():
    gc = gspread.authorize(credentials)
    spreadsheet_id = '1zh9n4FcZCqL3w3GK59WLxBz9IYeVx9hDDdzH-Koawdg'
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet('Antibodies [Human]') # take the  current third worksheet
    return worksheet.row_values(1)[1:] 
def add_antibody(): #
    # the antibody has been 
    1

def export_panel(new_sheet_name,headers,panel_data):
    try:
        new_worksheet = sheet.add_worksheet(title=new_sheet_name, rows=len(panel_data) + 10, cols=len(headers)+50)
    except gspread.exceptions.APIError as e:
        print(f"Error creating worksheet: {e}")
        return
    # de
    # Write the headers and panel data to the new worksheet
    new_worksheet.append_row(headers)
    for row in panel_data:
        new_worksheet.append_row(row)

    # Success message
    tk.messagebox.showinfo("Success", f"Panel '{new_sheet_name}' created successfully in Google Sheets!")
    