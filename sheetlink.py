############this is a sheetlink, all of  our code handling populating items to 
#############and  from the spreadsheet directly or for tasks that requires only light computation
############## WARNING:c THIS FILE SEEMS INEFFICENT. THAT IS INTENTIONAL. CHANGING VERIFICATION 
##############SLIGHTLY CAN LEAD TO COMPOUNDING ILL EFFECTS
               
import os
import tkinter as tk
from tkinter import ttk, messagebox
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Load the credentials from the downloaded JSON file

#script_dir = os.path.dirname(os.path.abspath(__file__))
#file_path = os.path.join("/Users/westtn/Downloads/licinewgui-a3841a5e622d.json")
script_dir = os.path.dirname(os.path.abspath(__file__))

# Define the path to the credentials.txt file
credentials_file_path = os.path.join(script_dir, "credential_location.txt")

# Read the file path from credentials.txt
try:
    # Attempt to read the file path from credentials.txt
    with open(credentials_file_path, "r") as file:
        file_path = file.read().strip()  # Remove any extra whitespace or newlines

    if not file_path:
        raise ValueError("The credentials.txt file is empty. Please provide a valid file path.")

    print(f"File path set to: {file_path}")

except FileNotFoundError:
    print(f"Error: The file 'credentials.txt' was not found in the directory {script_dir}.")
except PermissionError:
    print(f"Error: Permission denied while trying to read 'credentials.txt' in the directory {script_dir}.")
except ValueError as ve:
    print(f"Error: {ve}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
#credentials = Credentials.from_service_account_file(file_path, scopes=['https://www.googleapis.com/auth/spreadsheets'])
try:
    credentials = Credentials.from_service_account_file(file_path, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    gc = gspread.authorize(credentials)
    #spreadsheet_id = '1zh9n4FcZCqL3w3GK59WLxBz9IYeVx9hDDdzH-Koawdg'
    #sheet = gc.open_by_key(spreadsheet_id)
except FileNotFoundError:
    print(f"Error: The file 'credentials.txt' was not found in the directory {script_dir}.")
    credentials=None

def update_creds():
   
    credentials_file_path = os.path.join(script_dir, "credential_location.txt")
def update_protocol():
   
   protocols_file_path = os.path.join(script_dir, "protocol_location.txt")

# Read the file path from credentials.txt
   with open(credentials_file_path, "r") as file:
        file_path = file.read().strip()
   credentials = Credentials.from_service_account_file(file_path, scopes=['https://www.googleapis.com/auth/spreadsheets']) 
def fetch_authors():
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet('Authors') # take the  current third worksheet
    return worksheet.col_values(1)
def add_author(new_author):
    try:
         gc = gspread.authorize(credentials)
         try:
             sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
             with open(sheetname_file_path, "r") as file:
                 spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
         # Authorize and open the desired spreadsheet by its ID
         except Exception as e:
             print(f"An unexpected error occurred: {e}")
         sheet = gc.open_by_key(spreadsheet_id)
         worksheet = sheet.worksheet('Authors')  # Access the 'Authors' worksheet
    
         # Check if the author already exists
         existing_authors = worksheet.col_values(1)
         if new_author in existing_authors:
             messagebox.showinfo("Duplicate Entry", f"Author '{new_author}' already exists.")
             return
    
         # Add the new author to the next available row
         next_row = len(existing_authors) + 1
         worksheet.update_cell(next_row, 1, new_author)
         messagebox.showinfo("Success", f"Author '{new_author}' has been added.")
    except Exception as e:
         messagebox.showerror("Error", f"Failed to add author: {e}")
def delete_author(selected_author):
    try:
        # Authorize and access the spreadsheet
        gc = gspread.authorize(credentials)
        try:
            sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
            with open(sheetname_file_path, "r") as file:
                spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
        # Authorize and open the desired spreadsheet by its ID
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
        sheet = gc.open_by_key(spreadsheet_id)
        worksheet = sheet.worksheet('Authors')  # Access the 'Authors' worksheet

        # Get the list of authors
        authors = worksheet.col_values(1)

        if selected_author not in authors:
            messagebox.showwarning("Not Found", f"Author '{selected_author}' does not exist.")
            return

        # Find the row index of the selected author
        row_index = authors.index(selected_author) + 1  # Convert to 1-based index for Google Sheets
        worksheet.delete_rows(row_index)
        messagebox.showinfo("Success", f"Author '{selected_author}' has been deleted.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to delete author: {e}")
def fetch_dyes():   
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet('cytekDyes->Channels') # take the  current third worksheet
    return worksheet.col_values(1)
def return_channel(idx):
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet('cytekDyes->Channels') # take the  current third worksheet
    return worksheet.cell(idx + 1, 2).value  # idx + 1 because gspread uses 1-based indexing
def fetch_inventory(species):
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    if "Mouse" in species:  worksheet = sheet.worksheet('Antibodies [Mouse]') # take the  current third worksheet
    else:  worksheet = sheet.worksheet('Antibodies [Human]')
    return worksheet.get_all_values()
def fetch_numbers(species):
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    if "Mouse" in species:worksheet = sheet.worksheet('Antibodies [Mouse]') # take the  current third worksheet
    else:worksheet = sheet.worksheet('Antibodies [Human]')
    colvals=worksheet.col_values(1)
    return colvals#dont fetch header

def insert_row(target_index,species):
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    if "Mouse" in species:  worksheet = sheet.worksheet('Antibodies [Mouse]') # take the  current third worksheet
    else:  worksheet = sheet.worksheet('Antibodies [Human]')
    worksheet.insert_row([], target_index + 1)
def add_antibody(target_index,new_values,species):
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    if "Mouse" in species:  worksheet = sheet.worksheet('Antibodies [Mouse]') # take the  current third worksheet
    else:  worksheet = sheet.worksheet('Antibodies [Human]')
    worksheet.update(f'A{target_index + 1}:K{target_index + 1}', [new_values]) 
def polish_row(target_index,color,species): #clear and purdy up the cell for 
    print("polisihing row")
    color = [c / 255 for c in color]
    print("color before writing is")
    print(color)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    if "Mouse" in species:  worksheet = sheet.worksheet('Antibodies [Mouse]') # take the  current third worksheet
    else:  worksheet = sheet.worksheet('Antibodies [Human]')
    worksheet.format(f'{target_index}:{target_index}', {
    "backgroundColor": None,
    "textFormat": {
        "bold": False,
        "italic": False,
        "underline": False,
        "strikethrough": False
        }
    })    
    worksheet.format(f'D{target_index}', {
    "backgroundColor": {
        "red": color[0],
        "green": color[1],
        "blue": color[2]
    }
})

def fetch_panel_names():
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet_names = [worksheet.title for worksheet in sheet.worksheets()]

    # Filter worksheet names that contain the word "Panel:"
    filtered_worksheets = [name for name in worksheet_names if "Panel:" in name]
    return filtered_worksheets
def fetch_panel_data(selected_panel):
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet(selected_panel)
    return worksheet.get_all_values()[1:-1]
def fetch_pure_panel(selected_panel):
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet(selected_panel)
    return worksheet.get_all_values()[:-1]
def update_panel(sheetname,headers,panel_data):
    
    try:
        c = gspread.authorize(credentials)
        try:
            sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
            with open(sheetname_file_path, "r") as file:
                spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
        # Authorize and open the desired spreadsheet by its ID
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
        sheet = gc.open_by_key(spreadsheet_id)
        worksheet = sheet.worksheet(sheetname)
        worksheet.clear()
    except gspread.exceptions.APIError as e:
        print(f"Error creating worksheet: {e}")
        return
    # de
    # Write the headers and panel data to the new worksheet
    worksheet.append_row(headers)
    worksheet.format(f'A1:H1', {
    "backgroundColor": None,
    "textFormat": {
        "bold": True,
        "italic": False,
        "underline": False,
        "strikethrough": False
        }
    })    
    worksheet.freeze(rows=1, cols=None)
    index=2
    for row in panel_data:
        worksheet.append_row(row,value_input_option="USER_ENTERED")
        channel=row[3]
        if "YG" in channel: color=[146,208,80] #load in RGB order
        elif "UV" in channel:color=[217,217,217]
        elif "R" in channel:color=[255,126,121]
        elif "B"  in  channel:color=[0,176,240]
        elif "V" in channel:color=[216,131,255]
        else:color=[1,1,1]
        color = [c / 255 for c in color]
        print("color before writing is")
        print(color)
        worksheet.format(f'D{index}', {
            "backgroundColor": {
        "red": color[0],
            "green": color[1],
        "blue": color[2]
        }
        })
        index=index+1
# Success message
    # Success message
    tk.messagebox.showinfo("Success", f"Panel '{sheetname}' created successfully in Google Sheets!")
def export_panel(new_sheet_name,headers,panel_data): #creates a simple panel, no need to invoke "logic"
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    try:
        new_worksheet = sheet.add_worksheet(title=new_sheet_name, rows=len(panel_data) + 10, cols=len(headers)+50)
    except gspread.exceptions.APIError as e:
        print(f"Error creating worksheet: {e}")
        return
    # de
    # Write the headers  and bold them  and  wirtepanel data to the new worksheet
    new_worksheet.append_row(headers)
    new_worksheet.format(f'A1:H1', {
    "backgroundColor": None,
    "textFormat": {
        "bold": True,
        "italic": False,
        "underline": False,
        "strikethrough": False
        }
    })    
    new_worksheet.freeze(rows=1, cols=None)
    index=2
    for row in panel_data:
        new_worksheet.append_row(row,value_input_option="USER_ENTERED")
        channel=row[3]
        if "YG" in channel: color=[146,208,80] #load in RGB order
        elif "UV" in channel:color=[217,217,217]
        elif "R" in channel:color=[255,126,121]
        elif "B"  in  channel:color=[0,176,240]
        elif "V" in channel:color=[216,131,255]
        else:color=[1,1,1]
        color = [c / 255 for c in color]
        print("color before writing is")
        print(color)
        new_worksheet.format(f'D{index}', {
        "backgroundColor": {
            "red": color[0],
            "green": color[1],
            "blue": color[2]
        }
    })
        index=index+1
    # Success message
    tk.messagebox.showinfo("Success", f"Panel '{new_sheet_name}' created successfully in Google Sheets!")
def check_alerts():
    gc = gspread.authorize(credentials)
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    # Authorize and open the desired spreadsheet by its ID
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet('Antibodies [Mouse]') 
    vals=worksheet.get_all_values()
    worksheet = sheet.worksheet('Antibodies [Human]')
    temp=[[""]+item for item in worksheet.get_all_values()] #required to align lists
    vals=temp+vals[1:]
    vals=vals[1:]
    #print(vals)
    v=[item for item in vals if item[10] != '' ]
    #print( [f"{item[10]}" for item in v if int(float(item[10]))<5])

    alerts=[ f"Catalog {item[8]}, a {item[5]} antibody, is low with {item[10]} ul left" for item in v if int(float(item[10]))<5 ]
    print(alerts)
    return alerts
def update_value(species,search_number, subtract_value):
    print(species)
    print(search_number)
    print(subtract_value)
    gc = gspread.authorize(credentials)
    
    # Read spreadsheet ID from file
    try:
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            spreadsheet_id = file.read().strip()  # Remove any extra whitespace or newlines
    except Exception as e:
        print(f"An unexpected error occurred while reading sheet_name.txt: {e}")
        exit(1)
    
    # Open the spreadsheet
    if "ouse" in species:
        sheetname = "Antibodies [Mouse]"  # Change to the actual sheet name if needed
    else:
        sheetname = 'Antibodies [Human]'
    sheet = gc.open_by_key(spreadsheet_id)
    worksheet = sheet.worksheet(sheetname)
    # Get all values from column A

    # Get all values from column A
    column_A = worksheet.col_values(1)
    print(column_A )

    try:
        # Find row index (Google Sheets is 1-based index)
        row_index = column_A.index(str(search_number)) + 1  # Convert to 1-based index
        
        # Get the value in column J of the same row
        current_value = worksheet.cell(row_index, 11).value  # Column J (10th column)
        
        if current_value is None or current_value == '':
            print(f"No value found in column J at row {row_index}.")
            return
        
        # Convert to number
        current_value = float(current_value)

        # Perform subtraction
        new_value = current_value - subtract_value

        # Update cell in column J
        worksheet.update_cell(row_index, 11, new_value)
        print(f"Updated row {row_index}: {current_value} - {subtract_value} = {new_value}")

    except ValueError:
        print("Number not found in column A.")