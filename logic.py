import pandas as pd
from tkinter import messagebox
import numpy as np

EXCEL_FILE = "antibody_data.xlsx"
SHEET_NAME = "Antibodies"

def load_antibody_data():
    """Load the antibody data from the Excel sheet into a DataFrame."""
    try:
        # Read the Excel sheet into a DataFrame
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, engine='openpyxl')
        print("DataFrame loaded from Excel:")
        print(df)
    except FileNotFoundError:
        # If the file doesn't exist, create an empty DataFrame with the correct columns
        df = pd.DataFrame(columns=['Antibody Number', 'Antibody Specificity', 'Label', 'Supplier', 'Catalog Number', 'Target Species', 'Host Species', 'Original Volume', 'Volume', 'Number of Vials'])
        print("Excel file not found. Created a new DataFrame.")
    return df

def save_antibody_data(df):
    """Save the antibody DataFrame back to the Excel file."""
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    print("DataFrame saved to Excel.")

# Load the initial data
antibody_df = load_antibody_data()

def find_next_antibody_number(df):
    """Find the lowest available antibody number not already taken."""
    if df.empty:
        return 1
    existing_numbers = set(df['Antibody Number'])
    # Starting at 1, find the first number not in the set of existing numbers
    antibody_number = 1
    while antibody_number in existing_numbers:
        antibody_number += 1
    return antibody_number

def add_antibody_to_dataframe(data):
    """Add a new antibody entry to the DataFrame."""
    global antibody_df

    # Extract catalog number and supplier from the data
    catalog_number = data.get('Catalog Number')
    supplier = data.get('Supplier')
    
    # Check if the entry with the same catalog number and supplier already exists
    match = antibody_df[(antibody_df['Catalog Number'] == catalog_number) & (antibody_df['Supplier'] == supplier)]
    
    if not match.empty:
        # If a match is found, update the number of vials
        current_vials = match.iloc[0]['Number of Vials']
        new_vials = int(current_vials) + int(data.get('Number of Vials', 0))
        
        # Update the DataFrame with the new number of vials
        antibody_df.loc[(antibody_df['Catalog Number'] == catalog_number) & (antibody_df['Supplier'] == supplier), 'Number of Vials'] = new_vials
        
        # Show a dialog box with the updated vials info
        antibody_number = match.iloc[0]['Antibody Number']
        messagebox.showinfo("Info", f"Updated Number of Vials for Antibody {antibody_number}: Catalog Number {catalog_number} and Supplier {supplier}. New total: {new_vials}.")
    else:
        # Find the next available antibody number
        next_antibody_number = find_next_antibody_number(antibody_df)
    
        # Add the antibody number to the data
        data['Antibody Number'] = next_antibody_number

        # Initialize 'Original Volume' with the same value as 'Volume'
        data['Original Volume'] = data['Volume']
    
        # Append the new data as a row in the DataFrame
        antibody_df = antibody_df.append(data, ignore_index=True)

        messagebox.showinfo("Info", f"Created New Antibody {next_antibody_number}")
    # Save the updated DataFrame to the Excel sheet
    save_antibody_data(antibody_df)
    


def log_antibody_issue(antibody_number, error_message):
    global antibody_df
    
    # Ensure there's an Error Message column in the DataFrame
    if 'Error Message' not in antibody_df.columns:
        antibody_df['Error Message'] = ''
    
    # Find the row corresponding to the antibody number
    row_index = antibody_df.index[antibody_df['Antibody Number'] == antibody_number].tolist()
    
    if not row_index:
        return False  # Antibody number not found

    # Update the Error Message column
    row_index = row_index[0]  # Get the first match
    antibody_df.at[row_index, 'Error Message'] = error_message
    
    # Save the DataFrame
    save_antibody_data(antibody_df)
    
    return True

def use_antibody(antibody_number, volume_used):
    """Update the volume of the specified antibody and return the new volume and vials left."""
    global antibody_df

    # Find the row with the specified antibody number
    row_index = antibody_df.index[antibody_df['Antibody Number'] == antibody_number]
    if not row_index.empty:
        print('Test')
        row_index = row_index[0]
        current_volume = float(antibody_df.at[row_index, 'Volume'])
        original_volume = float(antibody_df.at[row_index, 'Original Volume'])
        vials = float(antibody_df.at[row_index, 'Number of Vials'])

        # Subtract the volume used
        new_volume = current_volume - np.ceil(volume_used)

        # Check if a vial is completely used up
        if new_volume <= 0 and vials > 0:
            new_volume = float(original_volume) + new_volume
            vials -= 1  # Decrease the vial count
        
        if vials == 0:
            antibody_df = antibody_df.drop(index=row_index)
        else:
            # Update the DataFrame
            antibody_df.at[row_index, 'Volume'] = new_volume
            antibody_df.at[row_index, 'Number of Vials'] = vials

        # Save the updated DataFrame to the Excel sheet
        save_antibody_data(antibody_df)
        
        if new_volume < 10 and vials == 1:
            antibody_number = antibody_df.at[row_index, 'Antibody Number']
            message = f'Please order antibody {antibody_number} as we are running low'
            messagebox.showwarning("Low Volume Warning", message)
        
        return True, vials, new_volume
    else:
        return None, None, None

def check_low_volume_antibodies():
    global antibody_df  # Declare the use of the global variable

    if antibody_df is None or antibody_df.empty:
        # Handle the case where antibody_df is not available or empty
        return
    
    """Check for antibodies that have less than 10μL volume and only 1 vial."""
    low_volume_antibodies = antibody_df[(antibody_df['Volume'] <= 10) & (antibody_df['Number of Vials'] == 1)]
    # Include Supplier and Catalog Number in the result
    return low_volume_antibodies[['Antibody Number', 'Supplier', 'Catalog Number', 'Volume']].to_dict('records')

def update_last_user(antibody_number, last_user):
    """Update the last user for the given antibody number in the antibody DataFrame."""
    # Ensure antibody_df exists
    if 'antibody_df' in globals():
        # Find the row corresponding to the antibody_number
        if antibody_number in antibody_df['Antibody Number'].values:
            antibody_df.loc[antibody_df['Antibody Number'] == antibody_number, 'Last User'] = last_user
        else:
            print(f"Antibody number {antibody_number} not found.")
    else:
        print("Antibody DataFrame not available.")