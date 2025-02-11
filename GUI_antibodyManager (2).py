import os
import tkinter as tk
from tkinter import ttk, messagebox, Label, Button, LabelFrame, Frame
import logic
import sheetlink
from tkinter import filedialog
import pandas as pd
from PIL import Image, ImageTk
import re
from decimal import Decimal, ROUND_HALF_UP
import pickle
import base64 
import requests
import json
import gspread #google sheets editor module. be wary of stability 
from google.oauth2.service_account import Credentials
# update to use main window
#anagha to add back buttons 
#set the  settings page to  add items to the list
#set the setting to allow you to select the credntials

import webbrowser

#%% Main Window
class AntibodyScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Antibody Scraper Main Menu")
        
        # Create a frame to hold the buttons
        self.frame = ttk.Frame(root)
        self.frame.pack(expand=True)
        
        # Main menu buttons
        ttk.Button(self.frame, text="Add an Antibody", command=self.open_add_window).pack(padx=10, pady=10)
        ttk.Button(self.frame, text="Create a Panel", command=self.open_create_panel_window).pack(padx=10, pady=10)
        ttk.Button(self.frame, text="Modify a Panel", command=self.open_modify_panel_window).pack(padx=10, pady=10)
        ttk.Button(self.frame, text="Run a Panel", command=self.open_run_panel_window).pack(padx=10, pady=10)
        ttk.Button(self.frame, text="Use Single Antibody", command=self.open_use_antibody_window).pack(padx=10, pady=10)
        ttk.Button(self.frame, text="Log Antibody Issue", command=self.open_log_issue_window).pack(padx=10, pady=10)
    
        # Perform low volume check at startup
        self.check_low_volume_antibodies()
    
    # def open_interim_antibody_window(self):
    #     destination="Add"
    #     InterimAntibodyWindow(self.root,destination)
    def open_add_window(self):
         AddAntibodyWindow(self.root)  
    def open_create_panel_window(self):
        #PanelCreatorApp(self.root)
       destination="Panel"
       InterimAntibodyWindow(self.root,destination)

    def open_modify_panel_window(self):
        PanelEditorApp(root,"Mouse")  # Pass the correct window as the master
        print("passed")
    def open_run_panel_window(self):
        RunPanelWindow(self.root)

    def open_use_antibody_window(self):
        UseAntibodyWindow(self.root)

    def open_log_issue_window(self):
        #LogIssueWindow(self.root)
        logic.determine_free_inventory()
    def check_low_volume_antibodies(self):
        """Check for antibodies with volume < 10μL and only 1 vial remaining."""
        low_volume_antibodies = logic.check_low_volume_antibodies()

        if low_volume_antibodies:
            message = "The following antibodies have less than 10μL left and only 1 vial remaining:\n"
            for antibody in low_volume_antibodies:
                message += (f"Antibody Number: {antibody['Antibody Number']}, "
                            f"Supplier: {antibody['Supplier']}, "
                            f"Catalog Number: {antibody['Catalog Number']}, "
                            f"Please order more.\n")
            messagebox.showwarning("Low Volume Warning", message)
            
            
#Dynamic 
class InterimAntibodyWindow:
    def __init__(self, root,destination):
        self.destination=destination
        self.interim_window = tk.Toplevel(root)
        self.interim_window.title("Add an Antibody")

        option_label = tk.Label(self.interim_window, text="Human or Mouse Antibody?")
        option_label.pack(pady=10)

        self.species_var = tk.StringVar()
        self.species_var.set("Mouse")
        options = ["Mouse", "Human"]
        dropdown = tk.OptionMenu(self.interim_window, self.species_var, *options)
        dropdown.pack()
        submit_button = tk.Button(self.interim_window, text="Submit", command=self.push_destination)
        submit_button.pack(pady=20)
   
    def open_panel_window(self):
        PanelCreatorApp(root,self.species_var.get())  # Pass the correct window as the master
        print("passed")
        self.interim_window.destroy()
   # def open_modify_window(self):
         # PanelEditorApp(root,self.species_var.get())  # Pass the correct window as the master
         # print("passed")
         # self.interim_window.destroy()
    def open_add_antibody_window(self):
        AddAntibodyWindow(root,self.species_var.get())  # Pass the correct window as the master
        print("passed")
        self.interim_window.destroy()
    def push_destination(self):
         if "Add" in self.destination:self.open_add_antibody_window()
         elif "Panel" in self.destination:self.open_panel_window()
         elif "Modify" in self.destination:self.open_modify_window()
 #%% Add an Antibody         
class AddAntibodyWindow():
    def __init__(self, root):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Add an Antibody")
        # Company input as a dropdown
        ttk.Label(self.new_window, text="Company:").grid(row=0, column=0, padx=10, pady=10)
        self.company_var = tk.StringVar()
        self.company_combobox = ttk.Combobox(self.new_window, textvariable=self.company_var)
        self.company_combobox['values'] = ("BD", "Biolegend","Thermofisher", "Tonbo", "Other (manual entry)")
        self.company_combobox.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(self.new_window, text="Species:").grid(row=1, column=0, padx=10, pady=10)
        self.species_var = tk.StringVar()
        self.species_combobox = ttk.Combobox(self.new_window, textvariable=self.species_var)
        self.species_combobox['values'] = ("Mouse", "Human")
        self.species_combobox.grid(row=1, column=1, padx=10, pady=10)
        # Catalog number input
        ttk.Label(self.new_window, text="Catalog Number:").grid(row=2, column=0, padx=10, pady=10)
        self.catalog_entry = ttk.Entry(self.new_window)
        self.catalog_entry.grid(row=2, column=1, padx=10, pady=10)
        
        # Scrape button
        self.scrape_button = ttk.Button(self.new_window, text="Scrape", command=self.scrape)
        self.scrape_button.grid(row=4, column=0, columnspan=2, pady=10)
    
    def scrape(self):
        company = self.company_var.get()
        catalog_number = self.catalog_entry.get()

        if not company or not catalog_number:
            messagebox.showerror("Input Error", "Please provide both company and catalog number.")
            return
    
        # If the company is one of the dropdown options, open manual entry with pre-filled data
        if company in [ "Other (manual entry)"]:
            prefilled_data = {
                "Supplier": company if company != "Other (manual entry)" else "",
                "Catalog Number": catalog_number
            }
            ManualEntryWindow(self.new_window,prefilled_data,self.species_var.get())
        elif company=="Tonbo":
           newdata=logic.TonboCytek_Scraper(catalog_number)
           prefilled_data = {
               "Supplier": company if company != "Other (manual entry)" else "",
               "Catalog Number": catalog_number}
           prefilled_data.update(newdata)
          # print(prefilled_data)
           ManualEntryWindow(self.new_window, prefilled_data,self.species_var.get())
        elif company=="Biolegend":
            newdata=logic.Biolegend_Scraper(catalog_number)
            prefilled_data = {
                "Supplier": company if company != "Other (manual entry)" else "",
                "Catalog Number": catalog_number}
            prefilled_data.update(newdata)
           # print(prefilled_data)
            ManualEntryWindow(self.new_window, prefilled_data,self.species_var.get())   
        elif company=="BD":
            newdata=logic.BD_scraper(catalog_number)
            prefilled_data = {
                "Supplier": company if company != "Other (manual entry)" else "",
                "Catalog Number": catalog_number}
            prefilled_data.update(newdata)
           # print(prefilled_data)
            ManualEntryWindow(self.new_window, prefilled_data,self.species_var.get())
        elif company=="Thermofisher":
            newdata=logic.Thermofisher_Scraper(catalog_number)
            prefilled_data = {
                "Supplier": company if company != "Other (manual entry)" else "",
                "Catalog Number": catalog_number}
            prefilled_data.update(newdata)
            #print(prefilled_data)
            ManualEntryWindow(self.new_window, prefilled_data,self.species_var.get())
            
        
        else:
            messagebox.showerror("Input Error", "Unknown company selected.")
       # self.new_window.destroy()
#%% Manual Entry
class ManualEntryWindow():

    def __init__(self, root, prefilled_data,species):
        self.species=species
        self.manual_entry_window = tk.Toplevel(root)
        self.manual_entry_window.title("Antibody Entry")
        # Use grid for the label and option menu to maintain consistency
        fillopt_label = ttk.Label(self.manual_entry_window, text="Select Ab Number")
        fillopt_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)
     
        first_missing, missing_numbers=logic.determine_free_inventory(self.species)
       
        
        self.antibody_num_var = tk.StringVar()
        self.antibody_num_var.set(first_missing)
        antibody_nums = missing_numbers
        antibody_num_dropdown = ttk.OptionMenu(self.manual_entry_window, self.antibody_num_var, *antibody_nums)
        antibody_num_dropdown.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

        # Define labels and corresponding data types
        self.labels = ["Antibody Specificity:", "Label:", "Peak Channel (CyTEK):", "Clone:", "Target Species:", "Host Species:", "Supplier:", "Catalog Number:", "Volume:", "Number of Vials:"]
        
        self.entries = {}

        # Create labels and entry fields with optional prefilled data
        for i, label in enumerate(self.labels):
            ttk.Label(self.manual_entry_window, text=label).grid(row=i+2, column=0, padx=10, pady=5)  # Start at row 2 to avoid overlap with previous widgets
            entry = ttk.Entry(self.manual_entry_window)
            entry.grid(row=i+2, column=1, padx=10, pady=5)

            # Insert prefilled data if available
            entry_value = prefilled_data.get(label.strip(':'), "") if prefilled_data else ""
            entry.insert(0, str(entry_value))  # Ensure that entry_value is a string
            
            self.entries[label] = entry
        
        # Submit button placed after all the entry fields
        submit_button = ttk.Button(self.manual_entry_window, text="Submit", command=self.submit_manual_entry)
        submit_button.grid(row=len(self.labels) + 2, column=0, columnspan=2, pady=10)
    
    def submit_manual_entry(self):
        # Collect data from entry fields
        data = {label.strip(':'): entry.get() for label, entry in self.entries.items()}
        data["Inventory Number"]=self.antibody_num_var.get()
        if not "" in data.values(): #make sure every thing is filled out
            try: 
                test=type(int(data["Volume"])) # check format
                test=type(int(data["Number of Vials"])) is int# check to make sure these are numbers 
                if "Mouse" in self.species:
                    test=type(int(data["Inventory Number"])) is int# check to make sure these are numbers 
                    data["Inventory Number"]= int(data["Inventory Number"])
                    
                else:
                    test=type(int(data["Inventory Number"].replace("H",""))) is int# check to make sure these are numbers by removing "H"
                    data["Inventory Number"]=(data["Inventory Number"])
                data["Volume"]= int(data["Volume"])
                data["Number of Vials"]= int(data["Number of Vials"])
            
                if not logic.check_existing_inventory(data,self.species): # check the worksheet for the preexisting column, if not found go ahead and export and give a little message
                    #rearrange the data  to the preferred order
                    fill_data=[data["Inventory Number"],data["Antibody Specificity"],data["Label"], data["Peak Channel (CyTEK)"], data["Clone"], data["Target Species"],data[ "Host Species"], data["Supplier"], data["Catalog Number"], data["Number of Vials"], data["Volume"]]#aka the "fill order" data
                    logic.insert_and_update_row(data["Peak Channel (CyTEK)"],data["Inventory Number"], fill_data,self.species)
                    messagebox.showinfo("Info", f"Manual entry submitted successfully! This antibody's ID is ")
                    
                    root.destroy() #kills the root window "new window" along with the current one
            except Exception as e:
                messagebox.showerror("Input Error", f"Wrong Format Volume or Vial #{e}")
        else: 
            messagebox.showerror("Input Error", "Missing Parameter")
        
      

        


#%% Create panel window
class CreatePanelWindow:
    def __init__(self, root,species):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Create a Panel")

        # Create a frame to hold the buttons and input fields
        self.frame = ttk.Frame(self.new_window)
        self.frame.pack(padx=10, pady=10)

        # User name input
        ttk.Label(self.frame, text="Your Name:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        self.user_name_var = tk.StringVar()
        self.user_name_combobox = ttk.Combobox(self.frame, textvariable=self.user_name_var)
        self.user_name_combobox['values'] = ("Anagha", "Tim", "Hannah")
        self.user_name_combobox.grid(row=0, column=1, padx=10, pady=10)

        # Panel name input
        ttk.Label(self.frame, text="Panel Name:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        self.panel_name_entry = ttk.Entry(self.frame)
        self.panel_name_entry.grid(row=1, column=1, padx=10, pady=10)

        # Add More Antibodies button
        self.add_more_button = ttk.Button(self.frame, text="Add More Antibodies", command=self.add_antibody_input)
        self.add_more_button.grid(row=2, column=0, columnspan=2, pady=10)

         # Submit button
        self.submit_button = ttk.Button(self.frame, text="Create Panel", command=self.create_panel)
        self.submit_button.grid(row=3, column=0, columnspan=2, pady=10)

        # Container for antibody inputs
        self.antibody_inputs_frame = ttk.Frame(self.frame)
        self.antibody_inputs_frame.grid(row=4, column=0, columnspan=2, pady=10)

        # Initialize input fields list
        self.antibody_entries = []

        # Add the first antibody input row
        self.add_antibody_input()

    def add_antibody_input(self):
        """Add a new row of antibody input fields."""
        row = len(self.antibody_entries)  # Determine the next row number
        row_frame = ttk.Frame(self.antibody_inputs_frame)
        row_frame.grid(row=row, column=0, columnspan=4, pady=5, sticky='w')

        ttk.Label(row_frame, text="Antibody Number:").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        antibody_entry = ttk.Entry(row_frame)
        antibody_entry.grid(row=0, column=1, padx=10, pady=5)

        ttk.Label(row_frame, text="Dilution:").grid(row=0, column=2, padx=10, pady=5, sticky='w')
        dilution_entry = ttk.Entry(row_frame)
        dilution_entry.grid(row=0, column=3, padx=10, pady=5)

        # Add an "X" button to remove the row
        remove_button = ttk.Button(row_frame, text="X", command=lambda: self.remove_antibody_input(row_frame))
        remove_button.grid(row=0, column=4, padx=5, pady=5)

        self.antibody_entries.append((antibody_entry, dilution_entry, row_frame))

  #  def remove_antibody_input(self, row_frame):
        #"""Remove an antibody input row."""
    
        #row_frame.grid_forget()  # Hide the row
      #  row_frame.destroy()  # Destroy the frame
       # self.antibody_entries = [entry for entry in self.antibody_entries if entry[2] != row_frame]  # Remove from list
    def remove_antibody_input(self, row_frame):
        """Remove an antibody input row and re-index the remaining ones."""
        row_frame.grid_forget()  # Hide the row
        row_frame.destroy()  # Destroy the frame
    
        # Remove the entry from the list
        self.antibody_entries = [entry for entry in self.antibody_entries if entry[2] != row_frame]
    
        # Re-index the remaining rows
        for idx, (antibody_entry, dilution_entry, frame) in enumerate(self.antibody_entries):
            frame.grid(row=idx, column=0, columnspan=4, pady=5, sticky='w')

    def create_panel(self):
        """Create a new panel with the provided antibody information."""
        user_name = self.user_name_var.get()
        panel_name = self.panel_name_entry.get()

        if not user_name or not panel_name:
            messagebox.showerror("Input Error", "Please provide both your name and the panel name.")
            return

        panel_data = []

        for antibody_entry, dilution_entry, _ in self.antibody_entries:
            antibody_number = antibody_entry.get()
            dilution = dilution_entry.get()
            print(antibody_entry.get())
            if antibody_number and dilution:
                # Retrieve antibody details from the DataFrame
                if int(antibody_number) in logic.antibody_df['Antibody Number'].values:
                    antibody_details = logic.antibody_df[logic.antibody_df['Antibody Number'] == int(antibody_number)]
                    antibody_specificity = antibody_details['Antibody Specificity'].values[0]
                    label = antibody_details['Label'].values[0]

                    panel_data.append({
                        'Antibody Number': antibody_number,
                        'Antibody Specificity': antibody_specificity,
                        'Label': label,
                        'Dilution': dilution
                    })

        if not panel_data:
            messagebox.showerror("Input Error", "No valid antibodies or dilutions provided.")
            return

        # Convert to DataFrame and save to Excel
        panel_df = pd.DataFrame(panel_data)
        user_folder = f"Panels/{user_name}"
        os.makedirs(user_folder, exist_ok=True)
        file_path = os.path.join(user_folder, f"{panel_name}.xlsx")
        panel_df.to_excel(file_path, index=False)

        messagebox.showinfo("Success", f"Panel created and saved as {file_path}")
        self.new_window.destroy()
#%% Modify a panel
class ModifyPanelWindow:
    def __init__(self, root):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Modify a Panel")
        ttk.Label(self.new_window, text="Panel modification functionality to be implemented").grid(row=0, column=0, padx=10, pady=10)

#%% Run a Panel



#%% Use antibody window
class UseAntibodyWindow:
    def __init__(self, root):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Use a Single Antibody")
        tk.Label(self.new_window, text="Species:").grid(row=0, column=0, padx=10, pady=10)
        self.species_var = tk.StringVar()
        self.species_combobox = ttk.Combobox(self.new_window, textvariable=self.species_var)
        self.species_combobox['values'] = ("Mouse", "Human")
        self.species_combobox.grid(row=0, column=1, padx=10, pady=10)
        # Antibody number input
        
        ttk.Label(self.new_window, text="Antibody Number:").grid(row=1, column=0, padx=10, pady=10)
        self.antibody_number_entry = ttk.Entry(self.new_window)
        self.antibody_number_entry.grid(row=1, column=1, padx=10, pady=10)
        
        # Volume used input
        ttk.Label(self.new_window, text="Volume Used:").grid(row=2, column=0, padx=10, pady=10)
        self.volume_used_entry = ttk.Entry(self.new_window)
        self.volume_used_entry.grid(row=2, column=1, padx=10, pady=10)
         
        # Submit button
        submit_button = ttk.Button(self.new_window, text="Submit", command=self.submit_usage)
        submit_button.grid(row=3, column=0, columnspan=2, pady=10)
    
    def submit_usage(self):
        # Get user input
        antibody_number = self.antibody_number_entry.get()
        volume_used = self.volume_used_entry.get()
    
        if not antibody_number or not volume_used:
            messagebox.showerror("Input Error", "Please provide both an antibody number and the volume used.")
            return
    
        # Validate inputs
        try:
            antibody_number = int(antibody_number)
            volume_used = float(volume_used)
        except ValueError:
            messagebox.showerror("Input Error", "Antibody number must be an integer, and volume used must be a number.")
            return
    
        # Update the volume in the DataFrame
        sheetlink.update_value(self.species_var.get(), antibody_number, volume_used)
    
    
        # Close the usage window
        self.new_window.destroy()


#%% Log an Issue
class LogIssueWindow:
    def __init__(self, root):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Log Issue")
        
        # Antibody number input
        ttk.Label(self.new_window, text="Antibody Number:").grid(row=0, column=0, padx=10, pady=10)
        self.antibody_number_entry = ttk.Entry(self.new_window)
        self.antibody_number_entry.grid(row=0, column=1, padx=10, pady=10)

        # Error message input
        ttk.Label(self.new_window, text="Error Message:").grid(row=1, column=0, padx=10, pady=10)
        self.error_message_entry = ttk.Entry(self.new_window)
        self.error_message_entry.grid(row=1, column=1, padx=10, pady=10)

        # Submit button
        submit_button = ttk.Button(self.new_window, text="Submit", command=self.submit_issue)
        submit_button.grid(row=2, column=0, columnspan=2, pady=10)
    
    def submit_issue(self):
        # Get user input
        antibody_number = self.antibody_number_entry.get()
        error_message = self.error_message_entry.get()
        
        if not antibody_number or not error_message:
            messagebox.showerror("Input Error", "Please provide both an antibody number and an error message.")
            return
        
        # Convert antibody_number to int and find the corresponding row in the DataFrame
        try:
            antibody_number = int(antibody_number)
        except ValueError:
            messagebox.showerror("Input Error", "Antibody number must be a valid integer.")
            return
        
        # Update the corresponding antibody's row with the error message
        success = logic.log_antibody_issue(antibody_number, error_message)
        
        if success:
            messagebox.showinfo("Success", "Error message logged successfully!")
        else:
            messagebox.showerror("Error", "No antibody found with the given number.")
        
        # Close the issue logging window
        self.new_window.destroy()





#%% Splash page


class LoadingWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("500x400")
        self.title("Loading...")
        self.configure(bg="white")

        # Load and display the image at the top (use the actual path to your image)
        self.image_path = "/Users/westtn/Desktop/Screenshot 2024-09-19 at 4.04.03 PM.png"
        self.load_image()

        # Create and pack a fake loading bar
        self.loading_label = tk.Label(self, text="Loading...", font=("Arial", 14), bg="white")
        self.loading_label.pack(pady=20)

        self.progress_canvas = tk.Canvas(self, width=400, height=30, bg="white", bd=0, highlightthickness=0)
        self.progress_canvas.pack(pady=10)

        # Simulate a loading bar
        self.progress_bar = self.progress_canvas.create_rectangle(2, 2, 2, 28, fill="green")

        self.fake_loading()

    def load_image(self):
        # Load the image
        image = Image.open(self.image_path)
        image = image.resize((500, 200), Image.Resampling.LANCZOS)  # Resize the image to fit the window width
        self.image_photo = ImageTk.PhotoImage(image)

        # Display the image in a label
        image_label = tk.Label(self, image=self.image_photo, bg="white")
        image_label.pack(pady=10)

    def fake_loading(self):
        # Fake loading animation
        current_width = 2
        max_width = 398

        def update_loading_bar():
            nonlocal current_width
            if current_width < max_width:
                current_width += 4  # Increment the loading bar width
                self.progress_canvas.coords(self.progress_bar, 2, 2, current_width, 28)
                self.after(50, update_loading_bar)  # Continue the animation every 50 milliseconds
            else:
                self.loading_label.config(text="Welcome!")  # Update text when done

        update_loading_bar()


#%% Welcome page
class WelcomeWindow:
    def __init__(self, root):
        self.root = root  # Assign the root window
        self.root.geometry("705x460")  # Increased window height to accommodate layout
        self.root.title("Welcome")
        
        # Load the banner image at the top (use an actual image file path)
        #self.banner_image_path = "/Users/westtn/Desktop/Screenshot 2024-09-20 at 1.51.27 PM.png"
        self.banner_image_path = '/Users/westtn/Desktop/Screenshot 2024-10-09 at 6.06.21 PM.png'
        self.load_banner_image()

        # Place the buttons in the center (existing functionality)
        self.create_button_frames()

        # Create the browse button at the bottom left
        self.create_browse_button()

        # Create round settings and alerts buttons at the bottom right
        self.create_round_buttons()
        #self.check_alerts()
    def load_banner_image(self):
    # Load banner image (use a path to your image file)
        banner_image = Image.open(self.banner_image_path)
        banner_image = banner_image.resize((705, 85), Image.Resampling.LANCZOS)  # Use LANCZOS instead of ANTIALIAS
        self.banner_photo = ImageTk.PhotoImage(banner_image)

    # Create a label to display the image
        banner_label = Label(self.root, image=self.banner_photo)
        banner_label.pack(pady=0)
    
#ttk.Button(self.frame, text="Add an Antibody", command=self.open_add_antibody_window).pack(padx=10, pady=10)
#ttk.Button(self.frame, text="Create a Panel", command=self.open_create_panel_window).pack(padx=10, pady=10)
#ttk.Button(self.frame, text="Modify a Panel", command=self.open_modify_panel_window).pack(padx=10, pady=10)
#ttk.Button(self.frame, text="Run a Panel", command=self.open_run_panel_window).pack(padx=10, pady=10)
#ttk.Button(self.frame, text="Use Single Antibody", command=self.open_use_antibody_window).pack(padx=10, pady=10)
#ttk.Button(self.frame, text="Log Antibody Issue", command=self.open_log_issue_window).pack(padx=10, pady=10)
    def create_button_frames(self):
        # Button 1 Info
        button1_frame = tk.Frame(root)
        button1_frame.pack(padx=10, pady=1, anchor="w")

        button1 = tk.Button(button1_frame, text="Add an Antibody", command=self.open_add_window)
        button1.pack(side="left", padx=5+10)

        label1_frame = tk.LabelFrame(button1_frame, text="Add Antibodies to Inventory")
        label1_frame.pack(side="left", padx=5)

        label1 = tk.Label(label1_frame, text="Add Antibodies to the Inventory based on availible lot numbers.                               ")
        label1.pack(padx=1, pady=1)

        # Button 2 Info
        button2_frame = tk.Frame(root)
        button2_frame.pack(padx=10, pady=1, anchor="w")

        button2 = tk.Button(button2_frame, text="Create a Panel", command=self.open_create_panel_window)
        button2.pack(side="left", padx=11+10)

        label2_frame = tk.LabelFrame(button2_frame, text="Compose A Panel Sheet")
        label2_frame.pack(side="left", padx=5)

        label2_text = "Compose a stain panel using the automated layout. Requires Wifi Connection.       "
        label2 = tk.Label(label2_frame, text=label2_text)
        label2.pack(padx=1, pady=1)

        # Button 3 Info
        button3_frame = tk.Frame(root)
        button3_frame.pack(padx=10, pady=1, anchor="w")

        button3 = tk.Button(button3_frame, text="Modify a Panel", command=self.open_modify_panel_window)
        button3.pack(side="left", padx=11+10)

        label3_frame = tk.LabelFrame(button3_frame, text="Modify a Antibody Panel")
        label3_frame.pack(side="left", padx=5)

        label3_text = "Edit an existing antibody panel. Requires Wifi.                                                             "
        label3 = tk.Label(label3_frame, text=label3_text)
        label3.pack(padx=1, pady=1)
        # Button 4 Info--------------------------------------
        button4_frame = tk.Frame(root)
        button4_frame.pack(padx=10, pady=1, anchor="w")
        
        button4 = tk.Button(button4_frame, text="Run a Panel", command=self.open_run_panel_window)
        button4.pack(side="left", padx=20+10)

        label4_frame = tk.LabelFrame(button4_frame, text="Run a Antibody Panel")
        label4_frame.pack(side="left", padx=5)

        label4_text = "Run an existing Antibody panel via opentrons. Requires Terminal/SSH connection."
        label4 = tk.Label(label4_frame, text=label4_text)
        label4.pack(padx=1, pady=1)
        # Button 5 Info---------------------------------------------
        button5_frame = tk.Frame(root)
        button5_frame.pack(padx=28, pady=1, anchor="w")
        
        button5 = tk.Button(button5_frame, text="Aliquot to Tube", command=self.open_log_issue_window)
        button5.pack(side="left", padx=7)

        label5_frame = tk.LabelFrame(button5_frame, text="Load a Rack from Tubes")
        label5_frame.pack(side="left", padx=12)

        label5_text = "Add your panel to a 96 Rack. Requires Wifi.                                                "
        label5 = tk.Label(label5_frame, text=label5_text)
        label5.pack(padx=1, pady=1)
        # Button 6 Info---------------------------------------------------
        button6_frame = tk.Frame(root)
        button6_frame.pack(padx=10, pady=1, anchor="w")
        
        button6 = tk.Button(button6_frame, text="Use Single Antibody", command=self.open_use_antibody_window)
        button6.pack(side="left", padx=5)

        label6_frame = tk.LabelFrame(button6_frame, text="Edit Single Volumes")
        label6_frame.pack(side="left", padx=5)

        label6_text = "Edit/Locate volume on a single antibody for seperate/non-automated use.              "
        label6 = tk.Label(label6_frame, text=label6_text)
        label6.pack(padx=1, pady=1)
        # Button 7 Info----------------------------------
        button7_frame = tk.Frame(root)
        button7_frame.pack(padx=11, pady=1, anchor="w")
        
        button7 = tk.Button(button7_frame, text="Log Antibody Issue", command=self.open_log_issue_window)
        button7.pack(side="left", padx=7)

        label7_frame = tk.LabelFrame(button7_frame, text="Add Notes to Antibodies")
        label7_frame.pack(side="left", padx=5)

        label7_text = "Add a non-usage note to any antibody. Requires Wifi.                                                "
        label7 = tk.Label(label7_frame, text=label7_text)
        label7.pack(padx=1, pady=1)

    def create_browse_button(self):
        browse_button = tk.Button(root, text="Browse", command=self.browse_files)
        browse_button.place(x=20, y=420)  # Place the browse button at the bottom left

    def create_round_buttons(self):
        # Create the round settings button
        settings_button = tk.Button(root, text="⚙️", command=self.open_settings, height=2, width=4, font=("Arial", 10))
        settings_button.place(x=630, y=420)  # Place the settings button at the bottom right

        # Create the round alerts button next to the settings button
        alerts_button = tk.Button(root, text="🔔", command=self.show_alerts, fg='red' ,height=2, width=4, font=("Arial", 10))
        alerts_button.config(bg="red", fg="red")
        alerts_button.place(x=570, y=420)
        print("""Check for alerts and update the Alerts button color.""")
        try:
            if sheetlink.check_alerts():  # Check if alerts exist
                self.alerts=sheetlink.check_alerts()
                print("hi")
                alerts_button.config(highlightbackground='red', fg="red")
            else:
                alerts_button.config(bg="SystemButtonFace", fg="black")  # Default button color
        except Exception as e:
            print(f"Error checking alerts: {e}")
          # Place the alerts button next to the settings button
    # Placeholder functions for button commands
    def open_add_window(self):
        
         AddAntibodyWindow(self.root)  
    def open_create_panel_window(self):
        #PanelCreatorApp(self.root)
       destination="Panel"
       InterimAntibodyWindow(self.root,destination)

    def open_modify_panel_window(self):
        PanelEditorApp(self.root,"Mouse")  # Pass the correct window as the master
        print("passed")
    def open_run_panel_window(self):
        RunPanelWindow(self.root)

    def open_use_antibody_window(self):
        UseAntibodyWindow(self.root)

    def open_log_issue_window(self):
        #LogIssueWindow(self.root)
        logic.determine_free_inventory()
    def open_settings(self):
        print("Settings button clicked")
        SettingsWindow(self.root)
    def show_alerts(self):
        print("Alerts button clicked")
        AlertsWindow(self.root,self.alerts)
        #self.create_round_buttons()
    def browse_files(self):
        BrowseWindow(self.root)
    def check_alerts(self):
         v=1
#%% New Panel Creator ---- needs  modifier and  mouse vs human interim page
class PanelCreatorApp:
    def __init__(self, root,species):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Google Spreadsheet Treeview with Panel Creation")
        self.species=species
        # Upper frame for user input
        self.upper_frame = ttk.Frame(self.new_window)
        self.upper_frame.pack(padx=10, pady=10, fill="x")

        # Dropdown for name selection
        ttk.Label(self.upper_frame, text="Select Your Name:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.name_var = tk.StringVar()
        self.name_combobox = ttk.Combobox(self.upper_frame, textvariable=self.name_var)
        self.name_combobox['values'] = sheetlink.fetch_authors()
        self.name_combobox.grid(row=0, column=1, padx=5, pady=5)

        # Field for panel name input
        ttk.Label(self.upper_frame, text="Panel Name:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.panel_name_entry = ttk.Entry(self.upper_frame)
        self.panel_name_entry.grid(row=0, column=3, padx=5, pady=5)

        # Field for total volume input
        ttk.Label(self.upper_frame, text="Total Volume:").grid(row=0, column=4, padx=5, pady=5, sticky='w')
        self.total_volume_entry = ttk.Entry(self.upper_frame)
        self.total_volume_entry.grid(row=0, column=5, padx=5, pady=5)

        # Load spreadsheet data
        
        self.sheet_data = sheetlink.fetch_inventory(self.species)
        #print(self.sheet_data)
        # Mid section for treeviews and buttons
        self.middle_frame = ttk.Frame(self.new_window)
        self.middle_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Create left treeview (with dilution column)
        self.left_treeview = self.create_treeview(self.middle_frame, include_dilution=True)
        self.left_treeview.grid(row=0, column=0, padx=5, pady=5)

        # Create right treeview (without dilution column)
        self.right_treeview = self.create_treeview(self.middle_frame, include_dilution=False)
        self.right_treeview.grid(row=0, column=2, padx=5, pady=5)

        # Add buttons between treeviews
        self.button_frame = ttk.Frame(self.middle_frame)
        self.button_frame.grid(row=0, column=1, padx=5, pady=5, sticky='n')

        self.add_button = ttk.Button(self.button_frame, text="<", command=self.move_to_left)
        self.add_button.pack(pady=5)

        self.remove_button = ttk.Button(self.button_frame, text=">", command=self.move_to_right)
        self.remove_button.pack(pady=5)

        # Create Panel button at the bottom
        self.create_panel_button = ttk.Button(self.new_window, text="Create Panel", command=self.create_panel)
        self.create_panel_button.pack(pady=10)

        # Load right treeview with spreadsheet data
        self.load_right_treeview()

        # Bind left treeview click for editing
        self.left_treeview.bind("<Double-1>", self.on_double_click_left)

    def create_treeview(self, parent, include_dilution=True):
        """Helper function to create a treeview. Optionally includes a 'Dilution' column."""
        if include_dilution:
            columns = ("Inventory Number", "Antibody Specificity","Label" ,"Peak Channel (CyTEK)", "Dilution")
        else:
            columns = ("Inventory Number", "Antibody Specificity","Label", "Peak Channel (CyTEK)")

        tree = ttk.Treeview(parent, columns=columns, show='headings', selectmode='extended')

        # Define column headings
        tree.heading("Inventory Number", text="Inventory Number")
        tree.heading("Antibody Specificity", text="Antibody Specificity")
        tree.heading("Label", text="Label")
        tree.heading("Peak Channel (CyTEK)", text="Peak Channel")

        if include_dilution:
            tree.heading("Dilution", text="Dilution")

        # Set column widths
        tree.column("Inventory Number", width=120)
        tree.column("Antibody Specificity", width=150)
        tree.column("Label", width=150)
        tree.column("Peak Channel (CyTEK)", width=100)

        if include_dilution:
            tree.column("Dilution", width=80)

        return tree

    def load_right_treeview(self):
        """Load the right treeview with spreadsheet data (without dilution column) and filter non-numeric Inventory Numbers."""
        for row in self.sheet_data[1:]:  # Skip the header row
            inventory_number = row[0]
            print(inventory_number)
            # Filter out non-numeric inventory numbers
            try:
                #int(inventory_number)  # Check if the value is numeric
                antibody_specificity = row[1]
                peak_channel = row[3]
                label=row[2]
                self.right_treeview.insert("", "end", values=(inventory_number, antibody_specificity, label ,peak_channel))
            except ValueError:
                # Skip non-numeric inventory numbers
                
                continue
    
        # Sort the right treeview after loading
        self.sort_treeview(self.right_treeview)
    
    def move_to_left(self):
        """Move selected rows from the right treeview to the left treeview and add editable dilution."""
        selected_items = self.right_treeview.selection()  # Get all selected items
        for item in selected_items:
            item_values = self.right_treeview.item(item, "values")
            
            # Move only if the Inventory Number is numeric or meets your condition
            try:
                self.left_treeview.insert("", "end", values=item_values + ("",))  # Add to left treeview with empty dilution
                self.right_treeview.delete(item)  # Remove from right treeview
            except ValueError:
                pass  # Skip if non-numeric, depending on your requirements
    
        # Sort the left treeview after all items have been moved
        self.sort_treeview(self.left_treeview)



    def move_to_right(self):
        """Move selected rows from the left treeview back to the right treeview."""
        selected_items = self.left_treeview.selection()  # Get all selected items
        for item in selected_items:
            item_values = self.left_treeview.item(item, "values")
            self.right_treeview.insert("", "end", values=item_values[:4])  # Add to right treeview without dilution
            self.left_treeview.delete(item)  # Remove from left treeview
    
        # Sort the right treeview after all items have been moved
        self.sort_treeview(self.right_treeview)


    def sort_treeview(self, treeview):
        """Sort the treeview by the first column (Inventory Number), handling mixed types."""
        items = [(treeview.item(item)["values"], item) for item in treeview.get_children()]
    
        # Define a function to extract numeric and non-numeric parts for sorting
        def sort_key(value):
            first_col_value = value[0][0]  # Extract the first column value (Inventory Number)
            
            # Use regex to split the string into numeric and non-numeric parts
            match = re.match(r"(\d+)(\D*)", str(first_col_value))
            if match:
                numeric_part = int(match.group(1))  # Convert numeric part to integer for sorting
                non_numeric_part = match.group(2)   # Keep the non-numeric part as a string
                return (numeric_part, non_numeric_part)
            else:
                # If the value does not match the pattern, return it as a string for lexicographical sorting
                return (float('inf'), str(first_col_value))
    
        # Sort the items using the sort_key
        sorted_items = sorted(items, key=sort_key)
    
        # Rearrange treeview items in sorted order
        for index, (_, item) in enumerate(sorted_items):
            treeview.move(item, "", index)

    def on_double_click_left(self, event):
        """Handle double-clicking to edit the dilution field in the left treeview."""
        region = self.left_treeview.identify_region(event.x, event.y)
        if region == 'cell':  # Check if clicking on a cell
            column = self.left_treeview.identify_column(event.x)
            if column == '#5':  # Check if the Dilution column was clicked
                row_id = self.left_treeview.identify_row(event.y)
                if row_id:
                    # Get the current value of the cell
                    current_value = self.left_treeview.item(row_id, 'values')[4]
                    # Create an entry widget to edit the value
                    entry = ttk.Entry(self.left_treeview)
                    entry.insert(0, current_value)
                    entry.focus()

                    # Position the entry widget on top of the clicked cell
                    bbox = self.left_treeview.bbox(row_id, column)
                    entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])

                    # Bind focus out or return key to update the cell value
                    entry.bind("<FocusOut>", lambda e: self.update_dilution(row_id, entry))
                    entry.bind("<Return>", lambda e: self.update_dilution(row_id, entry))

    def update_dilution(self, row_id, entry):
        """Update the dilution value in the treeview from the entry."""
        new_value = entry.get()
        # Update the dilution in the treeview
        current_values = list(self.left_treeview.item(row_id, 'values'))
        current_values[4] = new_value #index start at zero for the column
        self.left_treeview.item(row_id, values=current_values)
        entry.destroy()  # Remove the entry widget after updating
    def create_panel(self):
        """Create a panel, export to Google Sheets, and add a new worksheet with treeview data and total volume."""
    
        # Get the name, panel name, and total volume from the input fields
        user_name = self.name_var.get()
        panel_name = self.panel_name_entry.get()
        total_volume = self.total_volume_entry.get()
    
        # Combine the name and panel name for the new sheet
        new_sheet_name = f"Panel:{user_name}_{panel_name}"
    
        # Collect data from the left treeview (antibody information)
        panel_data = []
        headers = ["Inventory Number", "Antigen", "Dyes", "(Cytek) Peak Channel", "Dilution", "Volume (uL)"]

        for row_id in self.left_treeview.get_children():
            row = self.left_treeview.item(row_id, "values")
            inventory_number, antigen, dyes, peak_channel, dilution = row

        # Calculate the volume (total volume divided by dilution)
            try:
                dilution = float(dilution)  # Ensure dilution is a valid number
                volume_formula = f"={total_volume}/{dilution}" if dilution else ""
            except ValueError:
                tk.messagebox.showerror("Error", f"Invalid dilution value for {antigen}.")
                return
            panel_data.append([inventory_number, antigen, dyes, peak_channel, dilution,volume_formula])
        # Add Total Volume at the bottom right
        panel_data.append(["", "" ,"","",f"Total Volume:",f"{total_volume}"])
    
        print(panel_data)
        # Add a new worksheet with the generated sheet name
        sheetlink.export_panel(new_sheet_name,headers,panel_data)
    

#%%New Panel Modifier


class PanelEditorApp:
    def __init__(self, root, species):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Edit Existing Panel")
        self.species = species
        print("species="+species)

        # Upper frame for spreadsheet selection
        self.upper_frame = ttk.Frame(self.new_window)
        self.upper_frame.pack(padx=10, pady=10, fill="x")

        # Dropdown for spreadsheet selection
        ttk.Label(self.upper_frame, text="Select Panel Spreadsheet:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.spreadsheet_var = tk.StringVar()
        self.spreadsheet_combobox = ttk.Combobox(self.upper_frame, textvariable=self.spreadsheet_var)
        self.spreadsheet_combobox.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(self.upper_frame, text="Total Volume:").grid(row=0, column=4, padx=5, pady=5, sticky='w')
        self.total_volume_entry = ttk.Entry(self.upper_frame)
        self.total_volume_entry.grid(row=0, column=5, padx=5, pady=5)

        # Load spreadsheet names starting with "Panel:"
        self.load_panel_names()

        # Mid section for treeviews and buttons
        self.middle_frame = ttk.Frame(self.new_window)
        self.middle_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Create left treeview (with dilution column)
        self.left_treeview = self.create_treeview(self.middle_frame, include_dilution=True)
        self.left_treeview.grid(row=0, column=0, padx=5, pady=5)

        # Create right treeview (without dilution column)
        self.right_treeview = self.create_treeview(self.middle_frame, include_dilution=False)
        self.right_treeview.grid(row=0, column=2, padx=5, pady=5)

        # Add buttons between treeviews
        self.button_frame = ttk.Frame(self.middle_frame)
        self.button_frame.grid(row=0, column=1, padx=5, pady=5, sticky='n')

        self.add_button = ttk.Button(self.button_frame, text="<", command=self.move_to_left)
        self.add_button.pack(pady=5)

        self.remove_button = ttk.Button(self.button_frame, text=">", command=self.move_to_right)
        self.remove_button.pack(pady=5)

        # Submit Changes button at the bottom
        self.submit_changes_button = ttk.Button(self.new_window, text="Submit Changes", command=self.submit_changes)
        self.submit_changes_button.pack(pady=10)

        # Bind the dropdown selection event to load data
        self.spreadsheet_combobox.bind("<<ComboboxSelected>>", self.load_panel_data)
        self.left_treeview.bind("<Double-1>", self.on_double_click_left)

    def load_panel_names(self):
        
        """Fetch and display spreadsheet names that start with 'Panel:'."""
        panel_names = sheetlink.fetch_panel_names()  # Assuming this function returns a list of panel names
        self.spreadsheet_combobox['values'] = panel_names

    def load_panel_data(self, event=None):
        
        """Load the selected panel data into the left treeview and filter data for the right treeview."""
        
        selected_panel = self.spreadsheet_var.get()
       
        if not selected_panel:
            return
    
        # Load the data for the selected panel into the left treeview
        panel_data = sheetlink.fetch_panel_data(selected_panel)  # Assuming this returns panel data as a list of rows
        self.left_treeview.delete(*self.left_treeview.get_children())  # Clear existing entries
        
        for row in panel_data:
            self.left_treeview.insert("", "end", values=row)
    
        # Load the data for the right treeview, excluding those already in the left treeview
        all_data = sheetlink.fetch_inventory(self.species)  # Assuming this returns the entire inventory based on species
        self.update_species()
        left_treeview_items = {str(self.left_treeview.item(item)["values"][0]) for item in self.left_treeview.get_children()}
        self.right_treeview.delete(*self.right_treeview.get_children())  # Clear existing entries
        for row in all_data[1:]:  # Skip the header row
            if row[0] not in left_treeview_items:
                self.right_treeview.insert("", "end", values=(row[0], row[1], row[2], row[3]))
    def update_species(self):
        """Update the species and right treeview based on the contents of the left treeview."""
        # Check if any entry in the left treeview contains "H" in the Antibody Specificity column (index 1)
        print([str(self.left_treeview.item(item)["values"][0]) for item in self.left_treeview.get_children()])
        contains_human = any("H" in str(self.left_treeview.item(item)["values"][0]) for item in self.left_treeview.get_children())
            
        # Update the species based on the presence of "H" entries
        self.species = "Human" if contains_human else "Mouse"
    
    def create_treeview(self, parent, include_dilution=True):
        
        """Helper function to create a treeview with optional dilution column."""
        columns = ("Inventory Number", "Antibody Specificity", "Label", "Peak Channel (CyTEK)")
        if include_dilution:
            columns += ("Dilution",)

        tree = ttk.Treeview(parent, columns=columns, show='headings', selectmode='extended')
        for column in columns:
            tree.heading(column, text=column)
            tree.column(column, width=120 if column == "Inventory Number" else 150)

        return tree

    def move_to_left(self):
        
        """Move selected rows from the right treeview to the left treeview and add editable dilution."""
        selected_items = self.right_treeview.selection()
        for item in selected_items:
            item_values = self.right_treeview.item(item, "values")
            self.left_treeview.insert("", "end", values=item_values + ("",))
            self.right_treeview.delete(item)
        self.sort_treeview(self.left_treeview)
      
    def move_to_right(self):
        
        """Move selected rows from the left treeview back to the right treeview."""
        selected_items = self.left_treeview.selection()
        for item in selected_items:
            item_values = self.left_treeview.item(item, "values")
            self.right_treeview.insert("", "end", values=item_values[:4])  # Exclude the Dilution column
            self.left_treeview.delete(item)
        self.sort_treeview(self.right_treeview)
    def on_double_click_left(self, event):
        
        """Handle double-clicking to edit the dilution field in the left treeview."""
        region = self.left_treeview.identify_region(event.x, event.y)
        if region == 'cell':  # Check if clicking on a cell
            column = self.left_treeview.identify_column(event.x)
            if column == '#5':  # Check if the Dilution column was clicked
                row_id = self.left_treeview.identify_row(event.y)
                if row_id:
                    # Get the current value of the cell
                    current_value = self.left_treeview.item(row_id, 'values')[4]
                    # Create an entry widget to edit the value
                    entry = ttk.Entry(self.left_treeview)
                    entry.insert(0, current_value)
                    entry.focus()

                    # Position the entry widget on top of the clicked cell
                    bbox = self.left_treeview.bbox(row_id, column)
                    entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])

                    # Bind focus out or return key to update the cell value
                    entry.bind("<FocusOut>", lambda e: self.update_dilution(row_id, entry))
                    entry.bind("<Return>", lambda e: self.update_dilution(row_id, entry))

    def update_dilution(self, row_id, entry):
       
        """Update the dilution value in the treeview from the entry."""
        new_value = entry.get()
        # Update the dilution in the treeview
        current_values = list(self.left_treeview.item(row_id, 'values'))
        current_values[4] = new_value #index start at zero for the column
        self.left_treeview.item(row_id, values=current_values)
        entry.destroy()  # Remove the entry widget after updating
    def submit_changes(self):
        
        
        total_volume = self.total_volume_entry.get()
    
        """Replace the existing source sheet with the updated left treeview data."""
        selected_panel = self.spreadsheet_var.get()
        if not selected_panel:
            tk.messagebox.showerror("No Panel Selected", "Please select a panel to submit changes.")
            return

        # Collect data from the left treeview
        updated_panel_data = []
        headers = ["Inventory Number", "Antibody Specificity", "Label", "Peak Channel (CyTEK)", "Dilution"]

        for row_id in self.left_treeview.get_children():
            row = self.left_treeview.item(row_id, "values")
            print(row)
            inventory_number, antigen, dyes, peak_channel, dilution = row[:-1]

        # Calculate the volume (total volume divided by dilution)
            try:
                dilution = float(dilution)  # Ensure dilution is a valid number
                volume_formula = f"={total_volume}/{dilution}" if dilution else ""
            except ValueError:
                tk.messagebox.showerror("Error", f"Invalid dilution value for {antigen}.")
                return
            updated_panel_data.append([inventory_number, antigen, dyes, peak_channel, dilution,volume_formula])
        # Add Total Volume at the bottom right
        updated_panel_data.append(["", "" ,"","",f"Total Volume:",f"{total_volume}"])
        # Replace the existing sheet data with the updated data
        # Assuming this function deletes the existing sheet
        sheetlink.update_panel(selected_panel, headers, updated_panel_data)  # Recreate the sheet with updated data

        tk.messagebox.showinfo("Success", f"Panel '{selected_panel}' has been successfully updated.")
    
    def sort_treeview(self, treeview):
        
        """Sort the treeview by the first column (Inventory Number), handling mixed types."""
        items = [(treeview.item(item)["values"], item) for item in treeview.get_children()]
    
        # Define a function to extract numeric and non-numeric parts for sorting
        def sort_key(value):
            first_col_value = value[0][0]  # Extract the first column value (Inventory Number)
            
            # Use regex to split the string into numeric and non-numeric parts
            match = re.match(r"(\d+)(\D*)", str(first_col_value))
            if match:
                numeric_part = int(match.group(1))  # Convert numeric part to integer for sorting
                non_numeric_part = match.group(2)   # Keep the non-numeric part as a string
                return (numeric_part, non_numeric_part)
            else:
                # If the value does not match the pattern, return it as a string for lexicographical sorting
                return (float('inf'), str(first_col_value))
    
        # Sort the items using the sort_key
        sorted_items = sorted(items, key=sort_key)
    
        # Rearrange treeview items in sorted order
        for index, (_, item) in enumerate(sorted_items):
            treeview.move(item, "", index)


#%% Browse page
class BrowseWindow:
    def __init__(self, root):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        sheetname_file_path = os.path.join(script_dir, "sheet_name.txt")
        with open(sheetname_file_path, "r") as file:
            extension = file.read().strip()  # Remove any extra whitespace or newlines
        # Create a new window as a child of root
        webbrowser.open(f"https://docs.google.com/spreadsheets/d/{extension}")
#%%Alerts

class AlertsWindow:
    def __init__(self, root,alerts):
        # Create a new window as a child of root
        self.Alerts_window = tk.Toplevel(root)
        self.Alerts_window.title("Alerts Management")
        self.Alerts_window.geometry("500x400")
        self.listbox = tk.Listbox(self.Alerts_window)
        
        # Add the strings to the Listbox
        for item in alerts:
            self.listbox.insert(tk.END, item)
        
        # Pack the Listbox widget
        self.listbox.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)
        self.go_back_button = tk.Button(self.Alerts_window, text="Back", command=self.close_window)
        self.go_back_button.pack(pady=10)

    # Method to close the Alerts window
    def close_window(self):
        self.Alerts_window.destroy()
#%%Settings page

class SettingsWindow:
    def __init__(self, root):
        # Create a new window as a child of root
        self.Settings_window = tk.Toplevel(root)
        self.Settings_window.title("Credential & Panel Author Management")
        self.Settings_window.geometry("500x450")

        # Initialize file paths
        self.credential_file = "credential_location.txt"
        self.protocol_file = "protocol_location.txt"
        self.sheet_name_file = "sheet_name.txt"
        self.authors = self.load_authors()

        # Instruction label
        tk.Label(self.Settings_window, text="Manage Credentials and Panel Authors", font=("Arial", 14)).pack(pady=10)

        # Credential file path selection
        credential_button = tk.Button(self.Settings_window, text="Select Credential File", command=self.select_credential_file)
        credential_button.pack(pady=5)
        
        # Protocol file path selection
        protocol_button = tk.Button(self.Settings_window, text="Select Protocol File", command=self.select_protocol_file)
        protocol_button.pack(pady=5)
        
        # Sheet name designation
        tk.Label(self.Settings_window, text="Designate Sheet Name:", font=("Arial", 12)).pack(pady=5)
        self.sheet_name_var = tk.StringVar()
        self.sheet_name_entry = tk.Entry(self.Settings_window, textvariable=self.sheet_name_var)
        self.sheet_name_entry.pack(pady=5)
        designate_button = tk.Button(self.Settings_window, text="Designate Sheet", command=self.designate_sheet)
        designate_button.pack(pady=5)

        # Author management section
        tk.Label(self.Settings_window, text="Manage Panel Authors:", font=("Arial", 12)).pack(pady=10)

        # ComboBox for authors
        self.author_var = tk.StringVar()
        self.author_combobox = ttk.Combobox(self.Settings_window, textvariable=self.author_var, values=self.authors)
        self.author_combobox.pack(pady=5)
        self.author_combobox.set("")

        # Buttons for adding and deleting authors
        button_frame = tk.Frame(self.Settings_window)
        button_frame.pack(pady=10)
        
        add_button = tk.Button(button_frame, text="Add Author", command=self.add_author)
        add_button.pack(side="left", padx=10)
        delete_button = tk.Button(button_frame, text="Delete Author", command=self.delete_author)
        delete_button.pack(side="left", padx=10)
        
        # Back button
        self.go_back_button = tk.Button(self.Settings_window, text="Back", command=self.close_window)
        self.go_back_button.pack(pady=10)

    def select_credential_file(self):
        """Allow user to select a file and save its full path to credential_location.txt."""
        file_path = filedialog.askopenfilename(title="Select Credential File")
        if file_path:
            with open(self.credential_file, "w") as file:
                file.write(file_path)
            messagebox.showinfo("Success", f"Credential file path saved to {self.credential_file}.")
            sheetlink.update_creds()

    def select_protocol_file(self):
        """Allow user to select a file and save its full path to protocol_location.txt."""
        file_path = filedialog.askopenfilename(title="Select Protocol File")
        if file_path:
            with open(self.protocol_file, "w") as file:
                file.write(file_path)
            messagebox.showinfo("Success", f"Protocol file path saved to {self.protocol_file}.")
            sheetlink.update_protocol()

    def designate_sheet(self):
        """Save the designated sheet name to sheet_name.txt."""
        sheet_name = self.sheet_name_var.get().strip()
        if not sheet_name:
            messagebox.showwarning("Input Error", "Sheet name cannot be empty.")
            return
        
        with open(self.sheet_name_file, "w") as file:
            file.write(sheet_name)
        messagebox.showinfo("Success", f"Sheet name saved to {self.sheet_name_file}.")
        sheetlink.update_sheet_name(sheet_name)

    def load_authors(self):
        """Load authors from the sheetlink.fetch_authors() method."""
        try:
            return sheetlink.fetch_authors()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch authors: {e}")
            return []

    def add_author(self):
        """Add a new author to the Google Sheets and refresh the ComboBox."""
        new_author = self.author_var.get().strip()
        if not new_author:
            messagebox.showwarning("Input Error", "Author name cannot be empty.")
            return
    
        try:
            sheetlink.add_author(new_author)
            self.load_authors()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add author: {e}")

    def delete_author(self):
        """Delete the selected author from the Google Sheets and refresh the ComboBox."""
        selected_author = self.author_var.get().strip()
        if not selected_author:
            messagebox.showwarning("Selection Error", "Please select an author to delete.")
            return
    
        try:
            sheetlink.delete_author(selected_author)
            self.load_authors()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete author: {e}")

    def close_window(self):
        """Close the settings window."""
        self.Settings_window.destroy()

#%% New run panel 
class RunPanelWindow:
    def __init__(self, root):
        self.run_window = tk.Toplevel(root)
        self.run_window.title("Run Management")
        self.run_window.geometry("500x640")
        
       
        
        panel_label = ttk.Label(self.run_window, text="Select a Panel") # this controls how the progam will handles volumes smaller than 
        panel_label.pack(pady=10)
        panel_options = sheetlink.fetch_panel_names()
        panel_options=panel_options+["Upload a panel"]
       
        self.panel_var = tk.StringVar()
        self.panel_var.set(panel_options[0])  # Safely set the default value
        panel_dropdown = tk.OptionMenu(self.run_window, self.panel_var, panel_options[0], *panel_options)
        panel_dropdown.pack()
        plates_label = ttk.Label(self.run_window, text="How many plates?")
        plates_label.pack(pady=10)
        self.plates_var = tk.StringVar()
        plates_dropdown = ttk.OptionMenu(self.run_window, self.plates_var, 0,*range(5)) # accept up to 4 plates at a time 
        plates_dropdown.pack()   
        # Label and Entry for Media Amount#2
        reagent_label = ttk.Label(self.run_window, text="How much of your tertiary solution do you want added to the plate?")
        reagent_label.pack(pady=10)
        self.reagent_volume = tk.StringVar()
        reagent_entry = ttk.Entry(self.run_window, textvariable=self.reagent_volume)
        reagent_entry.pack()
        #Label and Entry For Null Wells
        null_label = ttk.Label(self.run_window, text="List null wells(captial letters, separate with commas)")
        null_label.pack(pady=10)
        self.null_wells_raw = tk.StringVar()
        reagent_entry = ttk.Entry(self.run_window, textvariable=self.null_wells_raw)
        reagent_entry.pack()
        # Label and Dropdown for Small Volume Handling
        volume_label = ttk.Label(self.run_window, text="How should the protocol handle small volumes?(ul)") # this controls how the progam will handles volumes smaller than 
        volume_label.pack(pady=10)
        self.volume_var = tk.StringVar()
        self.volume_var.set("Dilute 1:100")
        volume_options = ["Omit Volumes", "Round Up Wells", "Double Transfer Approximation", "Proportional Increase","Dilute 1:100"] #omit volumes leaves out sub 1ul volumes, rounding round all volumes up to 1ul, double transfer approximation will add a plate to collect and  insert, dilution just dilutes out 
        volume_dropdown = tk.OptionMenu(self.run_window, self.volume_var, *volume_options)
        volume_dropdown.pack()
        fillopt_label = ttk.Label(self.run_window, text="Fill plates proportionally from collection tube or direct from rack?") # this controls how the program will handle fill the plate
        fillopt_label.pack(pady=10)
        self.fill_var = tk.StringVar()
        self.fill_var.set("Rack")
        fill_options = ["Collection Tube", "Rack" ]
        fill_dropdown = tk.OptionMenu(self.run_window, self.fill_var, *fill_options)
        fill_dropdown.pack()
        distribute_label = ttk.Label(self.run_window, text="How much of your collection tube volume  do you want added to each well?")
        distribute_label.pack(pady=10)
        self.distribute_volume = tk.StringVar()
        distribute_entry = ttk.Entry(self.run_window, textvariable=self.distribute_volume)
        distribute_entry.pack()
        format_label = ttk.Label(self.run_window, text="Is the rack 96 cryotubes or 1-in-4 tube rack?") # this controls how the program will handle fill the plate
        format_label.pack(pady=10)
        self.formatp = tk.StringVar()
        self.formatp.set("Tube Rack")
        format_options = ["96-Cryotube", "Tube Rack" ]
        format_dropdown = ttk.OptionMenu(self.run_window, self.formatp, *format_options)
        format_dropdown.pack()
        submit_button = ttk.Button(self.run_window, text="Submit", command=self.prepare_panel)
        submit_button.pack(pady=20)
    def prepare_panel(self) :
        if "Upload" in self.panel_var.get():# Get script directory
            filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
            try:
                if filepath.endswith('.xlsx'):
                    dataframe = pd.read_excel(filepath)
                    
                elif filepath.endswith('.csv'):
                    dataframe = pd.read_csv(filepath)
                  
            except Exception as e:
                messagebox.showerror("Error", "Failed to process the file.\n{}".format(e))
        else:# this is for 
            script_dir = os.path.dirname(os.path.abspath(__file__))
            try:
    
                # Get all values from the worksheet
                data = sheetlink.fetch_pure_panel(self.panel_var.get())
    
                # Convert to pandas DataFrame
                df = pd.DataFrame(data)
    
                # Set first row as column headers
                df.columns = df.iloc[0]  # Make first row the header
                df = df[1:]  # Remove the first row from the data
    
                # Define output file paths
                #csv_output_path = os.path.join(script_dir, "exported_data.csv")
                excel_output_path = os.path.join(script_dir, "exported_data.xlsx")
    
                # Export DataFrame to CSV
                df.to_csv(excel_output_path, index=False)
    
                # Export DataFrame to Excel
                df.to_excel(excel_output_path, index=False, engine="openpyxl")
            except FileNotFoundError:
                    print(f"Error: The file 'credentials.txt' was not found in the directory {script_dir}.")
            except PermissionError:
                    print(f"Error: Permission denied while trying to read 'credentials.txt' in the directory {script_dir}.")
            except ValueError as ve:
                    print(f"Error: {ve}")
            except Exception as e:
                    print(f"An unexpected error occurred: {e}")
            
        try:
                
                dataframe = pd.read_excel(excel_output_path)
        except Exception as e:
                messagebox.showerror("Error", "Failed to process the file.\n{}".format(e))

        columns = dataframe.columns.tolist()
        inventory_columns = []
        volume_columns = []
        channel_columns = []
        dye_columns = []
        marker_columns=[]
        dilution_columns=[]
        offset_columns=[]
        for column in columns:
            if "ab number" in column.lower() or "inventory" in column.lower():
                inventory_columns.append(dataframe[column].tolist())
            elif "vol"  in column.lower() or "amount" in column.lower():
                volume_columns.append(dataframe[column].tolist())
            elif "channel" in column.lower():
                channel_columns.append(dataframe[column].tolist())
            elif "dye" in column.lower():
                dye_columns.append(dataframe[column].tolist())
            elif "marker"in column.lower() or "specificity" in column.lower() or "antigen" in column.lower():
                marker_columns.append(dataframe[column].tolist())  
            elif "dil" in column.lower(): 
                dilution_columns.append(dataframe[column].tolist())  
            elif "offs" in column.lower(): 
                offset_columns.append(dataframe[column].tolist())  
        print(inventory_columns)
        print(volume_columns) 
        print(channel_columns) 
        print(dye_columns)
        print(marker_columns)
        print(dilution_columns)
        print(offset_columns)
        if len(channel_columns)!=len(volume_columns):
                messagebox.showerror("Error", "An error occurred: Mismatched Columns") 
        else:
           self.volumes=volume_columns[0]# extract the lists, perhaps could revise 
           self.inventories=inventory_columns[0]
           self.channels=channel_columns[0]
           self.dyes= dye_columns[0]
           self.markers=marker_columns[0]
           self.dilutions=dilution_columns[0]
           self.offset=offset_columns[0]
           print("AHAH! columns sorted")
        
            
        self.compose_protocol()
  
    def compose_protocol(self): # the following functions compose the listbox codes                                               
      #create labware
     
       labware_list=[]#initalize a labware list to insert
       if "Cryo"  in self.formatp.get():
           labware_list.append("Dock3:Thermo Matrix Plate")
           formatted_labware="Dock3:Thermo Matrix Plate"
       else:
           labware_list.append("Dock3:Centrifuge Tube Rack")
           formatted_labware="Dock3:Centrifuge Tube Rack"
       collection_tube="A1"
       
       self.media_volume=(self.dilutions[1]*self.volumes[1])-sum(self.volumes) #calculate media volume
       if sum(self.volumes)+float(self.media_volume)>1500 or float(self.media_volume)+float(self.reagent_volume.get())>0 and not sum(self.volumes)+float(self.reagent_volume.get())*float(self.plates_var.get())+float(self.media_volume)>12000: # #if buffer volume greater than 1.5 ml tell the  program to dispense  elsewhere 
           labware_list.append("Dock6:Falcon Tube Rack x15")
       elif sum(self.volumes)+float(self.media_volume)+float(self.reagent_volume.get())*float(self.plates_var.get())>12000:
           labware_list.append("Dock6:Falcon Tube Rack x6") 
       if "double" in self.volume_var.get() or "lut" in self.volume_var.get():  #double transfer or dilution  creates a dilution plate
           labware_list.append("Dock9:96 Well Plate")
       for platenum in range(int(self.plates_var.get())):#tell the program where and if to place plates 
              if platenum+1<3: #1 and 2 go in their slots
               labware_list.append(f"Dock{int(platenum)+1}:96 Well Plate")
              elif platenum+1<5:
                  labware_list.append(f"Dock{int(platenum)+2}:96 Well Plate")    
              elif  platenum+1<7:
                  labware_list.append(f"Dock{int(platenum)+3}:96 Well Plate")

       liquidpanel=[]# initalize a liquid panel list 
       
       if "Proportional" in  self.volume_var.get(): 
           #requires a proportion be used to raise all values to the volume minimum volume
           proportion= 2/min(self.volumes) #search for the minimum volume in ul and divide 2ul by it
           if float(self.media_volume)>0: #add media first if present, apply proportion, and  float the value round
               self.reserve_media_volume=float(self.media_volume)*proportion+float(self.reagent_volume.get())*int(self.plates_var.get()) #will use again,save the value
               liquidpanel.append(f"Solution/Reagent|{labware_list[1]}|'A2'|{self.reserve_media_volume}|#e5e1f1")     # all labware for the panel media/buffer is the second rack. This merely avoids technical debt by allowing conditional rack selection
           self.volumes = [ proportion * item for item in self.volumes ] #multiply all the values in the volumes to get renewed volume

           self.media_volume=Decimal(float(self.media_volume)*proportion).quantize(Decimal('1'), rounding=ROUND_HALF_UP)
           self.reagent_volume=float(self.reagent_volume.get())*proportion*float(self.plates_var.get()) 
      
       elif "Round" in  self.volume_var.get():
           if float(self.media_volume)>0: #add media first if present, float the value round
               self.reserve_media_volume=float(self.media_volume)+float(self.reagent_volume.get())*int(self.plates_var.get()) #will use again,save the value
               liquidpanel.append(f"Solution/Reagent|{labware_list[1]}|'A2'|{self.reserve_media_volume}|#e5e1f1")
           self.media_volume=Decimal(float(self.media_volume)).quantize(Decimal('1'), rounding=ROUND_HALF_UP)
           self.reagent_volume=float(self.reagent_volume)*float(self.plates_var.get()) 
           self.volumes = [ 1 if item<1 else item for item in self.volumes] #convert sub 1 values  to 1 in a seperate line
           self.volumes = [ round(item,1) for item in self.volumes ]  #round non sub 1 volumes
      
       elif "Double" in  self.volume_var.get ():     
           if float(self.media_volume)>0: #add media first if present, float the value round
               self.reserve_media_volume=float(self.media_volume.get())+float(self.reagent_volume.get())*int(self.plates_var.get()) #will use again,save the value
               liquidpanel.append(f"Solution/Reagent|{labware_list[1]}|'A2'|{self.reserve_media_volume}|#e5e1f1")
           self.volumes = [ .5 if item<.5 else item for item in self.volumes] #convert sub .7 values  to .5 in a seperate line
           #self.volumes = [ round(.5 * round(float(item)/.5),1) for item in self.volumes] #round the values in the volumes to .5 place
           self.media_volume=round(.5 * round(float(self.media_volume)/.5),1)# round to.5 for use in the protocol
           self.reagent_volume=round(.5 * round(float(self.reagent_volume.get())/.5),1)*float(self.plates_var.get())  #round the value to .5 for use in the protocol
       elif "Dil" in self.volume_var.get ():  #basically untreated     
           subvolumes=[ 100*item if item<.5 else 0 for item in self.volumes] #dilution solution to add
           origsubvolumes=[ item if item<.5 else 0 for item in self.volumes] #orignal volume you wanted for subvolume quanr
           submedia=[ 100 if item<.5 else 0 for item in self.volumes]
           totalsub=sum(subvolumes)  #total excess media
           submediatotal=sum(submedia)
           
           print(f"total sub is {totalsub}")
           self.media_volume=self.media_volume-totalsub+sum(origsubvolumes)  #total media is orignal media  minus the dilution  and plus the value you realley wanted to get the "fraction" of dilution 
           self.reserve_media_volume=self.media_volume+submediatotal+float(self.reagent_volume.get())*96*float(self.plates_var.get()) #calculate ALL MEDIA volume
           if float(self.media_volume)>0: #add media first if present, float the value round
               #self.reserve_media_volume=float(self.media_volume+float(self.reagent_volume.get())*int(self.plates_var.get())) #will use again,save the value
               liquidpanel.append(f"Solution/Reagent|{labware_list[1]}|'A2'|{self.reserve_media_volume}|#e5e1f1")
               self.reagent_volume=float(self.reagent_volume.get())#*float(self.plates_var.get())       
      
       else: #else is effectively  omit, round to .5
           if float(self.media_volume)>0: #add media first if present, float the value, round
               self.reserve_media_volume=round(float(self.media_volume)+float(self.reagent_volume.get())*int(self.plates_var.get())) #will use again,save the value
               liquidpanel.append(f"Solution/Reagent|{labware_list[1]}|'A2'|{self.reserve_media_volume}|#e5e1f1")
               self.media_volume=Decimal(float(self.media_volume)).quantize(Decimal('1'), rounding=ROUND_HALF_UP)
               self.reagent_volume=float(self.reagent_volume.get())*float(self.plates_var.get())       
      
      
       
      #construct the list of all wells, insert nulls
       #This is some code that is supposed to allow you to restart by writing all the  values from the columns before edits happens
       #WARNING JUST SAVING THE COLUMNS LIST TO AN ATTRIBUTE WILL STILL ALLOW  THE LOOPS TO DELETE OR INSERT VALUES IF THE VALUES ARE BASED ON THE SAME COLUMNLISTS
       self.v=[]
       self.i=[]
       self.m=[]
       self.c=[]
       self.d=[]
       self.l=[]
       self.o=[]
       for mark,dye,vol,inv,chan,dil,off in zip(self.markers,self.dyes,self.volumes,self.inventories,self.channels,self.dilutions, self.offset):
           self.v.append(vol)
           self.d.append(dye)
           self.c.append(chan)
           self.i.append(inv)
           self.m.append(mark) 
           self.l.append(dil)
           self.o.append(off)
       self.v=[self.v]
       self.i=[self.i]
       self.m=[self.m]
       self.c=[self.c]
       self.d=[self.d]
       self.l=[self.l]
       self.l=[self.o]
       if self.null_wells_raw.get():
           for item in self.null_wells_raw.get().split(","):
               row = ord(item[0]) - ord('A') + 1
               col = int(item[1:])
               well_number = (row - 1) * 12 + col-1
               self.volumes.insert(well_number,1)
               self.dyes.insert(well_number,"null")
               self.channels.insert(well_number,"null")
               self.inventories.insert(well_number,"null")
               self.markers.insert(well_number,"null")
       self.wellist=[]
       index=1
       for vol in self.volumes: #compose wells
           row = (index - 1) // 12  # Calculate the row index
           col = (index - 1) % 12  # Calculate the column index
           well_name = chr(row + ord('A')) + str(col + 1)         
           self.wellist.append(well_name)
           index=index+1
       print("after nulls check") 
      
       print(self.i, self.v, self.c, self.d, self.m)
        # getting hard to read, reduce the unique if statement and construct the  liquids
       #lists here will have equal length, nulls inserted, and processed volumes but no removed volumes or wells. this is good because of degbugging and for independent loopmanipulation
       omitted_dye=[]
       omitted_channel=[]
       omitted_volume=[]
       omitted_inventory=[]
       omitted_wells=[]
       omitted_markers=[]
       if "Omit" in  self.volume_var.get(): #omit is the most straight foward revison
           #tag the omitted  values with "#" in volume
           #self.volumes=["#"+f"{item}" if item<1 or "null" in item  else item for item in self.volumes]
           # x=0# reset the index, you can redo with enumerate or comprehension, this is a result of crunch and technical 
               for value in self.volumes[:]:
                   if value < 1:
                       index = self.volumes.index(value)# Get the index of the value in the original list1
                       # Add the value at index i from all lists to the standby list
                       omitted_dye.append(self.dyes[index])
                       omitted_channel.append(self.volumes[index])
                       omitted_volume.append(self.channels[index])
                       omitted_inventory.append(self.inventories[index])
                       omitted_wells.append(self.wellist[index])
                       omitted_markers.append(self.markers[index])
                       # Remove the value at index i from all lists
                       del self.dyes[index]
                       del self.volumes[index]
                       del self.channels[index]
                       del self.inventories[index]
                       del self.wellist[index] 
                       del self.markers[index]#lists at the end of this loop  are equal length, have null with volume 1 and all correpsonding cells done  via the omit method                                  
       #start the liquidizing, may want to do this by specificity instead liquids are pretty much done the same way as omitted wells already removed
       # now we can use logic to decide if we need to place the collection tube in the Thermo Matrix Plate, or the large tube rack; its volume to indicate through the opentrons  app that inital volume is 0
       if "Cryo" in self.formatp.get():# 
            liquidpanel.append(f"collection_tube|{labware_list[1]}|'A1'|0|#e5e1f1")
            collect_t='A1'
            collect_p=6
            max_wells = 96  # Maximum number of wells in rack
       else:
           #add a conditional to define based on volume
           if self.dilutions[1]*self.volumes[1]<1700: #not media volume, what is your total volume you want 
               liquidpanel.append(f"collection_tube|{labware_list[0]}|'D6'|0|#e5e1f1") #if not add collection to lower right 
               collect_t='D6'
               collect_p=3 #tube rack
               max_wells = 23  # Maximum number of wells in rack, must be 23 to stop double counting the collection tube
           if self.dilutions[1]*self.volumes[1]>1700:
               liquidpanel.append(f"collection_tube|{labware_list[1]}|'A1'|0|#e5e1f1") #if not add collection to second rack 
               collect_t='A1'
               collect_p=6# second rack
               max_wells = 24  # Maximum number of wells in rack, must be 24 to have all tube slots
        
       self.fillist=[] #create a empty fill list to make the logic easier, will search this to give the destination and 
       rackswap=""# add a rack swap indicator 
       rackswaplist=[]
       index=1 #gotta reindex here, can't use the well list becauese we saved nulls in the list and omitted wells
       for mark,vol,well,chan,inv in zip(self.markers,self.volumes,self.wellist,self.channels,self.inventories):
            print(chan)
            if "UV" in chan:color="#d9d9d9"
            elif "V" in chan: color="#c062ff"                      
            elif "B" in chan:color="#72a0eb" 
            elif "YG" in chan:color="#92d050"  
            elif  "R" in chan:   color="#ff7e79"  
            else:color="#d9d9d9"
            if "null" not in mark and "live" not in mark.lower(): #make sure a there is not null or live/dead (case insensitve) and proceed
                if "Cryo" in self.formatp.get():# 
                    liqw = chr((index - 1) // 12 + 1+ 64)+str((index - 1) % 12+1 ) #effectively one line the code
                else:
                    liqw = chr((index - 1) // 6 + 1+ 64)+str((index - 1) % 6+1 ) #effectively one line the code
                liquidpanel.append(f"{rackswap}{mark} #{inv}|{labware_list[0]}|'{liqw}'|{vol}|{color}")
                self.fillist.append(f"{liqw}")
                if "*" in f"{rackswap}": rackswaplist.append(f"{mark}:{liqw}") #this text search based approach is a legacy of the way rackswap logic worked. you can revise here to stop marking with asterisk 
                index=index+1
            else:
                self.fillist.append("skip") #tell it to skip dyes and nulls
            if index > max_wells:
                rackswap=rackswap+"*"
                index=1 #resetsthe index
            
       print(liquidpanel)
       #prepare a rackswap list
       print(rackswaplist)
       #already made total media volume,plate number,per plate media volume and per tube media. just  complete total dye(reagent) and write the  messagebox
       #self.destroy()
       response = messagebox.askyesnocancel("Question", f"Buffer per well:{self.reagent_volume}\n Buffer in Tube{self.media_volume} \n Total Media:{self.reserve_media_volume} \n Total dye use:{sum(self.volumes)}\nNumber of Plates:{self.plates_var.get()}\nLoad plate from:{self.fill_var.get()}\nVolume mode:{self.volume_var.get()}\n Rack Swap:\n Remember to rackswap these wells:{rackswaplist}\n Would you like to generate a list of rack swapped items for later? [Hit Cancel To Restart]  ")
       #now comes the time to insert your lists  and if necessary export the rackswap data 
       if response is True:
            print("Response: Yes")
            print(self.volumes)
            print(self.inventories)
            print(self.channels)
            print(self.dyes)
            print(self.markers)
            print(self.wellist)
            print (self.fillist)
            print(labware_list)
            print("volumelist")
            print(self.volumes)
            #create a multi line dictonary to create a pickle file 
            self.dispenseParameters = {
                'fillmode': self.fill_var.get(),
                'volmode': self.volume_var.get(),
                'inventories': self.inventories,
                'channels': self.channels,
                'dyes': self.dyes,
                'vol': self.volumes,
                'markers': self.markers,
                'offset': self.offset,
                'plates':  self.plates_var.get() ,
                'wellist': self.wellist,
                'fillist': self.fillist,
                'mediavolume':int(self.media_volume) ,
                'labware': labware_list ,
                'liquids': liquidpanel ,
                'collectionwell': collect_t,
                'collectionplate': collect_p,
                'reagent volume': int(self.reagent_volume),
                'rackswapwell':(max_wells+len(self.null_wells_raw.get())),
                'rackswaplist':rackswaplist,
                'platevolume':float(self.distribute_volume.get())
                }
            path = os.getcwd()
            pickle.dump(self.dispenseParameters, open(f"{path}/panel.pkl", "wb"))
            url = "http://169.254.120.95:48888/api/contents/panel.pkl"  # Replace with your captured Request URL
            file_path = f"{path}/panel.pkl"  # Path to the local file to upload

            # Headers (adjust based on captured headers)
            headers = {
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "Content-Type": "application/json",
                "X-Requested-With": "XMLHttpRequest",
                "X-XSRFToken": "2|f6413e68|5c5ec0d4cbcb19a618c4cd14fac760d8|1737063372",  # Replace with captured XSRF token
                "Cookie": "_xsrf=2|f6413e68|5c5ec0d4cbcb19a618c4cd14fac760d8|1737063372",  # Replace with captured cookie
            }

            # Read the file content in binary mode
            with open(file_path, "rb") as f:
                file_content = f.read()

            # Encode the binary content as Base64
            file_content_base64 = base64.b64encode(file_content).decode("utf-8")  # Decode ensures it's a string

            # Payload (adjust format if needed)
            payload = {
                "type": "file",
                "format": "base64",  # Base64 encoding for binary files
                "content": file_content_base64,
            }

            # Send the PUT request
            response = requests.put(url, data=json.dumps(payload), headers=headers)

            # Handle the response
            if response.status_code in [200, 201]:
                print("File uploaded successfully!")
                print(f"Response: {response.json()}")
                messagebox.showinfo("Success", "File uploaded successfully!\n" + f"Response: {response.json()}")
            else:
                print(f"Failed to upload file. Status code: {response.status_code}")
                print(f"Response: {response.text}")
                messagebox.showerror("Error", f"Failed to upload file. Status code: {response.status_code}\nResponse: {response.text}")
            import uuid
            import tkinter as tk
            #from tkinter import messagebox
            import time
            
            # Path to your protocol file
            script_dir = os.path.dirname(os.path.abspath(__file__))
          
            #protocol_path = "/Users/westtn/Library/Application Support/Opentrons/protocols/227bd841-0e2c-45f7-969e-00f594c91592/src/panelprotocol.py"
            protocol_file_path = os.path.join(script_dir, "protocol_location.txt")
            print("protocol_file_path")
            with open(protocol_file_path, "r") as file:
                protocol_path = file.read().strip()  # Remove any extra whitespace or newlines
            # Generate a new UUID
            #make sure you don't bungle the upload
            print(protocol_path)
            new_uuid = str(uuid.uuid4())
                
                # Read the file and modify the last line
            with open(protocol_path, 'r') as f:
                lines = f.readlines()
            print(f"updating at {lines}")
            if "/Library/Application Support/Opentrons/" in protocol_path:     
                print("beginning update")
                # Update the last line containing #update_uuid=
                updated = False
                for i in range(len(lines)):
                    if "#update_uuid=" in lines[i]:
                        lines[i] = f"#update_uuid={new_uuid}\n"
                        updated = True
                        print("update")
                        break  # Stop after updating the first match
                
                if updated:
                    # Write the modified content back to the file
                    with open(protocol_path, 'w') as f:
                        f.writelines(lines)
                
                    # Verify the update by re-reading the file
                    time.sleep(0.5)  # Small delay to ensure the OS has flushed the write
                    with open(protocol_path, 'r') as f:
                        if f"#update_uuid={new_uuid}\n" in f.readlines():
                            verification_passed = True
                        else:
                            verification_passed = False
                else:
                    verification_passed = False
                
                # Create a Tkinter message box to confirm or warn
                root = tk.Tk()
                root.withdraw()  # Hide the main window
                
                if verification_passed:
                    messagebox.showinfo("Protocol Updated", f"Successfully updated protocol UUID to:\n{new_uuid}")
                    print(f"✅ Successfully updated protocol UUID to: {new_uuid}")
                else:
                    messagebox.showwarning("Update Failed", "❌ Failed to update the protocol file. Please check manually.")
                    print("❌ Failed to update the protocol file.")
            else:
                  messagebox.showwarning("Update Failed", "Wrong paths for updated protocol. Please check manually.")
            print("done")
            
       elif response is False:
            print("Response: No")
       else:
            print("Response: Cancel")
            print("start new window")
            
   
#%%
  
    

# Method to close the Alerts window
    def close_window(self):
        self.Settings_window.destroy()
if __name__ == "__main__":
   root = tk.Tk()    
    #app = AntibodyScraperApp(root)
   # root.mainloop()
   app = WelcomeWindow(root)
   root.mainloop()