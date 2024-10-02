import os
import tkinter as tk
from tkinter import ttk, messagebox
import logic
import sheetlink
from tkinter import filedialog
import pandas as pd
from PIL import Image, ImageTk

#%% Main Window
class AntibodyScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Antibody Scraper Main Menu")
        
        # Create a frame to hold the buttons
        self.frame = ttk.Frame(root)
        self.frame.pack(expand=True)
        
        # Main menu buttons
        ttk.Button(self.frame, text="Add an Antibody", command=self.open_add_antibody_window).pack(padx=10, pady=10)
        ttk.Button(self.frame, text="Create a Panel", command=self.open_create_panel_window).pack(padx=10, pady=10)
        ttk.Button(self.frame, text="Modify a Panel", command=self.open_modify_panel_window).pack(padx=10, pady=10)
        ttk.Button(self.frame, text="Run a Panel", command=self.open_run_panel_window).pack(padx=10, pady=10)
        ttk.Button(self.frame, text="Use Single Antibody", command=self.open_use_antibody_window).pack(padx=10, pady=10)
        ttk.Button(self.frame, text="Log Antibody Issue", command=self.open_log_issue_window).pack(padx=10, pady=10)
    
        # Perform low volume check at startup
        self.check_low_volume_antibodies()
    
    def open_add_antibody_window(self):
        AddAntibodyWindow(self.root)

    def open_create_panel_window(self):
        PanelCreatorApp(self.root)

    def open_modify_panel_window(self):
        ModifyPanelWindow(self.root)

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
#%% Add an Antibody
class AddAntibodyWindow:
    def __init__(self, root):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Add an Antibody")
        
        # Company input as a dropdown
        ttk.Label(self.new_window, text="Company:").grid(row=0, column=0, padx=10, pady=10)
        self.company_var = tk.StringVar()
        self.company_combobox = ttk.Combobox(self.new_window, textvariable=self.company_var)
        self.company_combobox['values'] = ("BD", "Biolegend","Thermofisher", "Tonbo", "Other (manual entry)")
        self.company_combobox.grid(row=0, column=1, padx=10, pady=10)
        
        # Catalog number input
        ttk.Label(self.new_window, text="Catalog Number:").grid(row=1, column=0, padx=10, pady=10)
        self.catalog_entry = ttk.Entry(self.new_window)
        self.catalog_entry.grid(row=1, column=1, padx=10, pady=10)
        
        # Scrape button
        self.scrape_button = ttk.Button(self.new_window, text="Scrape", command=self.scrape)
        self.scrape_button.grid(row=2, column=0, columnspan=2, pady=10)

    def scrape(self):
        company = self.company_var.get()
        catalog_number = self.catalog_entry.get()
        
        if company == "Other (manual entry)":
            ManualEntryWindow(self.new_window)
            return
        
        if not company or not catalog_number:
            messagebox.showerror("Input Error", "Please provide both company and catalog number.")
            return
        
        # Call the scraping function (to be implemented)
        self.scrape_antibody_info(company, catalog_number)
    
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
        elif company=="Tonbo":
           newdata=logic.TonboCytek_Scraper(catalog_number)
           prefilled_data = {
               "Supplier": company if company != "Other (manual entry)" else "",
               "Catalog Number": catalog_number}
           prefilled_data.update(newdata)
           print(prefilled_data)
           ManualEntryWindow(self.new_window, prefilled_data)
        elif company=="Biolegend":
            newdata=logic.Biolegend_Scraper(catalog_number)
            prefilled_data = {
                "Supplier": company if company != "Other (manual entry)" else "",
                "Catalog Number": catalog_number}
            prefilled_data.update(newdata)
            print(prefilled_data)
            ManualEntryWindow(self.new_window, prefilled_data)   
        elif company=="BD":
            newdata=logic.BD_scraper(catalog_number)
            prefilled_data = {
                "Supplier": company if company != "Other (manual entry)" else "",
                "Catalog Number": catalog_number}
            prefilled_data.update(newdata)
            print(prefilled_data)
            ManualEntryWindow(self.new_window, prefilled_data)
        elif company=="Thermofisher":
            newdata=logic.Thermofisher_Scraper(catalog_number)
            prefilled_data = {
                "Supplier": company if company != "Other (manual entry)" else "",
                "Catalog Number": catalog_number}
            prefilled_data.update(newdata)
            print(prefilled_data)
            ManualEntryWindow(self.new_window, prefilled_data)
            
        
        else:
            messagebox.showerror("Input Error", "Unknown company selected.")
       # self.new_window.destroy()
#%% Manual Entry
class ManualEntryWindow:
    def __init__(self, root, prefilled_data):
        self.manual_entry_window = tk.Toplevel(root)
        self.manual_entry_window.title("Antibody Entry")
        
        # Define labels and corresponding data types
        self.labels = ["Antibody Specificity:", "Label:","Peak Channel (CyTEK):", "Clone:","Target Species:", "Host Species:","Supplier:", "Catalog Number:",  "Volume:", "Number of Vials:"]
        
        self.entries = {}

        # Create labels and entry fields with optional prefilled data
        for i, label in enumerate(self.labels):
            ttk.Label(self.manual_entry_window, text=label).grid(row=i, column=0, padx=10, pady=5)
            entry = ttk.Entry(self.manual_entry_window)
            entry.grid(row=i, column=1, padx=10, pady=5)

            # Insert prefilled data if available
            print(prefilled_data.get(label.strip(':'), ""))
            print(prefilled_data)
            entry_value = prefilled_data.get(label.strip(':'), "") if prefilled_data else ""
            entry.insert(0, str(entry_value))  # Ensure that entry_value is a string
            
            self.entries[label] = entry
            #print(entry_value)
        # Submit button placed after all the entry fields
        submit_button = ttk.Button(self.manual_entry_window, text="Submit", command=self.submit_manual_entry)
        submit_button.grid(row=len(self.labels), column=0, columnspan=2, pady=10)
    
    def submit_manual_entry(self):
        # Collect data from entry fields
        data = {label.strip(':'): entry.get() for label, entry in self.entries.items()}
        print(data)
        # Add the collected data to the DataFrame via logic function
        if not "" in data.values(): #make sure every thing is filled out
            if not logic.check_existing_inventory(data): # check the worksheet for the preexisting column, if not found go ahead and export and give a little message
                
            
                messagebox.showinfo("Info", f"Manual entry submitted successfully! This antibody's ID is ")
               
                self.root.destroy() #kills the root window "new window" along with the current one
        else: 
            messagebox.showerror("Input Error", "Missing Parameter")
        
      
        
        # Show a message box with the antibody ID
        #messagebox.showinfo("Info", f"Manual entry submitted successfully! This antibody's ID is {next_antibody_number}")
        
        # Close the manual entry window
        


#%% Create panel window
class CreatePanelWindow:
    def __init__(self, root):
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
class RunPanelWindow:
    def __init__(self, root):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Run a Panel")

        # User selection dropdown (folders in Panels directory)
        ttk.Label(self.new_window, text="Select User:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        self.user_var = tk.StringVar()
        self.user_combobox = ttk.Combobox(self.new_window, textvariable=self.user_var)
        self.user_combobox.grid(row=0, column=1, padx=10, pady=10)
        self.update_user_dropdown()

        # Bind user selection to update panel dropdown
        self.user_combobox.bind("<<ComboboxSelected>>", self.update_panel_dropdown)

        # Panel selection dropdown (files in the selected user's folder)
        ttk.Label(self.new_window, text="Select Panel:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        self.panel_var = tk.StringVar()
        self.panel_combobox = ttk.Combobox(self.new_window, textvariable=self.panel_var)
        self.panel_combobox.grid(row=1, column=1, padx=10, pady=10)

        # Stain volume input
        ttk.Label(self.new_window, text="Total Stain Volume:").grid(row=2, column=0, padx=10, pady=10, sticky='w')
        self.stain_volume_entry = ttk.Entry(self.new_window)
        self.stain_volume_entry.grid(row=2, column=1, padx=10, pady=10)

        # Date input
        ttk.Label(self.new_window, text="Enter Date (YYYYMMDD):").grid(row=3, column=0, padx=10, pady=10, sticky='w')
        self.date_entry = ttk.Entry(self.new_window)
        self.date_entry.grid(row=3, column=1, padx=10, pady=10)

        # Run button
        self.run_button = ttk.Button(self.new_window, text="Run Panel", command=self.run_panel)
        self.run_button.grid(row=4, column=0, columnspan=2, pady=10)

    def update_user_dropdown(self):
        """Update the user dropdown with the folders in the Panels directory."""
        panels_dir = "Panels"
        users = [f for f in os.listdir(panels_dir) if os.path.isdir(os.path.join(panels_dir, f))]
        self.user_combobox['values'] = users

    def update_panel_dropdown(self, event=None):
        """Update the panel dropdown with .xlsx files from the selected user's folder."""
        selected_user = self.user_var.get()
        user_folder = os.path.join("Panels", selected_user)
        if os.path.isdir(user_folder):
            panels = [f for f in os.listdir(user_folder) if f.endswith('.xlsx')]
            self.panel_combobox['values'] = panels

    def run_panel(self):
        """Run the selected panel and calculate the volumes for each antibody."""
        selected_user = self.user_var.get()
        selected_panel = self.panel_var.get()
        stain_volume = self.stain_volume_entry.get()
        date = self.date_entry.get()

        # Ensure all inputs are provided
        if not selected_user or not selected_panel or not stain_volume or not date:
            messagebox.showerror("Input Error", "Please select a user, a panel, provide a stain volume, and a date.")
            return

        try:
            stain_volume = float(stain_volume)
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid number for the stain volume.")
            return

        # Check if the date format is correct
        if not date.isdigit() or len(date) != 8:
            messagebox.showerror("Input Error", "Please enter the date in YYYYMMDD format.")
            return

        # Load the selected panel's DataFrame
        panel_file = os.path.join("Panels", selected_user, selected_panel)
        panel_df = pd.read_excel(panel_file)

        # Create an 'Experiment Runs' folder if it doesn't exist
        experiment_folder = os.path.join("Experiment Runs")
        if not os.path.exists(experiment_folder):
            os.makedirs(experiment_folder)

        # Create a DataFrame to store the antibody usage for this run
        usage_df = pd.DataFrame(columns=['Antibody Number', 'Antibody Specificity', 'Label', 'Dilution', 'Volume Used'])

        # Iterate through the panel DataFrame and process each antibody
        for _, row in panel_df.iterrows():
            antibody_number = row['Antibody Number']
            dilution = float(row['Dilution'])
            antibody_specificity = row['Antibody Specificity']
            label = row['Label']

            # Calculate the volume to use for this antibody (dilution + 1uL for single stains)
            volume_used = stain_volume / dilution + 1

            # Add the used antibody data to the usage DataFrame
            usage_df = usage_df.append({
                'Antibody Number': antibody_number,
                'Antibody Specificity': antibody_specificity,
                'Label': label,
                'Dilution': dilution,
                'Volume Used': volume_used
            }, ignore_index=True)

            # Use the antibody by invoking the use_antibody function
            logic.use_antibody(antibody_number, volume_used)

            # Update the last user in the antibody DataFrame
            logic.update_last_user(antibody_number, selected_user)

        # Save the usage DataFrame to an Excel file in the 'Experiment Runs' folder with the date as part of the filename
        experiment_file = os.path.join(experiment_folder, f"{date}_{selected_user}_{selected_panel[:-5]}.xlsx")
        usage_df.to_excel(experiment_file, index=False)

        messagebox.showinfo("Success", f"Panel run successfully. Antibody usage has been saved to {experiment_file}.")
        self.new_window.destroy()



#%% Use antibody window
class UseAntibodyWindow:
    def __init__(self, root):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Use a Single Antibody")
        
        # Antibody number input
        ttk.Label(self.new_window, text="Antibody Number:").grid(row=0, column=0, padx=10, pady=10)
        self.antibody_number_entry = ttk.Entry(self.new_window)
        self.antibody_number_entry.grid(row=0, column=1, padx=10, pady=10)
        
        # Volume used input
        ttk.Label(self.new_window, text="Volume Used:").grid(row=1, column=0, padx=10, pady=10)
        self.volume_used_entry = ttk.Entry(self.new_window)
        self.volume_used_entry.grid(row=1, column=1, padx=10, pady=10)
        
        # Submit button
        submit_button = ttk.Button(self.new_window, text="Submit", command=self.submit_usage)
        submit_button.grid(row=2, column=0, columnspan=2, pady=10)
    
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
        success, vials, volume = logic.use_antibody(antibody_number, volume_used)
    
        if success:
            message = f'There are {vials} vials remaining and a volume of {volume}uL'
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", 'This function was unsuccessful.')
    
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
class WelcomeWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("705x440")  # Increased window height to accommodate layout
        self.title("Welcome")

        # Load the banner image at the top (use an actual image file path)
        self.banner_image_path = "/Users/westtn/Desktop/Screenshot 2024-09-20 at 1.51.27 PM.png"
        self.load_banner_image()

        # Place the buttons in the center (existing functionality)
        self.create_button_frames()

        # Create the browse button at the bottom left
        self.create_browse_button()

        # Create round settings and alerts buttons at the bottom right
        self.create_round_buttons()

    def load_banner_image(self):
    # Load banner image (use a path to your image file)
        banner_image = Image.open(self.banner_image_path)
        banner_image = banner_image.resize((700, 100), Image.Resampling.LANCZOS)  # Use LANCZOS instead of ANTIALIAS
        self.banner_photo = ImageTk.PhotoImage(banner_image)

    # Create a label to display the image
        banner_label = tk.Label(self, image=self.banner_photo)
        banner_label.pack(pady=0)
    
#ttk.Button(self.frame, text="Add an Antibody", command=self.open_add_antibody_window).pack(padx=10, pady=10)
#ttk.Button(self.frame, text="Create a Panel", command=self.open_create_panel_window).pack(padx=10, pady=10)
#ttk.Button(self.frame, text="Modify a Panel", command=self.open_modify_panel_window).pack(padx=10, pady=10)
#ttk.Button(self.frame, text="Run a Panel", command=self.open_run_panel_window).pack(padx=10, pady=10)
#ttk.Button(self.frame, text="Use Single Antibody", command=self.open_use_antibody_window).pack(padx=10, pady=10)
#ttk.Button(self.frame, text="Log Antibody Issue", command=self.open_log_issue_window).pack(padx=10, pady=10)
    def create_button_frames(self):
        # Button 1 Info
        button1_frame = tk.Frame(self)
        button1_frame.pack(padx=10, pady=1, anchor="w")

        button1 = tk.Button(button1_frame, text="Add an Antibody", command=self.open_super_gui)
        button1.pack(side="left", padx=5+10)

        label1_frame = tk.LabelFrame(button1_frame, text="Add Antibodies to Inventory")
        label1_frame.pack(side="left", padx=5)

        label1 = tk.Label(label1_frame, text="Add Antibodies to the Inventory based on availible lot numbers.                               ")
        label1.pack(padx=1, pady=1)

        # Button 2 Info
        button2_frame = tk.Frame(self)
        button2_frame.pack(padx=10, pady=1, anchor="w")

        button2 = tk.Button(button2_frame, text="Create a Panel", command=self.select_csv_file)
        button2.pack(side="left", padx=11+10)

        label2_frame = tk.LabelFrame(button2_frame, text="Compose A Panel Sheet")
        label2_frame.pack(side="left", padx=5)

        label2_text = "Compose a stain panel using the automated layout. Requires Wifi Connection.       "
        label2 = tk.Label(label2_frame, text=label2_text)
        label2.pack(padx=1, pady=1)

        # Button 3 Info
        button3_frame = tk.Frame(self)
        button3_frame.pack(padx=10, pady=1, anchor="w")

        button3 = tk.Button(button3_frame, text="Modify a Panel", command=self.select_dataframe_file)
        button3.pack(side="left", padx=11+10)

        label3_frame = tk.LabelFrame(button3_frame, text="Modify a Antibody Panel")
        label3_frame.pack(side="left", padx=5)

        label3_text = "Edit an existing antibody panel. Requires Wifi.                                                             "
        label3 = tk.Label(label3_frame, text=label3_text)
        label3.pack(padx=1, pady=1)
        # Button 4 Info
        button4_frame = tk.Frame(self)
        button4_frame.pack(padx=10, pady=1, anchor="w")
        
        button4 = tk.Button(button4_frame, text="Run a Panel", command=self.select_dataframe_file)
        button4.pack(side="left", padx=20+10)

        label4_frame = tk.LabelFrame(button4_frame, text="Run a Antibody Panel")
        label4_frame.pack(side="left", padx=5)

        label4_text = "Run an existing Antibody panel via opentrons. Requires Terminal/SSH connection."
        label4 = tk.Label(label4_frame, text=label4_text)
        label4.pack(padx=1, pady=1)
        # Button 5 Info
        button5_frame = tk.Frame(self)
        button5_frame.pack(padx=10, pady=1, anchor="w")
        
        button5 = tk.Button(button5_frame, text="Use Single Antibody", command=self.select_dataframe_file)
        button5.pack(side="left", padx=5)

        label5_frame = tk.LabelFrame(button5_frame, text="Edit Single Volumes")
        label5_frame.pack(side="left", padx=5)

        label5_text = "Edit/Locate volume on a single antibody for seperate/non-automated use.              "
        label5 = tk.Label(label5_frame, text=label5_text)
        label5.pack(padx=1, pady=1)
        # Button 6 Info
        button6_frame = tk.Frame(self)
        button6_frame.pack(padx=11, pady=1, anchor="w")
        
        button6 = tk.Button(button6_frame, text="Log Antibody Issue", command=self.select_dataframe_file)
        button6.pack(side="left", padx=7)

        label6_frame = tk.LabelFrame(button6_frame, text="Add Notes to Antibodies")
        label6_frame.pack(side="left", padx=5)

        label6_text = "Add a non-usage note to any antibody. Requires Wifi.                                                "
        label6 = tk.Label(label6_frame, text=label6_text)
        label6.pack(padx=1, pady=1)

    def create_browse_button(self):
        browse_button = tk.Button(self, text="Browse", command=self.browse_files)
        browse_button.place(x=20, y=400)  # Place the browse button at the bottom left

    def create_round_buttons(self):
        # Create the round settings button
        settings_button = tk.Button(self, text="⚙️", command=self.open_settings, height=2, width=4, font=("Arial", 10))
        settings_button.place(x=630, y=400)  # Place the settings button at the bottom right

        # Create the round alerts button next to the settings button
        alerts_button = tk.Button(self, text="🔔", command=self.show_alerts, height=2, width=4, font=("Arial", 10))
        alerts_button.place(x=570, y=400)  # Place the alerts button next to the settings button

    # Placeholder functions for button commands
    def open_super_gui(self):
        print("New Protocol button clicked")

    def select_csv_file(self):
        print("Stain from CSV button clicked")

    def select_dataframe_file(self):
        print("Ramp from Dataframe button clicked")

    def browse_files(self):
        file_path = filedialog.askopenfilename()
        print("File selected:", file_path)

    def open_settings(self):
        print("Settings button clicked")

    def show_alerts(self):
        print("Alerts button clicked")
#%% New Panel Creator ---- needs  modifier and  mouse vs human interim page
class PanelCreatorApp:
    def __init__(self, root):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Google Spreadsheet Treeview with Panel Creation")

        # Upper frame for user input
        self.upper_frame = ttk.Frame(self.new_window)
        self.upper_frame.pack(padx=10, pady=10, fill="x")

        # Dropdown for name selection
        ttk.Label(self.upper_frame, text="Select Your Name:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.name_var = tk.StringVar()
        self.name_combobox = ttk.Combobox(self.upper_frame, textvariable=self.name_var)
        self.name_combobox['values'] = ("Anagha", "Tim", "Hannah")
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
        self.sheet_data = sheetlink.fetch_mouse_inventory()

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

        tree = ttk.Treeview(parent, columns=columns, show='headings')

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
            
            # Filter out non-numeric inventory numbers
            try:
                int(inventory_number)  # Check if the value is numeric
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
        """Move selected row from right treeview to left treeview and add editable dilution, filtering non-numeric values."""
        selected_item = self.right_treeview.selection()
        if selected_item:
            item_values = self.right_treeview.item(selected_item, "values")
            
            # Filter out non-numeric inventory numbers
            try:
                int(item_values[0])  # Check if the first value is numeric
                self.left_treeview.insert("", "end", values=item_values + ("",))  # Move values to left treeview with empty dilution
    
                # Remove from right treeview
                self.right_treeview.delete(selected_item)
    
                # Sort the left treeview after moving
                self.sort_treeview(self.left_treeview)
            except ValueError:
                # Skip if the value is non-numeric
                pass


    def move_to_right(self):
        """Remove selected row from left treeview and return it to right treeview."""
        selected_item = self.left_treeview.selection()
        if selected_item:
            item_values = self.left_treeview.item(selected_item, "values")
            self.right_treeview.insert("", "end", values=item_values[:3])  # Return values (without dilution) to right treeview

            # Remove from left treeview
            self.left_treeview.delete(selected_item)
            self.sort_treeview(self.right_treeview)

    def sort_treeview(self, treeview):
        """Sort the treeview by the first column (Inventory Number), handling mixed types."""
        items = [(treeview.item(item)["values"], item) for item in treeview.get_children()]
    
        # Sort by the first column (Inventory Number)
        def sort_key(value):
            first_col_value = value[0][0]  # Extract the first column value (Inventory Number)
    
            # Try to convert to an integer for numeric sorting, otherwise keep as string
            try:
                return int(first_col_value)  # Sort numerically if possible
            except ValueError:
                return str(first_col_value)  # Otherwise, sort lexicographically
    
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
            if column == '#4':  # Check if the Dilution column was clicked
                row_id = self.left_treeview.identify_row(event.y)
                if row_id:
                    # Get the current value of the cell
                    current_value = self.left_treeview.item(row_id, 'values')[3]
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
        current_values[3] = new_value
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
        headers = ["Inventory Number", "Antibody Specificity", "Dilution"]
    
        for row_id in self.left_treeview.get_children():
            row = self.left_treeview.item(row_id, "values")
            panel_data.append([row[0], row[1], row[3]])  # Only include Inventory Number, Antibody Specificity, and Dilution
    
        # Add Total Volume at the bottom right
        panel_data.append(["", "", f"Total Volume: {total_volume}"])
    
        
        # Add a new worksheet with the generated sheet name
       #     sheetlink.export_panel(new_sheet_name,headers,panel_data): 
    
#%% Browse page



#%%Settings page




if __name__ == "__main__":
    root = tk.Tk()
    app = AntibodyScraperApp(root)
    root.mainloop()
    #app = WelcomeWindow()
    #app.mainloop()