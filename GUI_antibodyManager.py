import os
import tkinter as tk
from tkinter import ttk, messagebox
import logic
import pandas as pd

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
        CreatePanelWindow(self.root)

    def open_modify_panel_window(self):
        ModifyPanelWindow(self.root)

    def open_run_panel_window(self):
        RunPanelWindow(self.root)

    def open_use_antibody_window(self):
        UseAntibodyWindow(self.root)

    def open_log_issue_window(self):
        LogIssueWindow(self.root)

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

class AddAntibodyWindow:
    def __init__(self, root):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Add an Antibody")
        
        # Company input as a dropdown
        ttk.Label(self.new_window, text="Company:").grid(row=0, column=0, padx=10, pady=10)
        self.company_var = tk.StringVar()
        self.company_combobox = ttk.Combobox(self.new_window, textvariable=self.company_var)
        self.company_combobox['values'] = ("BD", "Biolegend", "Tonbo", "Other (manual entry)")
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
        if company in ["BD", "Biolegend", "Tonbo", "Other (manual entry)"]:
            prefilled_data = {
                "Supplier": company if company != "Other (manual entry)" else "",
                "Catalog Number": catalog_number
            }
 
            ManualEntryWindow(self.new_window, prefilled_data)
        else:
            messagebox.showerror("Input Error", "Unknown company selected.")

class ManualEntryWindow:
    def __init__(self, root, prefilled_data=None):
        self.manual_entry_window = tk.Toplevel(root)
        self.manual_entry_window.title("Antibody Entry")
        
        # Define labels and corresponding data types
        self.labels = ["Antibody Specificity:", "Label:", "Supplier:", "Catalog Number:", "Target Species:", "Host Species:", "Volume:", "Number of Vials:"]
        
        self.entries = {}

        # Create labels and entry fields with optional prefilled data
        for i, label in enumerate(self.labels):
            ttk.Label(self.manual_entry_window, text=label).grid(row=i, column=0, padx=10, pady=5)
            entry = ttk.Entry(self.manual_entry_window)
            entry.grid(row=i, column=1, padx=10, pady=5)

            # Insert prefilled data if available
            entry_value = prefilled_data.get(label.strip(':'), "") if prefilled_data else ""
            entry.insert(0, str(entry_value))  # Ensure that entry_value is a string
            
            self.entries[label] = entry
        
        # Submit button placed after all the entry fields
        submit_button = ttk.Button(self.manual_entry_window, text="Submit", command=self.submit_manual_entry)
        submit_button.grid(row=len(self.labels), column=0, columnspan=2, pady=10)
    
    def submit_manual_entry(self):
        # Collect data from entry fields
        data = {label.strip(':'): entry.get() for label, entry in self.entries.items()}
        
        # Add the collected data to the DataFrame via logic function
        next_antibody_number = logic.add_antibody_to_dataframe(data)
        
        # Show a message box with the antibody ID
        #messagebox.showinfo("Info", f"Manual entry submitted successfully! This antibody's ID is {next_antibody_number}")
        
        # Close the manual entry window
        self.manual_entry_window.destroy()



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
        self.user_name_combobox['values'] = ("Anagha", "Tim", "Hannah")  # Replace with actual names
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
        ttk.Label(self.antibody_inputs_frame, text="Antibody Number:").grid(row=row, column=0, padx=10, pady=5, sticky='w')
        antibody_entry = ttk.Entry(self.antibody_inputs_frame)
        antibody_entry.grid(row=row, column=1, padx=10, pady=5)

        ttk.Label(self.antibody_inputs_frame, text="Dilution:").grid(row=row, column=2, padx=10, pady=5, sticky='w')
        dilution_entry = ttk.Entry(self.antibody_inputs_frame)
        dilution_entry.grid(row=row, column=3, padx=10, pady=5)

        self.antibody_entries.append((antibody_entry, dilution_entry))

    def create_panel(self):
        """Create a new panel with the provided antibody information."""
        user_name = self.user_name_var.get()
        panel_name = self.panel_name_entry.get()

        if not user_name or not panel_name:
            messagebox.showerror("Input Error", "Please provide both your name and the panel name.")
            return

        panel_data = []

        for antibody_entry, dilution_entry in self.antibody_entries:
            antibody_number = antibody_entry.get()
            dilution = dilution_entry.get()

            if antibody_number and dilution:
                # Retrieve antibody details from the DataFrame
                if int(antibody_number) in logic.antibody_df['Antibody Number'].values:
                    antibody_details = logic.antibody_df[logic.antibody_df['Antibody Number'] == int(antibody_number)]
                    print(antibody_details)
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



class ModifyPanelWindow:
    def __init__(self, root):
        self.new_window = tk.Toplevel(root)
        self.new_window.title("Modify a Panel")
        ttk.Label(self.new_window, text="Panel modification functionality to be implemented").grid(row=0, column=0, padx=10, pady=10)


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





if __name__ == "__main__":
    root = tk.Tk()
    app = AntibodyScraperApp(root)
    root.mainloop()
