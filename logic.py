########################### the is  the logic sheet that handles anything that isn't a terminal or direct sheet writing
##########################    includes webscraping and peak channel  detemrination
##########################
#########################
import pandas as pd
from tkinter import messagebox
import numpy as np
import asyncio
from requests_html import AsyncHTMLSession
import tkinter as tk
from tkinter import simpledialog, messagebox
import nest_asyncio
import pdfplumber
import re
import time
import sheetlink
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
def BD_scraper(catalog_number):
    nest_asyncio.apply()

    # Function to show a pop-up to choose a title
    def choose_option(options):
        root = tk.Tk()
        root.withdraw()  # Hide the main Tkinter window
        # Pop-up to select from the available options
        option = simpledialog.askstring("Choose a product", f"Options:\n" + "\n".join([f"{i+1}. {title}" for i, title in enumerate(options)]))
        root.destroy()  # Close the Tkinter window after the selection
        if option and option.isdigit():
            return int(option) - 1  # Convert user input to index
        else:
            return None

    # Function to show an error message
    def show_popup(message):
        root = tk.Tk()
        root.withdraw()  # Hide the main Tkinter window
        messagebox.showerror("Error", message)
        root.destroy()

    # Asynchronous function to scrape the search results
    async def scrape_search_results(search_term):
        print(f"searched{search_term}")
        search_url = f"https://www.bdbiosciences.com/en-us/search-results?searchKey={search_term}"
        print(f"searched{search_url}")
        # Create an AsyncHTMLSession
        session = AsyncHTMLSession()

        # Fetch the search results page
        response = await session.get(search_url)

        # Render JavaScript (since the content is dynamically loaded)
        await response.html.arender(timeout=20)

        # Find elements with class 'pdp-search-card__body'
        elements = response.html.find('.pdp-search-card__body')

        # Extract titles and links from each result
        titles = [element.find('.pdp-search-card__body-title', first=True).text for element in elements]
        links = [element.find('a', first=True).attrs['href'] for element in elements]
        
        # Filter out .pdf links
        filtered_titles = [title for title, link in zip(titles, links) if not link.lower().endswith(".pdf")]
        filtered_links = [link for link in links if not link.lower().endswith(".pdf")]

        return filtered_titles, filtered_links

    # Asynchronous function to scrape the product details
    async def scrape_product_details(link):
        session = AsyncHTMLSession()
        
        # Fetch the product details page
        response = await session.get(link)
        await response.html.arender(timeout=20)

        # Find the product details from 'product-details-list__container'
        product_details_element = response.html.find('.product-details-list__container', first=True)
        format_details_element = response.html.find('.col-lg-5.col-sm-12.format-details-list__info', first=True)
        clone_details_element = response.html.find('.antibody-details-list__title', first=True)

        
        if product_details_element: #and format_details_element:
            product_details = product_details_element.text #skims the  details table for reactivity and antigen daata
            format_details=format_details_element.text
            clone_details=clone_details_element.text
           #print(product_details)
            #print(format_details)
            return product_details+"\n"+format_details+"\nClone: "+clone_details+"\n"
        else:
            return "Product details not found."

    # Main asynchronous function to handle the flow
    async def main():
        # Get search term from user input
        search_term = catalog_number  # Or dynamically ask for this via a dialog

        try:
            # Scrape the search results
            filtered_titles, filtered_links = await scrape_search_results(search_term)

            # Show a pop-up to allow the user to choose one of the titles
            selected_index = choose_option(filtered_titles)

            if selected_index is not None and 0 <= selected_index < len(filtered_titles):
                # Navigate to the selected product's page and scrape the details
                selected_link = filtered_links[selected_index]
                product_details = await scrape_product_details(selected_link)
                print(product_details)
                # list of  variables needed
                #dictionary_terms=["Antibody Specificity:", "Label:","Peak Channel (CyTEK)", "Clone","Target Species:", "Host Species:"]
                
                text="Title:"+filtered_titles[selected_index]+"\n"+product_details #add the page title to the data
                print(text)
                antibodydict={}
                #####declare a Specificity 
                start_keyword = "Title:" #using "title" instead of the  index  allows the string to be human readable 
                start_index = text.find(start_keyword)
                if start_index != -1:
                    start_index += len(start_keyword)  # Move to the end of the keyword
                    end_index = text.find('\n', start_index)  # Find the next newline character
                    title=text[start_index:end_index].strip()  # Strip removes extra spaces
                    parts = title.rsplit(' ', 1) #get whatever is at the end
                    antibodydict["Antibody Specificity"]=parts[-1]
                else:
                    print("Specificty  not found")
                  ######  ###Declare a label
                start_keyword = "Format:" #comes from concatenated format table
                start_index = text.find(start_keyword)
                if start_index != -1:
                    start_index += len(start_keyword)  # Move to the end of the keyword
                    end_index = text.find('\n', start_index)  # Find the next newline character
                    antibodydict["Label"] = text[start_index:end_index].strip()  # Strip removes extra spaces
                   
                else:
                    print("Label keyword not found")

                #### Declare a clone
                start_keyword = "Clone:"
                start_index = text.find(start_keyword)
                if start_index != -1:
                    start_index += len(start_keyword)  # Move to the end of the keyword
                    end_index = text.find('\n', start_index)  # Find the next newline character
                    antibodydict["Clone"] = text[start_index:end_index].strip()  # Strip removes extra spaces
                    
                else:
                    print("Reactivity keyword not found")
                #### Declare a target
                start_keyword = "Reactivity:"
                start_index = text.find(start_keyword)
                if start_index != -1:
                    start_index += len(start_keyword)  # Move to the end of the keyword
                    end_index = text.find('\n', start_index)  # Find the next newline character
                    antibodydict["Target Species"]= text[start_index:end_index].strip()  # Strip removes extra spaces
                   
                else:
                    print("Reactivity keyword not found")
                #### Declare a host
                start_keyword = "Isotype:"
                start_index = text.find(start_keyword)
                if start_index != -1:
                    start_index += len(start_keyword)  # Move to the end of the keyword
                    end_index = text.find('\n', start_index)  # Find the next newline character
                    antibodydict["Host Species"]= text[start_index:end_index].strip()  # Strip removes extra spaces
                    
                else:
                    None
                # Display product details
                #print(f"Product details for '{filtered_titles[selected_index]}':\n{product_details}")
                return antibodydict#[[selected_index],[product_details]]
            else:
                show_popup("Invalid selection or no selection made.")

        except Exception as e:
            # If an error occurs, display a pop-up message
            show_popup(f"An error occurred: {str(e)}")

    # Run the asyncio event loop
    return asyncio.run(main())

def TonboCytek_Scraper(catalog_number):
    # Asynchronous function to scrape the search results
    nest_asyncio.apply()
    def show_popup(message):
        root = tk.Tk()
        root.withdraw()  # Hide the main Tkinter window
        messagebox.showerror("Error", message)
        root.destroy()
    # Asynchronous function to scrape the product details
    async def scrape_product_details(link):
        session = AsyncHTMLSession()
        
        # Fetch the product details page
        response = await session.get(link)
        await response.html.arender(timeout=20)
    
        # Find the product details from 'product-details-list__container'
        product_details_element = response.html.find('.ResultTitle', first=True)
        http=product_details_element.find('a', first=True).attrs['href'] 
        
        if http:
            response = await session.get(http)
            await response.html.arender(timeout=20)
            product_details_element = response.html.find('.data-table', first=True)
            if product_details_element:
                return product_details_element.text
        else:
            return "Product details not found."
    
    # Main asynchronous function to handle the flow
    async def main():
        # Get search term from user input
        search_term = catalog_number
        search_url = f"https://cytekbio.com/search?q={search_term}"
        try:
            
            product_details = await scrape_product_details(search_url)
            
            # Display product details
           # print(f"Product details for '{search_term}':\n{product_details}")
            text=product_details #add the page title to the data
            #print(text)
            antibodydict={}
            #####declare a Specificity 
            start_keyword = "Name\n" #using "title" instead of the  index  allows the string to be human readable 
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find(' (', start_index)  # Find the next newline character
                title=text[start_index:end_index].strip()  # Strip removes extra spaces
                parts = title.rsplit(' ', 1) #get whatever is at the end
                antibodydict["Antibody Specificity"]=parts[-1]
            else:
                print("Specificty  not found")
              ######  ###Declare a label
            start_keyword = "Format\n" #comes from concatenated format table
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('\n', start_index)  # Find the next newline character
                antibodydict["Label"] = text[start_index:end_index].strip()  # Strip removes extra spaces
               
            else:
                print("Label keyword not found")

            #### Declare a clone
            start_keyword = "Clone\n"
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('\n', start_index)  # Find the next newline character
                antibodydict["Clone"] = text[start_index:end_index].strip()  # Strip removes extra spaces
                
            else:
                print("Reactivity keyword not found")
            #### Declare a target
            start_keyword = "Reactivity\n"
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('\n', start_index)  # Find the next newline character
                antibodydict["Target Species"]= text[start_index:end_index].strip()  # Strip removes extra spaces
               
            else:
                print("Reactivity keyword not found")
            #### Declare a host
            start_keyword = "Isotype\n"
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('\n', start_index)  # Find the next newline character
                antibodydict["Host Species"]= text[start_index:end_index].strip()  # Strip removes extra spaces
                
            else:
                None
            # Display product details
            #print(f"Product details for '{filtered_titles[selected_index]}':\n{product_details}")
           
            print(antibodydict)
            return antibodydict#[[selected_index],[product_details]]
            
        except Exception as e:
            # If an error occurs, display a pop-up message
            show_popup(f"An error occurred: {str(e)}")
    
    # Run the asyncio event loop
    return asyncio.run(main())
def Biolegend_Scraper(catalog_number):
    # Asynchronous function to scrape the search results
    nest_asyncio.apply()
    def show_popup(message):
        root = tk.Tk()
        root.withdraw()  # Hide the main Tkinter window
        messagebox.showerror("Error", message)
        root.destroy()
    # Asynchronous function to scrape the product details
    async def scrape_product_details(link):
        session = AsyncHTMLSession()
        
        # Fetch the product details page
        response = await session.get(link)
        await response.html.arender(timeout=20)
    
        # Find the product details from 'product-details-list__container'
        name_details_element = response.html.find('h1.col-xs-12.noPadding', first=True)
        clone_details_element = response.html.find('.col-xs-8.noPaddingLeft', first=True)
        product_details_element = response.html.find('.col-xs-12.col-sm-9.pull-right.noPadding', first=True)
        
        if product_details_element:
            return "Title:"+name_details_element.text+'\n'+clone_details_element.text+product_details_element.text
        else:
            return "Product details not found."
    
    # Main asynchronous function to handle the flow
    async def main():
        # Get search term from user input
        search_term = catalog_number
        search_url = f"https://www.biolegend.com/en-us/search-results?Keywords={search_term}"
        try:
            
            product_details = await scrape_product_details(search_url)
            
            # Display product details
           # print(f"Product details for '{search_term}':\n{product_details}")
            text=product_details #add the page title to the data
            print(text)
            antibodydict={}
            #####declare a Specificity 
            start_keyword = "Title:" #using "title" instead of the  index  allows the string to be human readable 
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find(' Antibody', start_index)  # Find the next newline character
                title=text[start_index:end_index].strip()  # Strip removes extra spaces
                parts = title.rsplit(' ', 1) #get whatever is at the end
                antibodydict["Antibody Specificity"]=parts[-1]
            else:
                print("Specificty  not found")
              ######  ###Declare a label
            start_keyword = "Title:" #comes from concatenated format table
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('anti', start_index)  # Find the next newline character
                antibodydict["Label"] = text[start_index:end_index].strip()  # Strip removes extra spaces
                if antibodydict["Label"]:
                   antibodydict["Peak Channel (CyTEK)"]=peak_channel(antibodydict["Label"])
            else:
                print("Label keyword not found")

            #### Declare a clone
            start_keyword = "Clone\n"
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find(' (', start_index)  # terminate this at the (see availible)
                antibodydict["Clone"] = text[start_index:end_index].strip()  # Strip removes extra spaces
                
            else:
                print("Reactivity keyword not found")
            #### Declare a target
            start_keyword = "Reactivity\n"
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('\n', start_index)  # Find the next newline character
                antibodydict["Target Species"]= text[start_index:end_index].strip()  # Strip removes extra spaces
               
            else:
                print("Reactivity keyword not found")
            #### Declare a host
            start_keyword = "Host Species\n"
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('\n', start_index)  # Find the next newline character
                antibodydict["Host Species"]= text[start_index:end_index].strip()  # Strip removes extra spaces
                
            else:
                None
            # Display product details
            #print(f"Product details for '{filtered_titles[selected_index]}':\n{product_details}")
          #  print(antibodydict)
            return antibodydict#[[selected_index],[product_details]]
            
        except Exception as e:
            # If an error occurs, display a pop-up message
            show_popup(f"An error occurred: {str(e)}")
    
    # Run the asyncio event loop
    return asyncio.run(main())
def Thermofisher_Scraper(catalog_number):
    # Asynchronous function to scrape the search results
    nest_asyncio.apply()
    def show_popup(message):
        root = tk.Tk()
        root.withdraw()  # Hide the main Tkinter window
        messagebox.showerror("Error", message)
        root.destroy()
    # Asynchronous function to scrape the product details
    async def scrape_product_details(link):
        session = AsyncHTMLSession()
        
        # Fetch the product details page
        response = await session.get(link)
        await response.html.arender(timeout=20)
        time.sleep(2)
        final_url = response.url
        print(f"Final URL: {final_url}")
        # Find the product details from 'product-details-list__container'
        name_details_element = response.html.find('h1.product-name', first=True)
        product_details_element = response.html.find('.product-detail-section.tested-app-section.new-tbl-design.product-specific-section.table-spec', first=True)
        
        if product_details_element:
            return "Title:"+name_details_element.text+'\n'+product_details_element.text
        else:
            return "Product details not found."
    
    # Main asynchronous function to handle the flow
    async def main():
        # Get search term from user input
        search_term = catalog_number
        search_url = f"https://www.thermofisher.com/antibody/product/{search_term}"
        print(search_url)
        try:
            
            product_details = await scrape_product_details(search_url)
            
            # Display product details
           # print(f"Product details for '{search_term}':\n{product_details}")
            text=product_details #add the page title to the data
            print(text)
            antibodydict={}
            #####declare a Specificity 
            start_keyword = "Title:" #using "title" instead of the  index  allows the string to be human readable 
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find(' '   , start_index)  # Find the next newline character
                title=text[start_index:end_index].strip()  # Strip removes extra spaces
                parts = title.rsplit(' ', 1) #get whatever is at the end
                antibodydict["Antibody Specificity"]=parts[-1]
            else:
                print("Specificty  not found")
              ######  ###Declare a label
            start_keyword = "Conjugate\n" #comes from concatenated format table
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('\n', start_index)  # Find the next newline character
                antibodydict["Label"] = text[start_index:end_index].strip()  # Strip removes extra spaces
                if antibodydict["Label"]:
                   antibodydict["Peak Channel (CyTEK)"]=peak_channel(antibodydict["Label"])
            else:
                print("Label keyword not found")

            #### Declare a clone
            start_keyword = "Clone\n"
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('\n', start_index)  # terminate this at the (see availible)
                antibodydict["Clone"] = text[start_index:end_index].strip()  # Strip removes extra spaces
                
            else:
                print("Reactivity keyword not found")
            #### Declare a target
            start_keyword = "Reactivity\n"
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('\n', start_index)  # Find the next newline character
                antibodydict["Target Species"]= text[start_index:end_index].strip()  # Strip removes extra spaces
               
            else:
                print("Reactivity keyword not found")
            #### Declare a host
            start_keyword = "Isotype\n"
            start_index = text.find(start_keyword)
            if start_index != -1:
                start_index += len(start_keyword)  # Move to the end of the keyword
                end_index = text.find('\n', start_index)  # Find the next newline character
                antibodydict["Host Species"]= text[start_index:end_index].strip()  # Strip removes extra spaces
                
            else:
                None
            # Display product details
            #print(f"Product details for '{filtered_titles[selected_index]}':\n{product_details}")
          #  print(antibodydict)
            return antibodydict#[[selected_index],[product_details]]
            
        except Exception as e:
            # If an error occurs, display a pop-up message
            show_popup(f"An error occurred: {str(e)}")
    
    # Run the asyncio event loop
    return asyncio.run(main())
def peak_channel(inputchannel): #external peak channel determining because I may move it 
    phrases = ["Brilliant Violet", "Brilliant Blue", "Cyanine","Brilliant Ultra Violet","Alexa Fluor 700"]

    # Corresponding abbreviations
    abbreviations = ["BV", "BB", "Cy","BUV","AF700"]
    inputchannel = inputchannel.replace("Â®", "")
    for phrase, abbreviation in zip(phrases, abbreviations):
            inputchannel = inputchannel.replace(phrase, abbreviation)
            inputchannel = inputchannel.replace("/", "-")
            
            print(inputchannel)
            if "BUV" or "BV" in inputchannel:inputchannel = inputchannel.replace(" ", "")
            inputchannel = inputchannel.replace("â„¢", "")
    #print(inputchannel+"input channel")
    pdf_path = "/Users/westtn/Downloads/Cytek Aurora 5L Fluorescent Guide.pdf"  # Replace with your PDF file path
    with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()  # Extract tables from each page
                if table:
                    # Loop through each row in the table
                    for row in table:
                        # Assuming search phrase is in the 4th or 5th column (index 3 or 4 in a zero-indexed list)
                        if row[3] and re.search(inputchannel, row[3], re.IGNORECASE) or row [4] and re.search(inputchannel, row[4], re.IGNORECASE) :
                            # Return the text from the 2nd column (index 1) of the same row
                            return row[1]  # Extracting text from 2nd column (index 1)
def determine_free_inventory():
    data=sheetlink.fetch_mouse_inventory()
    print(data[1])
def check_existing_inventory(dictionary): #check the existing inventory for duplicates
# Function to display Yes/No dialog
    data=sheetlink.fetch_mouse_inventory() #fetech all the data as a list of lists 
    
    def show_dialog(list_content):
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        result = messagebox.askyesno("Alert", f"Suspected Duplicate found: {list_content}\nContinue?")
        root.destroy()  # Destroy the Tk window
        if result:
            return True
        else:
            messagebox.showerror("Error", "Action Canceled")
            return False
    
    for lst in data:
        
            if (dictionary["Label"] and dictionary["Antibody Specificity"] and dictionary["Clone"] in lst) or dictionary["Catalog Number"] in lst:
     
                return show_dialog(lst)
    
    return False   # No match found