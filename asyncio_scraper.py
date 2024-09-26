import asyncio
from requests_html import AsyncHTMLSession
import tkinter as tk
from tkinter import simpledialog, messagebox
import nest_asyncio

# Apply nest_asyncio to allow asyncio.run within an existing event loop
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
    search_url = f"https://www.bdbiosciences.com/en-us/search-results?searchKey={search_term}"
    
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
    
    if product_details_element:
        product_details = product_details_element.text
        return product_details
    else:
        return "Product details not found."

# Main asynchronous function to handle the flow
async def main():
    # Get search term from user input
    search_term = "552775"  # Or dynamically ask for this via a dialog

    try:
        # Scrape the search results
        filtered_titles, filtered_links = await scrape_search_results(search_term)

        # Show a pop-up to allow the user to choose one of the titles
        selected_index = choose_option(filtered_titles)

        if selected_index is not None and 0 <= selected_index < len(filtered_titles):
            # Navigate to the selected product's page and scrape the details
            selected_link = filtered_links[selected_index]
            product_details = await scrape_product_details(selected_link)
            
            # Display product details
            print(f"Product details for '{filtered_titles[selected_index]}':\n{product_details}")
        else:
            show_popup("Invalid selection or no selection made.")

    except Exception as e:
        # If an error occurs, display a pop-up message
        show_popup(f"An error occurred: {str(e)}")

# Run the asyncio event loop
asyncio.run(main())
