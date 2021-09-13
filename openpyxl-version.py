#################################################################################################
"""New York Times Bestsellers List Grabber

Author: JJ Small
Date: 9/5/21
Company: Anacortes Public Library

This script will create an excel spreadsheet for the sole purpose of printing out and displaying
in the library. As of now it only processes fiction and nonfiction data.

Using the NYT API we grab data from certain endpoints and update a template excel file with the 
current NYT bestseller information. If you're reading this then I must be dead...
"""
#################################################################################################

from datetime import datetime
import io
import openpyxl
import requests
from requests import HTTPError

# Specific to jjsmall009 NYT API account
API_KEY = "gibberish"
API_URL = "https://api.nytimes.com/svc/books/v3/"
DATE = datetime.today()
 
def get_list_data(book_url):
    """
    Fetch all book data from the specified url and then grab the data we want from it. 

    Parameters:
        book_url (string): The API url based on the type of book you're requesting. Fiction, etc.
    
    Returns:
        List: Returns a list of books, where each book is a dictionary of certain parameters and values
    """

    request = requests.get(url=book_url)
    request.raise_for_status()

    try:
        print("Attempting to grab data from the NYT API...")
        data = request.json()["results"]["books"]
    except KeyError as e:
        # If for some reason we get data back that isn't in the correct format we'll crash
        print("Oops... the data isn't correct. Use the correct book type")
        print(e)
        exit()
    else:
        # Grab the data we want and not all of the other useless junk
        final_data = []
        keys = ["rank", "description", "title", "author", "book_image"]
        for book in data:
            book_dict_data = {info:book[info] for info in keys}
            final_data.append(book_dict_data)
        
        print("Finished grabbing data....")
        return final_data


def open_spreadsheet(file_name):
    """
    Open up a spreadsheet and use it for updating the data
    
    Parameters:
        file_name (string): The name of the template excel file

    Returns:
        Openpyxl Workbook: Returns this object full of fancy excel data
    
    """

    return openpyxl.load_workbook(filename=file_name)


def update_spreadsheet(book_data, workbook, sheet_name):
    """
    Update the specified workbook object with the new data we get from the list API. Basically it
    iterates over the list of books and will modify the matching cells in the specified sheet to 
    contain the new data. We later write the modified workbook to a new file for todays date.

    Parameters:
        book_data (list): The formatted list of books, where book = dict
        workbook (Openpyxl Workbook): The template excel file in workbook format
        sheet_name (string): The sheet name corresponds to the type of list (Fiction, etc.)
    """

    print(f"Processing {sheet_name} data.......")
    sheet = workbook[sheet_name]

    # Update second row with current date
    sheet["A2"] = f"The New York Times - Hardcover {sheet_name} - {DATE.strftime('%B %d, %Y')}"

    # For each row in the sheet, starting at the third row, update the contents of the row with
    # the new book data we pulled in from the API.
    for index, row in enumerate(sheet.iter_rows(min_row=3)):
        book = book_data[index]
        row[0].value = "" # This is the "Have/On Order" field and we have to manually do this later
        row[1].value = book["rank"]
        row[2].value = f"{book['title']} by {book['author']}\n{book['description']}"
        
        # Add in the image
        r = requests.get(url=book["book_image"])
        image_file = io.BytesIO(r.content)
        img = openpyxl.drawing.image.Image(image_file)
        img.width = 110
        img.height = 165
        spot = f"L{index + 3}"
        sheet.add_image(img, spot)
    print("Finished processing data...")
        

# Craft our endpoints
fiction_endpoint = f"{API_URL}lists/current/hardcover-fiction.json?api-key={API_KEY}"
non_endpoint = f"{API_URL}lists/current/hardcover-nonfiction.json?api-key={API_KEY}"

# Get and process all of the data and do the stuff
fiction_data = get_list_data(fiction_endpoint)
non_data = get_list_data(non_endpoint)

# Open up our template spreadsheet and update it with the current list of bestsellers
workbook = open_spreadsheet("TEMPLATE.xlsx")
update_spreadsheet(fiction_data, workbook, "Fiction")
update_spreadsheet(non_data, workbook, "Nonfiction")

# Write to new file with the new data
new_file = f"{DATE.strftime('%B %d')} New York Bestsellers.xlsx"
workbook.save(filename=new_file)
workbook.close()