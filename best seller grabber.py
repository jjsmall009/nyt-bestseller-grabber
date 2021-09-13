#################################################################################################
"""New York Times Bestsellers List Grabber

Author: JJ Small
Date: 9/6/21
Company: Anacortes Public Library

This script will create an excel spreadsheet for the sole purpose of printing out and displaying
in the library. As of now it only processes fiction and nonfiction data.

Using the NYT API we grab data from certain endpoints and update a template excel file with the 
current NYT bestseller information. If you're reading this then I must be dead...
"""
#################################################################################################

from xlsxwriter import Workbook
from datetime import datetime
from io import BytesIO
from requests import get

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

    request = get(url=book_url)
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

def create_spreadsheet(file_name):
    
    bold = wb.add_format({"bold": True})

    return wb

def update_spreadsheet(book_data, sheet):
    """
    Update the specified sheet with the new data we get from the list API. Basically it
    iterates over the list of books and will add in all of the new books data. 

    Parameters:
        book_data (list): The formatted list of books, where book = dict
        sheet_name (sheet): The sheet object (Fiction, etc.)
    """

    print(f"Processing {sheet} data.......")
    sheet.set_column(0,0, width=16.25)
    sheet.set_column(16,16, width=14.2)
    merge_format = wb.add_format({"valign": "vcenter", "text_wrap": True})

    # Add first row header
    sheet.set_row(0, height=40)
    sheet.merge_range("A1:S1", "")
    header_bold = wb.add_format({
        "font_name": "Cooper Hewitt Book",
        "font_size": 28,
        "font_color": "#44546A",
        "bold": True
    })

    header_style = wb.add_format({
        "font_name": "Cooper Hewitt Book",
        "font_size": 28,
        "font_color": "#44546A"
    })
    sheet.write_rich_string(
        "A1", 
        header_bold,
        "BEST SELLERS",
        header_style,
        " In the Anacortes Public Library Collection")

    # Add second row with current date
    sheet.set_row(1, height=95)
    sheet.merge_range("A2:Q2", "")
    date_header = wb.add_format({
        "font_name": "Cooper Hewitt Book",
        "font_size": 22,
        "font_color": "#4472C4",
        "align": "center",
        "valign": "vcenter"
    })
    sheet.write(
        "A2", 
        f"The New York Times - Hardcover {sheet.get_name()} - {DATE.strftime('%B %d, %Y')}",
        date_header)

    # For each book in our list, add in the data to the proper cells and format accordingly
    for index, book in enumerate(book_data):
        status_style = wb.add_format({
            "font_name": "Cooper Hewitt Book",
            "font_size": 22,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True
        })

        book_style = wb.add_format({
            "font_name": "Cooper Hewitt Book",
            "font_size": 22,
        })

        title_bold = wb.add_format({
            "font_name": "Cooper Hewitt Book",
            "font_size": 22,
            "bold": True,
        })
        row = index + 2
        sheet.set_row(row, height=125)
        sheet.merge_range(row, 2, row, 15, "", merge_format)

        sheet.write(row, 0, "-", status_style) # This is the "Have/On Order" field and we have to manually do this later
        sheet.write(row, 1, book["rank"], status_style)
        sheet.write_rich_string( # Format the book information with a bold title and such
            row, 
            2, 
            title_bold,
            f"{book['title']}",
            book_style,
            f" by {book['author']}\n{book['description']}",
            merge_format)
        
        # Add in the image
        r = get(url=book["book_image"])
        url = "google.com"
        image_file = BytesIO(r.content)
        sheet.insert_image(row, 16, url, {"image_data": image_file, "x_scale": .32, "y_scale": .32})
    print("Finished processing data...")
        

# Craft our endpoints
fiction_endpoint = f"{API_URL}lists/current/hardcover-fiction.json?api-key={API_KEY}"
non_endpoint = f"{API_URL}lists/current/hardcover-nonfiction.json?api-key={API_KEY}"

# Get and process all of the data and do the stuff
fiction_data = get_list_data(fiction_endpoint)
non_data = get_list_data(non_endpoint)

# Open up our template spreadsheet and update it with the current list of bestsellers
new_file = f"results/{DATE.strftime('%B %d')} New York Bestsellers.xlsx"
wb = Workbook(new_file)
f_sheet = wb.add_worksheet("Fiction")
nf_sheet = wb.add_worksheet("Nonfiction")

update_spreadsheet(fiction_data, f_sheet)
update_spreadsheet(non_data, nf_sheet)

wb.close()
