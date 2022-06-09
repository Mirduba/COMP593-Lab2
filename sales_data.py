# import functions
from sys import argv, exit
import os, re
from datetime import date
import pandas as ps

# Function to check argv value and file exists
def get_csvfile():
    if len(argv) >= 2:
        salescsv_path = argv[1]  # Save path in the variable
        if os.path.isfile(salescsv_path): # Check the path has valid file
            return salescsv_path
        else:  # Show warning as path doesn't have file
            print("\nWarning! The file doesn't exists. \n")
            exit("Script aborted\n")
    else: # Show warning as csv file path not provided 
        print("\nWarning! The csv file path is not provided\n")
        exit("Script aborted \n")

# Function to create order directory name and path to save extracted sales data 
def order_directory(getcsv_filepath): # Send path variable
    sales_csvpath = os.path.dirname(getcsv_filepath) # Save the sales path
    current_date = date.today().isoformat() # Get current date in iso format
    orderfolder_namepath = os.path.join(sales_csvpath, 'Orders_' + current_date) # join path and date for order folder name
    if not os.path.exists(orderfolder_namepath): # Checks whether folder exists or not
        os.makedirs(orderfolder_namepath) # if not makes a folder
    return orderfolder_namepath # returns the order name

# Function to split sales data into order files
def split_salescsv(getcsv_filepath, orderfolder_path): # Csv file path to extract and 2nd parameter to save the extract data
    read_salesdatacsv = ps.read_csv(getcsv_filepath) # Read the csv file and save in variable
    # Remove few columns from sales csv file
    read_salesdatacsv.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)
    # Add column for TOTAL PRICE in csv file and calculate item quantity and price
    read_salesdatacsv.insert(7, 'TOTAL PRICE', read_salesdatacsv['ITEM QUANTITY'] * read_salesdatacsv['ITEM PRICE'])
    # Iterate condition for same ORDER ID in sales data csv
    for order_id, order_csv in read_salesdatacsv.groupby('ORDER ID'): 
        order_csv.drop(columns=['ORDER ID'], inplace=True) # Remove ORDER ID column in new order csv file
        order_csv.sort_values(by='ITEM NUMBER', inplace=True) # Sort value by ITEM NUMBER
        grandtotal = order_csv['TOTAL PRICE'].sum() # Add the TOTAL PRICE column
        grandtotal_df = ps.DataFrame({'ITEM PRICE':['GRAND TOTAL'], 'TOTAL PRICE': [grandtotal]}) # Add row for GRAND TOTAL
        order_csv = ps.concat([order_csv, grandtotal_df]) # Append the GRAND TOTAL to the order csv file

        # Code to determine the csv file name for extract data (in Orders folder)
        customer_name = order_csv['CUSTOMER NAME'].values[0] # Get customer name 
        customer_name =re.sub(r'\W', '', customer_name) # RE removes non alphabets character in the customer name
        ordercsv_filename = 'Order' + str(order_id) + '_' + customer_name + '.xlsx' # Join order id and customer name for order csv file
        ordercsv_filepath = os.path.join(orderfolder_path, ordercsv_filename)

        # Code to rename the excel sheet
        sheet_name = 'Order #' + str(order_id)
        
        # Code to format the excel sheets
        excel_writer = ps.ExcelWriter(ordercsv_filepath, engine='xlsxwriter') 
        order_csv.to_excel(excel_writer, index=False, sheet_name=sheet_name)
        excel_book = excel_writer.book
        excel_sheet = excel_writer.sheets[sheet_name]
        excel_sheet.set_zoom(120) # Number depends on screen size
        money_format = excel_book.add_format({'num_format': '$#,##0.00'})
        excel_sheet.set_column('A:A',12) # ORDER DATE
        excel_sheet.set_column('B:E',14) # ITEM NUMBER, PRODUCT LINE, PRODUCT CODE, ITEM QUANTITY
        excel_sheet.set_column('I:I',24) # CUSTOMER NAME
        excel_sheet.set_column('F:G',12,money_format) # ITEM PRICE, TOTAL PRICE
        excel_writer.save() # Save the workbook


# Function calling

getcsv_filepath = get_csvfile()
#print("\n The path for csv file is: \n", getcsv_filepath) # Prints path

orderfolder_path = order_directory(getcsv_filepath)
#print("\n Order folder Created: ", orderfolder_path) # Prints order folder path

split_salescsv(getcsv_filepath,orderfolder_path)
