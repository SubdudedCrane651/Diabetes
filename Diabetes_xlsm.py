import openpyxl
from openpyxl import Workbook
from datetime import datetime
import sys
import pyodbc
import xlwings as xw


# Define the function to fetch data from Access and process it
def fetch_data_from_access(query, connection_string):
    cnxn = pyodbc.connect(connection_string, autocommit=True)
    crsr = cnxn.cursor()
    crsr.execute(query)
    rows = crsr.fetchall()
    crsr.close()
    cnxn.close()
    return rows


# Define the function to write data to Excel
def write_to_excel(sheet, start_column, rows, header, days):
    sheet[f"{start_column}3"] = header
    Datevar = []
    Average_Reading = []
    Timevar = []
    x = 4
    count = 0

    for row in rows:
        x += 1
        Datevar.append(row[0])
        sheet[f"{start_column}{x}"] = Datevar[count]
        Timevar.append(row[1])
        sheet[f"{chr(ord(start_column) + 1)}{x}"] = Timevar[count]
        Average_Reading.append(float(row[2]))
        sheet[f"{chr(ord(start_column) + 2)}{x}"] = Average_Reading[count]
        count += 1

# Main function to create the Excel file
def CreateDiabetes_xlsm():
    import xlwings as xw
    try:
            xl_app = xw.App(visible=False, add_book=False)
            wb = xl_app.books.open("C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm")

            run_macro = wb.app.macro('DeleteSelection')
            run_macro()

            wb.save()
            wb.close()

            xl_app.quit()

    except Exception as ex:
            template = "An exception of type {0} occurred. Arguments:\n{1!r}"
            message = template.format(type(ex).__name__, ex.args)
            print(message)
    
    connection_string = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=C:\Users\rchrd\Documents\Richard\Richards_Health.mdb;'
    )

    # Prepare the workbook
    workbook = openpyxl.load_workbook(
        "C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm", read_only=False, keep_vba=True
    )
    sheet = workbook.active
    
    # Add the number of days to cell "N2" (as an integer)
    sheet["N2"] = int(days)

    # Call and write Mourning Glucose Reading
    mourning_query = "{CALL Mourning_Gluclose_Reading}"
    rows = fetch_data_from_access(mourning_query, connection_string)
    write_to_excel(sheet, "A", rows, "Richard's Mourning Glucose Reading", days)

    # Call and write Afternoon Glucose Reading
    afternoon_query = "{CALL Afternoon_Glucose_Reading}"
    rows = fetch_data_from_access(afternoon_query, connection_string)
    write_to_excel(sheet, "E", rows, "Richard's Afternoon Glucose Reading", days)

    # Call and write Evening Glucose Reading
    evening_query = "{CALL Evening_Gluclose_Reading}"
    rows = fetch_data_from_access(evening_query, connection_string)
    write_to_excel(sheet, "I", rows, "Richard's Evening Glucose Reading", days)
    
    # Save the workbook
    workbook.save(filename="C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm")

    # Use xlwings to open the file
    xl_app = xw.App(visible=True, add_book=False)
    wb = xl_app.books.open("C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm")
    #xl_app.quit()


# Command-line argument for days
try:
    days = sys.argv[1]
except IndexError:
    days = 22

# Execute the function
CreateDiabetes_xlsm()