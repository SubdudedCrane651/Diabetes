import openpyxl
import pyodbc
import xlwings as xw
import sys
from datetime import datetime

# Function to fetch data from Access
def fetch_data_from_access(query, connection_string):
    cnxn = pyodbc.connect(connection_string, autocommit=True)
    crsr = cnxn.cursor()
    crsr.execute(query)
    rows = crsr.fetchall()
    crsr.close()
    cnxn.close()
    return rows

# Function to calculate and write morning (2 AM) averages per date
def write_morning_avg_to_excel(sheet, start_column, rows, header):
    sheet[f"{start_column}3"] = header
    daily_avg = {}

    for row in rows:
        date_str = row[0]
        reading = round(float(row[2]), 1)
        if date_str not in daily_avg:
            daily_avg[date_str] = []
        daily_avg[date_str].append(reading)

    x = 5
    for date, readings in daily_avg.items():
        sheet[f"{start_column}{x}"] = date
        sheet[f"{chr(ord(start_column) + 1)}{x}"] = "2:00 AM"
        sheet[f"{chr(ord(start_column) + 2)}{x}"] = round(sum(readings) / len(readings), 1) if readings else None
        x += 1

# Function to calculate and write evening (10 PM) averages per date
def write_evening_avg_to_excel(sheet, start_column, rows, header):
    sheet[f"{start_column}3"] = header
    daily_avg = {}

    for row in rows:
        date_str = row[0]
        reading = round(float(row[2]), 1)
        if date_str not in daily_avg:
            daily_avg[date_str] = []
        daily_avg[date_str].append(reading)

    x = 5
    for date, readings in daily_avg.items():
        sheet[f"{start_column}{x}"] = date
        sheet[f"{chr(ord(start_column) + 1)}{x}"] = "10:00 PM"
        sheet[f"{chr(ord(start_column) + 2)}{x}"] = round(sum(readings) / len(readings), 1) if readings else None
        x += 1

# Function to calculate and write afternoon (12 PM & 2 PM) averages per date
def write_afternoon_avg_to_excel(sheet, start_column, rows, header):
    sheet[f"{start_column}3"] = header
    daily_avg = {}

    for row in rows:
        date_str = row[0]
        time_str = row[1]
        reading = round(float(row[2]), 1)

        if date_str not in daily_avg:
            daily_avg[date_str] = {"12:00 PM": [], "2:00 PM": []}

        time_obj = time_str.time()

        #if datetime.strptime("13:00:00", "%H:%M:%S").time() <= time_obj < datetime.strptime("21:00:00", "%H:%M:%S").time():
        #    daily_avg[date_str]["2:00 PM"].append(reading)
            
        if datetime.strptime("9:00:00", "%H:%M:%S").time() <= time_obj < datetime.strptime("12:00:00", "%H:%M:%S").time():
            daily_avg[date_str]["12:00 PM"].append(reading)
       
        if datetime.strptime("12:00:00", "%H:%M:%S").time() <= time_obj < datetime.strptime("18:00:00", "%H:%M:%S").time():    
            daily_avg[date_str]["2:00 PM"].append(reading)

    x = 5
    for date, readings in daily_avg.items():
        sheet[f"{start_column}{x}"] = date
        sheet[f"{chr(ord(start_column) + 1)}{x}"] = "2:00 PM"
        sheet[f"{chr(ord(start_column) + 2)}{x}"] = round(sum(readings["2:00 PM"]) / len(readings["2:00 PM"]), 1) if readings["2:00 PM"] else None

        x += 1

        sheet[f"{start_column}{x}"] = date
        sheet[f"{chr(ord(start_column) + 1)}{x}"] = "12:00 PM"
        sheet[f"{chr(ord(start_column) + 2)}{x}"] = round(sum(readings["12:00 PM"]) / len(readings["12:00 PM"]), 1) if readings["12:00 PM"] else None

        x += 1

# Main function to create the Excel file
def CreateDiabetes_xlsm(days):
    try:
        xl_app = xw.App(visible=False, add_book=False)
        wb = xl_app.books.open("C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm")

        run_macro = wb.app.macro('DeleteSelection')
        run_macro()

        wb.save()
        wb.close()
        xl_app.quit()

    except Exception as ex:
        print(f"Error running macro: {ex}")

    connection_string = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=C:\Users\rchrd\Documents\Richard\Richards_Health.mdb;'
    )

    workbook = openpyxl.load_workbook("C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm", read_only=False, keep_vba=True)
    sheet = workbook.active
    
    sheet["N2"] = int(days)

    morning_query = "{CALL Mourning_Gluclose_Reading}"
    rows = fetch_data_from_access(morning_query, connection_string)
    print(rows)
    write_morning_avg_to_excel(sheet, "A", rows, "Richard's Morning Glucose Reading")

    afternoon_query = "{CALL Afternoon_Glucose_Reading}"
    rows = fetch_data_from_access(afternoon_query, connection_string)
    print(rows)
    write_afternoon_avg_to_excel(sheet, "E", rows, "Richard's Afternoon Glucose Reading")

    evening_query = "{CALL Evening_Gluclose_Reading}"
    rows = fetch_data_from_access(evening_query, connection_string)
    print(rows)
    write_evening_avg_to_excel(sheet, "I", rows, "Richard's Evening Glucose Reading")

    workbook.save(filename="C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm")

    xl_app = xw.App(visible=True, add_book=False)
    wb = xl_app.books.open("C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm")

try:
    days = sys.argv[1]
    CreateDiabetes_xlsm(days)
except IndexError:
    days = 14
    CreateDiabetes_xlsm(days)