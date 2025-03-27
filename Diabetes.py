import openpyxl
from openpyxl import Workbook
from unicodedata import decimal
import numpy as np
from datetime import datetime
import pyodbc
import sys
import pyodbc
import win32com.client

def run_access_subroutine(days):
    # Path to your Access database file
    database_path = r"C:\Users\rchrd\Documents\Richard\Richards_Health.mdb"
    
    try:
        # Create a COM object to interact with Access
        access_app = win32com.client.Dispatch("Access.Application")

        # Open the Access database
        access_app.OpenCurrentDatabase(database_path)

        # Run the Access VBA subroutine and pass the parameter dynamically
        access_app.Run("CreateDiabetesQrys", int(days))

        # Close the database
        access_app.CloseCurrentDatabase()

        # Quit the Access application
        access_app.Quit()

        print(f"Subroutine 'CreateDiabetesQrys' executed successfully for {days} days.")
    except Exception as e:
        print(f"Error executing subroutine: {e}")

# Example usage: Run the subroutine with 7 days as the parameter

# Example usage: Pass the number of days dynamically

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

cnxn = pyodbc.connect(connection_string, autocommit=True)
crsr = cnxn.cursor()

try:
   DateStart=sys.argv[1]
except:
   DateStart=22
   
run_access_subroutine(int(DateStart))

# SELECT Diabetes.Datevar, Diabetes.Timevar, Diabetes.Reading
# FROM Diabetes
# WHERE (((Diabetes.Datevar)>#8/28/2022#) And ((Diabetes.Timevar)<=#12/30/1899 7:0:0#) And ((Year(Diabetes.Datevar))=2022))
# ORDER BY Diabetes.Datevar DESC;

SQL1="""\
{CALL Mourning_Gluclose_Reading}
"""

crsr.execute(SQL1)

rows = crsr.fetchall()

print(rows)

workbook = openpyxl.load_workbook("C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm",read_only=False,keep_vba=True)

sheet = workbook.active

Datevar = []
Average_Reading = []
Timevar = []
x=4
count=0

sheet["A3"]="Richard's Mourning Glucose Reading"

for row in rows:
   x+=1
   Datevar.append((row[0]))
   sheet["A"+str(x)]=Datevar[count]
   Timevar.append((row[1]))
   sheet["B"+str(x)]=Timevar[count]
   Average_Reading.append(float(row[2]))
   sheet["C"+str(x)]=Average_Reading[count]
   count+=1

cnxn = pyodbc.connect(connection_string, autocommit=True)
crsr = cnxn.cursor()

sheet["n2"]=DateStart

# SELECT Diabetes.Datevar, Diabetes.Timevar, Diabetes.Reading
# FROM Diabetes
# WHERE (((Diabetes.Datevar)>#8/28/2022#) AND ((Diabetes.Timevar)>=#12/30/1899 7:0:0# And (Diabetes.Timevar)<#12/30/1899 17:0:0#))
# ORDER BY Diabetes.Datevar DESC;

SQL1="""\
{CALL Afternoon_Glucose_Reading}
"""

crsr.execute(SQL1)

rows = crsr.fetchall()

print(rows)

Datevar = []
Average_Reading = []
Timevar = []
x=4
count=0

sheet["E3"]="Richard's Afternoon Glucose Reading"

for row in rows:
   x+=1
   Datevar.append((row[0]))
   sheet["E"+str(x)]=Datevar[count]
   Timevar.append((row[1]))
   sheet["F"+str(x)]=Timevar[count]
   Average_Reading.append(float(row[2]))
   sheet["G"+str(x)]=Average_Reading[count]
   count+=1   

cnxn = pyodbc.connect(connection_string, autocommit=True)
crsr = cnxn.cursor()

# SELECT Diabetes.Datevar, Diabetes.Timevar, Diabetes.Reading
# FROM Diabetes
# WHERE (((Diabetes.Datevar)>#8/28/2022#) And ((Diabetes.Timevar)>=#12/30/1899 17:0:0#) And ((Year(Diabetes.Datevar))=2022))
# ORDER BY Diabetes.Datevar DESC;

SQL1="""\
{CALL Evening_Gluclose_Reading}
"""

crsr.execute(SQL1)

rows = crsr.fetchall()

print(rows)

Datevar = []
Average_Reading = []
Timevar = []
x=4
count=0

sheet["i3"]="Richard's Evening Glucose Reading"

for row in rows:
   x+=1
   Datevar.append((row[0]))
   sheet["i"+str(x)]=Datevar[count]
   Timevar.append((row[1]))
   sheet["j"+str(x)]=Timevar[count]
   Average_Reading.append(float(row[2]))
   sheet["k"+str(x)]=Average_Reading[count]
   count+=1      

workbook.save(filename="C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm")

xl_app = xw.App(visible=True, add_book=False)
wb = xl_app.books.open("C:\\Users\\rchrd\\Documents\\Richard\\Diabetes.xlsm")
   
   