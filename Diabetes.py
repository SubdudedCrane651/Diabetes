import openpyxl
from openpyxl import Workbook
from unicodedata import decimal
import numpy as np
from datetime import datetime
import pyodbc
import sys
import pyodbc
import win32com.client

# Connect to the Access application
db_path = r"C:\Users\rchrd\Documents\Richard\Richards_Health.mdb"  # Replace with the path to your database
access_app = win32com.client.Dispatch("Access.Application")
access_app.OpenCurrentDatabase(db_path)

# Specify the number of days for filtering
days = 7  # Adjust as needed

# Function to delete a query
def delete_query(query_name):
    try:
        access_app.CurrentDb().QueryDefs.Delete(query_name)
    except Exception as e:
        print(f"Error deleting query '{query_name}': {e}")

# Function to create a query
def create_query(query_name, query_sql):
    try:
        delete_query(query_name)  # First delete the query if it exists
        access_app.CurrentDb().CreateQueryDef(query_name, query_sql)
    except Exception as e:
        print(f"Error creating query '{query_name}': {e}")

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
   days=sys.argv[1]
except:
   days=22
   
# Define SQL for each query
mourning_sql = f"""
SELECT Diabetes.Datevar, Diabetes.Timevar, Diabetes.Reading
FROM Diabetes
WHERE Diabetes.Datevar >= Date() - {days}
      AND Diabetes.Timevar <= #11:59:00 AM#
      AND YEAR(Diabetes.Datevar) = 2025
ORDER BY Diabetes.Datevar DESC;
"""

afternoon_sql = f"""
SELECT Diabetes.Datevar, Diabetes.Timevar, Diabetes.Reading
FROM Diabetes
WHERE Diabetes.Datevar >= Date() - {days}
      AND Diabetes.Timevar >= #11:59:00 AM#
      AND Diabetes.Timevar <= #5:00:00 PM#
      AND YEAR(Diabetes.Datevar) = 2025
ORDER BY Diabetes.Datevar DESC;
"""

evening_sql = f"""
SELECT Diabetes.Datevar, Diabetes.Timevar, Diabetes.Reading
FROM Diabetes
WHERE Diabetes.Datevar >= Date() - {days}
      AND Diabetes.Timevar > #5:00:00 PM#
      AND YEAR(Diabetes.Datevar) = 2025
ORDER BY Diabetes.Datevar DESC;
"""   
   
# Execute the functions
create_query("Mourning_Gluclose_Reading", mourning_sql)
create_query("Afternoon_Glucose_Reading", afternoon_sql)
create_query("Evening_Gluclose_Reading", evening_sql)

# Close the Access application
access_app.CloseCurrentDatabase()
access_app.Quit()

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

sheet["n2"]=days

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
   
   