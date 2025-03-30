import openpyxl
from openpyxl import Workbook
from datetime import datetime
import sys
import pyodbc

try:
   days=sys.argv[1]
except:
   days=22
   
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

    cnxn = pyodbc.connect(connection_string, autocommit=True)
    crsr = cnxn.cursor()        

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
    
CreateDiabetes_xlsm()