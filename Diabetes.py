import openpyxl
from openpyxl import Workbook
from datetime import datetime
import sys
import pyodbc
import tkinter as tk
from tkinter import messagebox
import win32com.client


# Define the function to update queries
def update_queries():
    try:
        # Get the number of days from the entry widget
        days = int(entry.get())

        # Connect to the Access application
        db_path = r"C:\Users\rchrd\Documents\Richard\Richards_Health.mdb"  # Update with your database path
        access_app = win32com.client.Dispatch("Access.Application")
        access_app.OpenCurrentDatabase(db_path)

        # Function to delete a query
        def delete_query(query_name):
            try:
                access_app.CurrentDb().QueryDefs.Delete(query_name)
            except Exception as e:
                print(f"Error deleting query '{query_name}': {e}")

        # Function to create a query
        def create_query(query_name, query_sql):
            try:
                delete_query(query_name)  # Delete the existing query
                access_app.CurrentDb().CreateQueryDef(query_name, query_sql)
            except Exception as e:
                print(f"Error creating query '{query_name}': {e}")

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

        # Create or update the queries
        create_query("Mourning_Gluclose_Reading", mourning_sql)
        create_query("Afternoon_Glucose_Reading", afternoon_sql)
        create_query("Evening_Gluclose_Reading", evening_sql)

        # Close the Access database
        access_app.CloseCurrentDatabase()
        access_app.Quit()

        # Show success message
        messagebox.showinfo("Success", "Queries updated successfully!")
        CreateDiabetes_xlsm(days)
        quit_application()
    except ValueError:
        messagebox.showerror("Error", "Please enter a valid number of days.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def quit_application():
    window.destroy()  # Destroy the tkinter window
    sys.exit()        # Exit the script
 
def CreateDiabetes_xlsm(days):
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

# Create the tkinter GUI
window = tk.Tk()
window.title("Update Queries")

# Bind the quit function to the window's close button (X)
window.protocol("WM_DELETE_WINDOW", quit_application)

# Create a label for the entry field
label = tk.Label(window, text="Enter number of days:")
label.pack(pady=5)

# Create the entry field
entry = tk.Entry(window)
entry.pack(pady=5)

# Create the button to trigger the update function
button = tk.Button(window, text="Update Queries", command=update_queries)
button.pack(pady=10)

# Run the tkinter GUI event loop
window.mainloop()