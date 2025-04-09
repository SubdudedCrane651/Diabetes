import openpyxl
from openpyxl import Workbook
from datetime import datetime
import sys
import pyodbc
import tkinter as tk
from tkinter import messagebox
import win32com.client
import Diabetes_xlsm as db

# Function to retrieve the default days from the Access database
def get_default_days():
    try:
        db_path = r"C:\Users\rchrd\Documents\Richard\Richards_Health.mdb"  # Update with your database path
        conn = pyodbc.connect(f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};")
        cursor = conn.cursor()

        cursor.execute("SELECT TOP 1 Days FROM Days")
        row = cursor.fetchone()
        default_days = row[0] if row else 7  # Default to 7 if the table is empty
        
        cursor.close()
        conn.close()
        return default_days

    except Exception as e:
        print(f"Error retrieving default days: {e}")
        return 7  # Default to 7 on error

# Function to update queries
def update_queries():
    try:
        # Get the number of days from the entry widget
        days = int(entry.get())

        # Connect to the Access application
        db_path = r"C:\Users\rchrd\Documents\Richard\Richards_Health.mdb"  # Update with your database path
        access_app = win32com.client.Dispatch("Access.Application")
        access_app.OpenCurrentDatabase(db_path)

        # Update the first record instead of adding a new one
        conn = pyodbc.connect(f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};")
        cursor = conn.cursor()

        cursor.execute("UPDATE Days SET Days = ? WHERE ID = (SELECT MIN(ID) FROM Days)", days)
        conn.commit()

        cursor.close()
        conn.close()

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
        def delete_query(query_name):
            try:
                access_app.CurrentDb().QueryDefs.Delete(query_name)
            except Exception:
                pass

        def create_query(query_name, query_sql):
            try:
                delete_query(query_name)
                access_app.CurrentDb().CreateQueryDef(query_name, query_sql)
            except Exception as e:
                print(f"Error creating query '{query_name}': {e}")

        create_query("Mourning_Gluclose_Reading", mourning_sql)
        create_query("Afternoon_Glucose_Reading", afternoon_sql)
        create_query("Evening_Gluclose_Reading", evening_sql)

        # Close Access database
        access_app.CloseCurrentDatabase()
        access_app.Quit()

        # Show success message
        messagebox.showinfo("Success", "Queries updated successfully!")
        db.CreateDiabetes_xlsm(days)
        quit_application()

    except ValueError:
        messagebox.showerror("Error", "Please enter a valid number of days.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def quit_application():
    window.destroy()
    sys.exit()

# Create tkinter GUI
window = tk.Tk()
window.title("Update Queries")

# Bind quit function to window close button (X)
window.protocol("WM_DELETE_WINDOW", quit_application)

# Create label for entry field
label = tk.Label(window, text="Enter number of days:")
label.pack(pady=5)

# Get default value from the database
default_days = get_default_days()

# Create entry field with default value
entry = tk.Entry(window)
entry.insert(0, str(default_days))  # Set default value
entry.pack(pady=5)

# Create button to trigger update function
button = tk.Button(window, text="Update Queries", command=update_queries)
button.pack(pady=10)

# Run tkinter GUI event loop
window.mainloop()