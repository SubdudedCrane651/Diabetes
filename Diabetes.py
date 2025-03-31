import openpyxl
from openpyxl import Workbook
from datetime import datetime
import sys
import pyodbc
import tkinter as tk
from tkinter import messagebox
import win32com.client
import Diabetes_xlsm as db


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
        db.CreateDiabetes_xlsm(days)
        quit_application()
    except ValueError:
        messagebox.showerror("Error", "Please enter a valid number of days.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def quit_application():
    window.destroy()  # Destroy the tkinter window
    sys.exit()        # Exit the script
 
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