from unicodedata import decimal
import numpy as np
from datetime import datetime
import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D  # Required for 3D plotting

# --------------------------
# Database connection setup
# --------------------------
connection_string = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\rchrd\Documents\Richard\Richards_Health.mdb;'
)
cnxn = pyodbc.connect(connection_string, autocommit=True)
crsr = cnxn.cursor()

# Optional: List available tables (for debugging)
for i in crsr.tables(tableType='TABLE'):
    print(i.table_name)

# --------------------------
# Execute the query (stored procedure)
# --------------------------
SQL = """\
{CALL AverageControl1}
"""
crsr.execute(SQL)
rows = crsr.fetchall()

# --------------------------
# Prepare data lists
# --------------------------
Datevar = []           # will hold date strings (expected in "dd-mm-yyyy" format)
Average_Reading = []   # numeric values

for row in rows:
    Datevar.append(str(row[0]))
    Average_Reading.append(float(row[1]))

print(rows)

# --------------------------
# Convert date strings to datetime objects
# --------------------------
# Adjust the date format if necessary.
dates = []
for d in Datevar:
    try:
        dates.append(datetime.strptime(d, "%d-%m-%Y"))
    except Exception as e:
        print(f"Error parsing date '{d}': {e}")
        # If parsing fails, append None and you might want to filter these out later. 
        dates.append(None)

# --------------------------
# Define positions and properties for bars (3D)
# --------------------------
ind = np.arange(len(Average_Reading))  # x positions, one per reading
width = 0.35   # width of each bar (in x-direction)
depth = 0.5    # depth of each bar (in y-direction)

# Determine conditional colors for each bar:
colors = []
for value in Average_Reading:
    if value > 7:
        colors.append('red')
    elif value < 3:
        colors.append('blue')
    else:
        colors.append('green')

# --------------------------
# Create a 3D bar chart
# --------------------------
fig = plt.figure()
ax = fig.add_subplot(111, projection='3d')

# Draw bars using bar3d:
# - x: positions along date axis (ind)
# - y: dummy positions (0 for each since we have only one row of data)
# - z: the base (0)
# - dx: width, dy: depth, dz: height (Average_Reading)
ax.bar3d(ind, 
         np.zeros_like(ind), 
         np.zeros_like(ind), 
         width, 
         depth, 
         Average_Reading, 
         color=colors, 
         shade=True)

# --------------------------
# Setting custom x-axis ticks: Only one label per month
# --------------------------
tick_positions = []
tick_labels = []
last_month = None

# Loop over positions and parsed dates; because data are assumed sorted,
# record a tick when the month changes.
for i, d in enumerate(dates):
    if d is None:
        continue
    current_month = d.strftime('%Y-%m')
    if current_month != last_month:
        # Place tick at the center of the bar
        tick_positions.append(i + width / 2)
        # Format label as "Mon YYYY", e.g., "Jan 2022"
        tick_labels.append(d.strftime('%b %Y'))
        last_month = current_month

# Use our custom tick positions and labels on the x-axis:
ax.set_xticks(tick_positions)
ax.set_xticklabels(tick_labels, rotation=45, fontsize=10)

# --------------------------
# Labeling and formatting axes
# --------------------------
#ax.set_xlabel('Date')
ax.set_ylabel('')  # not used; can be hidden
ax.set_zlabel('Average Reading')
plt.title('Average Reading per Day (3D)')

# Hide dummy y-axis ticks as we have no meaningful y variable.
ax.set_yticks([])

plt.tight_layout()

# --------------------------
# Save the chart as a PNG file before showing it
# --------------------------
plt.savefig("AverageReading3D.png", dpi=300)

plt.show()
