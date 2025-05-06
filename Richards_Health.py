from datetime import datetime
import numpy as np
import pyodbc
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

# --------------------------
# Execute the query (stored procedure)
# --------------------------
SQL = """\
{CALL AverageControl1}
"""
crsr.execute(SQL)
rows = crsr.fetchall()

# --------------------------
# Prepare and sort data lists
# --------------------------
Datevar = []
Average_Reading = []

for row in rows:
    Datevar.append(str(row[0]))  # Store date as string first
    Average_Reading.append(float(row[1]))  # Store numeric reading

print(rows)

# Convert date strings to datetime objects for sorting
parsed_dates = []
for d in Datevar:
    try:
        parsed_dates.append(datetime.strptime(d, "%d-%m-%Y"))
    except Exception as e:
        print(f"Error parsing date '{d}': {e}")
        parsed_dates.append(None)

# Zip and sort data by date
sorted_data = sorted(zip(parsed_dates, Average_Reading), key=lambda x: x[0] if x[0] else datetime.max)

# Unzip sorted data back into lists
dates_sorted, Average_Reading_sorted = zip(*sorted_data)

# --------------------------
# Define positions and properties for bars (3D)
# --------------------------
ind = np.arange(len(Average_Reading_sorted))  # x positions, one per reading
width = 0.35   # width of each bar
depth = 0.5    # depth of each bar

# Determine conditional colors for each bar
colors = ['red' if val > 10 else 'blue' if val < 3 else 'green' for val in Average_Reading_sorted]

# --------------------------
# Create a 3D bar chart
# --------------------------
fig = plt.figure()
ax = fig.add_subplot(111, projection='3d')

ax.bar3d(ind, 
         np.zeros_like(ind), 
         np.zeros_like(ind), 
         width, 
         depth, 
         Average_Reading_sorted, 
         color=colors, 
         shade=True)

# --------------------------
# Setting custom x-axis ticks: One label per month in order
# --------------------------
tick_positions = []
tick_labels = []
last_month = None

for i, d in enumerate(dates_sorted):
    if d is None:
        continue
    current_month = d.strftime('%Y-%m')
    if current_month != last_month:
        tick_positions.append(i + width / 2)
        tick_labels.append(d.strftime('%b %Y'))
        last_month = current_month

ax.set_xticks(tick_positions)
ax.set_xticklabels(tick_labels, rotation=45, fontsize=10)

# --------------------------
# Labeling and formatting axes
# --------------------------
ax.set_zlabel('Average Reading')
plt.title('Average Reading per Day (3D)')

# Hide dummy y-axis ticks as we have no meaningful y variable.
ax.set_yticks([])

plt.tight_layout()

# --------------------------
# Save the chart as a PNG file before showing it
# --------------------------
plt.savefig(r"C:\Users\rchrd\Documents\Richard\Diabetes\AverageReading3D.png", dpi=300)
plt.show()