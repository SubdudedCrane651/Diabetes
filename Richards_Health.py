from unicodedata import decimal
import numpy as np
from datetime import datetime
import pyodbc
import pandas as pd
import matplotlib.pyplot as plt

fig = plt.figure()
ax = fig.add_subplot(111)

connection_string = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\rchrd\Documents\Richard\Richards_Health.mdb;'
)
cnxn = pyodbc.connect(connection_string, autocommit=True)
crsr = cnxn.cursor()

for i in crsr.tables(tableType='TABLE'):
    print(i.table_name)

#SQL="""\
#SELECT Format([datevar],"dd-mm-yyyy") AS Expr1, Round(Avg(diabetes.Reading),1) AS AveragePerMonth FROM diabetes WHERE (((diabetes.Datevar)>=#1/1/2022# And (diabetes.Datevar)<=#12/31/2022#)) GROUP BY Format([datevar],"dd-mm-yyyy") ORDER BY Format([datevar],"dd-mm-yyyy");
#"""

SQL="""\
{CALL AverageControl1}
"""

crsr.execute(SQL)

rows = crsr.fetchall()

Datevar = []
Average_Reading = []

for row in rows:
   Datevar.append(str(row[0]))
   Average_Reading.append(float(row[1]))

print(rows)

## necessary variables
ind = np.arange(len(Average_Reading))                # the x locations for the groups
width = 0.35                      # the width of the bars


## the bars
rects1 = ax.bar(ind, Average_Reading, width,
                color='black',
                error_kw=dict(elinewidth=6.5,ecolor='red'))

# axes and labels
ax.set_xlim(-width,len(ind)+width)
ax.set_ylim(0,20)


ax.set_ylabel('Avearge Reading')
ax.set_xlabel('Date')
ax.set_title('Average Reading per day')

ax.set_xticks(ind+width)
Dates = ax.set_xticklabels(Datevar)
plt.setp(Dates, rotation=45, fontsize=10)


plt.show()