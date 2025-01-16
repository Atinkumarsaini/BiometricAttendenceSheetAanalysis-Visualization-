import pandas as pd 
import numpy as np
from datetime import datetime
df = pd.read_excel("attendance.xlsx")
df["Out Duration( In Hrs)"] = df["Out Duration( In Hrs)"].str.replace('-1','0')
df.fillna("null", inplace = True)
df['attendance'] = np.where(df['Punch Records'] == 'null', '', 'present')

def clean_and_modify(row):
    if row != 'null':
        a = row.replace('(', '').replace(')', '').split(',')
        if len(a) % 2 == 0:
            return row + ',18:30:out(appic)'
    return row
df["Punch Records"] = df["Punch Records"].apply(clean_and_modify)
c = []
names = []
for index,row in df.iterrows():
    if row["attendance"] == 'present':
        c.append(row["Attendance Date"])
        names.append(row["Name"])
df['check'] = df["Punch Records"].apply(lambda x: len(x.replace('(', '').replace(')', '').split(','))-1 if len(x.replace('(', '').replace(')', '').split(','))-3>0 else 0)
df.to_excel("final.xlsx", index=False)

name = []
for i in names:
    if i not in name:
        name.append(i)
name_attendance = []
for i in name:
    name_attendance.append(names.count(i))
date =[]
for i in c:
    if i not in date:
        date.append(i)
attendance = []
for i in date:
    attendance.append(c.count(i))
for i in date:
    day_numbers = [datetime.strptime(date_str, '%Y-%m-%d').strftime('%d') for date_str in date]
date = day_numbers

import matplotlib.pyplot as plt
fig = plt.figure(figsize=(8, 6)) 
ax = fig.add_subplot(111)
for i, j in zip(date, attendance):
    ax.annotate('%s' % j, xy=(i, j), xytext=(5, 0), textcoords='offset points')
plt.plot(date, attendance, marker='s', linestyle='-', color='b', label='candidate') 
plt.xlabel('Date')
plt.ylabel('attendance')
plt.title('Attendance vs Date')
plt.grid(True)
plt.legend()
plt.show()


fig, ax = plt.subplots(figsize=(50, 12))
ax.bar(name, 
       name_attendance)
bars = ax.bar(name, name_attendance)
for bar, attendance in zip(bars, name_attendance):
    yval = bar.get_height()
    ax.annotate(attendance,
                xy=(bar.get_x() + bar.get_width() / 2, yval),
                xytext=(0, 0),  
                textcoords="offset points",
                ha='center', va='bottom')
ax.set_xticks([])

# Add custom x-axis labels inside the bars
for bar, label in zip(bars, name):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() / 2,
            str(label), ha='center', va='center', rotation=90, color='Black', fontsize=12)
plt.title("Employee's attendance")
plt.show()



