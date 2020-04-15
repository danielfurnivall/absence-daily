import pandas as pd
import os
import numpy as np
import datetime

dir = os.listdir('W:/Coronavirus Daily Absence/MICROSTRATEGY/Bank Files')
df = pd.read_excel('W:/Coronavirus Daily Absence/MICROSTRATEGY/Bank Files/'+dir[0])
print(len(df))

print(df.columns)

for i in dir[1:]:
    df_working = pd.read_excel('W:/Coronavirus Daily Absence/MICROSTRATEGY/Bank Files/'+i)
    df = df.append(df_working)
    print(len(df))
df.to_csv('W:/Coronavirus Daily Absence/MICROSTRATEGY/allbank'+datetime.datetime.now().strftime('%Y-%m-%d')+'.csv', index=False)



x = pd.pivot_table(df, index=['Valid Date', 'Staff Group'], columns='Bank / Agency', values='Work Time', aggfunc=np.sum)
print(x)
x.to_csv('W:/Coronavirus Daily Absence/MICROSTRATEGY/bankpiv-'+datetime.datetime.now().strftime('%Y-%m-%d')+'.csv')
# df = pd.read_excel('W:/Coronavirus Daily Absence/MICROSTRATEGY/Bank example 2nd March.xlsx')
# #sd = pd.read_excel('W:/Workforce Monthly Reports/Monthly_Reports/Feb-20 Snapshot/Staff Download/2020-02 - Staff Download - GGC.xls')
# print(df.columns)
# print(df[df['Resource Requirement'].str.contains('Band 2|Band 3|Band 4')])
# print(df['Resource Requirement'].value_counts())