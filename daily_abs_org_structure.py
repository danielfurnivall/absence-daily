import pandas as pd
from datetime import date


newpath = 'W:/Daily_Absence/' + (date.today()).strftime("%Y-%m-%d") + '.xls'
df = pd.read_excel(newpath, skiprows=4)
sd = pd.read_excel(
    'W:/Workforce Monthly Reports/Monthly_Reports/Sep-20 Snapshot/Staff Download/2020-09 - Staff Download - GGC.xls')
deptWTE_lookup = {}
for i in sd['department'].unique().tolist():
    dfx = sd[sd['department'] == i]
    wte_count = dfx['WTE'].sum(axis=0).round(2)
    deptWTE_lookup[i] = wte_count
print(deptWTE_lookup)
sd['DeptWTE'] = sd['department'].map(deptWTE_lookup)

df = df.rename(columns={'Pay No': 'Pay_Number'})
df = df.merge(sd, on='Pay_Number', how='left')
# df.drop(df.columns[0, 1], axis=1, inplace=True)
df.to_excel('W:/Daily_Absence/all_abs_hierarchy'+(date.today().strftime('%Y-%m-%d'))+'.xlsx', index=False)
