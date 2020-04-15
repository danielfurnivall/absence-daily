import pandas as pd
import numpy as np
from tkinter.filedialog import askopenfilename


absence_data = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Daily_Absence/RTW_Fortnightly/',
                           filetypes=(("Excel File", "*.xls"), ("All Files", "*.*")),
                           title="Choose the relevant absence extract."
                           )
df = pd.read_excel(absence_data, skiprows=4)
sd = 'W:/Workforce Monthly Reports/Monthly_Reports/Mar-20 Snapshot/Staff Download/2020-03 - Staff Download - GGC.xls'
sd = pd.read_excel(sd)
df.rename(columns={'Pay No':'Pay_Number'}, inplace=True)
df = df.merge(sd, on='Pay_Number', how='left')
print(df.columns)
df['StartDate'] = pd.to_datetime(df['Absence Episode Start Date']).dt.strftime('%d-%m-%Y')
df['Absence Episode Start Date'] = pd.to_datetime(df['StartDate'], dayfirst=True)
print(df['Absence Episode Start Date'].value_counts())

abs_reasons = df['AbsenceReason Description'].unique()
print(abs_reasons)

df = df[df['AbsenceReason Description'].isin(['Coronavirus – Household Related – Self Isolating',
                                              'Coronavirus – Self displaying symptoms – Self Isolating'
                                              ])]

print(df['AbsenceReason Description'].value_counts())

df['Proj_Abs_Period'] = np.where(
    df['AbsenceReason Description'] == 'Coronavirus – Household Related – Self Isolating', 14, 7)

#this is very clever - vectorised method of adding days to date
df['Proj_RTW_Date'] = df['Absence Episode Start Date'] + df['Proj_Abs_Period'].astype('timedelta64[D]')


print(df['Proj_RTW_Date'].value_counts())

df_piv = pd.pivot_table(df, index='Sector/Directorate/HSCP', values='Proj_Abs_Period', aggfunc=np.sum)
print(df_piv)