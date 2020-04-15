import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import date
from tkinter.filedialog import askopenfilename

absence_data = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Daily_Absence/Weekly/',
                           filetypes=(("Excel File", "*.xls"), ("All Files", "*.*")),
                           title="Choose the relevant absence extract."
                           )

df = pd.read_excel(absence_data, skiprows=4)
sd = pd.read_excel('W:/Workforce Monthly Reports/Monthly_Reports/Feb-20 Snapshot/Staff Download/2020-02 - Staff Download - GGC.xls')

df = df.rename(columns={'Pay No':'Pay_Number'})
print(len(df))
df = df.merge(sd, on='Pay_Number', how='left')
print(len(df))
print(df.columns)

df = df[df['AbsenceReason Description'].isin(['Coronavirus',
                                           'Infectious diseases',
                                           'Coronavirus – Covid 19 Positive',
                                           'Coronavirus – Underlying Health Condition',
                                           'Coronavirus – Household Related – Self Isolating',
                                           'Coronavirus – Self displaying symptoms – Self Isolating'])]

west_dun_piv = pd.pivot_table(df[df['Sector/Directorate/HSCP'] == 'West Dunbartonshire HSCP'],
                              index=['Sub-Directorate 1', 'Sub-Directorate 2', 'department','Job_Family',
                                     'Sub_Job_Family', 'AbsenceReason Description'],
                              values=['Working Days Lost', 'Pay_Number'],
                              aggfunc={'Working Days Lost':np.sum, 'Pay_Number':'count'})
west_dun_piv.to_excel(absence_data[:-4]+"-pivot.xlsx")