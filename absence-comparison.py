import pandas as pd

df_allabs = pd.read_excel('W:/Daily_Absence/feb23-mar29-abs.xls', skiprows=4)
df_today = pd.read_excel('W:/Daily_Absence/2020-03-29.xls', skiprows=4)
sd = pd.read_excel('W:/Workforce Monthly Reports/Monthly_Reports/Feb-20 Snapshot/Staff Download/2020-02 - Staff Download - GGC.xls')
phones = pd.read_excel('W:/MFT/phone number lookup.xlsx')
manager = pd.read_excel('W:/Daily_Absence/manager_lookup.xlsx')
manager = manager[['Pay_Number', 'Supervisor email address']]
df_allabs = df_allabs.rename(columns={'Pay No': 'Pay_Number'})
df_today = df_today.rename(columns={'Pay No': 'Pay_Number'})
df_allabs = df_allabs.merge(sd, on='Pay_Number', how='left')
df_allabs = df_allabs.merge(phones, on="Pay_Number", how='left')
df_allabs = df_allabs.merge(manager, on="Pay_Number", how='left')


print(df_allabs.columns)

df_allabs = df_allabs[(df_allabs['AbsenceReason Description'] == 'Infectious diseases') |
                      (df_allabs['AbsenceReason Description'] == 'Coronavirus – Covid 19 Positive')]

print(df_allabs['AbsenceReason Description'].value_counts())

df_today = df_today[(df_today['AbsenceReason Description'] == 'Infectious diseases') |
                      (df_today['AbsenceReason Description'] == 'Coronavirus – Covid 19 Positive')]

print(df_today['AbsenceReason Description'].value_counts())
absent_today = df_today['Pay_Number'].unique()

not_absent = df_allabs[~df_allabs['Pay_Number'].isin(absent_today)]
absent_today = df_allabs[df_allabs['Pay_Number'].isin(absent_today)]
print(len(absent_today))
print(len(df_allabs))
print(len(not_absent))


writer = pd.ExcelWriter('w:/Daily_Absence/all_positives.xlsx', engine='xlsxwriter')
not_absent[['Pay_Number','Supervisor email address','Forename','Surname','AbsenceReason Description',
                              'Sector/Directorate/HSCP','Sub-Directorate 1','department','Job_Family',
                              'Absence Episode Start Date','Address_Line_1','Address_Line_2','Address_Line_3',
                              'Postcode', 'Best Phone']].to_excel(writer, sheet_name='Not currently absent', index=False)
absent_today[['Pay_Number','Supervisor email address','Forename','Surname','AbsenceReason Description',
                              'Sector/Directorate/HSCP','Sub-Directorate 1','department','Job_Family',
                              'Absence Episode Start Date','Address_Line_1','Address_Line_2','Address_Line_3',
                              'Postcode', 'Best Phone']].to_excel(writer, sheet_name='Currently absent', index=False)
writer.save()


