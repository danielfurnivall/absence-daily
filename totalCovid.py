'''Quick file to report on multi-day covid absence pulled from SSTS'''
import pandas as pd

df = pd.read_excel('W:/Daily_Absence/nareen-allCoviddata.xlsx', skiprows=4)
sd = pd.read_excel('W:/Workforce Monthly Reports/Monthly_Reports/Mar-20 Snapshot/Staff Download/2020-03 - Staff Download - GGC.xls')

df = df.rename(columns={'Pay No': 'Pay_Number'})
df = df.merge(sd, on='Pay_Number', how='left')
df = df[['Pay_Number', 'Roster Location',
       'Absence Type', 'AbsenceReason Description',
       'Absence Episode Start Date', 'Absence Episode End Date',
       'Absence Episode Days', 'Working Days Lost', 'Hours Lost', 'Area',
       'Sector/Directorate/HSCP_Code', 'Sector/Directorate/HSCP',
       'Sub-Directorate 1', 'Sub-Directorate 2', 'department', 'Cost_Centre',
       'Surname', 'Forename', 'Base', 'Job_Family_Code', 'Job_Family',
       'Sub_Job_Family', 'Post_Descriptor', 'Conditioned_Hours',
       'Contracted_Hours', 'WTE', 'Contract_Description', 'NI_Number', 'Age',
       'Date_of_Birth', 'Date_Started', 'Contract Planned Contract End Date',
       'Annual_Salary', 'Date_To_Grade', 'Date_Superannuation_Started',
       'SB_Number', 'Sick_Date_Entitlement_From', 'Description',
       'Marital_Status', 'Sex', 'Job_Description', 'Grade', 'Group_Code',
       'Pay_Scale', 'Pay_Band', 'Scale_Point', 'Pay_Point', 'Incremental Date',
       'Address_Line_1', 'Address_Line_2', 'Address_Line_3', 'Postcode',
       'Area_Pay_Division', 'Mental_Health_Y/N']]
print(df.columns)
all_isolating = df[df['AbsenceReason Description'] == 'Coronavirus – Self displaying symptoms – Self Isolating']
all_parental = df[df['AbsenceReason Description'] == 'Coronavirus']
all_positive = df[(df['AbsenceReason Description'] == 'Infectious diseases') | (df['AbsenceReason Description'] == 'Coronavirus – Covid 19 Positive')]
all_household_isolating = df[df['AbsenceReason Description'] == 'Coronavirus – Household Related – Self Isolating']
all_underlying = df[df['AbsenceReason Description'] == 'Coronavirus – Underlying Health Condition']



print("Self isolating - "+str(len(all_isolating)))
print("Underlying Conditions - "+str(len(all_underlying)))
print("Covid Parental Leave - "+str(len(all_parental)))
print("Household isolating - "+str(len(all_household_isolating)))
print("Covid Positive - "+str(len(all_positive)))

all_isolating.to_excel('W:/Daily_Absence/Nareen-self-isolating.xlsx', index=False)
all_household_isolating.to_excel('W:/Daily_Absence/Nareen-household-isolating.xlsx', index=False)