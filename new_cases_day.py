import pandas as pd
import datetime as dt
import warnings
warnings.filterwarnings('ignore')
sd = pd.read_excel('W:/Workforce Monthly Reports/Monthly_Reports/Feb-20 Snapshot/Staff Download/2020-02 - Staff Download - GGC.xls')
sd.rename(columns={'Pay_Number':'Pay No'}, inplace=True)
startdate = pd.to_datetime('29/03/20', dayfirst=True)
enddate = pd.to_datetime('05/04/20', dayfirst=True)

writer = pd.ExcelWriter('W:/Daily_Absence/New_Cases.xlsx', engine='xlsxwriter')
dd = [startdate + dt.timedelta(days=x) for x in range((enddate-startdate).days + 1)]
for i in dd:
    date = (i.strftime('%Y-%m-%d'))
    df = pd.read_excel('W:/Daily_Absence/'+date+'.xls', skiprows=4)
    df = df.merge(sd[['Pay No', 'Sector/Directorate/HSCP']], on='Pay No')

    df['date'] = pd.to_datetime(df['Absence Episode Start Date']).dt.strftime('%Y-%m-%d')
    df = df[df['date'] == date]

    df_self = df[df['AbsenceReason Description'] == 'Coronavirus – Self displaying symptoms – Self Isolating']
    df_household = df[df['AbsenceReason Description'] == 'Coronavirus – Household Related – Self Isolating']
    print(date, "\nNew Self Isolating - Self displaying symptoms: "+str(len(df_self)), "\nNew Self Isolating - Household Related: "+
          str(len(df_household)))
    self_piv = pd.pivot_table(df_self, index='Sector/Directorate/HSCP', values='Pay No', aggfunc='count')
    print(self_piv)
    household_piv = pd.pivot_table(df_household, index='Sector/Directorate/HSCP', values='Pay No', aggfunc='count')
    self_piv.to_excel(writer, sheet_name=date+'-self')
    household_piv.to_excel(writer, sheet_name=date+'-household')
writer.save()