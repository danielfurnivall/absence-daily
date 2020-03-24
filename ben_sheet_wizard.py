import pandas as pd
from datetime import date
import numpy as np
absence_data = '/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/0 Daily Absence - Absence Data _ 40661676.xls'
shift_checker = '/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/0 Daily Absence - Shift Checker - Everyone _ 40661694.xls'
sd = pd.read_excel('/media/wdrive/Workforce Monthly Reports/Monthly_Reports/Feb-20 Snapshot/Staff Download/2020-02 - Staff Download - GGC.xls')
def concatenate_excel(filename, output):
    df = pd.ExcelFile(filename)
    # df = pd.concat(pd.read_excel('/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/0 Daily Absence - Absence Data _ 40661676.xls', sheet_name=None), ignore_index=True)
    print(df.sheet_names)

    print(df.sheet_names[1:])
    fin_df = pd.read_excel(filename, skiprows=1)
    cols = fin_df.columns

    for i in df.sheet_names[1:]:
        df1 = pd.read_excel(filename, sheet_name=i)
        df1.columns = fin_df.columns
        fin_df = fin_df.append(df1)
        print(len(fin_df))
    print(len(fin_df))
    if output == 'shiftchecker':
        print(fin_df.columns)
        fin_df = fin_df.rename(columns={'Pay Number': 'Pay_Number'})


        fin_df = fin_df.merge(sd[['Pay_Number','Area','Sector/Directorate/HSCP','Job_Family',
                                  'Sub_Job_Family', 'Sub-Directorate 1', 'Sub-Directorate 2',
                                  'department', 'Pay_Band', 'Cost_Centre']], on="Pay_Number", how='left')
        print(fin_df['Shift Start Date  & Time'].value_counts())
        fin_df['Shift Start Date  & Time'] = fin_df['Shift Start Date  & Time'].astype(str).str[:10]
        print(fin_df['Shift Start Date  & Time'].value_counts())
        fin_df['Band Group'] = np.where(fin_df['Pay_Band'].isin(['1','2','3','4']), 'Non Registered', 'Registered')
        print(fin_df['Band Group'].value_counts())
        fin_df['Lookup_String'] = fin_df['Pay_Number'].astype(str)+fin_df['Shift Start Date  & Time'].astype(str)
    if output=='absence':
        print(fin_df.columns)
        fin_df['Absence Episode Start Date'] = fin_df['DerivedAbsence Start Date'].astype(str).str[:10]
        fin_df['Lookup_String'] = fin_df['Pay No'].astype(str)+fin_df['Absence Episode Start Date'].astype(str)
        print(fin_df['Lookup_String'].value_counts())


    fin_df.to_csv('/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/appended-'+output+'-'+str(date.today())+'.csv',
                  index=False)


def merger(absence_file, shiftcheck_file):
    df1 = pd.read_csv(absence_file)
    df2 = pd.read_csv(shiftcheck_file)
    df2 = df2.merge(df1[['Lookup_String','Absence Type','AbsenceReason Description','Hours Lost']], on='Lookup_String',
                    how='left')

    ccbase = pd.read_excel('/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/CCBase.xlsx')
    ccbase = {row[0]: row[1] for row in ccbase.values}
    df2['department'] = df2['Cost_Centre'].map(ccbase)

    df2 = df2.fillna(value={'Absence Type':'<blank>','AbsenceReason Description':'<blank>','Area':'<blank>',
                            'Sector/Directorate/HSCP':'<blank>','Job_Family':'<blank>','Sub_Job_Family':'<blank>',
                            'Sub-Directorate 1':'<blank>','Sub-Directorate 2':'<blank>','department':'<blank>',
                            'Pay_Band':'<blank>','Cost_Centre':'<blank>'})
    print(df2['Pay_Band'].value_counts(dropna=False))

    df2.to_csv('/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/Shift-Checker-Complete'+str(date.today())+'.csv',
               index=False)

def pivot(file):
    df = pd.read_csv(file)
    print(df.columns)
    df_piv = pd.pivot_table(df, index=['Shift Start Date  & Time','Area','department','Job_Family','Sub_Job_Family',
                                       'Band Group','Absence Type','AbsenceReason Description'],
                            values=['Basic Hours (Standard)        ','Excess Part-time Hours','Overtime T1/2',
                                    'Hours Lost'], aggfunc=np.sum)
    df_piv['Bank Hours'] = ''
    df_piv['Agency Hours'] = ''
    print(df_piv.columns)
    df_piv.reset_index(inplace=True)
    df_piv= df_piv.rename(columns={'Shift Start Date  & Time':'Rounded Date','Job_Family':'Job Family',
                                   'Sub_Job_Family':'Sub Family','Absence Type':'Absence_Reason',
                                   'AbsenceReason Description':'Abs_Desc',
                                   'Basic Hours (Standard)        ':'Sum of Basic Hours (Standard)        ',
                                   'Excess Part-time Hours':'Sum of Excess Part-time Hours',
                                   'Overtime T1/2':'Sum of Overtime T1/2','Hours Lost':'Sum of Hrs Lost'})

    df_piv = df_piv[['Rounded Date','Area','department','Job Family','Sub Family','Band Group', 'Absence_Reason',
                     'Abs_Desc','Sum of Basic Hours (Standard)        ','Sum of Excess Part-time Hours',
                     'Sum of Overtime T1/2','Sum of Hrs Lost','Bank Hours','Agency Hours']]
    df_piv.to_excel('/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/pivot'+str(date.today())+'.xlsx', index=False)


# concatenate_excel(absence_data, 'absence')
# concatenate_excel(shift_checker, 'shiftchecker')

abs_file = '/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/appended-'+'absence'+'-'+str(date.today())+'.csv'
shift_file = '/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/appended-'+'shiftchecker'+'-'+str(date.today())+'.csv'
merger(abs_file, shift_file)
file = '/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/Shift-Checker-Complete'+str(date.today())+'.csv'
pivot(file)
