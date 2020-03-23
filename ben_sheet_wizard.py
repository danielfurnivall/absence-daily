import pandas as pd
absence_data = '/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/0 Daily Absence - Absence Data _ 40661676.xls'
shift_checker = '/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/0 Daily Absence - Shift Checker - Everyone _ 40661694.xls'
def concatenate_excel(filename, output):
    df = pd.ExcelFile(filename)
    # df = pd.concat(pd.read_excel('/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/0 Daily Absence - Absence Data _ 40661676.xls', sheet_name=None), ignore_index=True)
    print(df.sheet_names)

    print(df.sheet_names[1:])
    fin_df = pd.read_excel(filename, skiprows=1)
    cols = fin_df.columns
    #
    for i in df.sheet_names[1:]:
        df1 = pd.read_excel(filename, sheet_name=i)
        df1.columns = fin_df.columns
        fin_df = fin_df.append(df1)
        print(len(fin_df))
    print(len(fin_df))
    fin_df.to_csv('/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/appended-'+output+'.csv', index=False)

concatenate_excel(absence_data, 'absence')
concatenate_excel(shift_checker, 'shiftchecker')