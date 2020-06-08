'''
This file takes two SSTS extracts (one absence extract and one shift checker extract). These files are parsed,
sheet by sheet and concatenated. It then adds relevant data from org. structure etc and creates a new master shift
checker file merged with the absence.
A pivot table is then produced for upload to microstrategy.
'''
import pandas as pd
from datetime import date
import numpy as np
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import os

# Read in files
absence_data = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Coronavirus Daily Absence/MICROSTRATEGY/',
                               filetypes=(("Excel File", "*.xls"), ("All Files", "*.*")),
                               title="Choose the relevant absence extract."
                               )
shift_checker = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Coronavirus Daily Absence/MICROSTRATEGY/',
                                filetypes=(("Excel File", "*.xls"), ("All Files", "*.*")),
                                title="Choose the relevant shift checker extract."
                                )

staff_download = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Staff Downloads/',
                                 filetypes=(("Excel File", "*.xlsx"), ("All Files", "*.*")),
                                 title="Choose the relevant Staff Download file."
                                 )
sd = pd.read_excel(staff_download)


def concatenate_excel(filename, output):
    '''
    This function takes in either absence or shift checker data then opens the first sheet and concats the rest in.
    '''

    # open the excelfile to get list of sheets within.
    df = pd.ExcelFile(filename)
    print(df.sheet_names)

    # open first sheet as master df
    master = pd.read_excel(filename, skiprows=1)
    cols = master.columns

    for i in df.sheet_names[1:]:
        # iterate across sheets, reading then appending each to master df
        df1 = pd.read_excel(filename, sheet_name=i)
        df1.columns = master.columns
        master = master.append(df1)

    print(len(master))
    # part 1 - deals with shiftchecker file
    if output == 'shiftchecker':
        # rename for merge
        master = master.rename(columns={'Pay Number': 'Pay_Number'})
        # add org structure data from staff download
        master = master.merge(sd[['Pay_Number', 'Area', 'Sector/Directorate/HSCP', 'Job_Family',
                                  'Sub_Job_Family', 'Sub-Directorate 1', 'Sub-Directorate 2',
                                  'department', 'Pay_Band', 'Cost_Centre']], on="Pay_Number", how='left')
        # cut useless datetime fragments
        master['Shift Start Date  & Time'] = master['Shift Start Date  & Time'].astype(str).str[:10]
        # mark registered/nonregistered employees
        master['Band Group'] = np.where(master['Pay_Band'].isin(['1', '2', '3', '4']), 'Non Registered',
                                        np.where(master['Pay_Band'] == '', '',
                                                 'Registered'))  # Added np.where - remove second np.where
        # fix for public holiday overtime - implemented w/c 04/05/20
        if 'Overtime T2' in master:
            master['Overtime T1/2'] = master['Overtime T1/2'] + master['Overtime T2']
        # create string to merge shifts across abs/shiftchecker
        master['Lookup_String'] = master['Pay_Number'].astype(str) + master['Shift Start Date  & Time'].astype(str)

    # part 2 - deals with absence file
    if output == 'absence':
        # change abs descs to reflect that "infectious diseases" is assumed as covid positive
        master['AbsenceReason Description'].replace({'Infectious diseases': 'Coronavirus – Covid 19 Positive'},
                                                    inplace=True)
        # derived is more accurate so replace
        master['Absence Episode Start Date'] = master['DerivedAbsence Start Date'].astype(str).str[:10]

        # build lookup string as before
        master['Lookup_String'] = master['Pay No'].astype(str) + master['Absence Episode Start Date'].astype(str)
        # remove all non-absences
        master = master[master['Hours Lost'] != 0]
        master.sort_values(by='Hours Lost', ascending=False)
    # drop duplicates
    master.drop_duplicates(inplace=True, keep='first')
    # publish temporary file
    master.to_csv('W:/Coronavirus Daily Absence/MICROSTRATEGY/appended-' + output + '-' + str(date.today()) + '.csv',
                  index=False)


def merger(absence_file, shiftcheck_file):
    """This function takes in an absence and shiftcheck file and merges them. We also use this part of the process to
       add base using cost-centre details """
    absence_data = pd.read_csv(absence_file)
    shiftchecker_data = pd.read_csv(shiftcheck_file)
    shiftchecker_data = shiftchecker_data.merge(
        absence_data[['Lookup_String', 'Absence Type', 'AbsenceReason Description', 'Hours Lost']], on='Lookup_String',
        how='left')
    # build dictionary of cost centres and bases
    ccbase = pd.read_excel('W:/Coronavirus Daily Absence/MICROSTRATEGY/CCBase.xlsx')
    ccbase = {row[0]: row[1] for row in ccbase.values}
    shiftchecker_data['department'] = shiftchecker_data['Cost_Centre'].map(ccbase)

    # replace zeros with blanks
    shiftchecker_data = shiftchecker_data.fillna(
        value={'Absence Type': '<blank>', 'AbsenceReason Description': '<blank>', 'Area': '<blank>',
               'Sector/Directorate/HSCP': '<blank>', 'Job_Family': '<blank>', 'Sub_Job_Family': '<blank>',
               'Sub-Directorate 1': '<blank>', 'Sub-Directorate 2': '<blank>', 'department': '<blank>',
               'Pay_Band': '<blank>', 'Cost_Centre': '<blank>'})

    # publish merged file file
    shiftchecker_data.to_csv(
        'W:/Coronavirus Daily Absence/MICROSTRATEGY/Shift-Checker-Complete - ' + str(date.today()) + '.csv',
        index=False)


def pivot(file):
    """This function takes in the completed shift checker complete file, and makes a nice pivot table"""
    df = pd.read_csv(file)
    df_piv = pd.pivot_table(df,
                            index=['Shift Start Date  & Time', 'Sector/Directorate/HSCP', 'department', 'Job_Family',
                                   'Sub_Job_Family',  # added sec/dir/hscp and removed area
                                   'Band Group', 'Absence Type', 'AbsenceReason Description'],
                            values=['Basic Hours (Standard)        ', 'Excess Part-time Hours', 'Overtime T1/2',
                                    'Hours Lost'], aggfunc=np.sum)

    '''these two columns need to be added for microstrategy. they were initially intended to be linked with actual bank
    and agency info but that fell by the wayside.'''
    df_piv['Bank Hours'] = ''
    df_piv['Agency Hours'] = ''

    # prettify output
    df_piv.reset_index(inplace=True)
    df_piv = df_piv.rename(columns={'Shift Start Date  & Time': 'Rounded Date', 'Job_Family': 'Job Family',
                                    'Sub_Job_Family': 'Sub Family', 'Absence Type': 'Absence_Reason',
                                    'AbsenceReason Description': 'Abs_Desc',
                                    'Basic Hours (Standard)        ': 'Sum of Basic Hours (Standard)        ',
                                    'Excess Part-time Hours': 'Sum of Excess Part-time Hours',
                                    'Overtime T1/2': 'Sum of Overtime T1/2', 'Hours Lost': 'Sum of Hrs Lost'})

    df_piv = df_piv[['Rounded Date', 'Sector/Directorate/HSCP', 'department', 'Job Family', 'Sub Family', 'Band Group',
                     'Absence_Reason',  # added sec/dir/hscp and removed area
                     'Abs_Desc', 'Sum of Basic Hours (Standard)        ', 'Sum of Excess Part-time Hours',
                     'Sum of Overtime T1/2', 'Sum of Hrs Lost', 'Bank Hours', 'Agency Hours']]

    # replace blank depts with other ggc sites
    df_piv['department'].replace({'<blank>': 'Other GGC Sites'}, inplace=True)
    df_piv.replace({'<blank>': ''}, inplace=True)
    df_piv = df_piv.rename(columns={'Sector/Directorate/HSCP': 'Area'})  # rename area

    # publish final
    df_piv.to_excel('W:/Coronavirus Daily Absence/MICROSTRATEGY/pivot - ' + str(date.today()) + '.xlsx', index=False)


def pivot2(file):
    '''Does exactly the same as pivot1 except outputs as newpivot and adds some extra org structure bits'''
    df = pd.read_csv(file)
    print(df.columns)
    df_piv = pd.pivot_table(df, index=['Shift Start Date  & Time', 'Sector/Directorate/HSCP', 'department',
                                       'Sub-Directorate 1', 'Sub-Directorate 2', 'Cost_Centre',
                                       'Job_Family', 'Sub_Job_Family',  # added sec/dir/hscp and removed area
                                       'Band Group', 'Absence Type', 'AbsenceReason Description'],
                            values=['Basic Hours (Standard)        ', 'Excess Part-time Hours', 'Overtime T1/2',
                                    'Hours Lost'], aggfunc=np.sum)
    df_piv['Bank Hours'] = ''
    df_piv['Agency Hours'] = ''
    print(df_piv.columns)
    df_piv.reset_index(inplace=True)
    df_piv = df_piv.rename(columns={'Shift Start Date  & Time': 'Rounded Date', 'Job_Family': 'Job Family',
                                    'Sub_Job_Family': 'Sub Family', 'Absence Type': 'Absence_Reason',
                                    'AbsenceReason Description': 'Abs_Desc',
                                    'Basic Hours (Standard)        ': 'Sum of Basic Hours (Standard)        ',
                                    'Excess Part-time Hours': 'Sum of Excess Part-time Hours',
                                    'Overtime T1/2': 'Sum of Overtime T1/2', 'Hours Lost': 'Sum of Hrs Lost'})

    df_piv = df_piv[['Rounded Date', 'Sector/Directorate/HSCP', 'department', 'Sub-Directorate 1', 'Sub-Directorate 2',
                     'Cost_Centre', 'Job Family', 'Sub Family', 'Band Group', 'Absence_Reason',
                     # added sec/dir/hscp and removed area
                     'Abs_Desc', 'Sum of Basic Hours (Standard)        ', 'Sum of Excess Part-time Hours',
                     'Sum of Overtime T1/2', 'Sum of Hrs Lost', 'Bank Hours', 'Agency Hours']]
    df_piv['department'].replace({'<blank>': 'Other GGC Sites'}, inplace=True)
    df_piv.replace({'<blank>': ''}, inplace=True)
    df_piv = df_piv.rename(columns={'Sector/Directorate/HSCP': 'Area'})  # rename area
    df_piv.to_excel('W:/Coronavirus Daily Absence/MICROSTRATEGY/newpivot - ' + str(date.today()) + '.xlsx', index=False)


def test(filename):
    """This is a test function that changes depending on what needs to be tested. You can safely ignore it."""
    df = pd.read_csv(filename)
    print(df.columns)
    print(df['Hours Lost'].value_counts(dropna=False))
    print(len(df))
    df = df[df['Hours Lost'] != 0]
    print(len(df))
    df.drop_duplicates(keep='first', inplace=True)
    print(len(df))


# input absence + shiftchecker files
concatenate_excel(absence_data, 'absence')
concatenate_excel(shift_checker, 'shiftchecker')

# open and merge files
abs_file = 'W:/Coronavirus Daily Absence/MICROSTRATEGY/appended-' + 'absence' + '-' + str(date.today()) + '.csv'
shift_file = 'W:/Coronavirus Daily Absence/MICROSTRATEGY/appended-' + 'shiftchecker' + '-' + str(date.today()) + '.csv'
merger(abs_file, shift_file)

# pivot files
file = 'W:/Coronavirus Daily Absence/MICROSTRATEGY/Shift-Checker-Complete - ' + str(date.today()) + '.csv'
pivot(file)
messagebox.showwarning("File complete", "First File is complete - let the relevant person know.")
pivot2(file)

# after pivots are run, delete all the bits
os.remove(abs_file)
os.remove(shift_file)
os.remove(file)
messagebox.showwarning("New File complete", "File is complete - let the relevant person know.")

