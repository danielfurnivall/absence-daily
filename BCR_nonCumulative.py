"""This script takes a folder full of SSTS extracts, then builds a new dataframe with a row for each,
capturing each covid absence codes as a column. Once this file is made, it builds a bar chart racer file"""

import pandas as pd
import os
import bar_chart_race as bcr

# list of files to iterate through
directory = os.listdir('W:/daily_absence/racer_files')

# initialise empty dataframe
master = pd.DataFrame()

# list of covid codes
covid_codes = ['Coronavirus', 'Coronavirus – Household Related – Self Isolating',
               'Coronavirus – Self displaying symptoms – Self Isolating',
               'Coronavirus – Underlying Health Condition',
               'Coronavirus – Covid 19 Positive', 'Coronavirus – Test and Protect Isolation', 'Coronavirus – Quarantine']


def grab_data(df):
    """Shaves dataframe down to only those with covid codes, then returns a dictionary with a counter for each."""
    selection = df[df['AbsenceReason Description'].isin(covid_codes)]
    # initialise empty dataframe
    our_dic = {}
    # iterate through covid codes and add count to dictionary
    for i in covid_codes:
        current_code = len(selection[selection['AbsenceReason Description'] == i])
        our_dic[i] = current_code
    return our_dic


for i in directory:
    # open file
    df = pd.read_excel('W:/daily_absence/racer_files/' + i, skiprows=4)

    # get dictionary for file
    curr = grab_data(df)

    # get date from the first half of filename
    date = i.split('.')[0]
    curr['Date'] = date

    # append to master file
    master = master.append(curr, ignore_index=True)

# reorder columns
master = master[['Date', 'Coronavirus', 'Coronavirus – Covid 19 Positive',
                 'Coronavirus – Household Related – Self Isolating',
                 'Coronavirus – Self displaying symptoms – Self Isolating',
                 'Coronavirus – Underlying Health Condition']]

# rename columns
master = master.rename(columns={'Coronavirus – Covid 19 Positive': 'Covid Positive',
                                'Coronavirus – Household Related – Self Isolating': 'Household-Isolating',
                                'Coronavirus – Self displaying symptoms – Self Isolating': 'Self Isolating',
                                'Coronavirus – Underlying Health Condition': 'Underlying Health Condition'})

# set date as index
master.set_index('Date', inplace=True)

# build bar chart race
z = bcr.bar_chart_race(master, filename='C:/tong/data.mp4', orientation='h', sort='asc',
                       title="NHS GGC Coronavirus Absence Coding")
