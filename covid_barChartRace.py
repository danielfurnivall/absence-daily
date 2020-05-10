'''This builds a nice bar graph racer for our covid daily absence data'''
import pandas as pd
import os
import bar_chart_race as bcr
from IPython.display import HTML

files = os.listdir('W:/daily_absence/racer_files')

master_data = pd.DataFrame()

self_iso = pd.DataFrame()
house_iso = pd.DataFrame()
cov_pos = pd.DataFrame()
coronavirus = pd.DataFrame()
underl = pd.DataFrame()


def grab_data_types(x):
    df = pd.read_excel(x, skiprows=4)

    # get data for self isolators
    self_isolators = df[df['AbsenceReason Description'] ==
                        'Coronavirus – Self displaying symptoms – Self Isolating'][['Pay No']]

    # get data for family symptoms
    household = df[df['AbsenceReason Description'] == 'Coronavirus – Household Related – Self Isolating'][['Pay No']]

    # get data for parental leave
    parental = df[df['AbsenceReason Description'] == 'Coronavirus'][['Pay No']]

    # get data for covid positives
    positive = df[df['AbsenceReason Description'] == 'Coronavirus – Covid 19 Positive'][['Pay No']]

    # get data for underlying health conditions
    underlying = df[df['AbsenceReason Description'] == 'Coronavirus – Underlying Health Condition'][['Pay No']]

    return {'Self displaying symptoms – Self Isolating': self_isolators,
            'Household Related – Self Isolating': household,
            'Coronavirus': parental, 'Covid-19 Positive': positive,
            'Underlying Health Condition': underlying}


print(files)

# code to test self isolating properties
x = grab_data_types('W:/daily_absence/racer_files/2020-04-01.xls')
x['Self displaying symptoms – Self Isolating'].to_csv('C:/APIT/selfisos.csv', index=False)
print(x['Household Related – Self Isolating'])



for this_file in files:
    print(this_file.split(".")[0])
    row_dic = grab_data_types('W:/daily_absence/racer_files/' + this_file)

    # populate date field
    row_dic['Date'] = this_file.split(".")[0]

    # self-isos
    this_row_self_isolators = row_dic['Self displaying symptoms – Self Isolating']
    new_self_isos = this_row_self_isolators#[~this_row_self_isolators.isin(self_iso)]
    print("New Self Isolators = " + str(len(new_self_isos)))
    self_iso = self_iso.append(new_self_isos, ignore_index=True)
    self_iso = self_iso.drop_duplicates()
    row_dic['Self displaying symptoms – Self Isolating'] = len(self_iso)

    #household
    this_row_household = row_dic['Household Related – Self Isolating']
    house_iso = house_iso.append(this_row_household, ignore_index=True)
    house_iso = house_iso.drop_duplicates()
    row_dic['Household Related – Self Isolating'] = len(house_iso)

    #"coronavirus"
    this_row_coronavirus = row_dic['Coronavirus']
    coronavirus = coronavirus.append(this_row_coronavirus)
    coronavirus = coronavirus.drop_duplicates()
    row_dic['Coronavirus'] = len(coronavirus)

    #underlying health issues
    this_row_underlying = row_dic['Underlying Health Condition']
    underl = underl.append(this_row_underlying)
    underl.drop_duplicates(inplace=True)
    row_dic['Underlying Health Condition'] = len(underl)

    #covid pos
    this_row_positive = row_dic['Covid-19 Positive']
    cov_pos = cov_pos.append(this_row_positive)
    cov_pos.drop_duplicates(inplace=True)
    row_dic['Covid-19 Positive'] = len(cov_pos)

    master_data = master_data.append(row_dic, ignore_index=True)
print(master_data)

self_iso.to_csv('C:/APIT/selfisos.csv')

master_data = master_data[['Date', 'Self displaying symptoms – Self Isolating', 'Household Related – Self Isolating',
                           'Coronavirus', 'Covid-19 Positive',
                           'Underlying Health Condition']]
master_data.set_index('Date', inplace=True)
master_data.to_csv('C:/APIT/CovidData.csv')

z = bcr.bar_chart_race(master_data, filename='C:/APIT/data.mp4', orientation='h', sort='asc',
                       title="NHS GGC Coronavirus Absence Coding")
