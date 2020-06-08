'''This file takes a single day's SSTS Absence 6a report output <automatically
produced by the accompanying absence-daily-script.py>.
 It produces several pivot tables then emails them to the relevant people as required.'''

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import date

newpath = 'W:/Daily_Absence/' + (date.today()).strftime("%Y-%m-%d") + '.xls'
df = pd.read_excel(newpath, skiprows=4)
sd = pd.read_excel(
    'W:/Workforce Monthly Reports/Monthly_Reports/Mar-20 Snapshot/Staff Download/2020-03 - Staff Download - GGC.xls')
phones = pd.read_excel('W:/MFT/phone number lookup.xlsx')
manager = pd.read_excel('W:/Daily_Absence/manager_lookup.xlsx')
manager = manager[['Pay_Number', 'Supervisor email address', 'Work Email Address']]
print(df.columns)
print(sd.columns)
print(manager.columns)

df = df.rename(columns={'Pay No': 'Pay_Number'})
df['AbsenceReason Description'].replace({'Infectious diseases': 'Coronavirus – Covid 19 Positive'},
                                        inplace=True)
df = df.merge(sd, on='Pay_Number', how='left')
df = df.merge(phones, on="Pay_Number", how='left')
df = df.merge(manager, on="Pay_Number", how='left')
print(df.columns)

print(df['AbsenceReason Description'].value_counts())

all_tpi = df[df['AbsenceReason Description'] == 'Coronavirus – Test and Protect Isolation']
tpi_piv = pd.pivot_table(all_tpi, values='Pay_Number',
                                  index='Sector/Directorate/HSCP',
                                  aggfunc='count',
                                  fill_value=0)
print(tpi_piv)

print(df['Job_Family'].value_counts())

df_nursedocs = df[(df['Job_Family'] == 'Nursing and Midwifery') | (df['Job_Family'] == 'Medical and Dental')]

west_dun = df[df['Sector/Directorate/HSCP'] == 'West Dunbartonshire HSCP']
west_dun_piv = pd.pivot_table(west_dun, index=['Sub-Directorate 1', 'Sub-Directorate 2', 'AbsenceReason Description',
                                               'department', 'Post_Descriptor'], values=['WTE', 'Pay_Number'],
                              aggfunc={'WTE': np.sum, 'Pay_Number': 'count'}).round(1)
west_dun_piv.reset_index(inplace=True)
print(west_dun_piv.columns)
west_dun_piv.rename(
    columns={'Pay_Number': 'Headcount', 'Sub-Directorate 1': 'Sub-Dir 1', 'Sub-Directorate 2': 'Sub-Dir 2',
             'AbsenceReason Description': 'Reason'}, inplace=True)
west_dun_piv['Reason'].replace({'Coronavirus – Self displaying symptoms – Self Isolating': 'Covid-19 - Self Isolating',
                                'Coronavirus': 'Covid-19 - Carer/Parental Leave',
                                'Coronavirus – Covid 19 Positive': 'Covid-19 - Confirmed Positive',
                                'Coronavirus – Underlying Health Condition': 'Covid-19 - Underlying Health Condition',
                                'Coronavirus – Household Related – Self Isolating': 'Covid-19 - Household Isolating'},
                               inplace=True)


west_dun_piv.to_csv('W:/Daily_Absence/West_Dun-' + (date.today()).strftime('%Y-%m-%d') + '.csv', index=False)

south_sector = df[df['Sector/Directorate/HSCP'] == 'South Sector']
south_sector_piv = pd.pivot_table(south_sector, index=['Sub-Directorate 1', 'Sub-Directorate 2',
                                                       'AbsenceReason Description'], values=['WTE', 'Pay_Number'],
                                  aggfunc={'WTE': np.sum, 'Pay_Number': 'count'}).round(1)
south_sector_piv.reset_index(inplace=True)
print(south_sector_piv.columns)
south_sector_piv.rename(columns={'Pay_Number': 'Headcount', 'Sub-Directorate 1': 'Sub-Dir 1',
                                 'Sub-Directorate 2': 'Sub-Dir 2',
                                 'AbsenceReason Description': 'Reason'}, inplace=True)
south_sector_piv['Reason'].replace(
    {'Coronavirus – Self displaying symptoms – Self Isolating': 'Covid-19 - Self Isolating',
     'Coronavirus': 'Covid-19 - Carer/Parental Leave',
     'Coronavirus – Covid 19 Positive': 'Covid-19 - Confirmed Positive',
     'Coronavirus – Underlying Health Condition': 'Covid-19 - Underlying Health Condition',
     'Coronavirus – Household Related – Self Isolating': 'Covid-19 - Household Isolating'}, inplace=True)



south_sector_piv.to_csv('W:/Daily_Absence/' + 'South-Sector-' + (date.today()).strftime('%Y-%m-%d') + '.csv',
                        index=False)

covid_pos_nursedocs = df_nursedocs[(df_nursedocs['AbsenceReason Description'] == 'Infectious diseases') |
                                   (df_nursedocs['AbsenceReason Description'] == 'Coronavirus – Covid 19 Positive')]

self_isolating_nursedocs = df_nursedocs[df_nursedocs['AbsenceReason Description'] ==
                                        'Coronavirus – Self displaying symptoms – Self Isolating']

all_household_isolating = df[df['AbsenceReason Description'] == 'Coronavirus – Household Related – Self Isolating']
all_underlying = df[df['AbsenceReason Description'] == 'Coronavirus – Underlying Health Condition']
nursedocs_underlying = df_nursedocs[
    df_nursedocs['AbsenceReason Description'] == 'Coronavirus – Underlying Health Condition']
nursedocs_household = df_nursedocs[df_nursedocs['AbsenceReason Description'] ==
                                   'Coronavirus – Household Related – Self Isolating']

covid_parental_nursedocs = df_nursedocs[df_nursedocs['AbsenceReason Description'] == 'Coronavirus']

all_parental = df[df['AbsenceReason Description'] == 'Coronavirus']
all_positive = df[(df['AbsenceReason Description'] == 'Infectious diseases') | (
            df['AbsenceReason Description'] == 'Coronavirus – Covid 19 Positive')]

all_isolating = df[df['AbsenceReason Description'] == 'Coronavirus – Self displaying symptoms – Self Isolating']

print(len(self_isolating_nursedocs))
all_isolators = df[(df['AbsenceReason Description'] == 'Coronavirus – Self displaying symptoms – Self Isolating') |
                   (df['AbsenceReason Description'] == 'Coronavirus – Underlying Health Condition') |
                   (df['AbsenceReason Description'] == 'Coronavirus – Household Related – Self Isolating') |
                   (df['AbsenceReason Description'] == 'Coronavirus – Test and Protect Isolation')
                   ]

covid_pos_piv = pd.pivot_table(covid_pos_nursedocs, values='Pay_Number',
                               index='Sector/Directorate/HSCP',
                               columns='Job_Family',
                               aggfunc='count',
                               fill_value=0)
print(covid_pos_piv)

covid_parental_piv = pd.pivot_table(covid_parental_nursedocs, values='Pay_Number',
                                    index='Sector/Directorate/HSCP',
                                    columns='Job_Family',
                                    aggfunc='count',
                                    fill_value=0)
print(covid_parental_piv)
self_isolating_piv = pd.pivot_table(self_isolating_nursedocs, values='Pay_Number',
                                    index='Sector/Directorate/HSCP',
                                    columns='Job_Family',
                                    aggfunc='count',
                                    fill_value=0)
all_positive_piv = pd.pivot_table(all_positive, values='Pay_Number',
                                  index='Sector/Directorate/HSCP',
                                  aggfunc='count',
                                  fill_value=0)

all_parental_piv = pd.pivot_table(all_parental, values='Pay_Number',
                                  index='Sector/Directorate/HSCP',
                                  aggfunc='count',
                                  fill_value=0)
all_isolating_piv = pd.pivot_table(all_isolating, values='Pay_Number',
                                   index='Sector/Directorate/HSCP',
                                   aggfunc='count',
                                   fill_value=0)

all_household_isolating_piv = pd.pivot_table(all_household_isolating, values='Pay_Number',
                                             index='Sector/Directorate/HSCP',
                                             aggfunc='count',
                                             fill_value=0)

all_underlying_piv = pd.pivot_table(all_underlying, values='Pay_Number',
                                    index='Sector/Directorate/HSCP',
                                    aggfunc='count',
                                    fill_value=0)

nursedocs_underlying_piv = pd.pivot_table(nursedocs_underlying, values='Pay_Number',
                                          index='Sector/Directorate/HSCP',
                                          columns='Job_Family',
                                          aggfunc='count',
                                          fill_value=0)
print(nursedocs_underlying_piv)
nursedocs_household_piv = pd.pivot_table(nursedocs_household, values='Pay_Number',
                                         index='Sector/Directorate/HSCP',
                                         columns='Job_Family',
                                         aggfunc='count',
                                         fill_value=0)
print(nursedocs_underlying_piv)
exit()


# print(self_isolating_piv)


def graph_maker_all(data, graph_title):
    plt.style.use('seaborn')
    ax = data.plot(kind='bar', color='#003087', legend=False)
    plt.xticks(fontsize=7)
    plt.title(graph_title)
    height = (max(data.values))
    print(data.columns)

    for index, z in enumerate(data['Pay_Number']):
        label = z
        plt.annotate(label,
                     xy=(index, z + height / 40),
                     ha='center')
        # plt.text(x=index, y=z+height/20, s=z)
    plt.setp(ax.get_xticklabels(), rotation=50, horizontalalignment='right')
    plt.tight_layout()
    plt.savefig('C:/Covid_Graphs/' + graph_title, dpi=300)
    plt.close()
    # for i, each in enumerate(data.index):
    #     for col in data.columns:
    #         y = round(data.ix[each][col], 1)
    #         ax.text(i + 0.05, y + jtmax / 20, y)


def graph_maker_docs_and_nurses(i, graph_title):
    nm = i['Nursing and Midwifery'].to_list()
    md = i['Medical and Dental'].to_list()

    plt.style.use('seaborn')
    width = 0.35
    ind = np.arange(len(md))

    fig, ax = plt.subplots()

    rects1 = ax.bar(ind - width / 2, nm, width, label='Nursing & Midwifery')
    rects2 = ax.bar(ind + width / 2, md, width, label='Medical & Dental')

    plt.title(graph_title)
    ax.set_xticks(ind)
    ax.set_xticklabels(i.index)
    plt.legend(loc='best')
    plt.ylabel('Headcount')
    plt.xticks(fontsize=7)
    autolabel(ax, rects1, "left")
    autolabel(ax, rects2, "right")
    plt.setp(ax.get_xticklabels(), rotation=50, horizontalalignment='right')
    plt.tight_layout()
    plt.savefig('C:/Covid_Graphs/' + graph_title, dpi=300)
    plt.close()


def autolabel(ax, rects, xpos='center'):
    ha = {'center': 'center', 'right': 'left', 'left': 'right'}
    offset = {'center': 0, 'right': 1, 'left': -1}
    for rect in rects:
        height = rect.get_height()
        ax.annotate('{}'.format(height),
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 3),  # 3 points vertical offset
                    textcoords="offset points",
                    ha='center', va='bottom')


all_underlying.to_excel('W:/Daily_Absence/underlying' + (date.today()).strftime('%Y-%m-%d') + '.xlsx')
graph_maker_docs_and_nurses(nursedocs_household_piv,
                            "Special Leave SP - Coronavirus – Household Related – Self Isolating - Clinical")
graph_maker_docs_and_nurses(self_isolating_piv,
                            "Special Leave SP - Coronavirus – Self displaying symptoms – Self Isolating - Clinical")
graph_maker_docs_and_nurses(covid_parental_piv, "Special Leave SP - Coronavirus Parental Leave - Clinical")

graph_maker_all(all_isolating_piv,
                "Special Leave SP - Coronavirus – Self displaying symptoms – Self Isolating - All Staff")
graph_maker_all(all_parental_piv, "Special Leave SP - Coronavirus Parental Leave - All staff")
graph_maker_all(all_positive_piv, "Special Leave SP - Coronavirus - Covid-19 Confirmed - All staff")
graph_maker_all(all_underlying_piv, "Special Leave SP - Coronavirus – Underlying Health Condition - All staff")
graph_maker_all(all_household_isolating_piv,
                "Special Leave SP - Coronavirus – Household Related – Self Isolating - All staff")
df_isolating_sheet = all_isolators[['Pay_Number', 'Supervisor email address', 'Forename', 'Surname', 'Date_of_Birth',
                                    'AbsenceReason Description', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
                                    'department', 'Job_Family', 'Absence Episode Start Date', 'Address_Line_1',
                                    'Address_Line_2', 'Address_Line_3', 'Postcode', 'Best Phone']]
df_isolating_sheet.to_excel('W:/daily_absence/isolators-' + (date.today()).strftime('%Y-%m-%d') + '.xlsx', index=False)
df_isolating_south = df_isolating_sheet[df_isolating_sheet['Sector/Directorate/HSCP'] == 'South Sector']
df_isolating_south.to_excel('W:/daily_absence/south-isolators-' + (date.today()).strftime('%Y-%m-%d') + '.xlsx',
                            index=False)
all_pos_sheet = all_positive[
    ['Pay_Number', 'Supervisor email address', 'Forename', 'Surname', 'AbsenceReason Description',
     'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'department', 'Job_Family',
     'Absence Episode Start Date', 'Address_Line_1', 'Address_Line_2', 'Address_Line_3',
     'Postcode', 'Best Phone', 'Date_of_Birth', 'Date_Started', 'Date_To_Grade',
     'Date_Superannuation_Started', 'SB_Number']]
all_pos_sheet.to_excel('W:/daily_absence/positive-' + (date.today()).strftime('%Y-%m-%d') + '.xlsx', index=False)
print("Self isolating - " + str(len(all_isolating)))
print("Underlying Conditions - " + str(len(all_underlying)))
print("Covid Parental Leave - " + str(len(all_parental)))
print("Household isolating - " + str(len(all_household_isolating)))
print("Covid Positive - " + str(len(all_positive)))
print("Covid - Test and protect - "+str(len(tpi_piv)))
graph_maker_all(tpi_piv, "Special Leave SP - Coronavirus – Test and Protect Isolation")
graph_maker_docs_and_nurses(nursedocs_underlying_piv,
                            "Special Leave SP - Coronavirus – Underlying Health Condition - Clinical")
graph_maker_docs_and_nurses(covid_pos_piv, "Special Leave SP - Coronavirus - Covid-19 Confirmed - Clinical")
