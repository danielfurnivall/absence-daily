'''This file takes a single day's SSTS Absence 6a report output <automatically
produced by the accompanying absence-daily-script.py>.
 It produces several pivot tables then emails them to the relevant people as required.'''

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import date
import sys

def graph_maker_all(data, graph_title):

    for i in dirs:
        if i in data.index:
            pass
        else:
            data.loc[i] = [0]
    data.sort_index(inplace=True)
    plt.style.use('seaborn')
    ax = data.plot(kind='bar', color='#003087', legend=False)
    plt.xticks(fontsize=7)
    plt.title(graph_title)
    height = (max(data.values))

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


# get path of today's abs file
newpath = 'W:/Daily_Absence/' + (date.today()).strftime("%Y-%m-%d") + '.xls'

# read in abs file and staff download
df = pd.read_excel(newpath, skiprows=4)
sd = pd.read_excel(
    'W:/Workforce Monthly Reports/Monthly_Reports/Sep-20 Snapshot/Staff Download/2020-09 - Staff Download - GGC.xls')
# read in phone number lookup
phones = pd.read_excel('W:/MFT/phone number lookup.xlsx')
# read in eESS emails
manager = pd.read_excel('W:/Daily_Absence/eESS-emails.xlsx')
manager = manager[['Pay_Number', 'Supervisor email address', 'Work Email Address']]

all_covid_reasons = ['Coronavirus – Household Related – Self Isolating', 'Coronavirus – Underlying Health Condition',
                     'Coronavirus – Covid 19 Positive', 'Coronavirus',
                     'Coronavirus – Self displaying symptoms – Self Isolating', 'Coronavirus – Quarantine',
                     'Coronavirus – Test and Protect Isolation'
                     ]



# rename infectious diseases to covid positive - this change has been active since march, and may be inaccurate now.
df = df.rename(columns={'Pay No': 'Pay_Number'})
df['AbsenceReason Description'].replace({'Infectious diseases': 'Coronavirus – Covid 19 Positive'},
                                        inplace=True)

# get path of yesterday's abs file
yesterday = 'W:/Daily_Absence/' + (date.today() - pd.DateOffset(days=1)).strftime("%Y-%m-%d") + '.xls'
df_yesterday = pd.read_excel(yesterday, skiprows=4)
df_yesterday['AbsenceReason Description'].replace({'Infectious diseases': 'Coronavirus – Covid 19 Positive'},
                                        inplace=True)
df_yesterday = df_yesterday.rename(columns={'Pay No': 'Pay_Number'})

# merge files into df
df = df.merge(sd, on='Pay_Number', how='left')
df = df.merge(phones, on="Pay_Number", how='left')
df = df.merge(manager, on="Pay_Number", how='left')
df_yesterday = df_yesterday.merge(sd, on='Pay_Number', how='left')
df_today = df
df_today = df_today[df_today['AbsenceReason Description'].isin(all_covid_reasons)]
df_yesterday = df_yesterday[df_yesterday['AbsenceReason Description'].isin(all_covid_reasons)]
df_yesterday = df_yesterday[['Pay_Number', 'Forename', 'Surname', 'Roster Location', 'Absence Type', 'AbsenceReason Description', 'Absence Episode Start Date',
         'Absence Episode End Date','Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2', 'department',
         'Base', 'Job_Family', 'Sub_Job_Family', 'Post_Descriptor']]
df_yesterday['AbsenceReason Description'] = df_yesterday['AbsenceReason Description'].map(
        {'Coronavirus – Self displaying symptoms – Self Isolating': 'Self Isolating',
         'Coronavirus': 'Carer and Parental Leave',
         'Coronavirus – Covid 19 Positive': 'Covid Positive',
         'Coronavirus – Underlying Health Condition': 'Underlying Health Condition',
         'Coronavirus – Household Related – Self Isolating': 'Household Isolating',
         'Coronavirus – Test and Protect Isolation': 'Test and protect',
         'Coronavirus – Quarantine':'Quarantine'})
df_today = df_today[['Pay_Number', 'Forename', 'Surname', 'Roster Location', 'Absence Type', 'AbsenceReason Description', 'Absence Episode Start Date',
         'Absence Episode End Date','Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2', 'department',
         'Base', 'Job_Family', 'Sub_Job_Family', 'Post_Descriptor']]
df_today['AbsenceReason Description'] = df_today['AbsenceReason Description'].map(
        {'Coronavirus – Self displaying symptoms – Self Isolating': 'Self Isolating',
         'Coronavirus': 'Carer and Parental Leave',
         'Coronavirus – Covid 19 Positive': 'Covid Positive',
         'Coronavirus – Underlying Health Condition': 'Underlying Health Condition',
         'Coronavirus – Household Related – Self Isolating': 'Household Isolating',
         'Coronavirus – Test and Protect Isolation': 'Test and protect',
         'Coronavirus – Quarantine':'Quarantine'})

# Gillian Ayling-Whitehouse's daily absence file
df_today1 = df_today[~(df_today['Pay_Number'].isin(df_yesterday['Pay_Number'].unique()))]
df_yesterday = df_yesterday[~(df_yesterday['Pay_Number'].isin(df_today['Pay_Number'].unique()))]
df_today = df_today1
print(len(df_today))
print(len(df_yesterday))

with pd.ExcelWriter('W:/daily_absence/new_old_covid-'+(pd.Timestamp.now()).strftime('%Y-%m-%d') + '.xlsx',
                    engine='xlsxwriter') as writer:
    workbook = writer.book
    header_format = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 16})
    subheader = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 14, 'text_wrap': 1})
    subheadernoBold = workbook.add_format({'font_name': 'Arial', 'font_size': 14})
    table_format_ul = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white', 'underline': True})
    table_format = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white', 'align':'center'})
    back_button = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white',
                                       'underline': True, 'align': 'center', 'valign': 'vcenter'})
    cluster_format = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'red'})

    worksheet = workbook.add_worksheet('Contents')
    worksheet.hide_gridlines(2)
    worksheet.set_column('A:A', 26)
    worksheet.set_column('B:H', 20)

    # write headers
    worksheet.write(0, 0, 'Coronavirus - Daily Change Report', header_format)
    worksheet.write(1, 0, 'Date:', subheader)
    worksheet.write(1, 1, f'{pd.Timestamp.now().strftime("%d %B %Y")}', subheadernoBold)
    worksheet.insert_image('D1', 'W:/Danny/ggclogo.jpg', {'x_scale': 0.25, 'y_scale': 0.25})

    worksheet.write(5, 0, f'Covid Positive', subheader)
    worksheet.write(6, 0, f'Underlying Health Condition', subheader)
    worksheet.write(7, 0, f'Household Isolating', subheader)
    worksheet.write(8, 0, f'Self Isolating', subheader)
    worksheet.write(9, 0, f'Test & Protect', subheader)
    worksheet.write(10, 0, f'Quarantine', subheader)
    worksheet.write(11, 0, f'Parental/Carer Leave', subheader)
    worksheet.write(4, 1, f'New Cases', subheader)
    worksheet.write(4, 2, f'Expiring Cases', subheader)
    worksheet.write(5, 1, len(df_today[df_today["AbsenceReason Description"]=="Covid Positive"]), table_format)
    worksheet.write(6, 1, len(df_today[df_today["AbsenceReason Description"] == "Underlying Health Condition"]), table_format)
    worksheet.write(7, 1, len(df_today[df_today["AbsenceReason Description"] == "Household Isolating"]), table_format)
    worksheet.write(8, 1, len(df_today[df_today["AbsenceReason Description"] == "Self Isolating"]), table_format)
    worksheet.write(9, 1, len(df_today[df_today["AbsenceReason Description"] == "Test and Protect"]), table_format)
    worksheet.write(10, 1, len(df_today[df_today["AbsenceReason Description"] == "Quarantine"]), table_format)
    worksheet.write(11, 1, len(df_today[df_today["AbsenceReason Description"] == "Carer and Parental Leave"]), table_format)
    worksheet.write(5, 2, len(df_yesterday[df_yesterday["AbsenceReason Description"] == "Covid Positive"]), table_format)
    worksheet.write(6, 2, len(df_yesterday[df_yesterday["AbsenceReason Description"] == "Underlying Health Condition"]),
                    table_format)
    worksheet.write(7, 2, len(df_yesterday[df_yesterday["AbsenceReason Description"] == "Household Isolating"]), table_format)
    worksheet.write(8, 2, len(df_yesterday[df_yesterday["AbsenceReason Description"] == "Self Isolating"]), table_format)
    worksheet.write(9, 2, len(df_yesterday[df_yesterday["AbsenceReason Description"] == "Test and Protect"]), table_format)
    worksheet.write(10, 2, len(df_yesterday[df_yesterday["AbsenceReason Description"] == "Quarantine"]), table_format)
    worksheet.write(11, 2, len(df_yesterday[df_yesterday["AbsenceReason Description"] == "Carer and Parental Leave"]),
                    table_format)

    df_today.to_excel(writer, index=False, startrow=3, sheet_name = 'New')
    worksheet = writer.sheets['New']
    worksheet.write('A1', 'New Covid Absences - ' + pd.Timestamp.now().strftime('%d %B %Y'), header_format)
    worksheet.set_column('A:P', 20)
    df_yesterday.to_excel(writer, index=False, startrow=3, sheet_name = 'Expiring')
    worksheet = writer.sheets['Expiring']
    worksheet.write('A1', 'Expiring Covid Absences - ' + pd.Timestamp.now().strftime('%d %B %Y'), header_format)
    worksheet.set_column('A:P', 20)





# New file for all absence - Steven 04-11-20
all_abs = df[['Pay_Number', 'Forename', 'Surname', 'Roster Location', 'Absence Type', 'AbsenceReason Description', 'Absence Episode Start Date',
         'Absence Episode End Date','Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2', 'department',
         'Base', 'Job_Family', 'Sub_Job_Family', 'Post_Descriptor']]


all_abs['Forename'].loc[all_abs['Sector/Directorate/HSCP'].isna()] = "New staff - No payroll data yet"
with pd.ExcelWriter('W:/daily_absence/all_absence-' + (pd.Timestamp.now()).strftime('%Y-%m-%d') + '.xlsx',
                    engine='xlsxwriter') as writer:
    all_abs.to_excel(writer, index=False, startrow=3, sheet_name='Data')

    header_format = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 16})
    worksheet = writer.sheets['Data']
    worksheet.set_column('A:P', 20)
    worksheet.insert_image('D1', 'W:/Danny/ggclogo.jpg', {'x_scale': 0.25, 'y_scale': 0.25})
    worksheet.write('A1', 'Daily Absence Data - '+ pd.Timestamp.now().strftime('%d %B %Y'), header_format)
    worksheet.freeze_panes(3, 0)
    worksheet.autofilter(f'A4:P{len(df)+4}')

df_abs = df[['Pay_Number', 'Forename', 'Surname', 'Roster Location', 'Absence Type', 'AbsenceReason Description', 'Absence Episode Start Date',
         'Absence Episode End Date','Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2', 'department',
         'Base', 'Job_Family', 'Sub_Job_Family', 'Post_Descriptor']]
df_abs = df_abs[df_abs['AbsenceReason Description'].isin(all_covid_reasons)]
df_abs['Forename'].loc[df_abs['Sector/Directorate/HSCP'].isna()] = "New staff - No payroll data yet"
df_abs['Sector/Directorate/HSCP'].loc[df_abs['Sector/Directorate/HSCP'].isna()] = "New staff"
df_abs['Sector/Directorate/HSCP'].loc[df_abs['Sector/Directorate/HSCP'] == 'East Dunbartonshire - Oral Health'] = 'East Dun Oral Health'
df_abs['Sector/Directorate/HSCP'].loc[df_abs['Sector/Directorate/HSCP'] == "Women & Children's"] = "Women and Children's"
with pd.ExcelWriter('W:/daily_absence/all_covid_absence-' + (pd.Timestamp.now()).strftime('%Y-%m-%d') + '.xlsx',
                    engine='xlsxwriter') as writer:
    workbook = writer.book
    # cover sheet
    header_format = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 16})
    subheader = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 14, 'text_wrap':1})
    subheadernoBold = workbook.add_format({'font_name': 'Arial', 'font_size': 14})
    table_format_ul = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white', 'underline': True})
    table_format = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white'})
    back_button = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'white',
                                       'underline': True, 'align': 'center', 'valign': 'vcenter'})
    cluster_format = workbook.add_format({'bg_color': '#005EB8', 'font_color': 'red'})
    # create cover page
    worksheet = workbook.add_worksheet('Contents')
    worksheet.hide_gridlines(2)
    worksheet.set_column('A:A', 26)
    worksheet.set_column('B:H', 20)

    # write headers
    worksheet.write(0, 0, 'Coronavirus - Sector Summary Report', header_format)
    worksheet.write(1, 0, 'Date:', subheader)
    worksheet.write(1, 1, f'{pd.Timestamp.now().strftime("%d %B %Y")}', subheadernoBold)
    worksheet.insert_image('D1', 'W:/Danny/ggclogo.jpg', {'x_scale': 0.25, 'y_scale': 0.25})

    worksheet.write(3, 0, f'Sector/Directorate/HSCP', subheader)
    worksheet.write(3, 1, f'Covid Positive', subheader)
    worksheet.write(3, 2, f'Underlying Health Condition', subheader)
    worksheet.write(3, 3, f'Household Isolating', subheader)
    worksheet.write(3, 4, f'Self Isolating', subheader)
    worksheet.write(3, 5, f'Test & Protect', subheader)
    worksheet.write(3, 6, f'Quarantine', subheader)
    worksheet.write(3, 7, f'Covid Parental/Carer Leave', subheader)

    rownum = 4
    for i in df_abs['Sector/Directorate/HSCP'].unique():
        worksheet.write_url(rownum, 0, "internal:'" + i + "'!A1", string=i, cell_format=table_format_ul)
        worksheet.write(rownum, 1, len(df_abs[(df_abs['Sector/Directorate/HSCP'] == i) &
                                              (df_abs['AbsenceReason Description'] == 'Coronavirus – Covid 19 Positive')]), table_format)
        worksheet.write(rownum, 2, len(df_abs[(df_abs['Sector/Directorate/HSCP'] == i) &
                                              (df_abs['AbsenceReason Description'] == 'Coronavirus – Underlying Health Condition')]), table_format)
        worksheet.write(rownum, 3, len(df_abs[(df_abs['Sector/Directorate/HSCP'] == i) &
                                              (df_abs['AbsenceReason Description'] == 'Coronavirus – Household Related – Self Isolating')]), table_format)
        worksheet.write(rownum, 4, len(df_abs[(df_abs['Sector/Directorate/HSCP'] == i) &
                                              (df_abs['AbsenceReason Description'] == 'Coronavirus – Self displaying symptoms – Self Isolating')]), table_format)
        worksheet.write(rownum, 5, len(df_abs[(df_abs['Sector/Directorate/HSCP'] == i) &
                                              (df_abs['AbsenceReason Description'] == 'Coronavirus – Test and Protect Isolation')]), table_format)
        worksheet.write(rownum, 6, len(df_abs[(df_abs['Sector/Directorate/HSCP'] == i) &
                                              (df_abs['AbsenceReason Description'] == 'Coronavirus – Quarantine')]), table_format)
        worksheet.write(rownum, 7, len(df_abs[(df_abs['Sector/Directorate/HSCP'] == i) &
                                              (df_abs['AbsenceReason Description'] == 'Coronavirus')]), table_format)

        rownum+=1
    df_abs['AbsenceReason Description'] = df_abs['AbsenceReason Description'].map(
        {'Coronavirus – Self displaying symptoms – Self Isolating': 'Self Isolating',
         'Coronavirus': 'Carer and Parental Leave',
         'Coronavirus – Covid 19 Positive': 'Covid Positive',
         'Coronavirus – Underlying Health Condition': 'Underlying Health Condition',
         'Coronavirus – Household Related – Self Isolating': 'Household Isolating',
         'Coronavirus – Test and Protect Isolation': 'Test and protect',
         'Coronavirus – Quarantine':'Quarantine'})
    for i in df_abs['Sector/Directorate/HSCP'].unique():
        print(i)


        df_abs[df_abs['Sector/Directorate/HSCP'] == i].to_excel(writer, index=False, startrow=3, sheet_name=i)

        worksheet = writer.sheets[i]
        worksheet.set_column('A:P', 20)
        worksheet.write('A1', f'{i} - Absence Data - '+ pd.Timestamp.now().strftime('%d %B %Y'), header_format)
        worksheet.freeze_panes(3, 0)
        worksheet.write_url('D2', "internal:'Contents'!A1",
                        string='Back', cell_format=back_button)
        worksheet.autofilter(f'A4:P{len(df_abs) + 4}')




# This is to build the all covid named list and the all covid pivot

all_covid_reasons = df[df['AbsenceReason Description'].isin(all_covid_reasons)]
all_covid_reasons['Sector/Directorate/HSCP'].loc[all_covid_reasons['Sector/Directorate/HSCP'].isna()] = "New staff - no org structure"
all_covid_reasons['Job_Family'].loc[all_covid_reasons['Sector/Directorate/HSCP'].isna()] = "New staff"
all_covid_reasons.to_csv('W:/daily_absence/all_covid_names.csv')


all_covid_piv = pd.pivot_table(all_covid_reasons, values='Pay_Number',
                               index='Sector/Directorate/HSCP',
                               aggfunc='count',
                               fill_value=0,
                               dropna=False)

all_covid_piv_jobfam = pd.pivot_table(all_covid_reasons, values='Pay_Number',
                               index='Sector/Directorate/HSCP',
                               columns='Job_Family',
                               aggfunc='count',
                               fill_value=0,
                                      dropna=False)
all_covid_piv_jobfam.to_excel('W:/daily_absence/all_covid'+(date.today()).strftime('%Y-%m-%d')+'.xlsx')


# get all quarantine staff
quarantine_new = df[df['AbsenceReason Description'] == 'Coronavirus – Quarantine']

# get all test and protect staff
all_tpi = df[df['AbsenceReason Description'] == 'Coronavirus – Test and Protect Isolation']
# build test and protect staff
tpi_piv = pd.pivot_table(all_tpi, values='Pay_Number',
                                  index='Sector/Directorate/HSCP',
                                  aggfunc='count',
                                  fill_value=0)

# build file for gillian gall
west_dun = df[df['Sector/Directorate/HSCP'] == 'West Dunbartonshire HSCP']
west_dun_piv = pd.pivot_table(west_dun, index=['Sub-Directorate 1', 'Sub-Directorate 2', 'AbsenceReason Description',
                                               'department', 'Post_Descriptor'], values=['WTE', 'Pay_Number'],
                              aggfunc={'WTE': np.sum, 'Pay_Number': 'count'}).round(1)
west_dun_piv.reset_index(inplace=True)

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


# south sector files for david dall and ruth campbell
south_sector = df[df['Sector/Directorate/HSCP'] == 'South Sector']
south_sector_piv = pd.pivot_table(south_sector, index=['Sub-Directorate 1', 'Sub-Directorate 2',
                                                       'AbsenceReason Description'], values=['WTE', 'Pay_Number'],
                                  aggfunc={'WTE': np.sum, 'Pay_Number': 'count'}).round(1)
south_sector_piv.reset_index(inplace=True)

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


# build datasets for each covid abs reason
all_household_isolating = df[df['AbsenceReason Description'] == 'Coronavirus – Household Related – Self Isolating']
all_underlying = df[df['AbsenceReason Description'] == 'Coronavirus – Underlying Health Condition']

all_parental = df[df['AbsenceReason Description'] == 'Coronavirus']
all_positive = df[(df['AbsenceReason Description'] == 'Infectious diseases') | (
            df['AbsenceReason Description'] == 'Coronavirus – Covid 19 Positive')]

all_isolating = df[df['AbsenceReason Description'] == 'Coronavirus – Self displaying symptoms – Self Isolating']


all_isolators = df[(df['AbsenceReason Description'] == 'Coronavirus – Self displaying symptoms – Self Isolating') |
                   (df['AbsenceReason Description'] == 'Coronavirus – Underlying Health Condition') |
                   (df['AbsenceReason Description'] == 'Coronavirus – Household Related – Self Isolating') |
                   (df['AbsenceReason Description'] == 'Coronavirus – Test and Protect Isolation')
                   ]

#build pivs to inform graphs
all_positive_piv = pd.pivot_table(all_positive, values='Pay_Number',
                                  index='Sector/Directorate/HSCP',
                                  aggfunc='count',
                                  fill_value=0)

quarantine_piv = pd.pivot_table(quarantine_new, values='Pay_Number',
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


dirs = df['Sector/Directorate/HSCP'].unique().tolist()
dirs = [dir for dir in dirs if str(dir) != "nan"]

all_underlying.to_excel('W:/Daily_Absence/underlying' + (date.today()).strftime('%Y-%m-%d') + '.xlsx')
all_parental.to_excel('W:/Daily_Absence/parental' + (date.today()).strftime('%Y-%m-%d') + '.xlsx')

graph_maker_all(all_isolating_piv,
                "Special Leave SP - Coronavirus – Self displaying symptoms – Self Isolating - All Staff")
graph_maker_all(all_parental_piv, "Special Leave SP - Coronavirus Parental Leave - All staff")
graph_maker_all(all_positive_piv, "Special Leave SP - Coronavirus - Covid-19 Confirmed - All staff")
graph_maker_all(all_underlying_piv, "Special Leave SP - Coronavirus – Underlying Health Condition - All staff")
graph_maker_all(all_household_isolating_piv,
                "Special Leave SP - Coronavirus – Household Related – Self Isolating - All staff")
graph_maker_all(quarantine_piv, "Special Leave SP - Coronavirus - Quarantine (new code)")
df_isolating_sheet = all_isolators[['Pay_Number', 'Supervisor email address', 'Forename', 'Surname', 'Date_of_Birth',
                                    'AbsenceReason Description', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
                                    'department', 'Job_Family', 'Absence Episode Start Date', 'Address_Line_1',
                                    'Address_Line_2', 'Address_Line_3', 'Postcode', 'Best Phone']]
df_isolating_sheet.to_excel('W:/daily_absence/isolators-' + (date.today()).strftime('%Y-%m-%d') + '.xlsx', index=False)
df_isolating_south = df_isolating_sheet[df_isolating_sheet['Sector/Directorate/HSCP'] == 'South Sector']
df_isolating_south.to_excel('W:/daily_absence/south-isolators-' + (date.today()).strftime('%Y-%m-%d') + '.xlsx',
                            index=False)
graph_maker_all(all_covid_piv, "All Covid-related Absence Reasons")

all_pos_ICData = all_positive[[
    'Pay_Number','Forename','Surname','Date_of_Birth', 'Roster Location', 'department', 'Sector/Directorate/HSCP',
    'Sub-Directorate 1', 'Sub-Directorate 2', 'Job_Family', 'AbsenceReason Description', 'Absence Episode Start Date'
]]
all_pos_ICData['Date_of_Birth'] = all_pos_ICData['Date_of_Birth'].dt.strftime('%d-%m-%Y')
all_pos_ICData.to_excel('W:/daily_absence/ICData-' + (date.today()).strftime('%Y-%m-%d') + '.xlsx', index=False)
all_pos_sheet = all_positive[
    ['Pay_Number', 'Supervisor email address', 'Forename', 'Surname', 'AbsenceReason Description',
     'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'department', 'Job_Family',
     'Absence Episode Start Date', 'Address_Line_1', 'Address_Line_2', 'Address_Line_3',
     'Postcode', 'Best Phone', 'Date_of_Birth', 'Date_Started', 'Date_To_Grade',
     'Date_Superannuation_Started', 'SB_Number']]
all_pos_sheet.to_excel('W:/daily_absence/positive-' + (date.today()).strftime('%Y-%m-%d') + '.xlsx', index=False)
with open('W:/daily_absence/raw_data'+ date.today().strftime('%Y-%m-%d') + '.txt', 'w') as f:
    sys.stdout = f
    print("Self isolating - " + str(len(all_isolating)))
    print("Underlying Conditions - " + str(len(all_underlying)))
    print("Covid Parental Leave - " + str(len(all_parental)))
    print("Household isolating - " + str(len(all_household_isolating)))
    print("Covid Positive - " + str(len(all_positive)))
    print("Covid - Test and protect - "+str(len(all_tpi)))
    print(f'Quarantine (New code) = {len(quarantine_new)}')

graph_maker_all(tpi_piv, "Special Leave SP - Coronavirus – Test and Protect Isolation")

