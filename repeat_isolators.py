""" This file aims to capture repeat isolators"""
import pandas as pd

all_covid_reasons = ['Coronavirus – Household Related – Self Isolating',
                     #'Coronavirus – Underlying Health Condition',
                     #'Coronavirus – Covid 19 Positive',
                     #'Coronavirus',
                     'Coronavirus – Self displaying symptoms – Self Isolating',
                     'Coronavirus – Quarantine',
                     'Coronavirus – Test and Protect Isolation'
                     ]
def read_in_files():
    sd = pd.read_excel('W:/Staff Downloads/2020-09 - Staff Download.xlsx')
    master_file = pd.read_excel('W:/daily_absence/week2.xls', skiprows=4)
    master_file = master_file[['Pay No', 'AbsenceReason Description', 'Absence Episode Start Date',
                               'Absence Episode End Date']]
    master_file.rename(columns={'Absence Episode Start Date':'Start', 'Absence Episode End Date':'End'}, inplace=True)

    print(master_file.columns)
    master_file = master_file[master_file['AbsenceReason Description'].isin(all_covid_reasons)]
    covid_afflicted = master_file['Pay No'].drop_duplicates().sort_values().to_frame()
    print(len(covid_afflicted))
    for abs_type in all_covid_reasons:

        df = master_file[master_file['AbsenceReason Description']==abs_type]

        # Attempt to remove all 14 day isolation periods
        start_date = {}
        pay_no_episodes = {}
        for i in df['Pay No'].unique():
            df_dups = df[df['Pay No'] == i]
            if len(df_dups) == 1:
                start_date[i] = df_dups['Start'].iloc[0].strftime('%d/%m') + " - " \
                                + df_dups['End'].iloc[0].strftime('%d/%m')

                pay_no_episodes[i] = 1
                continue
            dates = ""
            for row in df_dups.itertuples():
                # print(row)
                dates += row.Start.strftime('%d/%m')+"-" + row.End.strftime('%d/%m') + ',\n'
            dates = dates[:-2]
            #
            #df_dups.loc[:, Start'] = max(df_dups[Start'])
            pay_no_episodes[i] = len(df_dups)
            start_date[i] = dates
        covid_afflicted[abs_type + " Date"] = covid_afflicted['Pay No'].map(start_date)
        covid_afflicted[abs_type + " Episodes"] = covid_afflicted['Pay No'].map(pay_no_episodes)

    covid_afflicted.rename(columns={'Pay No':'Pay_Number',
                                    'Coronavirus – Household Related – Self Isolating Date':'Household Dates',
                                   'Coronavirus – Self displaying symptoms – Self Isolating Date':'Self symptoms Dates',
                                   'Coronavirus – Quarantine Date': 'Quarantine Dates',
                                   'Coronavirus – Test and Protect Isolation Date': 'T&P Dates',
                                   'Coronavirus – Household Related – Self Isolating Episodes':'Household Episodes',
                                   'Coronavirus – Self displaying symptoms – Self Isolating Episodes': 'Self symptoms Episodes',
                                   'Coronavirus – Quarantine Episodes': 'Quarantine Episodes',
                                   'Coronavirus – Test and Protect Isolation Episodes':'T&P Episodes'}, inplace=True)

    covid_afflicted['Total Episodes'] = covid_afflicted[['Household Episodes', 'Self symptoms Episodes',
       'Quarantine Episodes', 'T&P Episodes']].sum(axis=1)
    covid_afflicted = covid_afflicted.merge(sd, on='Pay_Number', how='left')
    covid_afflicted = covid_afflicted[['Pay_Number','department','Forename', 'Surname', 'Household Dates', 'Self symptoms Dates',
    'Quarantine Dates', 'T&P Dates', 'Household Episodes', 'Self symptoms Episodes', 'Quarantine Episodes',
    'T&P Episodes', 'Total Episodes', 'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2',
    'Cost_Centre', 'Base', 'Job_Family', 'Sub_Job_Family']]
    with pd.ExcelWriter('W:/Daily_Absence/covid_historical.xlsx', engine='xlsxwriter') as writer:
        workbook = writer.book
        wrap_format = workbook.add_format({'text_wrap':1})
        header_format = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 16})

        covid_afflicted.to_excel(writer, index=False, sheet_name='Data', startrow=3)
        # xlsxwriter bits
        page = writer.sheets['Data']
        page.insert_image('A1', 'W:/Danny/ggclogo.jpg', {'x_scale': 0.2, 'y_scale': 0.2, 'x_offset':20})
        page.write('B1',
                   f'Covid-19 - All covid absence codes - data valid until '
                   f'{max(master_file["Start"]).strftime("%d-%m-%Y")}', header_format)
        page.set_column('A:H', width=18, cell_format=wrap_format)
        page.set_column('N:T', width=18, cell_format=wrap_format)
        page.autofilter(f'A4:T{len(covid_afflicted) + 4}')
        page.freeze_panes(4, 0)

    print(covid_afflicted.columns)

read_in_files()