import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import date

newpath = '/media/wdrive/Daily_Absence/' + '2020-03-24' + '.xls'
df = pd.read_excel(newpath, skiprows=4)
sd = pd.read_excel('/media/wdrive/Workforce Monthly Reports/Monthly_Reports/Feb-20 Snapshot/Staff Download/2020-02 - Staff Download - GGC.xls')
phones = pd.read_excel('/media/wdrive/MFT/phone number lookup.xlsx')
manager = pd.read_excel('/media/wdrive/Daily_Absence/manager_lookup.xlsx')
manager = manager[['Pay_Number', 'Supervisor email address']]
print(df.columns)
print(sd.columns)
print(manager.columns)


df = df.rename(columns={'Pay No': 'Pay_Number'})
df = df.merge(sd, on='Pay_Number')
df = df.merge(phones, on="Pay_Number")
df = df.merge(manager, on="Pay_Number")
print(df.columns)

print(df['AbsenceReason Description'].value_counts())
print(df['Job_Family'].value_counts())

df_nursedocs = df[(df['Job_Family'] == 'Nursing and Midwifery') | (df['Job_Family'] == 'Medical and Dental')]


covid_pos_nursedocs = df_nursedocs[(df_nursedocs['AbsenceReason Description'] == 'Infectious diseases') |
                                   (df_nursedocs['AbsenceReason Description'] == 'Coronavirus – Covid 19 Positive')]

self_isolating_nursedocs = df_nursedocs[df_nursedocs['AbsenceReason Description'] == 'Coronavirus – Self Isolating']

covid_parental_nursedocs = df_nursedocs[df_nursedocs['AbsenceReason Description'] == 'Coronavirus']

all_parental = df[df['AbsenceReason Description'] == 'Coronavirus']
all_positive = df[(df['AbsenceReason Description'] == 'Infectious diseases') |
                                   (df['AbsenceReason Description'] == 'Coronavirus – Covid 19 Positive')]

all_isolating = df[df['AbsenceReason Description'] == 'Coronavirus – Self Isolating']

print(len(self_isolating_nursedocs))

covid_pos_piv = pd.pivot_table(covid_pos_nursedocs, values='Pay_Number',
                               index='Sector/Directorate/HSCP',
                               columns='Job_Family',
                               aggfunc = 'count',
                               fill_value=0)
print(covid_pos_piv)

covid_parental_piv = pd.pivot_table(covid_parental_nursedocs, values='Pay_Number',
                                    index='Sector/Directorate/HSCP',
                                    columns='Job_Family',
                                    aggfunc = 'count',
                                    fill_value=0)
print(covid_parental_piv)
self_isolating_piv = pd.pivot_table(self_isolating_nursedocs, values='Pay_Number',
                                    index='Sector/Directorate/HSCP',
                                    columns='Job_Family',
                                    aggfunc = 'count',
                                    fill_value=0)
all_positive_piv = pd.pivot_table(all_positive, values='Pay_Number',
                                    index='Sector/Directorate/HSCP',
                                    aggfunc = 'count',
                                    fill_value=0)

all_parental_piv = pd.pivot_table(all_parental, values='Pay_Number',
                                    index='Sector/Directorate/HSCP',
                                    aggfunc = 'count',
                                    fill_value=0)
all_isolating_piv = pd.pivot_table(all_isolating, values='Pay_Number',
                                    index='Sector/Directorate/HSCP',
                                    aggfunc = 'count',
                                    fill_value=0)
print(self_isolating_piv)



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
                     xy = (index, z+height/40),
                     ha='center')
        #plt.text(x=index, y=z+height/20, s=z)
    plt.setp(ax.get_xticklabels(), rotation=50, horizontalalignment='right')
    plt.tight_layout()
    plt.savefig('/home/danny/workspace/'+graph_title, dpi=300)
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

    rects1 = ax.bar(ind - width/2, nm, width, label='Nursing & Midwifery')
    rects2 = ax.bar(ind+width/2, md, width, label= 'Medical & Dental')

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
    plt.savefig('/home/danny/workspace/'+graph_title, dpi=300)
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


graph_maker_docs_and_nurses(self_isolating_piv, "Special Leave SP - Coronavirus - Self Isolating - Clinical")
graph_maker_docs_and_nurses(covid_parental_piv, "Special Leave SP - Coronavirus Parental Leave - Clinical")
graph_maker_docs_and_nurses(covid_pos_piv, "Special Leave SP - Coronavirus - Covid-19 Confirmed - Clinical")
graph_maker_all(all_isolating_piv, "Special Leave SP - Coronavirus - Self Isolating - All staff")
graph_maker_all(all_parental_piv, "Special Leave SP - Coronavirus Parental Leave - All staff")
graph_maker_all(all_positive_piv, "Special Leave SP - Coronavirus - Covid-19 Confirmed - All staff")

df_isolating_sheet = all_isolating[['Pay_Number','Supervisor email address','Forename','Surname','Sector/Directorate/HSCP','Sub-Directorate 1','department','Job_Family','Absence Episode Start Date','Address_Line_1','Address_Line_2','Address_Line_3','Postcode', 'Best Phone']]
df_isolating_sheet.to_excel('/media/wdrive/daily_absence/isolators-'+date.today().strftime('%Y-%m-%d')+'.xlsx', index=False)