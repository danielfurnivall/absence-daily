'''The aim of this file is to send covid-related emails to several people with relevant data each day'''

from exchangelib import Account, Configuration, Credentials, Mailbox, Message, FileAttachment
import configparser
from datetime import date
import time
import pandas as pd

config = configparser.ConfigParser()
config.read('C:/Tong/creds.ini')
smtp_server = 'mail.ggc.scot.nhs.uk'
uname = config.get('EMAIL', 'uname')
pword = config.get('EMAIL', 'pword')

creds = Credentials(uname, pword)
account = Account('daniel.furnivall@ggc.scot.nhs.uk', credentials=creds, autodiscover=True)


def send_email(to_address, subject, body, attachments):
    m = Message(
        account=account,
        folder=account.sent,
        subject=subject,
        body=body,

        to_recipients=[Mailbox(email_address=to_address)]
    )
    #representing attachments as dictionaries with filenames as key + alias as value
    for i in attachments:
        with open(i, 'rb') as f:
            file_content = f.read()
        m.attach(FileAttachment(name=attachments.get(i), content=file_content))
        m.save()
    m.send_and_save()


date = date.today().strftime('%Y-%m-%d')

# Gillian Gall
send_email('gillian.gall2@ggc.scot.nhs.uk', date, 'This email was sent automatically - if there are any issues, please let me know with an email.',
           {'W:/Daily_Absence/West_Dun-' + date + '.csv': date + '.csv'})

time.sleep(2)

# Morag Kinnear
for i in ['morag.kinnear@ggc.scot.nhs.uk', 'Gillian.Ayling-Whitehouse@ggc.scot.nhs.uk']:
    send_email(i, date, 'This email was sent automatically - if there are any issues, please let me know with an email.',
               {'W:/daily_absence/positive-' + date + '.xlsx': 'positive' + date + '.xlsx'})

time.sleep(2)

# Gillian Ayling Whitehouse & Steven
recips = ['Gillian.Ayling-Whitehouse@ggc.scot.nhs.uk', 'steven.munce@ggc.scot.nhs.uk']
for i in recips:
    send_email(i, date, 'This email was sent automatically - if there are any issues, please let me know with an email.',
               {'W:/daily_absence/new_old_covid-' + date + '.xlsx': 'New_Old_Covid' + date + '.xlsx'})

    time.sleep(2)

# Covid_Absence_Team
covid_team = ['Gillian.Ayling-Whitehouse@ggc.scot.nhs.uk', 'Colin.McGowan@ggc.scot.nhs.uk',
              'James.Farrelly@ggc.scot.nhs.uk', 'Tracy.Keenan2@ggc.scot.nhs.uk', 'Morag.Kinnear@ggc.scot.nhs.uk',
              'Karleen.Jackson@ggc.scot.nhs.uk', 'Alexsis.Boffey@ggc.scot.nhs.uk', 'David.Dall@ggc.scot.nhs.uk',
              'Joan.Smith@ggc.scot.nhs.uk', 'Stephen.Wallace@ggc.scot.nhs.uk', 'Steven.Munce@ggc.scot.nhs.uk',
              'Ruth.Campbell@ggc.scot.nhs.uk', 'Margaret.Pirie@ggc.scot.nhs.uk', 'Margaret.Glen@ggc.scot.nhs.uk']
for i in covid_team:
    send_email(i, date, 'This email was sent automatically - if there are any issues, please let me know with an email.'
                        '\n\nNew field added (11/12/20) - "Booked Absence Days" - this allows you to look more closely at '
                        'absences within the self-isolating and household isolating categories now the isolation period '
                        'has been reduced to 10 days.',
           {'W:/daily_absence/all_covid_absence-' + date + '.xlsx': 'all_covid_absence' + date + '.xlsx'})
    time.sleep(2)

# tracey, steven, nareen
main_email_recipients = ['tracey.carrey@ggc.scot.nhs.uk', 'steven.munce@ggc.scot.nhs.uk','lauren.kelly@ggc.scot.nhs.uk',
                         'ian.macrae@ggc.scot.nhs.uk']
# read the data file produced by fast_graphs.py
f = open('W:/daily_absence/raw_data' + date + '.txt', 'r')
body = f.read()

for i in main_email_recipients:
    send_email(i, date, body,
           {'W:/daily_absence/positive-' + date + '.xlsx': 'positive' + date + '.xlsx',
            'W:/daily_absence/isolators-' + date + '.xlsx': 'isolators' + date + '.xlsx',
            'W:/daily_absence/all_covid' + date + '.xlsx': 'all_covid_pivot' + date + '.xlsx',
            'C:/Covid_Graphs/'+'All Covid-related Absence Reasons.png': 'All Covid-related Absence Reasons.png',
            'C:/Covid_Graphs/'+'Special Leave SP - Coronavirus - Covid-19 Confirmed - All staff.png':
                'Special Leave SP - Coronavirus - Covid-19 Confirmed - All staff.png',
            'C:/Covid_Graphs/'+'Special Leave SP - Coronavirus - Quarantine (new code).png':
                'Special Leave SP - Coronavirus - Quarantine (new code).png',
            'C:/Covid_Graphs/'+'Special Leave SP - Coronavirus Parental Leave - All staff.png':
                'Special Leave SP - Coronavirus Parental Leave - All staff.png',
            'C:/Covid_Graphs/'+'Special Leave SP - Coronavirus – Household Related – Self Isolating - All staff.png':
                'Special Leave SP - Coronavirus – Household Related – Self Isolating - All staff.png',
            'C:/Covid_Graphs/'+'Special Leave SP - Coronavirus – Self displaying symptoms – Self Isolating - All Staff.png':
                'Special Leave SP - Coronavirus – Self displaying symptoms – Self Isolating - All Staff.png',
            'C:/Covid_Graphs/'+'Special Leave SP - Coronavirus – Test and Protect Isolation.png':
                'Special Leave SP - Coronavirus – Test and Protect Isolation.png',
            'C:/Covid_Graphs/'+'Special Leave SP - Coronavirus – Underlying Health Condition - All staff.png':
                'Special Leave SP - Coronavirus – Underlying Health Condition - All staff.png'
            })
    time.sleep(2)
