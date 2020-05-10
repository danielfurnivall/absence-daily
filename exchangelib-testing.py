'''This script is for testing ExchangeLib and how it functions under forticlient VPN'''

from exchangelib import Credentials, Account, Message, Mailbox, HTMLBody
import configparser
config = configparser.ConfigParser()
config.read('W:/Python/Danny/SSTS Extract/SSTSconf.ini')
uname = config.get('EXCHANGE', 'username')
pword = config.get('EXCHANGE', 'password')
print(pword)
credentials = Credentials(uname, pword)
account = Account('flexrequests@ggc.scot.nhs.uk', credentials=credentials, autodiscover=True)
m = Message(
                account=account,
                subject='Your flexible working request has been submitted.',
                body=HTMLBody('Hello.'),
                to_recipients=([Mailbox(email_address='daniel.furnivall@ggc.scot.nhs.uk')])
)

m.send()
