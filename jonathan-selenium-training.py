from selenium import webdriver
import configparser

chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory": r"W:\Python\Danny\bank_extract_folder",
         'safebrowsing.disable_download_protection': True}
chromeOptions.add_experimental_option("prefs", prefs)
initial_site = 'https://nww.ggcbank.allocate-cloud.com/BankStaff/(S(wqc1vwycovtfhhwo0d00v4ox))/UserLogin.aspx'

browser = webdriver.Chrome(executable_path="W:/Danny/Chrome Webdriver/chromedriver.exe",
                           options=chromeOptions)

# uname = input("What is your username?")
# pword = input("What is your password?")


config = configparser.ConfigParser()
config.read(r'W:\\Python\Danny\SSTS Extract\SSTSConf.ini')

def login():
    browser.get(initial_site)
    username = browser.find_element_by_id("ctl00_content_login_UserName")
    username.clear()
    username.send_keys(config.get('ALLOCATE','uname'))
    password = browser.find_element_by_id("ctl00_content_login_Password")
    password.send_keys(config.get('ALLOCATE','pword'))
    browser.find_element_by_id("ctl00_content_login_LoginButton").click()


def get_to_file():
    browser.find_element_by_id('ctl00_content_BookingStatus1_cmdFavorite').click()

login()