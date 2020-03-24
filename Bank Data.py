from selenium import webdriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
import configparser
import pandas as pd
import os


config = configparser.ConfigParser()
config.read('/media/wdrive/Python/Danny/SSTS Extract/SSTSConf.ini')
date = pd.datetime.now()
start_date = date - timedelta(days=7)
end_date = date + timedelta(days=7)
filename = '/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/Export.xls'
fp = webdriver.FirefoxProfile()
fp.set_preference("browser.download.folderList", 2)
fp.set_preference("browser.download.dir", '/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/')
fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")
browser = webdriver.Firefox(executable_path='/home/danny/workspace/geckodriver', firefox_profile=fp)


def login():
    browser.get('https://nww.ggcbank.allocate-cloud.com/BankStaff/(S(wqc1vwycovtfhhwo0d00v4ox))/UserLogin.aspx')
    WebDriverWait(browser, 90).until(
            ec.element_to_be_clickable(
                (By.ID, 'ctl00_content_login_UserName')))
    username = browser.find_element_by_id("ctl00_content_login_UserName")
    username.clear()
    username.send_keys(config.get('ALLOCATE','uname'))
    password = browser.find_element_by_id("ctl00_content_login_Password")
    password.clear()
    password.send_keys(config.get('ALLOCATE','pword'))
    browser.find_element_by_id("ctl00_content_login_LoginButton").click()

def puller(start_date, end_date):
    #browser.get('https://nww.ggcbank.allocate-cloud.com/BankStaff/(S(ac5tgl1fsh4jtkaijycwkrgx))/BookingsStatus.aspx?&navigation=BookingsStatus&TopNavigation=Requests&Action=All')
    browser.find_element_by_id('ctl00_topNavigation_TopNav_linkRequests').click()
    WebDriverWait(browser, 90).until(
            ec.presence_of_element_located(
                (By.ID, 'ctl00_content_BookingStatus1_cmdFavorite')))
    browser.find_element_by_id("ctl00_content_BookingStatus1_cmdFavorite").click()

    WebDriverWait(browser, 90).until(
            ec.element_to_be_clickable(
                (By.XPATH, ".//a[contains(@onclick, 'favouriteSelect_1943')]")))

    browser.find_element_by_xpath(".//a[contains(@onclick, 'favouriteSelect_1943')]").click()

    time.sleep(5)
    browser.find_element_by_id("ctl00_content_BookingStatus1_collapsibleImage").click()
    WebDriverWait(browser, 90).until(
            ec.element_to_be_clickable(
                (By.ID, "ctl00_content_BookingStatus1_StartDate")))

    startdate = browser.find_element_by_id("ctl00_content_BookingStatus1_StartDate")

    enddate = browser.find_element_by_id("ctl00_content_BookingStatus1_EndDate")

    startdate.clear()
    enddate.clear()
    startdate.send_keys(start_date.strftime('%d-%m-%y'))
    enddate.send_keys(end_date.strftime('%d-%m-%y'))

    browser.find_element_by_id("ctl00_content_BookingStatus1_cmdSubmitPrint").click()
    WebDriverWait(browser, 90).until(
            ec.element_to_be_clickable(
                (By.ID, "ctl00_content_BookingStatus1_cmdXLS")))
    browser.find_element_by_id("ctl00_content_BookingStatus1_cmdXLS").click()
    WebDriverWait(browser, 90).until(
            ec.element_to_be_clickable(
                (By.ID, "ctl00_content_ExportToXLS")))
    browser.find_element_by_id("ctl00_content_ExportToXLS").click()
    to_csv()

def to_csv():

    while not os.path.exists(filename):
        time.sleep(2)
        # if len(browser.window_handles) == 2:
        #     time.sleep(5)
        #     if len(browser.window_handles) == 2:
        #         browser.switch_to.window(browser.window_handles[1])
        #         browser.close()
        #         browser.switch_to.window(browser.window_handles[0])
        #         puller(start_date, end_date)
    if os.path.getsize(filename) < 4096:
        puller(start_date, end_date)
        os.remove(filename)
    if os.path.exists(filename):
        print("Extract Complete", "Extract complete - translating to xlsx")
        data = pd.read_html(filename,
                            converters={'Request Id': lambda x: f"{x:10}"})
        data = data[0]
        data['Date'] = pd.to_datetime(data['Date'], format="%d-%b-%Y")
        data.to_csv('/media/wdrive/Coronavirus Daily Absence/MICROSTRATEGY/'+start_date.strftime('%Y%m%d')+'-'+end_date.strftime('%Y%m%d')+'.csv', index=False, date_format='%d-%b-%Y')
        os.remove(filename)


login()
puller(start_date, end_date)
to_csv()