from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import ElementClickInterceptedException, ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import configparser
from selenium.webdriver import ActionChains
import datetime
import os
import shutil
import time
import pandas as pd

config = configparser.ConfigParser()
config.read('/home/danny/creds.ini')
fp = webdriver.FirefoxProfile()
fp.set_preference("browser.download.dir", '/media/wdrive/Daily_Absence/')
fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")
driver = webdriver.Firefox(executable_path='/home/danny/workspace/geckodriver', firefox_profile=fp)
driver.get("https://bo-wf.scot.nhs.uk/InfoViewApp/logon.jsp")


def login():
    username_id = "usernameTextEdit"
    password_id = "passwordTextEdit"
    button_xpath = '//*[@id="buttonTable"]/input'

    driver.switch_to.frame('infoView_home')
    WebDriverWait(driver, 90).until(
        ec.element_to_be_clickable((By.ID, username_id)))

    username_field = driver.find_element_by_id(username_id)
    password_field = driver.find_element_by_id(password_id)

    password_field.send_keys(config.get('SSTS', 'password'))
    username_field.send_keys(config.get('SSTS', 'username'))

    driver.find_element_by_xpath(button_xpath).click()

def boxi_iframe_switch():

    driver.switch_to.default_content()
    WebDriverWait(driver, 90).until(
        ec.frame_to_be_available_and_switch_to_it('headerPlusFrame'))
    WebDriverWait(driver, 90).until(
        ec.frame_to_be_available_and_switch_to_it('dataFrame'))
    WebDriverWait(driver, 90).until(
        ec.frame_to_be_available_and_switch_to_it('workspaceFrame'))
    WebDriverWait(driver, 90).until(
        ec.frame_to_be_available_and_switch_to_it('workspaceBodyFrame'))


def get_to_report():
    boxi_iframe_switch()
    print("hello")

    manager_button = 'ListingURE_treeNode7_name'
    report_selector = 'ListingURE_listColumn_1_0_1'
    staff_record_leave_6a = 'ListingURE_listColumn_28_0_1'
    WebDriverWait(driver, 90).until(
        ec.element_to_be_clickable((By.ID,manager_button)))
    driver.find_element_by_id(manager_button).click()

    chain = ActionChains(driver)
    chain.double_click(driver.find_element_by_id(report_selector)).perform()
    WebDriverWait(driver, 90).until(
        ec.element_to_be_clickable((By.ID, staff_record_leave_6a))
    )
    driver.execute_script("arguments[0].scrollIntoView();", driver.find_element_by_id(staff_record_leave_6a))
    chain = ActionChains(driver)
    chain.double_click(driver.find_element_by_id(staff_record_leave_6a)).perform()



def sickabs():
    WebDriverWait(driver, 10).until(
        ec.frame_to_be_available_and_switch_to_it('webiViewFrame'))
    yesterday = datetime.date.today()
    WebDriverWait(driver, 10).until(
        ec.element_to_be_clickable((By.ID, 'PV1')))
    start = driver.find_element_by_id('PV1')
    start.clear()
    start.send_keys(yesterday.strftime('%d/%m/%Y') + " 00:00:00")
    WebDriverWait(driver, 10).until(
        ec.element_to_be_clickable((By.ID, '_CWpromptstrLstElt1')))
    driver.find_element_by_id('_CWpromptstrLstElt1').click()
    end = driver.find_element_by_id('PV2')
    end.clear()
    end.send_keys(yesterday.strftime('%d/%m/%Y') + " 00:00:00")
    WebDriverWait(driver, 10).until(
        ec.element_to_be_clickable((By.ID, '_CWpromptstrLstElt2')))
    driver.find_element_by_id('_CWpromptstrLstElt2').click()
    firstelem = driver.find_element_by_xpath('//*[@id="mlst_bodyLPV3_lov"]/div/table/tbody/tr[1]/td/div')
    chain = ActionChains(driver)
    chain.key_down(Keys.SHIFT).click(firstelem).send_keys(Keys.END).key_up(Keys.SHIFT).perform()
    driver.find_element_by_id('theBttnIconPV3AddButton').click()
    driver.find_element_by_id('theBttnCenterImgpromptsOKButton').click()

def get_data():

    WebDriverWait(driver, 90).until(ec.invisibility_of_element_located((By.ID, 'modal_waitDlg')))
    # WebDriverWait(driver, 10).until(
    #     ec.frame_to_be_available_and_switch_to_it('_iframeleftPane'))
    # driver.find_element_by_xpath('/html/body/div/span/div[2]/div[3]/span').click()
    # time.sleep(0.5)
    # WebDriverWait(driver, 10).until(
    #     ec.frame_to_be_available_and_switch_to_it('webiViewFrame'))
    driver.find_element_by_id('IconImg_iconMenu_arrow_docMenu').click()
    hov = driver.find_element_by_id('iconMenu_menu_docMenu_span_text_saveReportComputerAs')
    ActionChains(driver).move_to_element(hov).perform()
    WebDriverWait(driver, 90).until(
        ec.element_to_be_clickable((By.ID, 'saveReportComputerAs_span_text_saveReportXLS'))).click()
    try:
        driver.find_element_by_id('saveReportComputerAs_span_text_saveReportXLS').click()
    except ElementNotInteractableException:
        while not os.path.exists('/home/danny/Downloads/6a. Staff Record Leave - You Choose Type(s) - All Locations.xls'):
            time.sleep(1)
    yesterday = str(datetime.date.today())
    # newpath = '/media/wdrive/Daily_Absence/' + yesterday + '.xls'
    # oldpath = '/home/danny/Downloads/6a. Staff Record Leave - You Choose Type(s) - All Locations.xls'
    # try:
    #     shutil.move(oldpath, newpath)
    # except PermissionError:
    #     print("Error Quashed.")
    # while not os.path.exists(newpath):
    #     time.sleep(1)
    # os.remove('/home/danny/Downloads/6a. Staff Record Leave - You Choose Type(s) - All Locations.xls')


def transform_data():
    yesterday = str(datetime.date.today())
    newpath = '/media/wdrive/Daily_Absence/' + yesterday + '.xls'
    oldpath = '/home/danny/Downloads/Marion-Absence.xls'
    df = pd.read_excel(newpath, skiprows=1)
    sd = pd.read_excel('media/wdrive/Workforce Monthly Reports/Monthly_Reports/Feb-20 Snapshot/Staff Download/2019-02 - Staff Download - GGC.xls')
    print(df.columns)
    print(sd.columns)

    # try:
    #     shutil.move(oldpath, newpath)
    # except PermissionError:
    #     print("Error Quashed.")
    # while not os.path.exists(newpath):
    #     time.sleep(1)
    # os.remove('/home/danny/Downloads/Marion-Absence.xls')


login()
while True:
    try:
        get_to_report()
        sickabs()
        get_data()
        #transform_data()
        break
    except ElementClickInterceptedException:
        print('Broken, retrying')
        time.sleep(10)
        driver.get('https://bo-wf.scot.nhs.uk/InfoViewApp/listing/main.do')
#driver.close()

#transform_data()