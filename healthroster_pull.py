from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import ElementClickInterceptedException, ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import configparser
from selenium.webdriver import ActionChains
import datetime
import time


config = configparser.ConfigParser()
config.read('W:/Python/Danny/SSTS Extract/SSTSconf.ini')
# fp = webdriver.FirefoxProfile()
# fp.set_preference("plugin.state.flash", 2)
# fp.set_preference("browser.download.dir", r"W:\Coronavirus Daily Absence\MICROSTRATEGY\\")
# fp.set_preference('dom.ipc.plugins.enabled.libflashplayer.so','true')
# fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")
# driver = webdriver.Firefox(executable_path='W:/MFT/geckodriver', firefox_profile=fp)

chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory": r"W:\Python\Danny\bank_extract_folder",
         'safebrowsing.disable_download_protection': True,
         "profile.default_content_setting_values.plugins": 1,
         "profile.content_settings.plugin_whitelist.adobe-flash-player": 1,
         "profile.content_settings.exceptions.plugins.*,*.per_resource.adobe-flash-player": 1,
         "PluginsAllowedForUrls": "https://nww.ggc.allocate-cloud.com"}
chromeOptions.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(executable_path="W:/Danny/Chrome Webdriver/chromedriver.exe",
                           options=chromeOptions)
driver.get("https://nww.ggc.allocate-cloud.com/HealthRoster/GGCLIVE/Login.aspx")


passwordbox = driver.find_element_by_id('txtPassword')
usernamebox = driver.find_element_by_id('txtUsername')


usernamebox.send_keys(config.get('HealthRoster', 'uname'))
passwordbox.send_keys(config.get('HealthRoster', 'pword'))
driver.find_element_by_id('btnLogin').click()

time.sleep(3)
windows = (driver.window_handles)
driver.switch_to.window(windows[1])