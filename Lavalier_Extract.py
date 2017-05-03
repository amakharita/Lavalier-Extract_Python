import glob
import os
import datetime
import time
import shutil

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import ConfigParser

# Config file
config = ConfigParser.ConfigParser()
config.readfp(open('config.ini'))
username = config.get('login', 'username')
password = config.get('login', 'password')

# File path
download_dir = 'S:\Lavalier_Report_Downloads'

# Selenium
chromeoptions = webdriver.ChromeOptions()
options = {'download.default_directory': download_dir}
chromeoptions.add_experimental_option('prefs', options)
browser = webdriver.Chrome('C:/chromedriver.exe', chrome_options=chromeoptions)

# LOGIN - selects login form, populates user and pass fields, and submits form
def lavalierLogin():
    browser.get('https://lavalier.dragonflyware.com/login/login')
    browser.find_element_by_class_name('form-signin')
    element = browser.find_element_by_name('username')
    element.send_keys(str(username))
    element = browser.find_element_by_name('password')
    element.send_keys(str(password))
    browser.find_element_by_class_name('btn-lg').submit()

# Get first day of previous month
def prev_month_start(when):
    # Return previuos month's start and end date
    when = datetime.datetime.today()
    # Find today
    first = datetime.date(day=1, month=when.month, year=when.year)
    # Use that to find the first day of this month
    prev_month_end = first - datetime.timedelta(days=1)
    prev_month_start = datetime.date(day=1, month= prev_month_end.month, year= prev_month_end.year)
    # Return previous month's start date in MM/DD/YYYY format (must be in this format for the form to run)
    start = prev_month_start.strftime('%m/%d/%Y')
    return start

# Get last day of previous month
def prev_month_end(when):
    # Return previous month's start and end date
    when = datetime.datetime.today()
    # Find today
    first = datetime.date(day=1, month=when.month, year=when.year)
    # Use that to find the first day of this month
    prev_month_end = first - datetime.timedelta(days=1)
    prev_month_start = datetime.date(day=1, month= prev_month_end.month, year= prev_month_end.year)
    # Return previous month's end date in MM/DD/YYYY format (must be in this format for the form to run)
    end = prev_month_end.strftime('%m/%d/%Y')
    return end

# Gets Policy Transaction XLS file
def getPolicyTransactions():
    browser.get('https://lavalier.dragonflyware.com/policy/app/reports?execution=e3s1')
    browser.find_element_by_xpath("//select[@name='form1:j_idt35']/option[text()='Policy Transactions']").click()
    # Calendar fields
    startdate = browser.find_element_by_id('form1:dateField3_input')
    ActionChains(browser).move_to_element(startdate).click().send_keys(prev_month_start(datetime.datetime.now())).perform()
    # search_btn = browser.find_element_by_id('ctl00_cphMain_btnSearchAll')
    enddate = browser.find_element_by_id('form1:dateField4_input')
    ActionChains(browser).move_to_element(enddate).click().click().perform()
    ActionChains(browser).move_to_element(enddate).click().send_keys(prev_month_end(datetime.datetime.now())).perform()
    browser.find_element_by_class_name('commandLink').click()
    # Wait 30 seconds for the BtnGenerateReport to return, by then the file will have downloaded.
    max_wait_in_seconds = 30
    WebDriverWait(browser, max_wait_in_seconds).until(EC.presence_of_element_located((By.CLASS_NAME, 'commandLink')))

# Gets Bordereau Item XLS file
def getBordereauItem():
    browser.get('https://lavalier.dragonflyware.com/policy/app/reports?execution=e3s1')
    browser.find_element_by_xpath("//select[@name='form1:j_idt35']/option[text()='Bordereau Policy']").click()
    # Calendar fields
    browser.find_element_by_id('form1:dateField3_input').clear()
    startdate = browser.find_element_by_id('form1:dateField3_input')
    ActionChains(browser).move_to_element(startdate).click().send_keys(prev_month_start(datetime.datetime.now())).perform()
    browser.find_element_by_id('form1:dateField4_input').clear()
    enddate = browser.find_element_by_id('form1:dateField4_input')
    ActionChains(browser).move_to_element(enddate).click().click().perform()
    ActionChains(browser).move_to_element(enddate).click().send_keys(prev_month_end(datetime.datetime.now())).perform()
    browser.find_element_by_class_name('commandLink').click()
    # Wait 30 seconds for the BtnGenerateReport to return, by then the file will have downloaded.
    max_wait_in_seconds = 30
    WebDriverWait(browser, max_wait_in_seconds).until(EC.presence_of_element_located((By.CLASS_NAME, 'commandLink')))

# Move Bordereau and Transaction files from temp folder to their respective folders
def moveFiles():
    # Temp folder
    init_dir = 'S:/Lavalier_Report_Downloads/'
     # Moves files to folders that the rest of the ETL process will use as a source
    for f in os.listdir(init_dir):
        if f == 'policytransactions.xls':
            shutil.move('S:/Lavalier_Report_Downloads/policytransactions.xls',
                    'S:/Lavalier_Report_Downloads/Policy_Transactions/policytransactions.xls')
        elif f == 'LavalierBordereauxPolicy.xls':
            shutil.move('S:/Lavalier_Report_Downloads/LavalierBordereauxPolicy.xls',
                    'S:/Lavalier_Report_Downloads/Bordereau_Policy/LavalierBordereauxPolicy.xls')
        #else:
            #send error email function here

# Rename files to something identifiable
def renameFiles():
    pt_dir = 'S:/Lavalier_Report_Downloads/Policy_Transactions/'
    bord_dir = 'S:/Lavalier_Report_Downloads/Bordereau_Policy/'
    start = prev_month_start(datetime.datetime.now())
    end = prev_month_end(datetime.datetime.now())
    # Replace slashes as it confuses the rename function
    rename_pt_to = 'PT' + '-' + start.replace('/', '') + '-' + end.replace('/', '') + '.xls'
    rename_bord_to = 'BORD' + '-' + start.replace('/', '') + '-' + end.replace('/', '') + '.xls'
    for f in os.listdir(pt_dir):
        if f == 'policytransactions.xls':
            os.rename(os.path.join(pt_dir, f), os.path.join(pt_dir, rename_pt_to))
    for f in os.listdir(bord_dir):
        if f == 'LavalierBordereauxPolicy.xls':
            os.rename(os.path.join(bord_dir, f), os.path.join(bord_dir, rename_bord_to))



lavalierLogin()
getPolicyTransactions()
getBordereauItem()
moveFiles()
renameFiles()

for filename in glob.glob(os.path.join(download_dir, '*.xls*')):
    print(filename)



