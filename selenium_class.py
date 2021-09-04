
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
import chromedriver_autoinstaller
 
import os
import glob
import shutil
import time
import configparser
import win32com.client
 
chromedriver_autoinstaller.install()
 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = outlook.Folders
user_email = str(accounts[0])
userid = user_email.split('@')[0]
print('*--------------------------------------*')
print(' User email: ', user_email, '\n UserID: ', userid, sep='', end='\n')
print('*--------------------------------------*')
 
def rm_dir(dir):
    try:
        shutil.rmtree(dir)
    except:
        pass 
 
def setting_driver(URL):
    options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : tmp_download_dir }
    options.add_experimental_option('prefs',prefs)
    driver = webdriver.Chrome(chrome_options = options)
    driver.get(URL)
    driver.maximize_window()
    time.sleep(5)
    try:
        driver.find_element_by_id('i0116').send_keys(user_email)
        driver.find_element_by_id('idSIButton9').click()
    except:
        print('Error')
    time.sleep(15)
    return driver
 
def Web_operation(Series,DID):
    dropdown = driver.find_element_by_id('inputSeries')
    select = Select(dropdown)
    select.select_by_visible_text(Series)
 
    time.sleep(3)
    dropdown = driver.find_element_by_id('inputDesignID')
    select = Select(dropdown)
    select.select_by_visible_text(DID)
 
    time.sleep(2)
    elem= driver.find_element_by_xpath("//ul[@class='nav nav-pills']")
    elem.find_element_by_xpath(".//a[@class='nav-link']").click()
 
    time.sleep(2)
    Search_btn = driver.find_element_by_xpath("(//button[@tooltip='Search'])")
    Search_btn.click()
 
    time.sleep(15)
    CSV_Export_btn = driver.find_element_by_xpath("(//button[@tooltip='CSV Export'])")
    CSV_Export_btn.click()


    time.sleep(5)
    driver.find_element_by_xpath("(//input[@id='dpwweeks'])").send_keys(Keys.BACK_SPACE)
    driver.find_element_by_xpath("(//input[@id='dpwweeks'])").send_keys("20")
 
    time.sleep(5)
    CSV_Export2_btn = driver.find_element_by_xpath("(//button[@class='btn btn-primary ng-star-inserted'])")
    CSV_Export2_btn.click()
    time.sleep(30)
    driver.quit()
 
    Download_FileName = glob.glob(f'{tmp_download_dir}\\*.csv')
    OutPutDIR = 'H:\MMJ\Secure\PIE\All-Write\YE\Module\YLR\YIP_Download'
    for file in Download_FileName:
        try:
            OutFileName = OutPutDIR + "/" + DID + "_" + file.split('\\')[-1][0:-12] + "_dieloss.csv"       
            shutil.copy(file, OutFileName)
            print('-'*100)
            print(' Correctly Done: ', OutFileName, sep='', end='\n')
            print('-'*100)
        except FileNotFoundError:
            pass
        except OSError:
            pass
    try:
        shutil.rmtree(tmp_download_dir)
    except:
        pass
 
if __name__ == "__main__":
    os.chdir('C:/Users/' + userid + '/Downloads')
    current_dir = os.getcwd()
    tmp_download_dir = f'{current_dir}\\tmpDownload'
    rm_dir(tmp_download_dir)
 
    DIDs = ["A"]

    URL = 'https://***********'
    for DID in DIDs:
        driver = setting_driver(URL)
        Web_operation('120s', DID)
 
    DIDs = ["A"]
    URL = 'https://***********'
    for DID in DIDs:
        driver = setting_driver(URL)
        Web_operation('130s', DID)
