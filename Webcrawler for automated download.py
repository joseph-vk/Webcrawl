import os
import sys
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as expectedConditions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException

###############################################
username = 'username'
password = 'password'

###############################################

def chromeOptions(show = False):
    chromeOptions = webdriver.ChromeOptions()
    path = str(input(r'Enter path name to download reports to :'))
    if os.path.isdir(path) == True:
        prefs = {"download.default_directory" : path}
    else:
        print(r'Path not found...\n  Using default path C:\Users\User\Documents\Location')
        path = 'C:\\Users\\User\\Documents\\Location'
        prefs = {"download.default_directory" : path}

    chromeOptions.add_experimental_option("prefs",prefs)
    # chromeOptions.add_argument("--window-size=10,10")
    chromeOptions.add_experimental_option("prefs", {
                                                        "download.default_directory": path,
                                                      "download.prompt_for_download": False,
                                                       'safebrowsing.enabled': False,
                                                    })
    if show == False:
        chromeOptions.add_argument("--headless")
        chromeOptions.add_argument("--window-size=1920,1080")
        
    driver = webdriver.Chrome(options=chromeOptions)
    driver.set_window_position(-2500,-1)
    return driver, path

#################################################

def waitingTime(element, time=10, ping=2, clickable = False):
    waitingTime = wait(driver,time,ping)
    waitingTime.until(expectedConditions.presence_of_element_located((By.ID,element)))
    if clickable == True:
        waitingTime.until(expectedConditions.element_to_be_clickable((By.ID,element)))

##################################################


def downloadReport(reportLink): # temporarily suspende
    driver.get(reportLink)
    header = driver.find_element_by_xpath("//*[contains(@id,'hdr_')]/th[1]")
    Rightclick = webdriver.ActionChains(driver)
    Rightclick.context_click(header).perform()
    excelbefore = driver.find_element_by_xpath("//*[contains(@id,'context_list_header')]/div[last()]").click()
    sleep(2)
    try:
        excel = driver.find_element_by_xpath("//*[contains(text(),'Excel')]").click()
    except:
        excel = driver.find_element_by_xpath("//*[contains(@class,'context_menu')][2]/div[2]")
    print('File Found')
    #### Number of rows >10,000 check
    numRows = driver.find_element_by_xpath("//*[contains(@id,'_total_rows')]")
    print('There are ',numRows.text,' rows available to download... ')
    print('Starting download... Pease wait')
    if (int(numRows.text) >= 10000):
        element = 'export_wait'
        waitingTime(element)
        waitforDownload = driver.find_element_by_xpath("//*[@id='export_wait']").click()
        sleep(1)
    else:
        element  = 'poll_text'
        waitingTime(element)
        progressText = driver.find_element_by_xpath("//*[@id='poll_text']")
        print(progressText.text)

        while progressText.text != '':
            sleep(10)
            print(progressText.text)  
    sleep(2)
    download = driver.find_element_by_xpath("//*[contains(text(),'Download')]")
    # download = driver.find_element_by_xpath("//*[id='download_button']")
    # driver.execute_script('arguments[0].click();',download)     # backup _ run using javascript
    download.click()
    sleep(5)
    print('File downloaded to :',path)

##############################################


def login(driver,username,password):
    user = driver.find_element_by_id('username')
    passwd = driver.find_element_by_id('password')
    login = driver.find_element_by_xpath("//button[@type='submit']")
    user.send_keys(username)
    passwd.send_keys(password)
    login.click()
    print('\nLogged in successfully')


################################################


def listReports():
    reportDict = {}
    counter = 0
    flag = True
    element = '1166104fc611227b00bb1bc49f4845f4'   # View/Run - unlikely to change 
    try:
        waitingTime(element)  # network connections can result in strange delays
    except NoSuchElementException:
        print('Connection too slow... Retrying to attempt connection')
        waitingTime(element)
        
    viewrun = driver.find_element_by_xpath("//*[@id='1166104fc611227b00bb1bc49f4845f4']/div/div").click()
    element = 'gsft_main'
    waitingTime(element)
    print('Finding reports...\n')
    iframe = driver.find_element_by_xpath("//iframe[@id='gsft_main']")
    driver.switch_to.frame(iframe)
    while (not reportDict) and (counter <3):    # Attempting to find list of reports atleast 3 times
        sleep(1)
        report =driver.find_elements_by_xpath("//a[@ng-bind-html='report.name']")
        for reportNames in report: 
            reportDict[reportNames.text] = reportNames.get_property('href')
        # print(reportDict)
        counter += 1
    if not reportDict:
        print('Unknown error!! No reports found in the location.\n Please re run the program')
        flag = False
    else:
        print('Report Location found \n')
    # print(reportDict, flag)
    return (reportDict,flag)

##############################################



def getReportName(report):
    print('\nEnter the report number you wish to run (1,2,3 etc..)')
    try:
        reportnum = int(input('Enter the 0 to Quit:'))
    except:
        print('Invalid Entry')
        reportnum = len(report)+1
    finally:
        if reportnum == 0 : raise SystemExit('Closing Connection')
        else:    
            while reportnum > len(report):
                try:
                    reportnum = int(input('Report not found, please enter valid number or 0 to quit:'))
                except:
                    continue
                if reportnum == 0 : raise SystemExit('Closing Connection')
    reportname = list(report.keys())[reportnum-1]
    return reportname   



###############################################


(driver,path) = chromeOptions(show=True)
driver.get('https://your/service/url.com')
print('Attempting login')
for i in range(5):
	sleep(0.5)
	print(".",end='', flush = True) 
login(driver,username,password)
(report,reportFlag) = listReports()
if reportFlag == False: 
    driver.close()
    raise SystemExit('Report not found .... Quiting Application')
# getting report number and finding link
for key,reportnames in enumerate(report.keys()):
    print(key+1, ' : ', reportnames)
reportLink = report[getReportName(report)]
downloadReport(reportLink)
driver.close()




