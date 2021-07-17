from datetime import date
from selenium import webdriver
from bs4 import BeautifulSoup
import getpass
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import Select
import time
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook
wb = Workbook()
ws = wb.active


driver = webdriver.Chrome(executable_path='C:/Users/ALPA GARG/Desktop/scraping 2/chromedriver.exe')
driver.get('https://www.google.com/search?tbs=lf:1,lf_ui:14&tbm=lcl&q=marketing+companies+in+chandigarh&rflfq=1&num=10&sa=X&ved=2ahUKEwjcmJ-ulffwAhUgyzgGHVedDWMQjGp6BAgEEFk&biw=1536&bih=754#rlfi=hd:;si:;mv:[[30.7471298,76.7994542],[30.682378399999997,76.7190936]];tbs:lrf:!1m4!1u3!2m2!3m1!1e1!1m4!1u2!2m2!2m1!1e1!2m1!1e2!2m1!1e3!3sIAE,lf:1,lf_ui:14')
driver.maximize_window()

search_box = driver.find_element_by_name('q')

company_name, company_website,  Address, PhoneNumber,FB,Twitter, Insta, Web1, Web2, Web3, Owner_Name1, Owner_Name2, Owner_Name3= [], [], [], [], [], [], [],[],[],[],[],[],[]

time.sleep(3)
search_result = driver.find_elements_by_class_name('VkpGBb')

time.sleep(3)
#page = driver.find_element_by_xpath('/html/body/div[6]/div/div[8]/div[1]/div/div/div[2]/div[2]/div/div/div/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[4]/a')
#page.click()


for j in range (2):

    time.sleep(7)
    search_result = driver.find_elements_by_class_name('VkpGBb')

    for i in range(len(search_result)):

        #Click on the elements found
        try:
            search_result[i].click()
        except StaleElementReferenceException as exception:
            time.sleep(10)
            driver.find_elements_by_class_name('VkpGBb')
            search_result[i].click()

        #Extract the company name
        time.sleep(4)
        while True:
            try:
                company_name.append(driver.find_element_by_class_name('SPZz6b>h2').text)
                break
            except NoSuchElementException as exception:
                driver.refresh()
                time.sleep(4)
                try:
                    company_name.append(driver.find_element_by_class_name('SPZz6b>h2').text)
                    break
                except NoSuchElementException as exception:
                    print('In loop, plz check')

        #Extract the address of the company
        try:
            Address.append(driver.find_element_by_xpath(
                '/html/body/div[6]/div[2]/div[8]/div[3]/div/div[2]/async-local-kp/div/div/div[1]/div/div/div/div[1]/div/div[1]/div/div[3]/div/div[3]/div/div/span[2]').text)
        except NoSuchElementException as exception:
            Address.append('Not Mentioned')

        #Extract the website of the company
        try:
            company_website.append(driver.find_element_by_class_name('QqG1Sd>a').get_attribute('href'))
        except NoSuchElementException as exception:
            company_website.append('Null')

        #Extract the Phone no.
        try:
            time.sleep(2)
            PhoneNumber.append(driver.find_element_by_link_text('Phone').get_attribute('href'))
        except NoSuchElementException as exception:
            PhoneNumber.append('Null')

        #Extract the Facebook Profile
        try:
            FB.append(driver.find_element_by_link_text('Facebook').get_attribute('href'))
        except NoSuchElementException as exception:
            FB.append('Null')

        #Extract the Twitter Profile
        try:
            Twitter.append(driver.find_element_by_link_text('Twitter').get_attribute('href'))
        except NoSuchElementException as exception:
            Twitter.append('Null')

        #Extract the Instagram Profile
        try:
            Insta.append(driver.find_element_by_link_text('Instagram').get_attribute('href'))
        except NoSuchElementException as exception:
            Insta.append('Null')

        #Extract the first 3 web results
        fields = (driver.find_elements_by_class_name("RFlwHf"))
        time.sleep(3)
        Web1.append(fields[0].get_attribute('href'))
        Web2.append(fields[1].get_attribute('href'))
        try:
            Web3.append(fields[2].get_attribute('href'))
        except (IndexError):
            Web3.append('Null')


    #To go to next pages of the google maps
    time.sleep(3)
    try:
        page = driver.find_elements_by_class_name('d6cvqb>a')
        if(len(page)>1):
            page[1].click()
        else:
            page[0].click()
    except StaleElementReferenceException as exception:
        time.sleep(15)
        page = driver.find_elements_by_class_name('d6cvqb>a')
        if (len(page) > 1):
            page[1].click()
        else:
            page[0].click()


#Serach for the linkedin profile of the owner of different companies and extract the top 3 profiles
driver.get('https://www.google.com/')
time.sleep(3)
linkedin1, linkedin2, linkedin3 = [], [],[]

class TimeoutExcept0ion(object):
    pass
a=0
for i in company_name:
    a=a+1
    search_box = driver.find_element_by_name('q')
    search_box.clear()
    search_box.send_keys(i + ' India CEO/Owner/director linkedin')
    while True:
        try:
            search_box.submit()
            break
        except TimeoutException as exception:
            continue
    time.sleep(10)

    fields = (driver.find_elements_by_xpath('//div[@class="yuRUbf"]/a'))
    time.sleep(2)
    linkedin1.append(fields[0].get_attribute('href'))
    #except (IndexError):
    #link1.append('Null')
    linkedin2.append(fields[1].get_attribute('href'))
    linkedin3.append(fields[2].get_attribute('href'))
    if(a%15== 0):
        time.sleep(180)
    print(a)


#Login the linkedin account
driver.get('https://www.linkedin.com/')
time.sleep(3)
username = driver.find_element_by_name('session_key')
username.send_keys('naina**********@gmail.com')
password = driver.find_element_by_name('session_password')
password.send_keys('**************')
while True:
    try:
        password.submit()
        break
    except TimeoutException as exception:
        driver.refresh()
        time.sleep(3)
        username = driver.find_element_by_name('session_key')
        username.send_keys('naina*******@gmail.com')
        password = driver.find_element_by_name('session_password')
        password.send_keys('**********')
        password.submit()


#Extract the owner name from their linkedin page
for i in range(len(linkedin1)):
    driver.get(linkedin1[i])
    time.sleep(4)
    try:
        Owner_Name1.append(driver.find_element_by_xpath('//li[@class="inline t-24 t-black t-normal break-words"]').text)
    except NoSuchElementException as exception:
        time.sleep(3)
        try:
            Owner_Name1.append(
                driver.find_element_by_xpath('//li[@class="inline t-24 t-black t-normal break-words"]').text)
        except NoSuchElementException as exception:
            Owner_Name1.append('Null')



DATA = []
for i in range(len(company_name)):
    DATA.append([company_name[i], company_website[i],linkedin1[i], linkedin2[i], FB[i] , Twitter[i],Insta[i], Owner_Name1[i],Web1[i],Web2[i], Web3[i], PhoneNumber[i], Address[i], linkedin3[i]])

ws = pd.DataFrame(DATA, columns=['Company Name', 'Website', 'link 1', 'link 2','FB','Twitter','Ínsta', 'Owner name', 'Web 1','Web 2', 'Web 3', 'phone no.', 'address','link 3'])

#saved the data in a filed named Data mined
ws.to_excel('Data mined.xlsx')
