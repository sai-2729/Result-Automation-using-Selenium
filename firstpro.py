from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import time

browser = webdriver.Chrome(executable_path='C:\webdrivers\chromedriver.exe')

hallticketno="179F1A0520"
for i in range(0,4):
    num=hallticketno[7:]
    num=int(num)+1
    hallticketno=hallticketno[0:len(hallticketno)-3]+str(num)
    print(hallticketno)
    
    browser.get('https://jntuaresults.ac.in/view-results-56736237.html')
    search=browser.find_element_by_xpath('/html/body/div/div[1]/div/div/center/table/tbody/tr/th/center/input[2]')

    hallticket=browser.find_element_by_id('ht')
    hallticket.send_keys(hallticketno)
    search.click()

    try:
        browser.implicitly_wait(10)
        browser.find_element_by_xpath('/html/body/div/div[1]/div/div/center/div[1]/table')
    except NoSuchElementException:
        print('Fail')
        continue
    time.sleep(1)
    print('pass')
