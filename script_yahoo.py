"""
Created on Sat Jan 16 21:36:44 2021

@author: youssef ouahman
"""

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import Workbook
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import requests
import pyautogui
#from bs4 import BeautifulSoup



#driver = webdriver.Chrome(ChromeDriverManager().install())
driver = webdriver.Firefox()
#driver = webdriver.Chrome("C:/chromedriver.exe")

def yaloo(driver):
    workbook = Workbook()
    screenWidth, screenHeight = 1790,664
    driver.maximize_window()
    place = input("entry name thr city:")
    start = input("entry number the first  page:")
    end = input("entry number the finish page:")
    total = [["company","addresse","phone"]]

    for i in range(int(start),int(end)):
            for j in range(5,38):
                peice =[]
                try:
                    driver.get("https://www.yelo.ma/location/{}/{}".format(place,i))
                    driver.implicitly_wait(10)
                    ser = driver.find_element(By.XPATH,"/html/body/main/section/div[4]/div[{}]/h4/a".format(j))
                    driver.implicitly_wait(5)
                    ser.click()
                    time.sleep(5)
                    time.sleep(5)
                    pyautogui.click(screenWidth, screenHeight )
                    company = driver.find_element_by_xpath("//*[@id='company_name']")
                    addresse =  driver.find_element_by_xpath("/html/body/main/section/div[3]/div[1]/div[2]/div/div[3]/div[2]")
                    phone =  driver.find_element_by_xpath("/html/body/main/section/div[3]/div[1]/div[2]/div/div[4]/div[2]")
                    peice.append(company.text)
                    peice.append(addresse.text)
                    peice.append(phone.text)
                    total.append(peice)
                except:
                    print("probleme")
                    pass    

            sheet = workbook.active    
            for row in total:
                 sheet.append(row)
            workbook.save(filename="data_{}_{}_{}.xlsx".format(place,start,end))




yaloo(driver)
