from cmath import exp
from lib2to3.pgen2.driver import Driver
from time import time
from tkinter.messagebox import NO
from turtle import delay
from xml.dom import DOMSTRING_SIZE_ERR
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait 
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
import os
import sys
import shutil
import re
from selenium.webdriver.support import expected_conditions as EC
import time
from bipdict import subjects





class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


todays_fa = []

wb = openpyxl.load_workbook('fa.xlsx')


sheet = wb.active

# get the values from excel 
for i in sheet['A']:
    name = str(i.value)
# if length of the subject code is greater than 7 then 
    if len(name) >7:
        kill = name.replace(" ", "")
        name2 = re.split(',', kill)
        for na in name2:
            if len(na) == 7:
                todays_fa.append(na.upper())
            elif len(na) == 9:
                todays_fa.append(na.upper())
            else:
                print("the code should be recheck" +str(na))

        
    elif len(name) == 7:
        todays_fa.append(name)
    elif len(name) == 9:
        todays_fa.append(name)
    else:
        print("kindly check "+str(name))    
        

print(len(todays_fa))

not_down = []
no_added = []
path2=path ="C:\\Users\\Hxtreme\\Downloads\\"

#Extraction of data:


sheet = wb.active

J = 1
driver = webdriver.Chrome(executable_path="D:\\python files for moodle work\\chromedriver.exe")
driver.get('https://accounts.google.com/ServiceLogin/identifier?continue=https%3A%2F%2Fmail.google.com&sacu=1&passive=1209600&hl=en&acui=0&flowName=GlifWebSignIn&flowEntry=ServiceLogin&cid=1&TL=AKqFyY9fS1dRw8Qp3lKdBHb1jqC_m2CKQe6WwEU_FRewwf9xyEqNOvyklJyp7Jt8')
driver.implicitly_wait(100)

loginBox = driver.find_element_by_xpath('//*[@id ="identifierId"]')
loginBox.send_keys("email")

nextButton = driver.find_element_by_xpath('//*[@id ="identifierNext"]')
nextButton.click()

passWordBox = driver.find_element_by_xpath('//*[@id ="password"]/div[1]/div / div[1]/input')
passWordBox.send_keys("Password")


nextButton = driver.find_element_by_xpath('//*[@id ="passwordNext"]')
nextButton.click()
time.sleep(10)

driver.implicitly_wait(5000)

driver.execute_script("window.open('');")
driver.switch_to.window(driver.window_handles[J])

driver.get('https://moodle.bitsathy.ac.in/grade/export/xls/index.php?id=10793')
self_in = driver.find_element_by_xpath("/html[1]/body[1]/div[4]/div[1]/div[1]/div[1]/section[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[3]/div[1]/a[1]")
self_in.click()


for list in range(0, len(todays_fa)):
    try:
        code = subjects.get(todays_fa[list])
        if code == None:
            no_added.append(code)
            print(bcolors.FAIL+ str(todays_fa[list])+ "is not added in directory" +bcolors.FAIL)
        else:
            path = "https://moodle.bitsathy.ac.in/grade/export/xls/index.php?id="+ code 
            driver.get(path)
            time.sleep(4)
            export = driver.find_element_by_xpath("//input[@id='id_submitbutton']")
            export.click()
            time.sleep(4)
            print("has been downloaded" +str(todays_fa[list]))
            not_down.append(code)
            driver.implicitly_wait(20)
            stat1 = driver.find_element_by_xpath("//a[@id='yui_3_17_2_1_1690358424244_393']")
            stat1.click()
            
            
            
                        
    except:
        print(bcolors.WARNING +"download manually" +str(todays_fa[list]) + bcolors.ENDC)

driver.implicitly_wait(20)
driver.close()
# CLosing Data 
print("the total number of FA downloaded:" + " " +str(len(not_down)))
total_FA = len(todays_fa)
print("the total number of files should be downloaded:" + " " +str(total_FA))

not_added = len(no_added)
print("the total number of fa codes are not added till now" +" " +str(not_added))



