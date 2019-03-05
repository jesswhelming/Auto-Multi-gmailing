import pandas as pd
import datetime
import sys
from selenium import webdriver
import time
from selenium.webdriver.chrome.options import Options

#current date
date = datetime.datetime.today().strftime('%Y-%m-%d')

#run webengine
driver = webdriver.Chrome('/Users/jeesuyang/Documents/project/AutoGmailing/webdriver/chrome/chromedriver')
driver.set_window_size(1280, 1280)
driver.implicitly_wait(30)
time.sleep(2)

driver.get("https://google.com/gmail")
driver.implicitly_wait(30)
driver.find_element_by_id('identifierId').send_keys('YourGmailID')
time.sleep(1)
driver.implicitly_wait(30)
driver.find_element_by_xpath('//*[@id="identifierNext"]').click()
time.sleep(1)
driver.implicitly_wait(30)
driver.find_element_by_name('password').send_keys('YourGmailPassword')
time.sleep(2)
driver.implicitly_wait(30)
driver.find_element_by_xpath('//*[@id="passwordNext"]').click()
time.sleep(2)
driver.implicitly_wait(30)
# time.sleep(3)


#read excelfile
data = pd.read_excel (r'/Users/jeesuyang/Documents/project/AutoGmailing/email_info.xlsx')
excel_info = pd.DataFrame(data, columns= ['emailAdd', 'childName','courseName']).values.tolist()
print(excel_info)

#auto gmailing
for e in range(len(excel_info)):
    emailAdd = excel_info[e][0]
    childName = excel_info[e][1]
    courseName = excel_info[e][2]
    driver.find_element_by_css_selector("div.T-I.J-J5-Ji.T-I-KE.L3").click()
    driver.implicitly_wait(30)
    time.sleep(2)
    driver.find_element_by_name('to').send_keys(emailAdd)
    driver.implicitly_wait(30)
    driver.find_element_by_name('subjectbox').send_keys('Absence notice')
    driver.implicitly_wait(30)
    driver.find_element_by_class_name('LW-avf').send_keys('''Dear parents,
    Your child '''+childName+ ''' received an absence on '''+ date +''' in '''+courseName+'''.
    Please review your child's Attendance on Power School for the exact number of absences your
    child has received thus far during the school year. Please be reminded of our attendance policy
    which indicates the permitted number of absences. If you have any further questions please contact
    to the school office.''')
    time.sleep(2)
    driver.implicitly_wait(30)
    driver.find_element_by_css_selector("div.T-I.J-J5-Ji.aoO.T-I-atl.L3").click()
    time.sleep(2)
    driver.implicitly_wait(30)

print("finish")
driver.close()
