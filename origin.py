#-*- coding:utf-8 -*-

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
import time, datetime
import schedule
from pytz import timezone
import pytz
from random import *
from openpyxl import load_workbook
from selenium.webdriver.chrome.options import Options
options = Options()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
#made by 1301 all rights reserve
#use selenium, openpyxl

#random start time 7:20~7:40
print("자가진단 프로세스 시작")

def jaga():
    load_wb = load_workbook("./inform.xlsx", data_only=True)
    sheet1 = load_wb['Sheet1']
    namelist = []
    birlist = []
    passlist = []
    list1 = []
    passlist1 = []
    banlist = []
    # 0 김동현 1 서원준 2 이준엽 3 조현성 4 박재범 5 김재용
    for i in sheet1.rows:
        name = i[0].value
        namelist.append(name)
    for i in sheet1.rows:
        bir = i[1].value
        birlist.append(bir)
    for i in sheet1.rows:
        pass1 = i[2].value
        passlist.append(pass1)
    i = 0
    #save as namelist, birlist, passlist
    browser = webdriver.Chrome('./chromedriver.exe',options = options)
    browser.implicitly_wait(15)

    #random user pick
    for k in range(len(namelist)):
            a = randint(0,len(namelist)-1)
            while a in list1:
                a =randint(0,len(namelist)-1)
            list1.append(a)
    print(list1)

    while i <= len(namelist) - 1:
        #random timedealy each person
        t = list1[i]
        if t not in banlist:
            timedealy = randint(10,30)
            passlist1 = list(map(int,' '.join(str(passlist[t])).split()))
            emptylist = [(4,0),(5,1),(5,2),(5,3),(5,4),(6,0),(7,0),(8,1),(8,2),(8,3),(8,4),(9,0)]

            browser.get("https://hcs.eduro.go.kr/#/loginHome")
            #school select
            browser.delete_all_cookies()
            browser.find_element_by_id("btnConfirm2").click()
            browser.find_element_by_class_name("searchBtn").click()
            Select(browser.find_element_by_id("sidolabel")).select_by_value("03")
            Select(browser.find_element_by_id("crseScCode")).select_by_value("4")
            scb = browser.find_element_by_id("orgname")
            scb.send_keys("대구일과학고등학교")
            scb.send_keys(Keys.RETURN)
            kk = browser.find_element_by_xpath('//*[@id="softBoardListLayer"]/div[2]/div[1]/ul/li/a')
            kk.click()
            #user input
            browser.find_element_by_xpath('//*[@id="softBoardListLayer"]/div[2]/div[1]/ul/li/a').click()
            browser.find_element_by_xpath('//*[@id="softBoardListLayer"]/div[2]/div[2]/input').click()
            browser.find_element_by_xpath('//*[@id="user_name_input"]').send_keys(namelist[t])
            browser.find_element_by_xpath('//*[@id="birthday_input"]').send_keys(birlist[t])
            browser.find_element_by_xpath('//*[@id="btnConfirm"]').click()
            time.sleep(5)
            browser.find_element_by_xpath('//*[@id="WriteInfoForm"]/table/tbody/tr/td/div/button/img').click()
            for j in range(4,10):
                if j == 5 or j == 8:
                    k = 1
                    for k in range(1,5):
                        tar = '//*[@id="password_mainDiv"]/div['+str(j) + ']/a[' + str(k) + ']'
                        checkkey = browser.find_element_by_xpath(tar)
                        if '빈칸' in checkkey.get_attribute('aria-label'):
                            emptylist.remove((j,k))
                else:
                    k = 0
                    tar = '//*[@id="password_mainDiv"]/div['+str(j) + ']/a'
                    checkkey = browser.find_element_by_xpath(tar)
                    if '빈칸' in checkkey.get_attribute('aria-label'):
                        emptylist.remove((j,k))
            for s in passlist1:
                (j,k) = emptylist[s]
                if k ==0:
                    tar = '//*[@id="password_mainDiv"]/div['+str(j) + ']/a'
                    browser.find_element_by_xpath(tar).click()
                else:
                    tar = '//*[@id="password_mainDiv"]/div['+str(j) + ']/a[' + str(k) + ']'
                    browser.find_element_by_xpath(tar).click()
            browser.find_element_by_xpath('//*[@id="btnConfirm"]').click()
            time.sleep(3)
            try:
                CheckPoint = browser.find_element_by_xpath('//*[@id="container"]/div/section[2]/div[2]/ul/li')
            except:
                CheckPoint = browser.find_element_by_xpath('//*[@id="container"]/div/section[2]/div[2]/ul/li[1]')
            time.sleep(3)
            if 'active' in CheckPoint.get_attribute('class'):
                browser.find_element_by_xpath('//*[@id="topMenuBtn"]').click()
                browser.find_element_by_xpath('//*[@id="topMenuWrap"]/ul/li[5]/button').click()
                time.sleep(1)
                Alert(browser).accept()
                browser.find_element_by_xpath('/html/body/app-root/div/div[1]/div/button').click()
                time.sleep(1)
                Alert(browser).accept()
                i = i + 1
                print(namelist[t] + " 자가진단 본인이 완료" + str(i))
            else:
                browser.find_element_by_xpath('//*[@id="container"]/div/section[2]/div[2]/ul/li[1]/a/span[1]').click()
                browser.find_element_by_xpath('//*[@id="survey_q1a1"]').click()
                browser.find_element_by_xpath('//*[@id="survey_q2a3"]').click()
                browser.find_element_by_xpath('//*[@id="survey_q3a1"]').click()
                browser.find_element_by_xpath('//*[@id="survey_q4a1"]').click()
                browser.find_element_by_xpath('//*[@id="survey_q5a1"]').click()
                browser.find_element_by_xpath('//*[@id="btnConfirm"]').click()
                browser.find_element_by_xpath('//*[@id="topMenuBtn"]').click()
                browser.find_element_by_xpath('//*[@id="topMenuWrap"]/ul/li[5]/button').click()
                time.sleep(1)
                Alert(browser).accept()
                browser.find_element_by_xpath('/html/body/app-root/div/div[1]/div/button').click()
                time.sleep(1)
                Alert(browser).accept()
                i = i + 1
                print(namelist[t] + " 자가진단 완료" + str(timedealy) + " " + str(i))
                time.sleep(timedealy)
    print("자가진단 완료"+ str(nowDay))
    browser.close()
schedule.every().day.at('06:55').do(jaga)
while True:
    nowDay = time.strftime('%D')
    schedule.run_pending()
    time.sleep(1)
