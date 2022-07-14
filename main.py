import openpyxl
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import subprocess
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from bs4 import BeautifulSoup
import time
import datetime
import re
import sys
import os
import requests
from PyQt5.QtWidgets import QWidget, QApplication,QTreeView,QFileSystemModel,QVBoxLayout,QPushButton,QInputDialog,QLineEdit,QMainWindow,QMessageBox
from PyQt5.QtCore import QCoreApplication
from selenium.webdriver import ActionChains
from datetime import datetime,date,timedelta
import time
import requests
from bs4 import BeautifulSoup as bs
import sys,os,shutil
from PyQt5.QtWidgets import QWidget, QApplication,QTreeView,QFileSystemModel,QVBoxLayout,QPushButton,QInputDialog,QLineEdit,QMainWindow
import chromedriver_autoinstaller
import random


chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
driver_path = f'./{chrome_ver}/chromedriver.exe'
if os.path.exists(driver_path):
    print(f"chromedriver is installed: {driver_path}")
else:
    print(f"install the chrome driver(ver: {chrome_ver})")
    chromedriver_autoinstaller.install(True)

# print(driver_path)
options = webdriver.ChromeOptions()
# options.add_argument('headless')
options.add_experimental_option("detach", True)
options.add_experimental_option("excludeSwitches", ["enable-logging"])


#키워드입력

keyword="guitar"     #키워드입력
requiredNoItem=1000   #검색갯수



browser = webdriver.Chrome(driver_path, options=options)
browser.maximize_window()
url="https://www.youtube.com/results?search_query={}".format(keyword)
# url="https://www.youtube.com"
browser.get(url)
browser.implicitly_wait(5)
time.sleep(3)
print("페이지 열기 완료")

# 스크롤 계속 내리면서 갯수확인하는 로직

prevNoItem=0
i=1
while True:
    print(i,"번째 스크롤 다운중...")
    browser.find_element(By.TAG_NAME,'body').send_keys(Keys.PAGE_DOWN)
    time.sleep(0.5)
    if i%10==0:
        soup=BeautifulSoup(browser.page_source,'lxml')
        totalDiv=soup.find('div',attrs={"id":"contents"})
        eachItem=totalDiv.find_all('ytd-video-renderer')
        noItem=len(eachItem)
        print("포스팅갯수:", noItem, "현재시간:", datetime.now())
        if noItem==prevNoItem:
            break
        elif noItem>=requiredNoItem+20:
            break
        prevNoItem=noItem
        time.sleep(random.randint(3,8))
    i=i+1



#-------------------------------------------------------------------------------------------------
#동영상 검색 로직
wb=openpyxl.Workbook()
ws=wb.active
first_row=['키워드','동영상제목','누적조회수','재생시간','실시간 여부','등록일']
ws.append(first_row)
#비디오 태그를 모두 찾는다
videoDivs=browser.find_elements(By.TAG_NAME,'ytd-video-renderer')

#비디오 태그내에서 정보를 추출
for index,videoDiv in enumerate(videoDivs):
    print(index, "번째 확인중")
    isRealTime="해당없음"
    title=videoDiv.find_element(By.TAG_NAME,'yt-formatted-string').text # 타이틀을 추출
    print(title)
    clicks=videoDiv.find_element(By.CLASS_NAME,'style-scope ytd-video-meta-block').text
    print("rawclicks",clicks)

    if clicks.find("시청")>=0: # 클릭수에 시청중이라고 되어있으며 라이브 방송임
        clicks="실시간이라 해당 없음"
        isRealTime="실시간"
    elif clicks==0:
        clicks="없음"
    else:
        clicks=clicks.replace("조회수","").split("\n")[0].strip() # 텍스트에서 클릭수만 간추림
        if clicks.find("만") >= 0:
            clicks = re.sub(r'[^0-9.]', '', clicks).replace("회", "")
            clicks = float(clicks) * 10000
            clicks = (format(int(clicks), ","))
        elif clicks.find("천") >= 0:
            clicks = re.sub(r'[^0-9.]', '', clicks).replace("회", "")
            clicks = float(clicks) * 1000
            clicks = (format(int(clicks), ","))
        elif clicks.find("없음")>=0:
            clicks="조회수 없음"
        else:
            try:
                clicks = re.sub(r'[^0-9.]', '', clicks).replace("회", "")
                clicks = (format(int(clicks), ","))
            except:
                clicks = re.sub(r'[^0-9.]', '', clicks).replace("회", "")

    print(clicks)
    try:
        runningTime=videoDiv.find_element(By.CLASS_NAME,'style-scope ytd-thumbnail-overlay-time-status-renderer').text
        print(runningTime)
        if runningTime=="":
            print("끝")
            break
    except:
        runningTime="없음"
        print(runningTime)
    postingDate=videoDiv.find_element(By.CLASS_NAME,'style-scope ytd-video-meta-block').text
    if postingDate.find("시청")>=0:
        postingDate="없음"
    elif postingDate==0:
        postingDate="없음"
    else:
        postingDate = postingDate.replace("조회수", "").split("\n")[-1].strip()
    print(isRealTime)
    print(postingDate)
    data=[keyword,title,clicks,runningTime,isRealTime,postingDate]
    ws.append(data)
    print("---------------")
wb.save('result_videos.xlsx')
print("동영상 작업완료")
#-----------------------------------------------------------------------------



#채널 검색 로직
xb=openpyxl.Workbook()
xs=xb.active
first_row=['키워드','채널명','누적구독자수','누적동영상수']
xs.append(first_row)

channelDivs=browser.find_elements(By.TAG_NAME,'ytd-channel-renderer')

if len(channelDivs)>=1:
    for index,channelDiv in enumerate(channelDivs):
        print(index, "번째 확인중")
        # print(channelDiv)
        title = channelDiv.find_element(By.CLASS_NAME, 'style-scope ytd-channel-name').text  # 타이틀을 추출
        print(title)
        subscribers = channelDiv.find_element(By.ID, 'subscribers').text.replace("구독자","").strip()
        if subscribers=="":
            subscribers="없음"
        if subscribers.find("만")>=0:
            subscribers.replace("명","")
            subscribers = re.sub(r'[^0-9.]', '', subscribers)
            subscribers=int(float(subscribers)*10000)
        elif subscribers.find("천")>=0:
            subscribers.replace("명","")
            subscribers = re.sub(r'[^0-9.]', '', subscribers)
            subscribers=int(float(subscribers)*1000)
        else:
            subscribers=subscribers.replace("명","")
        print(subscribers)
        totalVideos = channelDiv.find_element(By.ID, 'video-count').text.replace("동영상","").strip()
        print(totalVideos)
        data = [keyword, title, subscribers, totalVideos]
        xs.append(data)
        print("---------------")
    xb.save('result_channel.xlsx')
    print("채널 작업완료")
else:
    print("채널 없음")