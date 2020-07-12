# -*- coding: utf-8 -*-
"""
Created on Wed May 20 14:13:56 2020

@author: Tom Lee
"""

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By
from selenium.common.exceptions import ElementNotVisibleException

import datetime #날짜
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import os
import csv
import threading
import pandas as pd
from time import sleep

options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('disable-gpu')
options.add_argument("lang=ko_KR")

url = 'https://www.k-startup.go.kr/common/announcement/announcementList.do?mid=30004&bid=701&searchAppAt=A'
keyword = ['ICT', 'IoT', 'AI', '4차산업', 'SW', 'S/W', '인공지능', '정보통신', '기술혁신',  
                '소프트웨어', '전자', '딥러닝', '머신러닝', '기술개발', 'R&D', '소상공인', '4차 산업', '인공 지능',
                '정보 통신', '기술 혁신', '기술 개발', '소상 공인']

def KstartUpCrawling():
    global final_df;

    driver = webdriver.Chrome(executable_path=os.path.abspath("chromedriver.exe"), options=options);
    driver.get(url);
    is_none = '';
    while 1:
        try:
            if is_none == 'display: none;':
                raise StopIteration
            more_btn = driver.find_element_by_css_selector('#listPlusAdd')
            driver.find_element_by_css_selector('#listPlusAdd > a').click()   #공고 더보기
            is_none = more_btn.get_attribute('style')
        except StopIteration:
            print("태그가 더이상 존재하지 않습니다.")
            break    
    list_num = int(driver.find_element_by_xpath('//*[@id="searchAnnouncementVO"]/div[2]/div[1]/span').text)
    ann_list = pd.DataFrame([list(map(lambda a: a, range(2)))] * list_num)
    for i in range(list_num):
        ann_list[0][i] = driver.find_element_by_css_selector(f'#liArea{i} > h4 > a').get_attribute('text').strip('\n\t')
        ann_list[1][i] = driver.find_element_by_css_selector(f'#liArea{i} > ul > li:nth-child(3)').text.strip('마감일자 ')
    ann_list.columns = ['Announcement', 'DeadLines']

   # 키워드로 찾은 지원사업명 및 기한 list up
    search_list1 = []
    search_list2 = []
    for n in range(0, len(keyword)):
        for cmp_val in ann_list['Announcement']:
            if keyword[n] in cmp_val:                           
                found_idx = ann_list[ann_list['Announcement'] == cmp_val].index
                search_list1.append(cmp_val)
                search_list2.append(ann_list['DeadLines'][found_idx].values)
    
    search_list1 = pd.DataFrame(search_list1)
    search_list2 = pd.DataFrame(search_list2)
    search_list1.columns = ['Announcement']
    search_list2.columns = ['DeadLines']
    # final_df = 최종 data
    final_df = pd.concat([search_list1, search_list2], axis=1)
    final_df.sort_values(by='DeadLines', inplace=True)
    final_df.drop_duplicates(['Announcement'], keep='last', inplace=True)

def main():
    KstartUpCrawling()
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    json_file_name = 'my-project-1560401214468-2fe2c8ff3630.json'
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
    gc = gspread.authorize(credentials)

    spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1kb-_2ewhsZh9y_-5z2fRuCCoVxO-7ctRIcR9qsvV7cw'
    # 스프레스시트 문서 가져오기 
    doc = gc.open_by_url(spreadsheet_url)
    sleep(0.5);
    # 시트 선택하기
    ret = 0;
    worksheet_list = doc.worksheets()
    for seek_to_worksheet in worksheet_list:
        if 'K-StartUp' in seek_to_worksheet.title:
            old_worksheet = doc.worksheet(seek_to_worksheet.title)
            ret = 1
            break
        else:
            ret = 0;
     
    date = datetime.datetime.now().strftime('%Y-%m-%d')
    sheet_name = 'K-StartUp_' + date       
    
    if ret == 1:
        doc.del_worksheet(old_worksheet)

    new_worksheet = doc.add_worksheet(title=sheet_name, rows=100, cols=20)
    for i in range(0, len(final_df['Announcement'])):
        new_worksheet.insert_row([final_df['Announcement'].iloc[i], final_df['DeadLines'].iloc[i]], i + 1)
        sleep(0.5);
    
    print('\nCrawling is successfully done')
    now_time = datetime.datetime.now()
    now_date = 'Finish time is : ' + now_time.strftime('%Y-%m-%d %H:%M:%S')
    print(now_date)

    set_time = 86400
    threading.Timer(set_time, KstartUpCrawling).start()  # 설정한 시간 후 다시 crawling하는 부분 (Unit : sec)

main()    

