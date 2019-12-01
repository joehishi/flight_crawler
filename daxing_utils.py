import requests, time
from urllib.parse import urlencode
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import time
import pandas as pd
import xlwt
import datetime
import numpy as np
import os


class DX_spider:
    def __init__(self,driver_path,Today):
        self.headers = {
'User-Agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36'
        }
        self.url_arr = 'https://www.bdia.com.cn/#/flightarr'
        self.url_dep = 'https://www.bdia.com.cn/#/flightdep'

        self.row_msg = '//div[@class="flight-block owh"] '
        self.plantime_msg = '//span[@class="plan-time flight-t"]'
        self.actualtime_msg = '//span[@class="actual-time flight-t"]'
        self.port_msg = '//div[@class="boarding-box"]/span'
        self.dst_msg = '//div[@class="destination-place"]/span'
        self.esttime_msg = '//span[@class="estimate-time flight-t"]'

        self.more_msg = '//div[@class="selectmore"]'

        self.hangban_msg = '//div[@class="flight-number block-li"]'

        self.day_col = '//input[@placeholder="选择日期"]'
        self.yesterday_col = '/html/body/div[2]/div[1]/div[1]/ul/li[1]/span'

        self.driver_path = driver_path

        self.plan_times,self.actual_times,self.dsts,self.ports = [],[],[],[]

        self.dom_arr_byhour = np.zeros((24))    #国内到达全部 by hour
        self.dom_dep_byhour = np.zeros((24))    #仅国内出发 登机口A 的全部航班 by hour
        self.intl_dep_byhour = np.zeros((24))    #仅国际出发 登机口E 的全部航班 by hour

        self.today = int(Today)

        self.filename = './output/大兴机场_航班时间'

        self.row = 0

        if(self.today == False):
            self.cur = datetime.date.today() + datetime.timedelta(-1)
        else:
            self.cur = datetime.datetime.now()

    def run_all(self):
        if self.today == True:
            print("开始爬取今天大兴机场航班信息")
        else:
            print("开始爬取昨天大兴机场航班信息")
        timelist = str(self.cur.year) + str(self.cur.month) + str(self.cur.day)
        self.filename = self.filename + timelist + '.xls'
        workbook = xlwt.Workbook(encoding='utf-8')
        if self.today == True:                
            self.open_page(self.url_arr)
            self.parse_page_base(arr = True)
            sheet = self.write_arr_raw_data(workbook,'daxing raw data')
            self.plan_times,self.actual_times,self.dsts,self.ports = [],[],[],[]
            self.open_page(self.url_dep)
            self.parse_page_base(arr = False)
            self.write_dep_raw_data(sheet)
            self.write_hour_data(workbook,'A-by ',self.dom_dep_byhour)
            self.write_hour_data(workbook,'Dom arr ',self.dom_arr_byhour)
            self.write_hour_data(workbook,'E-by ',self.intl_dep_byhour)

        else:                        #昨天的仅能获取到出发信息
            sheet = workbook.add_sheet('daxing raw data')
            self.open_page(self.url_dep)
            self.parse_page_base(arr = False)
            self.write_dep_raw_data(sheet)
            self.write_hour_data(workbook,'A-by ',self.dom_dep_byhour)
            self.write_hour_data(workbook,'E-by ',self.intl_dep_byhour)
        if not os.path.exists('output'):
            os.mkdir('output')
        workbook.save(self.filename)
        print('航班信息写入excel成功')

    def open_page(self,url):                #打开页面
        self.driver = webdriver.Firefox(executable_path=self.driver_path)
        self.driver.get(url)
        time.sleep(3)
        if self.today == False:
            print('选择昨天')
            self.driver.find_element_by_xpath(self.day_col).click()
            time.sleep(1)
            self.driver.find_element_by_xpath(self.yesterday_col).click()  #选择昨天
            time.sleep(2)
        else:
            print('选择今天')
        while(len(self.driver.find_elements_by_xpath(self.more_msg))):
            time.sleep(2)
            self.driver.find_element_by_xpath(self.more_msg + '/span').click()
        time.sleep(1)

    def parse_page_base(self,arr):
        hangbans = self.driver.find_elements_by_xpath(self.hangban_msg)
        total_num = len(hangbans)
        sel = set(range(total_num)) 

        ports = self.driver.find_elements_by_xpath(self.port_msg)
        if arr == True:
            for i in range(total_num):    #剔除国际到达   基本也没有国际到达的..
                if(ports[i].text[0] == 'Ｅ'):
                    sel.discard(i)

        pltimes = self.driver.find_elements_by_xpath(self.plantime_msg)
        actimes = self.driver.find_elements_by_xpath(self.actualtime_msg)
        estimes = self.driver.find_elements_by_xpath(self.esttime_msg)
        #ports = driver.find_elements_by_xpath(port_msg)
        dsts = self.driver.find_elements_by_xpath(self.dst_msg)

        self.plan_times = [pltimes[i].text for i in list(sel)]
        self.actual_times = [actimes[i].text for i in list(sel) if i < len(actimes)]
        #estimes = [estimes[i].text for i in list(sel)]
        self.actual_times = self.actual_times 
        self.ports = [ports[i].text for i in list(sel)]
        self.dsts = [dsts[i].text for i in list(sel)]

        self.driver.close()


    def write_arr_raw_data(self,workbook,name):
        sheet = workbook.add_sheet(name)
        head = ['计划时间', '实际时间', '登机口字母', '登机口数字' ,'国内/国际' , '到达/出发' , '目的地/始发地' ,'是否辐射范围内']    #表头
        for h in range(len(head)):
            sheet.write(0, h, head[h])
        i = 1                             #行号
        for j in range(len(self.plan_times)):               #到达部分!!!
            #print(j)
            sheet.write(i, 0, self.plan_times[j])
            if(j < len(self.actual_times)):
                sheet.write(i, 1, self.actual_times[j][-5:])
            else:
                sheet.write(i, 1, None)
            sheet.write(i, 2, self.ports[j][0])
            sheet.write(i, 3, self.ports[j][1:])
            if(self.ports[j][0] == 'E'):
                sheet.write(i, 4, 'Internatinal')
            else:
                sheet.write(i, 4, 'Domestic')
                self.dom_arr_byhour[int(self.plan_times[j][0:2])] = self.dom_arr_byhour[int(self.plan_times[j][0:2])] + 1   #国内到达全部 by hour
            sheet.write(i, 5, 'Arrival')
            sheet.write(i, 6, self.dsts[j])
            if(self.ports[j][0] == 'E'):
                sheet.write(i, 7, 'No')
            else:
                sheet.write(i, 7, 'Yes')
            i+=1
        self.row = i
        return sheet

    def write_dep_raw_data(self,sheet):
        i = self.row
        for j in range(len(self.plan_times)):           #出发部分!!!
            #print(j)
            sheet.write(i, 0, self.plan_times[j])
    
            if(j < len(self.actual_times)):
                sheet.write(i, 1, self.actual_times[j][-5:])
            else:
                sheet.write(i, 1, None) 

            if(len(self.ports[j]) > 1):  #port 不为空
                sheet.write(i, 2, self.ports[j][0])
                sheet.write(i, 3, self.ports[j][1:])
                if(self.ports[j][0] == 'E'):
                    sheet.write(i, 4, 'Internatinal')
                    self.intl_dep_byhour[int(self.plan_times[j][0:2])] = self.intl_dep_byhour[int(self.plan_times[j][0:2])] + 1     #仅国际出发 登机口E 的全部航班 by hour
                else:
                    sheet.write(i, 4, 'Domestic')
                if(self.ports[j][0] == 'E' or self.ports[j][0] == 'A'):
                    sheet.write(i, 7, 'Yes')
                else:
                    sheet.write(i, 7, 'No')

                if(self.ports[j][0] == 'A'):
                    self.dom_dep_byhour[int(self.plan_times[j][0:2])] = self.dom_dep_byhour[int(self.plan_times[j][0:2])] + 1        #仅国内出发 登机口A 的全部航班 by hour
            else:                         #port 为空
                sheet.write(i, 2, None)
                sheet.write(i, 3, None)
                sheet.write(i, 4, 'Domestic')
                sheet.write(i, 7, 'No')
    
            sheet.write(i, 5, 'Departure')
            sheet.write(i, 6, self.dsts[j])
            i+=1

    def write_hour_data(self,workbook,name,seq):
        sheet_by_hour = workbook.add_sheet(name + 'hour')
        head = ['Daily Total', '05:00', '06:00', '07:00' ,'08:00' , '09:00' , '10:00' ,'11:00','12:00', '13:00', '14:00' ,'15:00' , 
            '16:00' , '17:00' ,'18:00','19:00', '20:00', '21:00' ,'22:00' , '23:00' ]    #表头
        for h in range(len(head)):
            sheet_by_hour.write(0, h, head[h])
        sheet_by_hour.write(1, 0, sum(seq))
        for ii in range(1,19):
            sheet_by_hour.write(1, ii, seq[ii+4])
                         
        heet_by_period = workbook.add_sheet(name + 'period')
        head = ['Daily Total', '05:00-09:59', '10:00-14:59', '15:00-16:59' ,'17:00-21:59' , '22:00 after' ]    #表头
        for h in range(len(head)):
            heet_by_period.write(0, h, head[h])
        heet_by_period.write(1, 0, sum(seq))
        heet_by_period.write(1, 1, sum(seq[5:10]))
        heet_by_period.write(1, 2, sum(seq[10:15]))
        heet_by_period.write(1, 3, sum(seq[15:17]))
        heet_by_period.write(1, 4, sum(seq[17:22]))
        heet_by_period.write(1, 5, sum(seq[22:]))