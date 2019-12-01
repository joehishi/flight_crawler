import requests, time
from urllib.parse import urlencode
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import time
import pandas as pd
import xlwt
from lxml import etree
import numpy as np
import datetime
import os


class PD_spider:
    def __init__(self,driver_path,Today):
        self.headers = {
'User-Agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36'
        }
        self.url = 'https://www.shanghaiairport.com/cn/flights.html'

        self.plan_time_msg = '//td[@class="TD1"]'
        self.actual_time_msg = '//td[@class="TD7"]'
        self.detail_msg = '//td[@class="TD8"]/a'
        #在详细信息页中
        self.port_msg = '//td[@class="TD2"]'  
        self.dst_msg = '//div[@class="GoDestination GoBox"]//span'  
        self.plane_msg = '//td[@class="TD3"]'  

        self.hangban_msg = '//td[@class="TD2"]'  
        self.hangzhan_msg = '//td[@class="TD4"]'  

        self.select_day_col = '//div[@class="SelectBox SelecttimeDays z2"]'
        self.yesterday_col = '//div[@class="SelectBox SelecttimeDays z2"]/dl/dt[1]/a'

        self.start_button = '//a[@class="SpecialTipsClose"]'
        self.search_button = '//a[@id="btnSearch"]'
        self.next_button = '//a[@class="next"]'

        self.airport_col = '//div[@class="SelectBox SelectFlightAirport z3"]'
        self.pudong_airport_col = '//div[@class="SelectBox SelectFlightAirport z3"]/dl/dt[3]'

        self.direction_col = '//span[@id="direction"]'
        self.dom_dep_col = '//div[@class="SelectBox SelectDirection z3"]//dt[1]'
        #self.dom_arr_col = '//div[@class="SelectBox SelectDirection z3"]//dt[2]'
        self.intl_dep_col = '//div[@class="SelectBox SelectDirection z3"]//dt[3]'
        #self.intl_arr_col = '//div[@class="SelectBox SelectDirection z3"]//dt[4]'

        self.driver_path = driver_path

        self.plan_times,self.actual_times,self.dsts,self.ports,self.planes = [],[],[],[],[]
        self.Dom_byhour = np.zeros((24))    #国内出发137-147 by hour
        self.Intl_byhour = np.zeros((24))    #国际出发 by hour

        self.filename = './output/浦东机场_航班时间'

        self.today = int(Today)

    def run_all(self):
        if self.today == True:
            print("开始爬取今天浦东机场航班信息")
        else:
            print("开始爬取昨天浦东机场航班信息")
        print("打开国内出发航班页面")
        self.open_page(Dom = True)
        self.crawl_all()
        timelist = str(self.cur.year) + str(self.cur.month) + str(self.cur.day)
        self.filename = self.filename + timelist + '.xls'
        workbook = xlwt.Workbook(encoding='utf-8')
        self.write_raw_dom_data(workbook,'PVG S1 Dom raw data')
        self.write_dom_byhour(workbook,'PVG S1 Dom daily by hour')
        self.write_dom_byperiod(workbook,'PVG S1 Dom daily by period')
        self.plan_times,self.actual_times,self.dsts,self.ports,self.planes = [],[],[],[],[]
        print("打开国际出发航班页面")
        self.open_page(Dom = False)
        self.crawl_all()
        self.write_raw_intl_data(workbook,'PVG S1 intl raw data')
        self.write_intl_byhour(workbook,'PVG S1 intl daily by hour')
        self.write_intl_byperiod(workbook,'PVG S1 intl daily by period')
        if not os.path.exists('output'):
            os.mkdir('output')
        workbook.save(self.filename)
        print('航班信息写入excel成功')

        
    def open_page(self,Dom):                #打开页面
        self.driver = webdriver.Firefox(executable_path=self.driver_path)
        self.driver.get(self.url)
        time.sleep(3)
        if(len(self.driver.find_elements_by_xpath(self.start_button))):   #处理弹出页面
            self.driver.find_element_by_xpath(self.start_button).click()
        time.sleep(3)
        self.driver.find_element_by_xpath(self.airport_col).click()   
        time.sleep(1)
        self.driver.find_element_by_xpath(self.pudong_airport_col).click()   #选浦东机场
        time.sleep(1)
        if(self.today == False):
            print('选择昨天')
            time.sleep(3)
            self.driver.find_element_by_xpath(self.select_day_col).click()      #选昨天
            time.sleep(1)
            self.driver.find_element_by_xpath(self.yesterday_col).click()  
            time.sleep(1)
            self.cur = datetime.date.today() + datetime.timedelta(-1)
        else:
            print('选择今天')
            self.cur = datetime.datetime.now()
            
        if(Dom == False):  
            time.sleep(3)                                                   #国际出发
            self.driver.find_element_by_xpath(self.direction_col).click()
            time.sleep(1)
            self.driver.find_element_by_xpath(self.intl_dep_col).click()        
            time.sleep(1)
        self.driver.find_element_by_xpath(self.search_button).click()                #开始搜索
        time.sleep(1)

    def crawl_all(self):                           #爬虫内容
        while(len(self.driver.find_elements_by_xpath(self.next_button))):
            plan_times_t,actual_times_t,dsts_t,ports_t,planes_t = self.parse_page_base()
            self.plan_times = self.plan_times + plan_times_t
            self.actual_times = self.actual_times + actual_times_t
            self.dsts = self.dsts + dsts_t
            self.ports = self.ports + ports_t
            self.planes = self.planes + planes_t
            time.sleep(1)
            self.driver.find_element_by_xpath(self.next_button).click()
            time.sleep(1)
        plan_times_t,actual_times_t,dsts_t,ports_t,planes_t = self.parse_page_base()
        self.plan_times = self.plan_times + plan_times_t
        self.actual_times = self.actual_times + actual_times_t
        self.dsts = self.dsts + dsts_t
        self.ports = self.ports + ports_t
        self.planes = self.planes + planes_t
        self.driver.close()

    def write_all(self):
        timelist = str(self.cur.year) + str(self.cur.month) + str(self.cur.day)
        self.filename = self.filename + timelist + '.xls'
        workbook = xlwt.Workbook(encoding='utf-8')
        self.write_raw_dom_data(workbook,'PVG S1 Dom raw data')
        self.write_dom_byhour(workbook,'PVG S1 Dom daily by hour')
        self.write_dom_byperiod(workbook,'PVG S1 Dom daily by period')
        self.write_raw_intl_data(workbook,'PVG S1 intl raw data')
        self.write_intl_byhour(workbook,'PVG S1 intl daily by hour')
        self.write_intl_byperiod(workbook,'PVG S1 intl daily by period')
        workbook.save(self.filename)
        print('航班信息写入excel成功')


    def parse_page_base(self):
        hangbans = self.driver.find_elements_by_xpath(self.hangban_msg)
        sel = set(range(len(hangbans))) 
        #for i in range(len(hangbans)):    #剔除联合航班 
        #    if(len(hangbans[i].text.split()) > 1):
        #        sel.discard(i)
        hangzhans = self.driver.find_elements_by_xpath(self.hangzhan_msg)
        for i in list(sel):    #剔除TS1以外 
            if(hangzhans[i].text[-4:-1] != 'TS1'):
                sel.discard(i)
        plans = self.driver.find_elements_by_xpath(self.plan_time_msg)
        time.sleep(1)
        plan_times = [plans[i].text for i in list(sel)]
        actuals = self.driver.find_elements_by_xpath(self.actual_time_msg)
        time.sleep(1)
        actual_times = [actuals[i].text[-5:] for i in list(sel)]
        details = self.driver.find_elements_by_xpath(self.detail_msg)
        time.sleep(1)
        detail_pages = [details[i].get_property('href') for i in list(sel)]
        dsts , ports , planes = [] , [] , []
        for detail_page in detail_pages:
            dst,port,plane = self.further_page(detail_page,headers=self.headers)
            time.sleep(0.5)
            dsts.append(dst)
            if(port is not None):
                ports.append(port[1:])
            else:
                ports.append(port)
            planes.append(plane)
        return plan_times,actual_times,dsts,ports,planes

    def further_page(self,url,headers):
        #访问查看更多选项的网站,获取目的地、登机口、机型
        response = requests.get(url,headers = headers)
        html_doc = response.content.decode("utf-8")
        tree = etree.HTML(html_doc)
        dst = tree.xpath(self.dst_msg)
        port = tree.xpath(self.port_msg)
        plane = tree.xpath(self.plane_msg)
        return dst[0].text,port[0].text,plane[0].text

    def write_raw_dom_data(self,workbook,name):
        sheet = workbook.add_sheet(name)
        head = ['计划时间', '实际时间', '登机口', '目的地','机型','是否辐射范围内']    #表头
        for h in range(len(head)):
            sheet.write(0, h, head[h])
        i = 1
        for j in range(len(self.plan_times)):
            #print(j)
            sheet.write(i, 0, self.plan_times[j])
            sheet.write(i, 1, self.actual_times[j])
            sheet.write(i, 2, self.ports[j])
            sheet.write(i, 3, self.dsts[j])
            sheet.write(i, 4, self.planes[j])
            if(self.ports[j] is not None):
                if(int(self.ports[j]) >= 137 and int(self.ports[j]) <= 147):   #国内出发137-147 by hour
                    sheet.write(i, 5, 'Yes')
                    self.Dom_byhour[int(self.plan_times[j][0:2])] = self.Dom_byhour[int(self.plan_times[j][0:2])] + 1  
                else:
                    sheet.write(i, 5, 'No')
            else:
                sheet.write(i, 5, 'No')
            i+=1

    def write_dom_byhour(self,workbook,name):
        sheet_by_hour = workbook.add_sheet(name)
        head = ['Daily Total', '05:00', '06:00', '07:00' ,'08:00' , '09:00' , '10:00' ,'11:00','12:00', '13:00', '14:00' ,'15:00' , 
        '16:00' , '17:00' ,'18:00','19:00', '20:00', '21:00' ,'22:00' , '23:00' ]    #表头
        for h in range(len(head)):
            sheet_by_hour.write(0, h, head[h])
        sheet_by_hour.write(1, 0, sum(self.Dom_byhour))
        for ii in range(1,19):
            sheet_by_hour.write(1, ii, self.Dom_byhour[ii+4])

    def write_dom_byperiod(self,workbook,name):
        sheet_by_period = workbook.add_sheet('PVG S1 Dom daily by period')
        head = ['Daily Total', '05:00-09:59', '10:00-14:59', '15:00-16:59' ,'17:00-21:59' , '22:00 after' ]    #表头
        for h in range(len(head)):
            sheet_by_period.write(0, h, head[h])
        sheet_by_period.write(1, 0, sum(self.Dom_byhour))
        sheet_by_period.write(1, 1, sum(self.Dom_byhour[5:10]))
        sheet_by_period.write(1, 2, sum(self.Dom_byhour[10:15]))
        sheet_by_period.write(1, 3, sum(self.Dom_byhour[15:17]))
        sheet_by_period.write(1, 4, sum(self.Dom_byhour[17:22]))
        sheet_by_period.write(1, 5, sum(self.Dom_byhour[22:]))

    def write_raw_intl_data(self,workbook,name):
        sheet = workbook.add_sheet(name)
        head = ['计划时间', '实际时间', '登机口', '目的地','机型']    #表头
        for h in range(len(head)):
            sheet.write(0, h, head[h])
        i = 1
        for j in range(len(self.plan_times)):
            #print(j)
            sheet.write(i, 0, self.plan_times[j])
            sheet.write(i, 1, self.actual_times[j])
            sheet.write(i, 2, self.ports[j])
            sheet.write(i, 3, self.dsts[j])
            sheet.write(i, 4, self.planes[j])
            self.Intl_byhour[int(self.plan_times[j][0:2])] = self.Intl_byhour[int(self.plan_times[j][0:2])] + 1  #国际出发
            i+=1

    def write_intl_byhour(self,workbook,name):
        sheet_intl_by_hour = workbook.add_sheet(name)
        head = ['Daily Total', '05:00', '06:00', '07:00' ,'08:00' , '09:00' , '10:00' ,'11:00','12:00', '13:00', '14:00' ,'15:00' , 
        '16:00' , '17:00' ,'18:00','19:00', '20:00', '21:00' ,'22:00' , '23:00' ]    #表头
        for h in range(len(head)):
            sheet_intl_by_hour.write(0, h, head[h])
        sheet_intl_by_hour.write(1, 0, sum(self.Intl_byhour))
        for ii in range(1,19):
            sheet_intl_by_hour.write(1, ii, self.Intl_byhour[ii+4])

    def write_intl_byperiod(self,workbook,name):
        sheet_intl_by_period = workbook.add_sheet(name)
        head = ['Daily Total', '05:00-09:59', '10:00-14:59', '15:00-16:59' ,'17:00-21:59' , '22:00 after' ]    #表头
        for h in range(len(head)):
            sheet_intl_by_period.write(0, h, head[h])
        sheet_intl_by_period.write(1, 0, sum(self.Intl_byhour))
        sheet_intl_by_period.write(1, 1, sum(self.Intl_byhour[5:10]))
        sheet_intl_by_period.write(1, 2, sum(self.Intl_byhour[10:15]))
        sheet_intl_by_period.write(1, 3, sum(self.Intl_byhour[15:17]))
        sheet_intl_by_period.write(1, 4, sum(self.Intl_byhour[17:22]))
        sheet_intl_by_period.write(1, 5, sum(self.Intl_byhour[22:]))