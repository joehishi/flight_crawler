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

url1 = 'https://www.shanghaiairport.com/cn/flights.html'
headers = {
'User-Agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36'
}


# browser driver
chrome_driver = '/Users/huaiyuchen/Downloads/selenium_rel/chromedriver'
firefox_mac = '/Users/huaiyuchen/Downloads/selenium_rel/geckodriver'
firefox_linux = './geckodriver'
firefox_win = './geckodriver-v0.26.0-win64/geckodriver.exe'
driver_path = firefox_linux

plan_time_msg = '//td[@class="TD1"]'
actual_time_msg = '//td[@class="TD7"]'
detail_msg = '//td[@class="TD8"]/a'
#在详细信息页中
port_msg = '//td[@class="TD2"]'  
dst_msg = '//div[@class="GoDestination GoBox"]//span'  
plane_msg = '//td[@class="TD3"]'  

hangban_msg = '//td[@class="TD2"]'  
hangzhan_msg = '//td[@class="TD4"]'  

select_day_msg = '//div[@class="SelectBox SelecttimeDays z2"]'
yesterday_msg = '//div[@class="SelectBox SelecttimeDays z2"]/dl/dt[1]/a'


start_button = '//a[@class="SpecialTipsClose"]'
search_button = '//a[@id="btnSearch"]'
next_button = '//a[@class="next"]'

airport_col = '//div[@class="SelectBox SelectFlightAirport z3"]'
pudong_airport = '//div[@class="SelectBox SelectFlightAirport z3"]/dl/dt[3]'

def Init_geturl(url):
    # 初始化一个driver，并且指定chromedriver的路径 完成指定搜索浦东机场指令
    driver = webdriver.Firefox(executable_path=driver_path)
    # 请求网页
    driver.get(url)
    time.sleep(5)

    if(len(driver.find_elements_by_xpath(start_button))):   #处理弹出页面
        #print("here")
        driver.find_element_by_xpath(start_button).click()
    time.sleep(1)
    driver.find_element_by_xpath(airport_col).click()   #选浦东机场
    driver.find_element_by_xpath(pudong_airport).click()
    time.sleep(1)
    driver.find_element_by_xpath(select_day_msg).click()      #选昨天
    driver.find_element_by_xpath(yesterday_msg).click()  
    time.sleep(1)
    driver.find_element_by_xpath(search_button).click()   #默认国内出发
    time.sleep(1)
    return driver

def further_page(url,headers):
    #访问查看更多选项的网站,获取目的地、登机口、机型
    response = requests.get(url,headers = headers)
    html_doc = response.content.decode("utf-8")
    tree = etree.HTML(html_doc)
    # dst[0].text => str
    dst = tree.xpath(dst_msg)
    port = tree.xpath(port_msg)
    plane = tree.xpath(plane_msg)
    return dst[0].text,port[0].text,plane[0].text


driver = Init_geturl(url1)
dsts = []
ports = []
planes = []
plan_times = []
actual_times = []

print('国内出发=======================>')


while(len(driver.find_elements_by_xpath(next_button))):
    hangbans = driver.find_elements_by_xpath(hangban_msg)
    sel = set(range(len(hangbans))) 
    #for i in range(len(hangbans)):    #剔除联合航班 
    #    if(len(hangbans[i].text.split()) > 1):
    #        sel.discard(i)

    hangzhans = driver.find_elements_by_xpath(hangzhan_msg)
    for i in list(sel):    #剔除TS1以外 
        #print(hangzhans[i].text[-4:-1])
        if(hangzhans[i].text[-4:-1] != 'TS1'):
            sel.discard(i)

    plans = driver.find_elements_by_xpath(plan_time_msg)
    time.sleep(1)
    plan_times = plan_times + [plans[i].text for i in list(sel)]
    actuals = driver.find_elements_by_xpath(actual_time_msg)
    time.sleep(1)
    actual_times = actual_times + [actuals[i].text[-5:] for i in list(sel)]

    details = driver.find_elements_by_xpath(detail_msg)
    time.sleep(1)
    detail_pages = [details[i].get_property('href') for i in list(sel)]
    for detail_page in detail_pages:
        dst,port,plane = further_page(detail_page,headers=headers)
        time.sleep(0.5)
        dsts.append(dst)
        if(port is not None):
            ports.append(port[1:])
        else:
            ports.append(port)
        planes.append(plane)
    #break
    time.sleep(1)
    driver.find_element_by_xpath(next_button).click()
    time.sleep(1)

#跳出循环再来一次
hangbans = driver.find_elements_by_xpath(hangban_msg)
sel = set(range(len(hangbans))) 
#for i in range(len(hangbans)):    #剔除联合航班 
#    if(len(hangbans[i].text.split()) > 1):
#        sel.discard(i)

hangzhans = driver.find_elements_by_xpath(hangzhan_msg)
for i in list(sel):    #剔除TS1以外 
    #print(hangzhans[i].text[-4:-1])
    if(hangzhans[i].text[-4:-1] != 'TS1'):
        sel.discard(i)

plans = driver.find_elements_by_xpath(plan_time_msg)
time.sleep(1)
plan_times = plan_times + [plans[i].text for i in list(sel)]
actuals = driver.find_elements_by_xpath(actual_time_msg)
time.sleep(1)
actual_times = actual_times + [actuals[i].text[-5:] for i in list(sel)]

details = driver.find_elements_by_xpath(detail_msg)
time.sleep(1)
detail_pages = [details[i].get_property('href') for i in list(sel)]
for detail_page in detail_pages:
    dst,port,plane = further_page(detail_page,headers=headers)
    time.sleep(0.5)
    dsts.append(dst)
    if(port is not None):
        ports.append(port[1:])
    else:
        ports.append(port)
    planes.append(plane)
#break


driver.close()

Dom_byhour = np.zeros((24))    #国内出发137-147 by hour

workbook = xlwt.Workbook(encoding='utf-8')
sheet = workbook.add_sheet('PVG S1 Dom raw data')
head = ['计划时间', '实际时间', '登机口', '目的地','机型','是否辐射范围内']    #表头
for h in range(len(head)):
    sheet.write(0, h, head[h])
i = 1
for j in range(len(plan_times)):
    print(j)
    sheet.write(i, 0, plan_times[j])
    sheet.write(i, 1, actual_times[j])
    sheet.write(i, 2, ports[j])
    sheet.write(i, 3, dsts[j])
    sheet.write(i, 4, planes[j])
    if(ports[j] is not None):
        if(int(ports[j]) >= 137 and int(ports[j]) <= 147):
            sheet.write(i, 5, 'Yes')
            Dom_byhour[int(plan_times[j][0:2])] = Dom_byhour[int(plan_times[j][0:2])] + 1  #国内出发137-147 by hour
        else:
            sheet.write(i, 5, 'No')
    else:
        sheet.write(i, 5, 'No')
    i+=1

sheet_by_hour = workbook.add_sheet('PVG S1 Dom daily by hour')
head = ['Daily Total', '05:00', '06:00', '07:00' ,'08:00' , '09:00' , '10:00' ,'11:00','12:00', '13:00', '14:00' ,'15:00' , 
        '16:00' , '17:00' ,'18:00','19:00', '20:00', '21:00' ,'22:00' , '23:00' ]    #表头
for h in range(len(head)):
    sheet_by_hour.write(0, h, head[h])
sheet_by_hour.write(1, 0, sum(Dom_byhour))
for ii in range(1,19):
    sheet_by_hour.write(1, ii, Dom_byhour[ii+4])
                         
sheet_by_period = workbook.add_sheet('PVG S1 Dom daily by period')
head = ['Daily Total', '05:00-09:59', '10:00-14:59', '15:00-16:59' ,'17:00-21:59' , '22:00 after' ]    #表头
for h in range(len(head)):
    sheet_by_period.write(0, h, head[h])
sheet_by_period.write(1, 0, sum(Dom_byhour))
sheet_by_period.write(1, 1, sum(Dom_byhour[5:10]))
sheet_by_period.write(1, 2, sum(Dom_byhour[10:15]))
sheet_by_period.write(1, 3, sum(Dom_byhour[15:17]))
sheet_by_period.write(1, 4, sum(Dom_byhour[17:22]))
sheet_by_period.write(1, 5, sum(Dom_byhour[22:]))

########################################################################
print('国际出发=======================>')

plan_time_msg = '//td[@class="TD1"]'
actual_time_msg = '//td[@class="TD7"]'
detail_msg = '//td[@class="TD8"]/a'
#在详细信息页中
port_msg = '//td[@class="TD2"]'  
dst_msg = '//div[@class="GoDestination GoBox"]//span'  
plane_msg = '//td[@class="TD3"]'  

hangban_msg = '//td[@class="TD2"]'  
hangzhan_msg = '//td[@class="TD4"]'  


start_button = '//a[@class="SpecialTipsClose"]'
search_button = '//a[@id="btnSearch"]'
next_button = '//a[@class="next"]'

direction_button = '//span[@id="direction"]'
domestic_dep_button = '//div[@class="SelectBox SelectDirection z3"]//dt[1]'
domestic_arr_button = '//div[@class="SelectBox SelectDirection z3"]//dt[2]'
foreign_dep_button = '//div[@class="SelectBox SelectDirection z3"]//dt[3]'
foreign_arr_button = '//div[@class="SelectBox SelectDirection z3"]//dt[4]'

airport_col = '//div[@class="SelectBox SelectFlightAirport z3"]'
pudong_airport = '//div[@class="SelectBox SelectFlightAirport z3"]/dl/dt[3]'

def Init_geturl2(url):
    # 初始化一个driver，并且指定chromedriver的路径 完成指定搜索浦东机场指令
    driver = webdriver.Firefox(executable_path=driver_path)
    # 请求网页
    driver.get(url)
    time.sleep(5)

    if(len(driver.find_elements_by_xpath(start_button))):   #处理弹出页面
        #print("here")
        driver.find_element_by_xpath(start_button).click()
    time.sleep(1)
    driver.find_element_by_xpath(direction_button).click()
    driver.find_element_by_xpath(foreign_dep_button).click()   #国际出发
    time.sleep(1)
    driver.find_element_by_xpath(airport_col).click()
    driver.find_element_by_xpath(pudong_airport).click()   #浦东机场
    time.sleep(1)
    driver.find_element_by_xpath(select_day_msg).click()      #选昨天
    driver.find_element_by_xpath(yesterday_msg).click()  
    time.sleep(1)
    driver.find_element_by_xpath(search_button).click()   #默认国内出发
    time.sleep(1)
    return driver

def further_page2(url,headers):
    #访问查看更多选项的网站,获取目的地、登机口、机型
    response = requests.get(url,headers = headers)
    html_doc = response.content.decode("utf-8")
    tree = etree.HTML(html_doc)
    # dst[0].text => str
    dst = tree.xpath(dst_msg)
    port = tree.xpath(port_msg)
    plane = tree.xpath(plane_msg)
    return dst[0].text,port[0].text,plane[0].text


driver = Init_geturl2(url1)
dsts = []
ports = []
planes = []
plan_times = []
actual_times = []


while(len(driver.find_elements_by_xpath(next_button))):
    hangbans = driver.find_elements_by_xpath(hangban_msg)
    sel = set(range(len(hangbans))) 
    #for i in range(len(hangbans)):    #剔除联合航班 
    #    if(len(hangbans[i].text.split()) > 1):
    #        sel.discard(i)

    hangzhans = driver.find_elements_by_xpath(hangzhan_msg)
    for i in list(sel):    #剔除TS1以外 
        #print(hangzhans[i].text[-4:-1])
        if(hangzhans[i].text[-4:-1] != 'TS1'):
            sel.discard(i)

    plans = driver.find_elements_by_xpath(plan_time_msg)
    time.sleep(1)
    plan_times = plan_times + [plans[i].text for i in list(sel)]
    actuals = driver.find_elements_by_xpath(actual_time_msg)
    time.sleep(1)
    actual_times = actual_times + [actuals[i].text[-5:] for i in list(sel)]

    details = driver.find_elements_by_xpath(detail_msg)
    time.sleep(1)
    detail_pages = [details[i].get_property('href') for i in list(sel)]
    for detail_page in detail_pages:
        dst,port,plane = further_page(detail_page,headers=headers)
        time.sleep(0.5)
        dsts.append(dst)
        if(port is not None):
            ports.append(port[1:])
        else:
            ports.append(port)
        planes.append(plane)
    #break
    time.sleep(1)
    driver.find_element_by_xpath(next_button).click()
    time.sleep(1)

#跳出循环再来一次
hangbans = driver.find_elements_by_xpath(hangban_msg)
sel = set(range(len(hangbans))) 
#for i in range(len(hangbans)):    #剔除联合航班 
#    if(len(hangbans[i].text.split()) > 1):
#        sel.discard(i)

hangzhans = driver.find_elements_by_xpath(hangzhan_msg)
for i in list(sel):    #剔除TS1以外 
    #print(hangzhans[i].text[-4:-1])
    if(hangzhans[i].text[-4:-1] != 'TS1'):
        sel.discard(i)

plans = driver.find_elements_by_xpath(plan_time_msg)
time.sleep(1)
plan_times = plan_times + [plans[i].text for i in list(sel)]
actuals = driver.find_elements_by_xpath(actual_time_msg)
time.sleep(1)
actual_times = actual_times + [actuals[i].text[-5:] for i in list(sel)]

details = driver.find_elements_by_xpath(detail_msg)
time.sleep(1)
detail_pages = [details[i].get_property('href') for i in list(sel)]
for detail_page in detail_pages:
    dst,port,plane = further_page(detail_page,headers=headers)
    time.sleep(0.5)
    dsts.append(dst)
    if(port is not None):
        ports.append(port[1:])
    else:
        ports.append(port)
    planes.append(plane)
    #break


driver.close()

cur = datetime.date.today() + datetime.timedelta(-1)#datetime.datetime.now()

foreign_byhour = np.zeros((24))    #国际出发 by hour

output_file = './output/浦东机场_航班时间'
timelist = str(cur.year) + str(cur.month) + str(cur.day)
output_file = output_file + timelist + '.xls'

sheet = workbook.add_sheet('PVG S1 Intl raw data')
head = ['计划时间', '实际时间', '登机口', '目的地','机型']    #表头
for h in range(len(head)):
    sheet.write(0, h, head[h])
i = 1
for j in range(len(plan_times)):
    print(j)
    sheet.write(i, 0, plan_times[j])
    sheet.write(i, 1, actual_times[j])
    sheet.write(i, 2, ports[j])
    sheet.write(i, 3, dsts[j])
    sheet.write(i, 4, planes[j])
    foreign_byhour[int(plan_times[j][0:2])] = foreign_byhour[int(plan_times[j][0:2])] + 1  #国际出发
    i+=1

sheet_intl_by_hour = workbook.add_sheet('PVG S1 Intl daily by hour')
head = ['Daily Total', '05:00', '06:00', '07:00' ,'08:00' , '09:00' , '10:00' ,'11:00','12:00', '13:00', '14:00' ,'15:00' , 
        '16:00' , '17:00' ,'18:00','19:00', '20:00', '21:00' ,'22:00' , '23:00' ]    #表头
for h in range(len(head)):
    sheet_intl_by_hour.write(0, h, head[h])
sheet_intl_by_hour.write(1, 0, sum(foreign_byhour))
for ii in range(1,19):
    sheet_intl_by_hour.write(1, ii, foreign_byhour[ii+4])
                         
sheet_intl_by_period = workbook.add_sheet('PVG S1 Intl daily by period')
head = ['Daily Total', '05:00-09:59', '10:00-14:59', '15:00-16:59' ,'17:00-21:59' , '22:00 after' ]    #表头
for h in range(len(head)):
    sheet_intl_by_period.write(0, h, head[h])
sheet_intl_by_period.write(1, 0, sum(foreign_byhour))
sheet_intl_by_period.write(1, 1, sum(foreign_byhour[5:10]))
sheet_intl_by_period.write(1, 2, sum(foreign_byhour[10:15]))
sheet_intl_by_period.write(1, 3, sum(foreign_byhour[15:17]))
sheet_intl_by_period.write(1, 4, sum(foreign_byhour[17:22]))
sheet_intl_by_period.write(1, 5, sum(foreign_byhour[22:]))

workbook.save(output_file)
print('昨天航班写入excel成功')



