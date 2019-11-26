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

url_arr = 'https://www.bdia.com.cn/#/flightarr'
url_dep = 'https://www.bdia.com.cn/#/flightdep'
headers = {
'User-Agent':'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36'
}

########################################################################
#以下位到达部分

#browser driver
firefox_mac = '/Users/huaiyuchen/Downloads/selenium_rel/geckodriver'
firefox_win = './geckodriver-v0.26.0-win64/geckodriver.exe'
firefox_linux = './geckodriver'
driver_path = firefox_linux

msg_row = '//div[@class="flight-block owh"] '
plan_time = '//span[@class="plan-time flight-t"]'
actual_time = '//span[@class="actual-time flight-t"]'
port_msg = '//div[@class="boarding-box"]/span'
dst_msg = '//div[@class="destination-place"]/span'
est_time = '//span[@class="estimate-time flight-t"]'

more_msg = '//div[@class="selectmore"]'

hangban_msg = '//div[@class="flight-number block-li"]'

day_button = '//input[@placeholder="选择日期"]'
yesterday_button = '/html/body/div[2]/div[1]/div[1]/ul/li[1]/span'

driver = webdriver.Firefox(executable_path=driver_path)

print('到达航班=======================>')

# 请求网页
driver.get(url_arr)
time.sleep(3)


while(len(driver.find_elements_by_xpath(more_msg))):
    time.sleep(1)
    driver.find_element_by_xpath(more_msg + '/span').click()
time.sleep(1)


hangbans = driver.find_elements_by_xpath(hangban_msg)
total_num = len(hangbans)
sel = set(range(total_num)) 

ports = driver.find_elements_by_xpath(port_msg)
for i in range(total_num):    #剔除国际到达
    if(ports[i].text[0] == 'Ｅ'):
        sel.discard(i)

pltimes = driver.find_elements_by_xpath(plan_time)
actimes = driver.find_elements_by_xpath(actual_time)
estimes = driver.find_elements_by_xpath(est_time)
#ports = driver.find_elements_by_xpath(port_msg)
dsts = driver.find_elements_by_xpath(dst_msg)
print(total_num)

pltimes = [pltimes[i].text for i in list(sel)]
actimes = [actimes[i].text for i in list(sel) if i < len(actimes)]
#estimes = [estimes[i].text for i in list(sel)]
actimes = actimes 
ports = [ports[i].text for i in list(sel)]
dsts = [dsts[i].text for i in list(sel)]

driver.close()

########################################################################
#以下时出发部分

print('出发航班=======================>')


msg_row = '//div[@class="flight-block owh"] '
plan_time = '//span[@class="plan-time flight-t"]'
actual_time = '//span[@class="actual-time flight-t"]'
port_msg = '//div[@class="boarding-box"]/span'
dst_msg = '//div[@class="destination-place"]/span'
est_time = '//span[@class="estimate-time flight-t"]'

more_msg = '//div[@class="selectmore"]'

hangban_msg = '//div[@class="flight-number block-li"]'


day_button = '//input[@placeholder="选择日期"]'
yesterday_button = '/html/body/div[2]/div[1]/div[1]/ul/li[1]/span'

# 初始化一个driver，并且指定chromedriver的路径
# 请求网页
driver = webdriver.Firefox(executable_path=driver_path)
driver.get(url_dep)
time.sleep(3)

while(len(driver.find_elements_by_xpath(more_msg))):
    time.sleep(1)
    driver.find_element_by_xpath(more_msg + '/span').click()

hangbans = driver.find_elements_by_xpath(hangban_msg)
sel = set(range(len(hangbans)))

pltimes2 = driver.find_elements_by_xpath(plan_time)
actimes2 = driver.find_elements_by_xpath(actual_time)
estimes2 = driver.find_elements_by_xpath(est_time)
ports2 = driver.find_elements_by_xpath(port_msg)
dsts2 = driver.find_elements_by_xpath(dst_msg)

pltimes2 = [pltimes2[i].text for i in list(sel)]
actimes2 = [actimes2[i].text for i in list(sel) if i < len(actimes2)]
#estimes = [estimes[i].text for i in list(sel)]
actimes2 = actimes2 
ports2 = [ports2[i].text for i in list(sel)]
dsts2 = [dsts2[i].text for i in list(sel)] 

driver.close()

########################################################################
#excel写入部分

workbook = xlwt.Workbook(encoding='utf-8')    # gbk fow win
sheet = workbook.add_sheet('Today arr')

head = ['计划时间', '实际时间', '登机口字母', '登机口数字' ,'国内/国际' , '到达/出发' , '目的地/始发地' ,'是否辐射范围内']    #表头
for h in range(len(head)):
    sheet.write(0, h, head[h])
i = 1                             #行号

Dom_arr_byhour = np.zeros((24))    #国内到达全部 by hour

for j in range(len(pltimes)):               #到达部分!!!
    print(j)
    sheet.write(i, 0, pltimes[j])
    if(j < len(actimes)):
        sheet.write(i, 1, actimes[j][-5:])
    else:
        sheet.write(i, 1, None)
    sheet.write(i, 2, ports[j][0])
    sheet.write(i, 3, ports[j][1:])
    if(ports[j][0] == 'E'):
        sheet.write(i, 4, 'Internatinal')
    else:
        sheet.write(i, 4, 'Domestic')
        Dom_arr_byhour[int(pltimes[j][0:2])] = Dom_arr_byhour[int(pltimes[j][0:2])] + 1   #国内到达全部 by hour
    sheet.write(i, 5, 'Arrival')
    sheet.write(i, 6, dsts[j])
    if(ports[j][0] == 'E'):
        sheet.write(i, 7, 'No')
    else:
        sheet.write(i, 7, 'Yes')
    i+=1

byhour = np.zeros((24))    #仅国内出发 登机口A 的全部航班 by hour
forn_byhour = np.zeros((24))    #仅国际出发 登机口E 的全部航班 by hour

for j in range(len(pltimes2)):           #出发部分!!!
    print(j)
    sheet.write(i, 0, pltimes2[j])
    
    if(j < len(actimes2)):
        sheet.write(i, 1, actimes2[j][-5:])
    else:
        sheet.write(i, 1, None) 

    if(len(ports2[j]) > 1):  #port 不为空
        sheet.write(i, 2, ports2[j][0])
        sheet.write(i, 3, ports2[j][1:])
        if(ports2[j][0] == 'E'):
            sheet.write(i, 4, 'Internatinal')
            forn_byhour[int(pltimes2[j][0:2])] = forn_byhour[int(pltimes2[j][0:2])] + 1     #仅国际出发 登机口E 的全部航班 by hour
        else:
            sheet.write(i, 4, 'Domestic')
        if(ports2[j][0] == 'E' or ports2[j][0] == 'A'):
            sheet.write(i, 7, 'Yes')
        else:
            sheet.write(i, 7, 'No')

        if(ports2[j][0] == 'A'):
            byhour[int(pltimes2[j][0:2])] = byhour[int(pltimes2[j][0:2])] + 1        #仅国内出发 登机口A 的全部航班 by hour
        
    else:                         #port 为空
        sheet.write(i, 2, None)
        sheet.write(i, 3, None)
        sheet.write(i, 4, 'Domestic')
        sheet.write(i, 7, 'No')
    
    sheet.write(i, 5, 'Departure')
    sheet.write(i, 6, dsts2[j])
    i+=1

sheet_by_hour = workbook.add_sheet('A-by hour')
head = ['Daily Total', '05:00', '06:00', '07:00' ,'08:00' , '09:00' , '10:00' ,'11:00','12:00', '13:00', '14:00' ,'15:00' , 
        '16:00' , '17:00' ,'18:00','19:00', '20:00', '21:00' ,'22:00' , '23:00' ]    #表头
for h in range(len(head)):
    sheet_by_hour.write(0, h, head[h])
sheet_by_hour.write(1, 0, sum(byhour))
for ii in range(1,19):
    sheet_by_hour.write(1, ii, byhour[ii+4])
                         
heet_by_period = workbook.add_sheet('A-by period')
head = ['Daily Total', '05:00-09:59', '10:00-14:59', '15:00-16:59' ,'17:00-21:59' , '22:00 after' ]    #表头
for h in range(len(head)):
    heet_by_period.write(0, h, head[h])
heet_by_period.write(1, 0, sum(byhour))
heet_by_period.write(1, 1, sum(byhour[5:10]))
heet_by_period.write(1, 2, sum(byhour[10:15]))
heet_by_period.write(1, 3, sum(byhour[15:17]))
heet_by_period.write(1, 4, sum(byhour[17:22]))
heet_by_period.write(1, 5, sum(byhour[22:]))

sheet_Domarr_by_hour = workbook.add_sheet('Dom arr -by hour')
head = ['Daily Total', '05:00', '06:00', '07:00' ,'08:00' , '09:00' , '10:00' ,'11:00','12:00', '13:00', '14:00' ,'15:00' , 
        '16:00' , '17:00' ,'18:00','19:00', '20:00', '21:00' ,'22:00' , '23:00' ]    #表头
for h in range(len(head)):
    sheet_Domarr_by_hour.write(0, h, head[h])
sheet_Domarr_by_hour.write(1, 0, sum(Dom_arr_byhour))
for ii in range(1,19):
    sheet_Domarr_by_hour.write(1, ii, Dom_arr_byhour[ii+4])

sheet_Domarr_by_period = workbook.add_sheet('Dom arr -by period')
head = ['Daily Total', '05:00-09:59', '10:00-14:59', '15:00-16:59' ,'17:00-21:59' , '22:00 after' ]    #表头
for h in range(len(head)):
    sheet_Domarr_by_period.write(0, h, head[h])
sheet_Domarr_by_period.write(1, 0, sum(Dom_arr_byhour))
sheet_Domarr_by_period.write(1, 1, sum(Dom_arr_byhour[5:10]))
sheet_Domarr_by_period.write(1, 2, sum(Dom_arr_byhour[10:15]))
sheet_Domarr_by_period.write(1, 3, sum(Dom_arr_byhour[15:17]))
sheet_Domarr_by_period.write(1, 4, sum(Dom_arr_byhour[17:22]))
sheet_Domarr_by_period.write(1, 5, sum(Dom_arr_byhour[22:]))

sheet_forein_by_hour = workbook.add_sheet('E-by hour')
head = ['Daily Total', '05:00', '06:00', '07:00' ,'08:00' , '09:00' , '10:00' ,'11:00','12:00', '13:00', '14:00' ,'15:00' , 
        '16:00' , '17:00' ,'18:00','19:00', '20:00', '21:00' ,'22:00' , '23:00' ]    #表头
for h in range(len(head)):
    sheet_forein_by_hour.write(0, h, head[h])
sheet_forein_by_hour.write(1, 0, sum(forn_byhour))
for ii in range(1,19):
    sheet_forein_by_hour.write(1, ii, forn_byhour[ii+4])


sheet_forein_by_period = workbook.add_sheet('E-by period')
head = ['Daily Total', '05:00-09:59', '10:00-14:59', '15:00-16:59' ,'17:00-21:59' , '22:00 after' ]    #表头
for h in range(len(head)):
    sheet_forein_by_period.write(0, h, head[h])
sheet_forein_by_period.write(1, 0, sum(forn_byhour))
sheet_forein_by_period.write(1, 1, sum(forn_byhour[5:10]))
sheet_forein_by_period.write(1, 2, sum(forn_byhour[10:15]))
sheet_forein_by_period.write(1, 3, sum(forn_byhour[15:17]))
sheet_forein_by_period.write(1, 4, sum(forn_byhour[17:22]))
sheet_forein_by_period.write(1, 5, sum(forn_byhour[22:]))




cur=datetime.datetime.now()

output_file = './output/大兴机场_航班时间'
timelist = str(cur.year) + str(cur.month) + str(cur.day) 
output_file = output_file + timelist +'.xls'

workbook.save(output_file)
print('今天航班写入excel成功')

