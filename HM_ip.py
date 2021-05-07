import time
import xlrd
import re
import xlwt
from xlutils.copy import copy
from selenium import webdriver
data = xlrd.open_workbook('md5.xls')
sheet1 = data.sheet_by_index(1)
data1 = copy(data)
sheet2 = data1.get_sheet('sheet2')
md5_list = sheet1.col_values(0)
option2 = webdriver.ChromeOptions()
option2.add_argument(r'user-data-dir=C:\Users\zouyingtong\AppData\Local\Google\Chrome\User Data1')
driver2 = webdriver.Chrome(options=option2)
count=0
for it in md5_list:
    count=count+1
    driver2.get('https://sc.b.qihoo.net/ip/'+it)
    time.sleep(3)
    div=driver2.find_element_by_class_name('logo-box')
    img=div.find_element_by_tag_name('img')
    text=img.get_attribute('src')
    #print(text)
    if(text=='https://s7.qhres2.com/static/690caa1e021e7040.svg'):
        print("安全")
       # print("1")
        sheet2.write(count-1,1,"安全")
    elif(text=='https://s7.qhres2.com/static/f33b5998c9d32e14.svg'):
        print('危险')
        sheet2.write(count - 1, 1, "危险")
    elif (text == 'https://s7.qhres2.com/static/f2b955340cb1a4da.svg'):
        print('普通')
        sheet2.write(count - 1, 1, "普通")
    elif (text == 'https://s6.qhres2.com/static/bda7ac71808f75d7.svg'):
        print('未知')
        sheet2.write(count - 1, 1, "未知")
data1.save('D:/pachong/HM_ip.xls')