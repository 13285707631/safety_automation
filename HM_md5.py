import time
import xlrd
import re
import xlwt
from xlutils.copy import copy
from selenium import webdriver
data = xlrd.open_workbook('md5.xls')
sheet1 = data.sheet_by_index(0)
data1 = copy(data)
sheet2 = data1.get_sheet('sheet1')
md5_list = sheet1.col_values(0)
option2 = webdriver.ChromeOptions()
option2.add_argument(r'user-data-dir=C:\Users\zouyingtong\AppData\Local\Google\Chrome\User Data1')
driver2 = webdriver.Chrome(options=option2)
count=0
for it in md5_list:
    count=count+1
    driver2.get('https://sc.b.qihoo.net/file/'+it)
    time.sleep(2)
    info=driver2.find_element_by_class_name('levelInfo')
    print(info.text[:2])
    sheet2.write(count-1,1,info.text[:2])
data1.save('D:/pachong/HM_md5.xls')