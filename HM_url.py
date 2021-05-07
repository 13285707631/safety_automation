import time
import xlrd
import re
import xlwt
from xlutils.copy import copy
from selenium import webdriver
data = xlrd.open_workbook('md5.xls')
sheet1 = data.sheet_by_index(3)
data1 = copy(data)
sheet2 = data1.get_sheet('sheet4')
md5_list = sheet1.col_values(0)
option2 = webdriver.ChromeOptions()
option2.add_argument(r'user-data-dir=C:\Users\zouyingtong\AppData\Local\Google\Chrome\User Data1')
driver2 = webdriver.Chrome(options=option2)
count=0
def isElementExist(class_name):
    try:
        driver2.find_element_by_class_name(class_name)
        return True
    except:
        return False
def hasA(element):
    try:
        element.find_element_by_tag_name('a')
        return True
    except:
        return False
for it in md5_list:
    time.sleep(2)
    count=count+1
    driver2.get('https://sc.b.qihoo.net/domain/assuredshippings.com')
    time.sleep(2)
    div2=driver2.find_element_by_class_name('top-input.el-input.el-input--small.el-input--prefix.el-input--suffix')
    input=div2.find_element_by_class_name('el-input__inner')
    input.clear()
    input.send_keys(it)
    button=driver2.find_element_by_class_name('el-input__icon.el-icon-search').click()
    input.clear()
    time.sleep(5)
    if(isElementExist('logo-box')):
        div=driver2.find_element_by_class_name('logo-box')
        img = div.find_element_by_tag_name('img')
    elif(isElementExist('levelLogo')):
        div = driver2.find_element_by_class_name('levelLogo')
        img = div.find_element_by_tag_name('img')
    text=img.get_attribute('src')
    #print(text)
    #print(text)
    if(text=='https://s5.qhres2.com/static/eab58e5d2150a5d2.svg'):
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
data1.save('D:/pachong/HM_url.xls')