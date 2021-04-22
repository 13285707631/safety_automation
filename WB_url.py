import time
import xlrd
import re
import xlwt
from xlutils.copy import copy
import random
import datetime
from selenium import webdriver
import selenium

data = xlrd.open_workbook('md5.xls')
sheet1 = data.sheet_by_index(3)
data1 = copy(data)
sheet2 = data1.get_sheet('sheet4')
md5_list = sheet1.col_values(0)
driver2 = webdriver.Chrome('C:/Users/zouyingtong/Downloads/chromedriver.exe')  # Optional argument, if not specified will search path.
option2 = webdriver.ChromeOptions()
option2.add_argument('headless')
count = 0
browser2 = webdriver.Chrome(chrome_options=option2)
#for it in md5_list:a
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
    count=count+1
    driver2.get('https://s.threatbook.cn/')
    time.sleep(3)
    input=driver2.find_element_by_xpath('//input')
    input.clear()
    input.send_keys(it)
    #print(it)
    button=driver2.find_element_by_class_name('search-btn')
    button.click()
    time.sleep(13)
    if(isElementExist('sandbox-info__result')):
        sandbox_info_result = driver2.find_element_by_class_name('sandbox-info__result')
        span = sandbox_info_result.find_element_by_tag_name('span')
        res = span.text[-2:]
        print(res)
        sheet2.write(count-1,1,res)
        item_value = driver2.find_elements_by_xpath("//div[@class='sandbox-info__item-value ellipsis']")
        if(len(item_value)>0):
            n=2
            for i in item_value:
                if(hasA(i)):
                    a=i.find_element_by_tag_name('a')
                    print(a.text)
                    sheet2.write(count-1,n,a.text)
                    n=n+1
                else:
                    sheet2.write(count-1,n,i.text)
                    print(i.text)
                    n=n+1
        # data1.save('D:/pachong/md5_WB_res.xls')
    #data1.save('D:/pachong/md5_WB_res.xls')
    else:
        print('null')
data1.save('D:/pachong/WB_url.xls')