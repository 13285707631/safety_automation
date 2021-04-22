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
def isElementExistClassName(class_name):
    try:
        driver2.find_element_by_class_name(class_name)
        return True
    except:
        return False
def isElementExistId(id):
    try:
        driver2.find_element_by_id(id)
        return True
    except:
        return False
driver2.get('https://www.virustotal.com/gui/url/f5dad51013a324487a9bef99c2acf44007fa1dfba083f18b14976379ba952029/detection')

for it in md5_list:
    count=count+1
    time.sleep(3)
    body=driver2.find_element_by_tag_name('body')
    shadowRoot1=driver2.execute_script('return arguments[0].children[0].shadowRoot',body)
    searchBar=shadowRoot1.find_element_by_id('searchbar')
    toolBar=shadowRoot1.find_element_by_id('toolbar')
    shadowRoot4=driver2.execute_script('return arguments[0].shadowRoot',toolBar)

    searchButton=shadowRoot4.find_element_by_id('searchIcon')

    shadowRoot2=driver2.execute_script('return arguments[0].shadowRoot',searchBar)
    searchInput=shadowRoot2.find_element_by_id('searchInput')
    shadowRoot3=driver2.execute_script('return arguments[0].shadowRoot',searchInput)
    input=shadowRoot3.find_element_by_id('input')
    input.clear()
    input.send_keys(it)
    button=searchButton.click()
    input.clear()
    time.sleep(10)
    if (isElementExistId('url-view')):
        s = driver2.find_element_by_id('url-view')
        shadowroot1 = driver2.execute_script('return arguments[0].shadowRoot', s)
        sh2 = shadowroot1.find_element_by_id('report')  # 首先进行百分比的查找
        shadowroot4 = driver2.execute_script('return arguments[0].shadowRoot.children[0]', sh2)
        shadowroot5 = shadowroot4.find_element_by_class_name('veredict-widget')
        shadowroot6 = driver2.execute_script('return arguments[0].children[0].shadowRoot', shadowroot5)
        posi = shadowroot6.find_element_by_class_name('positives')
        total = shadowroot6.find_element_by_class_name('total')
        shadowroot3 = driver2.execute_script('return arguments[0].children[0].shadowRoot', sh2)
        time.sleep(1)
        string = "" + posi.text + total.text
        sheet2.write(count - 1, 1, string)
        shadowroot3 = driver2.execute_script('return arguments[0].children[0].shadowRoot', sh2)
        time.sleep(1)
        grey = shadowroot3.find_element_by_class_name('small.grey.filled')
        contains_pe = driver2.execute_script('return arguments[0].shadowRoot', grey)
        # 通用部分
        chips = contains_pe.find_elements_by_class_name('chip')
        res_string = ""
        for it in chips:
            res_string = res_string + '/' + it.text
        sheet2.write(count - 1, 2, res_string)
        print(str(count) + ":" + string + "    " + res_string)
    else:
        sheet2.write(count - 1, 1, 'null')
        sheet2.write(count - 1, 2, 'null')
data1.save('D:/pachong/VT_url.xls')