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
sheet1 = data.sheet_by_index(0)
data1 = copy(data)
sheet2 = data1.get_sheet('sheet1')
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
for it in md5_list:
    count=count+1
    driver2.get('https://s.threatbook.cn/')
    input=driver2.find_element_by_xpath('//input')
    input.send_keys(it)
    button=driver2.find_element_by_class_name('search-btn')
    button.click()
    input.clear()
    time.sleep(4)#结果为null时停止4秒钟来判断结果
    if(isElementExist('sample-page__nodata')):
        print(str(count)+': null')
        sheet2.write(count-1,1,'null')
        sheet2.write(count-1,2,'null')
        sheet2.write(count-1,3,'null')
        sheet2.write(count-1,4,'null')
        sheet2.write(count-1,5,'null')
        print('-----')
        continue
    elif(isElementExist('no-authority-tip')):
        print(str(count)+': null')
        sheet2.write(count - 1, 1, 'null')
        sheet2.write(count - 1, 2, 'null')
        sheet2.write(count - 1, 3, 'null')
        sheet2.write(count - 1, 4, 'null')
        sheet2.write(count - 1, 5, 'null')
        print('-----')
        continue
    elif(isElementExist('sample-link')):
        box = driver2.find_element_by_class_name('table-box')
        td = box.find_element_by_xpath('//td')
        #print('td', td.get_attribute('class'))
        a = td.find_element_by_xpath("//a[@class='sample-link']")
        url = a.get_attribute('href')
        driver2.get(url)
    time.sleep(15)  # 当有结果出现时停止15秒进行获取
    app_mountpoint = driver2.find_element_by_id('app_mountpoint')
    global_ = app_mountpoint.find_element_by_class_name('global-drop-box.report-global-upload')
    send_box = global_.find_element_by_class_name('sandbox-info__basic')
    # send_box_item_labels=send_box.find_elements_by_class_name('sandbox-info__item-label')
    send_box_items_value = send_box.find_elements_by_class_name('sandbox-info__item-value.ellipsis')
    # print(send_box_item_labels)
    # sandbox_info_item_label=send_box_items[3].find_element_by_class_name('sandbox-info__item-label')
    # print(send_box_items[3])
    # sandbox_info_item_ellipsis=send_box_items[3].find_element_by_class_name('sandbox-info__item-value.ellipsis')
    # print('label',sandbox_info_item_label.text)

    score = driver2.find_element_by_class_name('sandbox-info__score')
    tags = driver2.find_element_by_xpath("//div[@class='sandbox-info__tags']")
    spans = tags.find_elements_by_tag_name('span')
    sheet2.write(count-1,1,send_box_items_value[3].text)
    print(str(count) + ':提交时间' + send_box_items_value[3].text + '   score: ' + score.text)
    s=re.findall("\d+",score.text)
    print('标签:')
    string=""
    for span in spans:
        string=string+span.text+"/"
    sheet2.write(count-1,5,string)
    print('威胁程度:')
    if len(s)>0:
        s1=s[0]
        so=int(s1,10)
        print(so)
        sheet2.write(count-1,2,so)
        if so>=60:
          print('恶意')
          sheet2.write(count-1,3,'恶意')
        elif so>=30:
            print('可疑')
            sheet2.write(count - 1, 3, '可疑')
        else:
            print('安全')
            sheet2.write(count - 1, 3, '安全')
    else:
        sheet2.write(count-1,2,'null')
        sheet2.write(count-1,3,'null')
    if isElementExist('threat-intelligence__none-td'):
        ioc = driver2.find_element_by_class_name('threat-intelligence__none-td')
        print('无有ioc: ',ioc.text)
        sheet2.write(count-1,4,'无ioc')
    else:
        print('有ioc')
        sheet2.write(count-1,4,'有ioc')
    print('-----')
data1.save('D:/pachong/md5_WB_res.xls')