import time
import xlrd
import xlwt
from xlutils.copy import copy
import random
from selenium import webdriver

data=xlrd.open_workbook('md5.xls')
sheet1=data.sheet_by_index(0)
data1=copy(data)
sheet2=data1.get_sheet('sheet1')
md5_list=sheet1.col_values(0)

#driverVT检索
driver = webdriver.Chrome('C:/Users/zouyingtong/Downloads/chromedriver.exe')  # Optional argument, if not specified will search path.
option = webdriver.ChromeOptions()
option.add_argument('headless')
browser = webdriver.Chrome(chrome_options=option)
count=0
for it in md5_list:
    count=count+1
    string_md5=it
    time.sleep(random.randint(0, 1))
    #VT部分的检索
    driver.get('https://www.virustotal.com/gui/file/'+string_md5+'/detection')##连接网站
    s=driver.find_element_by_id('file-view')
    shadowroot1 = driver.execute_script('return arguments[0].shadowRoot',s)
    sh2=shadowroot1.find_element_by_id('report')#首先进行百分比的查找
    shadowroot4=driver.execute_script('return arguments[0].shadowRoot.children[0]',sh2)
    shadowroot5=shadowroot4.find_element_by_class_name('veredict-widget')
    shadowroot6=driver.execute_script('return arguments[0].children[0].shadowRoot',shadowroot5)
    posi=shadowroot6.find_element_by_class_name('positives')
    total=shadowroot6.find_element_by_class_name('total')
    ##查找标签VT
    shadowroot3 = driver.execute_script('return arguments[0].children[0].shadowRoot', sh2)
    # sh_grey=driver.execute_script('return arguments[0].children[0].shadowRoot', sh2)
    time.sleep(1)
    grey = shadowroot3.find_element_by_class_name('small.grey.filled')
    contains_pe = driver.execute_script('return arguments[0].shadowRoot', grey)
    # 通用部分
    chips = contains_pe.find_elements_by_class_name('chip')
    res_string=""
    for it in chips:
        res_string=res_string+'/'+ it.text
    if (posi.is_displayed() == 0&total.is_displayed()==0):
        sheet2.write(count-1, 1, 'null')
        sheet2.write(count - 1, 2, 'null')
        print(str(count)+': no singal')
        continue
    text=posi.text
    if(total.text=='0'):
        sheet2.write(count-1,1,'null')
        print(str(count)+': null')
    elif(text=='0'):
        sheet2.write(count-1,1, '0'+total.text)
        print(str(count)+': 0'+total.text)
    else:
        sheet2.write(count-1,1, text + total.text)
       # sheet2.write(count - 1, 2, res_string)
        print(str(count)+': '+text+total.text)
    if(res_string==""):
        sheet2.write(count - 1, 2, 'null')
    elif(res_string!=""):
        sheet2.write(count - 1, 2, res_string)


data1.save('D:/pachong/')
