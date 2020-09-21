from selenium import webdriver
import json
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import os
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import Select

# -*- coding: utf-8 -*-
import xlrd

file = open("config.txt", "r",encoding="utf-8")
config = json.loads(file.read())
file.close()
print("config.txt 讀取完成.")

delayTime = config["delayTime"]
物件追蹤表excel = config["物件追蹤表excel"]
身分證字號excel = config["身分證字號excel"]
選擇縣市url = config["選擇縣市url"]
選擇區域url = config["選擇區域url"]
主頁面url = config["主頁面url"]


caps = DesiredCapabilities.INTERNETEXPLORER
caps['ignoreProtectedModeSettings'] = True
caps['ignoreZoomSetting'] = True
driver = webdriver.Ie(executable_path='IEDriverServer.exe',capabilities=caps)
driver.get(主頁面url)


data = xlrd.open_workbook(物件追蹤表excel)
id_excel = xlrd.open_workbook(身分證字號excel)
data_json = {};

print("請手動登入帳號，完成後按任意鍵繼續.")
os.system("pause")

#讀取資料表
for n in range(len(data.sheet_names())):
    table = data.sheets()[n]
    
    #資料表之中欄位
    if n == 0:
        data_row = table.row_values(0)
    for i in range(1,table.nrows):
        #取出所有row欄位到data_json
        for j in range(len(table.row_values(i))-1):
            if table.row_values(i)[j] != "" :
                data_json[data_row[j]] = table.row_values(i)[j]
        id_table = id_excel.sheets()[i-1]
        身分證字號清單 = []
        身分證字號清單 = id_table.col_values(0)
        for 身分證字號 in 身分證字號清單:
            段建號 = data_json["段建號"]
            #選擇地區
            driver.get(選擇縣市url)
            縣市 = 段建號[0:段建號.find("市")+1].replace("\n", "")
            縣市下拉式選單 = Select(driver.find_element_by_id('City_ID'))
            縣市下拉式選單.select_by_visible_text(縣市)
            print("縣市 : " + 縣市)
            time.sleep(delayTime)
            區域 = 段建號[段建號.find("市")+1:段建號.find("區")+1]
            地區下拉式選單 = Select(driver.find_element_by_id('area_id'))
            地區下拉式選單.select_by_visible_text(區域)
            print("區域 : " + 區域)
            time.sleep(delayTime)
            段小段 = 段建號[段建號.find("區")+1:段建號.find("區")+5]
            建號 = 段建號[段建號.find("建號")-10:段建號.find("建號")-1]
            地號 = data_json["地號"][-11:-2]
            申請用途 = Select(driver.find_element_by_id('applyfor'))
            申請用途.select_by_visible_text("其他")
            print("段小段 : " + 段小段)
            print("建號 : " + 建號)
            print("地號 : " + 地號)
            print("身分證字號 : " + 身分證字號)
            print("==================================================")
            driver.find_element_by_name("MAXPAGE").send_keys("\b\b\b1");
            driver.find_element_by_id("INPUT_011").clear()
            driver.find_element_by_id("INPUT_011").send_keys(段小段)
            driver.find_element_by_id("INPUT_013").send_keys(地號)
            driver.find_element_by_id("INPUT_014").send_keys(建號)
            driver.find_element_by_name("INPUT_021").send_keys(Keys.SPACE)
            driver.find_element_by_name("INPUT_015").send_keys(身分證字號)
            driver.find_element_by_name("btnnew").send_keys(Keys.ENTER)
            time.sleep(delayTime)
            driver.find_element_by_name("btnsend").send_keys(Keys.ENTER)
            time.sleep(delayTime)
os.system("pause")

