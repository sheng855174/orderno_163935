from selenium import webdriver
import json
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import os
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import Select
import xlrd
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy

def isElementExist(driver,element):
    flag=True
    try:
        driver.find_element_by_xpath(element)
        return flag
    except:
        flag=False
        return flag

file = open("config.txt", "r",encoding="utf-8")
config = json.loads(file.read())
file.close()
print("config.txt 讀取完成.")

delayTime = config["delayTime"]
物件追蹤表excel = config["物件追蹤表excel"]
身分證字號excel = config["身分證字號excel"]
輸出excel = config["輸出excel"]
選擇縣市url = config["選擇縣市url"]
主頁面url = config["主頁面url"]


caps = DesiredCapabilities.INTERNETEXPLORER
caps['ignoreProtectedModeSettings'] = True
caps['ignoreZoomSetting'] = True

driver = webdriver.Edge(executable_path='msedgedriver.exe')
driver.get(主頁面url)


data = xlrd.open_workbook(物件追蹤表excel)
id_excel = xlrd.open_workbook(身分證字號excel)
workbook = copy(open_workbook(物件追蹤表excel))

data_json = {};

print("請手動登入帳號，完成後按任意鍵繼續.")
os.system("pause")
driver.get(選擇縣市url)

身分證字號excel工作表索引 = 0
#讀取資料表
for n in range(len(data.sheet_names())):
    table = data.sheets()[n]
    sheet = workbook.get_sheet(n)
    #資料表之中欄位
    if n == 0:
        data_row = table.row_values(0)
    for i in range(1,table.nrows):
        #取出所有row欄位到data_json
        for j in range(len(table.row_values(i))):
            if table.row_values(i)[j] != "":
                data_json[data_row[j]] = table.row_values(i)[j]
        id_table = id_excel.sheets()[身分證字號excel工作表索引]
        身分證字號清單 = []
        身分證字號清單 = id_table.col_values(0)
        姓名 = table.row_values(i)[5]
        y = 1
        if 姓名!="" and 姓名!=None:
            for 身分證字號 in 身分證字號清單:
                print("總筆數 :" , table.nrows)
                print("現在正在嘗試excel第",i+1,"列")
                if y == 1:
                    段建號 = data_json["段建號"]
                    #選擇地區
                    time.sleep(delayTime)
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
                    start = 0
                    end = 0
                    for z in range(段建號.find("建號"),0,-1):
                        if 段建號[z].isdigit()==True:
                            end = z;
                            break
                    for z in range(end,0,-1):
                        if 段建號[z].isdigit()==False and 段建號[z]!="-":
                            start = z+1;
                            break
                    建號 = 段建號[start:end]
                    start = 0
                    end = 0
                    for z in range(len(data_json["地號"])-1,0,-1):
                        if data_json["地號"][z].isdigit()==True:
                            end = z;
                            break
                    for z in range(end,0,-1):
                        if data_json["地號"][z].isdigit()==False and data_json["地號"][z]!="-":
                            start = z+1;
                            break
                    地號 = data_json["地號"][start:end]
                    申請用途 = Select(driver.find_element_by_id('applyfor'))
                    申請用途.select_by_visible_text("購屋、貸款使用")
                    print("段小段 : " + 段小段)
                    print("建號 : " + 建號)
                    print("地號 : " + 地號)
                    print("身分證字號 : " + 身分證字號)
                    driver.find_element_by_name("MAXPAGE").send_keys("\b\b\b1");
                    driver.find_element_by_id("INPUT_011").clear()
                    driver.find_element_by_id("INPUT_011").send_keys(段小段)
                    driver.find_element_by_id("INPUT_013").send_keys(地號)
                    driver.find_element_by_id("INPUT_014").send_keys(建號)
                    driver.find_element_by_name("INPUT_021").send_keys(Keys.SPACE)
                    driver.find_element_by_name("INPUT_015").send_keys(身分證字號)
                else :
                    driver.find_element_by_name("INPUT_015").clear()
                    driver.find_element_by_name("INPUT_015").send_keys(身分證字號)
                
                driver.find_element_by_name("btnnew").send_keys(Keys.ENTER)
                time.sleep(delayTime)
                driver.find_element_by_name("btnsend").send_keys(Keys.ENTER)
                time.sleep(delayTime)
                y = y+1
                #檢查失敗 or 成功
                error = isElementExist(driver,"//*[@id=\"ErrorMsgArea\"]")
                if error!=True :
                    print("資料正確!寫入excel")
                    sheet.write(i, 42, 身分證字號)
                    workbook.save('output.xls')
                    print("==================================================")
                    break
                else :
                    print("不正確的資料!!!!!!")
                print("==================================================")
            身分證字號excel工作表索引 = 身分證字號excel工作表索引 + 1



print("=========================================================")
print("執行完畢")
os.system("pause")
