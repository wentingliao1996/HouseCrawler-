import os 
import pathlib
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import chromedriver_autoinstaller_fix
import warnings
import pathlib
import pandas as pd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from dotenv import load_dotenv
import time
from bs4 import BeautifulSoup
import pandas as pd
import re
import gspread
from google.oauth2.service_account import Credentials

warnings.filterwarnings("ignore")

def getyear(table_element):
    row = table_element.find_elements(By.TAG_NAME, 'tr')[4]
    cells = row.find_elements(By.TAG_NAME, 'td')
    row_data = [cell.text for cell in cells]
    if row_data[1]!='':
        year = float(row_data[1].split("年")[0])
    else:
        year=0
    return year
def getlosction(table_element):
    dontwant = ["新埔","湖口", "新豐"]
    row = table_element.find_elements(By.TAG_NAME, 'tr')[9]
    cells = row.find_elements(By.TAG_NAME, 'td')
    row_data = [cell.text for cell in cells]
    losction= row_data[0]
    for item in dontwant:
        if item in losction and '新竹' in losction:
            losction = losction.replace('新竹', '')
    return losction
def getmoney(table_element):
    row = table_element.find_elements(By.TAG_NAME, 'tr')[0]
    cells = row.find_elements(By.TAG_NAME, 'td')
    row_data = [cell.text for cell in cells]
    if "萬" in row_data[0] and row_data[0].strip()!='':
        money = float(row_data[0].split("萬")[0])
    else:
        money=0
    return int(money)
def getsquare(table_element):
    row = table_element.find_elements(By.TAG_NAME, 'tr')[1]
    cells = row.find_elements(By.TAG_NAME, 'td')
    row_data = [cell.text for cell in cells]
    square = row_data[0].split("坪")[0]
    if row_data[0].strip()!='':
        square=float(square)
    else:
        square=0
    return square
def gethold(table_element):
    row = table_element.find_elements(By.TAG_NAME, 'tr')[3]
    cells = row.find_elements(By.TAG_NAME, 'td')
    row_data = [cell.text for cell in cells]
    return row_data[-1]
def main(num):
    continuous_count = 0  # 連續超過10組的次數
    last_valid_row = None  # 上一個非空資料的索引
    
    ## 連結google sheet
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file("token.json", scopes=scopes)
    client = gspread.authorize(creds)
    sheet_id = "1mhNnsbtlpaU_UYieQ4JJoS7ryBUhRh28ioyTo3cCDQU"
    sheet = client.open_by_key(sheet_id)
    worksheet = sheet.get_worksheet(0)
    
    ## 爬蟲
    options = webdriver.ChromeOptions()
     
    current_path = str(pathlib.Path("__file__").parent.absolute())
     
    prefs = {
        "download.default_directory": current_path +"\RESULT",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(options=options)
    i =1
    while True:
        html = f"https://www.ycut.com.tw/ShareCase/CaseDetail.aspx?CaseID={num+i}"
        driver.get(html)
        WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="form1"]/div[3]/div[4]/div/table[1]'))
            )
        table_element = driver.find_element(By.XPATH, '//*[@id="form1"]/div[3]/div[4]/div/table[1]')
        # 解析表格內容
        try:
            money = getmoney(table_element)
        except:
            money=0
        # 如果 row_data['0'] 為空，則將 last_valid_row 設置為當前索引
        if money:
            last_valid_row = None
            location = getlosction(table_element)
            year = getyear(table_element)
            table2 = driver.find_element(By.XPATH, '//*[@id="form1"]/div[3]/div[4]/div/table[2]')
            square = getsquare(table2)
            hold = gethold(table2)
            if '新竹' in location and year<30 and money<2500 and square >0:
                name = driver.find_element(By.ID, 'lCaseName').text
                thisinfo = [name, html, square, hold]
                table_rows = table_element.find_elements(By.TAG_NAME, 'tr')
                for row in table_rows:
                    # 獲取每一行的單元格
                    cells = row.find_elements(By.TAG_NAME, 'td')
                    # 提取單元格文本，並添加到列表中
                    row_data = [cell.text for cell in cells]
                    thisinfo += row_data
                print(thisinfo)
                existing_rows = len(worksheet.get_all_values())
                # 添加新數據到下一行
                worksheet.append_row(thisinfo, value_input_option='RAW', insert_data_option='INSERT_ROWS', table_range=f'A{existing_rows + 1}')
                print("Data added successfully!")
        else:
            # 如果 row_data['0'] 不為空，則重置 last_valid_row
            last_valid_row = i
        
        # 如果 超過30組，則將 continuous_count 加1
        if not money and last_valid_row is not None:
            continuous_count += 1
            if continuous_count >= 30:
                print(f"Reached 10 consecutive empty rows. Pausing for 6 hours.")
                time.sleep(2 * 3600)  # 暫停 2 小時
                continuous_count = 0  # 重置計數器
                i = last_valid_row - 1  # 重置迴圈索引到上一個非空資料的索引
                
        else:
            continuous_count = 0  # 重置計數器
        time.sleep(3)
        i += 1