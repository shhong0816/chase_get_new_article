import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from bs4 import BeautifulSoup
import time
from openpyxl import Workbook
from openpyxl import load_workbook

print("START")
#--------------크롬실행-----------------
#Options for chromedriver
#headless
options = webdriver.ChromeOptions()
#options.add_argument('headless')
options.add_argument("disable-gpu")
#useragent
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36")
#chrome start
browser = webdriver.Chrome(options=options)
browser.maximize_window()
#-------------------Make Frame----------------
df_science = pd.DataFrame({'Title' : [], 'url' : []})##요약도 해야할듯??, category(tag)로도 구별 needed
df_naver_news = pd.DataFrame({'Title' : [], 'url' : []})
df_sciencetimes = pd.DataFrame({'Title' : [], 'category' : [], 'abstract' : [], 'url' : []})
###--------------------GET_DATA,put to dataframe------------------------------------------------
###SCIENCE###
browser.get("https://www.science.org/toc/science/current")
#browser.get("https://www.science.org/toc/science/376/6600")
df_science.loc["Publish date"] = browser.find_element_by_xpath('//*[@id="pb-page-content"]/div/div/main/section[1]/div/div[2]/div[2]/div[2]/div/ul/li[3]/span').text
#Check special paper - xpath 달라짐
try:
    if browser.find_element_by_xpath('//*[@id="pb-page-content"]/div/div/main/section[2]/div/div[1]/div[1]/div/div/span').text.upper() == "SPECIAL ISSUE":
        is_Special = True
    else:
        is_Special = False
except:
    is_Special = False
##Find data 시작
num=1
for i in range(2,7):
    for j in range(1,40):
        if is_Special == True:
            try:
                Title = browser.find_element_by_xpath('//*[@id="pb-page-content"]/div/div/main/section[2]/div/div[1]/div[2]/div[2]/div/section['+str(i)+']/div['+str(j)+']/div/div[1]/div[1]/h3/a').text
                url = browser.find_element_by_xpath('//*[@id="pb-page-content"]/div/div/main/section[2]/div/div[1]/div[2]/div[2]/div/section['+str(i)+']/div['+str(j)+']/div/div[1]/div[1]/h3/a').get_attribute('href')
                df_science.loc[str(num)] = [Title,url]
                num = num+1
            except:
                continue
        else:
            try:
                Title = browser.find_element_by_xpath('//*[@id="pb-page-content"]/div/div/main/section[2]/div/div[1]/div[1]/div[2]/div/section['+str(i)+']/div['+str(j)+']/div/div[1]/div[1]/h3/a').text
                url = browser.find_element_by_xpath('//*[@id="pb-page-content"]/div/div/main/section[2]/div/div[1]/div[1]/div[2]/div/section['+str(i)+']/div['+str(j)+']/div/div[1]/div[1]/h3/a').get_attribute('href')
                df_science.loc[str(num)] = [Title,url]
                num = num+1
            except:
                continue
print("science fin")

##사이언스타임즈###
num = 1
df_science.loc["Publish date"] = "최근 50개 저장 without category"
for page in range(5):
    browser.get('https://www.sciencetimes.co.kr/category/sci-tech/page/'+str(page+1)+'/')
    for i in range(10):
        Title = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/section/div[1]/ul/li['+str(i+1)+']/div/div/div/a/strong').text
        #category xpath방식이 2개네...
        try:
            try:
                category = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/section/div[1]/ul/li['+str(i+1)+']/div/div/div/span').text
            except:
                category = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/section/div[1]/ul/li['+str(i+1)+']/div/div/div/a[1]').text
        except:
            category = 'none'
        abstract = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/section/div[1]/ul/li['+str(i+1)+']/div/div/div/a/p').text
        url = browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/section/div[1]/ul/li['+str(i+1)+']/div/div/div/a').get_attribute('href')
        df_sciencetimes.loc[str(num)] = [Title,category,abstract,url]
        num = num+1
df_sciencetimes
print("sciencetimes_fin")
###--------------------Compare_DATA with reference data------------------------------------------------
print("Comparing started")
###SCIENCE###
df_science_ordinary = pd.read_excel("./Reference_Data/science.xlsx",index_col=0,dtype='object')
compare_data_1 = df_science_ordinary.iloc[0][0]##date 비교
compare_data_2 = df_science_ordinary.iloc[1][0]##Title 2개 비교
compare_data_3 = df_science_ordinary.iloc[2][0]
if compare_data_1 == df_science.iloc[0][0] and compare_data_2 == df_science.iloc[1][0] and compare_data_3 == df_science.iloc[2][0]:
    science_is_same = True
else:
    science_is_same = False
##사이언스타임즈###
df_sciencetimes_ordinary = pd.read_excel("./Reference_Data/sciencetimes.xlsx",index_col=0,dtype='object')
compare_data_1 = df_sciencetimes_ordinary.iloc[0][0]
compare_data_2 = df_sciencetimes_ordinary.iloc[1][0]##Title 3개 비교
compare_data_3 = df_sciencetimes_ordinary.iloc[2][0]
if compare_data_1 == df_sciencetimes.iloc[0][0] and compare_data_2 == df_sciencetimes.iloc[1][0] and compare_data_3 == df_sciencetimes.iloc[2][0]:
    sciencetimes_is_same = True
else:
    sciencetimes_is_same = False
###--------------------Save_DATA at reference data------------------------------------------------
print("Making started")
###SCIENCE###
if science_is_same == False:
    with pd.ExcelWriter("./Reference_Data/science.xlsx") as writer:
        df_science.to_excel(writer,sheet_name = 'science')
##사이언스타임즈###
if sciencetimes_is_same == False:
    with pd.ExcelWriter("./Reference_Data/sciencetimes.xlsx") as writer:
        df_sciencetimes.to_excel(writer,sheet_name = 'sciencetimes')


###-----------------Making_Excel통합본--------------------------------------
###SCIENCE###
##사이언스타임즈###
df_sciencetimes_ordinary = pd.read_excel("./Reference_Data/sciencetimes.xlsx",index_col=0,dtype='object')
df_sciencetimes_final = df_sciencetimes_ordinary.drop(df_sciencetimes_ordinary[(df_sciencetimes_ordinary['category'] != '과학기술')].index)##numbering에서 제거만 됨.

###MAKING EXCEL####
with pd.ExcelWriter('News.xlsx') as writer:
    df_science.to_excel(writer, sheet_name='science')
    df_sciencetimes_final.to_excel(writer, sheet_name='sciencetimes')

###--------------------------decorate excel------------------------------
wb = load_workbook("News.xlsx")
##science##
ws_science = wb["science"]
if science_is_same == True:
    ws_science.sheet_properties.tabColor = "ffffff"
else:
    ws_science.sheet_properties.tabColor = "83dcb7"
##사이언스타임즈###
ws_sciencetimes = wb["sciencetimes"]
if sciencetimes_is_same == True:
    ws_sciencetimes.sheet_properties.tabColor = "ffffff"
else:##New contents
    ws_sciencetimes.sheet_properties.tabColor = "83dcb7"






browser.quit()
wb.save("News.xlsx")
print("FIN")
