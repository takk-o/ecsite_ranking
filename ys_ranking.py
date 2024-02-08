# YAHOO!ショッピングのカテゴリーID指定により、その配下のカテゴリーについて、
# 上位10位ランキングデータを取得してExcelに登録する。
# （ページの構成により、レビュー数／アベレージがズレるケース有り）
# Selenium利用により上位30位に拡張可（要コード調整）

import requests
from bs4 import BeautifulSoup
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from time import sleep

from api import ys_api
from exl import exl_wb

SEL_USE = False

# 入力されたカテゴリーIDに従属するID全てをリスト化
ys = ys_api()

while True:
    try:
        id = input('カテゴリーID: ')
        break
    except ValueError:
        pass
ids = [id]
ys.out_ids = []
ys.list_get_ids(ids)
codes = ys.out_ids

del ys

if SEL_USE:
    # Selenium起動準備
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    options.add_argument('--window-size=3200,1800') 
    options.add_argument('--user-data-dir=/Users/ynurmj5e/Library/Application Support/Google/Chrome/Default')
    driver = webdriver.Chrome(options=options)

# Excelファイル出力準備
wb = exl_wb('output', 'SpreadSheet.xlsx')
if SEL_USE:
    wb.open_workbook(30)
else:
    wb.open_workbook()

# メイン処理
for code in codes:
    # ランキングページ取得
    url = f'https://shopping.yahoo.co.jp/categoryranking/{code}/list?sc_i=shopping-pc-web-list-ranking-nrwcgt-slctp_md-ranking'
    
    if SEL_USE:
        # Selenium操作
        try:
            driver.get(url)
            sleep(1)
            driver.execute_script("window.scrollTo(0, 2500);")      # 「もっと見る」ボタンが表示されるまでスクロール
            sleep(1)
            driver.execute_script('arguments[0].click();', driver.find_element(By.CLASS_NAME, "button-text"))
            sleep(1)
        except NoSuchElementException:                              # 「もっと見る」ボタンの無いページスキップ
            pass

        page = BeautifulSoup(driver.page_source, 'html.parser')
    else:
        page = BeautifulSoup(requests.get(url).content, 'html.parser')

    # パンくずリスト取得
    categories = page.find_all('li', class_ = 'Breadcrumb__item')
    l_cat = []
    for category in categories:
        l_cat.append(category.text)

    # ランキング情報取得
    ranks = page.find_all('span', class_ = 'rank-text')
    reviews = page.find_all('span', class_ = 'Review__count____2_0_101 Review__count--hasBrackets____2_0_101 review-count')
    averages = page.find_all('span', class_ = 'Review__average____2_0_101 review-average')
    prices = page.find_all('span', class_ = 'price-number')
    names = page.find_all('span', class_ = 'name-text')
    images = page.find_all('img', class_ = 'image')

    # Excel操作
    ws = wb.copy_ws()
    ws.title = code
    ws['C2'].hyperlink = url

    # パンくずリスト登録
    for j in range(len(l_cat)):
        ws.cell(row=6, column=j + 2, value=l_cat[j])

    # ランキング情報登録
    for i in range(len(reviews)):
        ws.cell(row=i + 9, column=3, value=reviews[i].text)
        ws.cell(row=i + 9, column=4, value=float(averages[i].text))
        ws.cell(row=i + 9, column=5, value=int(prices[i].text.replace(',', '')))
        ws.cell(row=i + 9, column=6, value=names[i].text)
        img_url = images[i].get('src')
        ws.cell(row=i + 9, column=7).hyperlink = img_url

# Excel終了
wb.save_workbook()

if SEL_USE:
    # Seleniumu終了
    driver.close()