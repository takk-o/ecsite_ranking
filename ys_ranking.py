# YAHOO!ショッピングのカテゴリーID指定により、その配下のカテゴリーについて、
# 上位10位ランキングデータを取得してExcelに登録する。
# （ページの構成により、レビュー数／アベレージがズレるケース有り）
# Selenium利用(SEL_USE=True)により上位30位に拡張可
# （動作環境により、sleep時間や、解像度の調整が必要）

from datetime import datetime
from pathlib import Path
import sys
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from time import sleep

from api import ys_api
from exl import exl_wb

SEL_USE = True

# 日時の取得
fd_date = datetime.now().strftime('%Y%m%d')
fd_time = datetime.now().strftime("%H%M%S")
# 親ディレクトリパスの取得
fd_path = Path(sys.argv[0]).parent
# Excelファイル準備
wb = exl_wb(fd_path.joinpath(fd_date), f'Ranking_{fd_time}.xlsx')
if SEL_USE:
    ws = wb.read_workbook(fd_path, 'ys_base30.xlsx')
else:
    ws = wb.read_workbook(fd_path, 'ys_base10.xlsx')
# カテゴリーコード取得
id = ws['D2'].value

# 取得したカテゴリーIDに従属するID全てをリスト化
ys = ys_api()

# while True:
#     try:
#         id = input('カテゴリーID: ')
#         break
#     except ValueError:
#         pass
ids = [int(id)]
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

# メイン処理
for code in codes:
    # ランキングページ取得
    url = f'https://shopping.yahoo.co.jp/categoryranking/{code}/list?sc_i=shopping-pc-web-list-ranking-nrwcgt-slctp_md-ranking'
    
    if SEL_USE:
        # Selenium操作
        try:
            driver.get(url)
            sleep(1)
            driver.execute_script("window.scrollTo(0, 3000);")      # 「もっと見る」ボタンが表示されるまでスクロール
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
    # ranks = page.find_all('span', class_ = 'rank-text')
    # reviews = page.find_all('span', class_ = 'Review__count____2_0_101 Review__count--hasBrackets____2_0_101 review-count')
    # averages = page.find_all('span', class_ = 'Review__average____2_0_101 review-average')
    # prices = page.find_all('span', class_ = 'price-number')
    # names = page.find_all('span', class_ = 'name-text')
    # images = page.find_all('img', class_ = 'image')

    ranks, reviews, averages, prices, names, images = [], [], [], [], [], []
    if SEL_USE:
        num = 30
    else:
        num = 10
    space = page.new_tag('span')
    space.string = ''
    zero = page.new_tag('span')
    zero.string ='0'
    for i in range(num):
        if i <= 10:
            selector = f'div > div > div:nth-child(1) > ul > li > div > div:nth-child({i + 1}) > div > div > div.column-left > div > span.rank-text'
            element = page.select_one(selector=selector)
            if element:
                ranks.append(element)
            else:
                ranks.append(zero)
            selector = f'div > div > div:nth-child(1) > ul > li > div > div:nth-child({i + 1}) > div > div > div.column-middle-right > a > span > span.Review__count____2_0_101.Review__count--hasBrackets____2_0_101.review-count'
            element = page.select_one(selector=selector)
            if element:
                reviews.append(element)
            else:
                reviews.append(space)
            selector = f'div > div > div:nth-child(1) > ul > li > div > div:nth-child({i + 1}) > div > div > div.column-middle-right > a > span > span.Review__average____2_0_101.review-average'
            element = page.select_one(selector=selector)
            if element:
                averages.append(element)
            else:
                averages.append(zero)
            selector = f'div > div > div:nth-child(1) > ul > li > div > div:nth-child({i + 1}) > div > div > div.column-middle > div > p > span.price-number'
            element = page.select_one(selector=selector)
            if element:
                prices.append(element)
            else:
                prices.append(zero)
            selector = f'div > div > div:nth-child(1) > ul > li > div > div:nth-child({i + 1}) > div > div > div.column-middle > p > a > span'
            element = page.select_one(selector=selector)
            if element:
                names.append(element)
            else:
                names.append(space)
            selector = f'div > div > div:nth-child(1) > ul > li > div > div:nth-child({i + 1}) > div > div > div.column-middle-left > a > img'
            element = page.select_one(selector=selector)
            if element:
                images.append(element)
            else:
                images.append(space)
        else:
            selector = f'div > div > div > ul > li > div > div:nth-child({i - 9}) > div > div > div.column-left > div > span.rank-text'
            element = page.select_one(selector=selector)
            if element:
                ranks.append(element)
            else:
                ranks.append(zero)
            selector = f'div > div > div > ul > li > div > div:nth-child({i - 9}) > div > div > div.column-middle-right > a > span > span.Review__count____2_0_101.Review__count--hasBrackets____2_0_101.review-count'
            element = page.select_one(selector=selector)
            if element:
                reviews.append(element)
            else:
                reviews.append(space)
            selector = f'div > div > div > ul > li > div > div:nth-child({i - 9}) > div > div > div.column-middle-right > a > span > span.Review__average____2_0_101.review-average'
            element = page.select_one(selector=selector)
            if element:
                averages.append(element)
            else:
                averages.append(zero)
            selector = f'div > div > div > ul > li > div > div:nth-child({i - 9}) > div > div > div.column-middle > div.price > p > span.price-number'
            element = page.select_one(selector=selector)
            if element:
                prices.append(element)
            else:
                prices.append(zero)
            selector = f'div > div > div > ul > li > div > div:nth-child({i - 9}) > div > div > div.column-middle > p > a > span'
            element = page.select_one(selector=selector)
            if element:
                names.append(element)
            else:
                names.append(space)
            selector = f'div > div > div > ul > li > div > div:nth-child({i - 9}) > div > div > div.column-middle-left > a > img'
            element = page.select_one(selector=selector)
            if element:
                images.append(element)
            else:
                images.append(space)           

    # Excel操作
    ws = wb.copy_ws()
    ws.title = str(code)
    ws['D2'] = int(code)

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