# 楽天市場のジャンルID指定により、その配下のジャンルについて、
# 上位30位ランキングデータを取得してExcelに登録する。

import requests
from bs4 import BeautifulSoup
from time import sleep
from pprint import pprint

from api import rt_api
from exl import exl_wb

# 入力されたジャンルIDに従属するID全てをリスト化
rt = rt_api()

while True:
    try:
        id = input('ジャンルID: ')
        break
    except ValueError:
        pass
ids = [int(id)]
rt.out_ids = []
rt.list_get_ids(ids)
codes = rt.out_ids

# Excelファイル出力準備
wb = exl_wb('output', 'SpreadSheet.xlsx')
wb.open_workbook(30)

# メイン処理
for code in codes:

    # パンくずリスト取得
    l_cat = rt.get_parent_names(code)

    # ランキング情報取得
    ranks = []
    reviews = []
    averages = []
    prices = []
    names = []
    images = []
    for item in rt.get_ranking(code):
        ranks.append(item['Item']['rank'])
        reviews.append(item['Item']['reviewCount'])
        averages.append(item['Item']['reviewAverage'])
        prices.append(item['Item']['itemPrice'])
        names.append(item['Item']['itemName'])
        images.append(item['Item']['smallImageUrls'][0])
    sleep(0.2)

    # Excel操作
    ws = wb.copy_ws()
    ws.title = str(code)
    ws['B2'] = ""
    ws['B6'] = "トップ"
    
    # # パンくずリスト登録
    for j in range(len(l_cat)):
        ws.cell(row=6, column=j + 3, value=l_cat[j])

    # # ランキング情報登録
    if len(ranks) > 30:
        counter = 30
    else:
        counter = len(ranks)
    for i in range(counter):
        ws.cell(row=i + 9, column=2, value=f'{ranks[i]}位')
        ws.cell(row=i + 9, column=3, value=reviews[i])
        ws.cell(row=i + 9, column=4, value=float(averages[i]))
        ws.cell(row=i + 9, column=5, value=int(prices[i].replace(',', '')))
        ws.cell(row=i + 9, column=6, value=names[i])
        ws.cell(row=i + 9, column=7).hyperlink = images[0]['imageUrl']
        # ws.cell(row=i + 9, column=8).value = f'=_xlfn.image(G{i + 9})'

# # Excel終了
wb.save_workbook()

del rt