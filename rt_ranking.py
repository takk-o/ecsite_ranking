# 楽天市場のジャンルID指定により、その配下のジャンルについて、
# 上位30位ランキングデータを取得してExcelに登録する。

from datetime import datetime
from pathlib import Path
import sys
from time import sleep

from api import rt_api
from exl import exl_wb

# 日時の取得
fd_date = datetime.now().strftime('%Y%m%d')
fd_time = datetime.now().strftime("%H%M%S")
# 親ディレクトリパスの取得
fd_path = Path(sys.argv[0]).parent
# Excelファイル準備
wb = exl_wb(fd_path.joinpath(fd_date), f'Ranking_{fd_time}.xlsx')
ws = wb.read_workbook(fd_path, 'rt_base.xlsx')
# カテゴリーコード取得
id = ws['D2'].value

# 取得したジャンルIDに従属するID全てをリスト化
rt = rt_api()

# while True:
#     try:
#         id = input('ジャンルID: ')
#         break
#     except ValueError:
#         pass
ids = [int(id)]
rt.out_ids = []
rt.list_get_ids(ids)
codes = rt.out_ids

# メイン処理
for code in codes:

    # パンくずリスト取得
    l_cat = rt.get_parent_names(code)
    l_cat.append(rt.get_current_name(code))

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
    ws['D2'] = int(id)
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

del rt

# # Excel終了
wb.save_workbook()
