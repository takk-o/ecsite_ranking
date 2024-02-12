# ecsite_ranking

## Overview
### rt_ranking.py
楽天市場のジャンルID指定により、その配下のジャンルについて、上位30位ランキングデータを取得してExcelに登録する。
### ys_ranking.py
YAHOO!ショッピングのカテゴリーID指定により、その配下のカテゴリーについて、上位10位ランキングデータを取得してExcelに登録する。  
Selenium利用(SEL_USE=True)により上位30位に拡張可（動作環境により、sleep時間や、解像度の調整が必要）

## Requirements
- Python 3.11.7
- requests 2.31.0
- openpyxl 3.1.2
- python-dotenv 1.0.1
- bs4 0.0.2
- selenium 4.17.2

## Usage
### rt_ranking.py
1. rt_base.xlsxの「ジャンルコード」に確認したいジャンルコードを登録("D2"セル)
1. rt_ranking.pyを起動
1. [当日日付]フォルダに、Ranking_[時刻].xlsxファイルが作成される
### ys_ranking.py
1. ys_base10.xlsxの「カテゴリーコード」に確認したいカテゴリーコードを登録("D2"セル)
1. ys_ranking.pyを起動
1. [当日日付]フォルダに、Ranking_[時刻].xlsxファイルが作成される

## Author
- takk-o
- Mail : ynurmj5e@gmail.com
