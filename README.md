# ecsite_ranking

## Overview
### rt_ranking.py
楽天市場のジャンルID指定により、その配下のジャンルについて、上位30位ランキングデータを取得してExcelに登録する。
### ys_ranking.py
YAHOO!ショッピングのカテゴリーID指定により、その配下のカテゴリーについて、上位10位ランキングデータを取得してExcelに登録する。  
Selenium利用(SEL_USE=True)により上位30位に拡張可（動作環境により、sleep時間や、解像度の調整が必要）

## Requirements
> 下記バージョンは `requirements.txt` の内容と同期しています。依存関係を更新した場合は本セクションも合わせて更新してください。

- Python 3.12
- requests 2.34.2
- openpyxl 3.1.5
- python-dotenv 1.2.2
- bs4 0.0.2
- selenium 4.23.1

## Usage
### rt_ranking.py
1. rt_base.xlsxの「ジャンルコード」に確認したいジャンルコードを登録("D2"セル)
1. rt_ranking.pyを起動
1. [当日日付]フォルダに、Ranking_[時刻].xlsxファイルが作成される
### ys_ranking.py
1. ys_base10.xlsxの「カテゴリーコード」に確認したいカテゴリーコードを登録("D2"セル)
1. ys_ranking.pyを起動
1. [当日日付]フォルダに、Ranking_[時刻].xlsxファイルが作成される
### ys_ranking.py(Selenium利用)
1. ys_base30.xlsxの「カテゴリーコード」に確認したいカテゴリーコードを登録("D2"セル)
1. ys_ranking.pyの SEL_USE = True に変更 
1. ys_ranking.pyを起動
1. [当日日付]フォルダに、Ranking_[時刻].xlsxファイルが作成される

## Author
- takk-o
- Mail : ynurmj5e@gmail.com

## Maintenance status
このツールは現在アクティブなメンテナンスを行っていません。
2026年7月時点で動作確認を試みたところ、楽天APIの認証エラー(401)が発生しました。
`.env`（APIキー）が未設定だったことが原因の可能性が高く、サイト側API仕様変更による不具合かは未検証です。
再検証時は楽天ウェブサービスでアプリID再取得のうえ `.env` にキーを設定してください。

セキュリティアラート対応として、2026年7月にurllib3/idna/python-dotenv/requestsを最新化済みです（Dependabot 7件解消）。
