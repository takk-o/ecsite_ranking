# APIクラス
import requests
from requests.exceptions import RequestException, ConnectionError, HTTPError, Timeout

import settings

class ys_api:

    # 出力カテゴリーIDリスト
    def __init__(self):
        self.out_ids = []

    # 引数に指定したカテゴリーIDの子要素を返す
    def get_child_ids(self, id):
        # Yahoo!Shopping API カテゴリID取得

        # appid（必須）         string      Client ID（アプリケーションID）
        # output               json        レスポンス形式の指定             json : json形式
        # category_id（必須）	integer     カテゴリIDを指定                category_id = 1 : ルートカテゴリ(第1階層)

        YahooCategorySearchURL = 'https://shopping.yahooapis.jp/ShoppingWebService/V1/json/categorySearch'
        params = {}
        params['appid'] = settings.ys_app_id
        params['output'] = 'json'
        params['category_id'] = id
        try:
            results = requests.get(YahooCategorySearchURL, params)
            results.raise_for_status()
        except ConnectionError as ce:           # 接続エラー
            print("Connection Error:", ce)
            return []
        except HTTPError as he:                 # HTTPステータスエラー
            print("HTTP Error:", he)
            return []
        except Timeout as te:                   # タイムアウト
            print("Timeout Error:", te)
            return []
        except RequestException as re:          # その他エラー
            print("Error:", re)
            return []
        results = results.json()

        #　指定カテゴリーIDの子要素を取得
        child_elements = results['ResultSet']['0']['Result']['Categories']['Children']

        # 最終要素は[]
        if child_elements == []:
            return child_elements
        
        # 子要素の'Id'をリストとして返す
        child_ids = []
        for child_element in child_elements.items():
            if child_element[1] != 'Child':
                child_ids.append(child_element[1]['Id'])

        return child_ids

    # 引数に指定したカテゴリーIDリストに従属するID全てをリスト化して返す
    def list_get_ids(self, list_id):
        for id in list_id:
            self.out_ids.append(id)
            work_ids = []
            work_ids = self.get_child_ids(id)
            if work_ids != []:
                self.list_get_ids(work_ids)

class rt_api:
    
    # 出力ジャンルIDリスト
    def __init__(self):
        self.out_ids = []

    def genre_search(self, id):
        # Rakuten ジャンル検索API
        # applicationId（必須） string      アプリID
        # output               json        レスポンス形式                   json : json形式
        # genreId	           integer     ジャンルID                      category_id = 1 : ルートカテゴリ(第1階層)

        RakutenGenreSearchURL = 'https://app.rakuten.co.jp/services/api/IchibaGenre/Search/20140222'
        params = {}
        params['applicationId'] = settings.rt_app_id
        params['format'] = 'json'
        params['genreId'] = id

        try:
            results = requests.get(RakutenGenreSearchURL, params)
            results.raise_for_status()
        except ConnectionError as ce:           # 接続エラー
            print("Connection Error:", ce)
            return []
        except HTTPError as he:                 # HTTPステータスエラー
            print("HTTP Error:", he)
            return []
        except Timeout as te:                   # タイムアウト
            print("Timeout Error:", te)
            return []
        except RequestException as re:          # その他エラー
            print("Error:", re)
            return []
        return results.json()

    # 引数に指定したジャンルIDの子要素IDを返す
    def get_child_ids(self, id):
        results = self.genre_search(id)

        child_ids = []
        for child in results['children']:
            child_ids.append(child['child']['genreId'])
        return child_ids

    # 引数に指定したジャンルIDリストに従属するID全てをリスト化して返す
    def list_get_ids(self, list_id):
        for id in list_id:
            self.out_ids.append(id)
            work_ids = []
            work_ids = self.get_child_ids(id)
            if work_ids != []:
                self.list_get_ids(work_ids)

    # 引数に指定したジャンルIDの親要素名を返す
    def get_parent_names(self, id):
        results = self.genre_search(id)

        parent_names = []
        for parent in results['parents']:
            parent_names.append(parent['parent']['genreName'])
        return parent_names

    # 引数に指定したジャンルIDの名前を返す
    def get_current_name(self, id):
        results = self.genre_search(id)

        current_name = results['current']['genreName']
        return current_name

    def get_ranking(self, id):
        
        RakutenGenreSearchURL = 'https://app.rakuten.co.jp/services/api/IchibaItem/Ranking/20220601'
        params = {}
        params['applicationId'] = settings.rt_app_id
        params['elements'] = 'rank,reviewCount,reviewAverage,itemPrice,itemName,smallImageUrls'
        params['format'] = 'json'
        params['genreId'] = id

        try:
            results = requests.get(RakutenGenreSearchURL, params)
            results.raise_for_status()
        except ConnectionError as ce:           # 接続エラー
            print("Connection Error:", ce)
            return []
        except HTTPError as he:                 # HTTPステータスエラー
            print("HTTP Error:", he)
            return []
        except Timeout as te:                   # タイムアウト
            print("Timeout Error:", te)
            return []
        except RequestException as re:          # その他エラー
            print("Error:", re)
            return []
        results = results.json()['Items']

        return results
