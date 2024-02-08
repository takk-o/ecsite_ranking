from os import environ
from pathlib import Path
from dotenv import load_dotenv

# .envファイル読み込み
dotenv_path = Path.cwd().joinpath('.env')
load_dotenv(dotenv_path)
# APP_ID取得
ys_app_id = environ.get('YS_APP_ID')
rt_app_id = environ.get('RT_APP_ID')
