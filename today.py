from logging import debug
import pandas as pd
import numpy as np
from pandas.core.frame import DataFrame
import streamlit as st
import plotly.figure_factory as ff
import plotly.graph_objects as go
import openpyxl
import math

import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

st.set_page_config(page_title='受注内容/日')
st.markdown('#### 受注内容/日')

#小数点以下１ケタ
pd.options.display.float_format = '{:.2f}'.format

#current working dir
cwd = os.path.dirname(__file__)

#**********************gdriveからエクセルファイルのダウンロード・df化
fname ='kita79j'

# Google Drive APIを使用するための認証情報を取得する
creds_dict = st.secrets["gcp_service_account"]
creds = service_account.Credentials.from_service_account_info(creds_dict)

# Drive APIのクライアントを作成する
#API名（ここでは"drive"）、APIのバージョン（ここでは"v3"）、および認証情報を指定
service = build("drive", "v3", credentials=creds)

# 指定したファイル名を持つファイルのIDを取得する
#Google Drive上のファイルを検索するためのクエリを指定して、ファイルの検索を実行します。
# この場合、ファイル名とMIMEタイプを指定しています。
file_name = f"{fname}.xlsx"
query = f"name='{file_name}' and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
#指定されたファイルのメディアを取得
results = service.files().list(q=query).execute()
items = results.get("files", [])

if not items:
    st.warning(f"No files found with name: {file_name}")
else:
    # ファイルをダウンロードする
    file_id = items[0]["id"]
    file = service.files().get(fileId=file_id).execute()
    file_content = service.files().get_media(fileId=file_id).execute()

    # ファイルを保存する
    file_path = os.path.join(cwd, 'data', file_name)
    with open(file_path, "wb") as f:
        f.write(file_content)


#************************ファイルのdf化・加工
# ***今期受注***
path_now = os.path.join(cwd, 'data', 'kita79j.xlsx')
df_now = pd.read_excel(
    path_now, sheet_name='受注委託移動在庫生産照会', \
        usecols=[3, 6, 8, 10, 14, 15, 16, 45]) #index　ナンバー不要　index_col=0

df_now.sort_values('受注日', ascending=False, inplace=True)
selected_day =st.selectbox(
    '受注日を選択',
    df_now['受注日'].unique()
)
df_selected_now = df_now[df_now['受注日'] == selected_day]

kh_sum = df_selected_now[df_selected_now['営業担当コード']==952]['金額'].sum()
st.write('星川合計')
st.write(kh_sum)

st.table(df_selected_now)



