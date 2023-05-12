from logging import debug
import pandas as pd
import numpy as np
from pandas.core.frame import DataFrame
import streamlit as st
import plotly.figure_factory as ff
import plotly.graph_objects as go
import openpyxl
# from streamlit.state.session_state import Value
import math

import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

st.set_page_config(page_title='売り上げ分析（星川/得意先別）')
st.markdown('#### 売り上げ分析（星川/得意先別）')

#小数点以下１ケタ
pd.options.display.float_format = '{:.1f}'.format

#current working dir
cwd = os.path.dirname(__file__)

#**********************gdriveからエクセルファイルのダウンロード・df化
fname_list = ['79s', 'kita79j', 'kita78j', '前期北日本j']
for fname in fname_list:
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
# ***今期出荷***
path_snow = os.path.join(cwd, 'data', '79s.xlsx')
df_snow = pd.read_excel(
    path_snow, sheet_name='受注委託移動在庫生産照会', usecols=[3, 6, 15, 16, 45]) #index　ナンバー不要　index_col=0


# ***ファイル読み込み 前期出荷***
df_slast = pd.read_excel(\
    '79期出荷ALL星川.xlsx', sheet_name='受注委託移動在庫生産照会', \
        usecols=[3, 6, 15, 16, 45]) #index　ナンバー不要　index_col=0

# ***今期受注***
path_jnow = os.path.join(cwd, 'data', 'kita79j.xlsx')
df_jnow = pd.read_excel(
    path_jnow, sheet_name='受注委託移動在庫生産照会', usecols=[3, 6, 15, 16, 45]) #index　ナンバー不要　index_col=0

# ***前期受注***
path_jlast = os.path.join(cwd, 'data', 'kita78j.xlsx')
df_jlast = pd.read_excel(
    path_jlast, sheet_name='受注委託移動在庫生産照会', usecols=[3, 6, 15, 16, 45]) #index　ナンバー不要　index_col=0

# ***前期受注年間***
path_jlast_full = os.path.join(cwd, 'data', '前期北日本j.xlsx')
df_jlast_full = pd.read_excel(
    path_jlast_full, sheet_name='受注委託移動在庫生産照会', usecols=[3, 6, 15, 16, 45])

# *** 出荷月、受注月列の追加***
df_snow['出荷月'] = df_snow['出荷日'].dt.month
df_snow['受注月'] = df_snow['受注日'].dt.month
df_slast['出荷月'] = df_slast['出荷日'].dt.month
df_slast['受注月'] = df_slast['受注日'].dt.month
df_jnow['出荷月'] = df_jnow['出荷日'].dt.month
df_jnow['受注月'] = df_jnow['受注日'].dt.month
df_jlast['出荷月'] = df_jlast['出荷日'].dt.month
df_jlast['受注月'] = df_jlast['受注日'].dt.month
df_jlast_full['出荷月'] = df_jlast_full['出荷日'].dt.month
df_jlast_full['受注月'] = df_jlast_full['受注日'].dt.month

# ***INT型への変更***
df_snow[['金額', '出荷月', '受注月']] = df_snow[[\
    '金額', '出荷月', '受注月']].fillna(0).astype('int64')
#fillna　０で空欄を埋める
df_slast[['金額', '出荷月', '受注月']] = df_slast[[\
    '金額', '出荷月', '受注月']].fillna(0).astype('int64')
#fillna　０で空欄を埋める

df_jnow[['金額', '出荷月', '受注月']] = df_jnow[[\
    '金額', '出荷月', '受注月']].fillna(0).astype('int64')
#fillna　０で空欄を埋める
df_jlast[['金額', '出荷月', '受注月']] = df_jlast[[\
    '金額', '出荷月', '受注月']].fillna(0).astype('int64')
#fillna　０で空欄を埋める
df_jlast_full[['金額', '出荷月', '受注月']] = df_jlast_full[[\
    '金額', '出荷月', '受注月']].fillna(0).astype('int64')
#fillna　０で空欄を埋める

#目標
target_list = [9000000, 10600000, 10300000, 7900000, 8600000, 9100000, \
          5500000, 6400000, 7100000, 8900000, 7500000,9100000] 

def tif():
    month_list = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    columns_list = ['受注/今期', '受注/前期', '受注/前期年間','対前年差', '対前年比']
    cust_list = [
        '㈱東京ｲﾝﾃﾘｱ 仙台港本店', '㈱東京ｲﾝﾃﾘｱ 福島店', '㈱東京ｲﾝﾃﾘｱ 郡山店', '㈱東京ｲﾝﾃﾘｱ いわき店', \
        '㈱東京ｲﾝﾃﾘｱ 山形店'
    ]

    jnow_list = []
    jlast_list = []
    jlast_full_list = []
    sales_diff_list = []
    sales_rate_list = []

    for cust in cust_list:
        df_jnow2 = df_jnow[df_jnow['得意先名']==cust]
        df_jlast2 = df_jlast[df_jlast['得意先名']==cust]
        df_jlast_full2 = df_jlast_full[df_jlast_full['得意先名']==cust]
        target_num = 0
        for month in month_list:
            target = target_list[target_num]
            jnow = df_jnow2[df_jnow2['受注月'].isin([month])]['金額'].sum()
            jlast = df_jlast2[df_jlast2['受注月'].isin([month])]['金額'].sum()
            jlast_full = df_jlast_full2[df_jlast_full2['受注月'].isin([month])]['金額'].sum()
            
            sales_diff = jnow - jlast
            sales_rate = f'{jnow / jlast * 100: 0.1f} %'

            jnow_list.append('{:,}'.format(jnow))
            jlast_list.append('{:,}'.format(jlast))
            jlast_full_list.append('{:,}'.format(jlast_full))

            sales_diff_list.append('{:,}'.format(sales_diff))
            sales_rate_list.append(sales_rate)

            target_num += 1

        df_month = pd.DataFrame(list(zip(\
            jnow_list, jlast_list, jlast_full_list, sales_diff_list, sales_rate_list)), \
                columns=columns_list, index=month_list)
        
        jnow_list = []
        jlast_list = []
        jlast_full_list = []
        sales_diff_list = []
        sales_rate_list = []
    

        #*****受注ベース可視化
        df_month2 = df_month.copy()

        #グラフ用にint化
        df_month2['受注/今期2'] = df_month2['受注/今期'].apply(lambda x: x.replace(',', '')).astype('int')
        df_month2['受注/前期2'] = df_month2['受注/前期'].apply(lambda x: x.replace(',', '')).astype('int')
        df_month2['受注/前期年間2'] = df_month2['受注/前期年間'].apply(lambda x: x.replace(',', '')).astype('int')

        st.write(f'受注ベース/売上: {cust}')
        #グラフを描くときの土台となるオブジェクト
        fig3 = go.Figure()
        #今期のグラフの追加
        for col in df_month2.columns[5: 8]:
            fig3.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_month2[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_month2[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig3.update_layout(
            title='月別',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig3, use_container_width=True)

       

        #*****累計 受注ベース可視化
        #グラフ用にint化
        df_month2['累計/受注/今期2'] = df_month2['受注/今期2'].cumsum()
        df_month2['累計/受注/前期2'] = df_month2['受注/前期2'].cumsum()
        df_month2['累計/受注/前期年間2'] = df_month2['受注/前期年間2'].cumsum()

        #table用にdiffとrate追加
        df_month2['累計/前年差'] = df_month2['累計/受注/今期2'] - df_month2['累計/受注/前期2']
        df_month2['累計/前年比'] = df_month2['累計/受注/今期2'] / df_month2['累計/受注/前期2']

        #グラフを描くときの土台となるオブジェクト
        fig4 = go.Figure()
        #今期のグラフの追加
        for col in df_month2.columns[8:11]:
            fig4.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_month2[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_month2[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig4.update_layout(
            title='累計',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig4, use_container_width=True) 

        with st.expander('詳細', expanded=False):
            col_list = ['受注/今期', '受注/前期', '対前年差', '対前年比', '累計/受注/今期2', '累計/受注/前期2',\
                        '累計/前年差', '累計/前年比']
            df_temp = df_month2[col_list]
            st.table(df_temp) 
def tif2():
    month_list = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    columns_list = ['受注/今期', '受注/前期', '受注/前期年間', '対前年差', '対前年比']
    cust_list = [
        '㈱東京ｲﾝﾃﾘｱ 下田店', '㈱東京ｲﾝﾃﾘｱ 秋田店', '㈱東京ｲﾝﾃﾘｱ 盛岡店',
        '㈱東京ｲﾝﾃﾘｱ 仙台泉店', '㈱東京ｲﾝﾃﾘｱ 仙台南店'
    ]

    jnow_list = []
    jlast_list = []
    jlast_full_list = []
    sales_diff_list = []
    sales_rate_list = []

    for cust in cust_list:
        df_jnow2 = df_jnow[df_jnow['得意先名']==cust]
        df_jlast2 = df_jlast[df_jlast['得意先名']==cust]
        df_jlast_full2 = df_jlast_full[df_jlast_full['得意先名']==cust]
        target_num = 0
        for month in month_list:
            target = target_list[target_num]
            jnow = df_jnow2[df_jnow2['受注月'].isin([month])]['金額'].sum()
            jlast = df_jlast2[df_jlast2['受注月'].isin([month])]['金額'].sum()
            jlast_full = df_jlast_full2[df_jlast_full2['受注月'].isin([month])]['金額'].sum()
            
            sales_diff = jnow - jlast
            sales_rate = f'{jnow / jlast * 100: 0.1f} %'

            jnow_list.append('{:,}'.format(jnow))
            jlast_list.append('{:,}'.format(jlast))
            jlast_full_list.append('{:,}'.format(jlast_full))

            sales_diff_list.append('{:,}'.format(sales_diff))
            sales_rate_list.append(sales_rate)

            target_num += 1

        df_month = pd.DataFrame(list(zip(\
            jnow_list, jlast_list, jlast_full_list, sales_diff_list, sales_rate_list)), \
                columns=columns_list, index=month_list)
        
        jnow_list = []
        jlast_list = []
        jlast_full_list = []
        sales_diff_list = []
        sales_rate_list = []
    

        #*****受注ベース可視化
        df_month2 = df_month.copy()

        #グラフ用にint化
        df_month2['受注/今期2'] = df_month2['受注/今期'].apply(lambda x: x.replace(',', '')).astype('int')
        df_month2['受注/前期2'] = df_month2['受注/前期'].apply(lambda x: x.replace(',', '')).astype('int')
        df_month2['受注/前期年間2'] = df_month2['受注/前期年間'].apply(lambda x: x.replace(',', '')).astype('int')

        st.write(f'受注ベース/売上: {cust}')
        #グラフを描くときの土台となるオブジェクト
        fig3 = go.Figure()
        #今期のグラフの追加
        for col in df_month2.columns[5: 8]:
            fig3.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_month2[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_month2[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig3.update_layout(
            title='月別',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig3, use_container_width=True)

       

        #*****累計 受注ベース可視化
        #グラフ用にint化
        df_month2['累計/受注/今期2'] = df_month2['受注/今期2'].cumsum()
        df_month2['累計/受注/前期2'] = df_month2['受注/前期2'].cumsum()
        df_month2['累計/受注/前期年間2'] = df_month2['受注/前期年間2'].cumsum()

        #table用にdiffとrate追加
        df_month2['累計/前年差'] = df_month2['累計/受注/今期2'] - df_month2['累計/受注/前期2']
        df_month2['累計/前年比'] = df_month2['累計/受注/今期2'] / df_month2['累計/受注/前期2']

        #グラフを描くときの土台となるオブジェクト
        fig4 = go.Figure()
        #今期のグラフの追加
        for col in df_month2.columns[8:11]:
            fig4.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_month2[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_month2[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig4.update_layout(
            title='累計',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig4, use_container_width=True) 

        with st.expander('詳細', expanded=False):
            col_list = ['受注/今期', '受注/前期', '対前年差', '対前年比', '累計/受注/今期2', '累計/受注/前期2',\
                        '累計/前年差', '累計/前年比']
            df_temp = df_month2[col_list]
            st.table(df_temp)             

def senmon():
    month_list = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    columns_list = ['受注/今期', '受注/前期', '受注/前期年間', '対前年差', '対前年比']
    cust_list = [
        '（有）ケンポク家具', '株式会社丸ほん', 'ラボット・プランナー株式会社', '㈱家具のオツタカ'
    ]

    jnow_list = []
    jlast_list = []
    jlast_full_list = []
    sales_diff_list = []
    sales_rate_list = []

    for cust in cust_list:
        df_jnow2 = df_jnow[df_jnow['得意先名']==cust]
        df_jlast2 = df_jlast[df_jlast['得意先名']==cust]
        df_jlast_full2 = df_jlast_full[df_jlast_full['得意先名']==cust]
        target_num = 0
        for month in month_list:
            target = target_list[target_num]
            jnow = df_jnow2[df_jnow2['受注月'].isin([month])]['金額'].sum()
            jlast = df_jlast2[df_jlast2['受注月'].isin([month])]['金額'].sum()
            jlast_full = df_jlast_full2[df_jlast_full2['受注月'].isin([month])]['金額'].sum()
            
            sales_diff = jnow - jlast
            sales_rate = f'{jnow / jlast * 100: 0.1f} %'

            jnow_list.append('{:,}'.format(jnow))
            jlast_list.append('{:,}'.format(jlast))
            jlast_full_list.append('{:,}'.format(jlast_full))

            sales_diff_list.append('{:,}'.format(sales_diff))
            sales_rate_list.append(sales_rate)

            target_num += 1

        df_month = pd.DataFrame(list(zip(\
            jnow_list, jlast_list, jlast_full_list, sales_diff_list, sales_rate_list)), \
                columns=columns_list, index=month_list)
        
        jnow_list = []
        jlast_list = []
        jlast_full_list = []
        sales_diff_list = []
        sales_rate_list = []
    

        #*****受注ベース可視化
        df_month2 = df_month.copy()

        #グラフ用にint化
        df_month2['受注/今期2'] = df_month2['受注/今期'].apply(lambda x: x.replace(',', '')).astype('int')
        df_month2['受注/前期2'] = df_month2['受注/前期'].apply(lambda x: x.replace(',', '')).astype('int')
        df_month2['受注/前期年間2'] = df_month2['受注/前期年間'].apply(lambda x: x.replace(',', '')).astype('int')

        st.write(f'受注ベース/売上: {cust}')
        #グラフを描くときの土台となるオブジェクト
        fig3 = go.Figure()
        #今期のグラフの追加
        for col in df_month2.columns[5: 8]:
            fig3.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_month2[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_month2[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig3.update_layout(
            title='月別',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig3, use_container_width=True)

       

        #*****累計 受注ベース可視化
        #グラフ用にint化
        df_month2['累計/受注/今期2'] = df_month2['受注/今期2'].cumsum()
        df_month2['累計/受注/前期2'] = df_month2['受注/前期2'].cumsum()
        df_month2['累計/受注/前期年間2'] = df_month2['受注/前期年間2'].cumsum()

        #table用にdiffとrate追加
        df_month2['累計/前年差'] = df_month2['累計/受注/今期2'] - df_month2['累計/受注/前期2']
        df_month2['累計/前年比'] = df_month2['累計/受注/今期2'] / df_month2['累計/受注/前期2']

        #グラフを描くときの土台となるオブジェクト
        fig4 = go.Figure()
        #今期のグラフの追加
        for col in df_month2.columns[8:11]:
            fig4.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_month2[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_month2[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig4.update_layout(
            title='累計',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig4, use_container_width=True) 

        with st.expander('詳細', expanded=False):
            col_list = ['受注/今期', '受注/前期', '対前年差', '対前年比', '累計/受注/今期2', '累計/受注/前期2',\
                        '累計/前年差', '累計/前年比']
            df_temp = df_month2[col_list]
            st.table(df_temp) 

def senmon2():
    month_list = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    columns_list = ['受注/今期', '受注/前期', '受注/前期年間', '対前年差', '対前年比']
    cust_list = [
        '青森木工商事㈱', '㈱七尾家具百貨店', '株式会社　かさい',
        '有限会社　木乃や家具', '㈱日進', '（有）遠野家具センター',
        '(有)相馬屋家具店', '㈱家具の橋本', '㈱カルタ'
    ]

    jnow_list = []
    jlast_list = []
    jlast_full_list = []
    sales_diff_list = []
    sales_rate_list = []

    for cust in cust_list:
        df_jnow2 = df_jnow[df_jnow['得意先名']==cust]
        df_jlast2 = df_jlast[df_jlast['得意先名']==cust]
        df_jlast_full2 = df_jlast_full[df_jlast_full['得意先名']==cust]
        target_num = 0
        for month in month_list:
            target = target_list[target_num]
            jnow = df_jnow2[df_jnow2['受注月'].isin([month])]['金額'].sum()
            jlast = df_jlast2[df_jlast2['受注月'].isin([month])]['金額'].sum()
            jlast_full = df_jlast_full2[df_jlast_full2['受注月'].isin([month])]['金額'].sum()
            
            sales_diff = jnow - jlast
            sales_rate = f'{jnow / jlast * 100: 0.1f} %'

            jnow_list.append('{:,}'.format(jnow))
            jlast_list.append('{:,}'.format(jlast))
            jlast_full_list.append('{:,}'.format(jlast_full))

            sales_diff_list.append('{:,}'.format(sales_diff))
            sales_rate_list.append(sales_rate)

            target_num += 1

        df_month = pd.DataFrame(list(zip(\
            jnow_list, jlast_list, jlast_full_list, sales_diff_list, sales_rate_list)), \
                columns=columns_list, index=month_list)
        
        jnow_list = []
        jlast_list = []
        jlast_full_list = []
        sales_diff_list = []
        sales_rate_list = []
    

        #*****受注ベース可視化
        df_month2 = df_month.copy()

        #グラフ用にint化
        df_month2['受注/今期2'] = df_month2['受注/今期'].apply(lambda x: x.replace(',', '')).astype('int')
        df_month2['受注/前期2'] = df_month2['受注/前期'].apply(lambda x: x.replace(',', '')).astype('int')
        df_month2['受注/前期年間2'] = df_month2['受注/前期年間'].apply(lambda x: x.replace(',', '')).astype('int')

        st.write(f'受注ベース/売上: {cust}')
        #グラフを描くときの土台となるオブジェクト
        fig3 = go.Figure()
        #今期のグラフの追加
        for col in df_month2.columns[5: 8]:
            fig3.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_month2[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_month2[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig3.update_layout(
            title='月別',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig3, use_container_width=True)

       

        #*****累計 受注ベース可視化
        #グラフ用にint化
        df_month2['累計/受注/今期2'] = df_month2['受注/今期2'].cumsum()
        df_month2['累計/受注/前期2'] = df_month2['受注/前期2'].cumsum()
        df_month2['累計/受注/前期年間2'] = df_month2['受注/前期年間2'].cumsum()

        #table用にdiffとrate追加
        df_month2['累計/前年差'] = df_month2['累計/受注/今期2'] - df_month2['累計/受注/前期2']
        df_month2['累計/前年比'] = df_month2['累計/受注/今期2'] / df_month2['累計/受注/前期2']

        #グラフを描くときの土台となるオブジェクト
        fig4 = go.Figure()
        #今期のグラフの追加
        for col in df_month2.columns[8: 11]:
            fig4.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_month2[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_month2[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig4.update_layout(
            title='累計',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig4, use_container_width=True) 

        with st.expander('詳細', expanded=False):
            col_list = ['受注/今期', '受注/前期', '対前年差', '対前年比', '累計/受注/今期2', '累計/受注/前期2',\
                        '累計/前年差', '累計/前年比']
            df_temp = df_month2[col_list]
            st.table(df_temp) 

def target():
    month_list = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    columns_list = ['目標', '出荷/今期', '出荷/前期', '対目標差', '対目標比', '対前年差', '対前年比']
    cust_list = [
        'ラボット・プランナー株式会社', '㈱家具のオツタカ'
    ]
    labo_target_list = [
        600000, 900000, 1900000, 400000, 1000000, 800000,
        400000, 400000, 200000, 900000, 1100000, 400000
    ]
    otutaka_target_list = [
        400000, 400000, 400000, 400000, 400000, 400000,
        400000, 400000, 450000, 450000, 450000, 450000,
    ]
    target_all_list = [labo_target_list, otutaka_target_list]

    snow_list = []
    slast_list = []
    target_diff_list = []
    target_rate_list = []
    sales_diff_list = []
    sales_rate_list = []

    for (cust, target_list) in zip(cust_list, target_all_list):
        df_snow2 = df_snow[df_snow['得意先名']==cust]
        df_slast2 = df_slast[df_slast['得意先名']==cust]
        target_num = 0
        for month in month_list:
            target = target_list[target_num]
            snow = df_snow2[df_snow2['出荷月'].isin([month])]['金額'].sum()
            slast = df_slast2[df_slast2['出荷月'].isin([month])]['金額'].sum()

            target_diff = snow - target
            target_rate = f'{snow / target * 100: 0.1f} %'
            sales_diff = snow - slast
            sales_rate = f'{snow / slast * 100: 0.1f} %'

            snow_list.append('{:,}'.format(snow))
            slast_list.append('{:,}'.format(slast))
            target_diff_list.append(target_diff)
            target_rate_list.append(target_rate)
            sales_diff_list.append('{:,}'.format(sales_diff))
            sales_rate_list.append(sales_rate)

            target_num += 1

        df_month = pd.DataFrame(list(zip(\
            target_list, snow_list, slast_list, target_diff_list, target_rate_list, \
                sales_diff_list, sales_rate_list)), columns=columns_list, index=month_list)
        
        snow_list = []
        slast_list = []
        target_diff_list = []
        target_rate_list = []
        sales_diff_list = []
        sales_rate_list = []
    

        #*****受注ベース可視化
        df_month2 = df_month.copy()

        #グラフ用にint化
        df_month2['出荷/今期2'] = df_month2['出荷/今期'].apply(lambda x: int(x.replace(',', '')))
        df_month2['出荷/前期2'] = df_month2['出荷/前期'].apply(lambda x: int(x.replace(',', '')))

        with st.expander('詳細', expanded=False):
            df_temp = df_month2[columns_list]
            st.table(df_temp)

        st.write(f'受注ベース/売上: {cust}')
        #グラフを描くときの土台となるオブジェクト
        fig3 = go.Figure()
        #今期のグラフの追加
        for col in ['目標', '出荷/今期2', '出荷/前期2']:
            fig3.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_month2[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_month2[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig3.update_layout(
            title='月別',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig3, use_container_width=True)

        #*****累計 受注ベース可視化
        #グラフ用にint化
        df_month2['累計/目標'] = df_month2['目標'].cumsum()
        df_month2['累計/出荷/今期2'] = df_month2['出荷/今期2'].cumsum()
        df_month2['累計/出荷/前期2'] = df_month2['出荷/前期2'].cumsum()

        #table用にdiffとrate追加
        df_month2['累計/目標差'] = df_month2['累計/出荷/今期2'] - df_month2['累計/目標']
        df_month2['累計/目標比'] = df_month2['累計/出荷/今期2'] / df_month2['累計/目標']
        df_month2['累計/目標比'] = df_month2['累計/目標比'].apply(lambda x: '{:.2f}'.format(x))

        df_month2['累計/前年差'] = df_month2['累計/出荷/今期2'] - df_month2['累計/出荷/前期2']
        df_month2['累計/前年比'] = df_month2['累計/出荷/今期2'] / df_month2['累計/出荷/前期2']
        df_month2['累計/前年比'] = df_month2['累計/前年比'].apply(lambda x: '{:.2f}'.format(x))

        with st.expander('詳細', expanded=False):
            col_list = ['累計/目標', '累計/出荷/今期2', '累計/出荷/前期2', '累計/目標差', '累計/目標比', \
                        '累計/前年差', '累計/前年比']
            df_temp = df_month2[col_list]
            st.table(df_temp)

        #グラフを描くときの土台となるオブジェクト
        fig4 = go.Figure()
        #今期のグラフの追加
        for col in ['累計/目標', '累計/出荷/今期2', '累計/出荷/前期2']:
            fig4.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_month2[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_month2[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig4.update_layout(
            title='累計',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig4, use_container_width=True) 
       

def main():
    # アプリケーション名と対応する関数のマッピング
    apps = {
        '-': None,
        'TIF/星川★': tif,
        'TIF/その他': tif2,
        '専門店/星川★': senmon,
        '専門店/その他': senmon2,
        '目標比★': target

    }
    selected_app_name = st.sidebar.selectbox(label='分析項目の選択',
                                             options=list(apps.keys()))

    if selected_app_name == '-':
        st.info('サイドバーから分析項目を選択してください')
        st.stop()

    link = '[home](http://linkpagekh.s3-website-ap-northeast-1.amazonaws.com/)'
    st.sidebar.markdown(link, unsafe_allow_html=True)
    st.sidebar.caption('homeに戻る')    

    # 選択されたアプリケーションを処理する関数を呼び出す
    render_func = apps[selected_app_name]
    render_func()

if __name__ == '__main__':
    main()