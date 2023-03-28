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

st.set_page_config(page_title='売り上げ分析（エリア別）')
st.markdown('#### 売り上げ分析（エリア別)')

#current working dir
cwd = os.path.dirname(__file__)

#**********************gdriveからエクセルファイルのダウンロード・df化
fname_list = ['79j', '78j']
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
# ***今期受注***
path_now = os.path.join(cwd, 'data', '79j.xlsx')
df_now = pd.read_excel(
    path_now, sheet_name='受注委託移動在庫生産照会', usecols=[1, 3, 6, 8, 10, 14, 15, 16, 28, 31, 42, 50, 51, 52]) #index　ナンバー不要　index_col=0

# ***前期受注***
path_last = os.path.join(cwd, 'data', '78j.xlsx')
df_last = pd.read_excel(
    path_last, sheet_name='受注委託移動在庫生産照会', usecols=[1, 3, 6, 8, 10, 14, 15, 16, 28, 31, 42, 50, 51, 52]) #index　ナンバー不要　index_col=0

# *** 出荷月、受注月列の追加***
df_now['出荷月'] = df_now['出荷日'].dt.month
df_now['受注月'] = df_now['受注日'].dt.month
df_now['商品コード2'] = df_now['商　品　名'].map(lambda x: x.split()[0]) #品番
df_now['商品コード3'] = df_now['商　品　名'].map(lambda x: str(x)[0:2]) #頭品番
df_now['張地'] = df_now['商　品　名'].map(lambda x: x.split()[2] if len(x.split()) >= 4 else '')
df_last['出荷月'] = df_last['出荷日'].dt.month
df_last['受注月'] = df_last['受注日'].dt.month
df_last['商品コード2'] = df_last['商　品　名'].map(lambda x: x.split()[0])
df_last['商品コード3'] = df_last['商　品　名'].map(lambda x: str(x)[0:2]) #頭品番
df_last['張地'] = df_last['商　品　名'].map(lambda x: x.split()[2] if len(x.split()) >= 4 else '')

# ***INT型への変更***
df_now[['数量', '単価', '金額', '出荷倉庫', '原価金額', '出荷月', '受注月']] = df_now[['数量', '単価', '金額', '出荷倉庫', '原価金額', '出荷月', '受注月']].fillna(0).astype('int64')
#fillna　０で空欄を埋める
df_last[['数量', '単価', '金額', '出荷倉庫', '原価金額', '出荷月', '受注月']] = df_last[['数量', '単価', '金額', '出荷倉庫', '原価金額', '出荷月', '受注月']].fillna(0).astype('int64')
#fillna　０で空欄を埋める

fukushima_list = ['（有）ケンポク家具', '㈱東京ｲﾝﾃﾘｱ 福島店']
koriyama_list = ['ラボット・プランナー株式会社', '㈱東京ｲﾝﾃﾘｱ 郡山店']
iwaki_list = ['株式会社丸ほん', '㈱東京ｲﾝﾃﾘｱ いわき店', '㈱吉田家具店']
yamagata_list = ['㈱東京ｲﾝﾃﾘｱ 山形店', '㈱家具のオツタカ']
sendai_list = ['㈱家具の橋本', '(有)相馬屋家具店', '㈱東京ｲﾝﾃﾘｱ 仙台港本店', '㈱東京ｲﾝﾃﾘｱ 仙台泉店', \
               '㈱東京ｲﾝﾃﾘｱ 仙台南店']
kh_list = [fukushima_list, koriyama_list, iwaki_list, yamagata_list, sendai_list]
name_list = ['福島市', '郡山市', 'いわき市', '山形市', '仙台市']

def sales():
    for (area_list, name) in zip(kh_list, name_list): 
        sum_list = []   
        for cust in area_list:
            now_cust_sum = df_now[df_now['得意先名']==cust]['金額'].sum()
            last_cust_sum = df_last[df_last['得意先名']==cust]['金額'].sum()
            temp_list = [last_cust_sum, now_cust_sum]
            sum_list.append(temp_list)

        df_results = pd.DataFrame(sum_list, columns=['前期', '今期'], index=area_list)
        df_results.loc['合計'] = df_results.sum()
        df_results['対前年比'] = df_results['今期'] / df_results['前期']
        df_results['対前年差'] = df_results['今期'] - df_results['前期']
        df_results = df_results.T

        ratio = '{:.2f}'.format(df_results.loc['対前年比', '合計'])
        diff = '{:,}'.format(int(df_results.loc['対前年差', '合計']))
        st.markdown(f'##### {name}')
        st.metric(label='対前年比', value=ratio, delta=diff)

        #可視化
        #グラフを描くときの土台となるオブジェクト
        fig = go.Figure()
        #今期のグラフの追加
        for col in df_results.columns:
            fig.add_trace(
                go.Scatter(
                    x=df_results.index[:2],
                    y=df_results[col][:2],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_results[col][:2]/10000),
                    textposition="top center",
                    name=col)
            )

        #レイアウト設定     
        fig.update_layout(
            title=f'エリア別売上（累計）',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig, use_container_width=True) 

def sales_month():
    month_list = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]

    for (area_list, name) in zip(kh_list, name_list):  

        df_now_cust = df_now[df_now['得意先名'].isin(area_list)]
        df_last_cust = df_last[df_last['得意先名'].isin(area_list)]

        sum_list = []
        for month in month_list:
            df_now_month = df_now_cust[df_now_cust['受注月']==month]['金額'].sum()
            df_last_month = df_last_cust[df_last_cust['受注月']==month]['金額'].sum()
            temp_list = [df_now_month, df_last_month]
            sum_list.append(temp_list)

        df_results = pd.DataFrame(sum_list, index=month_list, columns=['今期', '前期']) 

        st.markdown(f'##### {name}')
        #可視化
        #グラフを描くときの土台となるオブジェクト
        fig = go.Figure()
        #今期のグラフの追加
        for col in df_results.columns:
            fig.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df_results[col],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_results[col]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig.update_layout(
            title='エリア別売上',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig, use_container_width=True) 

def ld_comp():

    for (area_list, name) in zip(kh_list, name_list):

        df_now_cust = df_now[df_now['得意先名'].isin(area_list)]
        df_last_cust = df_last[df_last['得意先名'].isin(area_list)]

        now_cust_sum_l = df_now_cust[df_now_cust['商品分類名2'].isin(\
            ['クッション', 'リビングチェア', 'リビングテーブル'])]['金額'].sum()

        now_cust_sum_d = df_now_cust[df_now_cust['商品分類名2'].isin(\
            ['ダイニングテーブル', 'ダイニングチェア', 'ベンチ'])]['金額'].sum() 

        last_cust_sum_l = df_last_cust[df_last_cust['商品分類名2'].isin(\
            ['クッション', 'リビングチェア', 'リビングテーブル'])]['金額'].sum()

        last_cust_sum_d = df_last_cust[df_last_cust['商品分類名2'].isin(\
            ['ダイニングテーブル', 'ダイニングチェア', 'ベンチ'])]['金額'].sum() 
        temp_list = [[last_cust_sum_l, last_cust_sum_d], [now_cust_sum_l, now_cust_sum_d]] 

        df_results = pd.DataFrame(temp_list, index=['前期', '今期'], columns=['Living', 'Dining'])
        df_results.loc['対前年比'] = df_results.loc['今期'] / df_results.loc['前期']
        df_results.loc['対前年差'] = df_results.loc['今期'] - df_results.loc['前期']

        st.markdown(f'##### {name}')

        col1, col2 = st.columns(2)
        with col1:
            st.write('Living')
            ratio = '{:.2f}'.format(df_results.loc['対前年比', 'Living'])
            diff = '{:,}'.format(int(df_results.loc['対前年差', 'Living']))
            st.metric(label='対前年比', value=ratio, delta=diff)

        with col2:
            st.write('Dining')
            ratio = '{:.2f}'.format(df_results.loc['対前年比', 'Dining'])
            diff = '{:,}'.format(int(df_results.loc['対前年差', 'Dining']))
            st.metric(label='対前年比', value=ratio, delta=diff)    


        #可視化
        #グラフを描くときの土台となるオブジェクト
        fig = go.Figure()
        #今期のグラフの追加
        for col in df_results.columns:
            fig.add_trace(
                go.Scatter(
                    x=df_results.index[:2], #対前年比,対前年差を拾わないように[:2]
                    y=df_results[col][:2],
                    mode = 'lines+markers+text', #値表示
                    text=round(df_results[col][:2]/10000),
                    textposition="top center", 
                    name=col)
            )

        #レイアウト設定     
        fig.update_layout(
            title='LD別売上',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig, use_container_width=True)         



def main():
    # アプリケーション名と対応する関数のマッピング
    apps = {
        '-': None,
        '売上: 累計': sales,
        '売上: 月別': sales_month,
        'LD別売上: 累計': ld_comp

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