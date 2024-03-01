import pandas as pd
import numpy as np
from pandas.core.frame import DataFrame
import streamlit as st
import plotly.figure_factory as ff
import plotly.graph_objects as go
import datetime
from datetime import timedelta

import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from func_collection import Graph

# pip install pandas numpy streamlit plotly google-oauth2-tool google-api-python-client openpyxl

st.set_page_config(page_title='販売員分析')
st.markdown('#### 販売員分析')

#小数点以下１ケタ
pd.options.display.float_format = '{:.1f}'.format

#current working dir
cwd = os.path.dirname(__file__)

fname_list = [
        '年齢別-担当者別分析TIF郡山', '年齢別-担当者別分析TIF港', '年齢別-担当者別分析TIF山形',
        '年齢別-担当者別分析TIF福島', '年齢別-担当者別分析ラボット', '年齢別-担当者別分析丸ほん'
                ]

#**********************gdriveからエクセルファイルのダウンロード・df化
def download_files():

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

# ファイルのダウンロード及び保存
download_files()

# データ加工をする関数
def make_data(df):
    min_date = df['受注日'].min()
    max_date = df['受注日'].max()

    st.sidebar.write(f'{min_date} - {max_date}')

    # 半年前を算出
    today = datetime.datetime.today()
    before180days = today - timedelta(days=180)

    start_date = st.sidebar.date_input(
        'データ開始日',
        before180days
        # datetime.datetime(2023, 1, 1)
    )
    end_date = st.sidebar.date_input(
        'データ終了日',
        today
    )

    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    df2 = df[(df['受注日'] >= start_date) & (df['受注日'] <= end_date)]

    # ***データ調整***
    df2['金額'] = df2['金額'].astype(int) #float →　int
    df2['受注月'] = pd.to_datetime(df2['受注日']).dt.strftime("%Y-%m")

    cates = []
    for cate in df2['商品分類名2']:
        if cate in ['クッション', 'リビングチェア', 'リビングテーブル']:
            cates.append('l')
        elif cate in ['ダイニングテーブル', 'ダイニングチェア', 'ベンチ']:
            cates.append('d')
        else:
            cates.append('e') 

    #LD列追加
    df2['cate'] = cates

    return df2

#インスタンス化
graph = Graph()


# #************************ファイルのdf化・加工
def shop():
    selected_cust = st.sidebar.selectbox(
        'ファイル名を選択',
        fname_list
        )
    # ***df化***
    path_cust = os.path.join(cwd, 'data', f'{selected_cust}.xlsx')
    df = pd.read_excel(
        path_cust, sheet_name='貼りつけ', usecols=[15, 16, 42, 43, 10, 50, 51]) #index　ナンバー不要　index_col=0


    df2 = make_data(df)

    st.markdown('##### 担当者別売上')
    s_rep = df2.groupby('取引先担当')['金額'].sum()
    s_rep.sort_values(ascending=False, inplace=True)

    graph.make_bar(s_rep, s_rep.index)

    #データの期間取得
    period = (df2['受注日'].max() - df2['受注日'].min()).days
    #１年に対する比率
    rate_temp = period / 365
    #換算するために必要な率
    year_rate = 1 / rate_temp
    st.caption(f'１年に対する比率: {rate_temp} 年換算するため必要な率: {year_rate}')

    df2['金額/年換算'] = df2['金額'] * year_rate 

    st.markdown('##### 担当者別売上/年換算')
    s_rep = df2.groupby('取引先担当')['金額/年換算'].sum()
    s_rep.sort_values(ascending=False, inplace=True)

    graph.make_bar(s_rep, s_rep.index)

    # 月別販売者数
    num_sales_dict = {}
    for month in df2['受注月']:
        temp_df = df2[df2['受注月']==month]
        num_sales = temp_df['取引先担当'].nunique()
        num_sales_dict[month] = num_sales
    
    df_num_sales = pd.DataFrame(num_sales_dict, index=['販売員数']).T

    st.markdown('##### 月別販売員数')
    graph.make_bar(df_num_sales['販売員数'], df_num_sales.index)

    st.markdown('##### 販売員別販売傾向')
    # 販売した商品列の新設
    df2['商品分類'] = df2['cate'] + '_' + df['シリーズ名']

    for rep in s_rep.index:
        _df = df2[df2['取引先担当']==rep]
        st.write(rep)
        graph.make_pie(_df['金額'], _df['商品分類'].unique())
    
    st.markdown('##### 月別販売員別売上')

    for rep in s_rep.index:
        _df = df2[df2['取引先担当']==rep]
        _s = _df.groupby('受注月')['金額'].sum()
        st.write(rep)
        graph.make_bar(_s, _s.index)

def ranking():
    # ***df化***
    df_all = pd.DataFrame()

    for fname in fname_list:
        path_cust = os.path.join(cwd, 'data', f'{fname}.xlsx')
        df = pd.read_excel(
            path_cust, sheet_name='貼りつけ', usecols=[15, 16, 42, 43, 10, 50, 51]) #index　ナンバー不要　index_col=0
        cust_name = fname.split('分析')[1]

        # 得意先名抽出
        df['得意先名'] = cust_name
        # 得意先名＋担当者名
        df['得意先名/担当者名'] = df['得意先名'] + '/' + df['取引先担当']
        # dfの連結
        df_all = pd.concat([df_all, df], axis=0)

    
    df_all2 = make_data(df_all)

    st.markdown('##### 担当者別売上')
    s_rep = df_all2.groupby('得意先名/担当者名')['金額'].sum()
    s_rep.sort_values(ascending=False, inplace=True)

    graph.make_bar(s_rep, s_rep.index)

    #データの期間取得
    period = (df_all2['受注日'].max() - df_all2['受注日'].min()).days
    #１年に対する比率
    rate_temp = period / 365
    #換算するために必要な率
    year_rate = 1 / rate_temp
    st.caption(f'１年に対する比率: {rate_temp} 年換算するため必要な率: {year_rate}')

    df_all2['金額/年換算'] = df_all2['金額'] * year_rate 

    st.markdown('##### 担当者別売上/年換算')
    s_rep = df_all2.groupby('得意先名/担当者名')['金額/年換算'].sum()
    s_rep.sort_values(ascending=False, inplace=True)

    graph.make_bar(s_rep, s_rep.index)

    # 星川抜き
    df_nonkh = df_all2[df_all2['取引先担当'] != '星川']
    
    for shop in df_nonkh['得意先名'].unique():
        # 得意先名毎に集計
        _df = df_nonkh[df_nonkh['得意先名']== shop]

        _s = _df.groupby('取引先担当')['金額/年換算'].sum()

        df_calc = pd.DataFrame(
            {
                '300万以上': 0,
                '250万以上': 0,
                '200万以上': 0,
                '150万以上': 0,
                '100万以上': 0,
                '50万以上': 0,
                '50万未満': 0,

            },
            index=['販売員数']
        )
        for sale in _s:
            if sale >= 3000000:
                df_calc['300万以上'] = df_calc['300万以上'] + 1
            elif sale >= 2500000:
                df_calc['250万以上'] = df_calc['250万以上'] + 1
            elif sale >= 2000000:
                df_calc['200万以上'] = df_calc['200万以上'] + 1
            elif sale >= 1500000:
                df_calc['150万以上'] = df_calc['150万以上'] + 1
            elif sale >= 1000000:
                df_calc['100万以上'] = df_calc['100万以上'] + 1
            elif sale >= 500000:
                df_calc['50万以上'] = df_calc['50万以上'] +1
            else:
                df_calc['50万未満'] = df_calc['50万未満'] +1

        df_calc = df_calc.T
        
        st.write(shop)
        graph.make_bar(df_calc['販売員数'], df_calc.index)
    
    
    # 販売員構成比/全店
    _s = df_nonkh.groupby('得意先名/担当者名')['金額/年換算'].sum()

    df_calc = pd.DataFrame(
        {
            '300万以上': 0,
            '250万以上': 0,
            '200万以上': 0,
            '150万以上': 0,
            '100万以上': 0,
            '50万以上': 0,
            '50万未満': 0,

        },
        index=['販売員数']
    )
    for sale in _s:
        if sale >= 3000000:
            df_calc['300万以上'] = df_calc['300万以上'] + 1
        elif sale >= 2500000:
            df_calc['250万以上'] = df_calc['250万以上'] + 1
        elif sale >= 2000000:
            df_calc['200万以上'] = df_calc['200万以上'] + 1
        elif sale >= 1500000:
            df_calc['150万以上'] = df_calc['150万以上'] + 1
        elif sale >= 1000000:
            df_calc['100万以上'] = df_calc['100万以上'] + 1
        elif sale >= 500000:
            df_calc['50万以上'] = df_calc['50万以上'] +1
        else:
            df_calc['50万未満'] = df_calc['50万未満'] +1

    df_calc = df_calc.T
    
    st.write('構成比/全店')
    graph.make_bar(df_calc['販売員数'], df_calc.index)

    graph.make_pie(df_calc['販売員数'], df_calc.index)
        







def main():
    # アプリケーション名と対応する関数のマッピング
    apps = {
        '-': None,
        '店舗別分析': shop,
        'ランキング': ranking


    }
    selected_app_name = st.sidebar.selectbox(label='分析項目の選択',
                                             options=list(apps.keys()))

    if selected_app_name == '-':
        st.info('サイドバーから分析項目を選択してください')
        st.stop()

    link = '[home](https://cocosan1-hidastreamlit4-linkpage-kh-sn2d6j.streamlit.app/)'
    st.sidebar.markdown(link, unsafe_allow_html=True)
    st.sidebar.caption('homeに戻る')    

    # 選択されたアプリケーションを処理する関数を呼び出す
    render_func = apps[selected_app_name]
    render_func()

if __name__ == '__main__':
    main()