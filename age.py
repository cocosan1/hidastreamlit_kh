import pandas as pd
import numpy as np
from pandas.core.frame import DataFrame
import streamlit as st
import plotly.figure_factory as ff
import plotly.graph_objects as go
import datetime

import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from func_collection import Graph

st.set_page_config(page_title='売り上げ分析（年齢層）')
st.markdown('#### 売り上げ分析（年齢層）')

#小数点以下１ケタ
pd.options.display.float_format = '{:.1f}'.format

#current working dir
cwd = os.path.dirname(__file__)

#**********************gdriveからエクセルファイルのダウンロード・df化
fname_list = [
    '年齢別-担当者別分析TIF郡山79', '年齢別-担当者別分析TIF港79', '年齢別-担当者別分析TIF山形79',
    '年齢別-担当者別分析TIF福島79', '年齢別-担当者別分析オツタカ', '年齢別-担当者別分析ケンポク',
    '年齢別-担当者別分析ラボット79', '年齢別-担当者別分析丸ほん79'
              ]

# *** selectbox 得意先名***
selected_cust = st.sidebar.selectbox(
    '得意先名:',
    fname_list,   
) 


# Google Drive APIを使用するための認証情報を取得する
creds_dict = st.secrets["gcp_service_account"]
creds = service_account.Credentials.from_service_account_info(creds_dict)

# Drive APIのクライアントを作成する
#API名（ここでは"drive"）、APIのバージョン（ここでは"v3"）、および認証情報を指定
service = build("drive", "v3", credentials=creds)

# 指定したファイル名を持つファイルのIDを取得する
#Google Drive上のファイルを検索するためのクエリを指定して、ファイルの検索を実行します。
# この場合、ファイル名とMIMEタイプを指定しています。
file_name = f"{selected_cust}.xlsx"
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


# #************************ファイルのdf化・加工
# ***df化***
path_cust = os.path.join(cwd, 'data', f'{selected_cust}.xlsx')
df = pd.read_excel(
    path_cust, sheet_name='貼りつけ', usecols=[15, 16, 42, 43, 10, 50, 51]) #index　ナンバー不要　index_col=0

min_date = df['受注日'].min()
max_date = df['受注日'].max()

st.sidebar.write(f'{min_date} - {max_date}')

start_date = st.sidebar.date_input(
    'データ開始日',
    datetime.datetime(2023, 1, 1)
)
end_date = st.sidebar.date_input(
    'データ終了日',
    datetime.datetime(2023, 6, 30)
)

start_date = pd.to_datetime(start_date)
end_date = pd.to_datetime(end_date)

df2 = df[(df['受注日'] >= start_date) & (df['受注日'] <= end_date)]

# ***データ調整***
df2 = df2.dropna(how='any') #一つでも欠損値のある行を削除　all　全て欠損値の行を削除
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

#インスタンス化
graph = Graph()

#年代別df
df_30 = df2[df2['年代']== 30]
df_40 = df2[df2['年代']== 40]
df_50 = df2[df2['年代']== 50]

def age_whole():
    
    #全体
    st.markdown('##### ■ 全体')
    col1, col2 = st.columns(2)
    s_age = df2.groupby('年代')['金額'].sum()

    with col1:
        st.write('売上')
        graph.make_bar(s_age, s_age.index)
    with col2:
        st.write('構成比')
        graph.make_pie(s_age, s_age.index)
    
    #living
    st.markdown('##### ■ living')
    col1, col2 = st.columns(2)
    df_l = df2[df2['cate']=='l']

    s_agel = df_l.groupby('年代')['金額'].sum()
    
    with col1:
        st.write('売上')
        graph.make_bar(s_agel, s_agel.index)
    with col2:
        st.write('構成比')
        graph.make_pie(s_agel, s_agel.index)
    
    #dining
    st.markdown('##### ■ dining')
    col1, col2 = st.columns(2)
    df_d = df2[df2['cate']=='d']

    s_aged = df_d.groupby('年代')['金額'].sum()
    
    with col1:
        st.write('売上')
        graph.make_bar(s_aged, s_aged.index)
    with col2:
        st.write('構成比')
        graph.make_pie(s_aged, s_aged.index)

def age_series():
    st.markdown('##### 年齢層別シリーズ別')

    #30代
    st.markdown('##### ■ 30代')
    #living
    st.write('living')
    col1, col2 = st.columns(2)
    
    df_30l = df_30[df_30['cate']=='l']

    s_30l = df_30l.groupby('シリーズ名')['金額'].sum()

    with col1:
        st.write('売上')
        graph.make_bar(s_30l, s_30l.index)
    with col2:
        st.write('構成比')
        graph.make_pie(s_30l, s_30l.index)
    
    #dining
    st.write('dining')
    col1, col2 = st.columns(2)
    
    df_30d = df_30[df_30['cate']=='d']

    s_30d = df_30d.groupby('シリーズ名')['金額'].sum()

    with col1:
        st.write('売上')
        graph.make_bar(s_30d, s_30d.index)
    with col2:
        st.write('構成比')
        graph.make_pie(s_30d, s_30d.index)
    
    st.divider()
    #40代
    st.markdown('##### ■ 40代')
    #living
    st.write('living')
    col1, col2 = st.columns(2)
    
    df_40l = df_40[df_40['cate']=='l']

    s_40l = df_40l.groupby('シリーズ名')['金額'].sum()

    with col1:
        st.write('売上')
        graph.make_bar(s_40l, s_40l.index)
    with col2:
        st.write('構成比')
        graph.make_pie(s_40l, s_40l.index)
    
    #dining
    st.write('dining')
    col1, col2 = st.columns(2)
    
    df_40d = df_40[df_40['cate']=='d']

    s_40d = df_40d.groupby('シリーズ名')['金額'].sum()

    with col1:
        st.write('売上')
        graph.make_bar(s_40d, s_40d.index)
    with col2:
        st.write('構成比')
        graph.make_pie(s_40d, s_40d.index)
    
    st.divider()
    #50代
    st.markdown('##### ■ 50代')
    #living
    st.write('living')
    col1, col2 = st.columns(2)
    
    df_50l = df_50[df_50['cate']=='l']

    s_50l = df_50l.groupby('シリーズ名')['金額'].sum()

    with col1:
        st.write('売上')
        graph.make_bar(s_50l, s_50l.index)
    with col2:
        st.write('構成比')
        graph.make_pie(s_50l, s_50l.index)
    
    #dining
    st.write('dining')
    col1, col2 = st.columns(2)
    
    df_50d = df_50[df_50['cate']=='d']

    s_50d = df_50d.groupby('シリーズ名')['金額'].sum()

    with col1:
        st.write('売上')
        graph.make_bar(s_50d, s_50d.index)
    with col2:
        st.write('構成比')
        graph.make_pie(s_50d, s_50d.index)

def suii_month():
    st.markdown('##### 月別売上推移/年齢層')
    
    s_30 = df_30.groupby('受注月')['金額'].sum()
    s_40 = df_40.groupby('受注月')['金額'].sum()
    s_50 = df_50.groupby('受注月')['金額'].sum()

    ages = [s_30, s_40, s_50]
    graph.make_line(ages, ['30', '40', '50'], s_50.index)


def rep():
    st.markdown('##### 担当者別売上')
    s_rep = df2.groupby('取引先担当')['金額'].sum()
    s_rep.sort_values(ascending=False, inplace=True)

    graph.make_bar(s_rep, s_rep.index)

    

 



def main():
    # アプリケーション名と対応する関数のマッピング
    apps = {
        '-': None,
        '年齢構成比 全体': age_whole,
        '年齢ベース/シリーズ別構成比': age_series,
        '月別売上推移/年齢層': suii_month,
        '担当者別売上': rep,


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