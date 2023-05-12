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

st.set_page_config(page_title='売り上げ分析（得意先別）')
st.markdown('#### 売り上げ分析（得意先別)')

#小数点以下１ケタ
pd.options.display.float_format = '{:.2f}'.format

#current working dir
cwd = os.path.dirname(__file__)

#**********************gdriveからエクセルファイルのダウンロード・df化
fname_list = ['kita79j', 'kita78j']
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
path_now = os.path.join(cwd, 'data', 'kita79j.xlsx')
df_now = pd.read_excel(
    path_now, sheet_name='受注委託移動在庫生産照会', usecols=[1, 3, 6, 8, 10, 14, 15, 16, 28, 31, 42, 50, 51, 52]) #index　ナンバー不要　index_col=0

# ***前期受注***
path_last = os.path.join(cwd, 'data', 'kita78j.xlsx')
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
 
df_now_total = df_now['金額'].sum()
df_last_total = df_last['金額'].sum()

# *** selectbox 得意先名***
customer = df_now['得意先名'].unique()
option_customer = st.selectbox(
    '得意先名:',
    customer,   
) 

df_now_cust =df_now[df_now['得意先名']==option_customer]
df_last_cust =df_last[df_last['得意先名']==option_customer]
df_now_cust_total = df_now_cust['金額'].sum()
df_last_cust_total = df_last_cust['金額'].sum()

def earnings_comparison_year():
    total_cust_now = df_now[df_now['得意先名']==option_customer]['金額'].sum()
    total_cust_last = df_last[df_last['得意先名']==option_customer]['金額'].sum()
    total_comparison = f'{total_cust_now / total_cust_last * 100: 0.1f} %'
    diff = '{:,}'.format(total_cust_now - total_cust_last)
    
    with st.expander('詳細', expanded=False):
        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric('今期売上', value= '{:,}'.format(total_cust_now), delta=diff)
        with col2:
            st.metric('前期売上', value= '{:,}'.format(total_cust_last))
        with col3:
            st.metric('対前年比', value= total_comparison)

    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig = go.Figure()
    #今期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=['今期'],
            y=[total_cust_now],
            text=round(total_cust_now/10000),
            textposition="outside", 
            name='今期')
    )
    #前期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=['前期'],
            y=[total_cust_last],
            text=round(total_cust_last/10000),
            textposition="outside", 
            name='前期'
            )
    )
    #レイアウト設定     
    fig.update_layout(
        title='別売上（累計）',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig, use_container_width=True)        

def earnings_comparison_month():
    month_list = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    columns_list = ['今期', '前期', '対前年差', '対前年比']
    df_now_cust = df_now[df_now['得意先名']==option_customer]
    df_last_cust = df_last[df_last['得意先名']==option_customer]

    earnings_now = []
    earnings_last = []
    earnings_diff = []
    earnings_rate = []

    for month in month_list:
        earnings_month_now = df_now_cust[df_now_cust['受注月'].isin([month])]['金額'].sum()
        earnings_month_last = df_last_cust[df_last_cust['受注月'].isin([month])]['金額'].sum()
        earnings_diff_culc = earnings_month_now - earnings_month_last
        earnings_rate_culc = f'{earnings_month_now / earnings_month_last * 100: 0.1f} %'

        earnings_now.append('{:,}'.format(earnings_month_now))
        earnings_last.append('{:,}'.format(earnings_month_last))
        earnings_diff.append('{:,}'.format(earnings_diff_culc))
        earnings_rate.append(earnings_rate_culc)

    df_earnings_month = pd.DataFrame(list(zip(earnings_now, earnings_last, earnings_diff, earnings_rate)), columns=columns_list, index=month_list)
    
    with st.expander('詳細', expanded=False):
        st.caption('受注月ベース')
        st.table(df_earnings_month)

    #グラフ用にintのデータを用意
    df_earnings_month2 = df_earnings_month.copy()
    df_earnings_month2['今期'] = df_earnings_month2['今期'].apply(lambda x: int(x.replace(',', '')))
    df_earnings_month2['前期'] = df_earnings_month2['前期'].apply(lambda x: int(x.replace(',', '')))


    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig = go.Figure()
    #今期のグラフの追加
    for col in df_earnings_month2.columns[:2]:
        fig.add_trace(
            go.Scatter(
                x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                y=df_earnings_month2[col],
                mode = 'lines+markers+text', #値表示
                text=round(df_earnings_month2[col]/10000),
                textposition="top center", 
                name=col)
        )

    #レイアウト設定     
    fig.update_layout(
        title='月別売上',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig, use_container_width=True)

def mean_earning_month():
    st.write('#### 平均成約単価')
    month_list = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    columns_list = ['今期', '前期', '対前年差', '対前年比']
    df_now_cust = df_now[df_now['得意先名']==option_customer]
    df_last_cust = df_last[df_last['得意先名']==option_customer]

    order_num_now = []
    for num in df_now_cust['伝票番号']:
        num2 = num.split('-')[0]
        order_num_now.append(num2)
    df_now_cust['order_num'] = order_num_now

    order_num_last = []
    for num in df_last_cust['伝票番号']:
        num2 = num.split('-')[0]
        order_num_last.append(num2)
    df_last_cust['order_num'] = order_num_last


    earnings_now = []
    earnings_last = []
    earnings_diff = []
    earnings_rate = []

    for month in month_list:
        earnings_month_now = df_now_cust[df_now_cust['受注月'].isin([month])]
        order_sum_now = earnings_month_now.groupby('order_num')['金額'].sum()
        order_mean_now = order_sum_now.mean()

        earnings_month_last = df_last_cust[df_last_cust['受注月'].isin([month])]
        order_sum_last = earnings_month_last.groupby('order_num')['金額'].sum()
        order_mean_last = order_sum_last.mean()

        order_mean_diff = order_mean_now - order_mean_last
        if order_mean_last == 0:
            order_mean_rate = '0%'
        else:
            order_mean_rate = f'{(order_mean_now / order_mean_last)*100: 0.1f} %'
        
        earnings_now.append(order_mean_now)
        earnings_last.append(order_mean_last)
        earnings_diff.append(order_mean_diff)
        earnings_rate.append(order_mean_rate)

    df_mean_earninngs_month = pd.DataFrame(list(zip(earnings_now, earnings_last, earnings_diff, earnings_rate)), columns=columns_list, index=month_list)
    st.caption('受注月ベース')

    col1, col2 = st.columns(2)

    with col1:
        diff = int(df_mean_earninngs_month['今期'].mean()) - int(df_mean_earninngs_month['前期'].mean())
        st.metric('今期平均', value='{:,}'.format(int(df_mean_earninngs_month['今期'].mean())), \
            delta='{:,}'.format(diff))

    with col2:
        st.metric('前期平均', value='{:,}'.format(int(df_mean_earninngs_month['前期'].mean()))) 

    df_mean_earninngs_month.fillna(0, inplace=True)
    df_mean_earninngs_month['今期'] = \
        df_mean_earninngs_month['今期'].map(lambda x: '{:,}'.format(int(x))) 
    df_mean_earninngs_month['前期'] = \
        df_mean_earninngs_month['前期'].map(lambda x: '{:,}'.format(int(x))) 
    df_mean_earninngs_month['対前年差'] = \
        df_mean_earninngs_month['対前年差'].map(lambda x: '{:,}'.format(int(x)))   
    
    with st.expander('詳細', expanded=False):
        st.table(df_mean_earninngs_month) 

    #グラフ用にintのデータを用意
    df_mean_earninngs_month2 = df_mean_earninngs_month.copy()
    df_mean_earninngs_month2['今期'] = df_mean_earninngs_month2['今期'].apply(lambda x: int(x.replace(',', '')))
    df_mean_earninngs_month2['前期'] = df_mean_earninngs_month2['前期'].apply(lambda x: int(x.replace(',', '')))


    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig = go.Figure()
    #今期のグラフの追加
    for col in df_mean_earninngs_month2.columns[:2]:
        fig.add_trace(
            go.Scatter(
                x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                y=df_mean_earninngs_month2[col],
                mode = 'lines+markers+text', #値表示
                text=df_mean_earninngs_month2[col],
                textposition="top center", 
                name=col)
        )

    #レイアウト設定     
    fig.update_layout(
        title='平均成約単価',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig, use_container_width=True)         
    
def living_dining_comparison():
    st.markdown('##### LD 前年比/構成比')

    col1, col2 = st.columns(2)
    with col1:
        living_now = df_now_cust[df_now_cust['商品分類名2'].isin(['クッション', 'リビングチェア', 'リビングテーブル'])]['金額'].sum()
        living_last = df_last_cust[df_last_cust['商品分類名2'].isin(['クッション', 'リビングチェア', 'リビングテーブル'])]['金額'].sum()

        l_diff = living_now-living_last
        l_ratio = f'{living_now/living_last*100:0.1f} %'

        #可視化
        #グラフを描くときの土台となるオブジェクト
        fig = go.Figure()
        #今期のグラフの追加
        fig.add_trace(
            go.Bar(
                x=['今期'],
                y=[living_now],
                text=round(living_now/10000),
                textposition="outside", 
                name='今期')
        )
        #前期のグラフの追加
        fig.add_trace(
            go.Bar(
                x=['前期'],
                y=[living_last],
                text=round(living_last/10000),
                textposition="outside", 
                name='前期'
                )
        )
        #レイアウト設定     
        fig.update_layout(
            title='リビング',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig, use_container_width=True)  

            
            
        st.caption(f'対前年差 {l_diff}')
        st.caption(f'対前年比 {l_ratio}')
        st.caption('クッション/リビングチェア/リビングテーブル')

    with col2:
        dining_now = df_now_cust[df_now_cust['商品分類名2'].isin(['ダイニングテーブル', 'ダイニングチェア', 'ベンチ'])]['金額'].sum()
        dining_last = df_last_cust[df_last_cust['商品分類名2'].isin(['ダイニングテーブル', 'ダイニングチェア', 'ベンチ'])]['金額'].sum()

        d_diff = dining_now-dining_last
        d_ratio = f'{dining_now/dining_last*100:0.1f} %'

        #可視化
        #グラフを描くときの土台となるオブジェクト
        fig2 = go.Figure()
        #今期のグラフの追加
        fig2.add_trace(
            go.Bar(
                x=['今期'],
                y=[dining_now],
                text=round(dining_now/10000),
                textposition="outside", 
                name='今期')
        )
        #前期のグラフの追加
        fig2.add_trace(
            go.Bar(
                x=['前期'],
                y=[dining_last],
                text=round(dining_last/10000),
                textposition="outside", 
                name='前期'
                )
        )
        #レイアウト設定     
        fig2.update_layout(
            title='ダイニング',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig2, use_container_width=True) 
        
        st.caption(f'対前年差 {d_diff}')
        st.caption(f'対前年比 {d_ratio}')
        st.caption('ダイニングテーブル/ダイニングチェア/ベンチ')
    

    else_now = df_now_cust[df_now_cust['商品分類名2'].isin(['キャビネット類', 'その他テーブル', '雑品・特注品', 'その他椅子','デスク', '小物・その他'])]['金額'].sum()
    else_last = df_last_cust[df_last_cust['商品分類名2'].isin(['キャビネット類', 'その他テーブル', '雑品・特注品', 'その他椅子','デスク', '小物・その他'])]['金額'].sum()

    with col1:
        comp_now_list = [living_now, dining_now, else_now]
        comp_now_index = ['リビング', 'ダイニング', 'その他']
        comp_now_columns = ['分類']
        df_comp_now = pd.DataFrame(comp_now_index, columns=comp_now_columns)
        df_comp_now['金額'] = comp_now_list

        # グラフ
        st.markdown('###### LD比率(今期)')
        fig_ld_ratio_now = go.Figure(
            data=[
                go.Pie(
                    labels=df_comp_now['分類'],
                    values=df_comp_now['金額']
                    )])
        fig_ld_ratio_now.update_layout(
            showlegend=True, #凡例表示
            height=200,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_ld_ratio_now.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_ld_ratio_now, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    with col2:
        comp_last_list = [living_last, dining_last, else_last]
        comp_last_index = ['リビング', 'ダイニング', 'その他']
        comp_last_columns = ['分類']
        df_comp_last = pd.DataFrame(comp_last_index, columns=comp_last_columns)
        df_comp_last['金額'] = comp_last_list

        # グラフ
        st.markdown('###### LD比率(前期)')
        fig_ld_ratio_last = go.Figure(
            data=[
                go.Pie(
                    labels=df_comp_last['分類'],
                    values=df_comp_last['金額']
                    )])
        fig_ld_ratio_last.update_layout(
            showlegend=True, #凡例表示
            height=200,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_ld_ratio_last.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_ld_ratio_last, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

def living_dining_comparison_ld():

    # *** selectbox LD***
    category = ['リビング', 'ダイニング']
    option_category = st.selectbox(
        'category:',
        category,   
    ) 
    if option_category == 'リビング':
        df_now_cust_cate = df_now_cust[df_now_cust['商品分類名2'].isin(['クッション', 'リビングチェア', 'リビングテーブル'])]
        df_last_cust_cate = df_last_cust[df_last_cust['商品分類名2'].isin(['クッション', 'リビングチェア', 'リビングテーブル'])]
    elif option_category == 'ダイニング':
        df_now_cust_cate = df_now_cust[df_now_cust['商品分類名2'].isin(['ダイニングテーブル', 'ダイニングチェア', 'ベンチ'])]
        df_last_cust_cate = df_last_cust[df_last_cust['商品分類名2'].isin(['ダイニングテーブル', 'ダイニングチェア', 'ベンチ'])]

    index = []
    now_result = []
    last_result = []
    diff = []
    ratio = []
    series_list = df_now_cust_cate['シリーズ名'].unique()
    
    for series in series_list:
        index.append(series)
        now_culc = df_now_cust_cate[df_now_cust_cate['シリーズ名']==series]['金額'].sum()
        last_culc = df_last_cust_cate[df_last_cust_cate['シリーズ名']==series]['金額'].sum()
        now_result.append(now_culc)
        last_result.append(last_culc)
        diff_culc = '{:,}'.format(now_culc - last_culc)
        diff.append(diff_culc)
        ratio_culc = f'{now_culc/last_culc*100:0.1f} %'
        ratio.append(ratio_culc)
    df_result = pd.DataFrame(list(zip(now_result, last_result, ratio, diff)), index=index, columns=['今期', '前期', '対前年比', '対前年差'])
    
    with st.expander('一覧', expanded=False):
     st.dataframe(df_result)
     st.caption('列名クリックでソート')
    
    #**********構成比円グラフ***************
    col1, col2 = st.columns(2)

    with col1:
        # グラフ
        st.write(f'{option_category} 構成比(今期)')
        fig_ld_ratio_now = go.Figure(
            data=[
                go.Pie(
                    labels=df_result.index,
                    values=df_result['今期']
                    )])
        fig_ld_ratio_now.update_layout(
            showlegend=True, #凡例表示
            height=200,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_ld_ratio_now.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_ld_ratio_now, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    with col2:
        # グラフ
        st.write(f'{option_category} 構成比(前期)')
        fig_ld_ratio_last = go.Figure(
            data=[
                go.Pie(
                    labels=df_result.index,
                    values=df_result['前期']
                    )])
        fig_ld_ratio_last.update_layout(
            showlegend=True, #凡例表示
            height=200,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_ld_ratio_last.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_ld_ratio_last, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    #*********前年比棒グラフ**************
    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig2 = go.Figure()
    #今期のグラフの追加
    fig2.add_trace(
        go.Bar(
            x=df_result.index,
            y=df_result['今期'],
            text=round(df_result['今期']/10000),
            textposition="outside", 
            name='今期')
    )
    #前期のグラフの追加
    fig2.add_trace(
        go.Bar(
            x=df_result.index,
            y=df_result['前期'],
            text=round(df_result['前期']/10000),
            textposition="outside", 
            name='前期'
            )
    )
    #レイアウト設定     
    fig2.update_layout(
        title='シリーズ別売上',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig2, use_container_width=True)     

def series():
    # *** selectbox 商品分類2***
    category = df_now['商品分類名2'].unique()
    option_category = st.selectbox(
        'category:',
        category,   
    ) 
    st.caption('構成比は下段')
    categorybase_now = df_now[df_now['商品分類名2']==option_category]
    categorybase_last = df_last[df_last['商品分類名2']==option_category]
    categorybase_cust_now = categorybase_now[categorybase_now['得意先名']== option_customer]
    categorybase_cust_last = categorybase_last[categorybase_last['得意先名']== option_customer]

    # ***シリーズ別売り上げ ***
    series_now = categorybase_cust_now.groupby('シリーズ名')['金額'].sum().sort_values(ascending=False).head(12) #降順
    series_now2 = series_now.apply('{:,}'.format) #数値カンマ区切り　注意strになる　グラフ作れなくなる

    series_last = categorybase_cust_last.groupby('シリーズ名')['金額'].sum().sort_values(ascending=False).head(12) #降順
    series_last2 = series_last.apply('{:,}'.format) #数値カンマ区切り　注意strになる　グラフ作れなくなる

    col1, col2 = st.columns(2)

    with col1:
        # グラフ　シリーズ別売り上げ構成比
        st.markdown('###### シリーズ別売り上げ構成比(今期)')
        fig_series_ratio_now = go.Figure(
            data=[
                go.Pie(
                    labels=series_now.index,
                    values=series_now
                    )])
        fig_series_ratio_now.update_layout(
            showlegend=True, #凡例表示
            height=200,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_series_ratio_now.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_series_ratio_now, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    with col2:
        # グラフ　シリーズ別売り上げ構成比
        st.markdown('###### シリーズ別売り上げ構成比(前期)')
        fig_series_ratio_last = go.Figure(
            data=[
                go.Pie(
                    labels=series_last.index,
                    values=series_last
                    )])
        fig_series_ratio_last.update_layout(
            showlegend=True, #凡例表示
            height=200,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_series_ratio_last.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_series_ratio_last, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

     #*********前年比棒グラフ**************
    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig2 = go.Figure()
    #今期のグラフの追加
    fig2.add_trace(
        go.Bar(
            x=series_now.index,
            y=series_now,
            text=round(series_now/10000),
            textposition="outside", 
            name='今期')
    )
    #前期のグラフの追加
    fig2.add_trace(
        go.Bar(
            x=series_last.index,
            y=series_last,
            text=round(series_last/10000),
            textposition="outside", 
            name='前期'
            )
    )
    #レイアウト設定     
    fig2.update_layout(
        title='シリーズ別売上',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig2, use_container_width=True)      

def item_count_category():
    # *** selectbox 得意先名***
    categories = df_now_cust['商品分類名2'].unique()
    option_categories = st.selectbox(
    '商品分類名2:',
    categories,   
    )    

    index = []
    count_now = []
    count_last = []
    diff = []

    df_now_cust_categories = df_now_cust[df_now_cust['商品分類名2']==option_categories]
    df_last_cust_categories = df_last_cust[df_last_cust['商品分類名2']==option_categories]
    series_list = df_now_cust[df_now_cust['商品分類名2']==option_categories]['シリーズ名'].unique()
    for series in series_list:
        index.append(series)
        month_len = len(df_now['受注月'].unique())
        df_now_cust_categories_count_culc = df_now_cust_categories[df_now_cust_categories['シリーズ名']==series]['数量'].sum()
        df_last_cust_categories_count_culc = df_last_cust_categories[df_last_cust_categories['シリーズ名']==series]['数量'].sum()
        count_now.append(f'{df_now_cust_categories_count_culc/month_len: 0.1f}')
        count_last.append(f'{df_last_cust_categories_count_culc/month_len: 0.1f}')
        diff.append(f'{(df_now_cust_categories_count_culc/month_len) - (df_last_cust_categories_count_culc/month_len):0.1f}')

    with st.expander('一覧', expanded=False):
        st.write('回転数/月平均')
        df_item_count = pd.DataFrame(list(zip(count_now, count_last, diff)), index=index, columns=['今期', '前期', '対前年差'])
        st.table(df_item_count) #列幅問題未解決 

    #*********前年比棒グラフ**************
    df_item_count2 = df_item_count.copy()
    df_item_count2['今期'] = df_item_count2['今期'].apply(lambda x: float(x))
    df_item_count2['前期'] = df_item_count2['前期'].apply(lambda x: float(x))
    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig2 = go.Figure()
    #今期のグラフの追加
    fig2.add_trace(
        go.Bar(
            x=df_item_count2.index,
            y=df_item_count2['今期'],
            text=df_item_count2['今期'],
            textposition="outside", 
            name='今期')
    )
    #前期のグラフの追加
    fig2.add_trace(
        go.Bar(
            x=df_item_count2.index,
            y=df_item_count2['前期'],
            text=df_item_count2['前期'],
            textposition="outside", 
            name='前期'
            )
    )
    #レイアウト設定     
    fig2.update_layout(
        title='シリーズ別回転数/月平均',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig2, use_container_width=True)   

def category_count_month():
    #　回転数 商品分類別 月毎
    # *** selectbox シリーズ名***
    category_list = df_now_cust['商品分類名2'].unique()
    option_category = st.selectbox(
    '商品分類名:',
    category_list,   
    ) 
    df_now_cust_category = df_now_cust[df_now_cust['商品分類名2']==option_category]
    
    months = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    count_now = []
    series_list = df_now_cust_category['シリーズ名'].unique()
    df_count = pd.DataFrame(index=series_list)
    for month in months:
        for series in series_list:
            df_now_cust_category_ser = df_now_cust_category[df_now_cust_category['シリーズ名']==series]
            count = df_now_cust_category_ser[df_now_cust_category_ser['受注月']==month]['数量'].sum()
            count_now.append(count)
        df_count[month] = count_now
        count_now = []

    with st.expander('一覧', expanded=False): 
        st.caption('今期')
        st.write(df_count)
    

    #可視化
    df_count2 = df_count.T #index月　col　seirs名に転置
    #グラフを描くときの土台となるオブジェクト
    fig = go.Figure()
    #今期のグラフの追加
    for col in df_count2.columns:
        fig.add_trace(
            go.Scatter(
                x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                y=df_count2[col],
                mode = 'lines+markers+text', #値表示
                text=df_count2[col],
                textposition="top center", 
                name=col)
        )

    #レイアウト設定     
    fig.update_layout(
        title='月別回転数',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig, use_container_width=True)       



def hokkaido_fushi_kokusanzai():
    # *** 北海道比率　節材比率　国産材比率 ***
    col1, col2, col3 = st.columns(3)
    cust_now = df_now[df_now['得意先名']== option_customer]
    cust_last = df_last[df_last['得意先名']== option_customer]
    total_now = cust_now['金額'].sum()
    total_last = cust_last['金額'].sum()

    #分類の詳細
    with st.expander('分類の詳細'):
        st.write('【節あり】森のことば/LEVITA (ﾚｳﾞｨﾀ)/森の記憶/とき葉/森のことばIBUKI/森のことば ウォルナット')
        st.write('【国産材1】北海道民芸家具/HIDA/Northern Forest/北海道HMその他/杉座/ｿﾌｨｵ SUGI/風のうた\
            Kinoe/SUWARI/KURINOKI')
        st.write('【国産材2】SG261M/SG261K/SG261C/SG261AM/SG261AK/SG261AC/KD201M/KD201K/KD201C\
                 KD201AM/KD201AK/KD201AC')
    with col1:
        hokkaido_now = cust_now[cust_now['出荷倉庫']==510]['金額'].sum()
        hokkaido_last = cust_last[cust_last['出荷倉庫']==510]['金額'].sum()
        hokkaido_diff = f'{(hokkaido_now/total_now*100) - (hokkaido_last/total_last*100):0.1f} %'

        st.metric('北海道工場比率', value=f'{hokkaido_now/total_now*100: 0.1f} %', delta=hokkaido_diff) #小数点以下1ケタ
        st.caption(f'前年 {hokkaido_last/total_last*100: 0.1f} %')

    with col2:
        fushi_now = cust_now[cust_now['シリーズ名'].isin(['森のことば', 'LEVITA (ﾚｳﾞｨﾀ)', '森の記憶', 'とき葉', 
        '森のことばIBUKI', '森のことば ウォルナット'])]['金額'].sum()
        fushi_last = cust_last[cust_last['シリーズ名'].isin(['森のことば', 'LEVITA (ﾚｳﾞｨﾀ)', '森の記憶', 'とき葉', 
        '森のことばIBUKI', '森のことば ウォルナット'])]['金額'].sum()
        # sdソファ拾えていない isin その値を含む行　true
        fushi_diff = f'{(fushi_now/total_now*100) - (fushi_last/total_last*100):0.1f} %'
        st.metric('節材比率', value=f'{fushi_now/total_now*100: 0.1f} %', delta=fushi_diff) #小数点以下1ケタ
        st.caption(f'前年 {fushi_last/total_last*100:0.1f} %')

    with col3:
        kokusanzai_now1 = cust_now[cust_now['シリーズ名'].isin(['北海道民芸家具', 'HIDA', 'Northern Forest', '北海道HMその他', 
        '杉座', 'ｿﾌｨｵ SUGI', '風のうた', 'Kinoe', 'SUWARI', 'KURINOKI'])]['金額'].sum() #SHSカバ拾えていない
        kokusanzai_last1 = cust_last[cust_last['シリーズ名'].isin(['北海道民芸家具', 'HIDA', 'Northern Forest', '北海道HMその他', 
        '杉座', 'ｿﾌｨｵ SUGI', '風のうた', 'Kinoe', 'SUWARI', 'KURINOKI'])]['金額'].sum() #SHSカバ拾えていない

        kokusanzai_now2 = cust_now[cust_now['商品コード2'].isin(['SG261M', 'SG261K', 'SG261C', 'SG261AM', 'SG261AK', 'SG261AC', 'KD201M', 'KD201K', 'KD201C', 'KD201AM', 'KD201AK', 'KD201AC'])]['金額'].sum()
        kokusanzai_last2 = cust_last[cust_last['商品コード2'].isin(['SG261M', 'SG261K', 'SG261C', 'SG261AM', 'SG261AK', 'SG261AC', 'KD201M', 'KD201K', 'KD201C', 'KD201AM', 'KD201AK', 'KD201AC'])]['金額'].sum()
        
        kokusanzai_now3 = cust_now[cust_now['商品コード3']=='HJ']['金額'].sum()
        kokusanzai_last3 = cust_last[cust_last['商品コード3']=='HJ']['金額'].sum()

        kokusanzai_now_t = kokusanzai_now1 + kokusanzai_now2 + kokusanzai_now3
        kokusanzai_last_t = kokusanzai_last1 + kokusanzai_last2 + kokusanzai_last3 

        kokusanzai_diff = f'{(kokusanzai_now_t/df_now_total*100) - (kokusanzai_last_t/df_last_total*100): 0.1f} %'
        st.metric('国産材比率', value=f'{kokusanzai_now_t/df_now_total*100: 0.1f} %', delta=kokusanzai_diff) #小数点以下1ケタ
        st.caption(f'前年 {kokusanzai_last_t/df_last_total*100: 0.1f} %')

def profit_aroma():
    col1, col2, col3 = st.columns(3)
    cust_now = df_now[df_now['得意先名']== option_customer]
    cust_last = df_last[df_last['得意先名']== option_customer]
    total_now = cust_now['金額'].sum()
    total_last = cust_last['金額'].sum()
    cost_now = cust_now['原価金額'].sum()
    cost_last = cust_last['原価金額'].sum()
    cost_last2 = f'{(total_last-cost_last)/total_last*100: 0.1f} %'
    diff = f'{((total_now-cost_now)/total_now*100) - ((total_last-cost_last)/total_last*100): 0.1f} %'
    with col1:
        st.metric('粗利率', value=f'{(total_now-cost_now)/total_now*100: 0.1f} %', delta=diff)
        st.caption(f'前年 {cost_last2}')

    with col2:
        profit = '{:,}'.format(total_now-cost_now)
        dif_profit = int((total_now-cost_now) - (total_last-cost_last))
        st.metric('粗利額', value=profit, delta=dif_profit)
        st.caption(f'前年 {total_last-cost_last}')

    with col3:
        aroma_now = cust_now[cust_now['シリーズ名'].isin(['きつつき森の研究所'])]['金額'].sum()
        aroma_last = cust_last[cust_last['シリーズ名'].isin(['きつつき森の研究所'])]['金額'].sum()
        aroma_last2 = '{:,}'.format(aroma_last)
        aroma_diff = '{:,}'.format(aroma_now- aroma_last)
        st.metric('きつつき森の研究所関連', value=('{:,}'.format(aroma_now)), delta=aroma_diff)
        st.caption(f'前年 {aroma_last2}')

def color():
    df_now_cust = df_now[df_now['得意先名']==option_customer]
    df_last_cust = df_last[df_last['得意先名']==option_customer]

    df_now_cust = df_now_cust.dropna(subset=['塗色CD'])
    df_last_cust = df_last_cust.dropna(subset=['塗色CD'])


    with st.expander('売上グラフ', expanded=False):
        col1, col2 = st.columns(2)

        with col1:
            # ***塗色別売り上げ ***
            color_now = df_now_cust.groupby('塗色CD')['金額'].sum().sort_values(ascending=False) #降順
            #color_now2 = color_now.apply('{:,}'.format) #数値カンマ区切り　注意strになる　グラフ作れなくなる
            st.markdown('###### 塗色別売上(今期)')

            # グラフ
            fig_color_now = go.Figure()
            fig_color_now.add_trace(
                go.Bar(
                    x=color_now.index,
                    y=color_now,
                    )
            )
            fig_color_now.update_layout(
                height=500,
                width=2000,
            )        
            
            st.plotly_chart(fig_color_now, use_container_width=True)

        with col2:
            # ***塗色別売り上げ ***
            color_last = df_last_cust.groupby('塗色CD')['金額'].sum().sort_values(ascending=False) #降順
            #color_last2 = color_now.apply('{:,}'.format) #数値カンマ区切り　注意strになる　グラフ作れなくなる
            st.markdown('###### 塗色別売上(前期)')

            # グラフ
            fig_color_last = go.Figure()
            fig_color_last.add_trace(
                go.Bar(
                    x=color_last.index,
                    y=color_last,
                    )
            )
            fig_color_last.update_layout(
                height=500,
                width=2000,
            )        
            
            st.plotly_chart(fig_color_last, use_container_width=True)    

    col3, col4 = st.columns(2)
    with col3:    
        # グラフ　塗色別売り上げ
        st.markdown('###### 塗色別売上構成比(今期)')
        fig_color_now2 = go.Figure(
            data=[
                go.Pie(
                    labels=color_now.index,
                    values=color_now
                    )])
        fig_color_now2.update_layout(
            showlegend=True, #凡例表示
            height=290,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_color_now2.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_color_now2, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    with col4:    
        # グラフ　塗色別売り上げ
        st.markdown('###### 塗色別売上構成比(前期)')
        fig_color_last2 = go.Figure(
            data=[
                go.Pie(
                    labels=color_last.index,
                    values=color_last
                    )])
        fig_color_last2.update_layout(
            showlegend=True, #凡例表示
            height=290,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_color_last2.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_color_last2, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

def category_color():
    # *** selectbox 商品分類2***
    category = df_now['商品分類名2'].unique()
    option_category = st.selectbox(
        'category:',
        category,   
    ) 
    categorybase_now = df_now[df_now['商品分類名2']==option_category]
    categorybase_last = df_last[df_last['商品分類名2']==option_category]
    categorybase_cust_now = categorybase_now[categorybase_now['得意先名']== option_customer]
    categorybase_cust_last = categorybase_last[categorybase_last['得意先名']== option_customer]

    col1, col2 = st.columns(2)

    with col1:
        # ***塗色別数量 ***
        color_now = categorybase_cust_now.groupby('塗色CD')['数量'].sum().sort_values(ascending=False) #降順
        #color_now2 = color_now.apply('{:,}'.format) #数値カンマ区切り　注意strになる　グラフ作れなくなる
        st.markdown('###### 塗色別数量(今期)')

        # グラフ
        fig_color_now = go.Figure()
        fig_color_now.add_trace(
            go.Bar(
                x=color_now.index,
                y=color_now,
                )
        )
        fig_color_now.update_layout(
            height=500,
            width=2000,
        )        
        
        st.plotly_chart(fig_color_now, use_container_width=True)

    with col2:
        # ***塗色別数量 ***
        color_last = categorybase_cust_last.groupby('塗色CD')['数量'].sum().sort_values(ascending=False) #降順
        #color_last2 = color_now.apply('{:,}'.format) #数値カンマ区切り　注意strになる　グラフ作れなくなる
        st.markdown('###### 塗色別数量(前期)')

        # グラフ
        fig_color_last = go.Figure()
        fig_color_last.add_trace(
            go.Bar(
                x=color_last.index,
                y=color_last,
                )
        )
        fig_color_last.update_layout(
            height=500,
            width=2000,
        )        
        
        st.plotly_chart(fig_color_last, use_container_width=True)    

    with col1:    
        # グラフ　塗色別数量
        st.markdown('###### 塗色別数量構成比(今期)')
        fig_color_now2 = go.Figure(
            data=[
                go.Pie(
                    labels=color_now.index,
                    values=color_now
                    )])
        fig_color_now2.update_layout(
            showlegend=True, #凡例表示
            height=290,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_color_now2.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_color_now2, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    with col2:    
        # グラフ　塗色別数量
        st.markdown('###### 塗色別数量構成比(前期)')
        fig_color_last2 = go.Figure(
            data=[
                go.Pie(
                    labels=color_last.index,
                    values=color_last
                    )])
        fig_color_last2.update_layout(
            showlegend=True, #凡例表示
            height=290,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_color_last2.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_color_last2, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
            
def category_fabric():
     # *** selectbox***
    category = ['ダイニングチェア', 'リビングチェア']
    option_category = st.selectbox(
        'category:',
        category,   
    ) 
    categorybase_now = df_now[df_now['商品分類名2']==option_category]
    categorybase_last = df_last[df_last['商品分類名2']==option_category]
    categorybase_cust_now = categorybase_now[categorybase_now['得意先名']== option_customer]
    categorybase_cust_last = categorybase_last[categorybase_last['得意先名']== option_customer]
    categorybase_cust_now = categorybase_cust_now[categorybase_cust_now['張地'] != ''] #空欄を抜いたdf作成
    categorybase_cust_last = categorybase_cust_last[categorybase_cust_last['張地'] != '']

    col1, col2 = st.columns(2)
    with col1:
        # ***張地別数量 ***
        fabric_now = categorybase_cust_now.groupby('張地')['数量'].sum().sort_values(ascending=False).head(12) #降順
        #fabric2 = fabric_now.apply('{:,}'.format) #数値カンマ区切り　注意strになる　グラフ作れなくなる
        
        #脚カットの場合ファブリックの位置がずれる為、行削除
        for index in fabric_now.index:
            if index in ['ｾﾐｱｰﾑﾁｪｱ', 'ｱｰﾑﾁｪｱ', 'ﾁｪｱ']:
                fabric_now.drop(index=index, inplace=True)

        st.markdown('###### 張地別数量(今期)')

        # グラフ
        fig_fabric_now = go.Figure()
        fig_fabric_now.add_trace(
            go.Bar(
                x=fabric_now.index,
                y=fabric_now,
                )
        )
        fig_fabric_now.update_layout(
            height=500,
            width=2000,
        )        
        
        st.plotly_chart(fig_fabric_now, use_container_width=True)

    with col2:
        # ***張地別売り上げ ***
        fabric_last = categorybase_cust_last.groupby('張地')['数量'].sum().sort_values(ascending=False).head(12) #降順
        #fabric2 = fabric_now.apply('{:,}'.format) #数値カンマ区切り　注意strになる　グラフ作れなくなる

        #脚カットの場合ファブリックの位置がずれる為、行削除
        for index in fabric_last.index:
            if index in ['ｾﾐｱｰﾑﾁｪｱ', 'ｱｰﾑﾁｪｱ', 'ﾁｪｱ']:
                fabric_last.drop(index=index, inplace=True)

        st.markdown('###### 張地別数量(前期)')

        # グラフ
        fig_fabric_last = go.Figure()
        fig_fabric_last.add_trace(
            go.Bar(
                x=fabric_last.index,
                y=fabric_last,
                )
        )
        fig_fabric_last.update_layout(
            height=500,
            width=2000,
        )        
        
        st.plotly_chart(fig_fabric_last, use_container_width=True)    

    with col1:
        # グラフ　張地別数量
        st.markdown('張地別数量構成比(今期)')
        fig_fabric_now2 = go.Figure(
            data=[
                go.Pie(
                    labels=fabric_now.index,
                    values=fabric_now
                    )])
        fig_fabric_now2.update_layout(
            legend=dict(
                x=-1, #x座標　グラフの左下(0, 0) グラフの右上(1, 1)
                y=0.99, #y座標
                xanchor='left', #x座標が凡例のどの位置を指すか
                yanchor='top', #y座標が凡例のどの位置を指すか
                ),
            height=290,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_fabric_now2.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_fabric_now2, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    with col2:
        # グラフ　張地別数量
        st.markdown('張地別数量構成比(前期)')
        fig_fabric_last2 = go.Figure(
            data=[
                go.Pie(
                    labels=fabric_last.index,
                    values=fabric_last
                    )])
        fig_fabric_last2.update_layout(
            legend=dict(
                x=-1, #x座標　グラフの左下(0, 0) グラフの右上(1, 1)
                y=0.99, #y座標
                xanchor='left', #x座標が凡例のどの位置を指すか
                yanchor='top', #y座標が凡例のどの位置を指すか
                ),
            height=290,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig_fabric_last2.update_traces(textposition='inside', textinfo='label+percent') 
        #inside グラフ上にテキスト表示
        st.plotly_chart(fig_fabric_last2, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅    

def series_col_fab():
    # *** selectbox 商品分類2***
    category = df_now['商品分類名2'].unique()
    option_category = st.selectbox(
        'category:',
        category,   
    ) 
    categorybase_now = df_now[df_now['商品分類名2']==option_category]
    categorybase_last = df_last[df_last['商品分類名2']==option_category]
    categorybase_cust_now = categorybase_now[categorybase_now['得意先名']== option_customer]
    categorybase_cust_last = categorybase_last[categorybase_last['得意先名']== option_customer]

    # *** シリース別塗色別数量 ***
    series_color_now = categorybase_cust_now.groupby(['シリーズ名', '塗色CD', '張地'])['数量'].sum().sort_values(ascending=False).head(20) #降順
    series_color_now2 = series_color_now.apply('{:,}'.format) #数値カンマ区切り　注意strになる　グラフ作れなくなる
    
    # **シリーズ別塗色別数量 ***
    series_color_last = categorybase_cust_last.groupby(['シリーズ名', '塗色CD', '張地'])['数量'].sum().sort_values(ascending=False).head(20) #降順
    series_color_last2 = series_color_last.apply('{:,}'.format) #数値カンマ区切り　注意strになる　グラフ作れなくなる
    
    #数量2以上に限定
    df_series_color_now2 = series_color_now[series_color_now >=2]
    df_series_color_last2 = series_color_last[series_color_last >=2]
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('###### 売れ筋ランキング 商品分類別(今期)')
        st.table(df_series_color_now2)
    with col2:
        st.write('###### 売れ筋ランキング 商品分類別(前期)')
        st.table(df_series_color_last2)

def main():
    # アプリケーション名と対応する関数のマッピング
    apps = {
        '-': None,
        '★売上 対前年比(累計)●': earnings_comparison_year,
        '★売上 対前年比(月毎)●': earnings_comparison_month,
        '平均成約単価': mean_earning_month,
        '★LD 前年比/構成比●': living_dining_comparison,
        '★LD シリーズ別/売上構成比●': living_dining_comparison_ld,
        '商品分類 シリーズ別 売上/構成比●': series,
        '★回転数 商品分類別●': item_count_category,
        '★回転数 商品分類別 月毎●': category_count_month,
        '★比率 北海道工場/節あり材/国産材●': hokkaido_fushi_kokusanzai, 
        '★比率 粗利/アロマ関連●': profit_aroma,
        '塗色別　売上構成比': color,
        '塗色別 数量/構成比/商品分類別●': category_color,
        '張地別 数量/構成比●': category_fabric,
        '売れ筋ランキング 商品分類別/塗色/張地●': series_col_fab
  
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