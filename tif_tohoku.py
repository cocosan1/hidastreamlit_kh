from logging import debug
import pandas as pd
import numpy as np
from pandas.core.frame import DataFrame
import streamlit as st
import plotly.figure_factory as ff
import plotly.graph_objects as go
import openpyxl
from streamlit.elements import metric

import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

st.set_page_config(page_title='売り上げ分析（TIF 一覧）')
st.markdown('#### 売り上げ分析（TIF 東北）')

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
    path_now, sheet_name='受注委託移動在庫生産照会', usecols=[1, 3, 6, 9, 10, 14, 15, 16, 28, 31, 42, 50, 51, 52]) #index　ナンバー不要　index_col=0

# ***前期受注***
path_last = os.path.join(cwd, 'data', 'kita78j.xlsx')
df_last = pd.read_excel(
    path_last, sheet_name='受注委託移動在庫生産照会', usecols=[1, 3, 6, 9, 10, 14, 15, 16, 28, 31, 42, 50, 51, 52]) #index　ナンバー不要　index_col=0

# *** 出荷月、受注月列の追加***
df_now['出荷月'] = df_now['出荷日'].dt.month
df_now['受注月'] = df_now['受注日'].dt.month 
df_now['商品コード2'] = df_now['商品コード'].map(lambda x: x.split()[0]) 

df_last['出荷月'] = df_last['出荷日'].dt.month
df_last['受注月'] = df_last['受注日'].dt.month
df_last['商品コード2'] = df_last['商品コード'].map(lambda x: x.split()[0])

# ***INT型への変更***
df_now[['数量', '単価', '金額', '出荷倉庫', '原価金額', '出荷月', '受注月']] = \
    df_now[['数量', '単価', '金額', '出荷倉庫', '原価金額', '出荷月', '受注月']].fillna(0).astype('int64')
#fillna　０で空欄を埋める

df_last[['数量', '単価', '金額', '出荷倉庫', '原価金額', '出荷月', '受注月']] = \
    df_last[['数量', '単価', '金額', '出荷倉庫', '原価金額', '出荷月', '受注月']].fillna(0).astype('int64')
#fillna　０で空欄を埋める 

df_now_total = df_now['金額'].sum()
df_last_total = df_last['金額'].sum()

customer_list = ['㈱東京ｲﾝﾃﾘｱ 下田店', '㈱東京ｲﾝﾃﾘｱ 郡山店', '㈱東京ｲﾝﾃﾘｱ 山形店', '㈱東京ｲﾝﾃﾘｱ 秋田店', '㈱東京ｲﾝﾃﾘｱ 盛岡店',\
        '㈱東京ｲﾝﾃﾘｱ 仙台港本店', '㈱東京ｲﾝﾃﾘｱ 仙台泉店', '㈱東京ｲﾝﾃﾘｱ 仙台南店', '㈱東京ｲﾝﾃﾘｱ 福島店']

original_list = ['森の記憶', 'LEVITA (ﾚｳﾞｨﾀ)', '悠々', 'とき葉', '青葉', '東京ｲﾝﾃﾘｱｵﾘｼﾞﾅﾙ']

def earnings_comparison():
    
    index = []
    earnings_now = []
    earnings_last = []
    comparison_rate = []
    comparison_diff = []

    for customer in customer_list:
        index.append(customer)
        cust_earnings_total_now = df_now[df_now['得意先名']==customer]['金額'].sum()
        cust_earnings_total_last = df_last[df_last['得意先名']==customer]['金額'].sum()
        earnings_rate_culc = f'{cust_earnings_total_now/cust_earnings_total_last*100: 0.1f} %'
        comaparison_diff_culc = cust_earnings_total_now - cust_earnings_total_last

        earnings_now.append(cust_earnings_total_now)
        earnings_last.append(cust_earnings_total_last)
        comparison_rate.append(earnings_rate_culc)
        comparison_diff.append(comaparison_diff_culc)
    df_earnings_comparison = pd.DataFrame(list(zip(earnings_now, earnings_last, comparison_rate, comparison_diff)), index=index, columns=['今期', '前期', '対前年比', '対前年差'])    
    
    with st.expander('一覧', expanded=False):
        st.dataframe(df_earnings_comparison)

    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig = go.Figure()
    #今期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=df_earnings_comparison.index,
            y=df_earnings_comparison['今期'],
            text=round(df_earnings_comparison['今期']/10000),
            textposition="outside", 
            name='今期')
    )
    #前期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=df_earnings_comparison.index,
            y=df_earnings_comparison['前期'],
            text=round(df_earnings_comparison['前期']/10000),
            textposition="outside", 
            name='前期'
            )
    )

    #レイアウト設定     
    fig.update_layout(
        title='売上一覧（累計）',
        showlegend=True, #凡例表示
        )
 
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig, use_container_width=True) 

def earnings_comparison_month():

    # *** selectbox 得意先名***
    month = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    option_month = st.selectbox(
    '受注月:',
    month,   
) 

    index = []
    earnings_now = []
    earnings_last = []
    comparison_rate = []
    comparison_diff = []

    df_now2 =df_now[df_now['得意先名'].isin(customer_list)]
    df_last2 =df_last[df_last['得意先名'].isin(customer_list)]
    df_now_month = df_now2[df_now2['受注月']==option_month]
    df_last_month = df_last2[df_last2['受注月']==option_month]

    earnings_now_total = df_now_month['金額'].sum()
    earnings_last_total = df_last_month['金額'].sum()
    comparison_rate_total = f'{earnings_now_total/earnings_last_total *100: 0.1f} %'
    comparison_diff_total =earnings_now_total - earnings_last_total
    data_list = [earnings_now_total, earnings_last_total, comparison_rate_total, comparison_diff_total]
    earnings_comparison_total_list = pd.DataFrame(data=[[earnings_now_total, earnings_last_total, comparison_rate_total, comparison_diff_total]], columns=['今期', '前期', '対前年比', '対前年差'])
    st.markdown("###### 合計")
    st.table(earnings_comparison_total_list)

    for customer in customer_list:
        index.append(customer)
        cust_earnings_total_now_month = df_now_month[df_now_month['得意先名']==customer]['金額'].sum()
        cust_earnings_total_last_month = df_last_month[df_last_month['得意先名']==customer]['金額'].sum()
        earnings_rate_culc = f'{cust_earnings_total_now_month/cust_earnings_total_last_month *100: 0.1f} %'
        comaparison_diff_culc = cust_earnings_total_now_month - cust_earnings_total_last_month

        earnings_now.append(cust_earnings_total_now_month)
        earnings_last.append(cust_earnings_total_last_month)
        comparison_rate.append(earnings_rate_culc)
        comparison_diff.append(comaparison_diff_culc)
    df_earnings_comparison = pd.DataFrame(list(zip(earnings_now, earnings_last, comparison_rate, comparison_diff)), index=index, columns=['今期', '前期', '対前年比', '対前年差'])
    
    with st.expander('一覧', expanded=False):
        st.markdown("###### 得意先別")  
        st.dataframe(df_earnings_comparison)

    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig = go.Figure()
    #今期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=df_earnings_comparison.index,
            y=df_earnings_comparison['今期'],
            text=round(df_earnings_comparison['今期']/10000),
            textposition="outside", 
            name='今期')
    )
    #前期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=df_earnings_comparison.index,
            y=df_earnings_comparison['前期'],
            text=round(df_earnings_comparison['前期']/10000),
            textposition="outside", 
            name='前期'
            )
    )
    #レイアウト設定     
    fig.update_layout(
        title='売上一覧（月別）',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig, use_container_width=True) 

def earnings_comparison_month_suii():

    month_list = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    columns_list = ['今期', '前期', '対前年差', '対前年比']

    earnings_now = []
    earnings_last = []
    earnings_diff = []
    earnings_rate = []

    #グラフ用int
    earnings_now2 = []
    earnings_last2 = []

    with st.form('選択'):
        selected_list = st.multiselect(
            '得意先を選択(複数可)',
            customer_list)
        
        submitted = st.form_submit_button('submit')

    if submitted:
        #可視化
        #グラフを描くときの土台となるオブジェクト
        fig = go.Figure()    

        for cust in selected_list:  
            
            df_now_cust = df_now[df_now['得意先名']==cust]
            df_last_cust = df_last[df_last['得意先名']==cust]

            for month in month_list:
                earnings_month_now = df_now_cust[df_now_cust['受注月'].isin([month])]['金額'].sum()
                earnings_month_last = df_last_cust[df_last_cust['受注月'].isin([month])]['金額'].sum()
                earnings_diff_culc = earnings_month_now - earnings_month_last
                earnings_rate_culc = f'{earnings_month_now / earnings_month_last * 100: 0.1f} %'

                earnings_now.append('{:,}'.format(earnings_month_now))
                earnings_last.append('{:,}'.format(earnings_month_last))
                earnings_diff.append('{:,}'.format(earnings_diff_culc))
                earnings_rate.append(earnings_rate_culc)

                earnings_now2.append(earnings_month_now)
                earnings_last2.append(earnings_month_last)

                earnings_now3 = []
                earnings_last3 = []
                for i in range(len(earnings_now2)):
                    round_now = round(earnings_now2[i] /10000)
                    earnings_now3.append(round_now)
                    round_last = round(earnings_last2[i] /10000)
                    earnings_last3.append(round_last)


            fig.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=earnings_now2,
                    mode = 'lines+markers+text', #値表示
                    text=earnings_now3,
                    textposition="top center", 
                    name=cust)
            ) 

            earnings_now2 = []
            earnings_last2 = []


        #レイアウト設定     
        fig.update_layout(
            title='月別売上',
            showlegend=True #凡例表示
        )
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
        st.plotly_chart(fig, use_container_width=True)        

    # df_earnings_month = pd.DataFrame(list(zip(earnings_now, earnings_last, earnings_diff, earnings_rate)), columns=columns_list, index=month_list)
        

def ld_earnings_comp():

    index = []
    l_earnings = [] #リニング売り上げ
    l_comp = [] #リビング比率

    d_earnings = [] #ダイニング売り上げ
    d_comp = [] #ダイニング比率

    o_earnings = [] #その他売り上げ
    o_comp = [] #その他比率

    #数値のまま格納
    sum_list = []

    for customer in customer_list:
        index.append(customer)

        df_now_cust = df_now[df_now['得意先名']==customer]

        df_now_cust_sum = df_now_cust['金額'].sum() #得意先売り上げ合計

        df_now_cust_sum_l = df_now_cust[df_now_cust['商品分類名2'].isin(['クッション', 'リビングチェア', 'リビングテーブル'])]['金額'].sum()
        l_earnings.append('{:,}'.format(df_now_cust_sum_l))
        l_comp_culc = f'{df_now_cust_sum_l/df_now_cust_sum*100:0.1f} %'
        l_comp.append(l_comp_culc)

        df_now_cust_sum_d = df_now_cust[df_now_cust['商品分類名2'].isin(['ダイニングテーブル', 'ダイニングチェア', 'ベンチ'])]['金額'].sum()
        d_earnings.append('{:,}'.format(df_now_cust_sum_d))
        d_comp_culc = f'{df_now_cust_sum_d/df_now_cust_sum*100:0.1f} %'
        d_comp.append(d_comp_culc)

        df_now_cust_sum_o = df_now_cust[df_now_cust['商品分類名2'].isin(['キャビネット類', 'その他テーブル', '雑品・特注品', 'その他椅子', 'デスク', '小物・その他'])]['金額'].sum()
        o_earnings.append('{:,}'.format(df_now_cust_sum_o))
        o_comp_culc = f'{df_now_cust_sum_o/df_now_cust_sum*100:0.1f} %'
        o_comp.append(o_comp_culc)

        temp_list = [df_now_cust_sum_l, df_now_cust_sum_d]
        sum_list.append(temp_list)

    st.write('構成比')
    df_comp = pd.DataFrame(list(zip(l_comp, d_comp, o_comp)), index=index, columns=['L', 'D', 'その他'])
    st.dataframe(df_comp)

    #計算用df
    df_sum = pd.DataFrame(sum_list, index=customer_list, columns=['Living', 'Dining'])

    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig = go.Figure()
    #今期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=df_sum.index,
            y=df_sum['Living'],
            text=round(df_sum['Living']/10000),
            textposition="outside", 
            name='Living')
    )
    #前期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=df_sum.index,
            y=df_sum['Dining'],
            text=round(df_sum['Dining']/10000),
            textposition="outside", 
            name='Dining'
            )
    )
    #レイアウト設定     
    fig.update_layout(
        title='LD別売上（累計）',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig, use_container_width=True) 

def ld_comp():

    index = []
    sum_list = [] #売り上げ
    
    for customer in customer_list:
        index.append(customer)

        df_now_cust = df_now[df_now['得意先名']==customer]
        now_cust_sum_l = df_now_cust[df_now_cust['商品分類名2'].isin(\
            ['クッション', 'リビングチェア', 'リビングテーブル'])]['金額'].sum()
        now_cust_sum_d = df_now_cust[df_now_cust['商品分類名2'].isin(\
            ['ダイニングテーブル', 'ダイニングチェア', 'ベンチ'])]['金額'].sum()
        temp_list = [now_cust_sum_l, now_cust_sum_d]
        sum_list.append(temp_list)

    df_results = pd.DataFrame(sum_list, index=index, columns=['Living', 'Dining'])

    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig = go.Figure()
    #今期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=df_results.index,
            y=df_results['Living'],
            text=round(df_results['Living']/10000),
            textposition="outside", 
            name='Living')
    )
    #前期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=df_results.index,
            y=df_results['Dining'],
            text=round(df_results['Dining']/10000),
            textposition="outside", 
            name='Dining'
            )
    )
    #レイアウト設定     
    fig.update_layout(
        title='LD別売上（累計）',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig, use_container_width=True) 

def series_comp():
    index = []
    sum_list = [] #売り上げ
    
    series_list = df_now['シリーズ名'].unique()
    selected_series = st.selectbox(
        'series:',
        series_list,   
    ) 
    for customer in customer_list:
        index.append(customer)

        df_now_cust = df_now[df_now['得意先名']==customer]
        df_now_cust_sum_l = df_now_cust[df_now_cust['商品分類名2'].isin(\
            ['クッション', 'リビングチェア', 'リビングテーブル'])]
        df_now_cust_sum_d = df_now_cust[df_now_cust['商品分類名2'].isin(\
            ['ダイニングテーブル', 'ダイニングチェア', 'ベンチ'])]
        lseiries_sum = df_now_cust_sum_l[df_now_cust_sum_l['シリーズ名']==selected_series]['金額'].sum() 
        dseiries_sum = df_now_cust_sum_d[df_now_cust_sum_d['シリーズ名']==selected_series]['金額'].sum()
        temp_list = [lseiries_sum, dseiries_sum]
        sum_list.append(temp_list)

    df_results = pd.DataFrame(sum_list, index=index, columns=['Living', 'Dining'])

    #可視化
    #グラフを描くときの土台となるオブジェクト
    fig = go.Figure()
    #今期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=df_results.index,
            y=df_results['Living'],
            text=round(df_results['Living']/10000),
            textposition="outside", 
            name='Living')
    )
    #前期のグラフの追加
    fig.add_trace(
        go.Bar(
            x=df_results.index,
            y=df_results['Dining'],
            text=round(df_results['Dining']/10000),
            textposition="outside", 
            name='Dining'
            )
    )
    #レイアウト設定     
    fig.update_layout(
        title='シリーズ別売上（累計）',
        showlegend=True #凡例表示
    )
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    st.plotly_chart(fig, use_container_width=True) 




def main():
    # アプリケーション名と対応する関数のマッピング
    apps = {
        '-': None,
        '売上: 累計': earnings_comparison,
        '売上: 月別': earnings_comparison_month,
        '売上: 月別推移': earnings_comparison_month_suii,
        'LD別売上: 累計': ld_comp,
        'シリーズ別売上: 累計': series_comp
 
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