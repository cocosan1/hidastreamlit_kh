import pandas as pd
import numpy as np
from pandas.core.frame import DataFrame
import streamlit as st
import plotly.figure_factory as ff
import plotly.graph_objects as go

import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

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
selected_cust = st.selectbox(
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
    path_cust, sheet_name='貼りつけ', usecols=[15, 42, 43, 51, 52, 54]) #index　ナンバー不要　index_col=0


# ***データ調整***
df = df.dropna(how='any') #一つでも欠損値のある行を削除　all　全て欠損値の行を削除
df['金額'] = df['金額'].astype(int) #float →　int
df['年代'] = df['年代'].astype(int) #float →　int
df['月'] = df['月'].astype(int) #float →　int

def age_ratio():
    col1, col2 = st.columns(2)
    with col1:
        # ***DataFrame　年齢層別全体 ***
        comp_age = df[['年代', '金額']].groupby('年代').sum() #年代で集計

        st.write('年齢構成比 全体')
        st.dataframe(comp_age)
    with col2:
        # ***グラフ　年齢層別全体 ***
        fig = go.Figure(
            data=[
                go.Pie(
                    labels=comp_age.index,
                    values=comp_age['金額'],
                    )])
        fig.update_layout(
            showlegend=True, #凡例表示
            height=150,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        fig.update_traces(textposition='inside', textinfo='label+percent') #inside グラフ上にテキスト表示
        st.plotly_chart(fig, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    with col1:
        # ***年齢構成比　ダイニング***
        st.write('年齢構成比 ダイニング')

        comp_dage = df[df['LD別']=='D'].groupby('年代')['金額'].sum()
        st.dataframe(comp_dage)
    with col2:
        #グラフ　年齢別構成比　ダイニング
        figd = go.Figure(
            data=[
                go.Pie(
                    labels=comp_dage.index,
                    values=comp_dage,
                    )])
        figd.update_layout(
            showlegend=True, #凡例表示
            height=150,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        figd.update_traces(textposition='inside', textinfo='label+percent') #inside グラフ上にテキスト表示
        st.plotly_chart(figd, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
    with col1:
        # ***年齢構成比　リビング***
        st.write('年齢構成比 リビング')

        comp_lage = df[df['LD別']=='L'].groupby('年代')['金額'].sum()
        st.dataframe(comp_lage)
    
    with col2:
        #グラフ　年齢別構成比　リビング
        figl = go.Figure(
            data=[
                go.Pie(
                    labels=comp_lage.index,
                    values=comp_lage,
                    )])
        figl.update_layout(
            showlegend=True, #凡例表示
            height=150,
            margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
            )
        figl.update_traces(textposition='inside', textinfo='label+percent') #inside グラフ上にテキスト表示
        st.plotly_chart(figl, use_container_width=True) 
        #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

def seiriesbase_ageratio_d():

    # ***シリーズベース　年代別構成比　ダイニング　全項目俯瞰 ***
    st.write('シリーズベース　年代別構成比　ダイニング')
    series_age_alld = pd.crosstab(df[df['LD別']=='D']['シリーズ名'], df[df['LD別']=='D']['年代'], values=df[df['LD別']=='D']['金額'], aggfunc='sum').fillna(0).astype('int64')

    st.dataframe(series_age_alld)

    # ***シリーズベース 年代別構成比 ダイニング***
    st.write('ダイニング　シリーズベース年代別構成比')

    comp_ageseries_d = df[df['LD別']=='D'].groupby(['シリーズ名', '年代'])['金額'].sum()
    st.dataframe(comp_ageseries_d)
    # selectbox
    series_list_d = df[df['LD別']=='D']['シリーズ名'].unique()
    option_d = st.selectbox(
        'series:',
        series_list_d,   
    )

    st.write('You selected: ', option_d)

    #グラフ　シリーズベース年代別構成比 ダイニング
    series_age_d = df[(df['LD別']=='D')&(df['シリーズ名']==option_d)].groupby('年代').sum()
    fig_das = go.Figure(
        data=[
            go.Pie(
                labels=series_age_d.index,
                values=df[(df['LD別']=='D')&(df['シリーズ名']==option_d)].groupby('年代')['金額'].sum(),
                )])
    fig_das.update_layout(
        showlegend=True, #凡例表示
        height=200,
        margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
        )
    fig_das.update_traces(textposition='inside', textinfo='label+percent') #inside グラフ上にテキスト表示
    st.plotly_chart(fig_das, use_container_width=True) 
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

def seiriesbase_ageratio_l():

    # ***シリーズベース　年代別構成比　リビング　全項目俯瞰 ***
    st.write('シリーズベース　年代別構成比　リビング')
    series_age_alll = pd.crosstab(df[df['LD別']=='L']['シリーズ名'], df[df['LD別']=='L']['年代'], values=df[df['LD別']=='L']['金額'], aggfunc='sum').fillna(0).astype('int64')

    st.dataframe(series_age_alll)
    
    # ***シリーズベース 年代別構成比 リビング***
    st.write('リビング　シリーズベース年代別構成比')

    comp_ageseries_l = df[df['LD別']=='L'].groupby(['シリーズ名', '年代'])['金額'].sum()
    st.dataframe(comp_ageseries_l)

    # selectbox
    series_list_l = df[df['LD別']=='L']['シリーズ名'].unique()

    option_l = st.selectbox(
        'series:',
        series_list_l,   
    )
    st.write('You selected: ', option_l)

    #グラフ　シリーズベース 年代別構成比 リビング
    series_age_l = df[(df['LD別']=='L')&(df['シリーズ名']==option_l)].groupby('年代').sum()
    fig_las = go.Figure(
        data=[
            go.Pie(
                labels=series_age_l.index,
                values=df[(df['LD別']=='L')&(df['シリーズ名']==option_l)].groupby('年代')['金額'].sum(),
                )])
    fig_las.update_layout(
        showlegend=True, #凡例表示
        height=200,
        margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
        )
    fig_las.update_traces(textposition='inside', textinfo='label+percent') #inside グラフ上にテキスト表示
    st.plotly_chart(fig_las, use_container_width=True) 
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

def agebase_seriesratio_d():
    # *** 年齢ベース　シリーズ別構成比　ダイニング ***
    st.write('年齢層ベース　シリーズ別構成比 ダイニング')

    comp_agebase_d = df[df['LD別']=='D'].groupby(['年代', 'シリーズ名'])['金額'].sum()

    # selectbox
    series_list_aged = df[df['LD別']=='D']['年代'].unique()

    option_aged = st.selectbox(
        'series:',
        series_list_aged,   
    )
    st.write('You selected: ', option_aged)

    #グラフ　年齢ベース　シリーズ別年代別構成比
    series_names_d = df[(df['LD別']=='D')&(df['年代']==option_aged)].groupby('シリーズ名').sum()
    fig_aged = go.Figure(
        data=[
            go.Pie(
                labels=series_names_d.index,
                values=df[(df['LD別']=='D')&(df['年代']==option_aged)].groupby('シリーズ名')['金額'].sum(),
                )])
    fig_aged.update_layout(
        showlegend=True, #凡例表示
        height=200,
        margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
        )
    fig_aged.update_traces(textposition='inside', textinfo='label+percent') #inside グラフ上にテキスト表示
    st.plotly_chart(fig_aged, use_container_width=True) 
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    # ***年齢層ベース　シリーズ別構成比 リビング ***
def agebase_seriesratio_l():
    st.write('年齢層ベース　シリーズ別構成比 リビング')

    comp_agebase_l = df[df['LD別']=='L'].groupby(['年代', 'シリーズ名'])['金額'].sum()
    st.dataframe(comp_agebase_l)

    # selectbox
    series_list_agel = df[df['LD別']=='L']['年代'].unique()

    option_agel = st.selectbox(
        'series:',
        series_list_agel,   
    )
    st.write('You selected: ', option_agel)

    #グラフ　リビング　シリーズ別年代別構成比
    series_names_l = df[(df['LD別']=='L')&(df['年代']==option_agel)].groupby('シリーズ名').sum()
    fig_agel = go.Figure(
        data=[
            go.Pie(
                labels=series_names_l.index,
                values=df[(df['LD別']=='L')&(df['年代']==option_agel)].groupby('シリーズ名')['金額'].sum(),
                )])
    fig_agel.update_layout(
        showlegend=True, #凡例表示
        height=200,
        margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
        )
    fig_agel.update_traces(textposition='inside', textinfo='label+percent') #inside グラフ上にテキスト表示
    st.plotly_chart(fig_agel, use_container_width=True) 
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

def person_ratio():
    # ***担当者別売り上げ ***
    st.write('担当者別売り上げ')

    person = df.groupby('取引先担当')['金額'].sum()
    person = person.sort_values(ascending=False)
    with st.expander('詳細', expanded=False):
        st.dataframe(person)

# 担当者別売上　グラフ
    fig_person = go.Figure()
    fig_person.add_trace(
        go.Bar(
            x=person.index,
            y=person,
            )
    )
    st.plotly_chart(fig_person, use_container_width=True) 

    # ***担当者別売り上げ構成比 ***
    st.write('担当者別売り上げ構成比')

    #担当者別売上構成比　グラフ
    fig_person_ratio = go.Figure(
        data=[
            go.Pie(
                labels=person.index,
                values=person,
                )])
    fig_person_ratio.update_layout(
        showlegend=True, #凡例表示
        height=200,
        margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
        )
    fig_person_ratio.update_traces(textposition='inside', textinfo='label+percent') 
    #inside グラフ上にテキスト表示
    st.plotly_chart(fig_person_ratio, use_container_width=True) 
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    # *** 月別販売者数 ***

    st.write('月別販売者数')

    count_person = df.groupby('月')['取引先担当'].nunique()

    #月別販売者数　グラフ
    fig_count_person = go.Figure()
    fig_count_person.add_trace(
        go.Bar(
            x=count_person.index,
            y=count_person,
            )
    )
    fig_count_person.update_yaxes(tick0=0, dtick=1)#0開始　目盛り単位1
    st.plotly_chart(fig_count_person, use_container_width=True)

def personbase_series():
    # ***販売者ベース　シリーズ売り上げ　ダイニング***

    st.write('販売者ベース　シリーズ別売り上げ　ダイニング')
    # selectbox
    person_series_dlist = df[df['LD別']=='D']['取引先担当'].unique()

    op_person_series_dlist = st.selectbox(
        'series:',
        person_series_dlist,   
    )
    st.write('You selected: ', op_person_series_dlist)

    #グラフ　販売者ベース　シリーズ売り上げ　ダイニング
    series_names_d = df[(df['LD別']=='D')&(df['取引先担当']==op_person_series_dlist)].groupby('シリーズ名').sum()
    fig_person_series_d = go.Figure(
        data=[
            go.Pie(
                labels=series_names_d.index,
                values=df[(df['LD別']=='D')&(df['取引先担当']==op_person_series_dlist)].groupby('シリーズ名')['金額'].sum(),
                )])
    fig_person_series_d.update_layout(
        showlegend=True, #凡例表示
        height=200,
        margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
        )
    fig_person_series_d.update_traces(textposition='inside', textinfo='label+percent') #inside グラフ上にテキスト表示
    st.plotly_chart(fig_person_series_d, use_container_width=True) 
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

    # ***販売者ベース　シリーズ売り上げ　リビング***

    st.write('販売者ベース　シリーズ別売り上げ　リビング')
    # selectbox
    person_series_llist = df[df['LD別']=='L']['取引先担当'].unique()

    op_person_series_llist = st.selectbox(
        'series:',
        person_series_llist,   
    )
    st.write('You selected: ', op_person_series_llist)

    #グラフ　販売者ベース　シリーズ売り上げ　リビング
    series_names_l = df[(df['LD別']=='L')&(df['取引先担当']==op_person_series_llist)].groupby('シリーズ名').sum()
    fig_person_series_l = go.Figure(
        data=[
            go.Pie(
                labels=series_names_l.index,
                values=df[(df['LD別']=='L')&(df['取引先担当']==op_person_series_llist)].groupby('シリーズ名')['金額'].sum(),
                )])
    fig_person_series_l.update_layout(
        showlegend=True, #凡例表示
        height=200,
        margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
        )
    fig_person_series_l.update_traces(textposition='inside', textinfo='label+percent') #inside グラフ上にテキスト表示
    st.plotly_chart(fig_person_series_l, use_container_width=True) 
    #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

def main():
    # アプリケーション名と対応する関数のマッピング
    apps = {
        '-': None,
        '年齢構成比 全体/D／L★': age_ratio,
        'シリーズベース 年齢構成比 D': seiriesbase_ageratio_d,
        'シリーズベース 年齢構成比 L': seiriesbase_ageratio_l,
        '年齢ベース シリーズ別構成比 D': agebase_seriesratio_d,
        '年齢ベース シリーズ別構成比 L': agebase_seriesratio_l,
        '担当者別売上★': person_ratio,
        '担当者ベース シリーズ別売上': personbase_series,

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





