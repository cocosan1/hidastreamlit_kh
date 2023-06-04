import pandas as pd
import datetime
import streamlit as st
import os
import plotly.graph_objects as go


#************************************************************
#index商品分類/col得意先/val売上
@st.cache_data(ttl=datetime.timedelta(hours=1))
def make_data_cust(df):
    #LD分類カラムの新設
    cate_list = []
    for cate in df['商品分類名2']:
        if cate in ['ダイニングテーブル', 'ダイニングチェア', 'ベンチ']:
            cate_list.append('d')
        elif cate in ['リビングチェア', 'クッション', 'リビングテーブル']:
            cate_list.append('l')
        else:
            cate_list.append('none') 

    df['category'] = cate_list
    #シリーズ名＋LD　カラムの新設
    df['series2'] = df['category'] + '_' + df['シリーズ名']

    #伝票NOカラムの作成
    df['伝票番号2'] = df['伝票番号'].map(lambda x: x[:8])

    #必要なカラムだけに絞る
    df = df[['得意先名', '金額', 'series2', '伝票番号2']]

    # #series2のリスト作成
    # a_list = []
    # for a in df['series2'].unique():
    #     if 'none' in a:
    #         continue
    #     elif 'その他' in a:
    #         continue
    #     else:
    #         a_list.append(a)
    #     a_list.sort() 
    
    #商品リストの絞込み　手作業
    selected_list = ['l_森のことば',
    'l_穂高',
    'd_侭 JIN',
    'd_SEOTO',
    'd_L-CHAIR',
    'd_クレセント',
    'd_森のことば',
    'd_TUGUMI',
    'd_森のことばIBUKI',
    'l_CHIGUSA(ﾁｸﾞｻ）',
    'd_風のうた',
    'l_VIOLA (ｳﾞｨｵﾗ)',
    'l_SEOTO',
    'l_風のうた',
    'd_ALMO (ｱﾙﾓ)',
    'd_nae',
    'd_穂高',
    'd_PRESCELTO (ﾌﾟﾚｼｪﾙﾄ）',
    'd_VIOLA (ｳﾞｨｵﾗ)',
    'l_ｽﾀﾝﾀﾞｰﾄﾞｺﾚｸｼｮﾝ',
    'd_YURURI',
    'l_SEION 静穏',
    'd_SEOTO-EX',
    'd_BAGUETTE LB(ﾊﾞｹｯﾄｴﾙﾋﾞｰ)',
    'd_tsubura',
    'l_杜の詩',
    'd_ﾆｭｰﾏｯｷﾝﾚｲ',
    'd_COBRINA',
    'l_YURURI',
    'd_Northern Forest',
    'l_nae',
    'l_PRESCELTO (ﾌﾟﾚｼｪﾙﾄ）',
    'l_森のことばIBUKI',
    'd_AWASE',
    'l_AWASE',
    'd_CHIGUSA(ﾁｸﾞｻ）',
    'd_HIDA',
    'd_Kinoe',
    'l_Northern Forest',
    'l_Kinoe',
    'd_円空',
    'l_ﾆｭｰﾏｯｷﾝﾚｲ',
    'l_TUGUMI',
    'l_森のことば ウォルナット',
    'l_SEOTO-EX',
    'd_杜の詩',
    'l_ALMO (ｱﾙﾓ)',
    'd_侭SUGI',
    'd_森のことば ウォルナット',
    'l_COBRINA',]
    selected_list.sort()

    #得意先毎に集計
    df_calc = pd.DataFrame(index=selected_list)
    for cust in df['得意先名'].unique():
        val_list = []
        df_cust = df[df['得意先名']==cust]
        for cate in selected_list:
            df_cust2 = df_cust[df_cust['series2']==cate]
            if len(df_cust2) == 0:
                val_list.append(0)
            else:
                val = df_cust2['金額'].sum()
                val_list.append(val)
        df_temp = pd.DataFrame(val_list, columns=[cust], index=selected_list)
        df_calc = df_calc.merge(df_temp, left_index=True, right_index=True, how='outer') 

    df_calc.loc['sales'] = df_calc.sum()
    
    return df_calc
#*******************************************************************************item
@st.cache_data(ttl=datetime.timedelta(hours=1))
def make_data_item(df):
    #dfをtargetで絞込み

    #LD分類
    cate_list = []
    for cate in df['商品分類名2']:
        if cate in ['ダイニングテーブル', 'ダイニングチェア', 'ベンチ']:
            cate_list.append('d')
        elif cate in ['リビングチェア', 'クッション', 'リビングテーブル']:
            cate_list.append('l')
        else:
            cate_list.append('none') 
    
    #分類列を追加
    df['category'] = cate_list 
    #noneを削除
    df = df[df['category']!='none']
    #シリーズ名＋LD分類
    df['series2'] = df['category'] + '_' + df['シリーズ名']

    #必要なカラムに絞る
    df = df[['得意先名', '金額', 'series2']] 

    #商品リストの絞込み　手作業
    selected_list = ['d_ALMO (ｱﾙﾓ)',
                    'd_AWASE',
                    'd_BAGUETTE LB(ﾊﾞｹｯﾄｴﾙﾋﾞｰ)',
                    'd_CHIGUSA(ﾁｸﾞｻ）',
                    'd_COBRINA',
                    'd_HIDA',
                    'd_Kinoe',
                    'd_L-CHAIR',
                    'd_Northern Forest',
                    'd_PRESCELTO (ﾌﾟﾚｼｪﾙﾄ）',
                    'd_SEOTO',
                    'd_SEOTO-EX',
                    'd_TUGUMI',
                    'd_VIOLA (ｳﾞｨｵﾗ)',
                    'd_YURURI',
                    'd_nae',
                    'd_tsubura',
                    'd_クレセント',
                    'd_侭 JIN',
                    'd_侭SUGI',
                    'd_円空',
                    'd_杜の詩',
                    'd_森のことば',
                    'd_森のことば ウォルナット',
                    'd_森のことばIBUKI',
                    'd_穂高',
                    'd_風のうた',
                    'd_ﾆｭｰﾏｯｷﾝﾚｲ',
                    'l_ALMO (ｱﾙﾓ)',
                    'l_AWASE',
                    'l_CHIGUSA(ﾁｸﾞｻ）',
                    'l_COBRINA',
                    'l_Kinoe',
                    'l_Northern Forest',
                    'l_PRESCELTO (ﾌﾟﾚｼｪﾙﾄ）',
                    'l_SEION 静穏',
                    'l_SEOTO',
                    'l_SEOTO-EX',
                    'l_TUGUMI',
                    'l_VIOLA (ｳﾞｨｵﾗ)',
                    'l_YURURI',
                    'l_nae',
                    'l_杜の詩',
                    'l_森のことば',
                    'l_森のことば ウォルナット',
                    'l_森のことばIBUKI',
                    'l_穂高',
                    'l_風のうた',
                    'l_ｽﾀﾝﾀﾞｰﾄﾞｺﾚｸｼｮﾝ',
                    'l_ﾆｭｰﾏｯｷﾝﾚｲ'
                    ]
    

    #商品毎に集計
    df_calc = pd.DataFrame(index=selected_list)
    
    val_list = []
    for cate in selected_list:
        df_cate2 = df[df['series2']==cate]
        if len(df_cate2) == 0:
            val_list.append(0)
        else:
            val = df_cate2['金額'].sum()
            val_list.append(val)
    df_calc = pd.DataFrame(val_list, columns=['合計'], index=selected_list)

    df_calc.loc['sales'] = df_calc.sum()

    return df_calc
#**********************************************************************************item target
@st.cache_data(ttl=datetime.timedelta(hours=1))
def make_data_item_target(target, df):
    #dfをtargetで絞込み
    df = df[df['得意先名'] == target]

    # target_name = f'{target}_new'
    # df.rename(columns={'得意先名', target_name})
    #LD分類
    cate_list = []
    for cate in df['商品分類名2']:
        if cate in ['ダイニングテーブル', 'ダイニングチェア', 'ベンチ']:
            cate_list.append('d')
        elif cate in ['リビングチェア', 'クッション', 'リビングテーブル']:
            cate_list.append('l')
        else:
            cate_list.append('none') 
    
    #分類列を追加
    df['category'] = cate_list 

    #noneを削除
    df = df[df['category']!='none']

    #シリーズ名＋LD分類
    df['series2'] = df['category'] + '_' + df['シリーズ名']

    #必要なカラムに絞る
    df = df[['得意先名', '金額', 'series2']] 

    #商品リストの絞込み　手作業
    selected_list = ['d_ALMO (ｱﾙﾓ)',
                    'd_AWASE',
                    'd_BAGUETTE LB(ﾊﾞｹｯﾄｴﾙﾋﾞｰ)',
                    'd_CHIGUSA(ﾁｸﾞｻ）',
                    'd_COBRINA',
                    'd_HIDA',
                    'd_Kinoe',
                    'd_L-CHAIR',
                    'd_Northern Forest',
                    'd_PRESCELTO (ﾌﾟﾚｼｪﾙﾄ）',
                    'd_SEOTO',
                    'd_SEOTO-EX',
                    'd_TUGUMI',
                    'd_VIOLA (ｳﾞｨｵﾗ)',
                    'd_YURURI',
                    'd_nae',
                    'd_tsubura',
                    'd_クレセント',
                    'd_侭 JIN',
                    'd_侭SUGI',
                    'd_円空',
                    'd_杜の詩',
                    'd_森のことば',
                    'd_森のことば ウォルナット',
                    'd_森のことばIBUKI',
                    'd_穂高',
                    'd_風のうた',
                    'd_ﾆｭｰﾏｯｷﾝﾚｲ',
                    'l_ALMO (ｱﾙﾓ)',
                    'l_AWASE',
                    'l_CHIGUSA(ﾁｸﾞｻ）',
                    'l_COBRINA',
                    'l_Kinoe',
                    'l_Northern Forest',
                    'l_PRESCELTO (ﾌﾟﾚｼｪﾙﾄ）',
                    'l_SEION 静穏',
                    'l_SEOTO',
                    'l_SEOTO-EX',
                    'l_TUGUMI',
                    'l_VIOLA (ｳﾞｨｵﾗ)',
                    'l_YURURI',
                    'l_nae',
                    'l_杜の詩',
                    'l_森のことば',
                    'l_森のことば ウォルナット',
                    'l_森のことばIBUKI',
                    'l_穂高',
                    'l_風のうた',
                    'l_ｽﾀﾝﾀﾞｰﾄﾞｺﾚｸｼｮﾝ',
                    'l_ﾆｭｰﾏｯｷﾝﾚｲ'
                    ]
    

    val_list = []
    for cate in selected_list:
        df_cust2 = df[df['series2']==cate]
        if len(df_cust2) == 0:
            val_list.append(0)
        else:
            val = df_cust2['金額'].sum()
            val_list.append(val)
    df_calc = pd.DataFrame(val_list, columns=[f'{target}_now'], index=selected_list)

    df_calc.loc['sales'] = df_calc.sum()

    return df_calc

#*******************************************************************グラフ
class Graph():
        def make_bar(self, val_list, x_list):
            #可視化
            #グラフを描くときの土台となるオブジェクト
            fig = go.Figure()
            #今期のグラフの追加
            for (val, x) in zip(val_list, x_list):
                fig.add_trace(
                    go.Bar(
                        x=[x],
                        y=[val],
                        text=[round(val/10000) if int(val) >= 10000 else int(val)],
                        textposition="outside", 
                        name=x)
                )
            #レイアウト設定     
            fig.update_layout(
                showlegend=False #凡例表示
            )
            #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
            st.plotly_chart(fig, use_container_width=True) 

        #**********************************************************棒グラフ　今期前期
        def make_bar_nowlast(self, lists_now, lists_last, x_list):
            #可視化
            #グラフを描くときの土台となるオブジェクト
            fig = go.Figure()
            #今期のグラフの追加
            
            for (val_list, name) in zip([lists_now, lists_last], ['今期', '前期']) :
                fig.add_trace(
                    go.Bar(
                        x=x_list,
                        y=val_list,  
                        text=[round(val/10000) if val >= 10000 else int(val) for val in val_list],
                        textposition="outside", 
                        name=name)
                )
            #レイアウト設定     
            fig.update_layout(
                showlegend=True #凡例表示
            )
            #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
            st.plotly_chart(fig, use_container_width=True) 

        #**********************************************************棒グラフ　今期前期 小数
        def make_bar_nowlast_float(self, lists_now, lists_last, x_list):
            #可視化
            #グラフを描くときの土台となるオブジェクト
            fig = go.Figure()
            #今期のグラフの追加
            
            for (val_list, name) in zip([lists_now, lists_last], ['今期', '前期']) :
                fig.add_trace(
                    go.Bar(
                        x=x_list,
                        y=val_list,  
                        text=[round(val, 2) for val in val_list],
                        textposition="outside", 
                        name=name)
                )
            #レイアウト設定     
            fig.update_layout(
                showlegend=True #凡例表示
            )
            #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
            st.plotly_chart(fig, use_container_width=True) 
        
        #*************************************************************棒グラフ　横 基準線あり
         #可視化
        def make_bar_h(self, val_list, label_list, name, title, line_val, height):
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=val_list,
                y=label_list,
                marker_color='#87cefa',
                textfont={'color': '#696969'},
                name=name)
                )
            fig.update_traces(
                textposition='outside',
                texttemplate='%{x:0.2f}',
                orientation='h'
                )
            # 基準線の追加
            fig.add_shape(
                type="line",
                x0=line_val,  # 基準線の開始位置 (x座標)
                x1=line_val,  # 基準線の終了位置 (x座標)
                y0=label_list[0],  # 基準線の開始位置 (y座標)
                y1=label_list[-1],  # 基準線の終了位置 (y座標)
                line=dict(
                    color="red",
                    width=2,
                    dash="dash"  # 破線を使用する場合は "dash" を指定
        )
    )
            fig.update_layout(
                title=title,
                width=500,
                height=height,
                plot_bgcolor='white'
                )
            #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
            st.plotly_chart(fig, use_container_width=True) 
        
        #*************************************************************棒グラフ　横　基準線なし
         #可視化
        def make_bar_h_nonline(self, val_list, label_list, name, title, height):
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=val_list,
                y=label_list,
                marker_color='#87cefa',
                textfont={'color': '#696969'},
                name=name)
                )
            fig.update_traces(
                textposition='outside',
                texttemplate='%{x}',
                orientation='h'
                )

            fig.update_layout(
                title=title,
                width=500,
                height=height,
                plot_bgcolor='white'
                )
            #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
            st.plotly_chart(fig, use_container_width=True) 

        #**********************************************************折れ線
        def make_line(self, df_list, name_list, x_list):

            #グラフを描くときの土台となるオブジェクト
            fig = go.Figure()
            #今期のグラフの追加

            for (df, name) in zip(df_list, name_list):

                fig.add_trace(
                go.Scatter(
                    x=x_list, #strにしないと順番が崩れる
                    y=df,
                    mode = 'lines+markers+text', #値表示
                    text=[round(val/10000) if val >= 10000 else int(val) for val in df],
                    textposition="top center", 
                    name=name)
                    )  

            #レイアウト設定     
            fig.update_layout(
                showlegend=True #凡例表示
            )
            #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
            st.plotly_chart(fig, use_container_width=True) 
        
        #**********************************************折れ線　non_xlist    
        def make_line_nonXlist(self, df_list, name_list):
            #グラフを描くときの土台となるオブジェクト
            fig = go.Figure()
            #今期のグラフの追加

            for (df, name) in zip(df_list, name_list):

                fig.add_trace(
                go.Scatter(
                    x=['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'], #strにしないと順番が崩れる
                    y=df,
                    mode = 'lines+markers+text', #値表示
                    text=[round(val/10000) if val >= 10000 else int(val) for val in df],
                    textposition="top center", 
                    name=name)
                    )  

            #レイアウト設定     
            fig.update_layout(
                showlegend=True #凡例表示
            )
            #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅
            st.plotly_chart(fig, use_container_width=True) 
            
        #***************************************************************円
        def make_pie(self, vals, labels):

            # st.write(f'{option_category} 構成比(今期)')
            fig = go.Figure(
                data=[
                    go.Pie(
                        labels=labels,
                        values=vals
                        )])
            fig.update_layout(
                showlegend=True, #凡例表示
                height=290,
                margin={'l': 20, 'r': 60, 't': 0, 'b': 0},
                )
            fig.update_traces(textposition='inside', textinfo='label+percent') 
            #inside グラフ上にテキスト表示
            st.plotly_chart(fig, use_container_width=True) 
            #plotly_chart plotlyを使ってグラグ描画　グラフの幅が列の幅

#*******************************************アイテムの深堀
#******************データの絞込み
def fukabori(df_now, df_now2, graph):
    st.markdown('#### 品番検索: 売れ筋分析 塗色/張地/得意先')

    #品番検索
    part_word = st.text_input(
        '頭品番 例SN',
        key='pw'
    )

    if part_word != '':

        item_list = []
        for item in df_now2['品番'].unique():
            if part_word in item:
                item_list.append(item)

        selected_item = st.selectbox(
            '品番',
            item_list,
            key='sl'
        )
        df_select = df_now2[df_now2['品番'] == selected_item]

        #******************塗色分析
        st.markdown('##### 塗色別数量')
        s_color = df_select.groupby('塗色CD')['数量'].sum()

        s_color = s_color.sort_values(ascending=False)
        graph.make_bar(s_color, s_color.index)

        #*****************張地ランキング
        st.markdown('##### 張地別数量')
        selected_color = st.selectbox(
            '塗色を選択',
            s_color.index,
            key='sc'
        )

        df_color = df_select[df_select['塗色CD']==selected_color]
        s_fab = df_color.groupby('張地')['数量'].sum()

        s_fab = s_fab.sort_values(ascending=True)

        with st.expander('張地グラフ', expanded=False):
            graph.make_bar_h_nonline(s_fab, s_fab.index, '数量', '張地/数量順', 800)

        #******************塗色張地分析
        s_item = df_select.groupby('商　品　名')['数量'].sum()

        s_item = s_item.sort_values(ascending=True)

        with st.expander('売れ筋組み合わせ 塗色／張地', expanded=False):
            graph.make_bar_h_nonline(s_item, s_item.index, '数量', '売れ筋組み合わせ 塗色／張地', 800)

            st.dataframe(s_item)
        #********************購入している得意先
        s_sum = df_select.groupby('得意先名')['数量'].sum()

        s_sum = s_sum.sort_values(ascending=True)
        graph.make_bar_h_nonline(s_sum, s_sum.index, '数量', '得意先別数量', 800)

        #dataframe作成
        sales_list = []
        rep_list = []
        for cust in s_sum.index:
            #得意先の売上取得
            df_cust = df_now[df_now['得意先名'] == cust]
            sales_sum = df_cust['金額'].sum()
            sales_list.append(sales_sum)
            #得意先担当者名取得
            rep = df_cust.iloc[0]['営業担当者名']
            rep_list.append(rep)

        df_cust = pd.DataFrame(list(zip(s_sum, sales_list, rep_list)), index=s_sum.index, \
                        columns=['数量', '参考:売上/全商品合計', '担当者'])
        df_cust.sort_values('数量', ascending=False, inplace=True)

        with st.expander('一覧', expanded=False):
            st.dataframe(df_cust)
        
        #********************得意先別深堀
        st.markdown('##### 得意先を選択して明細を見る')
        selected_cust = st.selectbox(
            '得意先を選択',
            df_cust.index,
            key='scust')
        
        df_cust2 = df_select[df_select['得意先名']==selected_cust]
        s_cust = df_cust2.groupby('商　品　名')['数量'].sum()

        s_cust = s_cust.sort_values(ascending=True)

        with st.expander('選択した得意先の明細を見る: 組み合わせ 塗色／張地', expanded=False):
            graph.make_bar_h_nonline(s_cust, s_cust.index, '数量', '組み合わせ 塗色／張地', 500)
            st.dataframe(df_cust2)

        
        #アソシエーション分析へのリンク
        st.markdown('#### 同時に買われているアイテムを見る')
        df_concat = pd.DataFrame()
        for num in df_cust2['伝票番号2']:
            df = df_now[df_now['伝票番号2'] == num]
            df_concat = pd.concat([df_concat, df], join='outer')
        
        with st.expander('明細', expanded=False):
            col_list = ['得意先名', '商　品　名', '数量', '伝票番号2']
            st.table(df_concat[col_list])


        link = '[アソシエーション分析](https://cocosan1-association-fullhinban-cmy4cf.streamlit.app/)'
        st.markdown(link, unsafe_allow_html=True)
