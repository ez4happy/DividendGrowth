import streamlit as st
import pandas as pd
import numpy as np
from scipy.stats import percentileofscore
import plotly.express as px
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # Stochastic 컬럼 탐색
    stochastic_col = next((col for col in df.columns if 'stochastic' in col.lower()), None)
    if not stochastic_col:
        st.error("Stochastic 컬럼을 찾을 수 없습니다.")
        st.stop()

    # ROE 컬럼 3개 탐색
    roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c]
    if len(roe_cols) < 3:
        st.error(f"ROE 컬럼이 3개 필요합니다. 현재: {roe_cols}")
        st.stop()
    roe_cols = roe_cols[:3]

    # 등락률 포맷
    if '등락률' in df.columns:
        df['등락률'] = df['등락률'].apply(lambda x: f"{float(x)*100:.2f}%" 
                                          if not isinstance(x, str) or '%' not in x else x)

    # 숫자형 변환
    num_cols = ['현재가', 'BPS', '배당수익률', stochastic_col, '10년후BPS', '복리수익률'] \
               + roe_cols + ['추정ROE']
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '').str.replace('%', '').astype(float)

    # 계산 컬럼
    if '추정ROE' not in df.columns:
        df['추정ROE'] = df[roe_cols[0]]*0.4 + df[roe_cols[1]]*0.35 + df[roe_cols[2]]*0.25
    if '10년후BPS' not in df.columns:
        df['10년후BPS'] = (df['BPS'] * (1 + df['추정ROE']/100) ** 10).round(0)
    if '복리수익률' not in df.columns:
        df['복리수익률'] = (((df['10년후BPS'] / df['현재가'])**(1/10)) - 1) * 100
        df['복리수익률'] = df['복리수익률'].round(2)

    # 매력도 계산
    alpha = st.slider('복리수익률(성장성) : 저평가(분위수) 가중치 (%)',
                      0, 100, 80, 5, format="%d%%") / 100
    max_return = df['복리수익률'].max()
    min_return = 15.0
    df['복리수익률점수'] = ((df['복리수익률'] - min_return) / 
                           (max_return - min_return)).clip(lower=0) * 100

    # 스토캐스틱 분위수 계산
    df['Stochastic_percentile'] = df[stochastic_col].apply(
        lambda x: 100 - percentileofscore(df[stochastic_col], x, kind='mean'))

    # 매력도 산출
    df['매력도'] = (alpha * df['복리수익률점수'] + 
                   (1 - alpha) * df['Stochastic_percentile']).round(2)

    # 정렬
    df_sorted = df.sort_values(by='매력도', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index + 1
    df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

    # 표시 컬럼 구성
    main_cols = ['순위', '종목명', '현재가', '등락률'] + roe_cols + \
                ['BPS', '배당수익률', 'RN', '추정ROE', '10년후BPS', '복리수익률', '매력도']
    df_show = df_sorted[[c for c in main_cols if c in df_sorted.columns]]

    # 스타일링
    def highlight_high_return(row):
        return ['background-color: lightgreen' if col == '종목명' and row['복리수익률'] >= 15 else ''
                for col in row.index]

    format_dict = {
        '현재가': '{:,.0f}', 
        roe_cols[0]: '{:.2f}', 
        roe_cols[1]: '{:.2f}', 
        roe_cols[2]: '{:.2f}',
        'BPS': '{:,.0f}', 
        '배당수익률': '{:.2f}', 
        'RN': '{:.0f}', 
        '추정ROE': '{:.2f}',
        '10년후BPS': '{:,.0f}', 
        '복리수익률': '{:.2f}', 
        '매력도': '{:.2f}'
    }

    styled_df = (df_show.style.apply(highlight_high_return, axis=1)
                 .format(format_dict)
                 .set_properties(**{'text-align': 'center'})
                 .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}]))

    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

    # 차트 1: 복리수익률 vs RN
    fig_scatter = px.scatter(df_sorted, x='RN', y='복리수익률', color='매력도',
                             hover_name='종목명',
                             title='복리수익률 vs RN 산점도',
                             labels={'RN': 'RN (Stochastic %K)', '복리수익률': '복리수익률 (%)'},
                             color_continuous_scale='Viridis')
    st.plotly_chart(fig_scatter, use_container_width=True)

    # 차트 2: 매력도 상위 5개 종목
    top_df = df_sorted.head(5)
    fig_bar = px.bar(top_df, x='종목명', y='매력도',
                     title='매력도 상위 5개 종목',
                     color='매력도', color_continuous_scale='Viridis')
    st.plotly_chart(fig_bar, use_container_width=True)

else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.")
