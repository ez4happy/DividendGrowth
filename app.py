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

    # Stochastic 컬럼명 자동 탐색
    stochastic_col = None
    for col in df.columns:
        if 'stochastic' in col.lower():
            stochastic_col = col
            break
    if not stochastic_col:
        st.error("Stochastic 컬럼을 찾을 수 없습니다.")
        st.stop()

    # ROE 컬럼 자동 탐색
    roe_cols = [col for col in df.columns if 'ROE' in col and '평균' not in col and '최종' not in col]
    if len(roe_cols) < 3:
        st.error(f"ROE 컬럼이 3개 필요합니다. 현재: {roe_cols}")
        st.stop()
    roe_cols = roe_cols[:3]

    # 숫자형 컬럼 처리 (등락률은 이미 퍼센트 문자열이므로 제외)
    num_cols = ['현재가', 'BPS', '배당수익률', stochastic_col, '10년후BPS', '복리수익률'] + roe_cols + ['추정ROE']
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '').str.replace('%','').astype(float)

    # 추정ROE, 10년후BPS, 복리수익률 계산 (이미 있으면 생략)
    if '추정ROE' not in df.columns:
        df['추정ROE'] = df[roe_cols[0]]*0.4 + df[roe_cols[1]]*0.35 + df[roe_cols[2]]*0.25
    if '10년후BPS' not in df.columns and 'BPS' in df.columns and '추정ROE' in df.columns:
        df['10년후BPS'] = (df['BPS'] * (1 + df['추정ROE']/100) ** 10).round(0)
    if '복리수익률' not in df.columns and '10년후BPS' in df.columns and '현재가' in df.columns:
        df['복리수익률'] = (((df['10년후BPS'] / df['현재가']) ** (1/10)) - 1) * 100
        df['복리수익률'] = df['복리수익률'].round(2)

    # 가중치 슬라이더
    alpha = st.slider(
        '복리수익률(성장성) : 저평가(분위수) 가중치 (%)',
        min_value=0, max_value=100, value=80, step=5, format="%d%%"
    ) / 100

    # 복리수익률 점수 (15% 미만은 0점)
    max_return = df['복리수익률'].max()
    min_return = 15.0
    df['복리수익률점수'] = ((df['복리수익률'] - min_return) / (max_return - min_return)).clip(lower=0) * 100

    # 저평가 점수: 분위수(Percentile)
    df['Stochastic_percentile'] = df[stochastic_col].apply(lambda x: 100 - percentileofscore(df[stochastic_col], x, kind='mean'))

    # 매력도 계산
    df['매력도'] = (alpha * df['복리수익률점수'] + (1 - alpha) * df['Stochastic_percentile']).round(2)

    # 매력도 순 정렬 및 순위
    df_sorted = df.sort_values(by='매력도', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index + 1

    # 컬럼명 변경: 스톡캐스틱 %K 컬럼명을 'RN'으로 변경
    df_sorted = df_sorted.rename(columns={stochastic_col: 'RN'})

    # 표에 표시할 컬럼 순서 (등락률 포함, 스톡캐스틱 퍼센트 제외)
    main_cols = ['순위', '종목명', '현재가', '등락률'] + roe_cols + [
        'BPS', '배당수익률', 'RN', '추정ROE', '10년후BPS', '복리수익률', '매력도'
    ]
    final_cols = [col for col in main_cols if col in df_sorted.columns]
    df_show = df_sorted[final_cols]

    # 하이라이트 함수 (복리수익률 15% 이상 종목명 연두색)
    def highlight_high_return(row):
        color = 'background-color: lightgreen' if row['복리수익률'] >= 15 else ''
        return [color if col == '종목명' else '' for col in row.index]

    # 포맷 지정 (등락률은 이미 문자열이므로 제외)
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

    # 스타일 적용
    styled_df = (
        df_show.style
        .apply(highlight_high_return, axis=1)
        .format(format_dict)
        .set_properties(**{'text-align': 'center'})
        .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])
    )

    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

    # 복리수익률 vs RN 산점도
    fig_scatter = px.scatter(
        df_sorted,
        x='RN',
        y='복리수익률',
        color='매력도',
        hover_name='종목명',
        title='복리수익률 vs RN 산점도',
        labels={'RN': 'RN (Stochastic %K)', '복리수익률': '복리수익률 (%)'},
        color_continuous_scale='Viridis'
    )
    st.plotly_chart(fig_scatter, use_container_width=True)

    # 매력도 상위 종목 바 차트
    top_df = df_sorted.head(5)
    fig_bar = px.bar(
        top_df,
        x='종목명',
        y='매력도',
        title='매력도 상위 5개 종목',
        labels={'종목명': '종목명', '매력도': '매력도 점수'},
        color='매력도',
        color_continuous_scale='Viridis'
    )
    st.plotly_chart(fig_bar, use_container_width=True)

else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.\n\n해당 파일을 같은 폴더에 넣어주세요.")
