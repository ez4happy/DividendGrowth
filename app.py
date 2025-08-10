import streamlit as st
import pandas as pd
import numpy as np
from scipy.stats import percentileofscore
import plotly.express as px
import os
import requests
from bs4 import BeautifulSoup

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

def get_pbr_band(gicode):
    """
    FnGuide 페이지에서 PBR 밴드 값 5개를 가져오는 함수
    gicode: 'A005930' 형식의 종목코드
    """
    url = f"https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode={gicode}&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701"
    resp = requests.get(url)
    resp.encoding = 'utf-8'
    soup = BeautifulSoup(resp.text, 'html.parser')

    # TODO: 실제 FnGuide HTML 구조에서 PBR 밴드 값 가져오기
    # 아래는 예시값 -> 실제 웹 구조 분석 후 변경 필요
    band_list = [0.9, 1.1, 1.3, 1.5, 1.7]
    return band_list

def get_pbr(price, bps):
    try:
        return price / bps if bps != 0 else np.nan
    except:
        return np.nan

def get_position(pbr, bands):
    if pd.isna(pbr):
        return np.nan
    if pbr < bands[0]:
        return 1
    elif pbr < bands[1]:
        return 2
    elif pbr < bands[2]:
        return 3
    elif pbr < bands[3]:
        return 4
    elif pbr < bands[4]:
        return 5
    else:
        return 6

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

    # 등락률을 퍼센트 문자열로 변환
    def percent_format(x):
        try:
            if isinstance(x, str) and '%' in x:
                return x
            else:
                return f"{float(x)*100:.2f}%"
        except:
            return x

    if '등락률' in df.columns:
        df['등락률'] = df['등락률'].apply(percent_format)

    # 숫자형 변환
    num_cols = ['현재가', 'BPS', '배당수익률', stochastic_col, '10년후BPS', '복리수익률'] + roe_cols + ['추정ROE']
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '').str.replace('%','').astype(float)

    # 계산 컬럼
    if '추정ROE' not in df.columns:
        df['추정ROE'] = df[roe_cols[0]]*0.4 + df[roe_cols[1]]*0.35 + df[roe_cols[2]]*0.25
    if '10년후BPS' not in df.columns and 'BPS' in df.columns and '추정ROE' in df.columns:
        df['10년후BPS'] = (df['BPS'] * (1 + df['추정ROE']/100) ** 10).round(0)
    if '복리수익률' not in df.columns and '10년후BPS' in df.columns and '현재가' in df.columns:
        df['복리수익률'] = (((df['10년후BPS'] / df['현재가']) ** (1/10)) - 1) * 100
        df['복리수익률'] = df['복리수익률'].round(2)

    # position 컬럼 생성
    if '종목코드' in df.columns:
        for idx, row in df.iterrows():
            gicode = row['종목코드']
            price = row['현재가']
            bps = row['BPS']
            bands = get_pbr_band(gicode)
            pbr = get_pbr(price, bps)
            df.at[idx, 'position'] = get_position(pbr, bands)

    # 성장성 vs 저평가 가중치
    alpha = st.slider(
        '복리수익률(성장성) : 저평가(분위수) 가중치 (%)',
        min_value=0, max_value=100, value=80, step=5, format="%d%%"
    ) / 100

    max_return = df['복리수익률'].max()
    min_return = 15.0
    df['복리수익률점수'] = ((df['복리수익률'] - min_return) / (max_return - min_return)).clip(lower=0) * 100

    df['Stochastic_percentile'] = df[stochastic_col].apply(
        lambda x: 100 - percentileofscore(df[stochastic_col], x, kind='mean')
    )

    df['매력도'] = (alpha * df['복리수익률점수'] + (1 - alpha) * df['Stochastic_percentile']).round(2)

    df_sorted = df.sort_values(by='매력도', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index + 1
    df_sorted = df_sorted.rename(columns={stochastic_col: 'RN'})

    main_cols = ['순위', '종목명', '현재가', '등락률'] + roe_cols + [
        'BPS', '배당수익률', 'RN', '추정ROE', '10년후BPS', '복리수익률', 'position', '매력도'
    ]
    final_cols = [col for col in main_cols if col in df_sorted.columns]
    df_show = df_sorted[final_cols]

    # 하이라이트: 복리수익률 15% 이상
    def highlight_high_return(row):
        color = 'background-color: lightgreen' if row['복리수익률'] >= 15 else ''
        return [color if col == '종목명' else '' for col in row.index]

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
        '매력도': '{:.2f}',
        'position': '{:.0f}'
    }

    styled_df = (
        df_show.style
        .apply(highlight_high_return, axis=1)
        .format(format_dict)
        .set_properties(**{'text-align': 'center'})
        .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])
    )

    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

    # 시각화
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
