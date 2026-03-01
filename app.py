import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # Stochastic 컬럼 찾기
    stochastic_col = next((col for col in df.columns if 'stochastic' in col.lower()), None)
    if not stochastic_col:
        st.error("Stochastic 컬럼을 찾을 수 없습니다.")
        st.stop()

    # ROE 컬럼 3개 찾기
    roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c]
    if len(roe_cols) < 3:
        st.error(f"ROE 컬럼 3개가 필요합니다. 현재: {roe_cols}")
        st.stop()
    roe_cols = roe_cols[:3]

    # -----------------------------
    # 안전 숫자 변환
    # -----------------------------
    def to_numeric_safe(series):
        return (
            series.astype(str)
            .str.replace(',', '', regex=False)
            .str.replace('%', '', regex=False)
            .replace('', np.nan)
            .astype(float)
        )

    def normalize_percent(series):
        series = to_numeric_safe(series)
        series = np.where(series.abs() < 1, series * 100, series)
        return pd.Series(series).round(2)

    # 퍼센트 처리
    if '등락률' in df.columns:
        df['등락률'] = normalize_percent(df['등락률'])

    if '배당수익률' in df.columns:
        df['배당수익률'] = normalize_percent(df['배당수익률'])

    # 숫자 변환
    num_cols = ['현재가', 'BPS', stochastic_col] + roe_cols
    for col in num_cols:
        df[col] = to_numeric_safe(df[col])

    # 계산
    df['추정ROE'] = (
        df[roe_cols[0]]*0.4 +
        df[roe_cols[1]]*0.35 +
        df[roe_cols[2]]*0.25
    )

    df['10년후BPS'] = (
        df['BPS'] * (1 + df['추정ROE']/100) ** 10
    ).round(0)

    df['복리수익률'] = (
        (df['10년후BPS'] / df['현재가']) ** (1/10) - 1
    ) * 100
    df['복리수익률'] = df['복리수익률'].round(2)

    # 🔥 복리수익률 기준 정렬
    df_sorted = df.sort_values(by='복리수익률', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index + 1
    df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

    # 표시 컬럼 순서
    display_cols = [
        '순위', '종목명', '현재가', '등락률',
        '배당수익률', '추정ROE',
        'BPS', '10년후BPS',
        '복리수익률', 'RN'
    ]

    df_show = df_sorted[display_cols]

    # 하이라이트
    def highlight_high_return(row):
        return [
            'background-color: lightgreen'
            if row['복리수익률'] >= 15 else ''
            for _ in row
        ]

    format_dict = {
        '현재가': '{:,.0f}',
        '등락률': '{:.2f}%',
        '배당수익률': '{:.2f}%',
        '추정ROE': '{:.2f}',
        'BPS': '{:,.0f}',
        '10년후BPS': '{:,.0f}',
        '복리수익률': '{:.2f}%',
        'RN': '{:.0f}'
    }

    styled_df = (
        df_show.style
              .apply(highlight_high_return, axis=1)
              .format(format_dict)
              .set_properties(**{'text-align': 'center'})
              .set_table_styles(
                  [{'selector':'th','props':[('text-align','center')]}]
              )
    )

    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

    # 산점도 (15% 이상 색 구분)
    df_sorted['HighReturn'] = df_sorted['복리수익률'] >= 15

    fig_scatter = px.scatter(
        df_sorted,
        x='RN',
        y='복리수익률',
        color='HighReturn',
        hover_name='종목명',
        title='복리수익률 vs RN 산점도',
        labels={'RN':'RN(Stochastic %K)', '복리수익률':'복리수익률(%)'}
    )
    st.plotly_chart(fig_scatter, use_container_width=True)

else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.")
