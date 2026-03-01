import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

if not os.path.exists(file_path):
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.")
    st.stop()

# =========================
# 데이터 로드
# =========================
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()

required_cols = ['종목명', '현재가', 'BPS']
missing_cols = [c for c in required_cols if c not in df.columns]
if missing_cols:
    st.error(f"필수 컬럼이 없습니다: {missing_cols}")
    st.stop()

# =========================
# Stochastic 컬럼 탐색
# =========================
stochastic_col = next(
    (col for col in df.columns if 'stochastic' in col.lower() and '%k' in col.lower()),
    None
)

if not stochastic_col:
    st.error("Stochastic Fast %K 컬럼을 찾을 수 없습니다.")
    st.stop()

# =========================
# ROE 컬럼 3개 탐색
# =========================
roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c]
if len(roe_cols) < 3:
    st.error(f"ROE 컬럼 3개가 필요합니다. 현재: {roe_cols}")
    st.stop()

roe_cols = roe_cols[:3]

# =========================
# 숫자형 안전 변환 함수
# =========================
def to_float(series):
    return (
        series.astype(str)
              .str.replace(',', '', regex=False)
              .str.replace('%', '', regex=False)
              .replace('', np.nan)
              .astype(float)
    )

numeric_cols = ['현재가', 'BPS', '배당수익률', stochastic_col] + roe_cols

for col in numeric_cols:
    if col in df.columns:
        df[col] = to_float(df[col])

# =========================
# 계산 컬럼 생성
# =========================

# 추정 ROE
df['추정ROE'] = (
    df[roe_cols[0]] * 0.4 +
    df[roe_cols[1]] * 0.35 +
    df[roe_cols[2]] * 0.25
)

# 10년 후 BPS
df['10년후BPS'] = (
    df['BPS'] * (1 + df['추정ROE'] / 100) ** 10
)

# 복리수익률 (0 나누기 방지)
df['복리수익률'] = np.where(
    df['현재가'] > 0,
    ((df['10년후BPS'] / df['현재가']) ** (1/10) - 1) * 100,
    np.nan
)

df['복리수익률'] = df['복리수익률'].round(2)
df['10년후BPS'] = df['10년후BPS'].round(0)

# =========================
# 정렬 (복리수익률 기준)
# =========================
df_sorted = (
    df.sort_values(by='복리수익률', ascending=False)
      .reset_index(drop=True)
)

df_sorted['순위'] = df_sorted.index + 1
df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

# =========================
# 표시 컬럼 구성
# =========================
main_cols = (
    ['순위', '종목명', '현재가', '등락률'] +
    roe_cols +
    ['BPS', '배당수익률', 'RN', '추정ROE',
     '10년후BPS', '복리수익률']
)

df_show = df_sorted[[c for c in main_cols if c in df_sorted.columns]]

# =========================
# 스타일링
# =========================
def highlight_high_return(row):
    return [
        'background-color: lightgreen'
        if col == '종목명' and row['복리수익률'] >= 15
        else ''
        for col in row.index
    ]

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
    '복리수익률': '{:.2f}'
}

styled_df = (
    df_show.style
        .apply(highlight_high_return, axis=1)
        .format(format_dict)
        .set_properties(**{'text-align': 'center'})
        .set_table_styles(
            [{'selector': 'th', 'props': [('text-align', 'center')]}]
        )
)

st.dataframe(styled_df, use_container_width=True, height=550, hide_index=True)

# =========================
# 산점도 차트
# =========================
fig_scatter = px.scatter(
    df_sorted,
    x='RN',
    y='복리수익률',
    hover_name='종목명',
    title='복리수익률 vs RN 산점도',
    labels={'RN': 'RN (Stochastic %K)', '복리수익률': '복리수익률 (%)'}
)

st.plotly_chart(fig_scatter, use_container_width=True)
