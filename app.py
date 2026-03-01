import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

# --------------------------------------------------
# CSS 강제 가운데 정렬 (st.dataframe 대응)
# --------------------------------------------------
st.markdown(
    """
    <style>
    div[data-testid="stDataFrame"] th {
        text-align: center !important;
    }
    div[data-testid="stDataFrame"] td {
        text-align: center !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

if not os.path.exists(file_path):
    st.error(f"'{file_path}' 파일이 존재하지 않습니다.")
    st.stop()

# --------------------------------------------------
# 데이터 로드
# --------------------------------------------------
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()

if df.empty:
    st.error("엑셀 파일이 비어 있습니다.")
    st.stop()

# --------------------------------------------------
# 안전 숫자 변환
# --------------------------------------------------
def to_numeric_safe(series):
    return pd.to_numeric(
        series.astype(str)
              .str.replace(',', '', regex=False)
              .str.replace('%', '', regex=False)
              .replace('', np.nan),
        errors='coerce'
    )

def normalize_percent(series):
    series = to_numeric_safe(series)
    series = np.where(series.abs() <= 1, series * 100, series)
    return pd.Series(series).round(2)

# --------------------------------------------------
# 필수 컬럼 체크
# --------------------------------------------------
required_base_cols = ['종목명', '현재가', 'BPS']
for col in required_base_cols:
    if col not in df.columns:
        st.error(f"필수 컬럼 누락: {col}")
        st.stop()

# --------------------------------------------------
# Stochastic 컬럼 자동 탐색
# --------------------------------------------------
stochastic_col = next((c for c in df.columns if 'stochastic' in c.lower()), None)
if not stochastic_col:
    st.error("Stochastic 컬럼을 찾을 수 없습니다.")
    st.stop()

# --------------------------------------------------
# ROE 컬럼 자동 탐색 (3개 이상 중 상위 3개)
# --------------------------------------------------
roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c]
if len(roe_cols) < 3:
    st.error("ROE 컬럼이 3개 이상 필요합니다.")
    st.stop()

roe_cols = roe_cols[:3]

# --------------------------------------------------
# 퍼센트 처리
# --------------------------------------------------
if '등락률' in df.columns:
    df['등락률'] = normalize_percent(df['등락률'])

if '배당수익률' in df.columns:
    df['배당수익률'] = normalize_percent(df['배당수익률'])

# --------------------------------------------------
# 숫자 변환
# --------------------------------------------------
num_cols = ['현재가', 'BPS', stochastic_col] + roe_cols
for col in num_cols:
    df[col] = to_numeric_safe(df[col])

df.dropna(subset=['현재가', 'BPS'], inplace=True)

# --------------------------------------------------
# 계산
# --------------------------------------------------
df['추정ROE'] = (
    df[roe_cols[0]]*0.4 +
    df[roe_cols[1]]*0.35 +
    df[roe_cols[2]]*0.25
)

df['추정ROE'] = df['추정ROE'].fillna(0)

df['10년후BPS'] = df['BPS'] * (1 + df['추정ROE']/100) ** 10
df['10년후BPS'] = df['10년후BPS'].replace([np.inf, -np.inf], np.nan)

df['복리수익률'] = np.where(
    df['현재가'] > 0,
    ((df['10년후BPS'] / df['현재가']) ** (1/10) - 1) * 100,
    np.nan
)

df['복리수익률'] = df['복리수익률'].replace([np.inf, -np.inf], np.nan)
df['복리수익률'] = df['복리수익률'].round(2)

df.dropna(subset=['복리수익률'], inplace=True)

if df.empty:
    st.warning("계산 후 표시할 데이터가 없습니다.")
    st.stop()

# --------------------------------------------------
# 정렬 및 순위
# --------------------------------------------------
df_sorted = df.sort_values(by='복리수익률', ascending=False).reset_index(drop=True)
df_sorted['순위'] = df_sorted.index + 1
df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

# --------------------------------------------------
# 표시 컬럼
# --------------------------------------------------
display_cols = [
    '순위', '종목명', '현재가', '등락률',
    '배당수익률', '추정ROE',
    'BPS', '10년후BPS',
    '복리수익률', 'RN'
]

existing_cols = [c for c in display_cols if c in df_sorted.columns]
df_show = df_sorted[existing_cols]

# --------------------------------------------------
# 하이라이트
# --------------------------------------------------
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
)

# height 자동 계산 (스크롤 최소화)
row_height = 35
calculated_height = min(len(df_show) * row_height + 60, 1000)

st.dataframe(
    styled_df,
    use_container_width=True,
    height=calculated_height,
    hide_index=True
)

# --------------------------------------------------
# 산점도
# --------------------------------------------------
df_sorted['HighReturn'] = df_sorted['복리수익률'] >= 15

fig = px.scatter(
    df_sorted,
    x='RN',
    y='복리수익률',
    color='HighReturn',
    hover_name='종목명',
    title='복리수익률 vs RN',
    labels={'RN': 'RN(Stochastic %K)', '복리수익률': '복리수익률(%)'}
)

st.plotly_chart(fig, use_container_width=True)
