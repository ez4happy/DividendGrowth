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

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()

# ----------------------------
# 공통 숫자 변환 함수
# ----------------------------
def to_numeric_safe(series):
    return (
        series.astype(str)
        .str.replace(',', '', regex=False)
        .str.replace('%', '', regex=False)
        .replace('', np.nan)
        .astype(float)
    )

# ----------------------------
# 퍼센트 자동 판별 함수
# ----------------------------
def normalize_percent(series):
    if series.dtype == object:
        series = to_numeric_safe(series)
    else:
        series = pd.to_numeric(series, errors='coerce')

    # 절대값 1보다 작은 경우 엑셀 퍼센트 형식으로 간주
    series = np.where(series.abs() < 1, series * 100, series)
    return pd.Series(series).round(2)

# ----------------------------
# 필수 컬럼 체크
# ----------------------------
required_cols = ['종목명', '현재가', 'BPS']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"필수 컬럼 누락: {missing}")
    st.stop()

# ----------------------------
# Stochastic %K 자동 탐색
# ----------------------------
stochastic_col = next(
    (c for c in df.columns if 'stochastic' in c.lower() and '%k' in c.lower()),
    None
)

if not stochastic_col:
    st.error("Stochastic Fast %K 컬럼을 찾을 수 없습니다.")
    st.stop()

# ----------------------------
# ROE 컬럼 3개 탐색
# ----------------------------
roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c]

if len(roe_cols) < 3:
    st.error(f"ROE 컬럼 3개 필요. 현재: {roe_cols}")
    st.stop()

roe_cols = roe_cols[:3]

# ----------------------------
# 숫자 변환
# ----------------------------
numeric_basic = ['현재가', 'BPS']
for col in numeric_basic:
    df[col] = to_numeric_safe(df[col])

if '배당수익률' in df.columns:
    df['배당수익률'] = normalize_percent(df['배당수익률'])

if '등락률' in df.columns:
    df['등락률'] = normalize_percent(df['등락률'])

df[stochastic_col] = to_numeric_safe(df[stochastic_col])

for col in roe_cols:
    df[col] = to_numeric_safe(df[col])

# ----------------------------
# 계산
# ----------------------------
df['추정ROE'] = (
    df[roe_cols[0]] * 0.4 +
    df[roe_cols[1]] * 0.35 +
    df[roe_cols[2]] * 0.25
)

df['10년후BPS'] = np.where(
    df['BPS'] > 0,
    df['BPS'] * (1 + df['추정ROE'] / 100) ** 10,
    np.nan
)

df['복리수익률'] = np.where(
    (df['현재가'] > 0) & (df['10년후BPS'] > 0),
    ((df['10년후BPS'] / df['현재가']) ** (1/10) - 1) * 100,
    np.nan
)

df['복리수익률'] = df['복리수익률'].round(2)
df['10년후BPS'] = df['10년후BPS'].round(0)

# ----------------------------
# 정렬 및 순위
# ----------------------------
df_sorted = df.sort_values(by='복리수익률', ascending=False).reset_index(drop=True)
df_sorted['순위'] = df_sorted.index + 1
df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

# ----------------------------
# 표시 컬럼 순서 (고정)
# ----------------------------
display_cols = [
    '순위','종목명','현재가','등락률',
    '배당수익률','추정ROE','BPS',
    '10년후BPS','복리수익률','RN'
]

df_show = df_sorted[display_cols]

# ----------------------------
# 스타일링
# ----------------------------
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
        .format(format_dict)
        .set_properties(**{'text-align': 'center'})
        .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])
)

st.dataframe(styled_df, use_container_width=True, height=None, hide_index=True)

# ----------------------------
# 산점도 (15% 이상 색상 구분)
# ----------------------------
df_sorted['HighReturn'] = df_sorted['복리수익률'] >= 15

fig = px.scatter(
    df_sorted,
    x='RN',
    y='복리수익률',
    color='HighReturn',
    hover_name='종목명',
    title='복리수익률 vs RN 산점도',
    labels={'RN':'RN (Stochastic %K)', '복리수익률':'복리수익률 (%)'}
)

st.plotly_chart(fig, use_container_width=True)
