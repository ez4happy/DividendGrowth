import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import yfinance as yf
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
    df[roe_cols[0]]*0.3 +
    df[roe_cols[1]]*0.1 +
    df[roe_cols[2]]*0.6
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
# 와인스타인 4단계 순차 계산
# 2단계  30주 이평 위 + 20주 신고가
# 3단계  2단계였다가 MA150 아래 or 10주 신저가
# 4단계  30주 이평 아래 + 20주 신저가
# 1단계  4단계였다가 MA150 위 or 10주 신고가
# --------------------------------------------------
def calc_weinstein(ticker_code):
    try:
        ticker = str(int(float(ticker_code))).zfill(6) + ".KS"
        raw = yf.download(ticker, period="max", progress=False)
        if raw.empty:
            return "N/A"
        if isinstance(raw.columns, pd.MultiIndex):
            raw.columns = raw.columns.get_level_values(0)
        raw = raw.reset_index()
        raw['Date'] = pd.to_datetime(raw['Date']).dt.tz_localize(None)
        raw = raw[['Date','Close']].dropna()

        raw['MA150']  = raw['Close'].rolling(150).mean()
        raw['고가100'] = raw['Close'].rolling(100).max()
        raw['저가100'] = raw['Close'].rolling(100).min()
        raw['고가50']  = raw['Close'].rolling(50).max()
        raw['저가50']  = raw['Close'].rolling(50).min()

        stages = [None] * len(raw)
        for i in range(len(raw)):
            close   = raw['Close'].iloc[i]
            ma150   = raw['MA150'].iloc[i]
            high100 = raw['고가100'].iloc[i]
            low100  = raw['저가100'].iloc[i]
            high50  = raw['고가50'].iloc[i]
            low50   = raw['저가50'].iloc[i]

            if pd.isna(ma150) or pd.isna(high100) or pd.isna(low100):
                stages[i] = None
                continue

            prev = stages[i-1] if i > 0 else None

            if close > ma150 and close >= high100:
                stages[i] = "2단계"
            elif prev == "2단계" and (close < ma150 or close <= low50):
                stages[i] = "3단계"
            elif close < ma150 and close <= low100:
                stages[i] = "4단계"
            elif prev == "4단계" and (close > ma150 or close >= high50):
                stages[i] = "1단계"
            else:
                stages[i] = prev if prev else "1단계"

        return stages[-1] if stages[-1] else "N/A"

    except Exception:
        return "N/A"

# --------------------------------------------------
# 와인스타인 계산 (캐시)
# --------------------------------------------------
if '종목코드' in df.columns:
    with st.spinner("와인스타인 단계 계산 중..."):
        df['와인스타인'] = df['종목코드'].apply(calc_weinstein)
else:
    df['와인스타인'] = "N/A"

# --------------------------------------------------
# 정렬 및 순위
# --------------------------------------------------
df_sorted = df.sort_values(by='복리수익률', ascending=False).reset_index(drop=True)
df_sorted['순위'] = df_sorted.index + 1
df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

# --------------------------------------------------
# 표시 컬럼 (와인스타인 맨 끝에 추가)
# --------------------------------------------------
display_cols = [
    '순위', '종목명', '현재가', '등락률',
    '배당수익률', '추정ROE',
    'BPS', '10년후BPS',
    '복리수익률', 'RN',
    '와인스타인'
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

