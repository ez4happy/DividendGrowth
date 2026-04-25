import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from pykrx import stock as krx
from datetime import datetime, timedelta
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

# 데이터 기준일 표시용 placeholder (수집 후 업데이트)
date_placeholder = st.empty()

file_path = "1.xlsx"

# --------------------------------------------------
# CSS 강제 가운데 정렬
# --------------------------------------------------
st.markdown(
    """
    <style>
    div[data-testid="stDataFrame"] th { text-align: center !important; }
    div[data-testid="stDataFrame"] td { text-align: center !important; }
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
# 필수 컬럼 체크 (현재가는 API에서 가져옴, 종목코드 필수)
# --------------------------------------------------
required_base_cols = ['종목명', '종목코드', 'BPS']
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
# ROE 컬럼 자동 탐색
# --------------------------------------------------
roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c]
if len(roe_cols) < 3:
    st.error("ROE 컬럼이 3개 이상 필요합니다.")
    st.stop()
roe_cols = roe_cols[:3]

# --------------------------------------------------
# 퍼센트 처리 (배당수익률만 엑셀에서 읽음)
# --------------------------------------------------
if '배당수익률' in df.columns:
    df['배당수익률'] = normalize_percent(df['배당수익률'])

# --------------------------------------------------
# 숫자 변환 (현재가/등락률은 pykrx에서 받음)
# --------------------------------------------------
num_cols = ['BPS', stochastic_col] + roe_cols
for col in num_cols:
    df[col] = to_numeric_safe(df[col])
df.dropna(subset=['BPS'], inplace=True)

# --------------------------------------------------
# 와인스타인 4단계 계산
# --------------------------------------------------
def calc_weinstein_stages_from_df(raw):
    raw = raw.copy()
    raw['MA150']    = raw['Close'].rolling(150).mean()
    raw['고가99']    = raw['High'].shift(1).rolling(99).max()
    raw['저가99']    = raw['Low'].shift(1).rolling(99).min()
    raw['고가49']    = raw['High'].shift(1).rolling(49).max()
    raw['저가49']    = raw['Low'].shift(1).rolling(49).min()
    raw['신고가100'] = (raw['High'] > raw['고가99']).fillna(False)
    raw['신저가100'] = (raw['Low']  < raw['저가99']).fillna(False)
    raw['신고가50']  = (raw['High'] > raw['고가49']).fillna(False)
    raw['신저가50']  = (raw['Low']  < raw['저가49']).fillna(False)

    close_arr = raw['Close'].values
    ma150_arr = raw['MA150'].values
    nh100_arr = raw['신고가100'].values
    nl100_arr = raw['신저가100'].values
    nh50_arr  = raw['신고가50'].values
    nl50_arr  = raw['신저가50'].values

    n      = len(raw)
    stages = [None] * n

    for i in range(n):
        if np.isnan(ma150_arr[i]):
            stages[i] = None
            continue

        close = close_arr[i]
        ma150 = ma150_arr[i]
        nh100 = nh100_arr[i]
        nl100 = nl100_arr[i]
        nh50  = nh50_arr[i]
        nl50  = nl50_arr[i]
        prev  = stages[i-1] if i > 0 else None

        if close > ma150 and nh100:
            stages[i] = "2단계"
        elif close < ma150 and nl100:
            stages[i] = "4단계"
        elif prev == "2단계" and (close < ma150 or nl50):
            stages[i] = "3단계"
        elif prev == "4단계" and (close > ma150 or nh50):
            stages[i] = "1단계"
        else:
            stages[i] = prev if prev else "1단계"

    return stages[-1] if stages[-1] else "N/A"


# --------------------------------------------------
# pykrx에서 현재가 + 등락률 + 와인스타인 + 마지막 거래일
# --------------------------------------------------
@st.cache_data(ttl=3600)
def get_stock_data(ticker_code):
    """pykrx 일봉 데이터로 (현재가, 등락률%, 와인스타인 단계, 마지막 거래일) 반환"""
    try:
        code = str(int(float(ticker_code))).zfill(6)
        today     = datetime.today()
        from_date = (today - timedelta(days=365*3)).strftime("%Y%m%d")
        to_date   = today.strftime("%Y%m%d")

        raw = krx.get_market_ohlcv(from_date, to_date, code)
        if raw is None or raw.empty:
            return np.nan, np.nan, "N/A", pd.NaT

        raw = raw.reset_index().rename(columns={
            '날짜':   'Date',
            '시가':   'Open',
            '고가':   'High',
            '저가':   'Low',
            '종가':   'Close',
            '거래량': 'Volume',
            '등락률': 'ChangePct'
        })
        raw = raw[['Date', 'Close', 'High', 'Low', 'Volume', 'ChangePct']].dropna(
            subset=['Close', 'High', 'Low']
        )
        raw = raw.sort_values('Date').reset_index(drop=True)

        if len(raw) < 2:
            return np.nan, np.nan, "N/A", pd.NaT

        current_price = float(raw['Close'].iloc[-1])
        change_pct    = float(raw['ChangePct'].iloc[-1])
        last_date     = pd.to_datetime(raw['Date'].iloc[-1])

        stage = calc_weinstein_stages_from_df(raw) if len(raw) >= 150 else "N/A"
        return current_price, change_pct, stage, last_date
    except Exception:
        return np.nan, np.nan, "N/A", pd.NaT


# --------------------------------------------------
# 종목별 주가 데이터 일괄 수집
# --------------------------------------------------
with st.spinner("KRX에서 주가 데이터 수집 중... (종목 수에 따라 시간이 걸려요)"):
    progress = st.progress(0)
    total    = len(df)
    rows     = []
    for i, code in enumerate(df['종목코드'].tolist()):
        rows.append(get_stock_data(code))
        progress.progress((i + 1) / total)
    progress.empty()

results = pd.DataFrame(
    rows,
    columns=['현재가', '등락률', '와인스타인', '기준일'],
    index=df.index
)
df['현재가']     = results['현재가']
df['등락률']     = results['등락률'].round(2)
df['와인스타인'] = results['와인스타인']
df['기준일']     = results['기준일']

# 가장 최근 거래일 (= 데이터 기준일)
latest_date = df['기준일'].dropna().max()

# API 실패 종목 제거
df.dropna(subset=['현재가'], inplace=True)
if df.empty:
    st.warning("주가 데이터를 가져온 종목이 없습니다. 종목코드를 확인하세요.")
    st.stop()

# --------------------------------------------------
# 데이터 기준일 상단 표시
# --------------------------------------------------
if pd.notna(latest_date):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    date_placeholder.caption(
        f"📅 **데이터 기준일:** {latest_date.strftime('%Y-%m-%d')} "
        f"(KRX 종가 기준) &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"🔄 **조회 시각:** {now_str}"
    )

# --------------------------------------------------
# 계산
# --------------------------------------------------
df['추정ROE'] = (
    df[roe_cols[0]]*0.3 +
    df[roe_cols[1]]*0.1 +
    df[roe_cols[2]]*0.6
).fillna(0)
df['10년후BPS'] = df['BPS'] * (1 + df['추정ROE']/100) ** 10
df['10년후BPS'] = df['10년후BPS'].replace([np.inf, -np.inf], np.nan)
df['복리수익률'] = np.where(
    df['현재가'] > 0,
    ((df['10년후BPS'] / df['현재가']) ** (1/10) - 1) * 100,
    np.nan
)
df['복리수익률'] = df['복리수익률'].replace([np.inf, -np.inf], np.nan).round(2)
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
    '현재가':    '{:,.0f}',
    '등락률':    '{:.2f}%',
    '배당수익률': '{:.2f}%',
    '추정ROE':  '{:.2f}',
    'BPS':      '{:,.0f}',
    '10년후BPS': '{:,.0f}',
    '복리수익률': '{:.2f}%',
    'RN':       '{:.0f}'
}

styled_df = (
    df_show.style
          .apply(highlight_high_return, axis=1)
          .format(format_dict)
)

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
