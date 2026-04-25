import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import yfinance as yf
from datetime import datetime, timedelta
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

date_placeholder = st.empty()

file_path = "1.xlsx"

# --------------------------------------------------
# CSS
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

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()
if df.empty:
    st.error("엑셀 파일이 비어 있습니다.")
    st.stop()

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

required_base_cols = ['종목명', '종목코드', 'BPS']
for col in required_base_cols:
    if col not in df.columns:
        st.error(f"필수 컬럼 누락: {col}")
        st.stop()

stochastic_col = next((c for c in df.columns if 'stochastic' in c.lower()), None)
if not stochastic_col:
    st.error("Stochastic 컬럼을 찾을 수 없습니다.")
    st.stop()

roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c]
if len(roe_cols) < 3:
    st.error("ROE 컬럼이 3개 이상 필요합니다.")
    st.stop()
roe_cols = roe_cols[:3]

if '배당수익률' in df.columns:
    df['배당수익률'] = normalize_percent(df['배당수익률'])

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
# yfinance에서 데이터 가져오기
# 핵심 변경: auto_adjust=False → 원본 종가(KRX 표시값과 동일) 사용
# --------------------------------------------------
def _download_yf(code_with_suffix):
    """yfinance 다운로드 헬퍼. auto_adjust=False로 원본 OHLC 받음"""
    today     = datetime.today()
    from_date = (today - timedelta(days=365*3)).strftime("%Y-%m-%d")
    to_date   = (today + timedelta(days=1)).strftime("%Y-%m-%d")
    raw = yf.download(
        code_with_suffix,
        start=from_date,
        end=to_date,
        progress=False,
        auto_adjust=False,   # ← 중요: 수정종가 아닌 원본 종가
        actions=False,
    )
    if isinstance(raw.columns, pd.MultiIndex):
        raw.columns = raw.columns.get_level_values(0)
    return raw

@st.cache_data(ttl=3600)
def get_stock_data(ticker_code):
    """(현재가, 등락률%, 와인스타인 단계, 마지막 거래일) 반환"""
    try:
        code = str(int(float(ticker_code))).zfill(6)

        # 코스피(.KS) 먼저, 없으면 코스닥(.KQ)
        raw = _download_yf(code + ".KS")
        if raw is None or raw.empty:
            raw = _download_yf(code + ".KQ")
        if raw is None or raw.empty:
            return np.nan, np.nan, "N/A", pd.NaT

        raw = raw.reset_index()
        raw['Date'] = pd.to_datetime(raw['Date']).dt.tz_localize(None)
        raw = raw[['Date', 'Close', 'High', 'Low', 'Volume']].dropna(
            subset=['Close', 'High', 'Low']
        )
        raw = raw.sort_values('Date').reset_index(drop=True)

        if len(raw) < 2:
            return np.nan, np.nan, "N/A", pd.NaT

        current_price = float(raw['Close'].iloc[-1])
        prev_price    = float(raw['Close'].iloc[-2])
        change_pct    = ((current_price - prev_price) / prev_price * 100) if prev_price > 0 else np.nan
        last_date     = pd.to_datetime(raw['Date'].iloc[-1])

        stage = calc_weinstein_stages_from_df(raw) if len(raw) >= 150 else "N/A"
        return current_price, change_pct, stage, last_date
    except Exception:
        return np.nan, np.nan, "N/A", pd.NaT


# --------------------------------------------------
# 일괄 수집
# --------------------------------------------------
with st.spinner("Yahoo Finance에서 주가 데이터 수집 중..."):
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

latest_date = df['기준일'].dropna().max()

df.dropna(subset=['현재가'], inplace=True)
if df.empty:
    st.warning("주가 데이터를 가져온 종목이 없습니다.")
    st.stop()

if pd.notna(latest_date):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    date_placeholder.caption(
        f"📅 **데이터 기준일:** {latest_date.strftime('%Y-%m-%d')} "
        f"(Yahoo Finance 종가 기준) &nbsp;&nbsp;|&nbsp;&nbsp; "
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

df_sorted = df.sort_values(by='복리수익률', ascending=False).reset_index(drop=True)
df_sorted['순위'] = df_sorted.index + 1
df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

display_cols = [
    '순위', '종목명', '현재가', '등락률',
    '배당수익률', '추정ROE',
    'BPS', '10년후BPS',
    '복리수익률', 'RN',
    '와인스타인'
]
existing_cols = [c for c in display_cols if c in df_sorted.columns]
df_show = df_sorted[existing_cols]

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
