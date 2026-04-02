import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import requests
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

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
# ROE 컬럼 자동 탐색
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
# 와인스타인 4단계 (메인 코드와 동일 로직)
#
# 2단계  종가 > MA150 AND 100일 신고가 (고가 기준)
# 3단계  2단계였다가 AND (종가 < MA150 OR 50일 신저가 (저가 기준))
# 4단계  종가 < MA150 AND 100일 신저가 (저가 기준)
# 1단계  4단계였다가 AND (종가 > MA150 OR 50일 신고가 (고가 기준))
# 그 외  이전 단계 유지
#
# 신고가/신저가 컬럼 명시적 생성 (True/False)
# 고가99 = shift(1).rolling(99) -> 오늘 포함 100일
# 저가99 = shift(1).rolling(99)
# 고가49 = shift(1).rolling(49) -> 오늘 포함 50일
# 저가49 = shift(1).rolling(49)
# --------------------------------------------------
def calc_weinstein_stages_from_df(raw):
    raw = raw.copy()
    raw['MA150']   = raw['Close'].rolling(150).mean()
    raw['고가99']   = raw['High'].shift(1).rolling(99).max()
    raw['저가99']   = raw['Low'].shift(1).rolling(99).min()
    raw['고가49']   = raw['High'].shift(1).rolling(49).max()
    raw['저가49']   = raw['Low'].shift(1).rolling(49).min()
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
# KRX 직접 API 호출 (pykrx 대신)
# --------------------------------------------------
def get_krx_ohlcv(code, start_date, end_date):
    """KRX 정보데이터시스템 직접 호출 - 실제 주가 반환"""
    url = "http://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": "http://data.krx.co.kr/",
    }
    params = {
        "bld": "dbms/MDC/STAT/standard/MDCSTAT01701",
        "isuCd": code,
        "isuCd2": "",
        "strtDd": start_date,
        "endDd": end_date,
        "adjStkPrc_check": "N",   # 수정주가 미적용 -> 실제 주가
        "adjStkPrc": "1",
        "share": "1",
        "money": "1",
        "csvxls_isNo": "false",
    }
    try:
        resp = requests.post(url, headers=headers, data=params, timeout=30)
        data = resp.json()
        rows = data.get("output", [])
        if not rows:
            return pd.DataFrame()

        df = pd.DataFrame(rows)
        df = df.rename(columns={
            "TRD_DD":   "Date",
            "TDD_OPNPRC": "Open",
            "TDD_HGPRC":  "High",
            "TDD_LWPRC":  "Low",
            "TDD_CLSPRC": "Close",
            "ACC_TRDVOL": "Volume",
        })
        for col in ["Open","High","Low","Close","Volume"]:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(",",""), errors="coerce"
                )
        df["Date"] = pd.to_datetime(df["Date"], format="%Y/%m/%d", errors="coerce")
        df = df[["Date","Close","High","Low","Volume"]].dropna()
        df = df.sort_values("Date").reset_index(drop=True)
        return df
    except Exception:
        return pd.DataFrame()


today       = pd.Timestamp.today()
krx_start   = (today - pd.Timedelta(days=10*365)).strftime("%Y%m%d")
krx_end     = today.strftime("%Y%m%d")

@st.cache_data(ttl=3600)
def get_weinstein_stage(ticker_code):
    try:
        code = str(int(float(ticker_code))).zfill(6)
        raw  = get_krx_ohlcv(code, krx_start, krx_end)
        if raw.empty or len(raw) < 150:
            return "N/A"
        return calc_weinstein_stages_from_df(raw)
    except Exception:
        return "N/A"

# --------------------------------------------------
# 와인스타인 계산
# --------------------------------------------------
if '종목코드' in df.columns:
    with st.spinner("와인스타인 단계 계산 중... (pykrx 실제 주가 기준)"):
        df['와인스타인'] = df['종목코드'].apply(get_weinstein_stage)
else:
    df['와인스타인'] = "N/A"

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
