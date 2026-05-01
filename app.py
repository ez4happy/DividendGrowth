import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import FinanceDataReader as fdr
from datetime import datetime, timedelta
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

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
# 필수 컬럼 체크
# ※ stochastic_col 불필요 (25일 이격도로 대체)
# --------------------------------------------------
required_base_cols = ['종목명', '종목코드', 'BPS']
for col in required_base_cols:
    if col not in df.columns:
        st.error(f"필수 컬럼 누락: {col}")
        st.stop()

roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c]
if len(roe_cols) < 3:
    st.error("ROE 컬럼이 3개 이상 필요합니다.")
    st.stop()
roe_cols = roe_cols[:3]

if '배당수익률' in df.columns:
    df['배당수익률'] = normalize_percent(df['배당수익률'])

num_cols = ['BPS'] + roe_cols
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
# FinanceDataReader로 주가 데이터 가져오기
# ※ 25일 이격도 추가 (기존 RN 대체)
#    이격도 = 현재가 / 25일 이평 × 100
#    100 = 이평과 일치
#     91 = 이평 대비 -9% (BNF 매수 검토 구간)
#    110 = 이평 대비 +10% (과열 구간)
# --------------------------------------------------
@st.cache_data(ttl=3600)
def get_stock_data(ticker_code):
    """(현재가, 등락률%, 와인스타인 단계, 25일 이격도, 마지막 거래일) 반환"""
    try:
        code = str(int(float(ticker_code))).zfill(6)
        today     = datetime.today()
        from_date = (today - timedelta(days=365*3)).strftime("%Y-%m-%d")
        to_date   = today.strftime("%Y-%m-%d")

        raw = fdr.DataReader(code, from_date, to_date)
        if raw is None or raw.empty:
            return np.nan, np.nan, "N/A", np.nan, pd.NaT

        raw = raw.reset_index()
        raw = raw[['Date', 'Close', 'High', 'Low', 'Volume', 'Change']].dropna(
            subset=['Close', 'High', 'Low']
        )
        raw = raw.sort_values('Date').reset_index(drop=True)

        if len(raw) < 2:
            return np.nan, np.nan, "N/A", np.nan, pd.NaT

        current_price = float(raw['Close'].iloc[-1])
        change_pct    = float(raw['Change'].iloc[-1]) * 100
        last_date     = pd.to_datetime(raw['Date'].iloc[-1])

        # 25일 이격도 계산
        # 데이터가 25일 이상일 때만 계산
        if len(raw) >= 25:
            ma25      = raw['Close'].rolling(25).mean()
            ma25_last = float(ma25.iloc[-1])
            if ma25_last > 0:
                ikgyuk_25 = round(current_price / ma25_last * 100, 1)
            else:
                ikgyuk_25 = np.nan
        else:
            ikgyuk_25 = np.nan

        stage = calc_weinstein_stages_from_df(raw) if len(raw) >= 150 else "N/A"
        return current_price, change_pct, stage, ikgyuk_25, last_date

    except Exception:
        return np.nan, np.nan, "N/A", np.nan, pd.NaT


# --------------------------------------------------
# 일괄 수집
# --------------------------------------------------
with st.spinner("KRX 주가 데이터 수집 중..."):
    progress = st.progress(0)
    total    = len(df)
    rows     = []
    for i, code in enumerate(df['종목코드'].tolist()):
        rows.append(get_stock_data(code))
        progress.progress((i + 1) / total)
    progress.empty()

results = pd.DataFrame(
    rows,
    columns=['현재가', '등락률', '와인스타인', '이격도', '기준일'],
    index=df.index
)
df['현재가']     = results['현재가']
df['등락률']     = results['등락률'].round(2)
df['와인스타인'] = results['와인스타인']
df['이격도']     = results['이격도']   # ← 25일 이격도 (100 기준)
df['기준일']     = results['기준일']

latest_date = df['기준일'].dropna().max()

df.dropna(subset=['현재가'], inplace=True)
if df.empty:
    st.warning("주가 데이터를 가져온 종목이 없습니다.")
    st.stop()

if pd.notna(latest_date):
    now_str  = datetime.now().strftime("%Y-%m-%d %H:%M")
    date_str = latest_date.strftime("%Y-%m-%d")
    date_placeholder.caption(
        f"📅 **데이터 기준일:** {date_str} (KRX 종가 기준)  |  🔄 **조회 시각:** {now_str}"
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
# 정렬 / 표시
# --------------------------------------------------
df_sorted = df.sort_values(by='복리수익률', ascending=False).reset_index(drop=True)
df_sorted['순위'] = df_sorted.index + 1

display_cols = [
    '순위', '종목명', '현재가', '등락률',
    '배당수익률', '추정ROE',
    'BPS', '10년후BPS',
    '복리수익률', '이격도',   # ← RN 대신 이격도
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
    '현재가':     '{:,.0f}',
    '등락률':     '{:.2f}%',
    '배당수익률':  '{:.2f}%',
    '추정ROE':   '{:.2f}',
    'BPS':       '{:,.0f}',
    '10년후BPS':  '{:,.0f}',
    '복리수익률':  '{:.2f}%',
    '이격도':     '{:.1f}',   # ← 소수점 1자리 (예: 91.6)
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
# 이격도 가이드 표시
# --------------------------------------------------
st.caption(
    "📌 **이격도 (25일 이평 기준)**: "
    "100 = 이평과 일치  |  "
    "🟢 90 이하 = BNF 매수 검토 구간  |  "
    "🔴 110 이상 = 과열 구간"
)

# --------------------------------------------------
# 산점도: 이격도 vs 복리수익률
# ※ x축: 이격도(25일), y축: 복리수익률
# 이격도 낮음 + 복리수익률 높음 = 최적 매수 후보
# --------------------------------------------------
df_plot = df_sorted.dropna(subset=['이격도'])

if not df_plot.empty:
    df_plot['HighReturn'] = df_plot['복리수익률'] >= 15

    # 이격도 90~110 참조선 영역 표시
    fig = px.scatter(
        df_plot,
        x='이격도',
        y='복리수익률',
        color='HighReturn',
        color_discrete_map={True: '#2ecc71', False: '#3498db'},
        hover_name='종목명',
        hover_data={
            '이격도': ':.1f',
            '복리수익률': ':.2f',
            '배당수익률': ':.2f',
            '와인스타인': True,
            'HighReturn': False,
        },
        title='복리수익률 vs 25일 이격도',
        labels={
            '이격도': '이격도 (25일 이평 = 100)',
            '복리수익률': '복리수익률(%)',
            'HighReturn': '15% 이상'
        }
    )

    # 이격도 90 수직 참조선 (BNF 매수 검토 기준)
    fig.add_vline(
        x=90,
        line_dash="dash",
        line_color="green",
        annotation_text="이격도 90 (매수 검토)",
        annotation_position="top right",
        annotation_font_color="green"
    )

    # 이격도 110 수직 참조선 (과열 기준)
    fig.add_vline(
        x=110,
        line_dash="dash",
        line_color="red",
        annotation_text="이격도 110 (과열)",
        annotation_position="top left",
        annotation_font_color="red"
    )

    # 복리수익률 15% 수평 참조선
    fig.add_hline(
        y=15,
        line_dash="dot",
        line_color="orange",
        annotation_text="복리수익률 15%",
        annotation_position="right",
        annotation_font_color="orange"
    )

    fig.update_layout(
        xaxis=dict(title="이격도 (25일 이평 = 100)"),
        yaxis=dict(title="복리수익률(%)"),
        legend_title="복리 15% 이상",
    )

    st.plotly_chart(fig, use_container_width=True)

    # 최적 매수 후보 (이격도 ≤ 90 + 복리 ≥ 15%)
    st.subheader("🎯 최적 매수 후보 (이격도 ≤ 90 + 복리수익률 ≥ 15%)")
    best = df_plot[
        (df_plot['이격도'] <= 90) & (df_plot['복리수익률'] >= 15)
    ].sort_values('복리수익률', ascending=False)

    if best.empty:
        st.info("현재 조건을 만족하는 종목이 없습니다. (이격도 ≤ 90 + 복리 ≥ 15%)")
    else:
        best_cols = [c for c in ['종목명', '현재가', '이격도', '복리수익률',
                                  '배당수익률', '와인스타인'] if c in best.columns]
        st.dataframe(
            best[best_cols].reset_index(drop=True),
            use_container_width=True,
            hide_index=True
        )
