import inspect
from datetime import datetime, timedelta, timezone, time as dtime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
import FinanceDataReader as fdr

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

date_placeholder = st.empty()

# --------------------------------------------------
# 설정값 (여기만 수정)
# --------------------------------------------------
EXCEL_NAME     = "1.xlsx"
HIGH_RETURN    = 15            # 복리수익률 강조 기준(%)
DISP_BUY_LINE  = 90            # 30주 이격도: 매수 검토
DISP_HEAT_LINE = 110           # 30주 이격도: 과열
STALE_DAYS     = 7             # 마지막 거래일이 최신일보다 이만큼 오래되면 제외(거래정지 등)
HISTORY_DAYS   = 365 * 3       # 주가 조회 기간
ROE_WEIGHTS    = (0.3, 0.1, 0.6)

KST = timezone(timedelta(hours=9))   # 서버(UTC)와 무관하게 한국 시각 표시

# 제외/경고 내역: (종목명, 사유)
problems = []

# --------------------------------------------------
# CSS 가운데 정렬 (※ 캔버스 기반 dataframe에서는 적용되지 않을 수 있음)
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


# --------------------------------------------------
# 유틸
# --------------------------------------------------
def stretch_kw(fn):
    """Streamlit 버전별 폭 옵션: width='stretch'(신) / use_container_width=True(구)"""
    try:
        p = inspect.signature(fn).parameters.get("width")
        if p is not None and p.default == "stretch":
            return {"width": "stretch"}
    except (TypeError, ValueError):
        pass
    return {"use_container_width": True}


def to_numeric_safe(series):
    s = (series.astype(str)
               .str.strip()
               .str.replace(',', '', regex=False)
               .str.replace('%', '', regex=False)
               .str.replace(r'^\((.*)\)$', r'-\1', regex=True))   # (1,234) → -1234
    s = s.replace({'': np.nan, '-': np.nan, 'nan': np.nan, 'NaN': np.nan,
                   'None': np.nan, 'N/A': np.nan, '#N/A': np.nan})
    return pd.to_numeric(s, errors='coerce')


def normalize_percent(series):
    """컬럼 단위로 판단: 최대값이 1 이하이면 비율(0.025)로 보고 ×100.
    (셀 단위로 판단하면 1.0% 같은 정상 값이 100%로 바뀌는 오류가 생김)"""
    s = to_numeric_safe(series)
    mx = s.abs().max()
    if pd.notna(mx) and mx <= 1:
        s = s * 100
    return s.round(2)


def normalize_code(x):
    """종목코드 → 6자리 문자열. 숫자코드는 zfill, 영숫자 신규코드는 그대로."""
    s = str(x).strip().upper()
    if s in ("", "NAN", "NONE", "<NA>"):
        return None
    if s.endswith(".0"):
        s = s[:-2]
    if len(s) == 7 and s[0] == "A" and s[1:].isalnum():   # 'A005930' 형식
        s = s[1:]
    if s.isdigit() and len(s) <= 6:
        return s.zfill(6)
    if len(s) == 6 and s.isalnum():
        return s
    return None


def drop_rows(frame, mask, reason=None):
    """mask에 해당하는 행 제거. reason이 있으면 제외 내역에 기록(이미 기록된 경우 None)."""
    if reason:
        for name in frame.loc[mask, '종목명']:
            problems.append((name, reason))
    return frame[~mask].copy()


def render_problems():
    if problems:
        with st.expander(f"⚠️ 제외/경고 종목 {len(problems)}개"):
            st.dataframe(
                pd.DataFrame(problems, columns=['종목명', '사유']),
                hide_index=True,
                **stretch_kw(st.dataframe)
            )


def stop_with(msg, level="error"):
    getattr(st, level)(msg)
    render_problems()
    st.stop()


# --------------------------------------------------
# 데이터 로드
# ※ 이격도는 엑셀의 '이격도' 컬럼(30주 이격도)을 그대로 사용
# --------------------------------------------------
excel_path = next(
    (p for p in (Path(__file__).resolve().parent / EXCEL_NAME, Path(EXCEL_NAME)) if p.exists()),
    None
)
if excel_path is None:
    st.error(f"'{EXCEL_NAME}' 파일이 존재하지 않습니다.")
    st.stop()

try:
    df = pd.read_excel(excel_path)
except Exception as e:
    st.error(f"엑셀을 읽지 못했습니다: {type(e).__name__}: {e}")
    st.stop()

df.columns = df.columns.astype(str).str.strip()
if df.empty:
    st.error("엑셀 파일이 비어 있습니다.")
    st.stop()

required_base_cols = ['종목명', '종목코드', 'BPS', '이격도']
for col in required_base_cols:
    if col not in df.columns:
        st.error(f"필수 컬럼 누락: {col}")
        st.stop()

roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c]
if len(roe_cols) < 3:
    st.error("ROE 컬럼이 3개 이상 필요합니다.")
    st.stop()
roe_cols = roe_cols[:3]
st.caption(
    "ROE 가중 → " +
    " + ".join(f"{c}×{w}" for c, w in zip(roe_cols, ROE_WEIGHTS)) +
    "  |  30주 이격도는 1.xlsx 값(엑셀 갱신 시점 기준)이라 현재가와 시점이 다를 수 있음"
)

# 숫자 변환
df['종목명'] = df['종목명'].fillna(df['종목코드'].astype(str))
df['BPS'] = to_numeric_safe(df['BPS'])
for col in roe_cols:
    df[col] = to_numeric_safe(df[col])
df['이격도'] = to_numeric_safe(df['이격도']).round(2)
if '배당수익률' in df.columns:
    df['배당수익률'] = normalize_percent(df['배당수익률'])

df['_code'] = df['종목코드'].map(normalize_code)
df = drop_rows(df, df['_code'].isna(), "종목코드 형식 오류/결측")
df = drop_rows(df, df['BPS'].isna(), "BPS 결측")
df = drop_rows(df, df[roe_cols].isna().any(axis=1), "ROE 결측(3개 중 1개 이상)")
if df.empty:
    stop_with("유효한 종목이 없습니다.", "warning")

# 이격도 단위/결측 점검
if df['이격도'].notna().sum() == 0:
    st.warning("'이격도' 컬럼에 유효한 숫자가 없습니다. (엑셀 수식이 계산되지 않은 채 저장되면 이렇게 읽힙니다)")
else:
    med = df['이격도'].median()
    if not (50 <= med <= 200):
        st.warning(f"이격도 중앙값이 {med:.2f}입니다. 100 기준(30주 이평=100) 값이 맞는지 확인하세요.")
roe_med = np.nanmedian(np.abs(df[roe_cols].to_numpy(dtype=float)))
if roe_med <= 1:
    st.warning("ROE가 소수(0.12)로 저장된 것 같습니다. 이 앱은 % 단위(12)를 가정합니다.")


# --------------------------------------------------
# 와인스타인 4단계 계산
# --------------------------------------------------
def calc_weinstein_stage(raw):
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
            continue

        close, ma150 = close_arr[i], ma150_arr[i]
        nh100, nl100 = nh100_arr[i], nl100_arr[i]
        nh50,  nl50  = nh50_arr[i],  nl50_arr[i]
        prev = stages[i-1] if i > 0 else None

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
# FinanceDataReader로 주가 수집
# ※ 실패 시 예외를 그대로 올린다 → st.cache_data는 예외를 캐시하지 않으므로
#    일시적 네트워크 오류가 1시간 동안 굳어지지 않음
# --------------------------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def get_stock_data(code):
    """(현재가, 등락률%, 와인스타인 단계, 마지막 거래일) 반환. 실패 시 예외."""
    today     = datetime.now(KST).date()
    from_date = (today - timedelta(days=HISTORY_DAYS)).strftime("%Y-%m-%d")
    to_date   = today.strftime("%Y-%m-%d")

    raw = fdr.DataReader(code, from_date, to_date)
    if raw is None or raw.empty:
        raise ValueError("가격 데이터 없음")

    raw = raw.reset_index()
    raw = raw.rename(columns={raw.columns[0]: 'Date'})   # 인덱스(날짜)가 첫 컬럼
    for c in ('Close', 'High', 'Low'):
        if c not in raw.columns:
            raise ValueError(f"컬럼 누락: {c}")
        raw[c] = pd.to_numeric(raw[c], errors='coerce')

    raw['Date'] = pd.to_datetime(raw['Date'])
    raw = raw.dropna(subset=['Close', 'High', 'Low'])
    raw = raw[(raw['Close'] > 0) & (raw['High'] > 0) & (raw['Low'] > 0)]
    raw = (raw.sort_values('Date')
              .drop_duplicates('Date', keep='last')
              .reset_index(drop=True))
    if len(raw) < 2:
        raise ValueError("유효 데이터 2일 미만")

    current_price = float(raw['Close'].iloc[-1])
    last_date     = raw['Date'].iloc[-1]

    change = raw['Change'].iloc[-1] if 'Change' in raw.columns else np.nan
    if pd.isna(change):                                   # Change 컬럼이 없거나 결측이면 종가로 계산
        change = raw['Close'].pct_change().iloc[-1]
    change_pct = float(change) * 100

    stage = calc_weinstein_stage(raw) if len(raw) >= 150 else "N/A"
    return current_price, change_pct, stage, last_date


# --------------------------------------------------
# 일괄 수집
# --------------------------------------------------
rows = []
with st.spinner("KRX 주가 데이터 수집 중..."):
    progress = st.progress(0)
    total = len(df)
    for i, (name, code) in enumerate(zip(df['종목명'], df['_code'])):
        try:
            rows.append(get_stock_data(code))
        except Exception as e:
            rows.append((np.nan, np.nan, "N/A", pd.NaT))
            problems.append((name, f"주가 조회 실패: {type(e).__name__}: {e}"))
        progress.progress((i + 1) / total)
    progress.empty()

results = pd.DataFrame(rows, columns=['현재가', '등락률', '와인스타인', '기준일'], index=df.index)
df['현재가']     = results['현재가']
df['등락률']     = results['등락률'].round(2)
df['와인스타인'] = results['와인스타인']
df['기준일']     = pd.to_datetime(results['기준일'])
# ※ df['이격도']는 엑셀 값 유지 (덮어쓰지 않음)

df = drop_rows(df, df['현재가'].isna())   # 사유는 수집 단계에서 이미 기록됨
if df.empty:
    stop_with("주가 데이터를 가져온 종목이 없습니다. (네트워크/데이터 소스 확인)", "warning")

latest_date = df['기준일'].max()
stale = df['기준일'] < (latest_date - pd.Timedelta(days=STALE_DAYS))
df = drop_rows(df, stale, f"최근 {STALE_DAYS}일 이상 거래 없음(거래정지/상장폐지 의심)")

now_kst  = datetime.now(KST)
intraday = (latest_date.date() == now_kst.date()) and (now_kst.time() < dtime(15, 40))
price_label = "장중 시세 — 종가 아님" if intraday else "KRX 종가 기준"
date_placeholder.caption(
    f"📅 **데이터 기준일:** {latest_date.strftime('%Y-%m-%d')} ({price_label})  |  "
    f"🔄 **조회 시각(KST):** {now_kst.strftime('%Y-%m-%d %H:%M')}"
)

# --------------------------------------------------
# 계산
# --------------------------------------------------
w0, w1, w2 = ROE_WEIGHTS
df['추정ROE'] = df[roe_cols[0]] * w0 + df[roe_cols[1]] * w1 + df[roe_cols[2]] * w2

growth = 1 + df['추정ROE'] / 100
df['10년후BPS'] = (df['BPS'] * growth ** 10).where(growth > 0)
df['10년후BPS'] = df['10년후BPS'].replace([np.inf, -np.inf], np.nan)

valid = (df['현재가'] > 0) & (df['10년후BPS'] > 0)      # 음수 BPS 등 → 제외
ratio = (df['10년후BPS'] / df['현재가']).where(valid)
df['복리수익률'] = ((ratio ** 0.1 - 1) * 100).replace([np.inf, -np.inf], np.nan).round(2)
df = drop_rows(df, df['복리수익률'].isna(), "복리수익률 계산 불가(BPS≤0 등)")

if df.empty:
    stop_with("계산 후 표시할 데이터가 없습니다.", "warning")

# --------------------------------------------------
# 정렬 / 표시
# --------------------------------------------------
df_sorted = df.sort_values(by='복리수익률', ascending=False, kind='stable').reset_index(drop=True)
df_sorted['순위'] = df_sorted.index + 1

display_cols = [
    '순위', '종목명', '현재가', '등락률',
    '배당수익률', '추정ROE',
    'BPS', '10년후BPS',
    '복리수익률', '이격도',
    '와인스타인'
]
existing_cols = [c for c in display_cols if c in df_sorted.columns]
df_show = df_sorted[existing_cols].rename(columns={'이격도': '30주 이격도'})


def highlight_high_return(row):
    color = 'background-color: lightgreen' if row['복리수익률'] >= HIGH_RETURN else ''
    return [color] * len(row)


format_dict = {
    '현재가':      '{:,.0f}',
    '등락률':      '{:.2f}%',
    '배당수익률':   '{:.2f}%',
    '추정ROE':    '{:.2f}',
    'BPS':        '{:,.0f}',
    '10년후BPS':   '{:,.0f}',
    '복리수익률':   '{:.2f}%',
    '30주 이격도':  '{:.2f}',
}
format_dict = {k: v for k, v in format_dict.items() if k in df_show.columns}   # 없는 컬럼 KeyError 방지

styled_df = (
    df_show.style
           .apply(highlight_high_return, axis=1)
           .format(format_dict, na_rep="-")
)

calculated_height = min(len(df_show) * 35 + 60, 1000)
st.dataframe(
    styled_df,
    height=calculated_height,
    hide_index=True,
    **stretch_kw(st.dataframe)
)

# --------------------------------------------------
# 산점도: 30주 이격도 vs 복리수익률
# 이격도 낮음 + 복리수익률 높음 = 최적 매수 후보
# --------------------------------------------------
df_plot = df_sorted.dropna(subset=['이격도']).copy()
n_no_disp = len(df_sorted) - len(df_plot)

if not df_plot.empty:
    hi_label, lo_label = f"{HIGH_RETURN}% 이상", f"{HIGH_RETURN}% 미만"
    df_plot['구분'] = np.where(df_plot['복리수익률'] >= HIGH_RETURN, hi_label, lo_label)

    hover = {'이격도': ':.2f', '복리수익률': ':.2f', '와인스타인': True, '구분': False}
    if '배당수익률' in df_plot.columns:
        hover['배당수익률'] = ':.2f'

    fig = px.scatter(
        df_plot,
        x='이격도',
        y='복리수익률',
        color='구분',
        color_discrete_map={hi_label: '#2ecc71', lo_label: '#3498db'},
        category_orders={'구분': [hi_label, lo_label]},
        hover_name='종목명',
        hover_data=hover,
        title='복리수익률 vs 30주 이격도',
    )

    fig.add_vline(
        x=DISP_BUY_LINE, line_dash="dash", line_color="green",
        annotation_text=f"이격도 {DISP_BUY_LINE} (매수 검토)",
        annotation_position="top right", annotation_font_color="green"
    )
    fig.add_vline(
        x=DISP_HEAT_LINE, line_dash="dash", line_color="red",
        annotation_text=f"이격도 {DISP_HEAT_LINE} (과열)",
        annotation_position="top left", annotation_font_color="red"
    )
    fig.add_hline(
        y=HIGH_RETURN, line_dash="dot", line_color="orange",
        annotation_text=f"복리수익률 {HIGH_RETURN}%",
        annotation_position="right", annotation_font_color="orange"
    )
    fig.update_layout(
        xaxis=dict(title="30주 이격도 (30주 이평 = 100)"),
        yaxis=dict(title="복리수익률(%)"),
        legend_title=f"복리 {HIGH_RETURN}% 기준",
    )
    st.plotly_chart(fig, **stretch_kw(st.plotly_chart))
    if n_no_disp:
        st.caption(f"※ 30주 이격도가 없는 {n_no_disp}개 종목은 산점도에서 제외됨")
else:
    st.info("산점도에 표시할 이격도 값이 없습니다.")

render_problems()
