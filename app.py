import streamlit as st
import pandas as pd
import numpy as np
import requests, time
from bs4 import BeautifulSoup
from scipy.stats import percentileofscore
import plotly.express as px
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

# ===========================
# 최신월 데이터 크롤링 함수
# ===========================
def get_latest_band_prices(gicode, max_retries=2, delay=0.2):
    url = (f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp"
           f"?oid=pbrBandCht&cid=01_06&gicode={gicode}&filter=D&term=Y&etc=B&etc2=0")
    latest_dt, latest_cells = None, None

    for _ in range(max_retries):
        try:
            resp = requests.get(url, headers={"User-Agent":"Mozilla/5.0"}, timeout=5)
            resp.encoding = 'utf-8'
            if resp.status_code != 200:
                time.sleep(delay)
                continue
            rows = BeautifulSoup(resp.text, 'html.parser').select("table tbody tr")
            for tr in rows:
                cells = [td.get_text(strip=True).replace(',', '') for td in tr.find_all('td')]
                if len(cells) >= 7 and cells[0].replace('/', '').isdigit():
                    if latest_dt is None or cells[0] > latest_dt:
                        latest_dt, latest_cells = cells[0], cells
            if latest_cells:
                try:
                    adj_price = float(latest_cells[1]) if latest_cells[1] != '-' else None
                except:
                    adj_price = None
                band_prices = [float(v) for v in latest_cells[2:7] if v != '-']
                if len(band_prices) == 5:
                    return latest_dt, adj_price, band_prices
            time.sleep(delay)
        except:
            time.sleep(delay)
    return None, None, None

# ===========================
# Position 계산
# ===========================
def get_position(current_price, band_prices):
    if current_price is None or np.isnan(current_price) or not band_prices or len(band_prices) != 5:
        return np.nan
    if current_price < band_prices[0]:
        return 1
    for i in range(1,5):
        if band_prices[i-1] <= current_price < band_prices[i]:
            return i+1
    return 6

# ===========================
# 메인 실행
# ===========================
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # 종목코드 6자리 맞추기
    df['종목코드'] = df['종목코드'].astype(str).str.zfill(6)

    # 필수 컬럼 확인
    stochastic_col = next((c for c in df.columns if 'stochastic' in c.lower()), None)
    roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c][:3]
    if not stochastic_col or len(roe_cols) < 3:
        st.error("필수 컬럼(Stochastic, ROE 3개)이 누락됐습니다.")
        st.stop()

    # 숫자 변환
    for col in ['현재가','BPS','배당수익률',stochastic_col,'10년후BPS','복리수익률'] + roe_cols + ['추정ROE']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',','').str.replace('%','').astype(float)

    # 계산 컬럼
    if '추정ROE' not in df.columns:
        df['추정ROE'] = df[roe_cols[0]]*0.4 + df[roe_cols[1]]*0.35 + df[roe_cols[2]]*0.25
    if '10년후BPS' not in df.columns:
        df['10년후BPS'] = (df['BPS']*(1+df['추정ROE']/100)**10).round(0)
    if '복리수익률' not in df.columns:
        df['복리수익률'] = (((df['10년후BPS']/df['현재가'])**(1/10))-1)*100

    # ===== 임시 테이블 수집 =====
    temp_list = []
    for _, row in df.iterrows():
        gicode = f"A{row['종목코드']}"
        latest_date, _, bands = get_latest_band_prices(gicode)
        pos = get_position(row['현재가'], bands)
        temp_list.append({'종목코드': row['종목코드'],
                          '최신월': latest_date,
                          'bands': bands,
                          'position': pos})
        time.sleep(0.1)
    band_df = pd.DataFrame(temp_list)

    # 병합
    df = df.merge(band_df[['종목코드','position']], on='종목코드', how='left')

    # 매력도 계산
    alpha = st.slider('복리수익률(성장성) : 저평가(분위수) 가중치 (%)', 0,100,80,5,format="%d%%")/100
    max_return = df['복리수익률'].max(); min_return = 15.0
    df['복리수익률점수'] = ((df['복리수익률']-min_return)/(max_return-min_return)).clip(lower=0)*100
    df['Stochastic_percentile'] = df[stochastic_col].apply(lambda x: 100 - percentileofscore(df[stochastic_col], x, kind='mean'))
    df['매력도'] = (alpha*df['복리수익률점수'] + (1-alpha)*df['Stochastic_percentile']).round(2)

    # 정렬 및 표시
    df_sorted = df.sort_values(by='매력도', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index+1
    df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

    main_cols = ['순위','종목명','현재가','등락률'] + roe_cols + \
                ['BPS','배당수익률','RN','추정ROE','10년후BPS','복리수익률','position','매력도']
    st.dataframe(df_sorted[main_cols], use_container_width=True, height=500)

    # 차트
    st.plotly_chart(px.scatter(df_sorted, x='RN', y='복리수익률', color='매력도',
                               hover_name='종목명', title='복리수익률 vs RN',
                               color_continuous_scale='Viridis'), use_container_width=True)
    st.plotly_chart(px.bar(df_sorted.head(5), x='종목명', y='매력도', color='매력도',
                           title='매력도 상위 5개', color_continuous_scale='Viridis'), use_container_width=True)
else:
    st.error(f"현재 폴더에 '{file_path}' 파일이 없습니다.")
