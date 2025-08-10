import streamlit as st
import pandas as pd
import numpy as np
from scipy.stats import percentileofscore
import plotly.express as px
import os
import requests
from bs4 import BeautifulSoup
import time

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

############################################################
# FnGuide 팝업에서 최신 수정주가 + 밴드 주가 5개 추출
############################################################
def get_latest_band_prices(gicode, max_retries=3, delay=0.5):
    """
    gicode: 'A005930' 형태
    return: (adj_price, band_prices[5])
    """
    url = f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp?oid=pbrBandCht&cid=01_06&gicode={gicode}&filter=D&term=Y&etc=B&etc2=0"

    for attempt in range(max_retries):
        try:
            resp = requests.get(url, timeout=5, headers={"User-Agent": "Mozilla/5.0"})
            resp.encoding = 'utf-8'
            soup = BeautifulSoup(resp.text, 'html.parser')

            # 표 첫번째 데이터 행
            first_row = soup.select_one("table tbody tr")
            if not first_row:
                time.sleep(delay)
                continue

            cells = [c.get_text(strip=True).replace(',', '') for c in first_row.find_all('td')]
            # 구조: [날짜, 수정주가, 밴드1, 밴드2, 밴드3, 밴드4, 밴드5]
            try:
                adj_price = float(cells[1]) if cells[1] not in ('-', '') else None
            except:
                adj_price = None

            band_prices = []
            for val_str in cells[2:7]: # 밴드 주가 5개
                try:
                    if val_str not in ('-', ''):
                        band_prices.append(float(val_str))
                except:
                    pass

            if len(band_prices) == 5:
                return adj_price, band_prices
            time.sleep(delay)
        except:
            time.sleep(delay)

    return None, None

############################################################
# 현재가와 밴드 주가 5개의 비교로 position 산출
############################################################
def get_position_from_price(current_price, band_prices):
    if current_price is None or np.isnan(current_price) or not band_prices:
        return np.nan
    if current_price < band_prices[0]:
        return 1
    for i in range(1, len(band_prices)):
        if band_prices[i-1] <= current_price < band_prices[i]:
            return i + 1
    return 6

############################################################
# 메인 처리
############################################################
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # Stochastic 컬럼 탐색
    stochastic_col = next((col for col in df.columns if 'stochastic' in col.lower()), None)
    if not stochastic_col:
        st.error("Stochastic 컬럼 없음")
        st.stop()

    # ROE 컬럼 탐색
    roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c][:3]
    if len(roe_cols) < 3:
        st.error(f"ROE 컬럼 부족: {roe_cols}")
        st.stop()

    # 등락률 처리
    if '등락률' in df.columns:
        df['등락률'] = df['등락률'].apply(lambda x: f"{float(x)*100:.2f}%" if not isinstance(x, str) or '%' not in x else x)

    # 숫자형 변환
    num_cols = ['현재가','BPS','배당수익률', stochastic_col, '10년후BPS','복리수익률'] + roe_cols + ['추정ROE']
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',','').str.replace('%','').astype(float)

    # 계산 컬럼
    if '추정ROE' not in df.columns:
        df['추정ROE'] = df[roe_cols[0]]*0.4 + df[roe_cols[1]]*0.35 + df[roe_cols[2]]*0.25
    if '10년후BPS' not in df.columns:
        df['10년후BPS'] = (df['BPS']*(1+df['추정ROE']/100)**10).round(0)
    if '복리수익률' not in df.columns:
        df['복리수익률'] = (((df['10년후BPS']/df['현재가'])**(1/10))-1)*100

    # position 컬럼 계산
    df['position'] = np.nan
    if '종목코드' in df.columns:
        for idx, row in df.iterrows():
            gicode = f"A{str(row['종목코드']).zfill(6)}"
            adj_price, band_prices = get_latest_band_prices(gicode)
            pos = get_position_from_price(row['현재가'], band_prices)
            df.at[idx, 'position'] = pos
            time.sleep(0.2)  # 딜레이

    # 매력도 계산
    alpha = st.slider('복리수익률(성장성) : 저평가(분위수) 가중치 (%)',0,100,80,5,format="%d%%")/100
    max_return = df['복리수익률'].max()
    min_return = 15.0
    df['복리수익률점수'] = ((df['복리수익률']-min_return)/(max_return-min_return)).clip(lower=0)*100
    df['Stochastic_percentile'] = df[stochastic_col].apply(lambda x: 100 - percentileofscore(df[stochastic_col], x, kind='mean'))
    df['매력도'] = (alpha*df['복리수익률점수'] + (1-alpha)*df['Stochastic_percentile']).round(2)

    # 정렬
    df_sorted = df.sort_values(by='매력도', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index+1
    df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

    # 표시 컬럼
    main_cols = ['순위','종목명','현재가','등락률'] + roe_cols + ['BPS','배당수익률','RN','추정ROE','10년후BPS','복리수익률','position','매력도']
    df_show = df_sorted[[c for c in main_cols if c in df_sorted.columns]]

    # 스타일
    def highlight_high_return(row):
        return ['background-color: lightgreen' if col=='종목명' and row['복리수익률']>=15 else '' for col in row.index]
    format_dict = {'현재가':'{:,.0f}', roe_cols[0]:'{:.2f}', roe_cols[1]:'{:.2f}', roe_cols[2]:'{:.2f}',
                   'BPS':'{:,.0f}','배당수익률':'{:.2f}','RN':'{:,.0f}','추정ROE':'{:.2f}',
                   '10년후BPS':'{:,.0f}','복리수익률':'{:.2f}','매력도':'{:.2f}','position':'{:,.0f}'}
    styled_df = df_show.style.apply(highlight_high_return, axis=1).format(format_dict).set_properties(**{'text-align': 'center'})\
        .set_table_styles([{'selector':'th','props':[('text-align','center')]}])
    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

    # 시각화
    st.plotly_chart(px.scatter(df_sorted, x='RN', y='복리수익률', color='매력도', hover_name='종목명', title='복리수익률 vs RN',
                               color_continuous_scale='Viridis'), use_container_width=True)
    st.plotly_chart(px.bar(df_sorted.head(5), x='종목명', y='매력도', color='매력도', title='매력도 상위 5개 종목',
                           color_continuous_scale='Viridis'), use_container_width=True)
else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.")
