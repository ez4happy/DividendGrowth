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
st.title("📈 Dividend Growth Stock (Debug)")

file_path = "1.xlsx"

############################################################
# 디버깅용: FnGuide 팝업에서 표 전체를 읽어 각 행의 cells 내용까지 출력
############################################################
def get_latest_band_prices_debug(gicode, max_retries=2, delay=0.2):
    url = (
        f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp"
        f"?oid=pbrBandCht&cid=01_06&gicode={gicode}&filter=D&term=Y&etc=B&etc2=0"
    )
    latest_dt = None
    latest_cells = None

    for _ in range(max_retries):
        try:
            resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=5)
            resp.encoding = 'utf-8'
            if resp.status_code != 200:
                time.sleep(delay)
                continue
            soup = BeautifulSoup(resp.text, 'html.parser')
            rows = soup.select("table tbody tr")
            
            all_cells = []
            for tr in rows:
                cells = [td.get_text(strip=True).replace(',', '') for td in tr.find_all('td')]
                all_cells.append(cells)
                # 최신 날짜행 판별
                if len(cells) >= 7:
                    dt = cells[0]
                    if len(dt) == 10 and dt.replace('/', '').isdigit():
                        if latest_dt is None or dt > latest_dt:
                            latest_dt = dt
                            latest_cells = cells

            # 최신 행 결과 반환 + 전체 cells 리스트도 리턴
            if latest_cells:
                try:
                    adj_price = float(latest_cells[1]) if latest_cells[1] != '-' else None
                except:
                    adj_price = None
                band_prices = []
                for val in latest_cells[2:7]:
                    if val != '-':
                        band_prices.append(float(val))
                if len(band_prices) == 5:
                    return latest_dt, adj_price, band_prices, latest_cells, all_cells
            time.sleep(delay)
        except Exception as e:
            time.sleep(delay)
    return None, None, None, None, None

def get_position_from_price(price, band_prices):
    if (
        price is None or 
        np.isnan(price) or 
        not band_prices or 
        len(band_prices) != 5
    ):
        return np.nan
    if price < band_prices[0]:
        return 1
    for i in range(1, 5):
        if band_prices[i-1] <= price < band_prices[i]:
            return i + 1
    return 6

############################################################
# 메인
############################################################
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # Stochastic 자동 추출
    stochastic_col = next((col for col in df.columns if 'stochastic' in col.lower()), None)
    if not stochastic_col:
        st.error("Stochastic 컬럼을 찾을 수 없습니다."); st.stop()

    roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c]
    if len(roe_cols) < 3:
        st.error(f"ROE 컬럼이 3개 필요. 현재: {roe_cols}"); st.stop()
    roe_cols = roe_cols[:3]

    # 등락률 변환
    if '등락률' in df.columns:
        df['등락률'] = df['등락률'].apply(lambda x: f"{float(x)*100:.2f}%" if not isinstance(x, str) or '%' not in x else x)
    num_cols = ['현재가','BPS','배당수익률', stochastic_col,'10년후BPS','복리수익률'] + roe_cols + ['추정ROE']
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

    # =====================
    # 임시 테이블에 크롤 결과/파싱 로깅
    # =====================
    temp_data = []
    for idx, row in df.iterrows():
        gicode = f"A{str(row['종목코드']).zfill(6)}"
        latest_date, adj, bands, latest_cells, all_cells = get_latest_band_prices_debug(gicode)
        pos = get_position_from_price(row['현재가'], bands)
        temp_data.append({
            '종목코드': row['종목코드'],
            '종목명': row['종목명'],
            '현재가': row['현재가'],
            '최신월': latest_date,
            'adj_price': adj,
            'bands': bands,
            'position': pos,
            'latest_cells': latest_cells,
            'all_cells': str(all_cells)
        })
        time.sleep(0.1)

    band_df = pd.DataFrame(temp_data)
    # Streamlit에서 밴드값/position/크롤링 데이터를 바로 보여줌
    st.subheader('🔎 밴드 크롤링/파싱 결과 원시 로그 (디버깅)')
    st.dataframe(band_df[['종목코드','종목명','현재가','최신월','adj_price','bands','position','latest_cells','all_cells']].astype(str), height=500)

    # position만 메인 df에 merge
    df = df.merge(band_df[['종목코드','position']], on='종목코드', how='left')

    # 매력도 계산 등 원래 로직
    alpha = st.slider('복리수익률(성장성) : 저평가(분위수) 가중치 (%)',
                      0,100,80,5,format="%d%%")/100
    max_return = df['복리수익률'].max()
    min_return = 15.0
    df['복리수익률점수'] = ((df['복리수익률']-min_return)/(max_return-min_return)).clip(lower=0)*100
    df['Stochastic_percentile'] = df[stochastic_col].apply(
        lambda x: 100 - percentileofscore(df[stochastic_col], x, kind='mean'))
    df['매력도'] = (alpha*df['복리수익률점수'] +
                   (1-alpha)*df['Stochastic_percentile']).round(2)

    df_sorted = df.sort_values(by='매력도', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index+1
    df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

    main_cols = ['순위','종목명','현재가','등락률'] + roe_cols + \
                ['BPS','배당수익률','RN','추정ROE','10년후BPS','복리수익률','position','매력도']
    df_show = df_sorted[[c for c in main_cols if c in df_sorted.columns]]

    def highlight_high_return(row):
        return ['background-color: lightgreen' if col=='종목명' and row['복리수익률']>=15 else '' for col in row.index]
    format_dict = {'현재가':'{:,.0f}', roe_cols[0]:'{:.2f}', roe_cols[1]:'{:.2f}', roe_cols[2]:'{:.2f}',
                   'BPS':'{:,.0f}','배당수익률':'{:.2f}','RN':'{:,.0f}','추정ROE':'{:.2f}',
                   '10년후BPS':'{:,.0f}','복리수익률':'{:.2f}','매력도':'{:.2f}','position':'{:,.0f}'}
    styled_df = (df_show.style.apply(highlight_high_return, axis=1)
        .format(format_dict)
        .set_properties(**{'text-align': 'center'})
        .set_table_styles([{'selector':'th','props':[('text-align','center')]}]))
    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

    # 차트
    st.plotly_chart(px.scatter(df_sorted, x='RN', y='복리수익률', color='매력도',
                               hover_name='종목명', title='복리수익률 vs RN 산점도',
                               color_continuous_scale='Viridis'), use_container_width=True)
    st.plotly_chart(px.bar(df_sorted.head(5), x='종목명', y='매력도',
                           color='매력도', title='매력도 상위 5개 종목',
                           color_continuous_scale='Viridis'), use_container_width=True)

else:
    st.error(f"현재 폴더에 '{file_path}' 파일이 없습니다.")
