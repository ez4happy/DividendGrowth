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
# FnGuide PBR 밴드 팝업에서 5개 밴드값 + 현재 PBR 가져오기
############################################################
def get_pbr_band_and_pbr(gicode):
    popup_url = (
        f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp"
        f"?oid=pbrBandCht&cid=01_06&gicode={gicode}"
        f"&filter=D&term=Y&etc=B&etc2=0"
        f"&titleTxt=PBR%20Band&dateTxt=undefined&unitTxt="
    )
    try:
        resp = requests.get(popup_url, timeout=5.0, headers={"User-Agent": "Mozilla/5.0"})
        resp.encoding = 'utf-8'
        soup = BeautifulSoup(resp.text, 'html.parser')

        # 📌 팝업 상단에 위치한 5개 PBR 밴드 추출
        band_list = []
        for th in soup.select("div.chartData table thead tr th"):
            try:
                val = float(th.get_text(strip=True))
                band_list.append(val)
            except:
                continue

        if len(band_list) < 5:
            st.warning(f"{gicode} 밴드값 부족 → 기본값 사용")
            band_list = [0.9, 1.1, 1.3, 1.5, 1.7]

        # 현재 PBR은 메인 페이지에서 가져오는 게 정확
        main_url = f"https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode={gicode}"
        main_resp = requests.get(main_url, timeout=5.0, headers={"User-Agent": "Mozilla/5.0"})
        main_soup = BeautifulSoup(main_resp.text, 'html.parser')
        current_pbr = None
        tag = main_soup.find(lambda t: t.name == 'th' and ('PBR' in t.get_text() or '주가순자산비율' in t.get_text()))
        if tag:
            td = tag.find_next_sibling('td')
            try:
                current_pbr = float(td.get_text(strip=True).replace(',', ''))
            except:
                pass

        return band_list, current_pbr
    except Exception as e:
        st.warning(f"{gicode} PBR 밴드 크롤링 오류: {e}")
        return [0.9, 1.1, 1.3, 1.5, 1.7], None

############################################################
# Position 계산
############################################################
def get_position(pbr, bands):
    if pd.isna(pbr):
        return np.nan
    for i, band in enumerate(bands):
        if pbr < band:
            return i + 1
    return 6

############################################################
# 메인 로직
############################################################
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # Stochastic 자동 탐색
    stochastic_col = next((col for col in df.columns if 'stochastic' in col.lower()), None)
    if not stochastic_col:
        st.error("Stochastic 컬럼을 찾을 수 없습니다.")
        st.stop()

    # ROE 컬럼 자동 탐색
    roe_cols = [c for c in df.columns if 'ROE' in c and '평균' not in c and '최종' not in c][:3]
    if len(roe_cols) < 3:
        st.error(f"ROE 컬럼이 3개 필요합니다. 현재: {roe_cols}")
        st.stop()

    # 등락률 % 처리
    if '등락률' in df.columns:
        df['등락률'] = df['등락률'].apply(lambda x: f"{float(x)*100:.2f}%" if not isinstance(x, str) or '%' not in x else x)

    # 숫자형 변환
    num_cols = ['현재가','BPS','배당수익률', stochastic_col,'10년후BPS','복리수익률'] + roe_cols + ['추정ROE']
    for col in num_cols:
        if col in df.columns:
            df[col] = (df[col].astype(str).str.replace(',', '').str.replace('%','').astype(float))

    # 계산 컬럼
    if '추정ROE' not in df.columns:
        df['추정ROE'] = df[roe_cols[0]]*0.4 + df[roe_cols[1]]*0.35 + df[roe_cols[2]]*0.25
    if '10년후BPS' not in df.columns:
        df['10년후BPS'] = (df['BPS'] * (1+df['추정ROE']/100) ** 10).round(0)
    if '복리수익률' not in df.columns:
        df['복리수익률'] = (((df['10년후BPS'] / df['현재가'])**(1/10))-1)*100

    # Position 계산
    df['position'] = np.nan
    if '종목코드' in df.columns:
        for idx, row in df.iterrows():
            gicode = f"A{str(row['종목코드']).zfill(6)}"
            bands, pbr_now = get_pbr_band_and_pbr(gicode)
            if pbr_now is None and row['BPS'] and row['BPS'] != 0:
                pbr_now = row['현재가'] / row['BPS']
            df.at[idx, 'position'] = get_position(pbr_now, bands)
            time.sleep(0.3)

    # 매력도 계산
    alpha = st.slider('복리수익률(성장성) : 저평가(분위수) 가중치 (%)', 0, 100, 80, 5, format="%d%%")/100
    max_return = df['복리수익률'].max()
    min_return = 15.0
    df['복리수익률점수'] = ((df['복리수익률'] - min_return)/(max_return-min_return)).clip(lower=0)*100
    df['Stochastic_percentile'] = df[stochastic_col].apply(lambda x: 100 - percentileofscore(df[stochastic_col], x, kind='mean'))
    df['매력도'] = (alpha*df['복리수익률점수'] + (1-alpha)*df['Stochastic_percentile']).round(2)

    # 정렬
    df_sorted = df.sort_values(by='매력도', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index+1
    df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

    # 출력 컬럼
    main_cols = ['순위','종목명','현재가','등락률'] + roe_cols + ['BPS','배당수익률','RN','추정ROE','10년후BPS','복리수익률','position','매력도']
    df_show = df_sorted[[c for c in main_cols if c in df_sorted.columns]]

    # 스타일
    def highlight_high_return(r):
        return ['background-color: lightgreen' if col=='종목명' and r['복리수익률']>=15 else '' for col in r.index]
    format_dict = {'현재가': '{:,.0f}', roe_cols[0]: '{:.2f}', roe_cols[1]: '{:.2f}', roe_cols[2]: '{:.2f}',
                   'BPS': '{:,.0f}','배당수익률':'{:.2f}','RN':'{:,.0f}','추정ROE':'{:.2f}',
                   '10년후BPS':'{:,.0f}','복리수익률':'{:.2f}','매력도':'{:.2f}','position':'{:,.0f}'}
    styled_df = df_show.style.apply(highlight_high_return, axis=1).format(format_dict).set_properties(**{'text-align':'center'})\
        .set_table_styles([{'selector':'th','props':[('text-align','center')]}])
    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

    # 차트
    st.plotly_chart(px.scatter(df_sorted, x='RN', y='복리수익률', color='매력도',
                               hover_name='종목명', title='복리수익률 vs RN 산점도',
                               color_continuous_scale='Viridis'), use_container_width=True)
    st.plotly_chart(px.bar(df_sorted.head(5), x='종목명', y='매력도', color='매력도',
                           title='매력도 상위 5개 종목', color_continuous_scale='Viridis'),
                    use_container_width=True)
else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.")
