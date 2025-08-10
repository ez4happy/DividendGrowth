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
# FnGuide 팝업 차트에서 5개 PBR 밴드 숫자 추출 함수
############################################################
def get_pbr_band(gicode, max_retries=3, delay=0.5):
    popup_url = (
        f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp"
        f"?oid=pbrBandCht&cid=01_06&gicode={gicode}"
        f"&filter=D&term=Y&etc=B&etc2=0"
        f"&titleTxt=PBR%20Band&dateTxt=undefined&unitTxt="
    )
    basic_band = [0.9, 1.1, 1.3, 1.5, 1.7]

    for attempt in range(max_retries):
        try:
            resp = requests.get(popup_url, timeout=5.0, headers={"User-Agent": "Mozilla/5.0"})
            resp.encoding = 'utf-8'
            soup = BeautifulSoup(resp.text, 'html.parser')

            band_list = []
            # 팝업 페이지 상단 'div.chartData > table > thead > tr > th'에서 5개 밴드 값 추출
            for th in soup.select("div.chartData table thead tr th"):
                try:
                    val = float(th.get_text(strip=True).replace(',', ''))
                    band_list.append(val)
                except:
                    continue

            if len(band_list) == 5:
                return band_list
            else:
                time.sleep(delay)
        except Exception as e:
            time.sleep(delay)
    # 재시도 후에도 실패 시 기본값 사용
    return basic_band

############################################################
# FnGuide 메인 페이지에서 현재 PBR 가져오기 함수
############################################################
def get_current_pbr(gicode):
    try:
        main_url = f"https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode={gicode}"
        resp = requests.get(main_url, timeout=5.0, headers={"User-Agent": "Mozilla/5.0"})
        resp.encoding = 'utf-8'
        soup = BeautifulSoup(resp.text, 'html.parser')

        tag = soup.find(lambda t: t.name == 'th' and ('PBR' in t.get_text() or '주가순자산비율' in t.get_text()))
        if tag:
            td = tag.find_next_sibling('td')
            pbr_text = td.get_text(strip=True).replace(',', '')
            return float(pbr_text)
    except:
        pass
    return None

############################################################
# Position 계산 함수
############################################################
def get_position(pbr, bands):
    if pbr is None or np.isnan(pbr):
        return np.nan
    for i, band in enumerate(bands):
        if pbr < band:
            return i + 1
    return 6

############################################################
# 메인 애플리케이션 로직
############################################################
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # Stochastic 컬럼 자동 탐색
    stochastic_col = next((col for col in df.columns if 'stochastic' in col.lower()), None)
    if not stochastic_col:
        st.error("Stochastic 컬럼을 찾을 수 없습니다.")
        st.stop()

    # ROE 컬럼 자동 탐색 및 3개만 취함
    roe_cols = [col for col in df.columns if 'ROE' in col and '평균' not in col and '최종' not in col]
    if len(roe_cols) < 3:
        st.error(f"ROE 컬럼이 3개 필요합니다. 현재: {roe_cols}")
        st.stop()
    roe_cols = roe_cols[:3]

    # 등락률 퍼센트 포맷 변환
    def percent_format(x):
        try:
            if isinstance(x, str) and '%' in x:
                return x
            else:
                return f"{float(x) * 100:.2f}%"
        except:
            return x

    if '등락률' in df.columns:
        df['등락률'] = df['등락률'].apply(percent_format)

    # 숫자형 컬럼 변환
    num_cols = ['현재가', 'BPS', '배당수익률', stochastic_col, '10년후BPS', '복리수익률'] + roe_cols + ['추정ROE']
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '').str.replace('%', '').astype(float)

    # 계산 컬럼 생성
    if '추정ROE' not in df.columns:
        df['추정ROE'] = df[roe_cols[0]] * 0.4 + df[roe_cols[1]] * 0.35 + df[roe_cols[2]] * 0.25
    if '10년후BPS' not in df.columns:
        df['10년후BPS'] = (df['BPS'] * (1 + df['추정ROE'] / 100) ** 10).round(0)
    if '복리수익률' not in df.columns:
        df['복리수익률'] = (((df['10년후BPS'] / df['현재가']) ** (1 / 10)) - 1) * 100
        df['복리수익률'] = df['복리수익률'].round(2)

    # position 컬럼 생성 : 종목코드별로 FnGuide에서 PBR밴드 5개 + 현재 PBR 값을 크롤링하여 계산
    if '종목코드' in df.columns:
        df['position'] = np.nan
        for idx, row in df.iterrows():
            raw_code = str(row['종목코드']).zfill(6)  # 6자리 맞춤
            gicode = f"A{raw_code}"  # A 접두어 붙임 (주식코드)
            
            bands = get_pbr_band(gicode)
            current_pbr = get_current_pbr(gicode)
            
            # 만약 FnGuide에서 PBR 못 가져오면 계산
            if current_pbr is None and not pd.isna(row['BPS']) and row['BPS'] != 0:
                current_pbr = row['현재가'] / row['BPS']
            
            df.at[idx, 'position'] = get_position(current_pbr, bands)
            
            time.sleep(0.3)  # 크롤링 딜레이 (예의)

    # 가중치 슬라이더 (복리수익률 vs 저평가 분위수)
    alpha = st.slider(
        '복리수익률(성장성) : 저평가(분위수) 가중치 (%)',
        min_value=0, max_value=100, value=80, step=5, format="%d%%"
    ) / 100

    max_return = df['복리수익률'].max()
    min_return = 15.0
    df['복리수익률점수'] = ((df['복리수익률'] - min_return) / (max_return - min_return)).clip(lower=0) * 100

    df['Stochastic_percentile'] = df[stochastic_col].apply(
        lambda x: 100 - percentileofscore(df[stochastic_col], x, kind='mean')
    )

    df['매력도'] = (alpha * df['복리수익률점수'] + (1 - alpha) * df['Stochastic_percentile']).round(2)

    df_sorted = df.sort_values(by='매력도', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index + 1
    df_sorted.rename(columns={stochastic_col: 'RN'}, inplace=True)

    main_cols = ['순위', '종목명', '현재가', '등락률'] + roe_cols + [
        'BPS', '배당수익률', 'RN', '추정ROE', '10년후BPS', '복리수익률', 'position', '매력도'
    ]
    df_show = df_sorted[[col for col in main_cols if col in df_sorted.columns]]

    # 하이라이트 함수 (복리수익률 15% 이상 종목명 연두색)
    def highlight_high_return(row):
        color = 'background-color: lightgreen' if row['복리수익률'] >= 15 else ''
        return [color if col == '종목명' else '' for col in row.index]

    format_dict = {
        '현재가': '{:,.0f}',
        roe_cols[0]: '{:.2f}',
        roe_cols[1]: '{:.2f}',
        roe_cols[2]: '{:.2f}',
        'BPS': '{:,.0f}',
        '배당수익률': '{:.2f}',
        'RN': '{:.0f}',
        '추정ROE': '{:.2f}',
        '10년후BPS': '{:,.0f}',
        '복리수익률': '{:.2f}',
        '매력도': '{:.2f}',
        'position': '{:.0f}'
    }

    styled_df = (
        df_show.style
        .apply(highlight_high_return, axis=1)
        .format(format_dict)
        .set_properties(**{'text-align': 'center'})
        .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])
    )

    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

    # 복리수익률 vs RN
