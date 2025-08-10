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

def get_pbr_band_and_pbr(gicode):
    url = f"https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode={gicode}&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701"
    try:
        resp = requests.get(url, timeout=5.0)
        soup = BeautifulSoup(resp.text, 'html.parser')
        # 예시 selector: 실제 구조 'PBR Band' 테이블 확인 필요
        band_list = []
        # 예시: table 내 PBR 밴드 값을 row(행)에서 추출
        pbr_band_table = soup.find("table", string=lambda x: x and "PBR" in x)
        if pbr_band_table:
            for cell in pbr_band_table.find_all("td"):
                try:
                    band = float(cell.get_text(strip=True))
                    band_list.append(band)
                except:
                    continue
        # (실제로는 위에서 .find 또는 select로 band_list 파싱 필요! 아래는 예시)
        if len(band_list) < 5:
            band_list = [0.9, 1.1, 1.3, 1.5, 1.7] # 예시 기본값
        # 현재 PBR 추출(대개의 경우 "주요 투자지표" 섹션에 별도 표기)
        current_pbr = None
        keylabels = ["PBR", "주가순자산비율"]
        for label in keylabels:
            tag = soup.find(lambda tag: tag.name == 'th' and label in tag.get_text())
            if tag:
                td = tag.find_next('td')
                try:
                    current_pbr = float(td.get_text(strip=True))
                    break
                except:
                    continue
        return band_list, current_pbr
    except Exception as e:
        st.warning(f"{gicode} 크롤링 오류: {e}")
        return [0.9, 1.1, 1.3, 1.5, 1.7], None

def get_position(pbr, bands):
    if pd.isna(pbr): return np.nan
    for i, band in enumerate(bands):
        if pbr < band:
            return i+1
    return 6

if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # Stochastic 컬럼명 자동 탐색
    stochastic_col = None
    for col in df.columns:
        if 'stochastic' in col.lower():
            stochastic_col = col
            break
    if not stochastic_col:
        st.error("Stochastic 컬럼을 찾을 수 없습니다.")
        st.stop()

    # ROE 컬럼 자동 탐색
    roe_cols = [col for col in df.columns if 'ROE' in col and '평균' not in col and '최종' not in col]
    if len(roe_cols) < 3:
        st.error(f"ROE 컬럼이 3개 필요합니다. 현재: {roe_cols}")
        st.stop()
    roe_cols = roe_cols[:3]

    def percent_format(x):
        try:
            if isinstance(x, str) and '%' in x:
                return x
            else:
                return f"{float(x)*100:.2f}%"
        except:
            return x

    if '등락률' in df.columns:
        df['등락률'] = df['등락률'].apply(percent_format)

    num_cols = ['현재가', 'BPS', '배당수익률', stochastic_col, '10년후BPS', '복리수익률'] + roe_cols + ['추정ROE']
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '').str.replace('%','').astype(float)

    if '추정ROE' not in df.columns:
        df['추정ROE'] = df[roe_cols[0]]*0.4 + df[roe_cols[1]]*0.35 + df[roe_cols[2]]*0.25
    if '10년후BPS' not in df.columns and 'BPS' in df.columns and '추정ROE' in df.columns:
        df['10년후BPS'] = (df['BPS'] * (1 + df['추정ROE']/100) ** 10).round(0)
    if '복리수익률' not in df.columns and '10년후BPS' in df.columns and '현재가' in df.columns:
        df['복리수익률'] = (((df['10년후BPS'] / df['현재가']) ** (1/10)) - 1) * 100
        df['복리수익률'] = df['복리수익률'].round(2)

    # [핵심]: position 컬럼을 FnGuide에서 PBR밴드 실시간 파싱 후 넣기
    if '종목코드' in df.columns:
        df['position'] = np.nan
        for idx, row in df.iterrows():
            gicode = row['종목코드']
            if pd.isna(gicode): continue
            try:
                band_list, current_pbr = get_pbr_band_and_pbr(gicode)
                # bps 없는 경우 PBR 산출
                if current_pbr is None and not pd.isna(row['BPS']) and row['BPS']!=0:
                    current_pbr = row['현재가']/row['BPS']
                df.at[idx, 'position'] = get_position(current_pbr, band_list)
                time.sleep(0.4) # 접근 속도 제한
            except Exception as e:
                st.warning(f"{gicode} 파싱실패: {e}")

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
    df_sorted = df_sorted.rename(columns={stochastic_col: 'RN'})

    main_cols = ['순위', '종목명', '현재가', '등락률'] + roe_cols + [
        'BPS', '배당수익률', 'RN', '추정ROE', '10년후BPS', '복리수익률', 'position', '매력도'
    ]
    final_cols = [col for col in main_cols if col in df_sorted.columns]
    df_show = df_sorted[final_cols]

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

    fig_scatter = px.scatter(
        df_sorted,
        x='RN',
        y='복리수익률',
        color='매력도',
        hover_name='종목명',
        title='복리수익률 vs RN 산점도',
        labels={'RN': 'RN (Stochastic %K)', '복리수익률': '복리수익률 (%)'},
        color_continuous_scale='Viridis'
    )
    st.plotly_chart(fig_scatter, use_container_width=True)

    top_df = df_sorted.head(5)
    fig_bar = px.bar(
        top_df,
        x='종목명',
        y='매력도',
        title='매력도 상위 5개 종목',
        labels={'종목명': '종목명', '매력도': '매력도 점수'},
        color='매력도',
        color_continuous_scale='Viridis'
    )
    st.plotly_chart(fig_bar, use_container_width=True)

else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.\n\n해당 파일을 같은 폴더에 넣어주세요.")
