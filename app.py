import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import traceback

st.set_page_config(layout="wide")
st.title("📈 Dividend Growth Stock with POSITION (um_table 테이블 파싱)")

def get_position_from_html_table(code):
    url = f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp?oid=pbrBandCht&cid=01_06&gicode={code}&filter=D&term=Y&etc=B&etc2=0&titleTxt=PBR%20Band&dateTxt=undefined&unitTxt="
    try:
        r = requests.get(url, timeout=10)
        r.encoding = 'utf-8'

        soup = BeautifulSoup(r.text, 'html.parser')
        tables = soup.find_all('table', class_='um_table')
        if not tables:
            st.warning(f"{code}: 'um_table' 클래스 테이블을 찾을 수 없습니다.")
            return None

        table = tables[0]
        rows = table.find_all('tr')
        if len(rows) < 3:
            st.warning(f"{code}: 테이블 데이터 행이 부족합니다.")
            return None

        data_row = rows[1]
        cols = data_row.find_all('td')
        if len(cols) < 7:
            st.warning(f"{code}: 데이터 컬럼이 충분하지 않습니다.")
            return None

        price_str = cols[1].get_text().replace(',', '').strip()
        price = float(price_str)

        bands = []
        for i in range(2, 7):
            band_str = cols[i].get_text().replace(',', '').strip()
            bands.append(float(band_str))

        if all(price < b for b in bands):
            return 1
        elif all(price < b for b in bands[:-1]):
            return 2
        else:
            return 6

    except Exception as e:
        st.warning(f"{code} 파싱 중 에러 발생: {e}")
        return None

def calculate_attractiveness(row):
    try:
        score = 0
        if pd.notnull(row.get('PER')) and row['PER'] < 10:
            score += 2
        if pd.notnull(row.get('PBR')) and row['PBR'] < 1:
            score += 2
        if pd.notnull(row.get('ROE')) and row['ROE'] > 10:
            score += 2
        return score
    except:
        return None

file_path = "1.xlsx"

if st.button("데이터 불러오고 계산 시작"):
    try:
        df = pd.read_excel(file_path)
        st.success(f"{file_path} 파일 로드 완료")

        if '종목코드' not in df.columns:
            st.error("엑셀에 '종목코드' 컬럼이 없습니다.")
            st.stop()

        df['종목코드'] = df['종목코드'].astype(str).str.zfill(6)
        df['종목코드_A'] = 'A' + df['종목코드']

        st.info("POSITION 계산 중입니다. 시간이 걸릴 수 있습니다...")
        df['POSITION'] = df['종목코드_A'].apply(get_position_from_html_table)

        st.info("매력도 계산 중입니다...")
        df['매력도'] = df.apply(calculate_attractiveness, axis=1)

        if '매력도' in df.columns and 'POSITION' in df.columns:
            cols = df.columns.tolist()
            cols.remove('POSITION')
            cols.remove('매력도')
            insert_idx = cols.index('종목코드') + 1
            new_order = cols[:insert_idx] + ['POSITION', '매력도'] + cols[insert_idx:]
            df = df[new_order]

        st.dataframe(df)

        @st.cache_data
        def convert_to_excel(df):
            return df.to_excel(index=False)

        st.download_button(
            label="💾 엑셀 다운로드",
            data=convert_to_excel(df),
            file_name="1_with_position.xlsx"
        )

    except Exception as e:
        st.error(f"에러 발생: {e}")
        st.text(traceback.format_exc())

else:
    st.info(f"현재 작업 폴더에 '{file_path}' 파일이 있어야 실행 가능합니다.")
