import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import traceback

st.set_page_config(layout="wide")
st.title("📈 Dividend Growth Stock with POSITION (HTML 테이블 파싱 버전)")

# -------------------
# POSITION 계산 함수 (HTML 테이블 파싱)
# -------------------
def get_position_from_html_table(code):
    url = f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp?oid=pbrBandCht&cid=01_06&gicode={code}&filter=D&term=Y&etc=B&etc2=0&titleTxt=PBR%20Band&dateTxt=undefined&unitTxt="
    try:
        r = requests.get(url, timeout=10)
        r.encoding = 'utf-8'

        soup = BeautifulSoup(r.text, 'html.parser')
        table = soup.find('table')
        if not table:
            st.warning(f"{code}: 데이터 테이블을 찾을 수 없습니다.")
            return None

        rows = table.find_all('tr')
        if len(rows) < 3:
            st.warning(f"{code}: 테이블 데이터 행이 부족합니다.")
            return None

        # 데이터가 두 번째 행부터 시작하는 경우가 많음 (헤더 1행, 실제 데이터 2행~)
        # 최신 데이터는 보통 가장 위에 있으므로 2번째 행(인덱스 1) 사용
        data_row = rows[1]
        cols = data_row.find_all('td')
        if len(cols) < 7:
            st.warning(f"{code}: 데이터 컬럼이 충분하지 않습니다.")
            return None

        # 수정주가 (2번째 컬럼)
        price_str = cols[1].get_text().replace(',', '').strip()
        price = float(price_str)

        # 밴드 5개 (3~7번째 컬럼)
        bands = []
        for i in range(2, 7):
            band_str = cols[i].get_text().replace(',', '').strip()
            bands.append(float(band_str))

        # POSITION 계산
        if all(price < b for b in bands):
            return 1
        elif all(price < b for b in bands[:-1]):
            return 2
        else:
            return 6

    except Exception as e:
        st.warning(f"{code} 파싱 중 에러 발생: {e}")
        return None


# -------------------
# 기존 매력도 계산 함수 (예시)
# -------------------
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


# -------------------
# 메인 실행부
# -------------------
file_path = "1.xlsx"

if st.button("데이터 불러오고 계산 시작"):
    try:
        df = pd.read_excel(file_path)
        st.success(f"{file_path} 파일 로드 완료")

        # 종목코드 6자리 + A 접두사
        if '종목코드' not in df.columns:
            st.error("엑셀에 '종목코드' 컬럼이 없습니다.")
            st.stop()
        df['종목코드'] = df['종목코드'].astype(str).str.zfill(6)
        df['종목코드_A'] = 'A' + df['종목코드']

        # POSITION 계산 (HTML 테이블 파싱)
        st.info("POSITION 계산 중입니다. 시간이 걸릴 수 있습니다...")
        df['POSITION'] = df['종목코드_A'].apply(get_position_from_html_table)

        # 매력도 계산
        st.info("매력도 계산 중입니다...")
        df['매력도'] = df.apply(calculate_attractiveness, axis=1)

        # 매력도 앞에 POSITION 컬럼 배치
        if '매력도' in df.columns and 'POSITION' in df.columns:
            cols = df.columns.tolist()
            cols.remove('POSITION')
            cols.remove('매력도')
            insert_idx = cols.index('종목코드') + 1
            new_order = cols[:insert_idx] + ['POSITION', '매력도'] + cols[insert_idx:]
            df = df[new_order]

        # 결과 출력
        st.dataframe(df)

        # 엑셀 다운로드
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
