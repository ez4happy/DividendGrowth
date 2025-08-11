import streamlit as st
import pandas as pd
import requests
import re
import json
import traceback

st.set_page_config(layout="wide")
st.title("📈 Dividend Growth Stock with POSITION")

# -------------------
# POSITION 계산 함수
# -------------------
def get_position(code):
    """FN가이드 PBR Band 데이터 기반 POSITION 계산"""
    url = f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp?oid=pbrBandCht&cid=01_06&gicode={code}&filter=D&term=Y&etc=B&etc2=0&titleTxt=PBR%20Band&dateTxt=undefined&unitTxt="
    try:
        r = requests.get(url, timeout=5)
        r.encoding = 'utf-8'

        # chartData 추출
        match = re.search(r'var\s+chartData\s*=\s*(\[[^\]]+\])', r.text)
        if not match:
            st.warning(f"{code}: chartData 패턴 없음")
            return None

        data_str = match.group(1)
        data = json.loads(data_str)

        # 수정주가 리스트 추출
        prices = [item[1] for item in data if isinstance(item, list) and len(item) > 1]
        if len(prices) < 6:
            st.warning(f"{code}: price 데이터 부족 ({len(prices)}개)")
            return None

        current_price = prices[0]
        next_5 = prices[1:6]
        next_4 = prices[1:5]

        if all(current_price < p for p in next_5):
            return 1
        elif all(current_price < p for p in next_4):
            return 2
        else:
            return 6
    except Exception as e:
        st.warning(f"{code} 에러: {e}")
        return None

# -------------------
# 기존 매력도 계산 함수
# -------------------
def calculate_attractiveness(row):
    """
    기존 매력도 계산 로직을 이 함수 안에 넣으시면 됩니다.
    아래 예시는 간단한 점수 예시입니다.
    """
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

        # POSITION 계산
        st.info("POSITION 계산 중입니다. 잠시만 기다려 주세요...")
        df['POSITION'] = df['종목코드_A'].apply(get_position)

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
