import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import traceback

st.set_page_config(layout="wide")
st.title("📈 Dividend Growth Stock with Monthly Data & POSITION")

def fetch_monthly_data(code):
    url = f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp?oid=pbrBandCht&cid=01_06&gicode={code}&filter=D&term=Y&etc=B&etc2=0&titleTxt=PBR%20Band&dateTxt=undefined&unitTxt="
    try:
        headers = {
            "User-Agent": "Mozilla/5.0"
        }
        r = requests.get(url, headers=headers, timeout=10)
        r.encoding = 'utf-8'
        soup = BeautifulSoup(r.text, 'html.parser')

        # 테이블 전체 찾기
        tables = soup.find_all('table')
        if not tables:
            st.warning(f"{code}: 테이블을 찾을 수 없습니다.")
            return None

        # 테이블 중 유효한 데이터가 있는 테이블 찾기
        for table in tables:
            rows = table.find_all('tr')
            if len(rows) < 3:
                continue
            # 헤더인지 확인 (1번째 row)
            header_cols = rows[0].find_all(['th','td'])
            header_texts = [col.get_text().strip() for col in header_cols]
            # 기대하는 헤더 있는지 확인 (일자, 수정주가 등)
            if not ('일자' in header_texts and '수정주가' in header_texts):
                continue

            # 데이터 수집 시작
            data = []
            for row in rows[1:]:
                cols = row.find_all('td')
                if len(cols) < 7:
                    continue
                date_str = cols[0].get_text().strip()
                price_str = cols[1].get_text().replace(',', '').strip()
                bands = [cols[i].get_text().replace(',', '').strip() for i in range(2,7)]

                try:
                    price = float(price_str)
                    bands_f = [float(b) for b in bands]
                    data.append({
                        '종목코드': code,
                        '일자': date_str,
                        '수정주가': price,
                        '밴드1': bands_f[0],
                        '밴드2': bands_f[1],
                        '밴드3': bands_f[2],
                        '밴드4': bands_f[3],
                        '밴드5': bands_f[4],
                    })
                except:
                    continue

            if data:
                return pd.DataFrame(data)

        st.warning(f"{code}: 유효한 데이터 테이블을 찾지 못했습니다.")
        return None

    except Exception as e:
        st.warning(f"{code} 데이터 수집 중 에러: {e}")
        return None

def calc_position(row):
    price = row['수정주가']
    bands = [row[f'밴드{i}'] for i in range(1,6)]

    if all(price < b for b in bands):
        return 1
    elif all(price < b for b in bands[:-1]):
        return 2
    else:
        return 6

file_path = "1.xlsx"

if st.button("데이터 불러오고 계산 시작"):
    try:
        df_base = pd.read_excel(file_path)
        st.success(f"{file_path} 파일 로드 완료")

        if '종목코드' not in df_base.columns:
            st.error("'종목코드' 컬럼이 없습니다.")
            st.stop()

        df_base['종목코드'] = df_base['종목코드'].astype(str).str.zfill(6)
        df_base['종목코드_A'] = 'A' + df_base['종목코드']

        all_data = []
        for code in df_base['종목코드_A'].unique():
            st.info(f"{code} 데이터 수집중...")
            monthly_df = fetch_monthly_data(code)
            if monthly_df is not None:
                all_data.append(monthly_df)

        if not all_data:
            st.error("어떤 종목도 데이터를 수집하지 못했습니다.")
            st.stop()

        df_monthly = pd.concat(all_data, ignore_index=True)

        # 최신 월 데이터만 남기기 (일자 기준 내림차순)
        df_monthly['일자'] = pd.to_datetime(df_monthly['일자'], format='%Y/%m/%d')
        df_latest = df_monthly.sort_values('일자', ascending=False).groupby('종목코드').first().reset_index()

        # POSITION 계산
        df_latest['POSITION'] = df_latest.apply(calc_position, axis=1)

        # 원본 데이터와 합치기
        df_final = pd.merge(df_base, df_latest[['종목코드', 'POSITION']], how='left', left_on='종목코드_A', right_on='종목코드')
        df_final.drop(columns=['종목코드_y'], inplace=True)
        df_final.rename(columns={'종목코드_x':'종목코드'}, inplace=True)

        st.dataframe(df_final)

        @st.cache_data
        def to_excel(df):
            return df.to_excel(index=False)

        st.download_button("💾 엑셀 다운로드", to_excel(df_final), file_name="1_with_position.xlsx")

    except Exception as e:
        st.error(f"에러 발생: {e}")
        st.text(traceback.format_exc())

else:
    st.info(f"'{file_path}' 파일이 있어야 실행 가능합니다.")
