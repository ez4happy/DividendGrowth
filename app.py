from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import pandas as pd
import streamlit as st
import time

def fetch_table_with_selenium(code):
    url = f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp?oid=pbrBandCht&cid=01_06&gicode={code}&filter=D&term=Y&etc=B&etc2=0&titleTxt=PBR%20Band&dateTxt=undefined&unitTxt="
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # 창 안 띄우기
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    driver.get(url)
    time.sleep(3)  # 페이지 로딩 대기

    # 테이블 요소 찾기
    table = driver.find_element(By.CSS_SELECTOR, "table.um_table")

    rows = table.find_elements(By.TAG_NAME, "tr")

    data = []
    for row in rows[1:]:  # 헤더 제외
        cols = row.find_elements(By.TAG_NAME, "td")
        if len(cols) < 7:
            continue
        date = cols[0].text.strip()
        price = cols[1].text.strip().replace(',', '')
        bands = [cols[i].text.strip().replace(',', '') for i in range(2,7)]
        try:
            price_f = float(price)
            bands_f = [float(b) for b in bands]
            data.append({
                "일자": date,
                "수정주가": price_f,
                "밴드1": bands_f[0],
                "밴드2": bands_f[1],
                "밴드3": bands_f[2],
                "밴드4": bands_f[3],
                "밴드5": bands_f[4],
            })
        except:
            continue

    driver.quit()

    df = pd.DataFrame(data)
    return df

def calc_position(row):
    price = row['수정주가']
    bands = [row[f'밴드{i}'] for i in range(1,6)]
    if all(price < b for b in bands):
        return 1
    elif all(price < b for b in bands[:-1]):
        return 2
    else:
        return 6


# Streamlit 앱 예제
st.title("Dividend Growth Stock with Selenium POSITION")

file_path = "1.xlsx"

if st.button("데이터 불러오고 계산 시작"):
    import traceback
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
            df_monthly = fetch_table_with_selenium(code)
            if df_monthly is not None and not df_monthly.empty:
                df_monthly['종목코드'] = code
                all_data.append(df_monthly)

        if not all_data:
            st.error("데이터를 수집하지 못했습니다.")
            st.stop()

        df_monthly_all = pd.concat(all_data, ignore_index=True)
        df_monthly_all['일자'] = pd.to_datetime(df_monthly_all['일자'], format="%Y/%m/%d")

        df_latest = df_monthly_all.sort_values('일자', ascending=False).groupby('종목코드').first().reset_index()
        df_latest['POSITION'] = df_latest.apply(calc_position, axis=1)

        df_final = pd.merge(df_base, df_latest[['종목코드', 'POSITION']], left_on='종목코드_A', right_on='종목코드', how='left')
        df_final.drop(columns=['종목코드_y'], inplace=True)
        df_final.rename(columns={'종목코드_x': '종목코드'}, inplace=True)

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
