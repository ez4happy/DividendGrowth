import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # ROE 컬럼 자동 탐색
    roe_cols = [col for col in df.columns if 'ROE' in col and '평균' not in col and '최종' not in col]
    if len(roe_cols) < 3:
        st.error(f"ROE 컬럼이 3개 필요합니다. 현재: {roe_cols}")
        st.stop()
    roe_cols = roe_cols[:3]

    # 숫자형 컬럼 처리
    num_cols = ['현재가', 'BPS', '배당수익률', 'Stochastic', '10년후BPS', '복리수익률'] + roe_cols
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '').str.replace('%','').astype(float)

    # 추정ROE, 10년후BPS, 복리수익률 계산
    df['추정ROE'] = df[roe_cols[0]]*0.4 + df[roe_cols[1]]*0.35 + df[roe_cols[2]]*0.25
    df['10년후BPS'] = (df['BPS'] * (1 + df['추정ROE']/100) ** 10).round(0)
    df['복리수익률(%)'] = (((df['10년후BPS'] / df['현재가']) ** (1/10)) - 1) * 100
    df['복리수익률(%)'] = df['복리수익률(%)'].round(2)

    # 정렬 및 순위
    df_sorted = df.sort_values(by='복리수익률(%)', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index + 1

    # 표에 표시할 컬럼만 선택
    main_cols = ['순위', '종목명', '현재가', '등락률'] + roe_cols + ['추정ROE', 'BPS', '배당수익률', 'Stochastic', '10년후BPS', '복리수익률(%)']
    final_cols = [col for col in main_cols if col in df_sorted.columns]
    df_show = df_sorted[final_cols]

    # 복리수익률 15% 이상 종목명 녹색 표시 함수
    def highlight_high_return(row):
        color = 'background-color: lightgreen' if row['복리수익률(%)'] >= 15 else ''
        return [color if col == '종목명' else '' for col in row.index]

    # 스타일 적용
    st.dataframe(
        df_show.style.apply(highlight_high_return, axis=1),
        use_container_width=True,
        height=500
    )

else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.\n\n해당 파일을 같은 폴더에 넣어주세요.")
