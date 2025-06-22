import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # (생략) 데이터 전처리 및 계산...

    # 예시: 복리수익률 컬럼명은 상황에 따라 '복리수익률' 또는 '복리수익률(%)'
    rate_col = '복리수익률'
    if rate_col not in df.columns:
        rate_col = '복리수익률(%)'

    # 포맷 지정
    format_dict = {
        rate_col: '{:.2f}'
    }

    # 복리수익률 컬럼만 15 이상 하이라이트
    def highlight_return(val):
        color = 'background-color: lightgreen' if val >= 15 else ''
        return color

    styled_df = (
        df.style
        .format(format_dict)
        .applymap(highlight_return, subset=[rate_col])  # 복리수익률 컬럼만 적용
        .set_properties(**{'text-align': 'center'})
        .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])
    )

    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.\n\n해당 파일을 같은 폴더에 넣어주세요.")
