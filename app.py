import streamlit as st
import pandas as pd
import os

# KeyError 방지: styler.render.max_elements 옵션 크게 설정
pd.set_option("styler.render.max_elements", 999_999_999)

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # 복리수익률 컬럼명 자동 탐색
    rate_col = '복리수익률'
    if rate_col not in df.columns:
        rate_col = '복리수익률(%)'

    # 보기 좋은 숫자/퍼센트 포맷
    format_dict = {
        '현재가': '{:,.0f}',
        '등락률': '{:+.2f}%',
        'ROE': '{:.2f}',
        'BPS': '{:,.0f}',
        '배당수익률': '{:.2f}',
        'Stochastic': '{:.0f}',
        '추정ROE': '{:.2f}',
        '10년후BPS': '{:,.0f}',
        rate_col: '{:.2f}'
    }
    format_dict = {col: fmt for col, fmt in format_dict.items() if col in df.columns}

    # 복리수익률 15% 이상 셀만 하이라이트
    def highlight_return(val):
        try:
            return 'background-color: lightgreen' if float(val) >= 15 else ''
        except:
            return ''

    styled_df = (
        df.style
        .format(format_dict)
        .applymap(highlight_return, subset=[rate_col])
        .set_properties(**{'text-align': 'center'})
        .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])
    )

    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.\n\n해당 파일을 같은 폴더에 넣어주세요.")
