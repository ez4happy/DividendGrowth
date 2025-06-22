import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # ROE 컬럼 자동 탐색 (첨부 이미지 기준 3개)
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

    # 표에 표시할 컬럼 순서 (첨부 이미지 기준)
    main_cols = [
        '종목명', '현재가', '등락률',
        roe_cols[0], roe_cols[1], roe_cols[2],
        'BPS', '배당수익률', 'Stochastic',
        '추정ROE', '10년후BPS', '복리수익률(%)'
    ]
    final_cols = [col for col in main_cols if col in df.columns]
    df_show = df[final_cols]

    # 하이라이트 함수: 복리수익률(%) 15 이상이면 연두색
    def highlight_high_return(row):
        color = 'background-color: lightgreen' if row['복리수익률(%)'] >= 15 else ''
        return [color if col == '복리수익률(%)' else '' for col in row.index]

    # 포맷 지정 (이미지 참고)
    format_dict = {
        '현재가': '{:,.0f}',
        '등락률': '{:+.2f}%',
        roe_cols[0]: '{:.2f}',
        roe_cols[1]: '{:.2f}',
        roe_cols[2]: '{:.2f}',
        'BPS': '{:,.0f}',
        '배당수익률': '{:.2f}',
        'Stochastic': '{:.0f}',
        '추정ROE': '{:.2f}',
        '10년후BPS': '{:,.0f}',
        '복리수익률(%)': '{:.2f}'
    }

    # 가운데 정렬, 포맷, 하이라이트 적용
    styled_df = (
        df_show.style
        .apply(highlight_high_return, axis=1)
        .format(format_dict)
        .set_properties(**{'text-align': 'center'})
        .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])
    )

    # 인덱스 숨김 (Streamlit 1.23+)
    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.\n\n해당 파일을 같은 폴더에 넣어주세요.")
