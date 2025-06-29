import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Dividend Growth Stock", layout="wide")
st.title("📈 Dividend Growth Stock")

file_path = "1.xlsx"

if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()

    # Stochastic 컬럼명 자동 탐색 ("Stochastic Fast %K" 등)
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

    # 숫자형 컬럼 처리
    num_cols = ['현재가', 'BPS', '배당수익률', stochastic_col, '10년후BPS', '복리수익률'] + roe_cols + ['추정ROE']
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '').str.replace('%','').astype(float)

    # 추정ROE, 10년후BPS, 복리수익률 계산 (이미 있으면 생략)
    if '추정ROE' not in df.columns:
        df['추정ROE'] = df[roe_cols[0]]*0.4 + df[roe_cols[1]]*0.35 + df[roe_cols[2]]*0.25
    if '10년후BPS' not in df.columns and 'BPS' in df.columns and '추정ROE' in df.columns:
        df['10년후BPS'] = (df['BPS'] * (1 + df['추정ROE']/100) ** 10).round(0)
    if '복리수익률' not in df.columns and '10년후BPS' in df.columns and '현재가' in df.columns:
        df['복리수익률'] = (((df['10년후BPS'] / df['현재가']) ** (1/10)) - 1) * 100
        df['복리수익률'] = df['복리수익률'].round(2)

    # 매력도 계산 (복리수익률 15% 미만 컷오프, Stochastic 0~100 정규화, 가중합 8:2)
    alpha = 0.8  # 성장성 80%, 저평가 20%
    max_return = df['복리수익률'].max()
    min_return = 15.0  # 고정 컷오프

    # 복리수익률 점수 (15% 미만은 0점)
    df['복리수익률점수'] = ((df['복리수익률'] - min_return) / (max_return - min_return)).clip(lower=0) * 100

    # 저평가 점수: (100 - Stochastic)로 0~100점 (Stochastic 값 0~100 기준)
    df['저평가점수'] = (100 - df[stochastic_col]).clip(lower=0, upper=100)

    # 최종 매력도
    df['매력도'] = (alpha * df['복리수익률점수'] + (1 - alpha) * df['저평가점수']).round(2)

    # 매력도 순 정렬 및 순위(1부터)
    df_sorted = df.sort_values(by='매력도', ascending=False).reset_index(drop=True)
    df_sorted['순위'] = df_sorted.index + 1

    # 표에 표시할 컬럼 순서 (순위가 맨 앞)
    main_cols = ['순위', '종목명', '현재가', '등락률'] + roe_cols + [
        'BPS', '배당수익률', stochastic_col, '추정ROE', '10년후BPS', '복리수익률', '매력도'
    ]
    final_cols = [col for col in main_cols if col in df_sorted.columns]
    df_show = df_sorted[final_cols]

    # 하이라이트 함수 (복리수익률 15% 이상 종목명 연두색)
    def highlight_high_return(row):
        color = 'background-color: lightgreen' if row['복리수익률'] >= 15 else ''
        return [color if col == '종목명' else '' for col in row.index]

    # 포맷 지정
    format_dict = {
        '현재가': '{:,.0f}',
        '등락률': '{:+.2f}%',
        roe_cols[0]: '{:.2f}',
        roe_cols[1]: '{:.2f}',
        roe_cols[2]: '{:.2f}',
        'BPS': '{:,.0f}',
        '배당수익률': '{:.2f}',
        stochastic_col: '{:.0f}',
        '추정ROE': '{:.2f}',
        '10년후BPS': '{:,.0f}',
        '복리수익률': '{:.2f}',
        '매력도': '{:.2f}'
    }

    # 스타일 적용: 가운데 정렬, 하이라이트, 포맷
    styled_df = (
        df_show.style
        .apply(highlight_high_return, axis=1)
        .format(format_dict)
        .set_properties(**{'text-align': 'center'})
        .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])
    )

    # 인덱스 숨김 옵션 추가!
    st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)

else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.\n\n해당 파일을 같은 폴더에 넣어주세요.")
