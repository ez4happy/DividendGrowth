import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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

    # 컬럼 순서
    main_cols = ['순위', '종목명', '현재가', '등락률'] + roe_cols + ['추정ROE', 'BPS', '배당수익률', 'Stochastic', '10년후BPS', '복리수익률(%)']
    final_cols = [col for col in main_cols if col in df_sorted.columns]
    df_show = df_sorted[final_cols]

    # 대시보드 요약 정보
    col1, col2, col3 = st.columns(3)
    col1.metric("전체 종목 평균 추정ROE(%)", round(df_show['추정ROE'].mean(), 2))
    col2.metric("전체 종목 평균 복리수익률(%)", round(df_show['복리수익률(%)'].mean(), 2))
    col3.metric("종목 수", len(df_show))

    st.markdown("### 📋 종목별 데이터")
    st.dataframe(df_show, use_container_width=True, height=500)

    st.markdown("### 📊 복리수익률(%) 순위별 바 차트")
    fig1 = px.bar(
        df_show,
        x='순위',
        y='복리수익률(%)',
        hover_data=final_cols,
        labels={'순위': '순위', '복리수익률(%)': '복리수익률 (%)'},
        title='복리수익률 순위별 바 차트'
    )
    fig1.update_layout(xaxis_title='순위', yaxis_title='복리수익률 (%)', plot_bgcolor='#fcfcfc', font=dict(size=15))
    st.plotly_chart(fig1, use_container_width=True)

    st.markdown("### 📊 추정ROE(%) 순위별 바 차트")
    fig2 = px.bar(
        df_show,
        x='순위',
        y='추정ROE',
        hover_data=final_cols,
        labels={'순위': '순위', '추정ROE': '추정ROE (%)'},
        title='추정ROE 순위별 바 차트'
    )
    fig2.update_layout(xaxis_title='순위', yaxis_title='추정ROE (%)', plot_bgcolor='#fcfcfc', font=dict(size=15))
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### 📈 복리수익률(%) & 추정ROE(%) 혼합 그래프")
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(
        x=df_show['순위'],
        y=df_show['복리수익률(%)'],
        name='복리수익률(%)',
        marker_color='skyblue'
    ))
    fig3.add_trace(go.Scatter(
        x=df_show['순위'],
        y=df_show['추정ROE'],
        name='추정ROE(%)',
        yaxis='y2',
        mode='lines+markers',
        marker_color='orange'
    ))
    fig3.update_layout(
        title='복리수익률 & 추정ROE 순위별 혼합 차트',
        xaxis_title='순위',
        yaxis=dict(title='복리수익률(%)'),
        yaxis2=dict(title='추정ROE(%)', overlaying='y', side='right'),
        plot_bgcolor='#fcfcfc',
        font=dict(size=15),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
    )
    st.plotly_chart(fig3, use_container_width=True)

else:
    st.error(f"현재 작업 폴더에 '{file_path}' 파일이 없습니다.\n\n해당 파일을 같은 폴더에 넣어주세요.")
