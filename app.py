import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 페이지 설정 (세련된 느낌을 위해)
st.set_page_config(page_title="Yuyu Pharma HR Dashboard", layout="wide")
st.title("📊 전사 인원현황 대시보드")
st.markdown("---")

# 2. 데이터 불러오기 (보내주신 파일명 기준)
try:
    # CSV 파일을 읽습니다 (엑셀 시리얼 날짜 대응)
    df = pd.read_excel("test10.xlsx")

    # 엑셀 날짜 숫자(45301 등)를 실제 날짜로 바꾸는 마법
    def convert_date(date_val):
        if pd.isna(date_val) or date_val == "": return None
        try:
            return pd.to_datetime(date_val, unit='D', origin='1899-12-30').date()
        except:
            return date_val

    df['입사일'] = df['입사일'].apply(convert_date)
    df['퇴사일'] = df['퇴사일'].apply(convert_date)

    # 재직 상태 구분
    df['상태'] = df['퇴사일'].apply(lambda x: '퇴사' if pd.notna(x) and x != "" else '재직')
    active_df = df[df['상태'] == '재직']

    # --- 사이드바 (필터) ---
    st.sidebar.header("🔍 검색 필터")
    selected_dept = st.sidebar.multiselect("부서 선택", options=df['부서'].unique(), default=df['부서'].unique())
    selected_type = st.sidebar.multiselect("근무 구분", options=df['구분'].unique(), default=df['구분'].unique())

    # 필터 적용된 데이터
    filtered_df = active_df[(active_df['부서'].isin(selected_dept)) & (active_df['구분'].isin(selected_type))]

    # --- 상단 지표 (KPI) ---
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("현재 재직자", f"{len(filtered_df)}명")
    with m2:
        st.metric("남성 인원", f"{len(filtered_df[filtered_df['성별']=='남'])}명")
    with m3:
        st.metric("여성 인원", f"{len(filtered_df[filtered_df['성별']=='여'])}명")
    with m4:
        st.metric("퇴사자 (누적)", f"{len(df[df['상태']=='퇴사'])}명")

    st.divider()

    # --- 차트 영역 ---
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("🏢 부서별 인원 비중")
        fig_dept = px.pie(filtered_df, names='부서', hole=0.5, color_discrete_sequence=px.colors.sequential.RdBu)
        st.plotly_chart(fig_dept, use_container_width=True)

    with c2:
        st.subheader("🎖️ 직책별 분포")
        fig_rank = px.bar(filtered_df['직책'].value_counts().reset_index(), x='index', y='직책', 
                          labels={'index':'직책', '직책':'인원(명)'}, color='index')
        st.plotly_chart(fig_rank, use_container_width=True)

    # --- 데이터 상세 테이블 ---
    st.subheader("📋 상세 사원 명부")
    st.dataframe(filtered_df[['사원명', '구분', '부서', '직책', '성별', '입사일']], use_container_width=True)

except Exception as e:
    st.error(f"오류가 발생했어요: {e}")
    st.info("폴더 안에 'test10.xlsx - Sheet1.csv' 파일이 있는지 확인해 주세요!")
