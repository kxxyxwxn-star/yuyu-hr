import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# 1. 페이지 설정: A4 보고서 느낌을 위해 가로폭을 제한(centered)
st.set_page_config(page_title="Yuyu Pharma 인원현황 보고서", layout="centered")

# CSS 스타일: 필터 색상 파란색 변경 및 디자인 미세 조정
st.markdown("""
    <style>
    /* 지표(Metric) 글자 크기 및 색상 */
    [data-testid="stMetricValue"] { font-size: 28px; color: #004a99; font-weight: bold; }
    [data-testid="stMetricLabel"] { font-size: 14px; color: #666; }
    
    /* 멀티셀렉트 필터 색상 (파란색 계열) */
    .stMultiSelect div div div div { background-color: #e3f2fd; color: #0d47a1; border-radius: 5px; }
    
    /* 구분선 및 제목 간격 조절 */
    hr { margin-top: 1rem; margin-bottom: 1rem; }
    
    /* 출력 시 버튼 등 불필요한 요소 제거 */
    @media print {
        .no-print, .stButton, [data-testid="stExpander"] { display: none !important; }
        .main { padding: 0 !important; }
    }
    </style>
    """, unsafe_allow_html=True)

# 2. 데이터 불러오기 및 전처리
try:
    # 엑셀 파일 읽기
    df = pd.read_excel("test10.xlsx")

    # 엑셀 날짜 변환 함수 (숫자형 날짜 대응)
    def convert_date(date_val):
        if pd.isna(date_val) or date_val == "": return None
        try:
            return pd.to_datetime(date_val, unit='D', origin='1899-12-30')
        except:
            return pd.to_datetime(date_val)

    df['입사일'] = df['입as일'].apply(convert_date) if '입as일' in df.columns else df['입사일'].apply(convert_date)
    df['퇴사일'] = df['퇴사일'].apply(convert_date)

    # --- [사이드바: 기준월 설정 및 필터] ---
    st.sidebar.header("📅 보고서 기준 설정")
    # 기준 연도와 월 선택 (이 선택에 따라 '당월'이 결정됨)
    report_year = st.sidebar.selectbox("기준 연도", options=[2024, 2025, 2026], index=2)
    report_month = st.sidebar.slider("기준 월", 1, 12, datetime.now().month)
    
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 상세 필터")
    selected_dept = st.sidebar.multiselect("부서 선택", options=df['부서'].unique(), default=df['부서'].unique())
    selected_type = st.sidebar.multiselect("근무구분 선택", options=df['구분'].unique(), default=df['구분'].unique())

    # 데이터 필터링 로직
    # 1. 당월 입사자: 선택한 연/월에 입사한 사람
    monthly_in = df[(df['입사일'].dt.year == report_year) & (df['입사일'].dt.month == report_month)]
    
    # 2. 당월 퇴사자: 선택한 연/월에 퇴사한 사람
    monthly_out = df[(df['퇴사일'].dt.year == report_year) & (df['퇴사일'].dt.month == report_month)]
    
    # 3. 현재 재직자: 기준일(선택월 말일) 기준으로 입사일은 기준일 이전이고, 퇴사일은 없거나 기준일 이후인 사람
    # (단순화를 위해 현재 파일 내 '퇴사일'이 없는 사람 중 필터링된 인원으로 표시)
    active_df = df[df['퇴사일'].isna() & (df['부서'].isin(selected_dept)) & (df['구분'].isin(selected_type))]

    # --- [메인 대시보드 시작] ---
    st.title("📑 인원현황 보고서")
    st.subheader(f"{report_year}년 {report_month}월 기준")
    st.caption(f"발행일: {datetime.now().strftime('%Y-%m-%d')}")
    st.markdown("---")

    # [섹션 1: 핵심 요약 지표]
    col1, col2, col3 = st.columns(3)
    col1.metric("현재 재직자", f"{len(active_df)}명")
    col2.metric("당월 입사자", f"{len(monthly_in)}명")
    col3.metric("당월 퇴사자", f"{len(monthly_out)}명")
    
    st.markdown("---")

    # [섹션 2: 근무구분 및 직책별 현황]
    st.markdown("#### 📊 인력 구성 현황")
    c1, c2 = st.columns(2)
    with c1:
        fig_type = px.pie(active_df, names='구분', title="근무구분별 비중", hole=0.4, height=300)
        fig_type.update_layout(showlegend=True, margin=dict(t=30, b=0, l=0, r=0))
        st.plotly_chart(fig_type, use_container_width=True)
    with c2:
        rank_data = active_df['직책'].value_counts().reset_index()
        rank_data.columns = ['직책', '인원수']
        fig_rank = px.bar(rank_data, x='직책', y='인원수', title="직책별 인원", height=300, color='직책')
        fig_rank.update_layout(showlegend=False, margin=dict(t=30, b=0, l=0, r=0))
        st.plotly_chart(fig_rank, use_container_width=True)

    # [섹션 3: 성별 및 부서별 현황]
    c3, c4 = st.columns(2)
    with c3:
        fig_sex = px.pie(active_df, names='성별', title="성별 구성비", height=300, color_discrete_sequence=['#1976d2', '#ec407a'])
        fig_sex.update_layout(margin=dict(t=30, b=0, l=0, r=0))
        st.plotly_chart(fig_sex, use_container_width=True)
    with c4:
        dept_data = active_df['부서'].value_counts().reset_index()
        dept_data.columns = ['부서', '인원수']
        fig_dept = px.bar(dept_data, x='인원수', y='부서', orientation='h', title="부서별 인원", height=300)
        fig_dept.update_layout(margin=dict(t=30, b=0, l=0, r=0))
        st.plotly_chart(fig_dept, use_container_width=True)

    st.markdown("---")

    # [섹션 4: 부서별 입퇴사 현황 (HR 이슈 파악용)]
    st.markdown(f"#### ⚠️ 부서별 입·퇴사 변동 ({report_month}월)")
    
    dept_in = monthly_in['부서'].value_counts().reset_index()
    dept_in.columns = ['부서', '입사']
    dept_out = monthly_out['부서'].value_counts().reset_index()
    dept_out.columns = ['부서', '퇴사']
    
    hr_trend = pd.merge(dept_in, dept_out, on='부서', how='outer').fillna(0)
    if not hr_trend.empty:
        fig_trend = px.bar(hr_trend, x='부서', y=['입사', '퇴사'], barmode='group', 
                           height=350, color_discrete_map={'입사': '#4caf50', '퇴사': '#f44336'})
        st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.write("해당 월에는 입·퇴사 변동 내역이 없습니다.")

    st.markdown("---")

    # [섹션 5: 상세 명단]
    st.markdown("#### 📋 상세 사원 명부 (재직자)")
    st.table(active_df[['사원명', '부서', '직책', '성별', '입사일']].sort_values(by='입사일', ascending=False).head(30))

    st.markdown("<br><p style='text-align: center; color: #999;'>위 내용은 Yuyu Pharma 사원명부 데이터를 바탕으로 자동 생성된 보고서입니다.</p>", unsafe_allow_html=True)

except Exception as e:
    st.error(f"오류가 발생했습니다: {e}")
    st.info("test10.xlsx 파일의 칼럼명이 '사원명', '구분', '부서', '직책', '성별', '입사일', '퇴사일'인지 확인해주세요.")