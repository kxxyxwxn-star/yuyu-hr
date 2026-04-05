import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

# 1. 페이지 설정
st.set_page_config(page_title="유유제약 인원 현황", layout="centered")

# CSS 스타일: 요청하신 색상(파란색->하얀색, 빨간색->파란색) 정밀 타격
st.markdown("""
    <style>
    /* 1. 사이드바 전체 배경 및 텍스트 */
    [data-testid="stSidebar"] { background-color: #f8f9fa; }

    /* 2. 필터(MultiSelect) 선택된 항목의 박스 색상 수정 */
    /* 빨간색으로 나오던 배경을 파란색으로, 파란색이었던 전체 배경은 하얀색으로 */
    div[data-baseweb="select"] > div { 
        background-color: white !important; 
        border-color: #007bff !important; 
    }
    
    /* 선택된 항목(Chip)의 배경색을 빨간색에서 파란색으로 */
    span[data-baseweb="tag"] {
        background-color: #007bff !important;
        color: white !important;
    }

    /* 3. 메트릭 및 타이틀 스타일 */
    [data-testid="stMetricValue"] { font-size: 24px; color: #004a99; }
    .report-title { font-size: 30px; font-weight: bold; text-align: center; margin-bottom: 0px; }
    .report-subtitle { font-size: 18px; text-align: center; color: #555; margin-bottom: 20px; }
    
    /* 4. 테이블 헤더 스타일 */
    table { width: 100%; border-collapse: collapse; }
    th { background-color: #f1f8ff; color: #004a99; border-bottom: 2px solid #007bff; }
    
    @media print {
        .no-print { display: none !important; }
    }
    </style>
    """, unsafe_allow_html=True)

# 2. 데이터 불러오기
try:
    file_path = "test10.xlsx"
    df = pd.read_excel(file_path)
    
    # 발행일 자동 인식
    mtime = os.path.getmtime(file_path)
    file_date = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d')

    def convert_date(date_val):
        if pd.isna(date_val) or date_val == "": return None
        try:
            return pd.to_datetime(date_val, unit='D', origin='1899-12-30')
        except:
            return pd.to_datetime(date_val)

    df['입사일'] = df['입사일'].apply(convert_date)
    df['퇴사일'] = df['퇴사일'].apply(convert_date)

    # --- [사이드바 설정] ---
    st.sidebar.header("📅 기준연월 선택")
    report_year = st.sidebar.selectbox("연도", options=[2024, 2025, 2026], index=2)
    report_month = st.sidebar.slider("월", 1, 12, datetime.now().month)
    
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 상세 필터")
    
    all_depts = sorted(df['부서'].unique())
    all_types = ["임원", "영업", "내근", "공장"] 
    
    selected_dept = st.sidebar.multiselect("부서 선택", options=all_depts, default=all_depts)
    selected_type = st.sidebar.multiselect("근무구분 선택", options=all_types, default=all_types)

    # 데이터 필터링
    monthly_in = df[(df['입사일'].dt.year == report_year) & (df['입사일'].dt.month == report_month)].sort_values(by='입사일', ascending=True)
    monthly_out = df[(df['퇴사일'].dt.year == report_year) & (df['퇴사일'].dt.month == report_month)]
    active_df = df[df['퇴사일'].isna() & (df['부서'].isin(selected_dept)) & (df['구분'].isin(selected_type))]

    # --- [메인 보고서 시작] ---
    st.markdown('<p class="report-title">유유제약 인원 현황 (인사교육팀)</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="report-subtitle">{report_year}년 {report_month}월 기준</p>', unsafe_allow_html=True)
    st.write(f"**발행일:** {file_date}")
    st.markdown("---")

    # [상단 요약]
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("재직", f"{len(active_df)}명")
    col_m2.metric("당월 입사자", f"{len(monthly_in)}명")
    col_m3.metric("당월 퇴사자", f"{len(monthly_out)}명")
    st.markdown("---")

    # [중단 레이아웃]
    st.markdown("#### 📊 인원현황")
    left_col, right_col = st.columns([1, 1])

    with left_col:
        st.write("**[부서별 인원]**")
        dept_counts = active_df['부서'].value_counts().reindex(all_depts).fillna(0).astype(int).reset_index()
        dept_counts.columns = ['부서명', '인원']
        st.table(dept_counts)

    with right_col:
        # [구분] 정렬: 임원, 영업, 내근, 공장
        st.write("**[구분]**")
        type_order = ["임원", "영업", "내근", "공장"]
        type_counts = active_df['구분'].value_counts().reindex(type_order).fillna(0).astype(int).reset_index()
        type_counts.columns = ['항목', '명']
        st.table(type_counts)

        # [직급]
        st.write("**[직급]**")
        rank_counts = active_df['직책'].value_counts().reset_index()
        rank_counts.columns = ['직급', '명']
        st.table(rank_counts)

        # [성별]
        st.write("**[성별]**")
        sex_counts = active_df['성별'].value_counts().reset_index()
        sex_counts.columns = ['성별', '명']
        st.table(sex_counts)

    st.markdown("---")

    # [섹션: 부서별 입퇴사 변동]
    st.markdown(f"#### ⚠️ 부서별 입·퇴사 변동 ({report_month}월)")
    dept_in = monthly_in['부서'].value_counts().reset_index()
    dept_in.columns = ['부서', '입사']
    dept_out = monthly_out['부서'].value_counts().reset_index()
    dept_out.columns = ['부서', '퇴사']
    hr_trend = pd.merge(dept_in, dept_out, on='부서', how='outer').fillna(0)

    if not hr_trend.empty:
        fig_trend = px.bar(hr_trend, x='부서', y=['입사', '퇴사'], barmode='group', height=300,
                           color_discrete_map={'입사': '#1976d2', '퇴사': '#d32f2f'})
        fig_trend.update_layout(yaxis=dict(dtick=1, range=[0, 10]))
        st.plotly_chart(fig_trend, use_container_width=True)
    else:
        st.write("변동 내역 없음")

    st.markdown("---")

    # [하단: 당월 입사자 명부]
    st.markdown("#### 📝 당월 입사자 명부")
    if not monthly_in.empty:
        st.table(monthly_in[['입사일', '사원명', '부서', '직책', '구분']].assign(입사일=lambda x: x['입사일'].dt.strftime('%Y-%m-%d')))
    else:
        st.write("해당 월에 입사자가 없습니다.")

except Exception as e:
    st.error(f"오류 발생: {e}")