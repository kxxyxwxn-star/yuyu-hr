import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

# 1. 페이지 설정
st.set_page_config(page_title="Yuyu Pharma HR Intelligence", layout="wide")

# 2. 하이엔드 대시보드 스타일 CSS
st.markdown("""
    <style>
    /* 전체 배경색 및 폰트 설정 */
    @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;600;800&display=swap');
    html, body, [class*="css"] { font-family: 'Pretendard', sans-serif; background-color: #F4F7F9; }

    /* 사이드바 디자인 */
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E6E9EF; }
    
    /* 필터 색상 커스터마이징 (유유 블루) */
    div[data-baseweb="select"] > div { border-radius: 8px; border-color: #004a99 !important; }
    span[data-baseweb="tag"] { background-color: #004a99 !important; color: white !important; border-radius: 4px; }

    /* 대시보드 카드 스타일 */
    .metric-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border: 1px solid #E6E9EF;
        text-align: center;
    }
    
    /* 텍스트 스타일 */
    .main-title { font-size: 32px; font-weight: 800; color: #1A1C1E; margin-bottom: 5px; }
    .sub-title { font-size: 16px; color: #64748B; margin-bottom: 25px; }
    .section-header { 
        font-size: 20px; font-weight: 700; color: #1A1C1E; 
        padding-bottom: 10px; border-bottom: 2px solid #004a99; margin-top: 30px; margin-bottom: 20px;
    }

    /* 테이블 스타일 연마 */
    .stTable { background-color: white; border-radius: 10px; overflow: hidden; }
    thead tr th { background-color: #F8FAFC !important; color: #475569 !important; font-weight: 600 !important; }
    
    /* 지표 강조 */
    [data-testid="stMetricValue"] { font-weight: 800; color: #004a99; }
    </style>
    """, unsafe_allow_html=True)

# 3. 데이터 로드 및 처리
try:
    file_path = "test10.xlsx"
    df = pd.read_excel(file_path)
    
    # 발행일 (파일 수정 시간)
    mtime = os.path.getmtime(file_path)
    file_date = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d')

    def convert_date(date_val):
        if pd.isna(date_val) or date_val == "": return None
        try: return pd.to_datetime(date_val, unit='D', origin='1899-12-30')
        except: return pd.to_datetime(date_val)

    df['입사일'] = df['입사일'].apply(convert_date)
    df['퇴사일'] = df['퇴사일'].apply(convert_date)

    # --- [사이드바 필터] ---
    with st.sidebar:
        st.markdown("### ⚙️ REPORT SETTING")
        report_year = st.selectbox("기준 연도", options=[2024, 2025, 2026], index=2)
        report_month = st.slider("기준 월", 1, 12, datetime.now().month)
        st.markdown("---")
        all_depts = sorted(df['부서'].unique())
        all_types = ["임원", "영업", "내근", "공장"] 
        selected_dept = st.multiselect("부서 필터", options=all_depts, default=all_depts)
        selected_type = st.multiselect("구분 필터", options=all_types, default=all_types)

    # 데이터 필터링 로직
    monthly_in = df[(df['입사일'].dt.year == report_year) & (df['입사일'].dt.month == report_month)]
    monthly_out = df[(df['퇴사일'].dt.year == report_year) & (df['퇴사일'].dt.month == report_month)]
    active_df = df[df['퇴사일'].isna() & (df['부서'].isin(selected_dept)) & (df['구분'].isin(selected_type))]

    # --- [메인 대시보드 레이아웃] ---
    
    # 상단 헤더 섹션
    head_left, head_right = st.columns([3, 1])
    with head_left:
        st.markdown(f'<p class="main-title">Yuyu Pharma HR Intelligence</p>', unsafe_allow_html=True)
        st.markdown(f'<p class="sub-title">유유제약 전사 인원 현황 보고 | {report_year}년 {report_month}월 기준</p>', unsafe_allow_html=True)
    with head_right:
        if os.path.exists("yuyu_logo.png"):
            st.image("yuyu_logo.png", width=180)
        st.caption(f"Data Updated: {file_date}")

    # [섹션 1: Key Performance Indicators]
    st.markdown('<p class="section-header">📌 핵심 인원 지표</p>', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("총 재직 인원", f"{len(active_df)}명")
    k2.metric("당월 신규 입사", f"{len(monthly_in)}명", delta=None)
    k3.metric("당월 중도 퇴사", f"{len(monthly_out)}명", delta=None, delta_color="inverse")
    k4.metric("인력 순증감", f"{len(monthly_in) - len(monthly_out)}명")

    # [섹션 2: 데이터 그리드 레이아웃]
    st.markdown('<p class="section-header">📊 부문별 현황 분석</p>', unsafe_allow_html=True)
    
    col_left, col_right = st.columns([1, 1.2])

    with col_left:
        st.write("**[부서별 재직 현황]**")
        dept_counts = active_df['부서'].value_counts().reindex(all_depts).fillna(0).astype(int).reset_index()
        dept_counts.columns = ['부서명', '현재원']
        st.table(dept_counts)

    with col_right:
        # 상단 소형 표 3개 (구분, 직급, 성별)
        r_top1, r_top2 = st.columns(2)
        with r_top1:
            st.write("**[근무 구분]**")
            type_order = ["임원", "영업", "내근", "공장"]
            t_counts = active_df['구분'].value_counts().reindex(type_order).fillna(0).astype(int).reset_index()
            t_counts.columns = ['구분', '인원']
            st.table(t_counts)
        with r_top2:
            st.write("**[성별 구성]**")
            s_counts = active_df['성별'].value_counts().reset_index()
            s_counts.columns = ['성별', '인원']
            st.table(s_counts)
        
        st.write("**[직급별 분포]**")
        rank_counts = active_df['직책'].value_counts().reset_index()
        rank_counts.columns = ['직급', '인원']
        st.table(rank_counts.T) # 직급은 가로로 길게 표시하여 공간 활용

    # [섹션 3: 입퇴사 트렌드 (차트 분리)]
    st.markdown('<p class="section-header">📈 부서별 인력 변동 리포트</p>', unsafe_allow_html=True)
    g1, g2 = st.columns(2)

    with g1:
        d_in = monthly_in['부서'].value_counts().reindex(all_depts).fillna(0).reset_index()
        d_in.columns = ['부서', '명']
        fig_in = px.bar(d_in, x='부서', y='명', title="➕ 부서별 신규 입사", color_discrete_sequence=['#004a99'])
        fig_in.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
                             yaxis=dict(dtick=1, range=[0, 10]), height=280, margin=dict(t=40, b=0, l=0, r=0))
        st.plotly_chart(fig_in, use_container_width=True)

    with g2:
        d_out = monthly_out['부서'].value_counts().reindex(all_depts).fillna(0).reset_index()
        d_out.columns = ['부서', '명']
        fig_out = px.bar(d_out, x='부서', y='명', title="➖ 부서별 중도 퇴사", color_discrete_sequence=['#E11D48'])
        fig_out.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
                              yaxis=dict(dtick=1, range=[0, 10]), height=280, margin=dict(t=40, b=0, l=0, r=0))
        st.plotly_chart(fig_out, use_container_width=True)

    # [섹션 4: 당월 입사자 상세 정보]
    st.markdown('<p class="section-header">📝 당월 신규 입사자 명부</p>', unsafe_allow_html=True)
    if not monthly_in.empty:
        # 오름차순 정렬 및 날짜 포맷팅
        display_in = monthly_in[['입사일', '사원명', '부서', '직책', '구분']].sort_values(by='입사일')
        display_in['입사일'] = display_in['입사일'].dt.strftime('%Y-%m-%d')
        st.table(display_in)
    else:
        st.info("해당 월의 신규 입사 내역이 존재하지 않습니다.")

except Exception as e:
    st.error(f"대시보드 구성 중 오류가 발생했습니다: {e}")