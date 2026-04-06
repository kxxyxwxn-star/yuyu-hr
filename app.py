import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

# 1. 페이지 설정
st.set_page_config(page_title="Yuyu Pharma HR Dashboard", layout="wide")

# 2. 하이엔드 대시보드 스타일 CSS
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;600;800&display=swap');
    html, body, [class*="css"] { font-family: 'Pretendard', sans-serif; background-color: #F4F7F9; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E6E9EF; }
    div[data-baseweb="select"] > div { border-radius: 8px; border-color: #004a99 !important; }
    span[data-baseweb="tag"] { background-color: #004a99 !important; color: white !important; border-radius: 4px; }
    .main-title { font-size: 32px; font-weight: 800; color: #1A1C1E; margin-bottom: 5px; }
    .sub-title { font-size: 16px; color: #64748B; margin-bottom: 25px; }
    .section-header { 
        font-size: 20px; font-weight: 700; color: #1A1C1E; 
        padding-bottom: 10px; border-bottom: 2px solid #004a99; margin-top: 30px; margin-bottom: 20px;
    }
    thead tr th { background-color: #F8FAFC !important; color: #475569 !important; font-weight: 600 !important; }
    [data-testid="stMetricValue"] { font-weight: 800; color: #004a99; }
    </style>
    """, unsafe_allow_html=True)

# 합계 행 볼드 및 배경색 스타일 함수 (안전한 인덱스 기반)
def style_total_row(df):
    # 기본 스타일 (빈 문자열) 생성
    style_df = pd.DataFrame('', index=df.index, columns=df.columns)
    # 인덱스 이름에 '합계'가 포함된 행만 스타일 지정
    for i, idx in enumerate(df.index):
        if "합계" in str(idx):
            style_df.iloc[i, :] = 'background-color: #F1F5F9; font-weight: bold;'
    return style_df

# 3. 데이터 로드 및 처리
try:
    file_path = "test10.xlsx"
    df = pd.read_excel(file_path)
    
    mtime = os.path.getmtime(file_path)
    file_date = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d')

    def convert_date(date_val):
        if pd.isna(date_val) or date_val == "": return None
        try: return pd.to_datetime(date_val, unit='D', origin='1899-12-30')
        except: return pd.to_datetime(date_val)

    df['입사일'] = df['입사일'].apply(convert_date)
    df['퇴사일'] = df['퇴사일'].apply(convert_date)

    # [직급 통합] 명칭 변경
    rank_map = {
        '명예회장': '임원', '대표이사': '임원', '본부장': '임원', '연구부소장': '임원', '공장장': '임원', '고문': '임원',
        '미화원': '사원', '헬스키퍼': '사원', '반장': '사원', '인턴': '사원'
    }
    df['직책'] = df['직책'].replace(rank_map)
    df['부서'] = df['부서'].apply(lambda x: '임원' if '임원' in str(x) else x)

    # [정렬 순서 정의]
    custom_dept_order = [
        "임원", "종병1지점", "종병2지점", "종병3지점", "종병4지점", "OEM/ODM팀", "OTC도매팀",
        "ETC마케팅실", "ETC마케팅팀", "인사교육팀", "결산세무팀", "복지시설팀", "심사운영팀",
        "분석연구팀", "IT혁신팀", "제제연구팀", "개발팀", "디자인팀", "OTC마케팅팀",
        "e커머스영업마케팅팀", "IT운영팀", "제품연구실", "건기식개발팀",
        "법무감사팀", "펫사업팀", "SCM팀", "홍보팀", "자금팀", "영업기획팀", "대외영업팀",
        "사업개발팀", "수출팀", "공장관리팀", "생산관리팀", "물류팀", "공무팀", "생산실",
        "연질팀", "제조팀", "포장팀", "제품기술팀", "품질보증팀", "품질관리팀"
    ]
    custom_rank_order = ["임원", "지점장", "실장", "팀장", "매니저", "대리", "주임", "4급사원", "5급사원", "사원"]

    with st.sidebar:
        st.markdown("### ⚙️ REPORT SETTING")
        report_year = st.selectbox("기준 연도", options=[2024, 2025, 2026], index=2)
        report_month = st.slider("기준 월", 1, 12, datetime.now().month)
        st.markdown("---")
        all_depts_in_df = sorted(df['부서'].unique())
        type_order = ["임원", "영업", "내근", "공장"] 
        selected_dept = st.multiselect("부서 필터", options=all_depts_in_df, default=all_depts_in_df)
        selected_type = st.multiselect("구분 필터", options=type_order, default=type_order)

    active_df = df[df['퇴사일'].isna() & (df['부서'].isin(selected_dept)) & (df['구분'].isin(selected_type))]
    monthly_in = df[(df['입사일'].dt.year == report_year) & (df['입사일'].dt.month == report_month)]
    monthly_out = df[(df['퇴사일'].dt.year == report_year) & (df['퇴사일'].dt.month == report_month)]

    head_left, head_right = st.columns([3, 1])
    with head_left:
        st.markdown(f'<p class="main-title">Yuyu Pharma HR Dashboard</p>', unsafe_allow_html=True)
        st.markdown(f'<p class="sub-title">유유제약 인원 현황 (인사교육팀) | {report_year}년 {report_month}월 기준</p>', unsafe_allow_html=True)
    with head_right:
        if os.path.exists("yuyu_logo.png"): st.image("yuyu_logo.png", width=180)
        st.caption(f"Data Updated: {file_date}")

    st.markdown('<p class="section-header">📌 구분</p>', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("재직", f"{len(active_df)}명")
    k2.metric("입사", f"{len(monthly_in)}명")
    k3.metric("퇴사", f"{len(monthly_out)}명")
    k4.metric("증감", f"{len(monthly_in) - len(monthly_out)}명")

    st.markdown('<p class="section-header">📊 인원 현황</p>', unsafe_allow_html=True)
    col_left, col_right = st.columns([1, 1.2])

    with col_left:
        st.write("**[부서별 인원현황]**")
        dept_counts = active_df['부서'].value_counts().reset_index()
        dept_counts.columns = ['부서명', '현재원']
        dept_display = dept_counts[dept_counts['현재원'] > 0].copy()
        
        # 정렬 처리
        final_cats = [d for d in custom_dept_order if d in dept_display['부서명'].values] + \
                     sorted([d for d in dept_display['부서명'].values if d not in custom_dept_order])
        dept_display['부서명'] = pd.Categorical(dept_display['부서명'], categories=final_cats, ordered=True)
        dept_display = dept_display.sort_values(by='부서명').set_index('부서명')
        
        # 중복 인덱스 제거 후 합계 추가 (에러 방지 핵심)
        dept_display = dept_display[~dept_display.index.duplicated(keep='first')]
        dept_display.loc['합계'] = dept_display['현재원'].sum()
        st.table(dept_display.style.apply(style_total_row, axis=None))

    with col_right:
        r_top1, r_top2 = st.columns(2)
        with r_top1:
            st.write("**[구분]**")
            t_counts = active_df['구분'].value_counts().reindex(type_order).fillna(0).astype(int).to_frame(name='명')
            t_counts = t_counts[~t_counts.index.duplicated(keep='first')]
            t_counts.loc['합계'] = t_counts['명'].sum()
            st.table(t_counts.style.apply(style_total_row, axis=None))
            
        with r_top2:
            st.write("**[성별]**")
            s_counts = active_df['성별'].value_counts().to_frame(name='명')
            s_counts = s_counts[~s_counts.index.duplicated(keep='first')]
            s_counts.loc['합계'] = s_counts['명'].sum()
            st.table(s_counts.style.apply(style_total_row, axis=None))
        
        st.write("**[직급별]**")
        rank_counts = active_df['직책'].value_counts().reset_index()
        rank_counts.columns = ['직급', '인원']
        rank_counts['직급'] = pd.Categorical(rank_counts['직급'], categories=custom_rank_order, ordered=True)
        rank_display = rank_counts.sort_values(by='직급').set_index('직급')
        rank_display = rank_display[~rank_display.index.duplicated(keep='first')]
        rank_display.loc['합계'] = rank_display['인원'].sum()
        st.table(rank_display.style.apply(style_total_row, axis=None))

    st.markdown('<p class="section-header">📈 입퇴사 현황</p>', unsafe_allow_html=True)
    g1, g2 = st.columns(2)
    with g1:
        d_in = monthly_in['부서'].value_counts().reset_index()
        d_in.columns = ['부서', '명']
        fig_in = px.bar(d_in, x='부서', y='명', title="➕ 입사", color_discrete_sequence=['#004a99'])
        # y축 레이블 방향 정방향(0도) 설정
        fig_in.update_layout(yaxis_title="명", yaxis=dict(dtick=1, tickangle=0), height=280)
        st.plotly_chart(fig_in, use_container_width=True)
    with g2:
        d_out = monthly_out['부서'].value_counts().reset_index()
        d_out.columns = ['부서', '명']
        fig_out = px.bar(d_out, x='부서', y='명', title="➖ 퇴사", color_discrete_sequence=['#E11D48'])
        fig_out.update_layout(yaxis_title="명", yaxis=dict(dtick=1, tickangle=0), height=280)
        st.plotly_chart(fig_out, use_container_width=True)

    st.markdown('<p class="section-header">📝 당월 입사자 명부</p>', unsafe_allow_html=True)
    if not monthly_in.empty:
        display_in = monthly_in[['입사일', '사원명', '부서', '직책', '구분']].sort_values(by='입사일').reset_index(drop=True)
        display_in.index = display_in.index + 1
        display_in.index.name = 'No'
        display_in['입사일'] = display_in['입사일'].dt.strftime('%Y-%m-%d')
        st.table(display_in)
    else:
        st.info("해당 월의 신규 입사 내역이 존재하지 않습니다.")

except Exception as e:
    st.error(f"대시보드 구성 중 오류가 발생했습니다: {e}")