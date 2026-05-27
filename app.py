import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import glob, os

# ── 페이지 설정 ──────────────────────────────────────────────
st.set_page_config(
    page_title="명지대학교 교수연구 지표 대시보드",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 스타일 ───────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
.stApp { background-color: #F5F7FA; }
[data-testid="stSidebar"] { background-color: #0D2B5E; }
[data-testid="stSidebar"] * { color: #E8EDF5 !important; }
[data-testid="metric-container"] {
    background: white; border: 1px solid #E4E9F2;
    border-radius: 12px; padding: 16px 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
[data-testid="stMetricValue"] {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.9rem !important; color: #0D2B5E !important; font-weight: 600;
}
[data-testid="stMetricLabel"] {
    font-size: 11px !important; color: #6B7A99 !important;
    font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em;
}
[data-testid="stMetricDelta"] { font-size: 12px !important; }
.page-title    { font-size: 22px; font-weight: 700; color: #0D2B5E; margin-bottom: 4px; }
.page-subtitle { font-size: 13px; color: #6B7A99; margin-bottom: 24px; }
.section-header {
    font-size: 11px; font-weight: 700; color: #6B7A99;
    text-transform: uppercase; letter-spacing: 0.1em;
    margin: 28px 0 12px; padding-bottom: 8px; border-bottom: 1px solid #E4E9F2;
}
.card {
    background: white; border-radius: 12px; padding: 18px 22px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06); margin-bottom: 12px;
}
.card-title { font-size: 13px; font-weight: 700; color: #0D2B5E; margin-bottom: 6px; font-family: 'IBM Plex Mono', monospace; }
.card-body  { font-size: 12px; color: #4A5568; line-height: 1.7; }
.card-note  { font-size: 11px; color: #6B7A99; margin-top: 6px; }
.highlight-box {
    background: #EEF4FF; border-left: 4px solid #0D2B5E;
    border-radius: 0 10px 10px 0; padding: 12px 18px;
    font-size: 13px; color: #1E3A8A; margin-top: 10px; line-height: 1.7;
}
.warn-box {
    background: #FFF7ED; border-left: 4px solid #F97316;
    border-radius: 0 10px 10px 0; padding: 12px 18px;
    font-size: 13px; color: #9A3412; margin-top: 10px; line-height: 1.7;
}
.good-box {
    background: #F0FDF4; border-left: 4px solid #16A34A;
    border-radius: 0 10px 10px 0; padding: 12px 18px;
    font-size: 13px; color: #166534; margin-top: 10px; line-height: 1.7;
}
.footnote {
    font-size: 11px; color: #9AA3B2; margin-top: 24px;
    padding-top: 12px; border-top: 1px solid #E4E9F2;
}
.badge-mj {
    display: inline-block; background: #E8392A; color: white;
    font-size: 10px; font-weight: 700; padding: 2px 7px;
    border-radius: 4px; margin-left: 6px; vertical-align: middle;
}
.badge-jj {
    display: inline-block; background: #0D2B5E; color: white;
    font-size: 10px; font-weight: 700; padding: 2px 7px;
    border-radius: 4px; margin-left: 6px; vertical-align: middle;
}
</style>
""", unsafe_allow_html=True)

# ── 상수 ────────────────────────────────────────────────────
MJ_NAME   = "명지대"
MJ_COLOR  = "#E8392A"
BASE_COLOR = "#93B8E0"
JJ_COLOR  = "#0D2B5E"

LAYOUT = dict(
    font=dict(family="Noto Sans KR, sans-serif", size=12, color="#2D3748"),
    plot_bgcolor="white", paper_bgcolor="white",
    margin=dict(l=20, r=20, t=40, b=20),
    hoverlabel=dict(bgcolor="white", font_size=12),
)

# ── 파일 탐색 ────────────────────────────────────────────────
def find_excel():
    base = os.path.dirname(os.path.abspath(__file__))
    files = glob.glob(os.path.join(base, "*.xlsx")) + glob.glob(os.path.join(base, "*.xls"))
    return max(files, key=os.path.getmtime) if files else None

# ── 데이터 로드 ──────────────────────────────────────────────
@st.cache_data
def load_data(path: str):
    xl = pd.read_excel(path, sheet_name=None)

    # ── Sheet 1: 국제학술지 논문당 피인용 ──
    d1 = xl['국제학술지 논문당 피인용_only articles'].copy()
    d1.columns = d1.iloc[0]
    d1 = d1.iloc[1:].reset_index(drop=True)
    for c in ['Publications','Citations','Self-citations',
              'Citations Per Publication','FWCI',
              'FWCI times JoongAng Ratio','Average of FWCI times JoongAng Ratio']:
        d1[c] = pd.to_numeric(d1[c], errors='coerce')
    pub_col = '국제학술지 논문당 피인용\n(중앙일보 2025.11 공시)'
    d1[pub_col] = d1[pub_col].replace('없음', None).replace('-', None)
    d1[pub_col] = pd.to_numeric(d1[pub_col], errors='coerce')
    d1 = d1.sort_values('Average of FWCI times JoongAng Ratio', ascending=False).reset_index(drop=True)
    d1['산출순위'] = d1.index + 1
    d1['is_mj'] = d1['대학명'] == MJ_NAME
    # 자기인용 제외 비율
    d1['자기인용_비율'] = d1['Self-citations'] / d1['Citations'] * 100
    d1['상위저널_비율'] = d1['Publications in top journal percentiles'] / d1['Publications'] * 100
    d1['상위피인용_비율'] = d1['Publications in top citation percentiles'] / d1['Publications'] * 100

    # ── Sheet 2: 국제협업논문 ──
    d2 = xl['국제협업논문_only articles'].copy()
    d2.columns = d2.iloc[0]
    d2 = d2.iloc[1:].reset_index(drop=True)
    for c in ['Publications','Publications_국제협업','국제협업 논문']:
        d2[c] = pd.to_numeric(d2[c], errors='coerce')
    jj_col2 = '국제협업논문\n(중앙일보 2025.11 공시)'
    d2[jj_col2] = d2[jj_col2].replace({'없음': None, '-': None})
    d2[jj_col2] = pd.to_numeric(d2[jj_col2], errors='coerce')
    d2 = d2.sort_values('국제협업 논문', ascending=False).reset_index(drop=True)
    d2['산출순위'] = d2.index + 1
    d2['is_mj'] = d2['대학명'] == MJ_NAME

    # ── Sheet 3: 인문사회 국내논문당 피인용 ──
    d3 = xl['인문사회 국내논문당 피인용_only articles'].copy()
    d3.columns = d3.iloc[1]
    d3 = d3.iloc[2:].reset_index(drop=True)
    d3 = d3[d3['Id'].notna() & (d3['Id'] != '[기준값]')].reset_index(drop=True)
    # 중복 컬럼 인덱스로 처리
    cols = d3.columns.tolist()
    d3.columns = [
        'Id','대학명(영문)','대학명(한글)',
        '인문_논문수','사회_논문수','체육_논문수','총_논문수',
        'P_인문','P_사회','P_체육',
        'Avg_인문','Avg_사회','Avg_체육',
        'Score_인문','Score_사회','Score_체육',
        '실값','Z값','점수','순위',
        'JJ_실값','JJ_Z값','JJ_점수','JJ_순위'
    ]
    for c in ['인문_논문수','사회_논문수','체육_논문수','총_논문수',
              'Avg_인문','Avg_사회','Avg_체육',
              'Score_인문','Score_사회','Score_체육',
              '실값','Z값','점수','순위','JJ_실값','JJ_Z값','JJ_점수','JJ_순위']:
        d3[c] = pd.to_numeric(d3[c], errors='coerce')
    d3 = d3.sort_values('점수', ascending=False).reset_index(drop=True)
    d3['산출순위'] = d3.index + 1
    d3['is_mj'] = d3['대학명(한글)'] == MJ_NAME

    return d1, d2, d3

# ── 사이드바 ─────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🎓 명지대학교\n**교수연구 지표 대시보드**")
    st.markdown("---")
    page = st.radio(
        "페이지",
        ["📋 지표 해설",
         "📊 3개 지표 전체 순위",
         "🔍 지표별 심층 분석",
         "🏫 명지대 종합 프로파일"],
        label_visibility="collapsed",
    )
    st.markdown("---")
    st.markdown(
        "<div style='font-size:11px;color:#93A8D4;'>"
        "산출 기준: 2026년 5월<br>"
        "비교 기준: 중앙일보 2025년 11월 공시<br>"
        "대상: 54개 대학 (명지대 포함)</div>",
        unsafe_allow_html=True,
    )

# ── 파일 로드 ────────────────────────────────────────────────
excel_path = find_excel()
if excel_path is None:
    st.error("❌ 같은 폴더에 엑셀 파일(.xlsx)이 없습니다.")
    st.stop()

d1, d2, d3 = load_data(excel_path)
with st.sidebar:
    st.markdown(
        f"<div style='font-size:11px;color:#93A8D4;margin-top:8px;'>"
        f"📂 {os.path.basename(excel_path)}</div>",
        unsafe_allow_html=True,
    )

mj1 = d1[d1['is_mj']].iloc[0]
mj2 = d2[d2['is_mj']].iloc[0]
mj3 = d3[d3['is_mj']].iloc[0]
total = len(d1)

# ════════════════════════════════════════════════════════════
# PAGE 1 — 지표 해설
# ════════════════════════════════════════════════════════════
if page == "📋 지표 해설":
    st.markdown("<div class='page-title'>지표 해설</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>중앙일보 대학평가 교수연구 부문 3개 지표의 산출 기준 및 방법</div>", unsafe_allow_html=True)

    # ── 데이터 기준 안내 ──
    st.markdown("<div class='section-header'>데이터 기준 안내</div>", unsafe_allow_html=True)
    st.markdown("""
    <div class='card'>
        <div class='card-title'>🔵 산출값 (본 대시보드)</div>
        <div class='card-body'>
            Scholytics(스콜리틱스) 기반으로 <b>2026년 5월</b> 기준 직접 산출한 수치입니다.<br>
            중앙일보 선정 <b>53개 대학 + 명지대학교</b>, 총 <b>54개 대학</b>의 데이터가 포함됩니다.
        </div>
        <div class='card-note'>
            ※ 중앙일보 공시값은 2025년 11월 기준 선정 53개교에만 존재하며, 명지대는 해당 없습니다.<br>
            ※ 기준 시점 차이(2026.5 vs 2025.11)로 두 값의 직접 비교는 참고용입니다.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── 기본 지표 안내 ──
    st.markdown("<div class='section-header'>기본 지표 안내</div>", unsafe_allow_html=True)
    st.markdown("""
    <div class='card'>
        <div class='card-title'>Scholytics 원천 데이터 지표</div>
        <div class='card-body'>
            중앙일보 대학평가 교수연구 부문 산출에 사용되는 Scholytics 기본 지표입니다.<br>
            아래 지표들을 조합하여 각 평가 지표의 순위를 산출합니다.
        </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        <div class='card'>
            <div class='card-title'>Publications</div>
            <div class='card-body'>평가 기간 내 전체 논문 수. 대학의 연구 규모를 나타내는 기본 지표입니다.</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("""
        <div class='card'>
            <div class='card-title'>Citations</div>
            <div class='card-body'>전체 논문이 타 논문에 인용된 횟수. 연구의 영향력을 나타내는 절대적 수치입니다.</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown("""
        <div class='card'>
            <div class='card-title'>FWCI (Field-Weighted Citation Impact)</div>
            <div class='card-body'>
                동일 발행연도·분야·문서유형 기준 <b>세계 평균 대비 피인용 비율</b>입니다.<br>
                · 1.0 = 세계 평균<br>
                · 1.0 초과 = 세계 평균 이상<br>
                · 1.0 미만 = 세계 평균 이하
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("""
        <div class='card'>
            <div class='card-title'>Publications in top journal percentiles</div>
            <div class='card-body'>상위 저널에 게재된 논문 수입니다. 연구 성과의 질적 수준을 나타냅니다.</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("""
        <div class='card'>
            <div class='card-title'>Publications in top citation percentiles</div>
            <div class='card-body'>동일 발행연도·분야 기준 상위 피인용 논문 수입니다. 해당 분야에서 주목받는 논문의 수를 나타냅니다.</div>
        </div>
        """, unsafe_allow_html=True)

    # ── 지표 ① ──
    st.markdown("<div class='section-header'>지표 ① — 국제학술지 논문당 피인용</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        <div class='card'>
            <div class='card-title'>산출 조건</div>
            <div class='card-body'>
                · 데이터 출처: <b>Scholytics Index</b> (Scopus 기반)<br>
                · 평가 기간: 평가연도 기준 <b>-5년 ~ -2년 (총 4년)</b><br>
                · Self-citation(자기인용): <b>제외</b> (저자 기준)<br>
                · Potentially predatory journal(의심학술지): <b>제외</b>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class='card'>
            <div class='card-title'>순위 결정 기준 지표</div>
            <div class='card-body'>
                <b>Average of FWCI times JoongAng Ratio</b><br><br>
                = Σ(논문별 FWCI × 중앙일보 가중치) ÷ 전체 논문수<br><br>
                이 값이 높을수록 분야 평균 대비 피인용 영향력이 높은 논문을 많이 보유하고 있음을 의미하며,
                <b>기관별 순위의 결정 기준</b>이 됩니다.
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── 지표 ② ──
    st.markdown("<div class='section-header'>지표 ② — 국제협업논문</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        <div class='card'>
            <div class='card-title'>산출 조건</div>
            <div class='card-body'>
                · <b>2024년 중앙일보 대학평가부터 종합 평가 항목으로 신규 추가</b><br>
                · 발행연도·저널인덱스·자기인용 제외 조건은 지표 ①과 동일<br>
                · Collaboration Type: <b>International Collaboration</b> 적용
            </div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class='card'>
            <div class='card-title'>순위 결정 기준 지표</div>
            <div class='card-body'>
                <b>국제협업 비중</b><br><br>
                = 국제협업 문서수 ÷ 전체 문서수<br><br>
                전체 논문 중 해외 기관과 공동 저술한 논문의 비율로,
                국제 연구 네트워크의 활성화 정도를 나타냅니다.
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── 지표 ③ ──
    st.markdown("<div class='section-header'>지표 ③ — 인문사회 국내논문당 피인용</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        <div class='card'>
            <div class='card-title'>산출 조건</div>
            <div class='card-body'>
                · 데이터 출처: <b>KCI Premium + KCI + KCI Candidate</b> (KCI 계열 3개 인덱스)<br>
                · 평가 분야: <b>인문 / 사회 / 체육</b> 3개 분야 (이공계 제외)<br>
                · Subject Area: 중앙일보 카테고리 적용<br>
                · 발행연도: 지표 ①과 동일
            </div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class='card'>
            <div class='card-title'>산출 방법 (6단계)</div>
            <div class='card-body'>
                ① 분야별 <b>Average of Citations times JoongAng Ratio</b> 추출<br>
                &nbsp;&nbsp;&nbsp;(각 논문의 인용수 × 중앙일보 가중치 합산 ÷ 논문수)<br>
                ② 분야별 논문수 비율 산출<br>
                &nbsp;&nbsp;&nbsp;P_인문 / P_사회 / P_체육<br>
                ③ 분야별 중간 지표 산출 (Z-score 방식)<br>
                &nbsp;&nbsp;&nbsp;Score = (대학별 분야값 ÷ 전체 평가대학 분야 평균) × 논문비율<br>
                ④ 분야별 Score 합산 → 실값<br>
                ⑤ Z값 환산 → 점수 산출<br>
                ⑥ 점수 내림차순 → <b>순위 결정</b>
            </div>
            <div class='card-note'>※ Z기준값: 명지대를 제외한 53개교 평균 적용</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div class='footnote'>
        ※ 산출 기준: Scholytics Index (국제학술지·국제협업), KCI 계열 (인문사회) &nbsp;|&nbsp;
        산출 시점: 2026년 5월 &nbsp;|&nbsp; 중앙일보 공시: 2025년 11월
    </div>
    """, unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# PAGE 2 — 3개 지표 전체 순위
# ════════════════════════════════════════════════════════════
elif page == "📊 3개 지표 전체 순위":
    st.markdown("<div class='page-title'>3개 지표 전체 순위</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>54개 대학 산출값 기준 순위 · 명지대 위치 강조</div>", unsafe_allow_html=True)

    # 명지대 요약 카드
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric(
            "국제학술지 논문당 피인용",
            f"{int(mj1['산출순위'])}위 / {total}개교",
            f"상위 {mj1['산출순위']/total*100:.0f}%",
            delta_color="off"
        )
    with c2:
        st.metric(
            "국제협업논문",
            f"{int(mj2['산출순위'])}위 / {total}개교",
            f"상위 {mj2['산출순위']/total*100:.0f}%",
            delta_color="off"
        )
    with c3:
        st.metric(
            "인문사회 국내논문당 피인용",
            f"{int(mj3['산출순위'])}위 / {total}개교",
            f"상위 {mj3['산출순위']/total*100:.0f}%",
            delta_color="off"
        )

    st.markdown("""
    <div class='highlight-box'>
        <b>★ 명지대학교는 중앙일보 2025년 선정 53개교에 포함되지 않았으나</b>,
        동일 기준으로 산출 시 국제학술지 논문당 피인용 <b>8위</b>,
        인문사회 국내논문당 피인용 <b>22위</b> 수준의 연구 경쟁력을 보유하고 있습니다.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div class='section-header'>지표 선택</div>", unsafe_allow_html=True)
    tab1, tab2, tab3 = st.tabs(["① 국제학술지 논문당 피인용", "② 국제협업논문", "③ 인문사회 국내논문당 피인용"])

    # ── Tab 1 ──
    with tab1:
        top_n = st.selectbox("표시 대학 수", [total, 30, 20], index=0,
                             format_func=lambda x: f"전체 {x}개교" if x == total else f"상위 {x}개교",
                             key="t1")
        plot = d1.head(top_n).sort_values('Average of FWCI times JoongAng Ratio', ascending=True)
        colors = [MJ_COLOR if v else BASE_COLOR for v in plot['is_mj']]

        fig = go.Figure()
        fig.add_trace(go.Bar(
            y=plot['대학명'], x=plot['Average of FWCI times JoongAng Ratio'],
            orientation='h', marker_color=colors,
            customdata=list(zip(plot['산출순위'], plot['FWCI'],
                                plot['Citations Per Publication'],
                                plot['is_mj'].map({True:'★ 명지대학교', False:'비교대학'}))),
            hovertemplate="<b>%{y}</b><br>산출값: %{x:.4f}<br>순위: %{customdata[0]}위<br>FWCI: %{customdata[1]:.3f}<br>논문당피인용: %{customdata[2]:.2f}<br>%{customdata[3]}<extra></extra>",
            showlegend=False,
        ))
        fig.add_trace(go.Bar(x=[None], y=[None], orientation='h',
                             marker_color=BASE_COLOR, name='비교대학'))
        fig.add_trace(go.Bar(x=[None], y=[None], orientation='h',
                             marker_color=MJ_COLOR, name='★ 명지대학교'))
        fig.update_layout(**LAYOUT, height=max(420, top_n * 26),
            yaxis=dict(autorange=True, tickfont=dict(size=11)),
            xaxis=dict(title='산출값 (FWCI×중앙비율 평균)', gridcolor='#F0F0F0'),
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        )
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("""
        <div class='footnote'>
            산출값: FWCI × 중앙일보 반영 비율의 대학별 평균 &nbsp;|&nbsp;
            ★ 명지대학교는 중앙일보 2025년 53개교 미포함, 동일 기준 자체 산출
        </div>""", unsafe_allow_html=True)

    # ── Tab 2 ──
    with tab2:
        top_n2 = st.selectbox("표시 대학 수", [total, 30, 20], index=0,
                              format_func=lambda x: f"전체 {x}개교" if x == total else f"상위 {x}개교",
                              key="t2")
        plot2 = d2.head(top_n2).sort_values('국제협업 논문', ascending=True)
        colors2 = [MJ_COLOR if v else BASE_COLOR for v in plot2['is_mj']]

        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            y=plot2['대학명'], x=plot2['국제협업 논문'],
            orientation='h', marker_color=colors2,
            customdata=list(zip(plot2['산출순위'], plot2['Publications'],
                                plot2['Publications_국제협업'],
                                plot2['is_mj'].map({True:'★ 명지대학교', False:'비교대학'}))),
            hovertemplate="<b>%{y}</b><br>국제협업률: %{x:.4f}<br>순위: %{customdata[0]}위<br>전체논문: %{customdata[1]:,}편<br>협업논문: %{customdata[2]:,}편<br>%{customdata[3]}<extra></extra>",
            showlegend=False,
        ))
        fig2.add_trace(go.Bar(x=[None], y=[None], orientation='h',
                              marker_color=BASE_COLOR, name='비교대학'))
        fig2.add_trace(go.Bar(x=[None], y=[None], orientation='h',
                              marker_color=MJ_COLOR, name='★ 명지대학교'))
        fig2.update_layout(**LAYOUT, height=max(420, top_n2 * 26),
            yaxis=dict(autorange=True, tickfont=dict(size=11)),
            xaxis=dict(title='국제협업 논문 비율', gridcolor='#F0F0F0', tickformat='.3f'),
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        )
        st.plotly_chart(fig2, use_container_width=True)
        st.markdown("""
        <div class='footnote'>
            국제협업률 = 국제협업논문 수 / 전체논문 수 &nbsp;|&nbsp;
            ★ 명지대학교는 중앙일보 2025년 53개교 미포함, 동일 기준 자체 산출
        </div>""", unsafe_allow_html=True)

    # ── Tab 3 ──
    with tab3:
        top_n3 = st.selectbox("표시 대학 수", [total, 30, 20], index=0,
                              format_func=lambda x: f"전체 {x}개교" if x == total else f"상위 {x}개교",
                              key="t3")
        plot3 = d3.head(top_n3).sort_values('점수', ascending=True)
        colors3 = [MJ_COLOR if v else BASE_COLOR for v in plot3['is_mj']]

        fig3 = go.Figure()
        fig3.add_trace(go.Bar(
            y=plot3['대학명(한글)'], x=plot3['점수'],
            orientation='h', marker_color=colors3,
            customdata=list(zip(plot3['산출순위'], plot3['실값'], plot3['Z값'],
                                plot3['is_mj'].map({True:'★ 명지대학교', False:'비교대학'}))),
            hovertemplate="<b>%{y}</b><br>점수: %{x:.3f}<br>순위: %{customdata[0]}위<br>실값: %{customdata[1]:.4f}<br>Z값: %{customdata[2]:.4f}<br>%{customdata[3]}<extra></extra>",
            showlegend=False,
        ))
        fig3.add_trace(go.Bar(x=[None], y=[None], orientation='h',
                              marker_color=BASE_COLOR, name='비교대학'))
        fig3.add_trace(go.Bar(x=[None], y=[None], orientation='h',
                              marker_color=MJ_COLOR, name='★ 명지대학교'))
        fig3.update_layout(**LAYOUT, height=max(420, top_n3 * 26),
            yaxis=dict(autorange=True, tickfont=dict(size=11)),
            xaxis=dict(title='인문사회 피인용 점수', gridcolor='#F0F0F0'),
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        )
        st.plotly_chart(fig3, use_container_width=True)
        st.markdown("""
        <div class='footnote'>
            점수 = 인문·사회·체육 분야별 Z점수 합산 &nbsp;|&nbsp; Z기준: 명지대 제외 53개교 평균 &nbsp;|&nbsp;
            ★ 명지대학교는 중앙일보 2025년 53개교 미포함, 동일 기준 자체 산출
        </div>""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# PAGE 3 — 지표별 심층 분석
# ════════════════════════════════════════════════════════════
elif page == "🔍 지표별 심층 분석":
    st.markdown("<div class='page-title'>지표별 심층 분석</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>각 지표의 세부 구성 요소와 명지대 위치 분석</div>", unsafe_allow_html=True)

    sel = st.radio("분석 지표 선택", ["① 국제학술지 논문당 피인용", "② 국제협업논문", "③ 인문사회 국내논문당 피인용"],
                   horizontal=True)

    # ── 심층 ①: 국제학술지 ──
    if sel == "① 국제학술지 논문당 피인용":
        st.markdown("<div class='section-header'>순위 결정 지표(Average of FWCI times JoongAng Ratio) × 논문 규모 분포</div>", unsafe_allow_html=True)
        avg_pub  = d1['Publications'].mean()
        avg_afwci = d1['Average of FWCI times JoongAng Ratio'].mean()

        fig = px.scatter(
            d1, x='Publications', y='Average of FWCI times JoongAng Ratio',
            size='Citations', color='is_mj',
            hover_name='대학명',
            color_discrete_map={True: MJ_COLOR, False: BASE_COLOR},
            hover_data={
                'Publications': ':,',
                'Average of FWCI times JoongAng Ratio': ':.4f',
                'Citations': ':,',
                'is_mj': False
            },
            labels={
                'Publications': '논문 수 (편)',
                'Average of FWCI times JoongAng Ratio': 'Avg FWCI×중앙비율',
                'Citations': '피인용 수'
            },
            size_max=50,
        )
        fig.add_hline(y=avg_afwci, line_dash='dot', line_color='#718096', line_width=1,
                      annotation_text=f'Avg FWCI×중앙비율 평균({avg_afwci:.4f})',
                      annotation_font=dict(size=10, color='#718096'))
        fig.add_vline(x=avg_pub, line_dash='dot', line_color='#718096', line_width=1,
                      annotation_text=f'논문수 평균({int(avg_pub):,}편)',
                      annotation_font=dict(size=10, color='#718096'))
        fig.update_layout(**LAYOUT, height=500,
            showlegend=False,
            xaxis=dict(title='논문 수 (편)', gridcolor='#F0F0F0', tickformat=','),
            yaxis=dict(title='Avg FWCI × 중앙일보 비율', gridcolor='#F0F0F0'),
        )
        st.plotly_chart(fig, use_container_width=True)

        # 명지대 위치 해설
        mj_afwci = mj1['Average of FWCI times JoongAng Ratio']
        if mj1['Publications'] < avg_pub and mj_afwci >= avg_afwci:
            box_cls, msg = 'highlight-box', f"명지대는 <b>논문 규모 평균 이하({int(mj1['Publications']):,}편)</b>이지만, <b>순위 결정 지표(Avg FWCI×중앙비율)는 평균 이상({mj_afwci:.4f})</b>으로 소규모 고품질 연구 그룹에 위치합니다."
        elif mj1['Publications'] < avg_pub and mj_afwci < avg_afwci:
            box_cls, msg = 'warn-box', f"명지대는 논문 규모({int(mj1['Publications']):,}편)와 순위 결정 지표({mj_afwci:.4f}) 모두 평균 이하입니다. 연구 규모 및 질적 영향력 동시 제고가 필요합니다."
        elif mj1['Publications'] >= avg_pub and mj_afwci >= avg_afwci:
            box_cls, msg = 'good-box', f"명지대는 논문 규모({int(mj1['Publications']):,}편)와 순위 결정 지표({mj_afwci:.4f}) 모두 평균 이상의 우수한 그룹에 위치합니다."
        else:
            box_cls, msg = 'warn-box', f"명지대는 논문 규모({int(mj1['Publications']):,}편)는 평균 이상이나 순위 결정 지표({mj_afwci:.4f})는 평균 이하로, 질적 영향력 제고가 필요합니다."
        st.markdown(f"<div class='{box_cls}'>{msg}</div>", unsafe_allow_html=True)

        st.markdown("<div class='section-header'>자기인용 비율 vs 상위 피인용 논문 비율</div>", unsafe_allow_html=True)
        avg_self = d1['자기인용_비율'].mean()
        avg_top  = d1['상위피인용_비율'].mean()

        fig2 = px.scatter(
            d1, x='자기인용_비율', y='상위피인용_비율',
            color='is_mj', hover_name='대학명',
            color_discrete_map={True: MJ_COLOR, False: BASE_COLOR},
            hover_data={'자기인용_비율': ':.1f', '상위피인용_비율': ':.1f', 'is_mj': False},
            labels={'자기인용_비율': '자기인용 비율 (%)', '상위피인용_비율': '상위 피인용 논문 비율 (%)'},
        )
        fig2.add_hline(y=avg_top, line_dash='dot', line_color='#718096', line_width=1,
                       annotation_text=f'평균({avg_top:.1f}%)',
                       annotation_font=dict(size=10, color='#718096'))
        fig2.add_vline(x=avg_self, line_dash='dot', line_color='#718096', line_width=1,
                       annotation_text=f'평균({avg_self:.1f}%)',
                       annotation_font=dict(size=10, color='#718096'))
        fig2.update_layout(**LAYOUT, height=420, showlegend=False,
            xaxis=dict(title='자기인용 비율 (%)', gridcolor='#F0F0F0'),
            yaxis=dict(title='상위 피인용 논문 비율 (%)', gridcolor='#F0F0F0'),
        )
        st.plotly_chart(fig2, use_container_width=True)

        c1, c2, c3 = st.columns(3)
        c1.metric("명지대 자기인용 비율", f"{mj1['자기인용_비율']:.1f}%", f"평균 {avg_self:.1f}%")
        c2.metric("명지대 상위 피인용 논문 비율", f"{mj1['상위피인용_비율']:.1f}%", f"평균 {avg_top:.1f}%")
        c3.metric("명지대 상위 저널 논문 비율", f"{mj1['상위저널_비율']:.1f}%",
                  f"평균 {d1['상위저널_비율'].mean():.1f}%")

    # ── 심층 ②: 국제협업 ──
    elif sel == "② 국제협업논문":
        st.markdown("<div class='section-header'>논문 규모 vs 국제협업률</div>", unsafe_allow_html=True)
        avg_pub2 = d2['Publications'].mean()
        avg_rate = d2['국제협업 논문'].mean()

        fig = px.scatter(
            d2, x='Publications', y='국제협업 논문',
            size='Publications_국제협업', color='is_mj',
            hover_name='대학명',
            color_discrete_map={True: MJ_COLOR, False: BASE_COLOR},
            hover_data={'Publications': ':,', '국제협업 논문': ':.4f',
                        'Publications_국제협업': ':,', 'is_mj': False},
            labels={'Publications': '전체 논문 수 (편)', '국제협업 논문': '국제협업률',
                    'Publications_국제협업': '국제협업 논문 수'},
            size_max=50,
        )
        fig.add_hline(y=avg_rate, line_dash='dot', line_color='#718096', line_width=1,
                      annotation_text=f'평균 협업률({avg_rate:.3f})',
                      annotation_font=dict(size=10, color='#718096'))
        fig.add_vline(x=avg_pub2, line_dash='dot', line_color='#718096', line_width=1,
                      annotation_text=f'논문수 평균({int(avg_pub2):,}편)',
                      annotation_font=dict(size=10, color='#718096'))
        fig.update_layout(**LAYOUT, height=500, showlegend=False,
            xaxis=dict(title='전체 논문 수 (편)', gridcolor='#F0F0F0', tickformat=','),
            yaxis=dict(title='국제협업률', gridcolor='#F0F0F0'),
        )
        st.plotly_chart(fig, use_container_width=True)

        if mj2['국제협업 논문'] < avg_rate:
            st.markdown(f"""
            <div class='warn-box'>
                명지대 국제협업률 <b>{mj2['국제협업 논문']:.4f}</b>로 전체 평균({avg_rate:.4f}) 대비
                <b>{(avg_rate - mj2['국제협업 논문'])*100:.2f}%p 하회</b>합니다.
                전체 {int(mj2['산출순위'])}위 / 54개교로, 국제 연구 네트워크 확대가 필요한 영역입니다.
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class='good-box'>
                명지대 국제협업률 <b>{mj2['국제협업 논문']:.4f}</b>로 전체 평균({avg_rate:.4f}) 이상입니다.
                전체 {int(mj2['산출순위'])}위 / 54개교.
            </div>""", unsafe_allow_html=True)

        st.markdown("<div class='section-header'>국제협업률 상위 20개교 비교</div>", unsafe_allow_html=True)
        top20 = d2.head(20).copy()
        # 명지대가 20위 밖이면 추가
        if not top20['is_mj'].any():
            top20 = pd.concat([top20, d2[d2['is_mj']]]).reset_index(drop=True)
        top20 = top20.sort_values('국제협업 논문', ascending=False)
        colors_t = [MJ_COLOR if v else BASE_COLOR for v in top20['is_mj']]
        fig2 = go.Figure(go.Bar(
            x=top20['대학명'], y=top20['국제협업 논문'],
            marker_color=colors_t,
            hovertemplate="<b>%{x}</b><br>협업률: %{y:.4f}<extra></extra>",
        ))
        fig2.add_hline(y=avg_rate, line_dash='dash', line_color='#718096',
                       annotation_text=f'전체 평균({avg_rate:.3f})',
                       annotation_font=dict(size=10))
        fig2.update_layout(**LAYOUT, height=380,
            xaxis=dict(title='', tickangle=-35, tickfont=dict(size=10)),
            yaxis=dict(title='국제협업률', gridcolor='#F0F0F0'),
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── 심층 ③: 인문사회 ──
    elif sel == "③ 인문사회 국내논문당 피인용":
        st.markdown("<div class='section-header'>분야별 논문 구성 비율 (명지대)</div>", unsafe_allow_html=True)

        c1, c2 = st.columns([1, 2])
        with c1:
            labels = ['인문', '사회', '체육']
            vals   = [float(mj3['인문_논문수']), float(mj3['사회_논문수']), float(mj3['체육_논문수'])]
            fig_pie = go.Figure(go.Pie(
                labels=labels, values=vals,
                marker_colors=['#0D2B5E', '#93B8E0', '#E8392A'],
                hole=0.45, textinfo='label+percent',
                hovertemplate="%{label}: %{value:,}편 (%{percent})<extra></extra>",
            ))
            fig_pie.update_layout(**LAYOUT, height=280, showlegend=False,
                                  title=dict(text=f"총 {int(mj3['총_논문수']):,}편", x=0.5, font=dict(size=13)))
            st.plotly_chart(fig_pie, use_container_width=True)

        with c2:
            st.markdown("<div class='section-header'>분야별 논문당 평균 피인용 수</div>", unsafe_allow_html=True)
            avg_all_인문 = d3['Avg_인문'].mean()
            avg_all_사회 = d3['Avg_사회'].mean()
            avg_all_체육 = d3['Avg_체육'].mean()

            fig_bar = go.Figure()
            cats = ['인문', '사회', '체육']
            mj_vals  = [float(mj3['Avg_인문']), float(mj3['Avg_사회']), float(mj3['Avg_체육'])]
            avg_vals = [avg_all_인문, avg_all_사회, avg_all_체육]

            fig_bar.add_trace(go.Bar(name='명지대', x=cats, y=mj_vals,
                                     marker_color=MJ_COLOR,
                                     hovertemplate="%{x}: %{y:.3f}<extra>명지대</extra>"))
            fig_bar.add_trace(go.Bar(name='전체 평균', x=cats, y=avg_vals,
                                     marker_color=BASE_COLOR,
                                     hovertemplate="%{x}: %{y:.3f}<extra>전체 평균</extra>"))
            fig_bar.update_layout(**LAYOUT, height=280, barmode='group',
                yaxis=dict(title='논문당 평균 피인용', gridcolor='#F0F0F0'),
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        st.markdown("<div class='section-header'>분야별 Z·Score 기여도</div>", unsafe_allow_html=True)
        score_cols = ['Score_인문', 'Score_사회', 'Score_체육']
        avg_scores = [d3[c].mean() for c in score_cols]
        mj_scores  = [float(mj3[c]) for c in score_cols]

        fig_radar = go.Figure()
        fig_radar.add_trace(go.Scatterpolar(
            r=mj_scores + [mj_scores[0]], theta=['인문', '사회', '체육', '인문'],
            fill='toself', name='명지대',
            line_color=MJ_COLOR, fillcolor='rgba(232,57,42,0.15)',
        ))
        fig_radar.add_trace(go.Scatterpolar(
            r=avg_scores + [avg_scores[0]], theta=['인문', '사회', '체육', '인문'],
            fill='toself', name='전체 평균',
            line_color=BASE_COLOR, fillcolor='rgba(147,184,224,0.15)',
        ))
        fig_radar.update_layout(**LAYOUT, height=380,
            polar=dict(radialaxis=dict(visible=True, gridcolor='#E4E9F2')),
            legend=dict(orientation='h', yanchor='bottom', y=-0.15, xanchor='center', x=0.5),
        )
        st.plotly_chart(fig_radar, use_container_width=True)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("인문 Avg 피인용", f"{float(mj3['Avg_인문']):.3f}", f"평균 {avg_all_인문:.3f}")
        c2.metric("사회 Avg 피인용", f"{float(mj3['Avg_사회']):.3f}", f"평균 {avg_all_사회:.3f}")
        c3.metric("체육 Avg 피인용", f"{float(mj3['Avg_체육']):.3f}", f"평균 {avg_all_체육:.3f}")
        c4.metric("최종 Z값", f"{float(mj3['Z값']):.4f}", f"순위 {int(mj3['산출순위'])}위")

        st.markdown("""
        <div class='footnote'>
            Z기준값: 명지대 제외 53개교 평균 &nbsp;|&nbsp;
            점수 = 인문·사회·체육 분야별 Z점수 합산
        </div>""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# PAGE 4 — 명지대 종합 프로파일
# ════════════════════════════════════════════════════════════
elif page == "🏫 명지대 종합 프로파일":
    st.markdown("<div class='page-title'>명지대학교 종합 프로파일</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>3개 지표 통합 분석 · 2026년 5월 산출 기준</div>", unsafe_allow_html=True)

    # 핵심 요약 카드
    st.markdown("<div class='section-header'>핵심 지표 요약</div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        pct1 = (1 - mj1['산출순위'] / total) * 100
        st.metric("국제학술지 논문당 피인용",
                  f"{int(mj1['산출순위'])}위 / {total}개교",
                  f"상위 {pct1:.0f}% · FWCI {mj1['FWCI']:.3f}")
    with c2:
        pct2 = (1 - mj2['산출순위'] / total) * 100
        st.metric("국제협업논문",
                  f"{int(mj2['산출순위'])}위 / {total}개교",
                  f"상위 {pct2:.0f}% · 협업률 {mj2['국제협업 논문']:.4f}")
    with c3:
        pct3 = (1 - mj3['산출순위'] / total) * 100
        st.metric("인문사회 국내논문당 피인용",
                  f"{int(mj3['산출순위'])}위 / {total}개교",
                  f"상위 {pct3:.0f}% · Z값 {float(mj3['Z값']):.4f}")

    st.markdown("""
    <div class='highlight-box'>
        <b>※ 명지대학교는 중앙일보 2025년 대학평가 선정 53개교에 포함되지 않았습니다.</b><br>
        본 수치는 동일한 산출 기준을 적용하여 자체 산출한 결과이며,
        <b>국제학술지 논문당 피인용 지표에서 54개교 중 8위</b>의 경쟁력을 확인할 수 있습니다.
    </div>
    """, unsafe_allow_html=True)

    # 레이더차트 — 3개 지표 백분위
    st.markdown("<div class='section-header'>3개 지표 백분위 레이더</div>", unsafe_allow_html=True)

    pct_r1 = (1 - (mj1['산출순위'] - 1) / total) * 100
    pct_r2 = (1 - (mj2['산출순위'] - 1) / total) * 100
    pct_r3 = (1 - (mj3['산출순위'] - 1) / total) * 100
    avg_pct = [50, 50, 50]  # 평균 = 50백분위

    c1, c2 = st.columns([1.2, 1])
    with c1:
        fig_r = go.Figure()
        fig_r.add_trace(go.Scatterpolar(
            r=[pct_r1, pct_r2, pct_r3, pct_r1],
            theta=['국제학술지\n피인용', '국제협업논문', '인문사회\n피인용', '국제학술지\n피인용'],
            fill='toself', name='명지대',
            line_color=MJ_COLOR, fillcolor='rgba(232,57,42,0.2)',
            hovertemplate="%{theta}: 상위 %{r:.1f}%<extra>명지대</extra>",
        ))
        fig_r.add_trace(go.Scatterpolar(
            r=avg_pct + [avg_pct[0]],
            theta=['국제학술지\n피인용', '국제협업논문', '인문사회\n피인용', '국제학술지\n피인용'],
            fill='toself', name='중위 기준(50%)',
            line=dict(color=BASE_COLOR, dash='dash'),
            fillcolor='rgba(147,184,224,0.1)',
        ))
        fig_r.update_layout(**LAYOUT, height=400,
            polar=dict(radialaxis=dict(range=[0, 100], ticksuffix='%', gridcolor='#E4E9F2')),
            legend=dict(orientation='h', yanchor='bottom', y=-0.2, xanchor='center', x=0.5),
        )
        st.plotly_chart(fig_r, use_container_width=True)

    with c2:
        st.markdown("<div class='section-header'>강점 · 약점 분석</div>", unsafe_allow_html=True)

        strengths, weaknesses = [], []
        if pct_r1 >= 70:
            strengths.append(f"국제학술지 피인용 <b>상위 {pct_r1:.0f}%</b> (8위) — 소규모 대비 연구 질적 영향력 우수")
        else:
            weaknesses.append(f"국제학술지 피인용 상위 {pct_r1:.0f}% ({int(mj1['산출순위'])}위)")

        if pct_r2 >= 50:
            strengths.append(f"국제협업 <b>상위 {pct_r2:.0f}%</b> — 국제 연구 네트워크 양호")
        else:
            weaknesses.append(f"국제협업률 상위 {pct_r2:.0f}% ({int(mj2['산출순위'])}위) — 국제 네트워크 확대 필요")

        if pct_r3 >= 60:
            strengths.append(f"인문사회 피인용 <b>상위 {pct_r3:.0f}%</b> ({int(mj3['산출순위'])}위) — 인문사회 분야 연구 활발")
        else:
            weaknesses.append(f"인문사회 피인용 상위 {pct_r3:.0f}% ({int(mj3['산출순위'])}위)")

        if strengths:
            st.markdown(
                "<div class='good-box'><b>✅ 강점</b><br>" +
                "<br>".join(f"· {s}" for s in strengths) + "</div>",
                unsafe_allow_html=True
            )
        if weaknesses:
            st.markdown(
                "<div class='warn-box'><b>⚠️ 개선 필요</b><br>" +
                "<br>".join(f"· {w}" for w in weaknesses) + "</div>",
                unsafe_allow_html=True
            )

    # 세부 수치 테이블
    st.markdown("<div class='section-header'>지표별 세부 수치</div>", unsafe_allow_html=True)
    summary = pd.DataFrame([
        {
            "지표": "국제학술지 논문당 피인용",
            "산출값": f"{mj1['Average of FWCI times JoongAng Ratio']:.4f}",
            "FWCI": f"{mj1['FWCI']:.3f}",
            "논문수": f"{int(mj1['Publications']):,}편",
            "산출순위": f"{int(mj1['산출순위'])}위 / {total}개교",
            "중앙일보 공시": "미선정",
        },
        {
            "지표": "국제협업논문",
            "산출값": f"{mj2['국제협업 논문']:.4f}",
            "FWCI": "-",
            "논문수": f"{int(mj2['Publications']):,}편",
            "산출순위": f"{int(mj2['산출순위'])}위 / {total}개교",
            "중앙일보 공시": "미선정",
        },
        {
            "지표": "인문사회 국내논문당 피인용",
            "산출값": f"{float(mj3['실값']):.4f}",
            "FWCI": "-",
            "논문수": f"{int(mj3['총_논문수']):,}편",
            "산출순위": f"{int(mj3['산출순위'])}위 / {total}개교",
            "중앙일보 공시": "미선정",
        },
    ])
    st.dataframe(summary, use_container_width=True, hide_index=True)

    st.markdown("""
    <div class='footnote'>
        산출 기준: 2026년 5월 &nbsp;|&nbsp;
        중앙일보 공시: 2025년 11월 (53개교 선정, 명지대 미포함) &nbsp;|&nbsp;
        본 수치는 동일 기준 자체 산출값으로 참고용입니다.
    </div>
    """, unsafe_allow_html=True)
