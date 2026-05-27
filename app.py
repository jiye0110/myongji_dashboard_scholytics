
import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="중앙일보 대학평가 교수연구 3개 지표 자체 산출 대시보드", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
.stApp { background-color: #F5F7FA; }
[data-testid="stSidebar"] { background-color: #0D2B5E; }
[data-testid="stSidebar"] * { color: #E8EDF5 !important; }
[data-testid="metric-container"] { background: white; border: 1px solid #E4E9F2; border-radius: 12px; padding: 16px 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
[data-testid="stMetricValue"] { font-family: 'IBM Plex Mono', monospace; font-size: 1.8rem !important; color: #0D2B5E !important; font-weight: 600; }
[data-testid="stMetricLabel"] { font-size: 11px !important; color: #6B7A99 !important; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; }
.page-title { font-size: 22px; font-weight: 700; color: #0D2B5E; margin-bottom: 4px; }
.page-subtitle { font-size: 13px; color: #6B7A99; margin-bottom: 24px; }
.section-header { font-size: 11px; font-weight: 700; color: #6B7A99; text-transform: uppercase; letter-spacing: 0.1em; margin: 28px 0 12px; padding-bottom: 8px; border-bottom: 1px solid #E4E9F2; }
.card { background: white; border-radius: 12px; padding: 18px 22px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); margin-bottom: 12px; }
.card-title { font-size: 13px; font-weight: 700; color: #0D2B5E; margin-bottom: 6px; font-family: 'IBM Plex Mono', monospace; }
.card-body { font-size: 12px; color: #4A5568; line-height: 1.7; }
.card-note { font-size: 11px; color: #6B7A99; margin-top: 8px; line-height: 1.6; }
.highlight-box { background: #EEF4FF; border-left: 4px solid #0D2B5E; border-radius: 0 10px 10px 0; padding: 12px 18px; font-size: 13px; color: #1E3A8A; margin-top: 10px; line-height: 1.7; }
.warn-box { background: #FFF7ED; border-left: 4px solid #F97316; border-radius: 0 10px 10px 0; padding: 12px 18px; font-size: 13px; color: #9A3412; margin-top: 10px; line-height: 1.7; }
.good-box { background: #F0FDF4; border-left: 4px solid #16A34A; border-radius: 0 10px 10px 0; padding: 12px 18px; font-size: 13px; color: #166534; margin-top: 10px; line-height: 1.7; }
.info-box { background: #F8F9FA; border: 1px solid #E4E9F2; border-radius: 10px; padding: 14px 18px; font-size: 13px; color: #2D3748; margin-top: 10px; line-height: 1.8; }
.footnote { font-size: 11px; color: #9AA3B2; margin-top: 24px; padding-top: 12px; border-top: 1px solid #E4E9F2; }
</style>
""", unsafe_allow_html=True)

MJ_NAME, MJ_COLOR, BASE_COLOR = "명지대", "#E8392A", "#93B8E0"
DATA_FILE = "research_indicators_260521.xlsx"
REQUIRED_SHEETS = ["명지대 기본 지표 테이블", "국제학술지 논문당 피인용_only articles", "국제협업논문_only articles", "인문사회 국내논문당 피인용_only articles"]
LAYOUT = dict(font=dict(family="Noto Sans KR, sans-serif", size=12, color="#2D3748"), plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=20, r=20, t=40, b=20), hoverlabel=dict(bgcolor="white", font_size=12))
JJ_10 = [
    [1,"교수당 외부연구비",15,"1.363","50/54",False,"진단평가팀"],[2,"교수당 자체연구비",10,"0","54/54",False,"진단평가팀"],[3,"국제학술지 논문당 피인용",20,"0.6907","8/54",True,"연구전략기획팀"],[4,"교수당 국제학술지 논문",10,"0.652","52/54",False,"진단평가팀"],[5,"국제협업 논문",5,"0.2655","39/54",True,"연구전략기획팀"],[6,"인문사회 교수당 국내논문 및 저역서",5,"1.791","41/54",False,"진단평가팀"],[7,"인문사회 국내논문당 피인용",10,"2.5669","22/54",True,"연구전략기획팀"],[8,"과학기술교수당 기술이전수입액",5,"0.447","49/54",False,"진단평가팀"],[9,"기술이전건수당 수입액",5,"0.447","42/54",False,"진단평가팀"],[10,"과학기술교수당 산학협력 수익",10,"5.077","22/54",False,"진단평가팀"]]

def find_excel():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), DATA_FILE)
    return path if os.path.exists(path) else None

def need_cols(df, cols, sheet):
    miss = [c for c in cols if c not in df.columns]
    if miss:
        st.error(f"'{sheet}' 시트에 필요한 컬럼이 없습니다: {', '.join(miss)}")
        st.stop()

def clean_name(s):
    return s.astype(str).str.strip()

@st.cache_data
def load_data(path):
    xl = pd.read_excel(path, sheet_name=None)
    miss = [s for s in REQUIRED_SHEETS if s not in xl]
    if miss:
        st.error(f"엑셀 파일에 필요한 시트가 없습니다: {', '.join(miss)}")
        st.stop()
    d0 = xl[REQUIRED_SHEETS[0]].copy()
    need_cols(d0, ['대학명','Publications','Citations','FWCI','Publications in top journal percentiles','Publications in top citation percentiles'], REQUIRED_SHEETS[0])
    d0['대학명'] = clean_name(d0['대학명'])
    for c in ['Publications','Citations','FWCI','Publications in top journal percentiles','Publications in top citation percentiles']:
        d0[c] = pd.to_numeric(d0[c], errors='coerce')
    d0['is_mj'] = d0['대학명'].eq(MJ_NAME)
    d1 = xl[REQUIRED_SHEETS[1]].copy(); d1.columns = d1.iloc[0]; d1 = d1.iloc[1:].reset_index(drop=True)
    d1['대학명'] = clean_name(d1['대학명']); d1['Average of FWCI times JoongAng Ratio'] = pd.to_numeric(d1['Average of FWCI times JoongAng Ratio'], errors='coerce')
    d1 = d1.sort_values('Average of FWCI times JoongAng Ratio', ascending=False).reset_index(drop=True); d1['산출순위'] = d1.index + 1; d1['is_mj'] = d1['대학명'].eq(MJ_NAME)
    d2 = xl[REQUIRED_SHEETS[2]].copy(); d2.columns = d2.iloc[0]; d2 = d2.iloc[1:].reset_index(drop=True)
    d2['대학명'] = clean_name(d2['대학명'])
    for c in ['Publications','Publications_국제협업','국제협업 논문']:
        d2[c] = pd.to_numeric(d2[c], errors='coerce')
    d2 = d2.sort_values('국제협업 논문', ascending=False).reset_index(drop=True); d2['산출순위'] = d2.index + 1; d2['is_mj'] = d2['대학명'].eq(MJ_NAME)
    d3 = xl[REQUIRED_SHEETS[3]].copy(); d3.columns = d3.iloc[1]; d3 = d3.iloc[2:].reset_index(drop=True); d3 = d3[d3['Id'].notna() & (d3['Id'] != '[기준값]')].reset_index(drop=True)
    d3.columns = ['Id','대학명(영문)','대학명(한글)','인문_논문수','사회_논문수','체육_논문수','총_논문수','P_인문','P_사회','P_체육','Avg_인문','Avg_사회','Avg_체육','Score_인문','Score_사회','Score_체육','실값','Z값','점수','순위','JJ_실값','JJ_Z값','JJ_점수','JJ_순위']
    d3['대학명(한글)'] = clean_name(d3['대학명(한글)'])
    for c in ['인문_논문수','사회_논문수','체육_논문수','총_논문수','Avg_인문','Avg_사회','Avg_체육','Score_인문','Score_사회','Score_체육','실값','Z값','점수','순위']:
        d3[c] = pd.to_numeric(d3[c], errors='coerce')
    d3 = d3.sort_values('점수', ascending=False).reset_index(drop=True); d3['산출순위'] = d3.index + 1; d3['is_mj'] = d3['대학명(한글)'].eq(MJ_NAME)
    return d0, d1, d2, d3

def mj_row(df, col):
    rows = df[df[col].astype(str).str.strip().eq(MJ_NAME)]
    if rows.empty:
        st.error(f"{col} 컬럼에서 '{MJ_NAME}' 데이터를 찾을 수 없습니다. 엑셀의 대학명 표기를 확인하세요.")
        st.stop()
    return rows.iloc[0]

def top_pct(rank, total):
    return float(rank) / total * 100

def pct_score(rank, total):
    return (1 - (float(rank) - 1) / total) * 100

def hbar(df, y, x, is_mj='is_mj', title_x=None, rank_col='산출순위'):
    colors = [MJ_COLOR if v else BASE_COLOR for v in df[is_mj]]
    fig = go.Figure(go.Bar(y=df[y], x=df[x], orientation='h', marker_color=colors, customdata=df[rank_col] if rank_col in df else None, hovertemplate='<b>%{y}</b><br>%{x:.4f}<extra></extra>'))
    fig.update_layout(**LAYOUT, height=min(max(420, len(df)*26), 1200), yaxis=dict(autorange=True, tickfont=dict(size=10)), xaxis=dict(title=title_x or x, gridcolor='#F0F0F0'))
    return fig

excel_path = find_excel()
if excel_path is None:
    st.error(f"❌ 앱 폴더에서 데이터 파일을 찾을 수 없습니다: {DATA_FILE}")
    st.stop()
d0, d1, d2, d3 = load_data(excel_path)
mj0, mj1, mj2, mj3 = mj_row(d0,'대학명'), mj_row(d1,'대학명'), mj_row(d2,'대학명'), mj_row(d3,'대학명(한글)')
total = len(d1)

with st.sidebar:
    st.markdown("### 🎓 명지대학교\n**중앙일보 대학평가**<br>**교수연구 분야 3개 지표 자체 산출 대시보드**", unsafe_allow_html=True)
    st.markdown("---")
    page = st.radio("페이지", ["📋 지표 해설", "📊 기본 연구 관련 데이터 지표", "🏆 중앙일보 교수연구 3개 지표 순위", "🔍 3개 지표 심층 분석", "🏫 명지대 종합 프로파일"], label_visibility="collapsed")
    st.markdown("---")
    st.markdown("<div style='font-size:11px;color:#93A8D4;'>산출: 연구전략기획팀<br>목적: 2025.11 중앙일보 대학평가 미선정에 따른 원인 진단<br>기준: 중앙일보 평가 방식 준용, 2026.5 확보 가능 데이터로 자체 재산출<br>대상: 2025 중앙일보 평가 상위 53개교 + 명지대학교</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:11px;color:#93A8D4;margin-top:8px;'>📂 {os.path.basename(excel_path)}</div>", unsafe_allow_html=True)

def notice():
    st.markdown("""
    <div class='info-box'>중앙일보 대학평가 결과는 2025년 11월에 발표되었으며, 실제 평가는 발표 이전 시점까지 확보된 자료를 바탕으로 산출된 것으로 보입니다. 명지대학교는 해당 평가 대상/결과에 포함되지 않아 공식 점수가 없으므로, 본 대시보드는 동일·유사 기준을 적용해 현재 확보 가능한 데이터로 자체 산출한 <b>진단용 참고자료</b>입니다.</div>
    """, unsafe_allow_html=True)

if page == "📋 지표 해설":
    st.markdown("<div class='page-title'>지표 해설</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>중앙일보 교수연구 분야 10개 지표 및 연구전략기획팀 자체 산출 3개 지표 안내</div>", unsafe_allow_html=True)
    notice()
    rows = [{'NO':a,'지표명':('🔵 '+b if f else b),'가중치':c,'점수(자체 산출)':d,'순위':e,'담당':g} for a,b,c,d,e,f,g in JJ_10]
    df_jj = pd.DataFrame(rows)

    def highlight_jj(row):
        if str(row['지표명']).startswith('🔵'):
            return ['background-color:#DBEAFE;'] * len(row)
        return [''] * len(row)

    st.markdown("<div class='section-header'>중앙일보 대학평가 교수연구 분야 10개 지표</div>", unsafe_allow_html=True)
    st.dataframe(df_jj.style.apply(highlight_jj, axis=1), use_container_width=True, hide_index=True)
    st.markdown("""
    <div class='card-note' style='margin-top:8px;'>
        🔵 연구전략기획팀에서 Scholytics를 활용하여 직접 산출하는 지표 (3·5·7번) &nbsp;|&nbsp;
        3·5·7번 점수는 2026년 5월 기준 자체 산출값으로 중앙일보 공시값과 다를 수 있습니다.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div class='section-header'>3개 지표 산출 방법</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        <div class='card'>
            <div class='card-title'>① 국제학술지 논문당 피인용 &nbsp;<span style='font-size:11px;color:#6B7A99;font-weight:400;'>가중치 20</span></div>
            <div class='card-body'>
                <b>산출 조건</b><br>
                · Scholytics Index (Scopus 기반) · 평가연도 기준 -5년~-2년<br>
                · 자기인용(Self-citation) 및 의심학술지 제외<br><br>
                <b>순위 결정 기준</b><br>
                Average of FWCI times JoongAng Ratio<br>
                = Σ(논문별 FWCI × 중앙일보 가중치) ÷ 전체 논문수
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("""
        <div class='card'>
            <div class='card-title'>② 국제협업논문 &nbsp;<span style='font-size:11px;color:#6B7A99;font-weight:400;'>가중치 5</span></div>
            <div class='card-body'>
                <b>산출 조건</b><br>
                · 발행연도·저널인덱스·자기인용 제외 조건은 지표 ①과 동일<br>
                · 2024년 중앙일보 평가부터 종합 항목으로 신규 추가<br><br>
                <b>순위 결정 기준</b><br>
                국제협업 비중 = 국제협업 문서수 ÷ 전체 문서수
            </div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class='card'>
            <div class='card-title'>③ 인문사회 국내논문당 피인용 &nbsp;<span style='font-size:11px;color:#6B7A99;font-weight:400;'>가중치 10</span></div>
            <div class='card-body'>
                <b>산출 조건</b><br>
                · KCI Premium + KCI + KCI Candidate<br>
                · 인문 / 사회 / 체육 3개 분야 (이공계 제외)<br>
                · 중앙일보 카테고리 적용<br><br>
                <b>산출 방법</b><br>
                분야별 논문당 피인용을 추출한 후, 논문수 비율로 가중한
                Z-score를 합산하여 최종 점수를 산출합니다.
            </div>
            <div class='card-note'>※ Z기준값: 명지대를 제외한 53개교 평균 적용</div>
        </div>
        """, unsafe_allow_html=True)

elif page == "📊 기본 연구 관련 데이터 지표":
    st.markdown("<div class='page-title'>기본 연구 관련 데이터 지표</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>Scholytics 원천 데이터 기반 참고용 지표 · 명지대학교 현황</div>", unsafe_allow_html=True)
    notice()
    metrics = [('Publications','전체 논문 수','편'),('Citations','피인용 수','회'),('FWCI','세계 평균 대비 피인용 비율',''),('Publications in top journal percentiles','상위 저널 게재 논문 수','편'),('Publications in top citation percentiles','상위 피인용 논문 수','편')]
    tbl=[]
    for col, desc, unit in metrics:
        v, avg = mj0[col], d0[col].mean()
        fmt = f"{v:.3f}" if col=='FWCI' else f"{int(v):,}{unit}"
        afmt = f"{avg:.3f}" if col=='FWCI' else f"{int(avg):,}{unit}"
        tbl.append({'지표':col,'설명':desc,'명지대':fmt,'54개교 평균':afmt})
    st.dataframe(pd.DataFrame(tbl), use_container_width=True, hide_index=True)
    sel = st.selectbox('지표 선택', [m[0] for m in metrics])
    plot = d0.sort_values(sel, ascending=True)
    fig = hbar(plot, '대학명', sel, rank_col='대학명')
    avg = d0[sel].mean(); fig.add_vline(x=avg, line_dash='dot', line_color='#718096', annotation_text=f'평균({avg:,.3f})')
    st.plotly_chart(fig, use_container_width=True)

elif page == "🏆 중앙일보 교수연구 3개 지표 순위":
    st.markdown("<div class='page-title'>중앙일보 교수연구 3개 지표 순위</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>54개 대학 자체 산출값 기준 순위 · 명지대 위치 강조</div>", unsafe_allow_html=True)
    notice()
    c1,c2,c3=st.columns(3)
    c1.metric('국제학술지 논문당 피인용', f"{int(mj1['산출순위'])}위 / {total}개교", f"상위 {top_pct(mj1['산출순위'],total):.1f}% 이내 · {mj1['Average of FWCI times JoongAng Ratio']:.4f}", delta_color='off')
    c2.metric('국제협업논문', f"{int(mj2['산출순위'])}위 / {total}개교", f"상위 {top_pct(mj2['산출순위'],total):.1f}% 이내 · {mj2['국제협업 논문']:.4f}", delta_color='off')
    c3.metric('인문사회 국내논문당 피인용', f"{int(mj3['산출순위'])}위 / {total}개교", f"상위 {top_pct(mj3['산출순위'],total):.1f}% 이내 · Z값 {float(mj3['Z값']):.4f}", delta_color='off')
    st.markdown(f"<div class='highlight-box'><b>★ 명지대학교는 중앙일보 2025년 11월 발표 평가 결과에 포함되지 않았으나</b>, 동일·유사 기준 자체 산출 시 국제학술지 논문당 피인용 <b>{int(mj1['산출순위'])}위</b>, 인문사회 국내논문당 피인용 <b>{int(mj3['산출순위'])}위</b> 수준으로 나타납니다.</div>", unsafe_allow_html=True)
    def rank_options(total):
        return {f"전체 ({total}개 대학)": total, "상위 30개 대학": 30, "상위 20개 대학": 20}

    t1,t2,t3=st.tabs(['① 국제학술지 논문당 피인용','② 국제협업논문','③ 인문사회 국내논문당 피인용'])
    with t1:
        opts = rank_options(total)
        n = opts[st.selectbox('표시 대학 수', list(opts.keys()), key='n1')]
        st.plotly_chart(hbar(d1.head(n).sort_values('Average of FWCI times JoongAng Ratio'), '대학명', 'Average of FWCI times JoongAng Ratio'), use_container_width=True)
    with t2:
        opts = rank_options(total)
        n = opts[st.selectbox('표시 대학 수', list(opts.keys()), key='n2')]
        st.plotly_chart(hbar(d2.head(n).sort_values('국제협업 논문'), '대학명', '국제협업 논문'), use_container_width=True)
    with t3:
        opts = rank_options(total)
        n = opts[st.selectbox('표시 대학 수', list(opts.keys()), key='n3')]
        st.plotly_chart(hbar(d3.head(n).sort_values('점수'), '대학명(한글)', '점수'), use_container_width=True)

elif page == "🔍 3개 지표 심층 분석":
    st.markdown("<div class='page-title'>3개 지표 심층 분석</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>각 지표의 세부 구성 요소와 명지대 위치 분석</div>", unsafe_allow_html=True)
    sel = st.radio('분석 지표 선택', ['① 국제학술지 논문당 피인용','② 국제협업논문','③ 인문사회 국내논문당 피인용'], horizontal=True)
    if sel.startswith('①'):
        dm = d1.merge(d0[['대학명','Publications','Citations']], on='대학명', how='left')
        fig = px.scatter(dm, x='Publications', y='Average of FWCI times JoongAng Ratio', size='Citations', color='is_mj', hover_name='대학명', color_discrete_map={True:MJ_COLOR, False:BASE_COLOR})
        fig.add_hline(y=dm['Average of FWCI times JoongAng Ratio'].mean(), line_dash='dot'); fig.add_vline(x=dm['Publications'].mean(), line_dash='dot'); fig.update_layout(**LAYOUT, height=500, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
        d0r=d0.copy(); d0r['상위저널_비율']=d0r['Publications in top journal percentiles']/d0r['Publications']*100; d0r['상위피인용_비율']=d0r['Publications in top citation percentiles']/d0r['Publications']*100
        fig2=px.scatter(d0r,x='상위저널_비율',y='상위피인용_비율',color='is_mj',hover_name='대학명',color_discrete_map={True:MJ_COLOR,False:BASE_COLOR}); fig2.update_layout(**LAYOUT,height=420,showlegend=False); st.plotly_chart(fig2,use_container_width=True)
    elif sel.startswith('②'):
        fig=px.scatter(d2,x='Publications',y='국제협업 논문',size='Publications_국제협업',color='is_mj',hover_name='대학명',color_discrete_map={True:MJ_COLOR,False:BASE_COLOR}); fig.add_hline(y=d2['국제협업 논문'].mean(),line_dash='dot'); fig.add_vline(x=d2['Publications'].mean(),line_dash='dot'); fig.update_layout(**LAYOUT,height=500,showlegend=False); st.plotly_chart(fig,use_container_width=True)
    else:
        labels=['인문','사회','체육']; vals=[float(mj3['인문_논문수']),float(mj3['사회_논문수']),float(mj3['체육_논문수'])]
        c1,c2=st.columns([1,2])
        with c1:
            fig=go.Figure(go.Pie(labels=labels,values=vals,hole=.45,marker_colors=['#0D2B5E','#93B8E0','#E8392A'])); fig.update_layout(**LAYOUT,height=300); st.plotly_chart(fig,use_container_width=True)
        with c2:
            fig=go.Figure(); fig.add_bar(name='명지대',x=labels,y=[float(mj3['Avg_인문']),float(mj3['Avg_사회']),float(mj3['Avg_체육'])],marker_color=MJ_COLOR); fig.add_bar(name='전체 평균',x=labels,y=[d3['Avg_인문'].mean(),d3['Avg_사회'].mean(),d3['Avg_체육'].mean()],marker_color=BASE_COLOR); fig.update_layout(**LAYOUT,height=300,barmode='group'); st.plotly_chart(fig,use_container_width=True)

else:
    st.markdown("<div class='page-title'>명지대 종합 프로파일</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>3개 지표 통합 분석 · 자체 산출 진단용</div>", unsafe_allow_html=True)
    notice()
    c1,c2,c3=st.columns(3)
    c1.metric('국제학술지 논문당 피인용', f"{int(mj1['산출순위'])}위 / {total}개교", f"상위 {top_pct(mj1['산출순위'],total):.1f}% 이내")
    c2.metric('국제협업논문', f"{int(mj2['산출순위'])}위 / {total}개교", f"상위 {top_pct(mj2['산출순위'],total):.1f}% 이내")
    c3.metric('인문사회 국내논문당 피인용', f"{int(mj3['산출순위'])}위 / {total}개교", f"상위 {top_pct(mj3['산출순위'],total):.1f}% 이내")
    r=[pct_score(mj1['산출순위'],total),pct_score(mj2['산출순위'],total),pct_score(mj3['산출순위'],total)]
    fig=go.Figure(); fig.add_trace(go.Scatterpolar(r=r+[r[0]],theta=['국제학술지\n피인용','국제협업논문','인문사회\n피인용','국제학술지\n피인용'],fill='toself',name='명지대',line_color=MJ_COLOR)); fig.add_trace(go.Scatterpolar(r=[50,50,50,50],theta=['국제학술지\n피인용','국제협업논문','인문사회\n피인용','국제학술지\n피인용'],name='중위 기준',line=dict(color=BASE_COLOR,dash='dash'))); fig.update_layout(**LAYOUT,height=420,polar=dict(radialaxis=dict(range=[0,100],ticksuffix='%'))); st.plotly_chart(fig,use_container_width=True)
    summary=pd.DataFrame([{'지표':'국제학술지 논문당 피인용','산출값':f"Avg FWCI×중앙비율 {mj1['Average of FWCI times JoongAng Ratio']:.4f}",'산출순위':f"{int(mj1['산출순위'])}위 / {total}개교",'성격':'공식 점수 없음, 자체 진단값'},{'지표':'국제협업논문','산출값':f"협업률 {mj2['국제협업 논문']:.4f}",'산출순위':f"{int(mj2['산출순위'])}위 / {total}개교",'성격':'공식 점수 없음, 자체 진단값'},{'지표':'인문사회 국내논문당 피인용','산출값':f"실값 {float(mj3['실값']):.4f}",'산출순위':f"{int(mj3['산출순위'])}위 / {total}개교",'성격':'공식 점수 없음, 자체 진단값'}])
    st.dataframe(summary,use_container_width=True,hide_index=True)
    st.download_button('요약표 CSV 다운로드', summary.to_csv(index=False).encode('utf-8-sig'), 'mju_research_indicator_summary.csv', 'text/csv')
