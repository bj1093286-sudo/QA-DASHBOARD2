import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
from collections import Counter
import re
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────
# Page Config
# ─────────────────────────────────────────
st.set_page_config(
    page_title="QA Monitoring Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────
# Shadcn-inspired CSS
# ─────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Geist:wght@300;400;500;600;700&family=Geist+Mono:wght@400;500&display=swap');

:root {
  --background: #09090b;
  --card: #18181b;
  --card-border: #27272a;
  --primary: #3b82f6;
  --primary-light: #60a5fa;
  --muted: #71717a;
  --muted-bg: #27272a;
  --text: #fafafa;
  --text-muted: #a1a1aa;
  --text-dim: #71717a;
  --success: #22c55e;
  --warning: #f59e0b;
  --danger: #ef4444;
  --accent: #6366f1;
  --border-radius: 12px;
}

html, body, [class*="css"] {
    font-family: 'Geist', -apple-system, BlinkMacSystemFont, sans-serif !important;
    background-color: var(--background) !important;
    color: var(--text) !important;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background-color: #111113 !important;
    border-right: 1px solid var(--card-border) !important;
}
section[data-testid="stSidebar"] * {
    color: var(--text-muted) !important;
}
section[data-testid="stSidebar"] .stRadio label {
    font-size: 13px !important;
    font-weight: 500 !important;
    padding: 6px 0 !important;
    cursor: pointer;
    transition: color 0.15s;
}
section[data-testid="stSidebar"] .stRadio label:hover {
    color: var(--text) !important;
}
section[data-testid="stSidebar"] h1, 
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: var(--text) !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}

/* Main content */
.main .block-container {
    padding: 1.5rem 2rem !important;
    max-width: 1400px !important;
}

/* Cards */
.qa-card {
    background: var(--card);
    border: 1px solid var(--card-border);
    border-radius: var(--border-radius);
    padding: 20px 24px;
    margin-bottom: 16px;
}
.qa-card-title {
    font-size: 13px;
    font-weight: 500;
    color: var(--text-muted);
    margin-bottom: 4px;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}
.qa-card-value {
    font-size: 28px;
    font-weight: 700;
    color: var(--text);
    line-height: 1.1;
}
.qa-card-sub {
    font-size: 12px;
    color: var(--text-dim);
    margin-top: 4px;
}
.qa-badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 500;
}
.qa-badge-blue { background: #1d4ed820; color: #60a5fa; border: 1px solid #3b82f640; }
.qa-badge-green { background: #15803d20; color: #4ade80; border: 1px solid #22c55e40; }
.qa-badge-red { background: #b91c1c20; color: #f87171; border: 1px solid #ef444440; }

/* Section header */
.section-header {
    font-size: 20px;
    font-weight: 700;
    color: var(--text);
    margin-bottom: 4px;
}
.section-desc {
    font-size: 13px;
    color: var(--text-dim);
    margin-bottom: 20px;
}

/* Divider */
.qa-divider {
    border: none;
    border-top: 1px solid var(--card-border);
    margin: 20px 0;
}

/* Metric row */
.metric-row {
    display: flex;
    gap: 16px;
    flex-wrap: wrap;
    margin-bottom: 20px;
}

/* Selectbox / multiselect styling */
.stSelectbox > div > div,
.stMultiSelect > div > div {
    background: var(--card) !important;
    border: 1px solid var(--card-border) !important;
    border-radius: 8px !important;
    color: var(--text) !important;
}

/* Table */
.dataframe {
    background: var(--card) !important;
    color: var(--text) !important;
    border: 1px solid var(--card-border) !important;
    border-radius: var(--border-radius) !important;
    font-size: 12px !important;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 1px solid var(--card-border) !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: var(--text-dim) !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    padding: 8px 16px !important;
    border: none !important;
}
.stTabs [aria-selected="true"] {
    color: var(--text) !important;
    border-bottom: 2px solid var(--primary) !important;
}

/* Plotly override */
.js-plotly-plot .plotly .modebar {
    background: transparent !important;
}

/* Scrollbar */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: var(--background); }
::-webkit-scrollbar-thumb { background: var(--card-border); border-radius: 3px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# Plotly theme
# ─────────────────────────────────────────
PLOTLY_LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(24,24,27,0.6)',
    font=dict(family='Geist, sans-serif', color='#a1a1aa', size=12),
    title_font=dict(size=14, color='#fafafa', weight=700),
    xaxis=dict(gridcolor='#27272a', linecolor='#27272a', tickfont=dict(size=11)),
    yaxis=dict(gridcolor='#27272a', linecolor='#27272a', tickfont=dict(size=11)),
    legend=dict(bgcolor='rgba(24,24,27,0.8)', bordercolor='#27272a', borderwidth=1,
                font=dict(size=11, color='#a1a1aa')),
    colorway=['#3b82f6','#6366f1','#22c55e','#f59e0b','#ef4444','#ec4899','#06b6d4','#84cc16'],
    margin=dict(l=40, r=20, t=40, b=40),
    hoverlabel=dict(bgcolor='#18181b', bordercolor='#3b82f6', font=dict(color='#fafafa', size=12)),
)

def apply_layout(fig, title='', height=350):
    fig.update_layout(**PLOTLY_LAYOUT, title=title, height=height)
    return fig

# ─────────────────────────────────────────
# Data Loading
# ─────────────────────────────────────────
@st.cache_data
def load_data(filepath='통합_문서1.xlsx'):
    wb = openpyxl.load_workbook(filepath)
    
    # ── CALL ──────────────────────────────
    call_items = {
        '첫인사': (8, 2.5), '정보확인': (11, 5), '끝인사': (14, 2.5),
        '인사톤': (17, 2.5), '통화종료': (20, 2.5), '음성숙련도': (23, 5),
        '감정연출': (26, 5), '양해': (29, 5), '즉각호응': (32, 5),
        '대기': (35, 5), '언어표현': (38, 5), '경청': (41, 5),
        '문의파악': (44, 5), '맞춤설명': (47, 5), '정확한안내': (50, 10),
        '프로세스': (53, 10), '전산처리': (56, 10), '상담이력': (59, 10)
    }
    
    chat_items = {
        '첫인사': (6, 5), '정보확인': (9, 5), '끝인사': (12, 5),
        '양해': (15, 5), '즉각호응': (18, 5), '대기': (21, 5),
        '언어표현': (24, 10), '가독성': (27, 5), '문의파악': (30, 10),
        '맞춤설명': (33, 10), '정확한안내': (36, 10), '프로세스': (39, 5),
        '전산처리': (42, 10), '상담이력': (45, 10)
    }
    
    board_items = {
        '첫인사': (7, 2.5), '끝인사': (10, 2.5), '언어표현': (13, 10),
        '양해': (16, 10), '가독성': (19, 10), '문의파악': (22, 10),
        '맞춤설명': (25, 10), '정확한안내': (28, 10), '프로세스': (31, 10),
        '전산처리': (34, 10), '상담이력': (37, 15)
    }
    
    def parse_sheet(ws, items_map, agent_col, period_col, inq_cols, total_col, evaluator_col, bonus_col, penalty_col, skip_rows=3):
        records = []
        rows = list(ws.iter_rows(values_only=True))
        for row in rows[skip_rows:]:
            if row[agent_col] is None:
                continue
            rec = {
                '상담사': str(row[agent_col]).strip(),
                '평가기간': str(row[period_col]).strip() if row[period_col] else '',
                '문의유형대': str(row[inq_cols[0]]).strip() if row[inq_cols[0]] else '',
                '문의유형중': str(row[inq_cols[1]]).strip() if row[inq_cols[1]] else '',
                '문의유형소': str(row[inq_cols[2]]).strip() if row[inq_cols[2]] else '',
                '평가자': str(row[evaluator_col]).strip() if row[evaluator_col] else '',
                '가점': _num(row[bonus_col]),
                '감점': _num(row[penalty_col]),
            }
            # Parse TOTAL
            total_raw = row[total_col]
            if isinstance(total_raw, str) and total_raw.startswith('='):
                # Calculate from items
                total = sum(_num(row[c]) for c, _ in items_map.values())
                total += rec['가점'] or 0
                total -= rec['감점'] or 0
            else:
                total = _num(total_raw)
            rec['TOTAL'] = total
            
            for item_name, (col_idx, max_score) in items_map.items():
                score = _num(row[col_idx])
                deduct_reason = str(row[col_idx+1]).strip() if row[col_idx+1] else ''
                no_deduct_comment = str(row[col_idx+2]).strip() if row[col_idx+2] else ''
                rec[f'{item_name}_점수'] = score
                rec[f'{item_name}_최대'] = max_score
                rec[f'{item_name}_감점사유'] = deduct_reason if deduct_reason not in ('None','') else ''
                rec[f'{item_name}_미차감'] = no_deduct_comment if no_deduct_comment not in ('None','') else ''
            records.append(rec)
        return pd.DataFrame(records)
    
    def _num(v):
        if v is None: return None
        if isinstance(v, (int, float)): return float(v)
        try: return float(str(v).replace(',',''))
        except: return None
    
    ws_call = wb['CALL_누적']
    ws_chat = wb['CHAT_누적']
    ws_board = wb['게시판_누적']
    
    df_call = parse_sheet(ws_call, call_items,
        agent_col=3, period_col=1, inq_cols=[5,6,7],
        total_col=64, evaluator_col=65, bonus_col=62, penalty_col=63)
    df_call['채널'] = 'CALL'
    
    df_chat = parse_sheet(ws_chat, chat_items,
        agent_col=1, period_col=0, inq_cols=[3,4,5],
        total_col=50, evaluator_col=51, bonus_col=48, penalty_col=49)
    df_chat['채널'] = 'CHAT'
    
    df_board = parse_sheet(ws_board, board_items,
        agent_col=2, period_col=1, inq_cols=[4,5,6],
        total_col=42, evaluator_col=43, bonus_col=40, penalty_col=41)
    df_board['채널'] = '게시판'
    
    df_all = pd.concat([df_call, df_chat, df_board], ignore_index=True)
    
    # Period sort key
    def period_sort(p):
        import re
        m = re.search(r'(\d+)년\s*(\d+)월(\d+)차', p)
        if m:
            return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
        return (9999, 0, 0)
    
    df_all['기간_정렬'] = df_all['평가기간'].apply(period_sort)
    df_all = df_all.sort_values('기간_정렬')
    
    return df_call, df_chat, df_board, df_all, call_items, chat_items, board_items

# ─────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────
def get_score_cols(df):
    return [c for c in df.columns if c.endswith('_점수')]

def get_item_names(score_cols):
    return [c.replace('_점수','') for c in score_cols]

def metric_card(title, value, sub='', color='blue'):
    badge_map = {'blue':'qa-badge-blue','green':'qa-badge-green','red':'qa-badge-red'}
    badge = badge_map.get(color,'qa-badge-blue')
    st.markdown(f"""
    <div class="qa-card">
        <div class="qa-card-title">{title}</div>
        <div class="qa-card-value">{value}</div>
        <div class="qa-card-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

def section_header(title, desc=''):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)
    if desc:
        st.markdown(f'<div class="section-desc">{desc}</div>', unsafe_allow_html=True)

def filter_df(df, periods, channels, agents):
    out = df.copy()
    if periods:
        out = out[out['평가기간'].isin(periods)]
    if channels:
        out = out[out['채널'].isin(channels)]
    if agents:
        out = out[out['상담사'].isin(agents)]
    return out

# ─────────────────────────────────────────
# Load
# ─────────────────────────────────────────
try:
    df_call, df_chat, df_board, df_all, call_items, chat_items, board_items = load_data('통합_문서1.xlsx')
except Exception as e:
    st.error(f"❌ 파일 로딩 오류: {e}\n\n`통합_문서1.xlsx` 파일을 앱과 같은 폴더에 두세요.")
    st.stop()

all_periods = sorted(df_all['평가기간'].unique(), key=lambda p: df_all[df_all['평가기간']==p]['기간_정렬'].iloc[0] if len(df_all[df_all['평가기간']==p]) else (9999,0,0))
all_channels = df_all['채널'].unique().tolist()
all_agents = sorted(df_all['상담사'].unique().tolist())

# ─────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 QA Dashboard")
    st.markdown("<hr style='border-color:#27272a;margin:8px 0 16px'>", unsafe_allow_html=True)
    
    page = st.radio("Navigation", [
        "🏠 개요",
        "📈 점수 트렌드",
        "🗂 문의유형 분석",
        "📋 항목별 분석",
        "⚠️ 감점 분석",
        "👤 상담사 종합평가",
        "✅ 이행률 분석",
        "🔍 평가자 분석",
        "💬 코멘트 NLP",
    ], label_visibility='collapsed')
    
    st.markdown("<hr style='border-color:#27272a;margin:16px 0 12px'>", unsafe_allow_html=True)
    st.markdown("### 필터")
    
    sel_period = st.multiselect("평가기간", all_periods, default=all_periods)
    sel_channel = st.multiselect("채널", all_channels, default=all_channels)
    sel_agent = st.multiselect("상담사", all_agents, default=all_agents)
    
    st.markdown("<hr style='border-color:#27272a;margin:12px 0'>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:11px;color:#52525b'>총 {len(df_all)}건 로드됨</div>", unsafe_allow_html=True)

df = filter_df(df_all, sel_period, sel_channel, sel_agent)

# ─────────────────────────────────────────
# Page: 개요
# ─────────────────────────────────────────
if page == "🏠 개요":
    section_header("QA 모니터링 대시보드", "상담 품질 평가 현황 종합 요약")
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("총 평가건수", f"{len(df):,}건", f"전체 채널 합산")
    with c2: metric_card("평균 점수", f"{df['TOTAL'].dropna().mean():.1f}점" if len(df) else '-', "전체 평균")
    with c3: metric_card("상담사 수", f"{df['상담사'].nunique()}명", "활성 상담사")
    with c4: metric_card("평가 기간 수", f"{df['평가기간'].nunique()}개", "")
    
    col1, col2 = st.columns([3, 2])
    
    with col1:
        # Score trend by period & channel
        grp = df.groupby(['평가기간','채널'])['TOTAL'].mean().reset_index()
        grp = grp.merge(df[['평가기간','기간_정렬']].drop_duplicates(), on='평가기간')
        grp = grp.sort_values('기간_정렬')
        fig = px.line(grp, x='평가기간', y='TOTAL', color='채널',
                      markers=True, labels={'TOTAL':'평균 점수', '평가기간':'평가기간'})
        apply_layout(fig, '📈 채널별 평균 점수 추이', 300)
        fig.update_traces(line=dict(width=2.5), marker=dict(size=6))
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Channel distribution
        ch_cnt = df['채널'].value_counts().reset_index()
        ch_cnt.columns = ['채널','건수']
        fig2 = px.pie(ch_cnt, names='채널', values='건수',
                      hole=0.65, color_discrete_sequence=['#3b82f6','#6366f1','#22c55e'])
        apply_layout(fig2, '채널별 평가 비중', 300)
        fig2.update_traces(textfont=dict(size=12, color='#fafafa'))
        st.plotly_chart(fig2, use_container_width=True)
    
    # Agent summary table
    section_header("상담사별 요약")
    agent_grp = df.groupby(['상담사','채널'])['TOTAL'].agg(['mean','count']).reset_index()
    agent_grp.columns = ['상담사','채널','평균점수','평가건수']
    agent_grp['평균점수'] = agent_grp['평균점수'].round(1)
    
    pivot = agent_grp.pivot_table(index='상담사', columns='채널', values='평균점수').round(1)
    pivot['전체평균'] = df.groupby('상담사')['TOTAL'].mean().round(1)
    pivot['총평가건'] = df.groupby('상담사')['TOTAL'].count()
    st.dataframe(pivot.sort_values('전체평균', ascending=False), use_container_width=True)

# ─────────────────────────────────────────
# Page: 점수 트렌드
# ─────────────────────────────────────────
elif page == "📈 점수 트렌드":
    section_header("점수 트렌드", "기간별·상담사별 점수 변화 분석")
    
    tab1, tab2, tab3 = st.tabs(["기간별 추이", "상담사별 비교", "상자 분포"])
    
    with tab1:
        grp = df.groupby(['평가기간','상담사'])['TOTAL'].mean().reset_index()
        grp = grp.merge(df[['평가기간','기간_정렬']].drop_duplicates(), on='평가기간').sort_values('기간_정렬')
        fig = px.line(grp, x='평가기간', y='TOTAL', color='상담사',
                      markers=True, labels={'TOTAL':'평균 점수'})
        apply_layout(fig, '상담사별 점수 추이', 420)
        fig.update_traces(line=dict(width=2), marker=dict(size=5))
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        agent_avg = df.groupby('상담사')['TOTAL'].mean().sort_values(ascending=True).reset_index()
        fig = px.bar(agent_avg, x='TOTAL', y='상담사', orientation='h',
                     labels={'TOTAL':'평균 점수','상담사':'상담사'},
                     color='TOTAL', color_continuous_scale=['#ef4444','#f59e0b','#22c55e'])
        apply_layout(fig, '상담사별 평균 점수', max(300, len(agent_avg)*30))
        fig.update_coloraxes(showscale=False)
        # Add avg line
        overall_avg = df['TOTAL'].mean()
        fig.add_vline(x=overall_avg, line_dash='dash', line_color='#60a5fa',
                      annotation_text=f'평균 {overall_avg:.1f}', annotation_font_color='#60a5fa')
        st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        fig = px.box(df, x='상담사', y='TOTAL', color='채널',
                     labels={'TOTAL':'점수','상담사':'상담사'})
        apply_layout(fig, '상담사별 점수 분포 (Box Plot)', 400)
        st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────
# Page: 문의유형 분석
# ─────────────────────────────────────────
elif page == "🗂 문의유형 분석":
    section_header("문의유형 × 상담사 분석", "문의 유형별 점수 히트맵 및 드릴다운")
    
    tab1, tab2 = st.tabs(["유형 × 상담사 히트맵", "유형별 건수/점수"])
    
    with tab1:
        pivot = df.groupby(['문의유형대','상담사'])['TOTAL'].mean().unstack(fill_value=np.nan).round(1)
        fig = go.Figure(data=go.Heatmap(
            z=pivot.values,
            x=pivot.columns.tolist(),
            y=pivot.index.tolist(),
            colorscale=[[0,'#ef4444'],[0.5,'#f59e0b'],[1,'#22c55e']],
            text=pivot.values.round(1),
            texttemplate='%{text}',
            textfont=dict(size=11, color='white'),
            hoverongaps=False,
            colorbar=dict(tickfont=dict(color='#a1a1aa'))
        ))
        apply_layout(fig, '문의유형(대) × 상담사 평균점수 히트맵', max(300, len(pivot)*35))
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            cnt = df['문의유형대'].value_counts().reset_index()
            cnt.columns = ['유형','건수']
            fig = px.bar(cnt, x='유형', y='건수', color='건수',
                         color_continuous_scale='Blues')
            apply_layout(fig, '문의유형(대)별 평가 건수', 300)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            avg = df.groupby('문의유형대')['TOTAL'].mean().reset_index()
            avg.columns = ['유형','평균점수']
            fig = px.bar(avg, x='유형', y='평균점수', color='평균점수',
                         color_continuous_scale=['#ef4444','#22c55e'])
            apply_layout(fig, '문의유형(대)별 평균 점수', 300)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        
        # 중분류
        avg2 = df.groupby(['문의유형대','문의유형중'])['TOTAL'].mean().reset_index()
        fig = px.sunburst(df, path=['문의유형대','문의유형중'], values='TOTAL',
                          color='TOTAL', color_continuous_scale=['#ef4444','#22c55e'])
        apply_layout(fig, '문의유형 계층 구조 (선버스트)', 400)
        st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────
# Page: 항목별 분석
# ─────────────────────────────────────────
elif page == "📋 항목별 분석":
    section_header("항목별 점수 분석", "QA 평가 항목별 시계열 및 상담사 비교")
    
    # Collect all score items present in filtered df
    score_cols = get_score_cols(df)
    items = get_item_names(score_cols)
    
    tab1, tab2, tab3 = st.tabs(["항목별 평균점수 (전체)", "항목 × 기간 추이", "상담사 × 항목 히트맵"])
    
    with tab1:
        # Compute avg score and max score per item
        rows = []
        for item in items:
            sc = df[f'{item}_점수'].dropna()
            mx = df[f'{item}_최대'].dropna()
            if len(sc) == 0: continue
            avg_s = sc.mean()
            avg_m = mx.mean() if len(mx) else 1
            rows.append({'항목': item, '평균점수': avg_s, '최대점수': avg_m, '달성률': avg_s/avg_m*100})
        item_df = pd.DataFrame(rows).sort_values('달성률')
        
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=item_df['달성률'], y=item_df['항목'], orientation='h',
            marker_color=item_df['달성률'].apply(lambda x: '#22c55e' if x>=90 else '#f59e0b' if x>=75 else '#ef4444'),
            text=item_df['달성률'].round(1).astype(str)+'%',
            textposition='inside', textfont=dict(color='white', size=11),
            name='달성률'
        ))
        apply_layout(fig, '항목별 평균 달성률 (%)', max(300, len(item_df)*28))
        fig.update_layout(xaxis=dict(range=[0,105], title='달성률 (%)'))
        st.plotly_chart(fig, use_container_width=True)
        
        st.dataframe(item_df[['항목','평균점수','최대점수','달성률']].assign(
            평균점수=lambda x: x['평균점수'].round(2),
            달성률=lambda x: x['달성률'].round(1)
        ).sort_values('달성률'), use_container_width=True)
    
    with tab2:
        sel_items = st.multiselect("항목 선택", items, default=items[:5])
        if sel_items:
            rows = []
            for item in sel_items:
                grp = df.groupby(['평가기간','기간_정렬'])[f'{item}_점수'].mean().reset_index()
                grp = grp.sort_values('기간_정렬')
                grp['항목'] = item
                grp.rename(columns={f'{item}_점수':'점수'}, inplace=True)
                rows.append(grp)
            trend_df = pd.concat(rows)
            fig = px.line(trend_df, x='평가기간', y='점수', color='항목', markers=True)
            apply_layout(fig, '항목별 점수 추이', 380)
            fig.update_traces(line=dict(width=2), marker=dict(size=5))
            st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        agent_item = {}
        for item in items:
            s = df.groupby('상담사')[f'{item}_점수'].mean()
            m = df.groupby('상담사')[f'{item}_최대'].mean()
            agent_item[item] = (s/m*100).round(1)
        
        heat_df = pd.DataFrame(agent_item)
        fig = go.Figure(data=go.Heatmap(
            z=heat_df.values,
            x=heat_df.columns.tolist(),
            y=heat_df.index.tolist(),
            colorscale=[[0,'#ef4444'],[0.5,'#f59e0b'],[1,'#22c55e']],
            zmin=0, zmax=100,
            text=heat_df.values.round(0).astype(str),
            texttemplate='%{text}%',
            textfont=dict(size=10),
            colorbar=dict(title='달성률%', tickfont=dict(color='#a1a1aa'))
        ))
        apply_layout(fig, '상담사 × 항목 달성률 히트맵', max(350, len(heat_df)*35))
        st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────
# Page: 감점 분석
# ─────────────────────────────────────────
elif page == "⚠️ 감점 분석":
    section_header("감점 분석", "감점 사유 및 패턴 분석 (정량 + 정성)")
    
    score_cols = get_score_cols(df)
    items = get_item_names(score_cols)
    
    # Collect all deduction events
    deduct_records = []
    for item in items:
        sc_col = f'{item}_점수'
        mx_col = f'{item}_최대'
        reason_col = f'{item}_감점사유'
        if reason_col not in df.columns: continue
        sub = df[[sc_col, mx_col, reason_col, '상담사','평가기간','채널']].copy()
        sub.columns = ['점수','최대','감점사유','상담사','평가기간','채널']
        sub['항목'] = item
        sub['감점여부'] = (sub['점수'] < sub['최대']) & sub['점수'].notna()
        sub = sub[sub['감점여부'] & (sub['감점사유'] != '')]
        deduct_records.append(sub)
    
    deduct_df = pd.concat(deduct_records, ignore_index=True) if deduct_records else pd.DataFrame()
    
    tab1, tab2, tab3 = st.tabs(["항목별 감점 빈도", "상담사별 감점 패턴", "감점사유 텍스트 분석"])
    
    with tab1:
        if len(deduct_df):
            item_deduct = deduct_df['항목'].value_counts().reset_index()
            item_deduct.columns = ['항목','감점건수']
            fig = px.bar(item_deduct, x='항목', y='감점건수',
                         color='감점건수', color_continuous_scale='Reds')
            apply_layout(fig, '항목별 감점 발생 건수', 340)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
            
            # Deduction rate per item
            deduct_rate = []
            for item in items:
                total = df[f'{item}_점수'].dropna().shape[0]
                deducted = deduct_df[deduct_df['항목']==item].shape[0]
                if total: deduct_rate.append({'항목': item, '감점율': deducted/total*100})
            dr_df = pd.DataFrame(deduct_rate).sort_values('감점율', ascending=False)
            
            fig2 = px.bar(dr_df, x='항목', y='감점율', color='감점율',
                          color_continuous_scale=['#22c55e','#f59e0b','#ef4444'],
                          labels={'감점율':'감점 발생률 (%)'})
            apply_layout(fig2, '항목별 감점 발생률 (%)', 300)
            fig2.update_coloraxes(showscale=False)
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("감점 데이터가 없습니다.")
    
    with tab2:
        if len(deduct_df):
            agent_deduct = deduct_df.groupby(['상담사','항목']).size().unstack(fill_value=0)
            fig = go.Figure(data=go.Heatmap(
                z=agent_deduct.values,
                x=agent_deduct.columns.tolist(),
                y=agent_deduct.index.tolist(),
                colorscale=[[0,'#18181b'],[1,'#ef4444']],
                text=agent_deduct.values,
                texttemplate='%{text}',
                textfont=dict(size=11),
            ))
            apply_layout(fig, '상담사 × 항목 감점 횟수', max(300, len(agent_deduct)*35))
            st.plotly_chart(fig, use_container_width=True)
            
            # Top deducted agents
            agent_cnt = deduct_df['상담사'].value_counts().reset_index()
            agent_cnt.columns = ['상담사','감점건수']
            st.markdown("**상담사별 총 감점 발생 건수**")
            st.dataframe(agent_cnt, use_container_width=True)
        else:
            st.info("감점 데이터가 없습니다.")
    
    with tab3:
        if len(deduct_df):
            all_reasons = deduct_df['감점사유'].dropna().tolist()
            
            # Keyword frequency (simple split)
            from collections import Counter
            words = []
            for r in all_reasons:
                # Basic Korean tokenization - split on spaces and punctuation
                tokens = re.sub(r'[^\w\s가-힣]', ' ', r).split()
                tokens = [t for t in tokens if len(t) >= 2]
                words.extend(tokens)
            
            word_freq = Counter(words).most_common(30)
            wf_df = pd.DataFrame(word_freq, columns=['단어','빈도'])
            
            fig = px.bar(wf_df, x='빈도', y='단어', orientation='h',
                         color='빈도', color_continuous_scale='Blues')
            apply_layout(fig, '감점 사유 핵심 키워드 TOP 30', 500)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
            
            # Show raw comments
            st.markdown("**감점 사유 원본 코멘트**")
            show_df = deduct_df[['상담사','채널','항목','점수','최대','감점사유']].copy()
            show_df = show_df[show_df['감점사유'] != '']
            st.dataframe(show_df.reset_index(drop=True), use_container_width=True)
        else:
            st.info("감점 사유 데이터가 없습니다.")

# ─────────────────────────────────────────
# Page: 상담사 종합평가
# ─────────────────────────────────────────
elif page == "👤 상담사 종합평가":
    section_header("상담사 종합 평가", "레이더 차트로 보는 강점/약점 프로파일")
    
    agent_sel = st.selectbox("상담사 선택", all_agents)
    agent_df = df_all[df_all['상담사'] == agent_sel]
    
    c1, c2, c3 = st.columns(3)
    avg_total = agent_df['TOTAL'].mean()
    eval_cnt = len(agent_df)
    channels_used = ', '.join(agent_df['채널'].unique())
    with c1: metric_card("평균 점수", f"{avg_total:.1f}" if avg_total else '-', channels_used)
    with c2: metric_card("평가 건수", f"{eval_cnt}건", "전체 기간")
    with c3: metric_card("활동 채널", channels_used, "")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Radar chart
        score_cols = get_score_cols(agent_df)
        items = get_item_names(score_cols)
        radar_data = []
        for item in items:
            s = agent_df[f'{item}_점수'].dropna().mean()
            m = agent_df[f'{item}_최대'].dropna().mean()
            if m and m > 0:
                radar_data.append({'항목': item, '달성률': s/m*100 if s else 0})
        
        if radar_data:
            rd = pd.DataFrame(radar_data)
            categories = rd['항목'].tolist()
            values = rd['달성률'].tolist()
            values += [values[0]]  # close polygon
            categories += [categories[0]]
            
            fig = go.Figure()
            fig.add_trace(go.Scatterpolar(
                r=values, theta=categories, fill='toself',
                fillcolor='rgba(59,130,246,0.15)',
                line=dict(color='#3b82f6', width=2),
                marker=dict(size=5, color='#3b82f6'),
                name=agent_sel
            ))
            # Add avg line
            overall_radar = []
            for item in items:
                s = df[f'{item}_점수'].dropna().mean()
                m = df[f'{item}_최대'].dropna().mean()
                overall_radar.append(s/m*100 if (m and m>0 and s) else 0)
            overall_radar += [overall_radar[0]]
            fig.add_trace(go.Scatterpolar(
                r=overall_radar, theta=categories, fill='toself',
                fillcolor='rgba(99,102,241,0.08)',
                line=dict(color='#6366f1', width=1.5, dash='dash'),
                marker=dict(size=4, color='#6366f1'),
                name='전체평균'
            ))
            apply_layout(fig, f'{agent_sel} 역량 레이더', 420)
            fig.update_layout(polar=dict(
                bgcolor='rgba(24,24,27,0.6)',
                radialaxis=dict(visible=True, range=[0,100], tickfont=dict(size=9, color='#71717a'),
                                gridcolor='#27272a', linecolor='#27272a'),
                angularaxis=dict(tickfont=dict(size=10, color='#a1a1aa'), gridcolor='#27272a')
            ))
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Score trend over time
        trend = agent_df.groupby(['평가기간','기간_정렬'])['TOTAL'].mean().reset_index().sort_values('기간_정렬')
        fig2 = px.bar(trend, x='평가기간', y='TOTAL', labels={'TOTAL':'평균 점수'},
                      color='TOTAL', color_continuous_scale=['#ef4444','#22c55e'])
        apply_layout(fig2, f'{agent_sel} 기간별 점수', 300)
        fig2.update_coloraxes(showscale=False)
        st.plotly_chart(fig2, use_container_width=True)
        
        # Channel breakdown
        ch = agent_df.groupby('채널')['TOTAL'].agg(['mean','count']).reset_index()
        ch.columns = ['채널','평균점수','건수']
        st.dataframe(ch, use_container_width=True)
    
    # Deduction summary for agent
    st.markdown("<hr class='qa-divider'>", unsafe_allow_html=True)
    st.markdown("**감점 발생 항목**")
    score_cols_a = get_score_cols(agent_df)
    items_a = get_item_names(score_cols_a)
    ded_rows = []
    for item in items_a:
        sc = agent_df[f'{item}_점수'].dropna()
        mx = agent_df[f'{item}_최대'].dropna()
        reasons = agent_df[f'{item}_감점사유'].dropna()
        reasons = reasons[reasons != '']
        if len(sc) == 0: continue
        avg_s = sc.mean()
        avg_m = mx.mean() if len(mx) else None
        ded_rows.append({
            '항목': item, '평균점수': round(avg_s,2), '최대점수': avg_m,
            '달성률': f"{avg_s/avg_m*100:.1f}%" if avg_m else '-',
            '감점건수': (sc < mx).sum() if len(mx)==len(sc) else '-',
            '주요감점사유': reasons.iloc[0][:60] if len(reasons) else ''
        })
    ded_agent_df = pd.DataFrame(ded_rows)
    st.dataframe(ded_agent_df, use_container_width=True)

# ─────────────────────────────────────────
# Page: 이행률 분석
# ─────────────────────────────────────────
elif page == "✅ 이행률 분석":
    section_header("이행률 분석", "항목별·상담사별·기간별 이행률 (감점 미발생률)")
    
    # 이행률 = 만점 받은 비율 (= 1 - 감점발생률)
    score_cols = get_score_cols(df)
    items = get_item_names(score_cols)
    
    tab1, tab2, tab3 = st.tabs(["항목별 이행률", "상담사별 이행률", "기간별 추이"])
    
    def compliance(scores, maxes):
        """Returns % of evaluations where full marks achieved"""
        valid = [(s, m) for s, m in zip(scores, maxes) if s is not None and m is not None and m > 0]
        if not valid: return None
        full = sum(1 for s, m in valid if s >= m)
        return full / len(valid) * 100
    
    with tab1:
        rows = []
        for item in items:
            sc = df[f'{item}_점수'].tolist()
            mx = df[f'{item}_최대'].tolist()
            c = compliance(sc, mx)
            if c is not None:
                rows.append({'항목': item, '이행률': round(c, 1)})
        comp_df = pd.DataFrame(rows).sort_values('이행률')
        
        fig = go.Figure(go.Bar(
            x=comp_df['이행률'], y=comp_df['항목'], orientation='h',
            marker_color=comp_df['이행률'].apply(lambda x: '#22c55e' if x>=90 else '#f59e0b' if x>=75 else '#ef4444'),
            text=comp_df['이행률'].astype(str)+'%',
            textposition='outside',
            textfont=dict(size=11, color='#a1a1aa')
        ))
        apply_layout(fig, '항목별 이행률 (만점 달성 비율)', max(300, len(comp_df)*28))
        fig.update_layout(xaxis=dict(range=[0,108]))
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        agent_comp = {}
        for agent in df['상담사'].unique():
            adf = df[df['상담사']==agent]
            vals = []
            for item in items:
                sc = adf[f'{item}_점수'].tolist()
                mx = adf[f'{item}_최대'].tolist()
                c = compliance(sc, mx)
                if c is not None: vals.append(c)
            if vals: agent_comp[agent] = round(np.mean(vals),1)
        
        ac_df = pd.DataFrame(list(agent_comp.items()), columns=['상담사','이행률']).sort_values('이행률')
        fig = px.bar(ac_df, x='이행률', y='상담사', orientation='h',
                     color='이행률', color_continuous_scale=['#ef4444','#22c55e'],
                     labels={'이행률':'평균 이행률 (%)'})
        apply_layout(fig, '상담사별 평균 이행률', max(300, len(ac_df)*30))
        fig.update_coloraxes(showscale=False)
        avg_c = ac_df['이행률'].mean()
        fig.add_vline(x=avg_c, line_dash='dash', line_color='#60a5fa',
                      annotation_text=f'평균 {avg_c:.1f}%', annotation_font_color='#60a5fa')
        st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        period_comp = {}
        for period in df['평가기간'].unique():
            pdf = df[df['평가기간']==period]
            sort_key = df[df['평가기간']==period]['기간_정렬'].iloc[0]
            vals = []
            for item in items:
                sc = pdf[f'{item}_점수'].tolist()
                mx = pdf[f'{item}_최대'].tolist()
                c = compliance(sc, mx)
                if c is not None: vals.append(c)
            if vals: period_comp[period] = (round(np.mean(vals),1), sort_key)
        
        pc_df = pd.DataFrame([(k,v[0],v[1]) for k,v in period_comp.items()],
                              columns=['기간','이행률','정렬']).sort_values('정렬')
        fig = px.line(pc_df, x='기간', y='이행률', markers=True, labels={'이행률':'평균 이행률 (%)'})
        apply_layout(fig, '기간별 전체 이행률 추이', 320)
        fig.update_traces(line=dict(width=2.5, color='#3b82f6'), marker=dict(size=7, color='#60a5fa'))
        fig.add_hline(y=90, line_dash='dot', line_color='#22c55e', annotation_text='목표 90%',
                      annotation_font_color='#22c55e')
        st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────
# Page: 평가자 분석
# ─────────────────────────────────────────
elif page == "🔍 평가자 분석":
    section_header("평가자 행동 분석", "평가자별 점수 분포·엄격도·일관성 비교")
    
    tab1, tab2, tab3 = st.tabs(["평가자별 점수 분포", "엄격도 비교", "평가자 × 상담사"])
    
    with tab1:
        fig = px.box(df, x='평가자', y='TOTAL', color='평가자',
                     labels={'TOTAL':'점수', '평가자':'평가자'})
        apply_layout(fig, '평가자별 점수 분포', 360)
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
        
        eval_stat = df.groupby('평가자')['TOTAL'].agg(['mean','std','count','min','max']).round(2).reset_index()
        eval_stat.columns = ['평가자','평균','표준편차','평가건수','최솟값','최댓값']
        st.dataframe(eval_stat, use_container_width=True)
    
    with tab2:
        # Leniency = avg score given; strictness = lower avg
        leniency = df.groupby('평가자')['TOTAL'].mean().sort_values().reset_index()
        leniency.columns = ['평가자','평균점수']
        overall = df['TOTAL'].mean()
        leniency['편향'] = leniency['평균점수'] - overall
        leniency['방향'] = leniency['편향'].apply(lambda x: '관대' if x>0 else '엄격')
        
        fig = go.Figure(go.Bar(
            x=leniency['편향'], y=leniency['평가자'],
            orientation='h',
            marker_color=leniency['방향'].apply(lambda x: '#22c55e' if x=='관대' else '#ef4444'),
            text=leniency['편향'].round(2).astype(str),
            textposition='outside'
        ))
        apply_layout(fig, '평가자별 점수 편향 (전체 평균 대비)', 300)
        fig.add_vline(x=0, line_color='#60a5fa')
        fig.update_layout(xaxis_title='점수 편향 (+관대 / -엄격)')
        st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        pivot = df.groupby(['평가자','상담사'])['TOTAL'].mean().unstack(fill_value=np.nan).round(1)
        fig = go.Figure(data=go.Heatmap(
            z=pivot.values,
            x=pivot.columns.tolist(),
            y=pivot.index.tolist(),
            colorscale=[[0,'#ef4444'],[0.5,'#f59e0b'],[1,'#22c55e']],
            text=pivot.values.round(1),
            texttemplate='%{text}',
            textfont=dict(size=11),
        ))
        apply_layout(fig, '평가자 × 상담사 평균 점수 히트맵', max(300, len(pivot)*50))
        st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────
# Page: 코멘트 NLP
# ─────────────────────────────────────────
elif page == "💬 코멘트 NLP":
    section_header("코멘트 NLP 분석", "감점 사유 키워드 클러스터링 및 테마 분석")
    
    score_cols = get_score_cols(df)
    items = get_item_names(score_cols)
    
    # Collect all comments
    all_comments = {'감점사유': [], '미차감': []}
    for item in items:
        dc = f'{item}_감점사유'
        nc = f'{item}_미차감'
        if dc in df.columns:
            reasons = df[dc].dropna().tolist()
            all_comments['감점사유'].extend([(r, item) for r in reasons if r and r != ''])
        if nc in df.columns:
            ncomments = df[nc].dropna().tolist()
            all_comments['미차감'].extend([(c, item) for c in ncomments if c and c != ''])
    
    tab1, tab2, tab3 = st.tabs(["키워드 빈도", "클러스터 분석", "상담사별 코멘트 테마"])
    
    with tab1:
        col1, col2 = st.columns(2)
        
        def get_top_words(comment_list, n=25):
            words = []
            for text, _ in comment_list:
                tokens = re.sub(r'[^\w\s가-힣a-zA-Z]', ' ', text).split()
                tokens = [t for t in tokens if len(t) >= 2 and t not in ['None','있는','하는','되는','이는','으로','에서','에게','하여','하고','있는','없는','되어','하면','하는','는데','이고','하다','위한','관련','사용','안내','필요','이행','진행','처리','내용','사항','경우','기준','코멘트']]
                words.extend(tokens)
            return Counter(words).most_common(n)
        
        with col1:
            deduct_words = get_top_words(all_comments['감점사유'])
            dw_df = pd.DataFrame(deduct_words, columns=['단어','빈도'])
            fig = px.bar(dw_df, x='빈도', y='단어', orientation='h',
                         color='빈도', color_continuous_scale='Reds',
                         labels={'빈도':'빈도','단어':''})
            apply_layout(fig, '⚠️ 감점 사유 TOP 키워드', 450)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            nodeduct_words = get_top_words(all_comments['미차감'])
            if nodeduct_words:
                nw_df = pd.DataFrame(nodeduct_words, columns=['단어','빈도'])
                fig = px.bar(nw_df, x='빈도', y='단어', orientation='h',
                             color='빈도', color_continuous_scale='Greens',
                             labels={'빈도':'빈도','단어':''})
                apply_layout(fig, '✅ 미차감 코멘트 TOP 키워드', 450)
                fig.update_coloraxes(showscale=False)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("미차감 코멘트 데이터가 없습니다.")
    
    with tab2:
        comments_text = [r for r, _ in all_comments['감점사유'] if len(r) > 5]
        
        if len(comments_text) >= 3:
            n_clusters = st.slider("클러스터 수", 2, min(8, len(comments_text)//2), 4)
            
            # TF-IDF + KMeans
            try:
                # Character-level n-grams work well for Korean without tokenizer
                vectorizer = TfidfVectorizer(
                    analyzer='char_wb', ngram_range=(2,4),
                    max_features=200, min_df=1
                )
                X = vectorizer.fit_transform(comments_text)
                km = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
                labels = km.fit_predict(X)
                
                cluster_df = pd.DataFrame({'코멘트': comments_text, '클러스터': labels})
                
                # Show cluster samples
                for ci in range(n_clusters):
                    members = cluster_df[cluster_df['클러스터']==ci]['코멘트'].tolist()
                    # Top words in cluster
                    cluster_words = []
                    for txt in members:
                        tokens = re.sub(r'[^\w\s가-힣]', ' ', txt).split()
                        cluster_words.extend([t for t in tokens if len(t) >= 2])
                    top_words = [w for w, _ in Counter(cluster_words).most_common(5)]
                    
                    with st.expander(f"🔵 클러스터 {ci+1} — [{', '.join(top_words)}] ({len(members)}건)"):
                        for m in members[:8]:
                            st.markdown(f"<div style='background:#27272a;border-radius:6px;padding:8px 12px;margin:4px 0;font-size:12px;color:#a1a1aa'>{m}</div>", unsafe_allow_html=True)
                
                # Cluster size chart
                cs = cluster_df['클러스터'].value_counts().sort_index().reset_index()
                cs.columns = ['클러스터','건수']
                cs['클러스터'] = cs['클러스터'].apply(lambda x: f'클러스터 {x+1}')
                fig = px.pie(cs, names='클러스터', values='건수', hole=0.5)
                apply_layout(fig, '클러스터별 감점사유 분포', 300)
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.warning(f"클러스터링 오류: {e}")
        else:
            st.info("클러스터 분석을 위해 더 많은 데이터가 필요합니다.")
    
    with tab3:
        # Per-agent comment theme
        agent_sel_nlp = st.selectbox("상담사 선택", all_agents, key='nlp_agent')
        agent_comments = []
        for item in items:
            dc = f'{item}_감점사유'
            if dc in df_all.columns:
                sub = df_all[(df_all['상담사']==agent_sel_nlp) & (df_all[dc].notna()) & (df_all[dc]!='')]
                agent_comments.extend(sub[dc].tolist())
        
        if agent_comments:
            words = []
            for txt in agent_comments:
                tokens = re.sub(r'[^\w\s가-힣]', ' ', txt).split()
                words.extend([t for t in tokens if len(t) >= 2])
            
            agent_words = Counter(words).most_common(20)
            aw_df = pd.DataFrame(agent_words, columns=['단어','빈도'])
            
            fig = px.bar(aw_df, x='빈도', y='단어', orientation='h',
                         color='빈도', color_continuous_scale='Oranges')
            apply_layout(fig, f'{agent_sel_nlp} 감점 코멘트 핵심 키워드', 400)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("**원본 감점 코멘트**")
            for c in agent_comments[:15]:
                st.markdown(f"<div style='background:#27272a;border-radius:6px;padding:8px 12px;margin:4px 0;font-size:12px;color:#a1a1aa'>• {c}</div>", unsafe_allow_html=True)
        else:
            st.info(f"{agent_sel_nlp}의 감점 코멘트가 없습니다.")

# Footer
st.markdown("<div style='height:40px'></div>", unsafe_allow_html=True)
st.markdown("""
<div style='border-top:1px solid #27272a;padding:16px 0;text-align:center;color:#52525b;font-size:11px'>
  QA Monitoring Dashboard · Built with Streamlit · Powered by Anthropic
</div>""", unsafe_allow_html=True)
