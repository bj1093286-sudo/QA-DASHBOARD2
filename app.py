import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from collections import Counter
import re
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────
# Google Sheets 설정
# ─────────────────────────────────────────
SHEET_ID = "1oVvew_o-JTw7mZEQ47PLXxiryD5EwtTqm1tRudVu3Ag"
GID_CALL  = "0"
GID_CHAT  = "659345665"
GID_BOARD = "1192171371"

def gsheet_csv_url(sheet_id, gid):
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"

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
# CSS
# ─────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Geist:wght@300;400;500;600;700&display=swap');
:root {
  --bg: #09090b; --card: #18181b; --border: #27272a;
  --primary: #3b82f6; --text: #fafafa; --muted: #a1a1aa; --dim: #71717a;
  --success: #22c55e; --warning: #f59e0b; --danger: #ef4444;
}
html, body, [class*="css"] {
    font-family: 'Geist', sans-serif !important;
    background-color: var(--bg) !important;
    color: var(--text) !important;
}
section[data-testid="stSidebar"] {
    background-color: #111113 !important;
    border-right: 1px solid var(--border) !important;
}
section[data-testid="stSidebar"] * { color: var(--muted) !important; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: var(--text) !important; font-size: 11px !important;
    font-weight: 700 !important; text-transform: uppercase; letter-spacing: 0.1em;
}
.main .block-container { padding: 1.5rem 2rem !important; max-width: 1400px !important; }
.qa-card {
    background: var(--card); border: 1px solid var(--border);
    border-radius: 12px; padding: 20px 24px; margin-bottom: 16px;
}
.qa-card-title { font-size: 11px; font-weight: 600; color: var(--muted);
    margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.08em; }
.qa-card-value { font-size: 26px; font-weight: 700; color: var(--text); line-height: 1.1; }
.qa-card-sub { font-size: 11px; color: var(--dim); margin-top: 4px; }
.section-header { font-size: 20px; font-weight: 700; color: var(--text); margin-bottom: 4px; }
.section-desc { font-size: 13px; color: var(--dim); margin-bottom: 20px; }
.qa-divider { border: none; border-top: 1px solid var(--border); margin: 20px 0; }
.comment-box {
    background: #1c1c1f; border: 1px solid var(--border); border-radius: 8px;
    padding: 10px 14px; margin: 6px 0; font-size: 12px; color: var(--muted); line-height: 1.6;
}
.comment-box.deduct { border-left: 3px solid var(--danger); }
.comment-box.ok     { border-left: 3px solid var(--success); }
.repeat-card {
    background: #1f1215; border: 1px solid #7f1d1d50;
    border-radius: 10px; padding: 14px 18px; margin: 10px 0;
}
.repeat-card-title { font-size: 13px; font-weight: 600; color: #f87171; margin-bottom: 6px; }
.pattern-tag {
    display: inline-block; background: #27272a; border-radius: 999px;
    padding: 3px 10px; font-size: 11px; color: var(--muted); margin: 2px;
}
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important; border-bottom: 1px solid var(--border) !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important; color: var(--dim) !important;
    font-size: 13px !important; font-weight: 500 !important;
}
.stTabs [aria-selected="true"] {
    color: var(--text) !important; border-bottom: 2px solid var(--primary) !important;
}
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# Plotly 테마
# ─────────────────────────────────────────
PLOTLY_LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(24,24,27,0.6)',
    font=dict(family='Geist, sans-serif', color='#a1a1aa', size=12),
    title_font=dict(size=14, color='#fafafa'),
    xaxis=dict(gridcolor='#27272a', linecolor='#27272a', tickfont=dict(size=11)),
    yaxis=dict(gridcolor='#27272a', linecolor='#27272a', tickfont=dict(size=11)),
    legend=dict(bgcolor='rgba(24,24,27,0.8)', bordercolor='#27272a', borderwidth=1,
                font=dict(size=11, color='#a1a1aa')),
    colorway=['#3b82f6','#6366f1','#22c55e','#f59e0b','#ef4444','#ec4899','#06b6d4','#84cc16'],
    margin=dict(l=40, r=20, t=40, b=40),
    hoverlabel=dict(bgcolor='#18181b', bordercolor='#3b82f6', font=dict(color='#fafafa', size=12)),
)
CH_COLOR = {'CALL-IN':'#3b82f6','CALL-OB':'#6366f1','CHAT':'#22c55e','게시판':'#f59e0b'}

def apply_layout(fig, title='', height=350):
    fig.update_layout(**PLOTLY_LAYOUT, title=title, height=height)
    return fig

# ─────────────────────────────────────────
# 이슈 패턴 정의
# ─────────────────────────────────────────
ISSUE_PATTERNS = [
    ('습관어/말버릇',      r'습관어|말버릇|아무래도|저희쪽|구요~|~구요'),
    ('존칭 오류',          r'존칭|요조체|비정중|반말|말투|셨을까요|되시겠'),
    ('토막말/도치',        r'토막말|도치'),
    ('인사/끝인사 누락',   r'인사.*누락|끝인사.*누락|첫인사.*누락'),
    ('양해 표현 누락',     r'양해.*누락|양해.*미이행|양해표현'),
    ('호응 부족',          r'호응.*누락|호응.*부족|즉각.*호응|적극적.*호응'),
    ('안내 누락/오안내',   r'누락|오안내|미안내'),
    ('상담이력 미기재',    r'이력.*누락|이력.*미기재|기재.*누락|메모.*누락'),
    ('이모티콘 부적절',    r'이모티콘|이모지'),
    ('두괄식 미이행',      r'두괄식'),
    ('띄어쓰기/문법',      r'띄어쓰기|문법|오타|맞춤법'),
    ('처리 오류/지연',     r'처리.*오류|처리.*지연|수기.*취소|약속.*미준수'),
    ('사진/첨부 누락',     r'사진.*누락|첨부.*누락'),
    ('문의파악 미흡',      r'문의.*파악|니즈.*유추|재질의|재문의'),
    ('가독성 문제',        r'가독성|문단.*나눔|엔터'),
]

ROOT_CAUSE_MAP = {
    '습관어/말버릇':    '무의식적 언어 습관 → 모니터링 후 의식적 교정 훈련 필요',
    '존칭 오류':        '경어법 이해 부족 → 표준 멘트 암기 및 반복 훈련 필요',
    '토막말/도치':      '문장 구성 습관 문제 → 완전한 문장으로 말하는 연습 필요',
    '인사/끝인사 누락': '상담 프로세스 누락 → 체크리스트 활용 권장',
    '양해 표현 누락':   '고객 공감 표현 부족 → 양해 멘트 패턴 훈련 필요',
    '호응 부족':        '경청 자세 부족 → 적극적 호응 멘트 연습 필요',
    '안내 누락/오안내': '업무 숙지 미흡 → 주요 안내 사항 재교육 필요',
    '상담이력 미기재':  '기록 습관 부재 → 상담 후 이력 작성 루틴화 필요',
    '이모티콘 부적절':  '채널 특성 이해 부족 → 공식 상담 채널 예절 재교육 필요',
    '두괄식 미이행':    '설명 구조 습관 문제 → 결론 먼저 말하는 훈련 필요',
    '띄어쓰기/문법':    '문서 작성 능력 부족 → 글쓰기 교육 또는 템플릿 활용 권장',
    '처리 오류/지연':   '업무 처리 정확도 부족 → 프로세스 재확인 및 재교육 필요',
    '사진/첨부 누락':   '업무 체크리스트 미활용 → 처리 단계 확인 습관화 필요',
    '문의파악 미흡':    '청취·이해 능력 부족 → 문의 요약 후 확인하는 습관 훈련 필요',
    '가독성 문제':      '채팅 문서화 능력 부족 → 문단 나눔·구조화 글쓰기 교육 필요',
    '기타':             '개별 피드백 코멘트를 참고하여 맞춤 교육 설계 필요',
}

def classify_issue(text):
    results = []
    for label, pattern in ISSUE_PATTERNS:
        if re.search(pattern, str(text)):
            results.append(label)
    return results if results else ['기타']

def sentence_analysis(comments):
    categorized = {}
    for text in comments:
        if not text or len(str(text)) < 5: continue
        for cat in classify_issue(text):
            categorized.setdefault(cat, []).append(str(text))
    return {cat: list(dict.fromkeys(texts))[:5] for cat, texts in categorized.items()}

def analyze_repeat_mistakes(agent_name, df_src):
    score_cols = [c for c in df_src.columns if c.endswith('_점수')]
    items = [c.replace('_점수','') for c in score_cols]
    agent_df = df_src[df_src['상담사'] == agent_name]
    all_issues, item_issues, period_issues = [], {}, {}
    for item in items:
        dc = f'{item}_감점사유'
        if dc not in agent_df.columns: continue
        for _, row in agent_df.iterrows():
            r = row[dc]
            if not r or r in ('None','nan',''): continue
            cats = classify_issue(r)
            all_issues.extend(cats)
            for cat in cats:
                item_issues.setdefault(item, []).append(cat)
                period_issues.setdefault(row['평가기간'], []).append(cat)
    issue_count   = Counter(all_issues)
    repeat_issues = {k: v for k, v in issue_count.items() if v >= 2}
    return issue_count, repeat_issues, item_issues, period_issues

# ─────────────────────────────────────────
# 헬퍼
# ─────────────────────────────────────────
def _num(v):
    if v is None: return None
    if isinstance(v, float) and np.isnan(v): return None
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(',','').strip())
    except: return None

def get_score_cols(df): return [c for c in df.columns if c.endswith('_점수')]
def get_item_names(sc): return [c.replace('_점수','') for c in sc]

def metric_card(title, value, sub=''):
    st.markdown(f"""<div class="qa-card">
        <div class="qa-card-title">{title}</div>
        <div class="qa-card-value">{value}</div>
        <div class="qa-card-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

def section_header(title, desc=''):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)
    if desc: st.markdown(f'<div class="section-desc">{desc}</div>', unsafe_allow_html=True)

def period_sort_key(p):
    m = re.search(r'(\d+)년\s*(\d+)월(\d+)차', str(p))
    return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else (9999,0,0)

def filter_df(df, periods, channels, agents):
    out = df.copy()
    if periods:  out = out[out['평가기간'].isin(periods)]
    if channels: out = out[out['채널'].isin(channels)]
    if agents:   out = out[out['상담사'].isin(agents)]
    return out

def compliance(scores, maxes):
    valid = [(s,m) for s,m in zip(scores,maxes) if s is not None and m is not None and m>0]
    if not valid: return None
    return sum(1 for s,m in valid if s>=m)/len(valid)*100

# ─────────────────────────────────────────
# 데이터 로딩
# ─────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data():
    call_items = {
        '첫인사':    (8,  2.5), '정보확인':  (11, 5),  '끝인사':    (14, 2.5),
        '인사톤':    (17, 2.5), '통화종료':  (20, 2.5), '음성숙련도':(23, 5),
        '감정연출':  (26, 5),   '양해':      (29, 5),   '즉각호응':  (32, 5),
        '대기':      (35, 5),   '언어표현':  (38, 5),   '경청':      (41, 5),
        '문의파악':  (44, 5),   '맞춤설명':  (47, 5),   '정확한안내':(50,10),
        '프로세스':  (53,10),   '전산처리':  (56,10),   '상담이력':  (59,10),
    }
    chat_items = {
        '첫인사':    (6,  5),  '정보확인':  (9,  5),  '끝인사':    (12, 5),
        '양해':      (15, 5),  '즉각호응':  (18, 5),  '대기':      (21, 5),
        '언어표현':  (24,10),  '가독성':    (27, 5),  '문의파악':  (30,10),
        '맞춤설명':  (33,10),  '정확한안내':(36,10),  '프로세스':  (39, 5),
        '전산처리':  (42,10),  '상담이력':  (45,10),
    }
    board_items = {
        '첫인사':    (7,  2.5), '끝인사':    (10, 2.5), '언어표현':  (13,10),
        '양해':      (16,10),  '가독성':    (19,10),   '문의파악':  (22,10),
        '맞춤설명':  (25,10),  '정확한안내':(28,10),   '프로세스':  (31,10),
        '전산처리':  (34,10),  '상담이력':  (37,15),
    }

    def parse_csv(url, items_map, agent_col, period_col, inq_cols,
                  total_col, evaluator_col, bonus_col, penalty_col,
                  inout_col=None, skip_rows=3):
        try:
            raw = pd.read_csv(url, header=None)
        except Exception as e:
            st.error(f"구글 시트 로딩 실패: {e}\n\n시트를 **공개(링크 있는 모든 사용자)** 로 설정해주세요.")
            return pd.DataFrame()

        def safe(row, i):
            if i >= len(row): return ''
            v = row.iloc[i]
            return '' if pd.isna(v) else str(v).strip()

        records = []
        for _, row in raw.iloc[skip_rows:].iterrows():
            agent = safe(row, agent_col)
            if not agent or agent in ('None','nan'): continue
            rec = {
                '상담사':     agent,
                '평가기간':   safe(row, period_col),
                '문의유형대': safe(row, inq_cols[0]),
                '문의유형중': safe(row, inq_cols[1]),
                '문의유형소': safe(row, inq_cols[2]),
                '평가자':     safe(row, evaluator_col),
                '가점':       _num(row.iloc[bonus_col])   if bonus_col   < len(row) else None,
                '감점':       _num(row.iloc[penalty_col]) if penalty_col < len(row) else None,
                '평가대상':   safe(row, inout_col) if inout_col is not None else '',
            }
            total_raw = row.iloc[total_col] if total_col < len(row) else None
            if pd.isna(total_raw) or str(total_raw).startswith('='):
                total = sum(_num(row.iloc[c]) or 0 for c,_ in items_map.values())
                total += rec['가점'] or 0
                total -= rec['감점'] or 0
            else:
                total = _num(total_raw)
            rec['TOTAL'] = total

            for iname, (ci, mx) in items_map.items():
                rec[f'{iname}_점수']     = _num(row.iloc[ci]) if ci < len(row) else None
                rec[f'{iname}_최대']     = mx
                r1 = safe(row, ci+1); r2 = safe(row, ci+2)
                rec[f'{iname}_감점사유'] = r1 if r1 not in ('None','nan') else ''
                rec[f'{iname}_미차감']   = r2 if r2 not in ('None','nan') else ''
            records.append(rec)

        return pd.DataFrame(records) if records else pd.DataFrame()

    df_call = parse_csv(
        gsheet_csv_url(SHEET_ID, GID_CALL), call_items,
        agent_col=3, period_col=1, inq_cols=[5,6,7],
        total_col=64, evaluator_col=65, bonus_col=62, penalty_col=63,
        inout_col=2
    )
    if len(df_call):
        df_call['채널'] = df_call['평가대상'].apply(
            lambda x: 'CALL-IN' if str(x).upper().strip()=='IN'
                      else 'CALL-OB' if str(x).upper().strip()=='OB' else 'CALL')

    df_chat = parse_csv(
        gsheet_csv_url(SHEET_ID, GID_CHAT), chat_items,
        agent_col=1, period_col=0, inq_cols=[3,4,5],
        total_col=50, evaluator_col=51, bonus_col=48, penalty_col=49)
    if len(df_chat): df_chat['채널'] = 'CHAT'

    df_board = parse_csv(
        gsheet_csv_url(SHEET_ID, GID_BOARD), board_items,
        agent_col=2, period_col=1, inq_cols=[4,5,6],
        total_col=42, evaluator_col=43, bonus_col=40, penalty_col=41)
    if len(df_board): df_board['채널'] = '게시판'

    df_all = pd.concat([df_call, df_chat, df_board], ignore_index=True)
    df_all['기간_정렬'] = df_all['평가기간'].apply(period_sort_key)
    df_all = df_all.sort_values('기간_정렬').reset_index(drop=True)
    return df_call, df_chat, df_board, df_all, call_items, chat_items, board_items

# ─────────────────────────────────────────
# 실행
# ─────────────────────────────────────────
with st.spinner("📡 구글 시트에서 데이터 불러오는 중..."):
    df_call, df_chat, df_board, df_all, call_items, chat_items, board_items = load_data()

if df_all.empty:
    st.error("데이터가 없습니다. 구글 시트 공유 설정을 확인해주세요.")
    st.stop()

all_periods  = sorted(df_all['평가기간'].unique(), key=period_sort_key)
all_channels = sorted(df_all['채널'].unique().tolist())
all_agents   = sorted(df_all['상담사'].unique().tolist())

# ─────────────────────────────────────────
# 사이드바
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 QA Dashboard")
    st.markdown("<hr style='border-color:#27272a;margin:8px 0 16px'>", unsafe_allow_html=True)
    page = st.radio("페이지", [
        "🏠 개요", "📈 점수 트렌드", "🗂 문의유형 분석",
        "📋 항목별 분석", "⚠️ 감점 분석", "👤 상담사 종합평가",
        "✅ 이행률 분석", "🔍 평가자 분석",
        "💬 코멘트 분석", "🔁 반복 실수 분석",
    ], label_visibility='collapsed')
    st.markdown("<hr style='border-color:#27272a;margin:16px 0 12px'>", unsafe_allow_html=True)
    st.markdown("### 필터")
    sel_period  = st.multiselect("평가기간", all_periods,  default=all_periods)
    sel_channel = st.multiselect("채널",     all_channels, default=all_channels,
                                  help="CALL-IN: 인바운드 / CALL-OB: 아웃바운드")
    sel_agent   = st.multiselect("상담사",   all_agents,   default=all_agents)
    st.markdown("<hr style='border-color:#27272a;margin:12px 0'>", unsafe_allow_html=True)
    if st.button("🔄 새로고침"):
        st.cache_data.clear(); st.rerun()
    st.markdown(f"<div style='font-size:11px;color:#52525b;margin-top:8px'>총 {len(df_all)}건</div>",
                unsafe_allow_html=True)

df = filter_df(df_all, sel_period, sel_channel, sel_agent)

# ══════════════════════════════════════════════════════════
# 🏠 개요
# ══════════════════════════════════════════════════════════
if page == "🏠 개요":
    section_header("QA 모니터링 대시보드", "상담 품질 평가 현황 종합 요약")
    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: metric_card("총 평가건수", f"{len(df):,}건")
    with c2: metric_card("평균 점수",   f"{df['TOTAL'].dropna().mean():.1f}점" if len(df) else '-')
    with c3: metric_card("상담사 수",   f"{df['상담사'].nunique()}명")
    with c4: metric_card("평가 기간",   f"{df['평가기간'].nunique()}개")
    with c5: metric_card("활성 채널",   f"{df['채널'].nunique()}개")

    col1, col2 = st.columns([3,2])
    with col1:
        grp = df.groupby(['평가기간','채널','기간_정렬'])['TOTAL'].mean().reset_index().sort_values('기간_정렬')
        fig = px.line(grp, x='평가기간', y='TOTAL', color='채널',
                      markers=True, color_discrete_map=CH_COLOR, labels={'TOTAL':'평균 점수'})
        apply_layout(fig, '채널별 평균 점수 추이 (CALL-IN / CALL-OB 분리)', 320)
        fig.update_traces(line=dict(width=2.5), marker=dict(size=6))
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        ch_cnt = df['채널'].value_counts().reset_index(); ch_cnt.columns = ['채널','건수']
        fig2 = px.pie(ch_cnt, names='채널', values='건수', hole=0.65, color='채널', color_discrete_map=CH_COLOR)
        apply_layout(fig2, '채널별 비중', 320)
        st.plotly_chart(fig2, use_container_width=True)

    section_header("상담사별 요약")
    asum = df.groupby('상담사').agg(전체평균=('TOTAL','mean'), 평가건수=('TOTAL','count'),
                                     최고점=('TOTAL','max'), 최저점=('TOTAL','min')).round(1).sort_values('전체평균', ascending=False)
    st.dataframe(asum, use_container_width=True)

# ══════════════════════════════════════════════════════════
# 📈 점수 트렌드
# ══════════════════════════════════════════════════════════
elif page == "📈 점수 트렌드":
    section_header("점수 트렌드", "기간별·상담사별·채널별 점수 변화")
    tab1, tab2, tab3 = st.tabs(["기간별 추이", "상담사별 비교", "Box Plot"])
    with tab1:
        grp = df.groupby(['평가기간','상담사','기간_정렬'])['TOTAL'].mean().reset_index().sort_values('기간_정렬')
        fig = px.line(grp, x='평가기간', y='TOTAL', color='상담사', markers=True, labels={'TOTAL':'평균 점수'})
        apply_layout(fig, '상담사별 점수 추이', 400)
        fig.update_traces(line=dict(width=2), marker=dict(size=5))
        st.plotly_chart(fig, use_container_width=True)

        grp2 = df.groupby(['평가기간','채널','기간_정렬'])['TOTAL'].mean().reset_index().sort_values('기간_정렬')
        fig2 = px.line(grp2, x='평가기간', y='TOTAL', color='채널',
                       markers=True, color_discrete_map=CH_COLOR, labels={'TOTAL':'평균 점수'})
        apply_layout(fig2, '채널별 점수 추이', 300)
        fig2.update_traces(line=dict(width=2.5), marker=dict(size=6))
        st.plotly_chart(fig2, use_container_width=True)
    with tab2:
        agent_avg = df.groupby(['상담사','채널'])['TOTAL'].mean().reset_index()
        fig = px.bar(agent_avg, x='상담사', y='TOTAL', color='채널', barmode='group',
                     color_discrete_map=CH_COLOR, labels={'TOTAL':'평균 점수'})
        apply_layout(fig, '상담사 × 채널별 평균 점수', 380)
        fig.add_hline(y=df['TOTAL'].mean(), line_dash='dash', line_color='#60a5fa',
                      annotation_text=f'전체평균 {df["TOTAL"].mean():.1f}', annotation_font_color='#60a5fa')
        st.plotly_chart(fig, use_container_width=True)
    with tab3:
        fig = px.box(df, x='상담사', y='TOTAL', color='채널',
                     color_discrete_map=CH_COLOR, labels={'TOTAL':'점수'})
        apply_layout(fig, '상담사별 점수 분포', 400)
        st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════
# 🗂 문의유형 분석
# ══════════════════════════════════════════════════════════
elif page == "🗂 문의유형 분석":
    section_header("문의유형 분석", "유형별 점수 히트맵 및 분포")
    tab1, tab2 = st.tabs(["유형 × 상담사 히트맵", "유형별 건수/점수"])
    with tab1:
        pivot = df.groupby(['문의유형대','상담사'])['TOTAL'].mean().unstack(fill_value=np.nan).round(1)
        fig = go.Figure(data=go.Heatmap(
            z=pivot.values, x=pivot.columns.tolist(), y=pivot.index.tolist(),
            colorscale=[[0,'#ef4444'],[0.5,'#f59e0b'],[1,'#22c55e']],
            text=pivot.values.round(1), texttemplate='%{text}', textfont=dict(size=11),
        ))
        apply_layout(fig, '문의유형(대) × 상담사 평균점수', max(300, len(pivot)*35))
        st.plotly_chart(fig, use_container_width=True)
    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            cnt = df['문의유형대'].value_counts().reset_index(); cnt.columns=['유형','건수']
            fig = px.bar(cnt, x='유형', y='건수', color='건수', color_continuous_scale='Blues')
            apply_layout(fig, '문의유형별 건수', 300); fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            avg = df.groupby('문의유형대')['TOTAL'].mean().reset_index(); avg.columns=['유형','평균점수']
            fig = px.bar(avg, x='유형', y='평균점수', color='평균점수', color_continuous_scale=['#ef4444','#22c55e'])
            apply_layout(fig, '문의유형별 평균점수', 300); fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        df_sb = df[df['문의유형대']!='']
        if len(df_sb):
            fig3 = px.sunburst(df_sb, path=['문의유형대','문의유형중'], color='TOTAL',
                               color_continuous_scale=['#ef4444','#22c55e'])
            apply_layout(fig3, '문의유형 계층 구조', 420)
            st.plotly_chart(fig3, use_container_width=True)

# ══════════════════════════════════════════════════════════
# 📋 항목별 분석
# ══════════════════════════════════════════════════════════
elif page == "📋 항목별 분석":
    section_header("항목별 점수 분석", "QA 평가 항목별 달성률 및 시계열")
    score_cols = get_score_cols(df); items = get_item_names(score_cols)
    tab1, tab2, tab3 = st.tabs(["항목별 달성률", "항목 × 기간 추이", "상담사 × 항목 히트맵"])
    with tab1:
        rows = []
        for item in items:
            sc = df[f'{item}_점수'].dropna(); mx = df[f'{item}_최대'].dropna()
            if len(sc)==0: continue
            avg_s=sc.mean(); avg_m=mx.mean() if len(mx) else 1
            rows.append({'항목':item,'평균점수':avg_s,'최대점수':avg_m,'달성률':avg_s/avg_m*100 if avg_m else 0})
        item_df = pd.DataFrame(rows).sort_values('달성률')
        fig = go.Figure(go.Bar(
            x=item_df['달성률'], y=item_df['항목'], orientation='h',
            marker_color=item_df['달성률'].apply(lambda x:'#22c55e' if x>=90 else '#f59e0b' if x>=75 else '#ef4444'),
            text=item_df['달성률'].round(1).astype(str)+'%', textposition='inside',
            textfont=dict(color='white',size=11),
        ))
        apply_layout(fig,'항목별 평균 달성률', max(300, len(item_df)*28))
        fig.update_layout(xaxis=dict(range=[0,108],title='달성률 (%)'))
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(item_df.round(2), use_container_width=True)
    with tab2:
        sel_items = st.multiselect("항목 선택", items, default=items[:5])
        if sel_items:
            rows2 = []
            for item in sel_items:
                grp = df.groupby(['평가기간','기간_정렬'])[f'{item}_점수'].mean().reset_index().sort_values('기간_정렬')
                grp['항목']=item; grp.rename(columns={f'{item}_점수':'점수'}, inplace=True)
                rows2.append(grp)
            trend_df = pd.concat(rows2)
            fig = px.line(trend_df, x='평가기간', y='점수', color='항목', markers=True)
            apply_layout(fig,'항목별 점수 추이',380)
            st.plotly_chart(fig, use_container_width=True)
    with tab3:
        agent_item = {}
        for item in items:
            s = df.groupby('상담사')[f'{item}_점수'].mean()
            m = df.groupby('상담사')[f'{item}_최대'].mean()
            agent_item[item] = (s/m*100).round(1)
        heat_df = pd.DataFrame(agent_item)
        fig = go.Figure(data=go.Heatmap(
            z=heat_df.values, x=heat_df.columns.tolist(), y=heat_df.index.tolist(),
            colorscale=[[0,'#ef4444'],[0.5,'#f59e0b'],[1,'#22c55e']],
            zmin=0, zmax=100, text=heat_df.values.round(0), texttemplate='%{text}%',
            textfont=dict(size=10),
        ))
        apply_layout(fig,'상담사 × 항목 달성률 히트맵', max(350, len(heat_df)*35))
        st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════
# ⚠️ 감점 분석
# ══════════════════════════════════════════════════════════
elif page == "⚠️ 감점 분석":
    section_header("감점 분석", "감점 사유 패턴 — 정량 + 정성")
    score_cols = get_score_cols(df); items = get_item_names(score_cols)
    deduct_records = []
    for item in items:
        if f'{item}_감점사유' not in df.columns: continue
        sub = df[[f'{item}_점수',f'{item}_최대',f'{item}_감점사유','상담사','평가기간','채널']].copy()
        sub.columns = ['점수','최대','감점사유','상담사','평가기간','채널']
        sub['항목'] = item
        sub = sub[sub['점수'].notna()&sub['최대'].notna()]
        sub['감점여부'] = sub['점수']<sub['최대']
        deduct_records.append(sub[sub['감점여부']&(sub['감점사유']!='')])
    deduct_df = pd.concat(deduct_records, ignore_index=True) if deduct_records else pd.DataFrame()

    tab1, tab2, tab3 = st.tabs(["항목별 감점 빈도","상담사별 패턴","전체 감점 목록"])
    with tab1:
        if len(deduct_df):
            idc = deduct_df['항목'].value_counts().reset_index(); idc.columns=['항목','감점건수']
            fig = px.bar(idc, x='항목', y='감점건수', color='감점건수', color_continuous_scale='Reds')
            apply_layout(fig,'항목별 감점 발생 건수',340); fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
            dr = [{'항목':item,'감점율':deduct_df[deduct_df['항목']==item].shape[0]/df[f'{item}_점수'].dropna().shape[0]*100}
                  for item in items if df[f'{item}_점수'].dropna().shape[0]>0]
            dr_df = pd.DataFrame(dr).sort_values('감점율',ascending=False)
            fig2 = px.bar(dr_df, x='항목', y='감점율', color='감점율', color_continuous_scale=['#22c55e','#ef4444'])
            apply_layout(fig2,'항목별 감점 발생률 (%)',300); fig2.update_coloraxes(showscale=False)
            st.plotly_chart(fig2, use_container_width=True)
    with tab2:
        if len(deduct_df):
            ad = deduct_df.groupby(['상담사','항목']).size().unstack(fill_value=0)
            fig = go.Figure(data=go.Heatmap(
                z=ad.values, x=ad.columns.tolist(), y=ad.index.tolist(),
                colorscale=[[0,'#18181b'],[1,'#ef4444']],
                text=ad.values, texttemplate='%{text}', textfont=dict(size=11),
            ))
            apply_layout(fig,'상담사 × 항목 감점 횟수', max(300, len(ad)*35))
            st.plotly_chart(fig, use_container_width=True)
    with tab3:
        if len(deduct_df):
            ac = deduct_df['상담사'].value_counts().reset_index(); ac.columns=['상담사','감점건수']
            fig = px.bar(ac, x='상담사', y='감점건수', color='감점건수', color_continuous_scale='Reds')
            apply_layout(fig,'상담사별 총 감점 건수',300); fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(deduct_df[['상담사','채널','항목','점수','최대','감점사유']].reset_index(drop=True),
                         use_container_width=True)

# ══════════════════════════════════════════════════════════
# 👤 상담사 종합평가
# ══════════════════════════════════════════════════════════
elif page == "👤 상담사 종합평가":
    section_header("상담사 종합 평가", "레이더 차트로 보는 강점/약점 프로파일")
    agent_sel = st.selectbox("상담사 선택", all_agents)
    agent_df  = df_all[df_all['상담사']==agent_sel]
    c1,c2,c3 = st.columns(3)
    with c1: metric_card("평균 점수", f"{agent_df['TOTAL'].mean():.1f}점" if len(agent_df) else '-')
    with c2: metric_card("평가 건수", f"{len(agent_df)}건")
    with c3: metric_card("활동 채널", ', '.join(agent_df['채널'].unique()))

    col1, col2 = st.columns(2)
    with col1:
        sc = get_score_cols(agent_df); its = get_item_names(sc)
        rd = []
        for item in its:
            s = agent_df[f'{item}_점수'].dropna().mean()
            m = agent_df[f'{item}_최대'].dropna().mean()
            if m and m>0: rd.append({'항목':item,'달성률':s/m*100 if s else 0})
        if rd:
            rd = pd.DataFrame(rd)
            cats = rd['항목'].tolist()+[rd['항목'].iloc[0]]
            vals = rd['달성률'].tolist()+[rd['달성률'].iloc[0]]
            ov = []
            for item in its:
                s = df[f'{item}_점수'].dropna().mean() if f'{item}_점수' in df.columns else 0
                m = df[f'{item}_최대'].dropna().mean() if f'{item}_최대' in df.columns else 1
                ov.append(s/m*100 if m else 0)
            ov += [ov[0]]
            fig = go.Figure()
            fig.add_trace(go.Scatterpolar(r=vals, theta=cats, fill='toself',
                fillcolor='rgba(59,130,246,0.15)', line=dict(color='#3b82f6',width=2),
                marker=dict(size=5), name=agent_sel))
            fig.add_trace(go.Scatterpolar(r=ov, theta=cats, fill='toself',
                fillcolor='rgba(99,102,241,0.06)', line=dict(color='#6366f1',width=1.5,dash='dash'),
                marker=dict(size=4), name='전체평균'))
            apply_layout(fig, f'{agent_sel} 역량 레이더', 420)
            fig.update_layout(polar=dict(
                bgcolor='rgba(24,24,27,0.6)',
                radialaxis=dict(visible=True,range=[0,100],tickfont=dict(size=9,color='#71717a'),
                                gridcolor='#27272a',linecolor='#27272a'),
                angularaxis=dict(tickfont=dict(size=10,color='#a1a1aa'),gridcolor='#27272a')
            ))
            st.plotly_chart(fig, use_container_width=True)
    with col2:
        trend = agent_df.groupby(['평가기간','기간_정렬'])['TOTAL'].mean().reset_index().sort_values('기간_정렬')
        fig2 = px.bar(trend, x='평가기간', y='TOTAL', color='TOTAL', color_continuous_scale=['#ef4444','#22c55e'])
        apply_layout(fig2,'기간별 점수',300); fig2.update_coloraxes(showscale=False)
        st.plotly_chart(fig2, use_container_width=True)
        ch = agent_df.groupby('채널')['TOTAL'].agg(['mean','count']).reset_index()
        ch.columns=['채널','평균점수','건수']; ch['평균점수']=ch['평균점수'].round(1)
        st.dataframe(ch, use_container_width=True)

    st.markdown("<hr class='qa-divider'>", unsafe_allow_html=True)
    st.markdown("**항목별 상세 현황**")
    sc = get_score_cols(agent_df); its = get_item_names(sc)
    rows3 = []
    for item in its:
        s = agent_df[f'{item}_점수'].dropna(); m = agent_df[f'{item}_최대'].dropna()
        if len(s)==0: continue
        avg_s=s.mean(); avg_m=m.mean() if len(m) else None
        reasons = agent_df[f'{item}_감점사유'].dropna(); reasons=reasons[reasons!='']
        rows3.append({'항목':item,'평균점수':round(avg_s,2),'최대점수':avg_m,
                      '달성률':f"{avg_s/avg_m*100:.1f}%" if avg_m else '-',
                      '주요감점사유':reasons.iloc[0][:60] if len(reasons) else '없음'})
    st.dataframe(pd.DataFrame(rows3), use_container_width=True)

# ══════════════════════════════════════════════════════════
# ✅ 이행률 분석
# ══════════════════════════════════════════════════════════
elif page == "✅ 이행률 분석":
    section_header("이행률 분석", "만점 달성 비율 — 항목별·상담사별·기간별")
    score_cols = get_score_cols(df); items = get_item_names(score_cols)
    tab1, tab2, tab3 = st.tabs(["항목별","상담사별","기간별 추이"])
    with tab1:
        rows = [{'항목':item,'이행률':round(compliance(df[f'{item}_점수'].tolist(),df[f'{item}_최대'].tolist()),1)}
                for item in items if compliance(df[f'{item}_점수'].tolist(),df[f'{item}_최대'].tolist()) is not None]
        comp_df = pd.DataFrame(rows).sort_values('이행률')
        fig = go.Figure(go.Bar(
            x=comp_df['이행률'], y=comp_df['항목'], orientation='h',
            marker_color=comp_df['이행률'].apply(lambda x:'#22c55e' if x>=90 else '#f59e0b' if x>=75 else '#ef4444'),
            text=comp_df['이행률'].astype(str)+'%', textposition='outside',
            textfont=dict(size=11,color='#a1a1aa')
        ))
        apply_layout(fig,'항목별 이행률 (만점 달성 비율)', max(300, len(comp_df)*28))
        fig.update_layout(xaxis=dict(range=[0,108]))
        st.plotly_chart(fig, use_container_width=True)
    with tab2:
        ac2 = {}
        for agent in df['상담사'].unique():
            adf = df[df['상담사']==agent]
            vals = [compliance(adf[f'{item}_점수'].tolist(),adf[f'{item}_최대'].tolist()) for item in items]
            vals = [v for v in vals if v is not None]
            if vals: ac2[agent] = round(np.mean(vals),1)
        ac_df = pd.DataFrame(list(ac2.items()),columns=['상담사','이행률']).sort_values('이행률')
        fig = px.bar(ac_df, x='이행률', y='상담사', orientation='h', color='이행률',
                     color_continuous_scale=['#ef4444','#22c55e'])
        apply_layout(fig,'상담사별 평균 이행률', max(300, len(ac_df)*30))
        fig.update_coloraxes(showscale=False)
        avg_c = ac_df['이행률'].mean()
        fig.add_vline(x=avg_c, line_dash='dash', line_color='#60a5fa',
                      annotation_text=f'평균 {avg_c:.1f}%', annotation_font_color='#60a5fa')
        st.plotly_chart(fig, use_container_width=True)
    with tab3:
        pc = {}
        for period in df['평가기간'].unique():
            pdf = df[df['평가기간']==period]
            sk = pdf['기간_정렬'].iloc[0]
            vals = [compliance(pdf[f'{item}_점수'].tolist(),pdf[f'{item}_최대'].tolist()) for item in items]
            vals = [v for v in vals if v is not None]
            if vals: pc[period] = (round(np.mean(vals),1), sk)
        pc_df = pd.DataFrame([(k,v[0],v[1]) for k,v in pc.items()],columns=['기간','이행률','정렬']).sort_values('정렬')
        fig = px.line(pc_df, x='기간', y='이행률', markers=True)
        apply_layout(fig,'기간별 전체 이행률 추이',320)
        fig.update_traces(line=dict(width=2.5,color='#3b82f6'), marker=dict(size=7,color='#60a5fa'))
        fig.add_hline(y=90, line_dash='dot', line_color='#22c55e',
                      annotation_text='목표 90%', annotation_font_color='#22c55e')
        st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════
# 🔍 평가자 분석
# ══════════════════════════════════════════════════════════
elif page == "🔍 평가자 분석":
    section_header("평가자 행동 분석", "엄격도·일관성·편향 비교")
    tab1, tab2, tab3 = st.tabs(["점수 분포","엄격도 편향","평가자 × 상담사"])
    with tab1:
        fig = px.box(df, x='평가자', y='TOTAL', color='평가자', labels={'TOTAL':'점수'})
        apply_layout(fig,'평가자별 점수 분포',360); fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
        es = df.groupby('평가자')['TOTAL'].agg(['mean','std','count','min','max']).round(2).reset_index()
        es.columns=['평가자','평균','표준편차','평가건수','최솟값','최댓값']
        st.dataframe(es, use_container_width=True)
    with tab2:
        ln = df.groupby('평가자')['TOTAL'].mean().reset_index(); ln.columns=['평가자','평균점수']
        overall = df['TOTAL'].mean()
        ln['편향']=(ln['평균점수']-overall).round(2)
        ln['방향']=ln['편향'].apply(lambda x:'관대' if x>0 else '엄격')
        fig = go.Figure(go.Bar(
            x=ln['편향'], y=ln['평가자'], orientation='h',
            marker_color=ln['방향'].apply(lambda x:'#22c55e' if x=='관대' else '#ef4444'),
            text=ln['편향'].astype(str), textposition='outside'
        ))
        apply_layout(fig,'평가자별 점수 편향 (+관대 / -엄격)',300)
        fig.add_vline(x=0, line_color='#60a5fa')
        st.plotly_chart(fig, use_container_width=True)
    with tab3:
        pivot = df.groupby(['평가자','상담사'])['TOTAL'].mean().unstack(fill_value=np.nan).round(1)
        fig = go.Figure(data=go.Heatmap(
            z=pivot.values, x=pivot.columns.tolist(), y=pivot.index.tolist(),
            colorscale=[[0,'#ef4444'],[0.5,'#f59e0b'],[1,'#22c55e']],
            text=pivot.values.round(1), texttemplate='%{text}', textfont=dict(size=11),
        ))
        apply_layout(fig,'평가자 × 상담사 평균점수', max(300, len(pivot)*50))
        st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════
# 💬 코멘트 분석 (문장 단위)
# ══════════════════════════════════════════════════════════
elif page == "💬 코멘트 분석":
    section_header("코멘트 문장 분석", "감점 사유 문장을 카테고리별로 분류하고 패턴 도출")
    score_cols = get_score_cols(df); items = get_item_names(score_cols)

    all_deduct, all_ok = [], []
    for item in items:
        dc = f'{item}_감점사유'; nc = f'{item}_미차감'
        if dc in df.columns:
            for _, row in df.iterrows():
                r = row[dc]
                if r and r not in ('None','nan',''): all_deduct.append({'텍스트':r,'항목':item,'상담사':row['상담사'],'기간':row['평가기간']})
        if nc in df.columns:
            for _, row in df.iterrows():
                r = row[nc]
                if r and r not in ('None','nan',''): all_ok.append({'텍스트':r,'항목':item,'상담사':row['상담사']})

    tab1, tab2, tab3 = st.tabs(["감점 문장 카테고리 분석","미차감 문장 분석","클러스터 분석"])
    with tab1:
        st.markdown("#### 감점 사유 → 이슈 카테고리 자동 분류")
        if all_deduct:
            texts = [c['텍스트'] for c in all_deduct]
            cat_analysis = sentence_analysis(texts)
            cat_counts = {k:len(v) for k,v in cat_analysis.items()}
            cat_df = pd.DataFrame(list(cat_counts.items()),columns=['카테고리','건수']).sort_values('건수',ascending=False)
            fig = px.bar(cat_df, x='건수', y='카테고리', orientation='h',
                         color='건수', color_continuous_scale='Reds')
            apply_layout(fig,'이슈 카테고리별 감점 건수', max(300, len(cat_df)*32))
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("#### 카테고리별 대표 감점 문장")
            for cat, examples in sorted(cat_analysis.items(), key=lambda x:-len(x[1])):
                with st.expander(f"🔴 {cat} — {len(examples)}건"):
                    for ex in examples:
                        st.markdown(f'<div class="comment-box deduct">{ex}</div>', unsafe_allow_html=True)
        else:
            st.info("감점 사유 데이터가 없습니다.")

    with tab2:
        st.markdown("#### 미차감 코멘트 → 카테고리 분류")
        if all_ok:
            ok_texts = [c['텍스트'] for c in all_ok]
            ok_analysis = sentence_analysis(ok_texts)
            ok_df = pd.DataFrame([(k,len(v)) for k,v in ok_analysis.items()],columns=['카테고리','건수']).sort_values('건수',ascending=False)
            fig = px.bar(ok_df, x='건수', y='카테고리', orientation='h',
                         color='건수', color_continuous_scale='Greens')
            apply_layout(fig,'미차감 코멘트 카테고리 분포', max(280, len(ok_df)*32))
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
            for cat, examples in sorted(ok_analysis.items(), key=lambda x:-len(x[1])):
                with st.expander(f"🟢 {cat} — {len(examples)}건"):
                    for ex in examples:
                        st.markdown(f'<div class="comment-box ok">{ex}</div>', unsafe_allow_html=True)
        else:
            st.info("미차감 코멘트 데이터가 없습니다.")

    with tab3:
        st.markdown("#### KMeans 클러스터링 (문장 유사도 기반 그룹화)")
        texts_cl = [c['텍스트'] for c in all_deduct if len(c['텍스트'])>5]
        if len(texts_cl) >= 4:
            n_clusters = st.slider("클러스터 수", 2, min(8, len(texts_cl)//2), 4)
            try:
                vec = TfidfVectorizer(analyzer='char_wb', ngram_range=(2,4), max_features=300, min_df=1)
                X = vec.fit_transform(texts_cl)
                km = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
                labels = km.fit_predict(X)
                cl_df = pd.DataFrame({'문장':texts_cl,'클러스터':labels})
                for ci in range(n_clusters):
                    members = cl_df[cl_df['클러스터']==ci]['문장'].tolist()
                    cats_in = []
                    for txt in members: cats_in.extend(classify_issue(txt))
                    top_cats = [c for c,_ in Counter(cats_in).most_common(3)]
                    label_str = ' / '.join(top_cats) if top_cats else '기타'
                    with st.expander(f"🔵 클러스터 {ci+1} — {label_str} ({len(members)}건)"):
                        for m in members[:8]:
                            st.markdown(f'<div class="comment-box deduct">{m}</div>', unsafe_allow_html=True)
            except Exception as e:
                st.warning(f"클러스터링 오류: {e}")
        else:
            st.info("클러스터 분석을 위한 데이터가 부족합니다.")

# ══════════════════════════════════════════════════════════
# 🔁 반복 실수 분석
# ══════════════════════════════════════════════════════════
elif page == "🔁 반복 실수 분석":
    section_header("반복 실수 원인 분석", "상담사별 반복되는 실수 패턴과 근본 원인 도출")

    tab1, tab2 = st.tabs(["전체 상담사 현황", "개인별 심층 분석"])

    with tab1:
        all_issue_labels = [label for label,_ in ISSUE_PATTERNS]+['기타']
        heatmap_data = {}
        for agent in all_agents:
            ic,_,_,_ = analyze_repeat_mistakes(agent, df_all)
            heatmap_data[agent] = {label: ic.get(label,0) for label in all_issue_labels}
        heat_df = pd.DataFrame(heatmap_data).T
        heat_df = heat_df[heat_df.sum(axis=1)>0]
        if len(heat_df):
            fig = go.Figure(data=go.Heatmap(
                z=heat_df.values, x=heat_df.columns.tolist(), y=heat_df.index.tolist(),
                colorscale=[[0,'#18181b'],[0.3,'#7f1d1d'],[1,'#ef4444']],
                text=heat_df.values, texttemplate='%{text}', textfont=dict(size=11),
            ))
            apply_layout(fig,'상담사별 이슈 유형 발생 횟수', max(350, len(heat_df)*38))
            fig.update_layout(xaxis=dict(tickangle=-35))
            st.plotly_chart(fig, use_container_width=True)

        total_issues = Counter()
        for agent in all_agents:
            ic,_,_,_ = analyze_repeat_mistakes(agent, df_all)
            total_issues.update(ic)
        total_df = pd.DataFrame(total_issues.most_common(), columns=['이슈유형','총발생횟수'])
        col1, col2 = st.columns(2)
        with col1:
            fig2 = px.bar(total_df, x='총발생횟수', y='이슈유형', orientation='h',
                          color='총발생횟수', color_continuous_scale='Reds')
            apply_layout(fig2,'전체 이슈 유형 랭킹', max(300, len(total_df)*30))
            fig2.update_coloraxes(showscale=False)
            st.plotly_chart(fig2, use_container_width=True)
        with col2:
            st.dataframe(total_df, use_container_width=True)

    with tab2:
        agent_sel = st.selectbox("상담사 선택", all_agents, key='repeat_agent')
        ic, repeat_issues, item_issues, period_issues = analyze_repeat_mistakes(agent_sel, df_all)

        if not ic:
            st.success(f"✅ {agent_sel} 상담사는 분석 가능한 감점 데이터가 없습니다.")
        else:
            c1,c2,c3 = st.columns(3)
            with c1: metric_card("총 이슈 발생", f"{sum(ic.values())}건")
            with c2: metric_card("반복 이슈 유형", f"{len(repeat_issues)}개", "2회 이상 반복")
            with c3:
                top = ic.most_common(1)
                metric_card("최다 반복 이슈", top[0][0] if top else '-', f"{top[0][1]}회" if top else '')

            ic_df = pd.DataFrame(ic.most_common(), columns=['이슈유형','발생횟수'])
            ic_df['반복'] = ic_df['이슈유형'].apply(lambda x: x in repeat_issues)
            fig = go.Figure(go.Bar(
                x=ic_df['발생횟수'], y=ic_df['이슈유형'], orientation='h',
                marker_color=ic_df['반복'].apply(lambda x:'#ef4444' if x else '#f59e0b'),
                text=ic_df['발생횟수'], textposition='outside',
                textfont=dict(size=11,color='#a1a1aa')
            ))
            apply_layout(fig, f'{agent_sel} — 이슈 유형별 발생 횟수', max(300, len(ic_df)*30))
            st.plotly_chart(fig, use_container_width=True)

            if repeat_issues:
                st.markdown("#### 🔴 반복 이슈 상세 및 근본 원인")
                agent_df_all = df_all[df_all['상담사']==agent_sel]
                sc_all = get_score_cols(df_all); its_all = get_item_names(sc_all)
                for issue, cnt in sorted(repeat_issues.items(), key=lambda x:-x[1]):
                    affected = [item for item,cats in item_issues.items() if issue in cats]
                    related = []
                    for item in its_all:
                        dc = f'{item}_감점사유'
                        if dc not in agent_df_all.columns: continue
                        for r in agent_df_all[dc].dropna():
                            if r and r not in ('None','nan','') and issue in classify_issue(r):
                                related.append(r)
                    related_unique = list(dict.fromkeys(related))[:5]
                    root = ROOT_CAUSE_MAP.get(issue, '개별 피드백 참고 후 맞춤 교육 설계 필요')
                    affected_tags = ''.join([f'<span class="pattern-tag">{i}</span>' for i in set(affected)])
                    comments_html = ''.join([f'<div class="comment-box deduct">{c}</div>' for c in related_unique])
                    st.markdown(f"""
                    <div class="repeat-card">
                        <div class="repeat-card-title">🔴 {issue} — {cnt}회 반복 발생</div>
                        <div style="font-size:12px;color:#a1a1aa;margin-bottom:8px">발생 항목: {affected_tags}</div>
                        <div style="font-size:12px;color:#fbbf24;margin-bottom:10px;font-weight:500">📌 근본 원인: {root}</div>
                        <div style="font-size:11px;color:#71717a;font-weight:600;margin-bottom:4px">📝 실제 감점 문장</div>
                        {comments_html}
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("반복 이슈(2회 이상)가 없습니다.")

            if period_issues:
                st.markdown("#### 기간별 이슈 발생 추이")
                rows_p = [{'기간':p,'이슈':iss} for p,ilist in period_issues.items() for iss in ilist]
                if rows_p:
                    pif = pd.DataFrame(rows_p)
                    pif['정렬'] = pif['기간'].apply(period_sort_key)
                    pif = pif.sort_values('정렬')
                    pvt = pif.groupby(['기간','이슈']).size().unstack(fill_value=0).reset_index()
                    fig_p = go.Figure()
                    for ic2 in [c for c in pvt.columns if c!='기간']:
                        fig_p.add_trace(go.Scatter(x=pvt['기간'], y=pvt[ic2],
                                                   mode='lines+markers', name=ic2,
                                                   line=dict(width=2), marker=dict(size=5)))
                    apply_layout(fig_p, f'{agent_sel} — 기간별 이슈 추이', 320)
                    st.plotly_chart(fig_p, use_container_width=True)

# ─────────────────────────────────────────
# Footer
# ─────────────────────────────────────────
st.markdown("<div style='height:40px'></div>", unsafe_allow_html=True)
st.markdown("""
<div style='border-top:1px solid #27272a;padding:16px 0;text-align:center;color:#52525b;font-size:11px'>
  QA Monitoring Dashboard · Streamlit · Google Sheets 연동
</div>""", unsafe_allow_html=True)
