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
warnings.filterwarnings("ignore")

SHEET_ID = "1oVvew_o-JTw7mZEQ47PLXxiryD5EwtTqm1tRudVu3Ag"
GID_CALL  = "0"
GID_CHAT  = "659345665"
GID_BOARD = "1192171371"

def gsheet_csv_url(sid, gid):
    return f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid={gid}"

st.set_page_config(page_title="QA Monitoring", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700&display=swap');
html,body,[class*="css"],.stApp{font-family:'Noto Sans KR',sans-serif!important;background:#f4f6f9!important;color:#1e2330!important;}
section[data-testid="stSidebar"]{background:#ffffff!important;border-right:1px solid #e2e6ea!important;}
section[data-testid="stSidebar"]>div{padding:20px 12px!important;}
section[data-testid="stSidebar"] *{color:#374151!important;}
.main .block-container{padding:20px 28px!important;max-width:1400px!important;}

div[data-testid="stRadio"]>div{gap:3px!important;flex-direction:column!important;}
div[data-testid="stRadio"]>div>label{
    background:#f4f6f9!important;border:1px solid #e2e6ea!important;border-radius:8px!important;
    padding:9px 14px!important;cursor:pointer!important;transition:all .15s!important;
    font-size:13px!important;font-weight:500!important;color:#374151!important;width:100%!important;margin:0!important;
}
div[data-testid="stRadio"]>div>label:hover{background:#eef2ff!important;border-color:#6366f1!important;color:#6366f1!important;}
div[data-testid="stRadio"] input[type=radio]{display:none!important;}

.kpi-card{background:#fff;border:1px solid #e2e6ea;border-radius:12px;padding:16px 18px;height:100%;}
.kpi-label{font-size:11px;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px;}
.kpi-value{font-size:22px;font-weight:700;color:#1e2330;}
.kpi-sub{font-size:11px;color:#94a3b8;margin-top:2px;}
.kpi-up{color:#22c55e;font-size:11px;font-weight:600;}
.kpi-down{color:#ef4444;font-size:11px;font-weight:600;}

.alert-card{background:#fff;border:1px solid #fca5a5;border-left:4px solid #ef4444;border-radius:10px;padding:12px 14px;margin:5px 0;}
.alert-title{font-size:13px;font-weight:600;color:#dc2626;margin-bottom:3px;}
.alert-body{font-size:12px;color:#6b7280;line-height:1.5;}
.warn-card{background:#fff;border:1px solid #fcd34d;border-left:4px solid #f59e0b;border-radius:10px;padding:12px 14px;margin:5px 0;}
.warn-title{font-size:13px;font-weight:600;color:#d97706;margin-bottom:3px;}
.warn-body{font-size:12px;color:#6b7280;line-height:1.5;}
.good-card{background:#fff;border:1px solid #bbf7d0;border-left:4px solid #22c55e;border-radius:10px;padding:12px 14px;margin:5px 0;}
.good-title{font-size:13px;font-weight:600;color:#16a34a;margin-bottom:3px;}
.good-body{font-size:12px;color:#6b7280;line-height:1.5;}

.cbox{background:#f8fafc;border:1px solid #e2e6ea;border-left:3px solid #ef4444;border-radius:6px;padding:8px 12px;margin:3px 0;font-size:12px;color:#374151;line-height:1.5;}
.section-title{font-size:17px;font-weight:700;color:#1e2330;margin-bottom:2px;}
.section-sub{font-size:12px;color:#94a3b8;margin-bottom:14px;}
.tag{display:inline-block;background:#eef2ff;border-radius:999px;padding:2px 9px;font-size:11px;color:#6366f1;margin:2px;}
.stTabs [data-baseweb="tab-list"]{background:transparent!important;border-bottom:1px solid #e2e6ea!important;}
.stTabs [data-baseweb="tab"]{background:transparent!important;color:#94a3b8!important;font-size:13px!important;font-weight:500!important;}
.stTabs [aria-selected="true"]{color:#1e2330!important;border-bottom:2px solid #6366f1!important;font-weight:600!important;}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:3px;}
</style>
""", unsafe_allow_html=True)

# ── Plotly 테마 ──
def plt(fig, title="", h=340):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#f8fafc",
        font=dict(family="Noto Sans KR,sans-serif", color="#374151", size=12),
        title=dict(text=title, font=dict(size=13, color="#1e2330")) if title else None,
        xaxis=dict(gridcolor="#e2e6ea", linecolor="#e2e6ea", tickfont=dict(size=11, color="#6b7280")),
        yaxis=dict(gridcolor="#e2e6ea", linecolor="#e2e6ea", tickfont=dict(size=11, color="#6b7280")),
        legend=dict(bgcolor="rgba(255,255,255,.9)", bordercolor="#e2e6ea", borderwidth=1, font=dict(size=11)),
        colorway=["#6366f1","#22c55e","#f59e0b","#ef4444","#06b6d4","#ec4899","#84cc16","#8b5cf6"],
        margin=dict(l=36, r=16, t=36 if title else 16, b=36), height=h,
        hoverlabel=dict(bgcolor="#fff", bordercolor="#e2e6ea", font=dict(color="#1e2330", size=12)),
    )
    return fig

CH_COLOR = {"CALL-IN":"#6366f1","CALL-OB":"#06b6d4","CHAT":"#22c55e","게시판":"#f59e0b"}

ISSUE_PATTERNS = [
    ("습관어/말버릇",   r"습관어|말버릇|아무래도|저희쪽|구요~"),
    ("존칭 오류",       r"존칭|요조체|비정중|반말|말투|셨을까요"),
    ("토막말/도치",     r"토막말|도치"),
    ("인사 누락",       r"인사.*누락|끝인사.*누락|첫인사.*누락"),
    ("양해 누락",       r"양해.*누락|양해.*미이행"),
    ("호응 부족",       r"호응.*누락|호응.*부족|적극적.*호응"),
    ("안내 오류",       r"누락|오안내|미안내"),
    ("이력 미기재",     r"이력.*누락|이력.*미기재|기재.*누락"),
    ("이모티콘",        r"이모티콘|이모지"),
    ("두괄식 미이행",   r"두괄식"),
    ("문법/띄어쓰기",   r"띄어쓰기|문법|오타|맞춤법"),
    ("처리 오류",       r"처리.*오류|처리.*지연|약속.*미준수"),
    ("첨부 누락",       r"사진.*누락|첨부.*누락"),
    ("문의파악 미흡",   r"문의.*파악|재질의|재문의"),
    ("가독성",          r"가독성|문단.*나눔|엔터"),
]
ROOT_CAUSE = {
    "습관어/말버릇":  "무의식적 언어 습관 → 모니터링 후 의식적 교정",
    "존칭 오류":      "경어법 이해 부족 → 표준 멘트 암기 훈련",
    "토막말/도치":    "문장 구성 습관 → 완전한 문장 말하기 연습",
    "인사 누락":      "프로세스 누락 → 체크리스트 활용",
    "양해 누락":      "고객 공감 부족 → 양해 멘트 패턴 훈련",
    "호응 부족":      "경청 자세 부족 → 적극적 호응 연습",
    "안내 오류":      "업무 숙지 미흡 → 주요 안내 재교육",
    "이력 미기재":    "기록 습관 부재 → 상담 후 이력 루틴화",
    "이모티콘":       "채널 예절 이해 부족 → 공식 채널 예절 교육",
    "두괄식 미이행":  "설명 구조 습관 → 결론 먼저 말하는 훈련",
    "문법/띄어쓰기":  "문서 작성 능력 부족 → 글쓰기 교육",
    "처리 오류":      "업무 정확도 부족 → 프로세스 재확인",
    "첨부 누락":      "체크리스트 미활용 → 처리 단계 확인 습관화",
    "문의파악 미흡":  "청취 이해 부족 → 문의 요약 확인 습관 훈련",
    "가독성":         "채팅 문서화 부족 → 구조화 글쓰기 교육",
    "기타":           "개별 피드백 참고 후 맞춤 교육",
}

def classify_issue(text):
    res = [l for l, p in ISSUE_PATTERNS if re.search(p, str(text))]
    return res if res else ["기타"]

def _num(v):
    if v is None: return None
    if isinstance(v, float) and np.isnan(v): return None
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(",","").strip())
    except: return None

def _s(v):
    if v is None: return ""
    s = str(v).strip()
    return "" if s in ("None","nan","NaN","") else s

def period_key(p):
    m = re.search(r"(\d+)년\s*(\d+)월(\d+)차", str(p))
    return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else (9999,0,0)

def comp_rate(scores, maxes):
    v = [(s,m) for s,m in zip(scores,maxes) if s is not None and m is not None and m>0]
    if not v: return None
    return round(sum(1 for s,m in v if s>=m)/len(v)*100, 1)

def kpi(label, value, sub=""):
    st.markdown(f'''<div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>''', unsafe_allow_html=True)

def sh(title, sub=""):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
    if sub: st.markdown(f'<div class="section-sub">{sub}</div>', unsafe_allow_html=True)

# ── 채널별 독립 항목 정의 ──
CALL_ITEMS = {
    "첫인사":(8,2.5),"정보확인":(11,5),"끝인사":(14,2.5),
    "인사톤":(17,2.5),"통화종료":(20,2.5),"음성숙련도":(23,5),
    "감정연출":(26,5),"양해":(29,5),"즉각호응":(32,5),
    "대기":(35,5),"언어표현":(38,5),"경청":(41,5),
    "문의파악":(44,5),"맞춤설명":(47,5),"정확한안내":(50,10),
    "프로세스":(53,10),"전산처리":(56,10),"상담이력":(59,10),
}
CHAT_ITEMS = {
    "첫인사":(6,5),"정보확인":(9,5),"끝인사":(12,5),
    "양해":(15,5),"즉각호응":(18,5),"대기":(21,5),
    "언어표현":(24,10),"가독성":(27,5),"문의파악":(30,10),
    "맞춤설명":(33,10),"정확한안내":(36,10),"프로세스":(39,5),
    "전산처리":(42,10),"상담이력":(45,10),
}
BOARD_ITEMS = {
    "첫인사":(7,2.5),"끝인사":(10,2.5),"언어표현":(13,10),
    "양해":(16,10),"가독성":(19,10),"문의파악":(22,10),
    "맞춤설명":(25,10),"정확한안내":(28,10),"프로세스":(31,10),
    "전산처리":(34,10),"상담이력":(37,15),
}

@st.cache_data(ttl=300)
def load_data():
    def parse(url, items_map, agent_col, period_col, inq_cols,
               total_col, eval_col, bonus_col, pen_col, inout_col=None, skip=3):
        try:
            raw = pd.read_csv(url, header=None)
        except Exception as e:
            st.error(f"로딩 실패: {e}  →  구글 시트 공유(뷰어) 설정 확인")
            return pd.DataFrame()
        records = []
        for _, row in raw.iloc[skip:].iterrows():
            agent = _s(row.iloc[agent_col] if agent_col < len(row) else None)
            if not agent: continue
            rec = {
                "상담사": agent,
                "평가기간": _s(row.iloc[period_col]),
                "문의유형대": _s(row.iloc[inq_cols[0]]),
                "문의유형중": _s(row.iloc[inq_cols[1]]),
                "문의유형소": _s(row.iloc[inq_cols[2]]),
                "평가자": _s(row.iloc[eval_col] if eval_col < len(row) else None),
                "가점": _num(row.iloc[bonus_col] if bonus_col < len(row) else None),
                "감점": _num(row.iloc[pen_col] if pen_col < len(row) else None),
                "평가대상": _s(row.iloc[inout_col] if inout_col is not None and inout_col < len(row) else None),
            }
            tr = row.iloc[total_col] if total_col < len(row) else None
            if pd.isna(tr) if tr is not None else True or str(tr).startswith("="):
                total = sum(_num(row.iloc[c]) or 0 for c,_ in items_map.values())
                total += rec["가점"] or 0
                total -= rec["감점"] or 0
            else:
                total = _num(tr)
            rec["TOTAL"] = total
            for iname,(ci,mx) in items_map.items():
                rec[f"{iname}_점수"]    = _num(row.iloc[ci]) if ci < len(row) else None
                rec[f"{iname}_최대"]    = mx
                rec[f"{iname}_감점사유"] = _s(row.iloc[ci+1] if ci+1 < len(row) else None)
                rec[f"{iname}_미차감"]   = _s(row.iloc[ci+2] if ci+2 < len(row) else None)
            records.append(rec)
        return pd.DataFrame(records) if records else pd.DataFrame()

    df_call = parse(gsheet_csv_url(SHEET_ID, GID_CALL), CALL_ITEMS,
                    3,1,[5,6,7],64,65,62,63, inout_col=2)
    if len(df_call):
        df_call["채널"] = df_call["평가대상"].apply(
            lambda x: "CALL-IN" if str(x).upper().strip()=="IN"
                      else "CALL-OB" if str(x).upper().strip()=="OB" else "CALL")

    df_chat = parse(gsheet_csv_url(SHEET_ID, GID_CHAT), CHAT_ITEMS,
                    1,0,[3,4,5],50,51,48,49)
    if len(df_chat): df_chat["채널"] = "CHAT"

    df_board = parse(gsheet_csv_url(SHEET_ID, GID_BOARD), BOARD_ITEMS,
                     2,1,[4,5,6],42,43,40,41)
    if len(df_board): df_board["채널"] = "게시판"

    # 채널별 독립 데이터프레임 반환 + 전체 합산
    df_all = pd.concat([df_call, df_chat, df_board], ignore_index=True)
    df_all["기간_정렬"] = df_all["평가기간"].apply(period_key)
    df_all = df_all.sort_values("기간_정렬").reset_index(drop=True)
    return df_call, df_chat, df_board, df_all

with st.spinner("구글 시트 데이터 로딩 중..."):
    df_call, df_chat, df_board, df_all = load_data()

if df_all.empty:
    st.error("데이터 없음 — 구글 시트 공유(뷰어) 설정을 확인하세요.")
    st.stop()

all_periods  = sorted(df_all["평가기간"].unique(), key=period_key)
all_channels = sorted(df_all["채널"].unique().tolist())
all_agents   = sorted(df_all["상담사"].unique().tolist())
all_evals    = sorted(df_all["평가자"].dropna().unique().tolist())

# ══════════════════════════════
# 상단 기간 필터 (항상 표시)
# ══════════════════════════════
with st.container():
    fc1, fc2, fc3, fc4 = st.columns([3,2,2,1])
    with fc1:
        sel_period = st.multiselect("평가기간", all_periods, default=all_periods, label_visibility="collapsed",
                                     placeholder="평가기간 선택")
    with fc2:
        sel_channel = st.multiselect("채널", all_channels, default=all_channels, label_visibility="collapsed",
                                      placeholder="채널 선택")
    with fc3:
        sel_agent = st.multiselect("상담사", all_agents, default=all_agents, label_visibility="collapsed",
                                    placeholder="상담사 선택")
    with fc4:
        if st.button("새로고침", use_container_width=True):
            st.cache_data.clear(); st.rerun()

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

def flt(df):
    o = df.copy()
    if sel_period:  o = o[o["평가기간"].isin(sel_period)]
    if sel_channel: o = o[o["채널"].isin(sel_channel)]
    if sel_agent:   o = o[o["상담사"].isin(sel_agent)]
    return o

df = flt(df_all)
df_c = flt(df_call) if len(df_call) else pd.DataFrame()
df_h = flt(df_chat) if len(df_chat) else pd.DataFrame()
df_b = flt(df_board) if len(df_board) else pd.DataFrame()

# ══════════════════════════════
# 사이드바 메뉴
# ══════════════════════════════
with st.sidebar:
    st.markdown("<div style='font-size:15px;font-weight:700;color:#1e2330;padding-bottom:14px;margin-bottom:14px;border-bottom:1px solid #e2e6ea'>QA Monitoring</div>", unsafe_allow_html=True)
    page = st.radio("메뉴", [
        "개요 및 위기감지",
        "점수 트렌드",
        "문의유형 분석",
        "채널별 항목 분석",
        "감점 분석",
        "상담사 종합평가",
        "이행률 분석",
        "평가자 분석",
    ], label_visibility="collapsed")
    st.markdown(f"<div style='font-size:11px;color:#94a3b8;margin-top:16px;padding-top:12px;border-top:1px solid #e2e6ea'>총 {len(df_all)}건 로드됨</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
# 공통 유틸 — 채널별 항목 점수 추출
# ══════════════════════════════════════════════════════════
def get_channel_item_scores(ch_df, items_map):
    """채널 전용 df + 해당 채널 항목맵으로 감점/이행률 계산"""
    if ch_df is None or len(ch_df) == 0: return pd.DataFrame()
    rows = []
    for iname, (_, mx) in items_map.items():
        sc_col = f"{iname}_점수"
        if sc_col not in ch_df.columns: continue
        sc = ch_df[sc_col].dropna().tolist()
        mx_list = [mx] * len(sc)
        cr = comp_rate(sc, mx_list)
        deducted = sum(1 for s in sc if s < mx)
        rows.append({"항목": iname, "최대점수": mx, "평균점수": np.mean(sc) if sc else 0,
                     "이행률": cr or 0, "감점건수": deducted, "총평가": len(sc)})
    return pd.DataFrame(rows)

def get_deduct_df(src_df, items_map):
    """감점 발생 레코드 추출"""
    recs = []
    for iname, (_, mx) in items_map.items():
        sc_col = f"{iname}_점수"
        rs_col = f"{iname}_감점사유"
        if sc_col not in src_df.columns: continue
        sub = src_df[[sc_col, rs_col, "상담사","평가기간","채널"]].copy()
        sub.columns = ["점수","감점사유","상담사","평가기간","채널"]
        sub["항목"] = iname; sub["최대"] = mx
        sub = sub[sub["점수"].notna()]
        sub = sub[sub["점수"] < mx]
        sub = sub[sub["감점사유"] != ""]
        recs.append(sub)
    return pd.concat(recs, ignore_index=True) if recs else pd.DataFrame()

# ══════════════════════════════
# 위기감지 핵심 계산
# ══════════════════════════════
def crisis_analysis():
    """상담사별 위기 신호 계산"""
    results = []
    for agent in df_all["상담사"].unique():
        adf = df_all[df_all["상담사"] == agent]
        adf_f = flt(adf)
        if len(adf_f) == 0: continue

        # 기간별 점수 추이
        periods_sorted = sorted(adf_f["평가기간"].unique(), key=period_key)
        period_avgs = [adf_f[adf_f["평가기간"]==p]["TOTAL"].mean() for p in periods_sorted]

        # 최근 트렌드 (마지막 2기간 비교)
        trend = "유지"
        trend_val = 0.0
        if len(period_avgs) >= 2:
            trend_val = period_avgs[-1] - period_avgs[-2]
            if trend_val <= -3: trend = "악화"
            elif trend_val >= 3: trend = "개선"

        # 감점 이슈 집계
        all_issues = []
        for ch_df, items_map in [(df_c, CALL_ITEMS),(df_h, CHAT_ITEMS),(df_b, BOARD_ITEMS)]:
            if ch_df is None or len(ch_df)==0: continue
            a_ch = ch_df[ch_df["상담사"]==agent]
            a_ch = flt(a_ch)
            for iname in items_map:
                dc = f"{iname}_감점사유"
                if dc not in a_ch.columns: continue
                for r in a_ch[dc].dropna():
                    if r and r not in ("None","nan",""):
                        all_issues.extend(classify_issue(r))

        ic = Counter(all_issues)
        repeat = {k:v for k,v in ic.items() if v>=2}
        top_issue = ic.most_common(1)[0] if ic else ("없음", 0)

        # 채널별 평균
        ch_scores = {}
        for ch in adf_f["채널"].unique():
            avg = adf_f[adf_f["채널"]==ch]["TOTAL"].mean()
            ch_scores[ch] = round(avg, 1)

        results.append({
            "상담사": agent,
            "전체평균": round(adf_f["TOTAL"].mean(), 1),
            "평가건수": len(adf_f),
            "트렌드": trend,
            "트렌드값": round(trend_val, 1),
            "반복이슈수": len(repeat),
            "최다이슈": top_issue[0],
            "최다이슈횟수": top_issue[1],
            "채널별점수": ch_scores,
            "이슈카운터": ic,
        })
    return pd.DataFrame(results)

# ══════════════════════════════════════════════════════════════════
# 개요 및 위기감지
# ══════════════════════════════════════════════════════════════════
if page == "개요 및 위기감지":
    sh("개요 및 위기감지", "전체 현황 요약 · 집중 코칭이 필요한 상담사 자동 감지")

    # KPI
    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: kpi("총 평가건수", f"{len(df):,}건")
    with c2: kpi("전체 평균", f"{df['TOTAL'].dropna().mean():.1f}점" if len(df) else "-")
    with c3: kpi("상담사 수", f"{df['상담사'].nunique()}명")
    with c4: kpi("평가 기간", f"{df['평가기간'].nunique()}개")
    with c5: kpi("활성 채널", f"{df['채널'].nunique()}개")

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # 위기감지
    crisis_df = crisis_analysis()
    if len(crisis_df):
        danger  = crisis_df[(crisis_df["트렌드"]=="악화") | (crisis_df["반복이슈수"]>=3)]
        warning = crisis_df[(crisis_df["트렌드"]=="유지") & (crisis_df["반복이슈수"]>=2) & (~crisis_df["상담사"].isin(danger["상담사"]))]
        good    = crisis_df[crisis_df["트렌드"]=="개선"]

        col_a, col_b = st.columns([1,2])
        with col_a:
            st.markdown("#### 위기 감지 알림")
            if len(danger):
                for _, r in danger.iterrows():
                    ch_txt = "  ".join([f"{k} {v}점" for k,v in r["채널별점수"].items()])
                    st.markdown(f'''<div class="alert-card">
                        <div class="alert-title">즉시 코칭 필요 — {r["상담사"]}</div>
                        <div class="alert-body">
                            평균 {r["전체평균"]}점 · 트렌드 {r["트렌드값"]:+.1f}점<br>
                            반복이슈 {r["반복이슈수"]}개 · 최다: {r["최다이슈"]} ({r["최다이슈횟수"]}회)<br>
                            채널별: {ch_txt}
                        </div></div>''', unsafe_allow_html=True)
            if len(warning):
                for _, r in warning.iterrows():
                    st.markdown(f'''<div class="warn-card">
                        <div class="warn-title">모니터링 필요 — {r["상담사"]}</div>
                        <div class="warn-body">반복이슈 {r["반복이슈수"]}개 · 최다: {r["최다이슈"]}</div>
                    </div>''', unsafe_allow_html=True)
            if len(good):
                for _, r in good.iterrows():
                    st.markdown(f'''<div class="good-card">
                        <div class="good-title">개선 중 — {r["상담사"]}</div>
                        <div class="good-body">트렌드 {r["트렌드값"]:+.1f}점 상승</div>
                    </div>''', unsafe_allow_html=True)
            if not len(danger) and not len(warning):
                st.success("현재 위기 상담사 없음")

        with col_b:
            st.markdown("#### 상담사별 채널 점수 비교")
            ch_rows = []
            for _, r in crisis_df.iterrows():
                for ch, sc in r["채널별점수"].items():
                    ch_rows.append({"상담사": r["상담사"], "채널": ch, "평균점수": sc})
            if ch_rows:
                ch_df2 = pd.DataFrame(ch_rows)
                fig = px.bar(ch_df2, x="상담사", y="평균점수", color="채널",
                             barmode="group", color_discrete_map=CH_COLOR)
                plt(fig, "상담사별 채널 점수", 320)
                fig.add_hline(y=df["TOTAL"].mean(), line_dash="dash", line_color="#94a3b8",
                              annotation_text=f"전체평균", annotation_font_color="#94a3b8")
                st.plotly_chart(fig, use_container_width=True)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # 항목별 차감 기여자 (채널별 독립)
    st.markdown("#### 항목별 차감 기여 상담사 (채널 분리)")
    tab_ci, tab_ch, tab_cb = st.tabs(["CALL (IN+OB)", "CHAT", "게시판"])

    for tab, ch_src, items_map, label in [
        (tab_ci, df_c, CALL_ITEMS, "CALL"),
        (tab_ch, df_h, CHAT_ITEMS, "CHAT"),
        (tab_cb, df_b, BOARD_ITEMS, "게시판"),
    ]:
        with tab:
            if ch_src is None or len(ch_src)==0:
                st.info(f"{label} 데이터 없음")
                continue
            recs2 = []
            for iname,(ci,mx) in items_map.items():
                sc_col=f"{iname}_점수"
                if sc_col not in ch_src.columns: continue
                for agent in ch_src["상담사"].unique():
                    adf2 = ch_src[ch_src["상담사"]==agent]
                    sc = adf2[sc_col].dropna()
                    deducted = (sc < mx).sum()
                    if deducted > 0:
                        recs2.append({"상담사":agent,"항목":iname,"감점건수":int(deducted)})
            if recs2:
                d2 = pd.DataFrame(recs2)
                pvt = d2.pivot_table(index="상담사", columns="항목", values="감점건수", fill_value=0)
                fig = go.Figure(data=go.Heatmap(
                    z=pvt.values, x=pvt.columns.tolist(), y=pvt.index.tolist(),
                    colorscale=[[0,"#f8fafc"],[0.3,"#fef3c7"],[1,"#ef4444"]],
                    text=pvt.values, texttemplate="%{text}",
                    textfont=dict(size=11, color="#374151"),
                ))
                plt(fig, f"{label} 항목별 감점 횟수", max(260, len(pvt)*38))
                fig.update_layout(xaxis=dict(tickangle=-30))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("감점 데이터 없음")

    # 기간별 점수 추이
    st.markdown("#### 기간별 평균 점수 추이")
    grp = df.groupby(["평가기간","채널","기간_정렬"])["TOTAL"].mean().reset_index().sort_values("기간_정렬")
    fig = px.line(grp, x="평가기간", y="TOTAL", color="채널",
                  markers=True, color_discrete_map=CH_COLOR, labels={"TOTAL":"평균 점수"})
    plt(fig, "", 280)
    fig.update_traces(line=dict(width=2.5), marker=dict(size=6))
    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 점수 트렌드
# ══════════════════════════════════════════════════════════════════
elif page == "점수 트렌드":
    sh("점수 트렌드", "기간별·상담사별·채널별 점수 변화")
    tab1,tab2,tab3 = st.tabs(["기간별 추이","상담사별 비교","Box Plot"])
    with tab1:
        grp = df.groupby(["평가기간","상담사","기간_정렬"])["TOTAL"].mean().reset_index().sort_values("기간_정렬")
        fig = px.line(grp, x="평가기간", y="TOTAL", color="상담사", markers=True, labels={"TOTAL":"평균 점수"})
        plt(fig, "상담사별 점수 추이", 380)
        fig.update_traces(line=dict(width=2), marker=dict(size=5))
        st.plotly_chart(fig, use_container_width=True)
        grp2 = df.groupby(["평가기간","채널","기간_정렬"])["TOTAL"].mean().reset_index().sort_values("기간_정렬")
        fig2 = px.line(grp2, x="평가기간", y="TOTAL", color="채널", markers=True,
                       color_discrete_map=CH_COLOR, labels={"TOTAL":"평균 점수"})
        plt(fig2, "채널별 점수 추이", 280)
        fig2.update_traces(line=dict(width=2.5), marker=dict(size=6))
        st.plotly_chart(fig2, use_container_width=True)
    with tab2:
        aa = df.groupby(["상담사","채널"])["TOTAL"].mean().reset_index()
        fig = px.bar(aa, x="상담사", y="TOTAL", color="채널", barmode="group",
                     color_discrete_map=CH_COLOR, labels={"TOTAL":"평균 점수"})
        plt(fig, "상담사 × 채널별 평균 점수", 360)
        fig.add_hline(y=df["TOTAL"].mean(), line_dash="dash", line_color="#94a3b8")
        st.plotly_chart(fig, use_container_width=True)
    with tab3:
        fig = px.box(df, x="상담사", y="TOTAL", color="채널",
                     color_discrete_map=CH_COLOR, labels={"TOTAL":"점수"})
        plt(fig, "상담사별 점수 분포", 380)
        st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 문의유형 분석
# ══════════════════════════════════════════════════════════════════
elif page == "문의유형 분석":
    sh("문의유형 분석", "유형별 점수 히트맵 및 분포")
    tab1,tab2 = st.tabs(["유형 × 상담사 히트맵","유형별 건수/점수"])
    with tab1:
        pivot = df.groupby(["문의유형대","상담사"])["TOTAL"].mean().unstack(fill_value=np.nan).round(1)
        fig = go.Figure(data=go.Heatmap(
            z=pivot.values, x=pivot.columns.tolist(), y=pivot.index.tolist(),
            colorscale=[[0,"#ef4444"],[0.5,"#fef3c7"],[1,"#22c55e"]],
            text=[[f"{v:.1f}" if not np.isnan(v) else "" for v in row] for row in pivot.values],
            texttemplate="%{text}", textfont=dict(size=11, color="#374151"),
        ))
        plt(fig, "문의유형(대) × 상담사 평균점수", max(300, len(pivot)*35))
        st.plotly_chart(fig, use_container_width=True)
    with tab2:
        co1,co2 = st.columns(2)
        with co1:
            cnt = df["문의유형대"].value_counts().reset_index(); cnt.columns=["유형","건수"]
            fig = px.bar(cnt, x="유형", y="건수", color="건수", color_continuous_scale=["#e0e7ff","#6366f1"])
            plt(fig, "문의유형별 건수", 300); fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        with co2:
            avg = df.groupby("문의유형대")["TOTAL"].mean().reset_index(); avg.columns=["유형","평균점수"]
            fig = px.bar(avg, x="유형", y="평균점수", color="평균점수",
                         color_continuous_scale=["#ef4444","#22c55e"])
            plt(fig, "문의유형별 평균 점수", 300); fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 채널별 항목 분석 (채널별 독립)
# ══════════════════════════════════════════════════════════════════
elif page == "채널별 항목 분석":
    sh("채널별 항목 분석", "채널마다 다른 평가 항목을 독립적으로 분석합니다")

    ch_tab1, ch_tab2, ch_tab3 = st.tabs(["CALL (IN+OB)", "CHAT", "게시판"])

    for tab, ch_df_base, items_map, ch_label in [
        (ch_tab1, df_c, CALL_ITEMS, "CALL"),
        (ch_tab2, df_h, CHAT_ITEMS, "CHAT"),
        (ch_tab3, df_b, BOARD_ITEMS, "게시판"),
    ]:
        with tab:
            if ch_df_base is None or len(ch_df_base)==0:
                st.info(f"{ch_label} 데이터 없음")
                continue

            st.markdown(f"**{ch_label} 평가 항목별 이행률**")
            idf = get_channel_item_scores(ch_df_base, items_map)
            if len(idf):
                idf_sorted = idf.sort_values("이행률")
                fig = go.Figure(go.Bar(
                    x=idf_sorted["이행률"], y=idf_sorted["항목"], orientation="h",
                    marker_color=idf_sorted["이행률"].apply(
                        lambda x: "#22c55e" if x>=90 else "#f59e0b" if x>=75 else "#ef4444"),
                    text=idf_sorted["이행률"].apply(lambda x: f"{x:.1f}%"),
                    textposition="outside", textfont=dict(size=11, color="#374151"),
                ))
                plt(fig, f"{ch_label} 항목별 이행률", max(260, len(idf)*28))
                fig.update_layout(xaxis=dict(range=[0,108]))
                st.plotly_chart(fig, use_container_width=True)

            st.markdown(f"**{ch_label} 상담사별 × 항목별 달성률 히트맵**")
            heat_rows = {}
            for iname in items_map:
                sc_col = f"{iname}_점수"
                mx = items_map[iname][1]
                if sc_col not in ch_df_base.columns: continue
                for agent in ch_df_base["상담사"].unique():
                    adf3 = ch_df_base[ch_df_base["상담사"]==agent]
                    sc = adf3[sc_col].dropna()
                    if len(sc)==0: continue
                    cr = comp_rate(sc.tolist(), [mx]*len(sc))
                    heat_rows.setdefault(agent, {})[iname] = cr or 0

            if heat_rows:
                hdf = pd.DataFrame(heat_rows).T.fillna(0)
                fig = go.Figure(data=go.Heatmap(
                    z=hdf.values, x=hdf.columns.tolist(), y=hdf.index.tolist(),
                    colorscale=[[0,"#ef4444"],[0.5,"#fef3c7"],[1,"#22c55e"]],
                    zmin=0, zmax=100,
                    text=[[f"{v:.0f}%" for v in row] for row in hdf.values],
                    texttemplate="%{text}", textfont=dict(size=10, color="#374151"),
                ))
                plt(fig, f"{ch_label} 상담사 × 항목 이행률", max(260, len(hdf)*38))
                fig.update_layout(xaxis=dict(tickangle=-30))
                st.plotly_chart(fig, use_container_width=True)

            # 상담사별 기간별 항목 이행률 추이
            st.markdown(f"**{ch_label} 상담사별 항목 이행률 추이**")
            agent_sel2 = st.selectbox(f"상담사 선택 ({ch_label})", sorted(ch_df_base["상담사"].unique()), key=f"agent_{ch_label}")
            item_sel2 = st.multiselect(f"항목 선택 ({ch_label})", list(items_map.keys()),
                                        default=list(items_map.keys())[:4], key=f"item_{ch_label}")
            if item_sel2:
                adf4 = ch_df_base[ch_df_base["상담사"]==agent_sel2].copy()
                adf4["기간_정렬"] = adf4["평가기간"].apply(period_key)
                trend_rows = []
                for period in sorted(adf4["평가기간"].unique(), key=period_key):
                    pdf = adf4[adf4["평가기간"]==period]
                    for iname in item_sel2:
                        sc_col = f"{iname}_점수"
                        mx = items_map[iname][1]
                        if sc_col not in pdf.columns: continue
                        sc = pdf[sc_col].dropna().tolist()
                        cr = comp_rate(sc, [mx]*len(sc))
                        if cr is not None:
                            trend_rows.append({"기간":period,"항목":iname,"이행률":cr})
                if trend_rows:
                    tdf = pd.DataFrame(trend_rows)
                    fig = px.line(tdf, x="기간", y="이행률", color="항목", markers=True)
                    plt(fig, f"{agent_sel2} — {ch_label} 항목별 이행률 추이", 300)
                    fig.update_traces(line=dict(width=2), marker=dict(size=6))
                    fig.add_hline(y=90, line_dash="dot", line_color="#94a3b8", annotation_text="목표 90%")
                    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 감점 분석
# ══════════════════════════════════════════════════════════════════
elif page == "감점 분석":
    sh("감점 분석", "채널별 독립 감점 패턴 분석")
    ch_tab1, ch_tab2, ch_tab3 = st.tabs(["CALL", "CHAT", "게시판"])
    for tab, ch_src, items_map, lbl in [
        (ch_tab1, df_c, CALL_ITEMS, "CALL"),
        (ch_tab2, df_h, CHAT_ITEMS, "CHAT"),
        (ch_tab3, df_b, BOARD_ITEMS, "게시판"),
    ]:
        with tab:
            if ch_src is None or len(ch_src)==0:
                st.info(f"{lbl} 데이터 없음"); continue
            ded = get_deduct_df(ch_src, items_map)
            if len(ded)==0:
                st.info("감점 데이터 없음"); continue
            c1,c2 = st.columns(2)
            with c1:
                ic2 = ded["항목"].value_counts().reset_index(); ic2.columns=["항목","건수"]
                fig = px.bar(ic2, x="건수", y="항목", orientation="h", color="건수",
                             color_continuous_scale=["#fee2e2","#ef4444"])
                plt(fig, f"{lbl} 항목별 감점 건수", max(260, len(ic2)*30))
                fig.update_coloraxes(showscale=False)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                ad = ded.groupby(["상담사","항목"]).size().unstack(fill_value=0)
                fig = go.Figure(data=go.Heatmap(
                    z=ad.values, x=ad.columns.tolist(), y=ad.index.tolist(),
                    colorscale=[[0,"#f8fafc"],[1,"#ef4444"]],
                    text=ad.values, texttemplate="%{text}",
                    textfont=dict(size=11, color="#374151"),
                ))
                plt(fig, f"{lbl} 상담사 × 항목 감점", max(260, len(ad)*38))
                st.plotly_chart(fig, use_container_width=True)
            st.markdown("**감점 원본 코멘트**")
            show = ded[["상담사","항목","점수","최대","감점사유"]].copy()
            show = show.replace("", np.nan).dropna(subset=["감점사유"])
            st.dataframe(show.reset_index(drop=True), use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 상담사 종합평가 (위기감지 + 채널별 약점 통합)
# ══════════════════════════════════════════════════════════════════
elif page == "상담사 종합평가":
    sh("상담사 종합평가", "레이더 차트 · 채널별 약점 · 반복 이슈 · 코칭 포인트")
    agent_sel = st.selectbox("상담사 선택", all_agents)
    adf_all = df_all[df_all["상담사"]==agent_sel]

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi("전체 평균", f"{adf_all['TOTAL'].mean():.1f}점" if len(adf_all) else "-")
    with c2: kpi("평가 건수", f"{len(adf_all)}건")
    with c3: kpi("활동 채널", ", ".join(adf_all["채널"].unique()))
    with c4:
        ps = sorted(adf_all["평가기간"].unique(), key=period_key)
        if len(ps)>=2:
            v1 = adf_all[adf_all["평가기간"]==ps[-1]]["TOTAL"].mean()
            v0 = adf_all[adf_all["평가기간"]==ps[-2]]["TOTAL"].mean()
            diff = v1-v0
            kpi("최근 트렌드", f"{diff:+.1f}점", "전기 대비")
        else:
            kpi("최근 트렌드", "-")

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # 채널별 약점 분석
    st.markdown("#### 채널별 항목 약점")
    ch_tabs = st.tabs(["CALL-IN","CALL-OB","CHAT","게시판"])
    ch_map = {"CALL-IN":(df_c,CALL_ITEMS),"CALL-OB":(df_c,CALL_ITEMS),
              "CHAT":(df_h,CHAT_ITEMS),"게시판":(df_b,BOARD_ITEMS)}
    for i, ch_name in enumerate(["CALL-IN","CALL-OB","CHAT","게시판"]):
        with ch_tabs[i]:
            ch_src2, items_m = ch_map[ch_name]
            if ch_src2 is None or len(ch_src2)==0:
                st.info("데이터 없음"); continue
            if ch_name in ("CALL-IN","CALL-OB"):
                a_ch = ch_src2[(ch_src2["상담사"]==agent_sel)&(ch_src2["채널"]==ch_name)]
            else:
                a_ch = ch_src2[ch_src2["상담사"]==agent_sel]
            if len(a_ch)==0:
                st.info(f"{ch_name} 평가 없음"); continue
            rows_ch = []
            for iname,(ci,mx) in items_m.items():
                sc = a_ch[f"{iname}_점수"].dropna().tolist() if f"{iname}_점수" in a_ch.columns else []
                if not sc: continue
                cr = comp_rate(sc,[mx]*len(sc))
                rows_ch.append({"항목":iname,"이행률":cr or 0,"평균점수":np.mean(sc),"최대":mx})
            if rows_ch:
                rdf = pd.DataFrame(rows_ch).sort_values("이행률")
                fig = go.Figure(go.Bar(
                    x=rdf["이행률"], y=rdf["항목"], orientation="h",
                    marker_color=rdf["이행률"].apply(lambda x:"#22c55e" if x>=90 else "#f59e0b" if x>=75 else "#ef4444"),
                    text=rdf["이행률"].apply(lambda x:f"{x:.1f}%"),
                    textposition="outside", textfont=dict(size=11,color="#374151"),
                ))
                plt(fig, f"{agent_sel} — {ch_name} 항목 이행률", max(240, len(rdf)*28))
                fig.update_layout(xaxis=dict(range=[0,108]))
                st.plotly_chart(fig, use_container_width=True)

    # 기간별 점수
    trend2 = adf_all.groupby(["평가기간","기간_정렬"])["TOTAL"].mean().reset_index().sort_values("기간_정렬")
    fig2 = px.line(trend2, x="평가기간", y="TOTAL", markers=True, labels={"TOTAL":"평균 점수"})
    plt(fig2, f"{agent_sel} 기간별 점수 추이", 260)
    fig2.update_traces(line=dict(width=2.5,color="#6366f1"), marker=dict(size=7,color="#6366f1"))
    st.plotly_chart(fig2, use_container_width=True)

    # 반복 이슈
    st.markdown("#### 반복 이슈 및 코칭 포인트")
    all_issues2 = []
    issue_with_text = {}
    for ch_src3, items_m2 in [(df_c,CALL_ITEMS),(df_h,CHAT_ITEMS),(df_b,BOARD_ITEMS)]:
        if ch_src3 is None or len(ch_src3)==0: continue
        a3 = ch_src3[ch_src3["상담사"]==agent_sel]
        for iname in items_m2:
            dc = f"{iname}_감점사유"
            if dc not in a3.columns: continue
            for r in a3[dc].dropna():
                if not r or r in ("None","nan",""): continue
                cats = classify_issue(r)
                all_issues2.extend(cats)
                for cat in cats:
                    issue_with_text.setdefault(cat, []).append(r)

    ic3 = Counter(all_issues2)
    repeat3 = {k:v for k,v in ic3.items() if v>=2}
    if repeat3:
        for issue, cnt in sorted(repeat3.items(), key=lambda x:-x[1]):
            texts_ex = list(dict.fromkeys(issue_with_text.get(issue,[])))[:3]
            root = ROOT_CAUSE.get(issue,"개별 피드백 참고 후 맞춤 교육")
            boxes = "".join([f'<div class="cbox">{t}</div>' for t in texts_ex])
            st.markdown(f'''<div class="alert-card">
                <div class="alert-title">{issue} — {cnt}회 반복</div>
                <div class="alert-body">📌 {root}<br>{boxes}</div>
            </div>''', unsafe_allow_html=True)
    else:
        st.success("반복 이슈 없음")

# ══════════════════════════════════════════════════════════════════
# 이행률 분석
# ══════════════════════════════════════════════════════════════════
elif page == "이행률 분석":
    sh("이행률 분석", "채널별 독립 이행률 — 항목별·상담사별·기간별")
    ch_tab1, ch_tab2, ch_tab3 = st.tabs(["CALL","CHAT","게시판"])
    for tab, ch_src4, items_m3, lbl2 in [
        (ch_tab1, df_c, CALL_ITEMS, "CALL"),
        (ch_tab2, df_h, CHAT_ITEMS, "CHAT"),
        (ch_tab3, df_b, BOARD_ITEMS, "게시판"),
    ]:
        with tab:
            if ch_src4 is None or len(ch_src4)==0:
                st.info(f"{lbl2} 데이터 없음"); continue

            t1,t2,t3 = st.tabs([f"항목별",f"상담사별",f"기간별"])
            with t1:
                idf2 = get_channel_item_scores(ch_src4, items_m3)
                if len(idf2):
                    s2 = idf2.sort_values("이행률")
                    fig = go.Figure(go.Bar(
                        x=s2["이행률"], y=s2["항목"], orientation="h",
                        marker_color=s2["이행률"].apply(lambda x:"#22c55e" if x>=90 else "#f59e0b" if x>=75 else "#ef4444"),
                        text=s2["이행률"].apply(lambda x:f"{x:.1f}%"), textposition="outside",
                        textfont=dict(size=11,color="#374151"),
                    ))
                    plt(fig, f"{lbl2} 항목별 이행률", max(260, len(s2)*28))
                    fig.update_layout(xaxis=dict(range=[0,108]))
                    st.plotly_chart(fig, use_container_width=True)
            with t2:
                agent_comp2 = {}
                for ag in ch_src4["상담사"].unique():
                    adf5 = ch_src4[ch_src4["상담사"]==ag]
                    vals5 = [comp_rate(adf5[f"{i}_점수"].dropna().tolist(),[mx]*len(adf5[f"{i}_점수"].dropna()))
                             for i,(ci,mx) in items_m3.items() if f"{i}_점수" in adf5.columns]
                    vals5 = [v for v in vals5 if v is not None]
                    if vals5: agent_comp2[ag] = round(np.mean(vals5),1)
                if agent_comp2:
                    acd = pd.DataFrame(list(agent_comp2.items()),columns=["상담사","이행률"]).sort_values("이행률")
                    fig = px.bar(acd, x="이행률", y="상담사", orientation="h", color="이행률",
                                 color_continuous_scale=["#ef4444","#22c55e"])
                    plt(fig, f"{lbl2} 상담사별 이행률", max(260, len(acd)*30))
                    fig.update_coloraxes(showscale=False)
                    avg5 = acd["이행률"].mean()
                    fig.add_vline(x=avg5, line_dash="dash", line_color="#94a3b8",
                                  annotation_text=f"평균 {avg5:.1f}%")
                    st.plotly_chart(fig, use_container_width=True)
            with t3:
                pc2 = {}
                for period2 in ch_src4["평가기간"].unique():
                    pdf2 = ch_src4[ch_src4["평가기간"]==period2]
                    sk2 = period_key(period2)
                    vals6 = [comp_rate(pdf2[f"{i}_점수"].dropna().tolist(),[mx]*len(pdf2[f"{i}_점수"].dropna()))
                             for i,(ci,mx) in items_m3.items() if f"{i}_점수" in pdf2.columns]
                    vals6 = [v for v in vals6 if v is not None]
                    if vals6: pc2[period2] = (round(np.mean(vals6),1), sk2)
                if pc2:
                    pcd = pd.DataFrame([(k,v[0],v[1]) for k,v in pc2.items()],columns=["기간","이행률","정렬"]).sort_values("정렬")
                    fig = px.line(pcd, x="기간", y="이행률", markers=True)
                    plt(fig, f"{lbl2} 기간별 이행률 추이", 280)
                    fig.update_traces(line=dict(width=2.5,color="#6366f1"), marker=dict(size=7))
                    fig.add_hline(y=90, line_dash="dot", line_color="#22c55e", annotation_text="목표 90%")
                    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 평가자 분석
# ══════════════════════════════════════════════════════════════════
elif page == "평가자 분석":
    sh("평가자 행동 분석", "엄격도·일관성·편향 비교")
    tab1,tab2,tab3 = st.tabs(["점수 분포","엄격도 편향","평가자 × 상담사"])
    with tab1:
        fig = px.box(df, x="평가자", y="TOTAL", color="평가자", labels={"TOTAL":"점수"})
        plt(fig, "평가자별 점수 분포", 340); fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
        es = df.groupby("평가자")["TOTAL"].agg(["mean","std","count","min","max"]).round(2).reset_index()
        es.columns=["평가자","평균","표준편차","평가건수","최솟값","최댓값"]
        st.dataframe(es, use_container_width=True)
    with tab2:
        ln = df.groupby("평가자")["TOTAL"].mean().reset_index(); ln.columns=["평가자","평균점수"]
        ln["편향"]=(ln["평균점수"]-df["TOTAL"].mean()).round(2)
        ln["방향"]=ln["편향"].apply(lambda x:"관대" if x>0 else "엄격")
        fig = go.Figure(go.Bar(
            x=ln["편향"], y=ln["평가자"], orientation="h",
            marker_color=ln["방향"].apply(lambda x:"#22c55e" if x=="관대" else "#ef4444"),
            text=ln["편향"].astype(str), textposition="outside"
        ))
        plt(fig, "평가자별 점수 편향", 280)
        fig.add_vline(x=0, line_color="#94a3b8")
        st.plotly_chart(fig, use_container_width=True)
    with tab3:
        pivot3 = df.groupby(["평가자","상담사"])["TOTAL"].mean().unstack(fill_value=np.nan).round(1)
        fig = go.Figure(data=go.Heatmap(
            z=pivot3.values, x=pivot3.columns.tolist(), y=pivot3.index.tolist(),
            colorscale=[[0,"#ef4444"],[0.5,"#fef3c7"],[1,"#22c55e"]],
            text=[[f"{v:.1f}" if not np.isnan(v) else "" for v in row] for row in pivot3.values],
            texttemplate="%{text}", textfont=dict(size=11,color="#374151"),
        ))
        plt(fig, "평가자 × 상담사 평균 점수", max(280, len(pivot3)*45))
        st.plotly_chart(fig, use_container_width=True)

st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)
st.markdown("<div style='border-top:1px solid #e2e6ea;padding:14px 0;text-align:center;color:#94a3b8;font-size:11px'>QA Monitoring Dashboard · Google Sheets 연동</div>", unsafe_allow_html=True)
