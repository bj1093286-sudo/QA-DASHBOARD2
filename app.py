import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from collections import Counter
import re
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
/* 사이드바 멀티셀렉트 선택 태그 흰색 글씨 */
section[data-testid="stSidebar"] [data-baseweb="tag"] span{color:#ffffff!important;}
section[data-testid="stSidebar"] [data-baseweb="tag"]{background:#6366f1!important;}
section[data-testid="stSidebar"] [data-baseweb="tag"] [role="button"]{color:#ffffff!important;}
.main .block-container{padding:20px 28px!important;max-width:1400px!important;}

/* 사이드바 버튼 스타일 */
section[data-testid="stSidebar"] .stButton>button{
    background:#f4f6f9!important;border:1px solid #e2e6ea!important;border-radius:8px!important;
    padding:9px 14px!important;cursor:pointer!important;transition:all .15s!important;
    font-size:13px!important;font-weight:500!important;color:#374151!important;width:100%!important;
    margin:2px 0!important;text-align:left!important;
}
section[data-testid="stSidebar"] .stButton>button:hover{background:#eef2ff!important;border-color:#6366f1!important;color:#6366f1!important;}
section[data-testid="stSidebar"] .stButton>button[kind="primary"]{background:#eef2ff!important;border-color:#6366f1!important;color:#6366f1!important;font-weight:700!important;}

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
.info-card{background:#fff;border:1px solid #bfdbfe;border-left:4px solid #3b82f6;border-radius:10px;padding:12px 14px;margin:5px 0;}
.info-title{font-size:13px;font-weight:600;color:#2563eb;margin-bottom:3px;}
.info-body{font-size:12px;color:#6b7280;line-height:1.5;}

.cbox{background:#f8fafc;border:1px solid #e2e6ea;border-left:3px solid #ef4444;border-radius:6px;padding:8px 12px;margin:3px 0;font-size:12px;color:#374151;line-height:1.5;}
.section-title{font-size:17px;font-weight:700;color:#1e2330;margin-bottom:2px;}
.section-sub{font-size:12px;color:#94a3b8;margin-bottom:14px;}
.tag{display:inline-block;background:#eef2ff;border-radius:999px;padding:2px 9px;font-size:11px;color:#6366f1;margin:2px;}
.stTabs [data-baseweb="tab-list"]{background:transparent!important;border-bottom:1px solid #e2e6ea!important;}
.stTabs [data-baseweb="tab"]{background:transparent!important;color:#94a3b8!important;font-size:13px!important;font-weight:500!important;}
.stTabs [aria-selected="true"]{color:#1e2330!important;border-bottom:2px solid #6366f1!important;font-weight:600!important;}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:3px;}

.insight-box{background:linear-gradient(135deg,#f0f4ff,#fff);border:1px solid #c7d2fe;border-radius:12px;padding:16px 18px;margin:8px 0;}
.insight-title{font-size:14px;font-weight:700;color:#4338ca;margin-bottom:8px;}
.insight-item{font-size:12px;color:#374151;margin:4px 0;padding:4px 8px;background:rgba(255,255,255,.7);border-radius:6px;border-left:3px solid #6366f1;}
</style>
""", unsafe_allow_html=True)

# ── Plotly 테마 ──
def plt_theme(fig, title="", h=340):
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

# ── 그래프 값 불투명 박스 텍스트 공통 설정 ──
TEXT_FONT = dict(size=11, color="#1e2330")
TEXT_BGCOLOR = "rgba(255,255,255,0.85)"
TEXT_BORDERCOLOR = "#e2e6ea"

def add_text_box_style(fig):
    """모든 트레이스에 불투명 텍스트 박스 적용"""
    fig.update_traces(
        textfont=dict(size=11, color="#1e2330"),
        # texttemplate에 bgcolor 적용은 plotly bar/scatter에서 직접 지원 안 되므로
        # hoverlabel을 활용하고 텍스트는 outside + bgcolor 근사
    )
    return fig

CH_COLOR = {"CALL-IN":"#6366f1","CALL-OB":"#06b6d4","CHAT":"#22c55e","게시판":"#f59e0b"}

# ── 이슈 분류 패턴 (실제 감점사유 텍스트 기반, 정확도 최우선) ──
# 우선순위: 구체적·긴 패턴 먼저 → 짧은 패턴 나중
ISSUE_PATTERNS = [
    # ── 인사 관련 ──
    ("첫인사 누락",       r"첫인사\s*(누락|미이행|안\s*함|하지\s*않)|인사\s*누락(?!.*끝)"),
    ("끝인사 누락",       r"끝인사\s*(누락|미이행|안\s*함)|마무리\s*인사\s*(누락|미이행)"),
    ("통화종료 미흡",     r"통화\s*종료\s*(미흡|누락|먼저)|먼저\s*끊|종료\s*(미흡|오류)"),
    ("정보확인 누락",     r"정보\s*확인\s*(누락|미이행|안\s*함)|본인\s*확인\s*(누락|미이행)|인증\s*(누락|미이행)"),
    ("인사톤 미흡",       r"인사\s*톤|톤\s*(낮음|무감|단조|미흡)|발성\s*미흡"),

    # ── 언어·말투 ──
    ("습관어/말버릇",     r"습관어|말버릇|아무래도|저희\s*쪽|구요\s*~|그\s*니까|어\s*~|음\s*~|네\s*~\s*네|에\s*~"),
    ("존칭 오류",         r"존칭\s*(오류|누락|미사용)|요조체|비\s*정중|반말|하세요체\s*오류|합쇼체\s*오류|셨을까요|존댓말"),
    ("토막말/도치",       r"토막말|도치\s*(문장)?|문장\s*(불완전|끊김|짧음)|짧게\s*끊|끊어\s*말"),
    ("두괄식 미이행",     r"두괄식|결론\s*(먼저|후에|나중)|요점\s*먼저"),
    ("문법/띄어쓰기",     r"띄어\s*쓰기|문법\s*(오류|틀림)|맞춤법|오타|철자\s*(오류|틀림)"),

    # ── 고객응대 태도 ──
    ("양해 누락",         r"양해\s*(누락|미이행|안\s*함)|양해\s*멘트\s*(누락|미이행)|대기\s*양해|보류\s*양해|보류\s*전\s*양해"),
    ("대기 안내 누락",    r"대기\s*안내\s*(누락|미이행)|보류\s*안내\s*(누락|미이행)|홀드\s*안내\s*(누락|미이행)"),
    ("호응 부족",         r"호응\s*(누락|부족|미흡|없음)|즉각\s*호응\s*(누락|부족)|공감\s*(부족|미흡)|경청\s*(부족|미흡)"),
    ("감정연출 미흡",     r"감정\s*연출\s*(미흡|부족)|감성\s*(응대|멘트)\s*(미흡|부족)|공감\s*멘트\s*(누락|미흡)"),

    # ── 안내·업무 처리 ──
    ("오안내",            r"오\s*안내|잘못\s*(안내|설명)|틀린\s*안내|안내\s*오류|잘못된\s*정보"),
    ("안내 누락",         r"안내\s*(누락|미이행|빠짐)|미\s*안내(?!\s*대기)|설명\s*(누락|빠짐)"),
    ("문의파악 미흡",     r"문의\s*파악\s*(미흡|부족|안\s*됨)|재\s*질의|재\s*문의|문의\s*내용\s*(파악|확인)\s*(미흡|안\s*됨)|요청\s*파악\s*(미흡|부족)"),
    ("맞춤설명 부족",     r"맞춤\s*(설명|안내)\s*(부족|미흡|누락)|개인화\s*(설명|안내)|고객\s*상황\s*(미반영|미고려)"),
    ("정확한안내 미흡",   r"부정확|정확\s*(하지|하게)\s*(않|못)|정보\s*(불완전|부정확)|안내\s*(불충분|미흡|부족)"),
    ("프로세스 미준수",   r"프로세스\s*(미준수|위반|누락)|절차\s*(누락|미준수|위반)|순서\s*(오류|위반)|약속\s*(미준수|불이행)"),
    ("처리 오류/지연",    r"처리\s*(오류|지연|미흡|실수)|업무\s*(오류|실수|지연)|잘못\s*처리"),
    ("이력 미기재",       r"이력\s*(누락|미기재|미작성|안\s*씀)|상담\s*이력\s*(누락|미기재)|메모\s*(누락|미작성)|기재\s*(누락|안\s*됨)"),
    ("전산처리 오류",     r"전산\s*(오류|누락|미처리|실수)|시스템\s*(오류|누락)|입력\s*(오류|누락|실수)"),

    # ── 채팅/게시판 전용 ──
    ("가독성 부족",       r"가독성|문단\s*(나눔|없음|미흡)|엔터\s*(없음|미사용)|줄\s*바꿈\s*(없음|미흡)|단락\s*(없음|미구분)"),
    ("첨부 누락",         r"(사진|파일|첨부|이미지)\s*(누락|미첨부|빠짐)"),

    # ── 음성 전용 ──
    ("음성숙련도 미흡",   r"음성\s*숙련(도)?\s*(미흡|부족)|발음\s*(불명확|미흡|나쁨)|말\s*속도\s*(빠름|느림|미흡)|목소리\s*(미흡|작음|불명확)"),
]

ROOT_CAUSE = {
    "첫인사 누락":       "첫인사 체크리스트 미이행 → 콜 시작 루틴 점검",
    "끝인사 누락":       "끝인사 체크리스트 미이행 → 콜 마무리 루틴 점검",
    "통화종료 미흡":     "통화종료 절차 미숙 → 종료 멘트 패턴 훈련",
    "정보확인 누락":     "본인인증 절차 누락 → 프로세스 재교육",
    "인사톤 미흡":       "발화 톤 관리 부족 → 인사 톤 연습",
    "습관어/말버릇":     "무의식적 언어 습관 → 녹취 모니터링 후 의식적 교정",
    "존칭 오류":         "경어법 이해 부족 → 표준 멘트 암기 훈련",
    "토막말/도치":       "문장 구성 습관 → 완전한 문장 말하기 연습",
    "두괄식 미이행":     "설명 구조 습관 → 결론 먼저 말하는 훈련",
    "문법/띄어쓰기":     "문서 작성 능력 부족 → 글쓰기 교육",
    "양해 누락":         "고객 공감 부족 → 양해 멘트 패턴 훈련",
    "대기 안내 누락":    "대기 안내 절차 누락 → 보류 전 멘트 습관화",
    "호응 부족":         "경청 자세 부족 → 적극적 호응 연습",
    "감정연출 미흡":     "감성 응대 부족 → 감정연출 스크립트 교육",
    "오안내":            "업무 지식 오류 → 정확한 업무 매뉴얼 재숙지",
    "안내 누락":         "업무 숙지 미흡 → 주요 안내사항 재교육",
    "문의파악 미흡":     "청취 이해 부족 → 문의 요약 확인 습관 훈련",
    "맞춤설명 부족":     "고객 상황 파악 부족 → 개인화 응대 교육",
    "정확한안내 미흡":   "정보 정확도 부족 → 업무 지식 강화 교육",
    "프로세스 미준수":   "절차 이해 부족 → 프로세스 플로우 재교육",
    "처리 오류/지연":    "업무 정확도 부족 → 처리 단계별 확인 습관화",
    "이력 미기재":       "기록 습관 부재 → 상담 후 이력 루틴화",
    "전산처리 오류":     "시스템 활용 미숙 → 전산 처리 재교육",
    "가독성 부족":       "채팅 문서화 부족 → 구조화 글쓰기 교육",
    "첨부 누락":         "체크리스트 미활용 → 처리 단계 확인 습관화",
    "음성숙련도 미흡":   "발화 기술 부족 → 음성 훈련 및 모니터링",
    "기타":              "개별 피드백 참고 후 맞춤 교육",
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

# ── 채널별 독립 항목 정의 (만점 기준 재검수) ──
# CALL: 총 100점 기준
CALL_ITEMS = {
    "첫인사":    (8,  2.5),
    "정보확인":  (11, 5),
    "끝인사":    (14, 2.5),
    "인사톤":    (17, 2.5),
    "통화종료":  (20, 2.5),
    "음성숙련도":(23, 5),
    "감정연출":  (26, 5),
    "양해":      (29, 5),
    "즉각호응":  (32, 5),
    "대기":      (35, 5),
    "언어표현":  (38, 5),
    "경청":      (41, 5),
    "문의파악":  (44, 5),
    "맞춤설명":  (47, 5),
    "정확한안내":(50, 10),
    "프로세스":  (53, 10),
    "전산처리":  (56, 10),
    "상담이력":  (59, 10),
}
# CHAT: 총 100점 기준
CHAT_ITEMS = {
    "첫인사":    (6,  5),
    "정보확인":  (9,  5),
    "끝인사":    (12, 5),
    "양해":      (15, 5),
    "즉각호응":  (18, 5),
    "대기":      (21, 5),
    "언어표현":  (24, 10),
    "가독성":    (27, 5),
    "문의파악":  (30, 10),
    "맞춤설명":  (33, 10),
    "정확한안내":(36, 10),
    "프로세스":  (39, 5),
    "전산처리":  (42, 10),
    "상담이력":  (45, 10),
}
# 게시판: 총 100점 기준
BOARD_ITEMS = {
    "첫인사":    (7,  2.5),
    "끝인사":    (10, 2.5),
    "언어표현":  (13, 10),
    "양해":      (16, 10),
    "가독성":    (19, 10),
    "문의파악":  (22, 10),
    "맞춤설명":  (25, 10),
    "정확한안내":(28, 10),
    "프로세스":  (31, 10),
    "전산처리":  (34, 10),
    "상담이력":  (37, 10),   # 만점 10점 (수정)
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
                rec[f"{iname}_점수"]     = _num(row.iloc[ci]) if ci < len(row) else None
                rec[f"{iname}_최대"]     = mx
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
# 사이드바 — 메뉴 + 필터 통합
# ══════════════════════════════
with st.sidebar:
    st.markdown("<div style='font-size:15px;font-weight:700;color:#1e2330;padding-bottom:10px;margin-bottom:10px;border-bottom:1px solid #e2e6ea'>QA Monitoring</div>", unsafe_allow_html=True)

    # ── 기간/채널/상담사 필터 (사이드바 상단) ──
    st.markdown("<div style='font-size:11px;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px'>필터</div>", unsafe_allow_html=True)
    sel_period  = st.multiselect("평가기간",  all_periods,  default=all_periods,  label_visibility="collapsed", placeholder="평가기간 선택")
    sel_channel = st.multiselect("채널",      all_channels, default=all_channels, label_visibility="collapsed", placeholder="채널 선택")
    sel_agent   = st.multiselect("상담사",    all_agents,   default=all_agents,   label_visibility="collapsed", placeholder="상담사 선택")

    if st.button("🔄 새로고침", use_container_width=True):
        st.cache_data.clear(); st.rerun()

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    st.markdown("<div style='border-top:1px solid #e2e6ea;margin-bottom:10px'></div>", unsafe_allow_html=True)

    # ── 메뉴 버튼 ──
    st.markdown("<div style='font-size:11px;font-weight:600;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px'>메뉴</div>", unsafe_allow_html=True)

    MENU_ITEMS = [
        "개요 및 위기감지",
        "점수 트렌드",
        "문의유형 분석",
        "채널별 항목 분석",
        "감점 분석",
        "상담사 종합평가",
        "이행률 분석",
        "항목별 원인 추적",   # 신규
        "평가자 분석",
    ]

    if "current_page" not in st.session_state:
        st.session_state["current_page"] = "개요 및 위기감지"

    for m in MENU_ITEMS:
        is_active = st.session_state["current_page"] == m
        btn_type = "primary" if is_active else "secondary"
        if st.button(m, key=f"menu_{m}", use_container_width=True, type=btn_type):
            st.session_state["current_page"] = m
            st.rerun()

    st.markdown(f"<div style='font-size:11px;color:#94a3b8;margin-top:16px;padding-top:12px;border-top:1px solid #e2e6ea'>총 {len(df_all)}건 로드됨</div>", unsafe_allow_html=True)

page = st.session_state["current_page"]

def flt(df):
    o = df.copy()
    if sel_period:  o = o[o["평가기간"].isin(sel_period)]
    if sel_channel: o = o[o["채널"].isin(sel_channel)]
    if sel_agent:   o = o[o["상담사"].isin(sel_agent)]
    return o

df    = flt(df_all)
df_c  = flt(df_call)  if len(df_call)  else pd.DataFrame()
df_h  = flt(df_chat)  if len(df_chat)  else pd.DataFrame()
df_b  = flt(df_board) if len(df_board) else pd.DataFrame()

# ══════════════════════════════════════════════════════════════════
# 공통 유틸
# ══════════════════════════════════════════════════════════════════
def get_channel_item_scores(ch_df, items_map):
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
    recs = []
    for iname, (_, mx) in items_map.items():
        sc_col = f"{iname}_점수"
        rs_col = f"{iname}_감점사유"
        if sc_col not in src_df.columns: continue
        cols = [sc_col, rs_col, "상담사","평가기간","채널"]
        sub = src_df[cols].copy()
        sub.columns = ["점수","감점사유","상담사","평가기간","채널"]
        sub["항목"] = iname; sub["최대"] = mx
        sub = sub[sub["점수"].notna()]
        sub = sub[sub["점수"] < mx]
        sub = sub[sub["감점사유"] != ""]
        recs.append(sub)
    return pd.concat(recs, ignore_index=True) if recs else pd.DataFrame()

# ── 항목별 기간 이행률 추이 계산 ──
def get_item_trend_df(ch_df, items_map):
    rows = []
    if ch_df is None or len(ch_df)==0: return pd.DataFrame()
    for period in sorted(ch_df["평가기간"].unique(), key=period_key):
        pdf = ch_df[ch_df["평가기간"]==period]
        for iname,(ci,mx) in items_map.items():
            sc_col = f"{iname}_점수"
            if sc_col not in pdf.columns: continue
            sc = pdf[sc_col].dropna().tolist()
            cr = comp_rate(sc,[mx]*len(sc))
            if cr is not None:
                rows.append({"기간":period,"항목":iname,"이행률":cr,"기간_정렬":period_key(period)})
    return pd.DataFrame(rows).sort_values("기간_정렬") if rows else pd.DataFrame()

# ── 위기감지 ──
def crisis_analysis():
    results = []
    for agent in df_all["상담사"].unique():
        adf = df_all[df_all["상담사"] == agent]
        adf_f = flt(adf)
        if len(adf_f) == 0: continue
        periods_sorted = sorted(adf_f["평가기간"].unique(), key=period_key)
        period_avgs = [adf_f[adf_f["평가기간"]==p]["TOTAL"].mean() for p in periods_sorted]
        trend = "유지"; trend_val = 0.0
        if len(period_avgs) >= 2:
            trend_val = period_avgs[-1] - period_avgs[-2]
            if trend_val <= -3: trend = "악화"
            elif trend_val >= 3: trend = "개선"
        all_issues = []
        for ch_df2, items_map in [(df_c, CALL_ITEMS),(df_h, CHAT_ITEMS),(df_b, BOARD_ITEMS)]:
            if ch_df2 is None or len(ch_df2)==0: continue
            a_ch = ch_df2[ch_df2["상담사"]==agent]
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
        ch_scores = {}
        for ch in adf_f["채널"].unique():
            avg = adf_f[adf_f["채널"]==ch]["TOTAL"].mean()
            ch_scores[ch] = round(avg, 1)
        results.append({
            "상담사": agent, "전체평균": round(adf_f["TOTAL"].mean(), 1),
            "평가건수": len(adf_f), "트렌드": trend, "트렌드값": round(trend_val, 1),
            "반복이슈수": len(repeat), "최다이슈": top_issue[0],
            "최다이슈횟수": top_issue[1], "채널별점수": ch_scores, "이슈카운터": ic,
        })
    return pd.DataFrame(results)

# ══════════════════════════════════════════════════════════════════
# 개요 및 위기감지
# ══════════════════════════════════════════════════════════════════
if page == "개요 및 위기감지":
    sh("개요 및 위기감지", "전체 현황 요약 · 집중 코칭이 필요한 상담사 자동 감지")

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: kpi("총 평가건수", f"{len(df):,}건")
    with c2: kpi("전체 평균", f"{df['TOTAL'].dropna().mean():.1f}점" if len(df) else "-")
    with c3: kpi("상담사 수", f"{df['상담사'].nunique()}명")
    with c4: kpi("평가 기간", f"{df['평가기간'].nunique()}개")
    with c5: kpi("활성 채널", f"{df['채널'].nunique()}개")

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

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
                        <div class="alert-title">🚨 즉시 코칭 필요 — {r["상담사"]}</div>
                        <div class="alert-body">
                            평균 {r["전체평균"]}점 · 트렌드 {r["트렌드값"]:+.1f}점<br>
                            반복이슈 {r["반복이슈수"]}개 · 최다: {r["최다이슈"]} ({r["최다이슈횟수"]}회)<br>
                            채널별: {ch_txt}
                        </div></div>''', unsafe_allow_html=True)
            if len(warning):
                for _, r in warning.iterrows():
                    st.markdown(f'''<div class="warn-card">
                        <div class="warn-title">⚠️ 모니터링 필요 — {r["상담사"]}</div>
                        <div class="warn-body">반복이슈 {r["반복이슈수"]}개 · 최다: {r["최다이슈"]}</div>
                    </div>''', unsafe_allow_html=True)
            if len(good):
                for _, r in good.iterrows():
                    st.markdown(f'''<div class="good-card">
                        <div class="good-title">✅ 개선 중 — {r["상담사"]}</div>
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
                             barmode="group", color_discrete_map=CH_COLOR,
                             text="평균점수")
                plt_theme(fig, "상담사별 채널 점수", 300)
                fig.update_traces(
                    texttemplate="%{text:.1f}",
                    textposition="outside",
                    textfont=dict(size=10, color="#1e2330"),
                    cliponaxis=False
                )
                fig.add_hline(y=df["TOTAL"].mean(), line_dash="dash", line_color="#94a3b8",
                              annotation_text=f"전체평균", annotation_font_color="#94a3b8")
                st.plotly_chart(fig, use_container_width=True)

            # ── 채널별 항목 이행률 추이 (상담사별 채널 점수 비교 바로 아래) ──
            st.markdown("#### 채널별 항목 이행률 추이")
            ov_tab1, ov_tab2, ov_tab3 = st.tabs(["CALL", "CHAT", "게시판"])
            for ov_tab, ov_src, ov_map, ov_lbl in [
                (ov_tab1, df_c, CALL_ITEMS, "CALL"),
                (ov_tab2, df_h, CHAT_ITEMS, "CHAT"),
                (ov_tab3, df_b, BOARD_ITEMS, "게시판"),
            ]:
                with ov_tab:
                    if ov_src is None or len(ov_src)==0:
                        st.info(f"{ov_lbl} 데이터 없음"); continue
                    trend_df = get_item_trend_df(ov_src, ov_map)
                    if len(trend_df):
                        avg_trend = trend_df.groupby(["기간","기간_정렬"])["이행률"].mean().reset_index().sort_values("기간_정렬")
                        low_items = trend_df.groupby("항목")["이행률"].mean().nsmallest(5).index.tolist()
                        low_trend = trend_df[trend_df["항목"].isin(low_items)]
                        col1_ov, col2_ov = st.columns(2)
                        with col1_ov:
                            fig = px.line(avg_trend, x="기간", y="이행률", markers=True,
                                          labels={"이행률":"전체 이행률 (%)"},
                                          text="이행률")
                            plt_theme(fig, f"{ov_lbl} 전체 이행률 추이", 260)
                            fig.update_traces(line=dict(width=2.5,color="#6366f1"), marker=dict(size=7),
                                              texttemplate="%{text:.1f}%", textposition="top center",
                                              textfont=dict(size=10,color="#1e2330"))
                            fig.add_hline(y=90, line_dash="dot", line_color="#22c55e", annotation_text="목표 90%")
                            st.plotly_chart(fig, use_container_width=True)
                        with col2_ov:
                            if len(low_trend):
                                fig2 = px.line(low_trend, x="기간", y="이행률", color="항목",
                                               markers=True, labels={"이행률":"이행률 (%)"},
                                               text="이행률")
                                plt_theme(fig2, f"{ov_lbl} 하위 항목 이행률 추이 (하위 5개)", 260)
                                fig2.update_traces(line=dict(width=2), marker=dict(size=6),
                                                   texttemplate="%{text:.0f}%", textposition="top center",
                                                   textfont=dict(size=9,color="#1e2330"))
                                fig2.add_hline(y=90, line_dash="dot", line_color="#94a3b8")
                                st.plotly_chart(fig2, use_container_width=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── 항목별 감점 기여자 히트맵 ──
    st.markdown("#### 항목별 차감 기여 상담사 (채널 분리)")
    tab_ci, tab_ch2, tab_cb = st.tabs(["CALL (IN+OB)", "CHAT", "게시판"])

    for tab, ch_src, items_map, label in [
        (tab_ci, df_c, CALL_ITEMS, "CALL"),
        (tab_ch2, df_h, CHAT_ITEMS, "CHAT"),
        (tab_cb, df_b, BOARD_ITEMS, "게시판"),
    ]:
        with tab:
            if ch_src is None or len(ch_src)==0:
                st.info(f"{label} 데이터 없음"); continue
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
                plt_theme(fig, f"{label} 항목별 감점 횟수", max(260, len(pvt)*38))
                fig.update_layout(xaxis=dict(tickangle=-30))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("감점 데이터 없음")

    # ── 기간별 점수 추이 ──
    st.markdown("#### 기간별 평균 점수 추이")
    grp = df.groupby(["평가기간","채널","기간_정렬"])["TOTAL"].mean().reset_index().sort_values("기간_정렬")
    fig = px.line(grp, x="평가기간", y="TOTAL", color="채널",
                  markers=True, color_discrete_map=CH_COLOR,
                  labels={"TOTAL":"평균 점수"}, text="TOTAL")
    plt_theme(fig, "", 280)
    fig.update_traces(line=dict(width=2.5), marker=dict(size=6),
                      texttemplate="%{text:.1f}", textposition="top center",
                      textfont=dict(size=10,color="#1e2330"))
    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 점수 트렌드
# ══════════════════════════════════════════════════════════════════
elif page == "점수 트렌드":
    sh("점수 트렌드", "기간별·상담사별·채널별 점수 변화")
    tab1,tab2,tab3 = st.tabs(["기간별 추이","상담사별 비교","Box Plot"])
    with tab1:
        grp = df.groupby(["평가기간","상담사","기간_정렬"])["TOTAL"].mean().reset_index().sort_values("기간_정렬")
        fig = px.line(grp, x="평가기간", y="TOTAL", color="상담사", markers=True,
                      labels={"TOTAL":"평균 점수"}, text="TOTAL")
        plt_theme(fig, "상담사별 점수 추이", 380)
        fig.update_traces(line=dict(width=2), marker=dict(size=5),
                          texttemplate="%{text:.1f}", textposition="top center",
                          textfont=dict(size=9,color="#1e2330"))
        st.plotly_chart(fig, use_container_width=True)
        grp2 = df.groupby(["평가기간","채널","기간_정렬"])["TOTAL"].mean().reset_index().sort_values("기간_정렬")
        fig2 = px.line(grp2, x="평가기간", y="TOTAL", color="채널", markers=True,
                       color_discrete_map=CH_COLOR, labels={"TOTAL":"평균 점수"}, text="TOTAL")
        plt_theme(fig2, "채널별 점수 추이", 280)
        fig2.update_traces(line=dict(width=2.5), marker=dict(size=6),
                           texttemplate="%{text:.1f}", textposition="top center",
                           textfont=dict(size=10,color="#1e2330"))
        st.plotly_chart(fig2, use_container_width=True)
    with tab2:
        aa = df.groupby(["상담사","채널"])["TOTAL"].mean().reset_index()
        fig = px.bar(aa, x="상담사", y="TOTAL", color="채널", barmode="group",
                     color_discrete_map=CH_COLOR, labels={"TOTAL":"평균 점수"}, text="TOTAL")
        plt_theme(fig, "상담사 × 채널별 평균 점수", 360)
        fig.update_traces(texttemplate="%{text:.1f}", textposition="outside",
                          textfont=dict(size=10,color="#1e2330"), cliponaxis=False)
        fig.add_hline(y=df["TOTAL"].mean(), line_dash="dash", line_color="#94a3b8")
        st.plotly_chart(fig, use_container_width=True)
    with tab3:
        fig = px.box(df, x="상담사", y="TOTAL", color="채널",
                     color_discrete_map=CH_COLOR, labels={"TOTAL":"점수"})
        plt_theme(fig, "상담사별 점수 분포", 380)
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
        plt_theme(fig, "문의유형(대) × 상담사 평균점수", max(300, len(pivot)*35))
        st.plotly_chart(fig, use_container_width=True)
    with tab2:
        co1,co2 = st.columns(2)
        with co1:
            cnt = df["문의유형대"].value_counts().reset_index(); cnt.columns=["유형","건수"]
            fig = px.bar(cnt, x="유형", y="건수", color="건수",
                         color_continuous_scale=["#e0e7ff","#6366f1"], text="건수")
            plt_theme(fig, "문의유형별 건수", 300)
            fig.update_traces(textposition="outside", textfont=dict(size=10,color="#1e2330"), cliponaxis=False)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        with co2:
            avg = df.groupby("문의유형대")["TOTAL"].mean().reset_index(); avg.columns=["유형","평균점수"]
            fig = px.bar(avg, x="유형", y="평균점수", color="평균점수",
                         color_continuous_scale=["#ef4444","#22c55e"], text="평균점수")
            plt_theme(fig, "문의유형별 평균 점수", 300)
            fig.update_traces(texttemplate="%{text:.1f}", textposition="outside",
                              textfont=dict(size=10,color="#1e2330"), cliponaxis=False)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 채널별 항목 분석
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
                st.info(f"{ch_label} 데이터 없음"); continue

            st.markdown(f"**{ch_label} 평가 항목별 이행률 (만점 기준 표시)**")
            idf = get_channel_item_scores(ch_df_base, items_map)
            if len(idf):
                idf_sorted = idf.sort_values("이행률")
                # 항목명에 만점 기준 포함
                idf_sorted["항목_만점"] = idf_sorted.apply(
                    lambda r: f"{r['항목']} ({int(r['최대점수']) if r['최대점수']==int(r['최대점수']) else r['최대점수']}점)", axis=1)
                fig = go.Figure(go.Bar(
                    x=idf_sorted["이행률"], y=idf_sorted["항목_만점"], orientation="h",
                    marker_color=idf_sorted["이행률"].apply(
                        lambda x: "#22c55e" if x>=90 else "#f59e0b" if x>=75 else "#ef4444"),
                    text=idf_sorted["이행률"].apply(lambda x: f"{x:.1f}%"),
                    textposition="outside", textfont=dict(size=11, color="#1e2330"),
                ))
                plt_theme(fig, f"{ch_label} 항목별 이행률 (항목명 옆 만점 표시)", max(260, len(idf)*28))
                fig.update_layout(xaxis=dict(range=[0,112]))
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
                plt_theme(fig, f"{ch_label} 상담사 × 항목 이행률", max(260, len(hdf)*38))
                fig.update_layout(xaxis=dict(tickangle=-30))
                st.plotly_chart(fig, use_container_width=True)

            st.markdown(f"**{ch_label} 상담사별 항목 이행률 추이**")
            agent_sel2 = st.selectbox(f"상담사 선택 ({ch_label})", sorted(ch_df_base["상담사"].unique()), key=f"agent_{ch_label}")
            item_sel2 = st.multiselect(f"항목 선택 ({ch_label})", list(items_map.keys()),
                                        default=list(items_map.keys())[:4], key=f"item_{ch_label}")
            if item_sel2:
                adf4 = ch_df_base[ch_df_base["상담사"]==agent_sel2].copy()
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
                    fig = px.line(tdf, x="기간", y="이행률", color="항목", markers=True,
                                  text="이행률")
                    plt_theme(fig, f"{agent_sel2} — {ch_label} 항목별 이행률 추이", 300)
                    fig.update_traces(line=dict(width=2), marker=dict(size=6),
                                      texttemplate="%{text:.0f}%", textposition="top center",
                                      textfont=dict(size=9,color="#1e2330"))
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
                             color_continuous_scale=["#fee2e2","#ef4444"], text="건수")
                plt_theme(fig, f"{lbl} 항목별 감점 건수", max(260, len(ic2)*30))
                fig.update_traces(textposition="outside", textfont=dict(size=10,color="#1e2330"), cliponaxis=False)
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
                plt_theme(fig, f"{lbl} 상담사 × 항목 감점", max(260, len(ad)*38))
                st.plotly_chart(fig, use_container_width=True)

            # 감점 이슈 분류 분포
            all_iss = []
            for _, row in ded.iterrows():
                all_iss.extend(classify_issue(str(row["감점사유"])))
            if all_iss:
                ic_cnt = Counter(all_iss)
                iss_df = pd.DataFrame(ic_cnt.most_common(15), columns=["이슈유형","건수"])
                fig_iss = px.bar(iss_df, x="건수", y="이슈유형", orientation="h",
                                 color="건수", color_continuous_scale=["#fef3c7","#f59e0b"],
                                 text="건수")
                plt_theme(fig_iss, f"{lbl} 이슈 유형 분류 분포 (세분화)", max(300, len(iss_df)*28))
                fig_iss.update_traces(textposition="outside", textfont=dict(size=10,color="#1e2330"), cliponaxis=False)
                fig_iss.update_coloraxes(showscale=False)
                st.plotly_chart(fig_iss, use_container_width=True)

            st.markdown("**감점 원본 코멘트 (차수 포함)**")
            show = ded[["평가기간","상담사","항목","점수","최대","감점사유"]].copy()
            show = show.replace("", np.nan).dropna(subset=["감점사유"])
            show = show.rename(columns={"평가기간":"차수"})
            st.dataframe(show.reset_index(drop=True), use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 상담사 종합평가
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
        else: kpi("최근 트렌드", "-")

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

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
                rows_ch.append({"항목":iname,"이행률":cr or 0,"평균점수":np.mean(sc),"최대":mx,
                                 "항목_만점":f"{iname} ({int(mx) if mx==int(mx) else mx}점)"})
            if rows_ch:
                rdf = pd.DataFrame(rows_ch).sort_values("이행률")
                fig = go.Figure(go.Bar(
                    x=rdf["이행률"], y=rdf["항목_만점"], orientation="h",
                    marker_color=rdf["이행률"].apply(lambda x:"#22c55e" if x>=90 else "#f59e0b" if x>=75 else "#ef4444"),
                    text=rdf["이행률"].apply(lambda x:f"{x:.1f}%"),
                    textposition="outside", textfont=dict(size=11,color="#1e2330"),
                ))
                plt_theme(fig, f"{agent_sel} — {ch_name} 항목 이행률 (만점 포함)", max(240, len(rdf)*28))
                fig.update_layout(xaxis=dict(range=[0,112]))
                st.plotly_chart(fig, use_container_width=True)

    trend2 = adf_all.groupby(["평가기간","기간_정렬"])["TOTAL"].mean().reset_index().sort_values("기간_정렬")
    fig2 = px.line(trend2, x="평가기간", y="TOTAL", markers=True,
                   labels={"TOTAL":"평균 점수"}, text="TOTAL")
    plt_theme(fig2, f"{agent_sel} 기간별 점수 추이", 260)
    fig2.update_traces(line=dict(width=2.5,color="#6366f1"), marker=dict(size=7,color="#6366f1"),
                       texttemplate="%{text:.1f}", textposition="top center",
                       textfont=dict(size=10,color="#1e2330"))
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### 반복 이슈 및 코칭 포인트")
    all_issues2 = []
    issue_with_text = {}
    issue_with_period = {}
    for ch_src3, items_m2 in [(df_c,CALL_ITEMS),(df_h,CHAT_ITEMS),(df_b,BOARD_ITEMS)]:
        if ch_src3 is None or len(ch_src3)==0: continue
        a3 = ch_src3[ch_src3["상담사"]==agent_sel]
        for iname in items_m2:
            dc = f"{iname}_감점사유"
            if dc not in a3.columns: continue
            for _, row in a3.iterrows():
                r = row[dc]
                if not r or r in ("None","nan",""): continue
                cats = classify_issue(r)
                all_issues2.extend(cats)
                period_val = row.get("평가기간","")
                for cat in cats:
                    issue_with_text.setdefault(cat, []).append(r)
                    issue_with_period.setdefault(cat, []).append(period_val)

    ic3 = Counter(all_issues2)
    repeat3 = {k:v for k,v in ic3.items() if v>=2}
    if repeat3:
        for issue, cnt in sorted(repeat3.items(), key=lambda x:-x[1]):
            texts_ex = list(dict.fromkeys(issue_with_text.get(issue,[])))[:3]
            periods_ex = issue_with_period.get(issue,[])[:3]
            root = ROOT_CAUSE.get(issue,"개별 피드백 참고 후 맞춤 교육")
            boxes = "".join([f'<div class="cbox">[{p}] {t}</div>' for t,p in zip(texts_ex,periods_ex)])
            st.markdown(f'''<div class="alert-card">
                <div class="alert-title">{issue} — {cnt}회 반복</div>
                <div class="alert-body">📌 {root}<br>{boxes}</div>
            </div>''', unsafe_allow_html=True)
    else:
        st.success("반복 이슈 없음")

# ══════════════════════════════════════════════════════════════════
# 이행률 분석 (고도화)
# ══════════════════════════════════════════════════════════════════
elif page == "이행률 분석":
    sh("이행률 분석", "채널별 독립 이행률 — 항목별·상담사별·기간별 + 원인 인사이트")

    ch_tab1, ch_tab2, ch_tab3 = st.tabs(["CALL","CHAT","게시판"])
    for tab, ch_src4, items_m3, lbl2 in [
        (ch_tab1, df_c, CALL_ITEMS, "CALL"),
        (ch_tab2, df_h, CHAT_ITEMS, "CHAT"),
        (ch_tab3, df_b, BOARD_ITEMS, "게시판"),
    ]:
        with tab:
            if ch_src4 is None or len(ch_src4)==0:
                st.info(f"{lbl2} 데이터 없음"); continue

            t1,t2,t3,t4 = st.tabs([f"항목별",f"상담사별",f"기간별",f"📊 인사이트"])

            with t1:
                idf2 = get_channel_item_scores(ch_src4, items_m3)
                if len(idf2):
                    idf2["항목_만점"] = idf2.apply(
                        lambda r: f"{r['항목']} ({int(r['최대점수']) if r['최대점수']==int(r['최대점수']) else r['최대점수']}점)", axis=1)
                    s2 = idf2.sort_values("이행률")
                    fig = go.Figure(go.Bar(
                        x=s2["이행률"], y=s2["항목_만점"], orientation="h",
                        marker_color=s2["이행률"].apply(lambda x:"#22c55e" if x>=90 else "#f59e0b" if x>=75 else "#ef4444"),
                        text=s2["이행률"].apply(lambda x:f"{x:.1f}%"), textposition="outside",
                        textfont=dict(size=11,color="#1e2330"),
                    ))
                    plt_theme(fig, f"{lbl2} 항목별 이행률", max(260, len(s2)*28))
                    fig.update_layout(xaxis=dict(range=[0,112]))
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
                                 color_continuous_scale=["#ef4444","#22c55e"], text="이행률")
                    plt_theme(fig, f"{lbl2} 상담사별 이행률", max(260, len(acd)*30))
                    fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside",
                                      textfont=dict(size=10,color="#1e2330"), cliponaxis=False)
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
                    fig = px.line(pcd, x="기간", y="이행률", markers=True, text="이행률")
                    plt_theme(fig, f"{lbl2} 기간별 이행률 추이", 280)
                    fig.update_traces(line=dict(width=2.5,color="#6366f1"), marker=dict(size=7),
                                      texttemplate="%{text:.1f}%", textposition="top center",
                                      textfont=dict(size=10,color="#1e2330"))
                    fig.add_hline(y=90, line_dash="dot", line_color="#22c55e", annotation_text="목표 90%")
                    st.plotly_chart(fig, use_container_width=True)

            with t4:
                st.markdown("#### 📊 이행률 종합 인사이트")

                # 1) 전체 이행률 현황
                idf_ins = get_channel_item_scores(ch_src4, items_m3)
                if len(idf_ins)==0:
                    st.info("데이터 없음"); continue

                overall_rate = idf_ins["이행률"].mean()
                low_items_ins = idf_ins[idf_ins["이행률"]<75].sort_values("이행률")
                high_items_ins = idf_ins[idf_ins["이행률"]>=90].sort_values("이행률", ascending=False)

                col_ins1, col_ins2, col_ins3 = st.columns(3)
                with col_ins1: kpi("전체 평균 이행률", f"{overall_rate:.1f}%")
                with col_ins2: kpi("위험 항목 수 (<75%)", f"{len(low_items_ins)}개")
                with col_ins3: kpi("달성 항목 수 (≥90%)", f"{len(high_items_ins)}개")

                st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

                # 2) 항목별 악화/개선 추이 분석 (기간 2개 이상일 때)
                trend_ins = get_item_trend_df(ch_src4, items_m3)
                if len(trend_ins) > 0:
                    periods_ins = sorted(trend_ins["기간"].unique(), key=period_key)
                    if len(periods_ins) >= 2:
                        p_now = periods_ins[-1]
                        p_prev = periods_ins[-2]
                        now_df = trend_ins[trend_ins["기간"]==p_now].set_index("항목")["이행률"]
                        prev_df = trend_ins[trend_ins["기간"]==p_prev].set_index("항목")["이행률"]
                        common_items = now_df.index.intersection(prev_df.index)
                        diffs = (now_df[common_items] - prev_df[common_items]).sort_values()

                        improving = diffs[diffs > 5]
                        worsening = diffs[diffs < -5]

                        col_i1, col_i2 = st.columns(2)
                        with col_i1:
                            st.markdown(f"**📈 개선 중인 항목** ({p_prev} → {p_now})")
                            if len(improving):
                                for item, diff in improving.items():
                                    # 이 항목을 개선시킨 상담사 찾기
                                    contrib_agents = []
                                    sc_col = f"{item}_점수"
                                    mx = items_m3[item][1]
                                    for ag in ch_src4["상담사"].unique():
                                        ag_now = ch_src4[(ch_src4["상담사"]==ag)&(ch_src4["평가기간"]==p_now)]
                                        ag_prev = ch_src4[(ch_src4["상담사"]==ag)&(ch_src4["평가기간"]==p_prev)]
                                        if sc_col not in ch_src4.columns: continue
                                        sc_n = ag_now[sc_col].dropna().tolist()
                                        sc_p = ag_prev[sc_col].dropna().tolist()
                                        if sc_n and sc_p:
                                            cr_n = comp_rate(sc_n,[mx]*len(sc_n))
                                            cr_p = comp_rate(sc_p,[mx]*len(sc_p))
                                            if cr_n and cr_p and cr_n > cr_p:
                                                contrib_agents.append(f"{ag}(+{cr_n-cr_p:.0f}%)")
                                    agent_str = ", ".join(contrib_agents[:3]) if contrib_agents else "데이터 부족"
                                    st.markdown(f'''<div class="good-card">
                                        <div class="good-title">✅ {item} +{diff:.1f}%p</div>
                                        <div class="good-body">개선 기여: {agent_str}</div>
                                    </div>''', unsafe_allow_html=True)
                            else:
                                st.info("5%p 이상 개선 항목 없음")

                        with col_i2:
                            st.markdown(f"**📉 악화 중인 항목** ({p_prev} → {p_now})")
                            if len(worsening):
                                for item, diff in worsening.items():
                                    # 이 항목을 악화시킨 상담사 찾기
                                    problem_agents = []
                                    sc_col = f"{item}_점수"
                                    rs_col = f"{item}_감점사유"
                                    mx = items_m3[item][1]
                                    agent_comments = {}
                                    for ag in ch_src4["상담사"].unique():
                                        ag_now = ch_src4[(ch_src4["상담사"]==ag)&(ch_src4["평가기간"]==p_now)]
                                        ag_prev = ch_src4[(ch_src4["상담사"]==ag)&(ch_src4["평가기간"]==p_prev)]
                                        if sc_col not in ch_src4.columns: continue
                                        sc_n = ag_now[sc_col].dropna().tolist()
                                        sc_p = ag_prev[sc_col].dropna().tolist()
                                        if sc_n and sc_p:
                                            cr_n = comp_rate(sc_n,[mx]*len(sc_n))
                                            cr_p = comp_rate(sc_p,[mx]*len(sc_p))
                                            if cr_n and cr_p and cr_n < cr_p:
                                                # 감점 코멘트 수집
                                                cmts = []
                                                if rs_col in ag_now.columns:
                                                    cmts = [c for c in ag_now[rs_col].dropna() if c not in ("","None","nan")]
                                                problem_agents.append((ag, cr_n-cr_p, cmts))

                                    # 문제 상담사 출력
                                    prob_str_parts = []
                                    all_cmts = []
                                    for ag, delta, cmts in sorted(problem_agents, key=lambda x:x[1]):
                                        prob_str_parts.append(f"{ag}({delta:.0f}%)")
                                        all_cmts.extend(cmts[:1])

                                    prob_str = ", ".join(prob_str_parts[:3]) if prob_str_parts else "전반적 하락"
                                    cmt_boxes = "".join([f'<div class="cbox">{c}</div>' for c in all_cmts[:2]])
                                    st.markdown(f'''<div class="alert-card">
                                        <div class="alert-title">🚨 {item} {diff:.1f}%p 하락</div>
                                        <div class="alert-body">원인 상담사: {prob_str}<br>{cmt_boxes}</div>
                                    </div>''', unsafe_allow_html=True)
                            else:
                                st.info("5%p 이상 악화 항목 없음")

                # 3) 상담사별 이행률 문제 진단
                st.markdown("---")
                st.markdown("#### 👤 상담사별 이행률 진단 및 코칭 필요도")
                agent_diag = []
                for ag in ch_src4["상담사"].unique():
                    adf_d = ch_src4[ch_src4["상담사"]==ag]
                    item_rates = {}
                    for iname,(ci,mx) in items_m3.items():
                        sc_col = f"{iname}_점수"
                        if sc_col not in adf_d.columns: continue
                        sc = adf_d[sc_col].dropna().tolist()
                        cr = comp_rate(sc,[mx]*len(sc))
                        if cr is not None: item_rates[iname] = cr

                    if not item_rates: continue

                    avg_rate = np.mean(list(item_rates.values()))
                    low_cnt = sum(1 for v in item_rates.values() if v < 75)
                    low_items_ag = sorted([(k,v) for k,v in item_rates.items() if v<75], key=lambda x:x[1])

                    # 반복 이슈 수집
                    all_iss_ag = []
                    for iname in items_m3:
                        dc = f"{iname}_감점사유"
                        if dc not in adf_d.columns: continue
                        for r in adf_d[dc].dropna():
                            if r and r not in ("None","nan",""):
                                all_iss_ag.extend(classify_issue(r))
                    ic_ag = Counter(all_iss_ag)
                    repeat_ag = {k:v for k,v in ic_ag.items() if v>=2}

                    # 기간별 트렌드
                    trend_ag = "데이터부족"
                    periods_ag = sorted(adf_d["평가기간"].unique(), key=period_key)
                    if len(periods_ag)>=2:
                        rates_ag = []
                        for p in periods_ag:
                            p_df = adf_d[adf_d["평가기간"]==p]
                            vv = [comp_rate(p_df[f"{i}_점수"].dropna().tolist(),[mx]*len(p_df[f"{i}_점수"].dropna()))
                                  for i,(ci,mx) in items_m3.items() if f"{i}_점수" in p_df.columns]
                            vv = [v for v in vv if v is not None]
                            if vv: rates_ag.append(np.mean(vv))
                        if len(rates_ag)>=2:
                            delta_ag = rates_ag[-1] - rates_ag[-2]
                            trend_ag = f"{'↑' if delta_ag>0 else '↓'}{abs(delta_ag):.1f}%p"

                    agent_diag.append({
                        "상담사": ag, "평균이행률": round(avg_rate,1),
                        "위험항목수": low_cnt, "반복이슈수": len(repeat_ag),
                        "트렌드": trend_ag,
                        "저이행항목": low_items_ag, "반복이슈": ic_ag,
                    })

                agent_diag_df = pd.DataFrame(agent_diag).sort_values("평균이행률")

                if len(agent_diag_df):
                    # 요약 테이블
                    disp = agent_diag_df[["상담사","평균이행률","위험항목수","반복이슈수","트렌드"]].copy()
                    st.dataframe(disp.reset_index(drop=True), use_container_width=True)

                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

                    # 상세 진단 카드
                    st.markdown("**상담사별 상세 진단**")
                    sel_ag_ins = st.selectbox(f"상담사 선택 ({lbl2})", agent_diag_df["상담사"].tolist(), key=f"ins_{lbl2}")
                    ag_row = agent_diag_df[agent_diag_df["상담사"]==sel_ag_ins].iloc[0]

                    col_d1, col_d2 = st.columns(2)
                    with col_d1:
                        st.markdown(f'''<div class="insight-box">
                            <div class="insight-title">📋 {sel_ag_ins} 진단 요약</div>
                            <div class="insight-item">평균 이행률: {ag_row["평균이행률"]}%</div>
                            <div class="insight-item">위험 항목: {ag_row["위험항목수"]}개</div>
                            <div class="insight-item">반복 이슈: {ag_row["반복이슈수"]}개</div>
                            <div class="insight-item">기간 트렌드: {ag_row["트렌드"]}</div>
                        </div>''', unsafe_allow_html=True)

                        # 저이행 항목 상세
                        if ag_row["저이행항목"]:
                            st.markdown("**⚠️ 위험 항목 (이행률 <75%)**")
                            for iname, rate in ag_row["저이행항목"]:
                                mx = items_m3[iname][1]
                                # 해당 항목 감점 코멘트
                                sc_col = f"{iname}_점수"
                                rs_col = f"{iname}_감점사유"
                                adf_d2 = ch_src4[ch_src4["상담사"]==sel_ag_ins]
                                cmts2 = []
                                periods_cmts = []
                                if rs_col in adf_d2.columns:
                                    for _, r2 in adf_d2.iterrows():
                                        c2 = r2[rs_col]
                                        if c2 and c2 not in ("","None","nan"):
                                            cmts2.append(c2)
                                            periods_cmts.append(r2.get("평가기간",""))
                                cmt_html = "".join([f'<div class="cbox">[{p}] {c}</div>' for c,p in zip(cmts2[:3],periods_cmts[:3])])
                                root2 = ROOT_CAUSE.get(iname, "개별 피드백 참고")
                                st.markdown(f'''<div class="warn-card">
                                    <div class="warn-title">{iname} ({int(mx) if mx==int(mx) else mx}점) — 이행률 {rate:.1f}%</div>
                                    <div class="warn-body">📌 {root2}<br>{cmt_html}</div>
                                </div>''', unsafe_allow_html=True)

                    with col_d2:
                        # 반복 이슈 상세
                        ic_sel = ag_row["반복이슈"]
                        repeat_sel = {k:v for k,v in ic_sel.items() if v>=2}
                        if repeat_sel:
                            st.markdown("**🔁 반복 이슈 패턴**")
                            iss_df2 = pd.DataFrame(list(ic_sel.most_common(10)), columns=["이슈유형","횟수"])
                            fig_iss2 = px.bar(iss_df2, x="횟수", y="이슈유형", orientation="h",
                                             color="횟수", color_continuous_scale=["#fef3c7","#ef4444"],
                                             text="횟수")
                            plt_theme(fig_iss2, f"{sel_ag_ins} 이슈 분포", max(200, len(iss_df2)*26))
                            fig_iss2.update_traces(textposition="outside", textfont=dict(size=10,color="#1e2330"), cliponaxis=False)
                            fig_iss2.update_coloraxes(showscale=False)
                            st.plotly_chart(fig_iss2, use_container_width=True)

                            st.markdown("**🎯 코칭 우선순위**")
                            for issue2, cnt2 in sorted(repeat_sel.items(), key=lambda x:-x[1])[:5]:
                                root3 = ROOT_CAUSE.get(issue2, "개별 피드백 참고 후 맞춤 교육")
                                severity = "즉시 코칭" if cnt2 >= 3 else "모니터링"
                                card_class = "alert-card" if cnt2>=3 else "warn-card"
                                title_class = "alert-title" if cnt2>=3 else "warn-title"
                                st.markdown(f'''<div class="{card_class}">
                                    <div class="{title_class}">[{severity}] {issue2} — {cnt2}회</div>
                                    <div class="alert-body">📌 {root3}</div>
                                </div>''', unsafe_allow_html=True)
                        else:
                            st.success("반복 이슈 없음")

                        # 기간별 항목 이행률 추이 (선택 상담사)
                        ag_trend_rows = []
                        adf_ag_trend = ch_src4[ch_src4["상담사"]==sel_ag_ins]
                        for period_ag2 in sorted(adf_ag_trend["평가기간"].unique(), key=period_key):
                            pf = adf_ag_trend[adf_ag_trend["평가기간"]==period_ag2]
                            for iname2,(ci2,mx2) in items_m3.items():
                                sc2 = pf[f"{iname2}_점수"].dropna().tolist() if f"{iname2}_점수" in pf.columns else []
                                cr2 = comp_rate(sc2,[mx2]*len(sc2))
                                if cr2 is not None:
                                    ag_trend_rows.append({"기간":period_ag2,"항목":iname2,"이행률":cr2})
                        if ag_trend_rows:
                            ag_tdf = pd.DataFrame(ag_trend_rows)
                            # 하위 항목만 표시
                            low_ag_items = [k for k,v in ag_row["저이행항목"]][:5] if ag_row["저이행항목"] else list(items_m3.keys())[:5]
                            ag_tdf_low = ag_tdf[ag_tdf["항목"].isin(low_ag_items)]
                            if len(ag_tdf_low):
                                fig_ag_t = px.line(ag_tdf_low, x="기간", y="이행률", color="항목",
                                                   markers=True, text="이행률")
                                plt_theme(fig_ag_t, f"{sel_ag_ins} 위험 항목 추이", 240)
                                fig_ag_t.update_traces(line=dict(width=2), marker=dict(size=6),
                                                       texttemplate="%{text:.0f}%", textposition="top center",
                                                       textfont=dict(size=9,color="#1e2330"))
                                fig_ag_t.add_hline(y=75, line_dash="dot", line_color="#ef4444", annotation_text="위험선 75%")
                                fig_ag_t.add_hline(y=90, line_dash="dot", line_color="#22c55e", annotation_text="목표 90%")
                                st.plotly_chart(fig_ag_t, use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 항목별 원인 추적 (신규 메뉴)
# ══════════════════════════════════════════════════════════════════
elif page == "항목별 원인 추적":
    sh("항목별 원인 추적", "항목 악화/개선의 원인 상담사와 반복 패턴 심층 분석")

    ch_sel_ot = st.selectbox("채널 선택", ["CALL","CHAT","게시판"], key="ot_ch")
    ch_src_ot = {"CALL":df_c, "CHAT":df_h, "게시판":df_b}[ch_sel_ot]
    items_ot  = {"CALL":CALL_ITEMS, "CHAT":CHAT_ITEMS, "게시판":BOARD_ITEMS}[ch_sel_ot]

    if ch_src_ot is None or len(ch_src_ot)==0:
        st.info("데이터 없음")
    else:
        item_sel_ot = st.selectbox("항목 선택", list(items_ot.keys()), key="ot_item")
        mx_ot = items_ot[item_sel_ot][1]
        sc_col_ot = f"{item_sel_ot}_점수"
        rs_col_ot = f"{item_sel_ot}_감점사유"

        st.markdown(f"**{ch_sel_ot} > {item_sel_ot} (만점 {int(mx_ot) if mx_ot==int(mx_ot) else mx_ot}점) 심층 분석**")

        # 전체 이행률
        sc_all_ot = ch_src_ot[sc_col_ot].dropna().tolist() if sc_col_ot in ch_src_ot.columns else []
        if sc_all_ot:
            cr_all = comp_rate(sc_all_ot,[mx_ot]*len(sc_all_ot))
            col_ot1,col_ot2,col_ot3 = st.columns(3)
            with col_ot1: kpi("전체 이행률", f"{cr_all:.1f}%")
            with col_ot2: kpi("감점 건수", f"{sum(1 for s in sc_all_ot if s<mx_ot)}건")
            with col_ot3: kpi("평가 건수", f"{len(sc_all_ot)}건")

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        col_ot_a, col_ot_b = st.columns(2)
        with col_ot_a:
            # 상담사별 이행률 비교
            ag_rates_ot = []
            for ag in ch_src_ot["상담사"].unique():
                adf_ot = ch_src_ot[ch_src_ot["상담사"]==ag]
                sc_ag = adf_ot[sc_col_ot].dropna().tolist() if sc_col_ot in adf_ot.columns else []
                if sc_ag:
                    cr_ag = comp_rate(sc_ag,[mx_ot]*len(sc_ag))
                    deduct_cnt = sum(1 for s in sc_ag if s<mx_ot)
                    ag_rates_ot.append({"상담사":ag,"이행률":cr_ag or 0,"감점건수":deduct_cnt})
            if ag_rates_ot:
                ag_ot_df = pd.DataFrame(ag_rates_ot).sort_values("이행률")
                fig_ot = px.bar(ag_ot_df, x="이행률", y="상담사", orientation="h",
                                color="이행률", color_continuous_scale=["#ef4444","#22c55e"],
                                text="이행률")
                plt_theme(fig_ot, f"상담사별 [{item_sel_ot}] 이행률", max(240, len(ag_ot_df)*30))
                fig_ot.update_traces(texttemplate="%{text:.1f}%", textposition="outside",
                                     textfont=dict(size=10,color="#1e2330"), cliponaxis=False)
                fig_ot.update_coloraxes(showscale=False)
                avg_ot = ag_ot_df["이행률"].mean()
                fig_ot.add_vline(x=avg_ot, line_dash="dash", line_color="#94a3b8",
                                 annotation_text=f"평균 {avg_ot:.1f}%")
                st.plotly_chart(fig_ot, use_container_width=True)

        with col_ot_b:
            # 기간별 이행률 추이 (상담사별 분리)
            trend_ot_rows = []
            for ag in ch_src_ot["상담사"].unique():
                adf_ot2 = ch_src_ot[ch_src_ot["상담사"]==ag]
                for period_ot in sorted(adf_ot2["평가기간"].unique(), key=period_key):
                    pf_ot = adf_ot2[adf_ot2["평가기간"]==period_ot]
                    sc_ot2 = pf_ot[sc_col_ot].dropna().tolist() if sc_col_ot in pf_ot.columns else []
                    cr_ot2 = comp_rate(sc_ot2,[mx_ot]*len(sc_ot2))
                    if cr_ot2 is not None:
                        trend_ot_rows.append({"기간":period_ot,"상담사":ag,"이행률":cr_ot2})
            if trend_ot_rows:
                tdf_ot = pd.DataFrame(trend_ot_rows)
                fig_ot2 = px.line(tdf_ot, x="기간", y="이행률", color="상담사",
                                   markers=True, text="이행률")
                plt_theme(fig_ot2, f"[{item_sel_ot}] 상담사별 기간별 추이", 300)
                fig_ot2.update_traces(line=dict(width=2), marker=dict(size=6),
                                      texttemplate="%{text:.0f}%", textposition="top center",
                                      textfont=dict(size=9,color="#1e2330"))
                fig_ot2.add_hline(y=90, line_dash="dot", line_color="#22c55e", annotation_text="목표 90%")
                fig_ot2.add_hline(y=75, line_dash="dot", line_color="#ef4444", annotation_text="위험 75%")
                st.plotly_chart(fig_ot2, use_container_width=True)

        # 감점 코멘트 이슈 분류
        st.markdown("---")
        st.markdown(f"#### [{item_sel_ot}] 감점 코멘트 이슈 분류 분석")
        if rs_col_ot in ch_src_ot.columns:
            all_cmts_ot = []
            cmt_recs_ot = []
            for _, row_ot in ch_src_ot.iterrows():
                cmt = row_ot[rs_col_ot]
                sc_v = row_ot[sc_col_ot] if sc_col_ot in row_ot.index else None
                if cmt and cmt not in ("","None","nan") and _num(sc_v) is not None and _num(sc_v) < mx_ot:
                    cats = classify_issue(str(cmt))
                    all_cmts_ot.extend(cats)
                    for cat in cats:
                        cmt_recs_ot.append({
                            "이슈유형": cat,
                            "차수": row_ot.get("평가기간",""),
                            "상담사": row_ot.get("상담사",""),
                            "코멘트": cmt,
                            "점수": _num(sc_v),
                        })

            if all_cmts_ot:
                col_iss1, col_iss2 = st.columns(2)
                with col_iss1:
                    ic_ot = Counter(all_cmts_ot)
                    iss_ot_df = pd.DataFrame(ic_ot.most_common(12), columns=["이슈유형","건수"])
                    fig_ic_ot = px.bar(iss_ot_df, x="건수", y="이슈유형", orientation="h",
                                       color="건수", color_continuous_scale=["#fef3c7","#f59e0b"],
                                       text="건수")
                    plt_theme(fig_ic_ot, f"[{item_sel_ot}] 감점 이슈 유형", max(280, len(iss_ot_df)*28))
                    fig_ic_ot.update_traces(textposition="outside", textfont=dict(size=10,color="#1e2330"), cliponaxis=False)
                    fig_ic_ot.update_coloraxes(showscale=False)
                    st.plotly_chart(fig_ic_ot, use_container_width=True)
                with col_iss2:
                    # 상담사별 이슈 분포
                    if cmt_recs_ot:
                        cmt_df_ot = pd.DataFrame(cmt_recs_ot)
                        ag_iss_pvt = cmt_df_ot.groupby(["상담사","이슈유형"]).size().unstack(fill_value=0)
                        if len(ag_iss_pvt):
                            fig_ag_iss = go.Figure(data=go.Heatmap(
                                z=ag_iss_pvt.values, x=ag_iss_pvt.columns.tolist(), y=ag_iss_pvt.index.tolist(),
                                colorscale=[[0,"#f8fafc"],[0.3,"#fef3c7"],[1,"#ef4444"]],
                                text=ag_iss_pvt.values, texttemplate="%{text}",
                                textfont=dict(size=11,color="#374151"),
                            ))
                            plt_theme(fig_ag_iss, f"상담사 × 이슈 유형", max(260, len(ag_iss_pvt)*38))
                            fig_ag_iss.update_layout(xaxis=dict(tickangle=-30))
                            st.plotly_chart(fig_ag_iss, use_container_width=True)

                # 원본 코멘트 테이블
                st.markdown("**📋 감점 원본 코멘트 (차수 포함)**")
                if cmt_recs_ot:
                    cmt_show = pd.DataFrame(cmt_recs_ot)[["차수","상담사","이슈유형","점수","코멘트"]]
                    st.dataframe(cmt_show.reset_index(drop=True), use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# 평가자 분석
# ══════════════════════════════════════════════════════════════════
elif page == "평가자 분석":
    sh("평가자 행동 분석", "엄격도·일관성·편향 비교")
    tab1,tab2,tab3 = st.tabs(["점수 분포","엄격도 편향","평가자 × 상담사"])
    with tab1:
        fig = px.box(df, x="평가자", y="TOTAL", color="평가자", labels={"TOTAL":"점수"})
        plt_theme(fig, "평가자별 점수 분포", 340); fig.update_layout(showlegend=False)
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
            text=ln["편향"].apply(lambda x:f"{x:+.2f}"), textposition="outside",
            textfont=dict(size=10,color="#1e2330"),
        ))
        plt_theme(fig, "평가자별 점수 편향", 280)
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
        plt_theme(fig, "평가자 × 상담사 평균 점수", max(280, len(pivot3)*45))
        st.plotly_chart(fig, use_container_width=True)

st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)
st.markdown("<div style='border-top:1px solid #e2e6ea;padding:14px 0;text-align:center;color:#94a3b8;font-size:11px'>QA Monitoring Dashboard · Google Sheets 연동</div>", unsafe_allow_html=True)
