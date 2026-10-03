"""
팀플 메이트 - 바이브코딩으로 만든 팀 프로젝트 도구
시작(로그인) · 홈 · 파트 분배 · 자료 내 추적 · 기여도 평가 · AI 도우미
실행: py -m streamlit run app.py
"""
import base64
import copy
import hashlib
import hmac
import html
import json
import math
import os
import random
import re
import threading
import time
import uuid
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

try:
    import requests   # Gemini 호출용 (streamlit 설치 시 함께 설치됨)
except ImportError:
    requests = None

try:
    import anthropic   # 진짜 AI 기능 (pip install anthropic)
except ImportError:
    anthropic = None

try:
    from docx import Document
except ImportError:  # python-docx 미설치 시 .docx 업로드만 비활성화
    Document = None

st.set_page_config(page_title="팀플 메이트", page_icon="🧩", layout="wide")

CREDIT = "바이브코딩으로 만든 팀플 도구"   # 사이드바 하단 문구 (자유롭게 바꾸세요)

# ──────────────────────────────────────────────────────────────
# 기본 데이터
# ──────────────────────────────────────────────────────────────
TRAITS = ["리더형", "분석형", "창의형", "실행형", "조율형"]
TASKS = ["자료조사", "기획", "디자인", "발표", "글쓰기", "데이터분석"]

# 결과물 종류 → 역할 우선순위와 자료 평가 기준에 영향
#   weights: (적합성, 신뢰도, 분량) 가중치 / target: 만점 분량 / max: 이보다 길면 감점
#   cred_mult: 신뢰도 보정 (1보다 작으면 엄격) / boosts: 역할 우선순위 가산점
OUTPUT_TYPES = {
    "탐구보고서": {
        "weights": (0.40, 0.35, 0.25), "target": 2000, "max": 6000, "cred_mult": 1.0,
        "boosts": {"자료조사": 2, "데이터분석": 2, "글쓰기": 1},
        "focus": "근거와 분석 과정이 중요해요.",
    },
    "논문": {
        "weights": (0.30, 0.50, 0.20), "target": 4000, "max": 15000, "cred_mult": 0.85,
        "boosts": {"자료조사": 2, "글쓰기": 2, "데이터분석": 1},
        "focus": "출처와 인용을 가장 엄격하게 봐요.",
    },
    "발표자료": {
        "weights": (0.45, 0.25, 0.30), "target": 800, "max": 2500, "cred_mult": 1.1,
        "boosts": {"발표": 2, "디자인": 2, "기획": 1},
        "focus": "핵심만 간결하게. 너무 길면 감점돼요.",
    },
    "포스터·카드뉴스": {
        "weights": (0.50, 0.20, 0.30), "target": 400, "max": 1200, "cred_mult": 1.2,
        "boosts": {"디자인": 3, "기획": 1},
        "focus": "짧고 주제에 딱 맞는 문구가 중요해요.",
    },
    "영상": {
        "weights": (0.45, 0.20, 0.35), "target": 1000, "max": 3000, "cred_mult": 1.2,
        "boosts": {"디자인": 2, "기획": 2, "발표": 1},
        "focus": "대본 분량과 흐름을 중심으로 봐요.",
    },
    "기타": {
        "weights": (0.34, 0.33, 0.33), "target": 1500, "max": 6000, "cred_mult": 1.0,
        "boosts": {},
        "focus": "세 항목을 고르게 봐요.",
    },
}

ROLES = {
    "기획": {
        "title": "기획 리드",
        "desc": "전체 방향과 일정을 잡고, 파트 사이의 흐름을 연결합니다.",
        "todo": ["목차·일정표 만들기", "중간 점검 회의 진행", "파트별 결과물 취합"],
        "keywords": ["기획", "계획", "제안", "전략", "캠페인", "프로젝트", "아이디어", "정책", "해결", "개선", "방안"],
    },
    "자료조사": {
        "title": "리서처",
        "desc": "주제와 관련된 문헌·기사·통계를 찾아 출처와 함께 정리합니다.",
        "todo": ["핵심 자료 5건 이상 수집", "출처 목록 정리", "요약 노트 공유"],
        "keywords": ["조사", "연구", "역사", "사례", "현황", "문헌", "인터뷰", "탐구", "비교", "원인", "영향"],
    },
    "글쓰기": {
        "title": "보고서 작성자",
        "desc": "보고서 본문과 요약문, 발표 대본을 씁니다.",
        "todo": ["보고서 초안 작성", "문장·맞춤법 다듬기", "요약문 작성"],
        "keywords": ["보고서", "에세이", "글", "소논문", "기사", "대본", "스토리", "리포트", "논문", "신문"],
    },
    "발표": {
        "title": "발표자",
        "desc": "최종 결과를 정리해 발표하고 질의응답을 맡습니다.",
        "todo": ["발표 흐름 구성", "리허설 2회 이상", "예상 질문 정리"],
        "keywords": ["발표", "토론", "설득", "소개", "공유", "피칭", "시연"],
    },
    "디자인": {
        "title": "비주얼 디자이너",
        "desc": "슬라이드, 포스터, 인포그래픽 같은 시각 자료를 만듭니다.",
        "todo": ["슬라이드 템플릿 제작", "도표·이미지 제작", "최종 디자인 검수"],
        "keywords": ["디자인", "포스터", "홍보", "브랜드", "영상", "시각", "앱", "웹", "ui", "카드뉴스", "광고"],
    },
    "데이터분석": {
        "title": "데이터 분석가",
        "desc": "설문·통계 데이터를 정리하고 그래프로 해석합니다.",
        "todo": ["데이터 수집·정리", "그래프 제작", "분석 결과 해석"],
        "keywords": ["데이터", "통계", "설문", "분석", "실험", "측정", "그래프", "ai", "인공지능", "추이", "변화"],
    },
}

TRAIT_AFFINITY = {
    "리더형": {"기획": 3, "발표": 2},
    "분석형": {"데이터분석": 3, "자료조사": 2},
    "창의형": {"디자인": 3, "기획": 1, "글쓰기": 1},
    "실행형": {"자료조사": 2, "글쓰기": 2, "디자인": 1},
    "조율형": {"발표": 2, "기획": 2, "글쓰기": 1},
}

TRAIT_COLOR = {
    "리더형": "#3F44D9", "분석형": "#0E8C9B", "창의형": "#C2467E",
    "실행형": "#12A386", "조율형": "#C9822A",
}

# ──────────────────────────────────────────────────────────────
# 스타일
# ──────────────────────────────────────────────────────────────
st.markdown(
    """
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css');
:root {
  --ink: #1A1F36; --ink-soft: #5A6178; --indigo: #3F44D9; --indigo-soft: #ECEDFD;
  --mint: #12A386; --amber: #C9822A; --red: #C2453D; --line: #E3E7F0; --night: #151A33;
}
html, body, .stApp, [class*="css"], input, textarea, button, select {
  font-family: 'Pretendard Variable', Pretendard, -apple-system, 'Apple SD Gothic Neo',
               'Malgun Gothic', sans-serif !important;
}
.stApp { background: #F4F6FB; }
.block-container { padding-top: 2.2rem; max-width: 1080px; }
h1, h2, h3, h4 { color: var(--ink); letter-spacing: -0.02em; }

[data-testid="stSidebar"] { background: var(--night); }
[data-testid="stSidebar"] * { color: #D9DDF0 !important; }
[data-testid="stSidebar"] .brand { font-size: 1.35rem; font-weight: 800; color: #fff !important; letter-spacing: -0.03em; }
[data-testid="stSidebar"] .brand-sub { font-size: 0.82rem; color: #8F96B8 !important; margin-bottom: 1.6rem; }
[data-testid="stSidebar"] [role="radiogroup"] label { padding: 0.55rem 0.75rem; border-radius: 10px; margin-bottom: 0.2rem; width: 100%; }
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) { background: rgba(255,255,255,0.09); }
[data-testid="stSidebar"] [role="radiogroup"] label p { font-size: 0.98rem; font-weight: 600; }
[data-testid="stSidebar"] .side-stat {
  border-top: 1px solid rgba(255,255,255,0.08); margin-top: 1.4rem; padding-top: 1rem;
  font-size: 0.85rem; line-height: 1.8; color: #A9AFCC !important;
}
[data-testid="stSidebar"] .credit { font-size: 0.75rem; color: #6E7598 !important; margin-top: 2rem; }

.page-title { font-size: 2rem; font-weight: 800; color: var(--ink); letter-spacing: -0.035em; margin: 0; }
.page-lead { color: var(--ink-soft); font-size: 1rem; margin: 0.35rem 0 1.4rem; }

/* 안내 메모 */
.note {
  background: #fff; border: 1px solid var(--line); border-left: 4px solid var(--indigo);
  border-radius: 10px; padding: 0.8rem 1rem; color: var(--ink); line-height: 1.65; margin: 0.3rem 0 1.1rem;
}
.note b { color: var(--indigo); }
.note.warn { border-left-color: var(--amber); }
.ai-box {
  background: #fff; border: 1px solid #CFD1F7; border-radius: 12px; padding: 0.85rem 1rem;
  margin: 0.4rem 0 1rem; color: var(--ink); line-height: 1.7;
}
.ai-box .ai-tag { display: inline-block; font-size: 0.74rem; font-weight: 800; color: #fff;
  background: var(--indigo); border-radius: 6px; padding: 0.1rem 0.45rem; margin-bottom: 0.35rem; }
.ai-tip { margin-top: 0.6rem; font-size: 0.86rem; color: var(--ink); background: #F6F6FE;
  border-radius: 8px; padding: 0.5rem 0.65rem; line-height: 1.55; }
.ai-tip b { color: var(--indigo); }

.member-label { font-weight: 700; color: var(--ink); font-size: 0.95rem; }

.role-card {
  background: #fff; border: 1px solid var(--line); border-radius: 14px; border-top: 4px solid var(--indigo);
  padding: 1.1rem 1.15rem 1rem; margin-bottom: 1rem;
}
.role-card .who { display: flex; justify-content: space-between; align-items: center; }
.role-card .name { font-weight: 800; font-size: 1.1rem; color: var(--ink); }
.pill { font-size: 0.75rem; font-weight: 700; padding: 0.15rem 0.55rem; border-radius: 999px; color: #fff; }
.role-card .role { font-size: 1.35rem; font-weight: 800; color: var(--indigo); margin: 0.55rem 0 0.15rem; letter-spacing: -0.02em; }
.role-card .desc { color: var(--ink-soft); font-size: 0.9rem; line-height: 1.55; }
.role-card .why { margin: 0.7rem 0 0.4rem; display: flex; flex-wrap: wrap; gap: 0.3rem; }
.why-chip { background: var(--indigo-soft); color: var(--indigo); font-size: 0.76rem; font-weight: 600; padding: 0.18rem 0.5rem; border-radius: 6px; }
.role-card ul { margin: 0.3rem 0 0; padding-left: 1.1rem; color: var(--ink); font-size: 0.88rem; line-height: 1.7; }

.score-head { display: flex; justify-content: space-between; align-items: baseline; margin-top: 0.4rem; }
.score-name { font-weight: 700; color: var(--ink); font-size: 1rem; }
.score-name small { color: var(--ink-soft); font-weight: 600; margin-left: 0.3rem; }
.score-num { font-weight: 800; font-size: 1.6rem; color: var(--ink); letter-spacing: -0.03em; }
.score-num small { font-size: 0.8rem; color: var(--ink-soft); font-weight: 600; }
.score-note { color: var(--ink-soft); font-size: 0.84rem; margin: -0.3rem 0 0.9rem; }
.kw { display: inline-block; font-size: 0.78rem; padding: 0.1rem 0.45rem; border-radius: 5px; margin: 0 0.2rem 0.2rem 0; }
.kw.hit { background: #E3F6F1; color: #0B7A64; }
.kw.miss { background: #F1F2F6; color: #8A90A6; text-decoration: line-through; }

/* 게시판 */
.post-head { display: flex; justify-content: space-between; align-items: center; gap: 0.6rem; }
.post-title { font-weight: 800; font-size: 1.05rem; color: var(--ink); }
.post-meta { color: var(--ink-soft); font-size: 0.82rem; margin: 0.15rem 0 0.5rem; }
.total-badge { font-weight: 800; font-size: 0.85rem; padding: 0.2rem 0.6rem; border-radius: 8px; background: var(--indigo-soft); color: var(--indigo); white-space: nowrap; }
.vote-line { font-size: 0.86rem; color: var(--ink); margin: 0.2rem 0; }
.vote-yes { color: var(--mint); font-weight: 700; }
.vote-no { color: var(--red); font-weight: 700; }
.post-body {
  white-space: pre-wrap; background: #FAFBFD; border: 1px solid var(--line); border-radius: 8px;
  padding: 0.7rem 0.85rem; max-height: 240px; overflow-y: auto; font-size: 0.9rem; line-height: 1.65; color: var(--ink);
}
.mini-scores { font-size: 0.84rem; color: var(--ink-soft); margin: 0.3rem 0 0.6rem; }

.summary-line { background: #fff; border: 1px solid var(--line); border-radius: 12px; padding: 0.8rem 1rem; margin-bottom: 0.5rem; color: var(--ink); font-size: 0.98rem; }
.summary-line .tag { font-weight: 800; margin-right: 0.5rem; }

.stButton > button { border-radius: 10px; font-weight: 700; padding: 0.5rem 1.1rem; }
.stButton > button:focus-visible { outline: 3px solid #9EA2F5; outline-offset: 2px; }
/* 시작 페이지 */
.hero-title { font-size: 2.6rem; font-weight: 800; color: var(--ink); letter-spacing: -0.04em; margin: 0.6rem 0 0; }
.hero-lead { color: var(--ink-soft); font-size: 1.05rem; margin: 0.3rem 0 1.2rem; }
.stat-row { display: flex; flex-wrap: wrap; gap: 0.6rem; margin: 0.2rem 0 1.4rem; }
.stat { flex: 1 1 140px; background: #fff; border: 1px solid var(--line); border-radius: 12px; padding: 0.75rem 1rem; }
.stat-num, .big-count { font-size: 1.6rem; font-weight: 800; color: var(--ink); letter-spacing: -0.03em; }
.stat-label { font-size: 0.82rem; color: var(--ink-soft); }
.team-line { background: #fff; border: 1px solid var(--line); border-radius: 10px; padding: 0.55rem 0.8rem;
  margin-bottom: 0.4rem; color: var(--ink); font-size: 0.93rem; display: flex; gap: 0.3rem; }
.team-pct { margin-left: auto; font-weight: 800; color: var(--indigo); }
[data-testid="stSidebar"] .me-line { font-size: 0.92rem; margin: -0.8rem 0 1rem; color: #fff !important; }
.me-tag { font-size: 0.7rem; font-weight: 800; background: var(--indigo); color: #fff; border-radius: 5px;
  padding: 0.05rem 0.35rem; vertical-align: middle; }
.role-card.mine { box-shadow: 0 0 0 2px var(--indigo) inset; }
.home-role { font-size: 1.5rem; font-weight: 800; color: var(--indigo); letter-spacing: -0.02em; }
.desc { color: var(--ink-soft); font-size: 0.92rem; line-height: 1.55; }
/* ── 다크 모드 브라우저에서도 밝은 화면 고정 (config.toml 없어도 동작) ── */
:root { color-scheme: light; }
[data-testid="stHeader"] { background: transparent; }
[data-testid="stHeader"] [data-testid="stToolbar"] button,
[data-testid="stHeader"] [data-testid="stToolbar"] a,
[data-testid="stHeader"] [data-testid="stToolbar"] span { color: var(--ink) !important; }
.block-container { padding-top: 3.4rem; }
:is([data-testid="stMain"], section.stMain), [data-testid="stBottom"] > div { background: var(--bg); color: var(--ink); }
:is([data-testid="stMain"], section.stMain) [data-testid="stMarkdownContainer"] { color: var(--ink); }
:is([data-testid="stMain"], section.stMain) [data-testid="stMarkdownContainer"] :is(p, li, h1, h2, h3, h4, h5, h6, strong, em) { color: var(--ink); }
:is([data-testid="stMain"], section.stMain) p.page-lead, :is([data-testid="stMain"], section.stMain) p.hero-lead { color: var(--ink-soft); }
:is([data-testid="stMain"], section.stMain) [data-testid="stCaptionContainer"],
:is([data-testid="stMain"], section.stMain) [data-testid="stCaptionContainer"] p { color: var(--ink-soft) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stWidgetLabel"] p { color: var(--ink) !important; font-weight: 600; }

/* 입력칸 */
:is([data-testid="stMain"], section.stMain) [data-baseweb="input"], :is([data-testid="stMain"], section.stMain) [data-baseweb="base-input"],
:is([data-testid="stMain"], section.stMain) [data-baseweb="textarea"], :is([data-testid="stMain"], section.stMain) [data-baseweb="select"] > div,
[data-testid="stBottom"] [data-baseweb="textarea"], [data-testid="stChatInput"] {
  background-color: #fff !important; border-color: #D5DAE5 !important; }
:is([data-testid="stMain"], section.stMain) input, :is([data-testid="stMain"], section.stMain) textarea, [data-testid="stBottom"] textarea,
:is([data-testid="stMain"], section.stMain) [data-baseweb="select"] div, :is([data-testid="stMain"], section.stMain) [data-baseweb="select"] svg {
  color: var(--ink) !important; -webkit-text-fill-color: var(--ink); background-color: transparent; }
:is([data-testid="stMain"], section.stMain) input::placeholder, :is([data-testid="stMain"], section.stMain) textarea::placeholder,
[data-testid="stBottom"] textarea::placeholder { color: #9AA0B4 !important; -webkit-text-fill-color: #9AA0B4; }
:is([data-testid="stMain"], section.stMain) [data-testid="stNumberInput"] button { background: #fff !important; color: var(--ink) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stFileUploaderDropzone"] { background: #fff !important; border: 1px dashed #C9CEDB; }
:is([data-testid="stMain"], section.stMain) [data-testid="stFileUploaderDropzone"] :is(span, small, div) { color: var(--ink-soft) !important; }

/* 버튼 */
:is([data-testid="stMain"], section.stMain) button:is([kind="secondary"], [kind="secondaryFormSubmit"], [kind="tertiary"]) {
  background: #fff !important; color: var(--ink) !important; border: 1px solid #D5DAE5 !important; }
:is([data-testid="stMain"], section.stMain) button:is([kind="secondary"], [kind="secondaryFormSubmit"]):hover:not(:disabled) {
  border-color: var(--indigo) !important; color: var(--indigo) !important; }
button:is([kind="primary"], [kind="primaryFormSubmit"]) {
  background: var(--indigo) !important; border-color: var(--indigo) !important; color: #fff !important; }
button:is([kind="primary"], [kind="primaryFormSubmit"]):hover:not(:disabled) { background: #3337B8 !important; }
:is([data-testid="stMain"], section.stMain) button:is([kind="primary"], [kind="primaryFormSubmit"]) p { color: #fff !important; }
:is([data-testid="stMain"], section.stMain) button[kind="secondary"]:hover:not(:disabled) p { color: var(--indigo) !important; }
:is([data-testid="stMain"], section.stMain) button[kind^="segmented_control"] { background: #fff !important; color: var(--ink) !important; border-color: #D5DAE5 !important; }
:is([data-testid="stMain"], section.stMain) button[kind="segmented_controlActive"] { background: var(--indigo-soft) !important; border-color: var(--indigo) !important; }
:is([data-testid="stMain"], section.stMain) button[kind="segmented_controlActive"] p { color: var(--indigo) !important; font-weight: 700; }

/* 체크박스 · 진행바 · 탭 · 접는 상자 · 테두리 상자 */
:is([data-testid="stMain"], section.stMain) [data-testid="stCheckbox"] label > span + div { background-color: #fff; border: 1.5px solid #B8BDCC; }
:is([data-testid="stMain"], section.stMain) [data-testid="stCheckbox"] label:has(input:checked) > span + div { background-color: var(--indigo); border-color: var(--indigo); }
[data-testid="stProgressBarTrack"] { background-color: #E3E7F0 !important; }
[data-testid="stProgressBarTrack"] > div { background-color: var(--indigo) !important; }
:is([data-testid="stMain"], section.stMain) [data-baseweb="tab-highlight"] { background-color: var(--indigo) !important; }
:is([data-testid="stMain"], section.stMain) [data-baseweb="tab-border"] { background-color: var(--line) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stExpander"] details { background: #fff; border-color: var(--line) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stExpander"] summary,
:is([data-testid="stMain"], section.stMain) [data-testid="stExpander"] summary * { color: var(--ink) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stExpander"] summary:hover { background: #F6F7FB; }
:is([data-testid="stMain"], section.stMain) [data-testid="stVerticalBlockBorderWrapper"],
:is([data-testid="stMain"], section.stMain) [data-testid="stForm"] { border-color: var(--line) !important; }
:is([data-testid="stMain"], section.stMain) hr { border-color: var(--line) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stChatMessage"] { background: #fff; border: 1px solid var(--line); }

/* 사이드바 메뉴: 동그라미 대신 선택 줄 강조 */
[data-testid="stSidebar"] [data-testid="stRadioOption"] > div > div:first-child:not([data-testid]) { display: none; }
[data-testid="stSidebar"] button:is([kind="secondary"], [kind="tertiary"]) {
  background: rgba(255,255,255,0.07) !important; border: 1px solid rgba(255,255,255,0.16) !important; }
[data-testid="stSidebar"] button[kind="tertiary"] { background: transparent !important; border: none !important;
  padding: 0 !important; min-height: 0 !important; }
[data-testid="stSidebar"] button[kind="tertiary"] p { font-size: 0.75rem !important; color: #6E7598 !important; text-decoration: underline; }
[data-testid="stSidebar"] [data-baseweb="input"], [data-testid="stSidebar"] [data-baseweb="select"] > div {
  background-color: rgba(255,255,255,0.08) !important; }
/* Streamlit 최신 버전 입력칸 구조 */
:is([data-testid="stMain"], section.stMain) :is([data-testid="stTextInputRootElement"], [data-testid="stTextAreaRootElement"],
  [data-testid="stDateInputField"], [data-testid="stNumberInputContainer"]),
[data-testid="stBottom"] [data-testid="stChatInput"], [data-testid="stBottom"] [data-testid="stChatInput"] > div {
  background-color: #fff !important; border-color: #D5DAE5 !important; color: var(--ink) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stSelectbox"] div:not([data-testid="stWidgetLabel"] div) {
  background-color: #fff !important; color: var(--ink) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stSelectbox"] :is(span, svg) { color: var(--ink) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stSelectbox"] > div > div { border-color: #D5DAE5 !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stNumberInputContainer"] button { background: #fff !important; color: var(--ink) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stDateInputField"] * { color: var(--ink) !important; }
[data-testid="stBottom"] { background: var(--bg) !important; }
[data-testid="stBottom"] > div, [data-testid="stBottomBlockContainer"] { background: var(--bg) !important; }
/* 위쪽 보기 전환 버튼 */
:is([data-testid="stMain"], section.stMain) button[data-variant="segmented_control"] {
  background: #fff !important; border-color: #D5DAE5 !important; }
:is([data-testid="stMain"], section.stMain) button[data-variant="segmented_control"] p { color: var(--ink) !important; }
:is([data-testid="stMain"], section.stMain) button[data-variant="segmented_control"][aria-checked="true"] {
  background: var(--indigo-soft) !important; border-color: var(--indigo) !important; }
:is([data-testid="stMain"], section.stMain) button[data-variant="segmented_control"][aria-checked="true"] p { color: var(--indigo) !important; font-weight: 700; }
/* 큰 제목 크기 고정 (Streamlit 기본 글자 크기에 덮이지 않게) */
:is([data-testid="stMain"], section.stMain) p.page-title { font-size: 2rem !important; font-weight: 800 !important; line-height: 1.25 !important; margin: 0 !important; }
:is([data-testid="stMain"], section.stMain) p.hero-title { font-size: 2.6rem !important; font-weight: 800 !important; line-height: 1.2 !important; margin: 0.6rem 0 0 !important; }
:is([data-testid="stMain"], section.stMain) p.page-lead, :is([data-testid="stMain"], section.stMain) p.hero-lead { font-size: 1.02rem !important; margin: 0.35rem 0 1.3rem !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stCheckbox"] [data-testid="stWidgetLabel"] p { font-weight: 500; }
[data-testid="stSidebar"] [data-testid="stRadioOption"] { width: 100%; }
/* 탭 밑줄 · 본문 라디오 선택 표시를 남색으로 */
.react-aria-SelectionIndicator { background-color: var(--indigo) !important; border-color: var(--indigo) !important; }
[data-testid="stTab"][aria-selected="true"] { border-bottom-color: var(--indigo) !important; }
:is([data-testid="stMain"], section.stMain) [data-testid="stRadioOption"][data-selected="true"] > div > div:first-child {
  background-color: var(--indigo) !important; border-color: var(--indigo) !important; }
</style>
""",
    unsafe_allow_html=True,
)


# ──────────────────────────────────────────────────────────────
# 저장 파일 (팀 정보 + 게시판). 같은 서버에 접속한 팀원 모두가 공유
# ──────────────────────────────────────────────────────────────
EMPTY_DATA = {"team": None, "posts": [], "progress": {}}
GITHUB_API = "https://api.github.com/repos/{repo}/contents/{path}"
LEGACY_DOC = "teamplay_data.json"          # 팀 코드 도입 전, 팀 하나만 있던 시절의 저장 파일
LOCAL_DIR = Path(__file__).parent / "teamplay_data"   # 내 컴퓨터 저장 폴더 (팀마다 파일 하나)
CODE_CHARS = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"        # 헷갈리는 0/O, 1/I 제외


class StorageError(Exception):
    pass


def _gh_cfg():
    """secrets.toml에 [github_storage]가 있으면 GitHub에 저장, 없으면 내 컴퓨터 파일에 저장."""
    try:
        cfg = st.secrets.get("github_storage")
    except Exception:   # secrets.toml이 없을 때
        return None
    if not cfg or not cfg.get("token") or not cfg.get("repo"):
        return None
    return {"token": cfg["token"], "repo": cfg["repo"].strip().strip("/"), "branch": cfg.get("branch")}


def _gh_headers(cfg, raw=False):
    return {"Authorization": f"Bearer {cfg['token']}", "X-GitHub-Api-Version": "2022-11-28",
            "Accept": "application/vnd.github.raw+json" if raw else "application/vnd.github+json"}


def _clean(d):
    d = d if isinstance(d, dict) else {}
    for k, v in EMPTY_DATA.items():
        d.setdefault(k, copy.deepcopy(v))
    return d


def _gh_read(cfg, path):
    """(내용, sha). 파일이 없으면 (None, None)."""
    url = GITHUB_API.format(repo=cfg["repo"], path=path)
    params = {"ref": cfg["branch"]} if cfg["branch"] else None
    r = requests.get(url, headers=_gh_headers(cfg), params=params, timeout=20)
    if r.status_code == 404:
        meta = requests.get(f"https://api.github.com/repos/{cfg['repo']}", headers=_gh_headers(cfg), timeout=20)
        if meta.status_code != 200:
            raise StorageError("데이터 저장소를 찾지 못했어요. repo 이름과 토큰 권한을 확인해 주세요.")
        return None, None
    if r.status_code in (401, 403):
        raise StorageError("GitHub 토큰이 맞지 않거나 만료됐어요. 새 토큰을 만들어 Secrets에 넣어 주세요.")
    if r.status_code >= 400:
        raise StorageError(f"GitHub에서 읽지 못했어요 (코드 {r.status_code}).")
    meta = r.json()
    if meta.get("content") and meta.get("encoding") == "base64":
        text = base64.b64decode(meta["content"]).decode("utf-8")
    else:   # 1MB가 넘으면 내용이 따로 와서 한 번 더 받음
        text = requests.get(url, headers=_gh_headers(cfg, raw=True), params=params, timeout=30).text
    return json.loads(text or "{}"), meta.get("sha")


def _gh_write(cfg, path, d, sha):
    url = GITHUB_API.format(repo=cfg["repo"], path=path)
    body = {"message": f"팀플 메이트 자동 저장 {path} {datetime.now():%Y-%m-%d %H:%M:%S}",
            "content": base64.b64encode(json.dumps(d, ensure_ascii=False, indent=1).encode("utf-8")).decode()}
    if cfg["branch"]:
        body["branch"] = cfg["branch"]
    for _ in range(3):
        if sha:
            body["sha"] = sha
        else:
            body.pop("sha", None)
        r = requests.put(url, headers=_gh_headers(cfg), json=body, timeout=30)
        if r.status_code in (200, 201):
            return r.json()["content"]["sha"]
        if r.status_code in (409, 422):   # 버전이 어긋남 → 최신 버전 번호를 받아 다시 저장
            g = requests.get(url, headers=_gh_headers(cfg), timeout=20)
            sha = g.json().get("sha") if g.status_code == 200 else None
            continue
        if r.status_code in (401, 403):
            raise StorageError("GitHub에 저장할 권한이 없어요. 토큰의 Contents 권한(Read and write)을 확인해 주세요.")
        raise StorageError(f"GitHub에 저장하지 못했어요 (코드 {r.status_code}).")
    raise StorageError("GitHub 저장이 계속 충돌해요. 잠시 후 다시 시도해 주세요.")


def _local_file(path):
    return Path(__file__).parent / LEGACY_DOC if path == LEGACY_DOC else LOCAL_DIR / path


@st.cache_resource
def _store():
    """앱을 쓰는 모든 사람이 함께 보는 데이터 (서버 메모리). 바뀌면 저장소에 기록."""
    return {"docs": {}, "lock": threading.Lock(), "dirty": set(), "writer": None, "error": None,
            "legacy_checked": False, "keys": {}, "admin_fails": []}


def _read_doc(path):
    """lock 안에서 호출. 처음 읽는 문서만 저장소에서 가져옴."""
    s = _store()
    if path in s["docs"]:
        return s["docs"][path]
    cfg = _gh_cfg()
    if cfg:
        data, sha = _gh_read(cfg, path)
    else:
        f = _local_file(path)
        try:
            data, sha = (json.loads(f.read_text(encoding="utf-8")) if f.exists() else None), None
        except (json.JSONDecodeError, OSError):
            data, sha = None, None
    s["docs"][path] = {"data": data, "sha": sha}
    return s["docs"][path]


def _put_doc(path, data):
    """lock 안에서 호출. 메모리에 반영하고 저장 예약."""
    s = _store()
    doc = s["docs"].setdefault(path, {"data": None, "sha": None})
    doc["data"] = copy.deepcopy(data)
    cfg = _gh_cfg()
    if not cfg:
        f = _local_file(path)
        f.parent.mkdir(parents=True, exist_ok=True)
        tmp = f.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(f)
        return
    s["dirty"].add(path)
    if s["writer"] is None or not s["writer"].is_alive():
        s["writer"] = threading.Thread(target=_flush_loop, args=(s, cfg), daemon=True)
        s["writer"].start()


def _flush_loop(s, cfg):
    """GitHub 저장은 1초쯤 걸려서, 화면이 멈추지 않게 뒤에서 모아서 저장."""
    while True:
        with s["lock"]:
            if not s["dirty"]:
                s["writer"] = None
                return
            path = s["dirty"].pop()
            snapshot, sha = copy.deepcopy(s["docs"][path]["data"]), s["docs"][path]["sha"]
        try:
            new_sha = _gh_write(cfg, path, snapshot, sha)
            with s["lock"]:
                s["docs"][path]["sha"], s["error"] = new_sha, None
        except Exception as e:
            with s["lock"]:
                s["error"] = str(e)
                s["dirty"].add(path)
            time.sleep(5)   # 잠시 쉬었다 다시 시도
        time.sleep(0.3)


def team_doc(code):
    return f"teams/{code}.json"


def team_exists(code):
    """팀 코드가 있으면 True. 저장소를 못 읽으면 StorageError."""
    if not code or len(code) != 6 or any(c not in CODE_CHARS for c in code):
        return False
    s = _store()
    with s["lock"]:
        try:
            doc = _read_doc(team_doc(code))
        except StorageError as e:
            s["error"] = str(e)
            raise
        return bool(doc["data"] and doc["data"].get("team"))


def new_team_code():
    rng = random.SystemRandom()
    for _ in range(20):
        code = "".join(rng.choice(CODE_CHARS) for _ in range(6))
        if not team_exists(code):
            return code
    raise StorageError("팀 코드를 만들지 못했어요. 다시 시도해 주세요.")


def legacy_code():
    """팀 코드 도입 전에 만든 팀이 있으면 코드를 붙여 옮기고, 안내 중인 코드를 돌려줌."""
    s = _store()
    with s["lock"]:
        try:
            doc = _read_doc(LEGACY_DOC)
        except StorageError:
            return None
        d = doc["data"] or {}
        if d.get("team") and not d.get("migrated_to"):   # 처음 한 번만 옮김
            rng = random.SystemRandom()
            code = "".join(rng.choice(CODE_CHARS) for _ in range(6))
            pointer = {"migrated_to": code, "show_notice": True}
            cfg = _gh_cfg()
            if cfg:
                # 순서 중요: 새 팀 파일이 확실히 저장된 뒤에만 '옮겼음' 표시를 남김 (실패하면 다음에 다시 시도)
                try:
                    sha = _gh_write(cfg, team_doc(code), _clean(d), None)
                    s["docs"][team_doc(code)] = {"data": _clean(copy.deepcopy(d)), "sha": sha}
                    psha = _gh_write(cfg, LEGACY_DOC, pointer, doc["sha"])
                    s["docs"][LEGACY_DOC] = {"data": pointer, "sha": psha}
                except StorageError:
                    return None
            else:
                _put_doc(team_doc(code), _clean(d))
                _put_doc(LEGACY_DOC, pointer)
            d = s["docs"][LEGACY_DOC]["data"]
        return d.get("migrated_to") if d.get("show_notice") else None


def hide_legacy_notice():
    s = _store()
    with s["lock"]:
        doc = _read_doc(LEGACY_DOC)
        if doc["data"] and doc["data"].get("migrated_to"):
            _put_doc(LEGACY_DOC, {**doc["data"], "show_notice": False})


SETTINGS_DOC = "admin/settings.json"


def get_settings():
    """관리자 화면에서 정한 설정 (모든 팀 공통)."""
    s = _store()
    with s["lock"]:
        try:
            return dict(_read_doc(SETTINGS_DOC)["data"] or {})
        except StorageError:
            return {}


def save_settings(new):
    s = _store()
    with s["lock"]:
        _read_doc(SETTINGS_DOC)
        _put_doc(SETTINGS_DOC, new)


def admin_load(code):
    s = _store()
    with s["lock"]:
        doc = _read_doc(team_doc(code))
        return _clean(copy.deepcopy(doc["data"])) if doc["data"] and doc["data"].get("team") else None


def admin_save(code, d):
    s = _store()
    with s["lock"]:
        _read_doc(team_doc(code))
        _put_doc(team_doc(code), _clean(d))


def list_team_codes():
    """저장된 팀 코드 목록."""
    cfg = _gh_cfg()
    codes = set()
    if cfg:
        url = GITHUB_API.format(repo=cfg["repo"], path="teams")
        r = requests.get(url, headers=_gh_headers(cfg), timeout=20)
        if r.status_code == 200 and isinstance(r.json(), list):
            codes |= {f["name"][:-5] for f in r.json() if f.get("name", "").endswith(".json")}
        elif r.status_code not in (200, 404):
            raise StorageError(f"팀 목록을 읽지 못했어요 (코드 {r.status_code}).")
    elif (LOCAL_DIR / "teams").exists():
        codes |= {f.stem for f in (LOCAL_DIR / "teams").glob("*.json")}
    with _store()["lock"]:   # 방금 만들어져 아직 저장 중인 팀도 포함
        codes |= {k[6:-5] for k in _store()["docs"] if k.startswith("teams/")}
    return sorted(codes)


def storage_mode():
    return "GitHub" if _gh_cfg() else "내 컴퓨터"


def storage_status():
    s = _store()
    if s["error"]:
        return f"저장 오류: {s['error']}"
    if s["dirty"] or s["writer"]:
        return "저장하는 중..."
    return "저장됨"


def load_data():
    """지금 들어와 있는 팀(st.session_state.team_code)의 데이터. 팀이 없으면 빈 데이터."""
    code = st.session_state.get("team_code")
    if not code:
        return copy.deepcopy(EMPTY_DATA)
    s = _store()
    with s["lock"]:
        try:
            doc = _read_doc(team_doc(code))
        except StorageError as e:
            s["error"] = str(e)
            raise
        return _clean(copy.deepcopy(doc["data"])) if doc["data"] else copy.deepcopy(EMPTY_DATA)


def save_data(d):
    """지금 팀의 데이터를 저장. 저장소에서 먼저 읽은 적 있는 팀만 저장(덮어쓰기 방지)."""
    code = st.session_state.get("team_code")
    if not code:
        return
    s = _store()
    with s["lock"]:
        if team_doc(code) not in s["docs"]:
            return
        _put_doc(team_doc(code), _clean(d))


# ──────────────────────────────────────────────────────────────
# 진짜 AI (Claude API). 키가 없으면 규칙 기반으로만 동작
# ──────────────────────────────────────────────────────────────
AI_PROVIDERS = ["Gemini (무료)", "Claude"]
AI_MODELS = {   # Claude
    "빠르고 저렴함 (Haiku 4.5)": "claude-haiku-4-5-20251001",
    "더 똑똑함 (Sonnet 5.5)": "claude-sonnet-5-5",
}
GEMINI_MODELS = {   # 무료 등급에서 쓸 수 있는 Flash 계열
    "최신 Flash (추천)": "gemini-flash-latest",
    "Gemini 2.5 Flash": "gemini-2.5-flash",
    "가벼운 Flash-Lite (하루 횟수 넉넉)": "gemini-flash-lite-latest",
}
GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
AI_PERSONA = ("너는 중고등학생 팀 프로젝트를 돕는 AI '팀플 메이트'야. 한국어 해요체로, 학생이 바로 실행할 수 있게 "
              "짧고 구체적으로 답해. 과제를 대신 써 주기보다 방향과 방법을 알려 주는 쪽을 우선해.")


def provider():
    p = get_settings().get("provider")
    return p if p in AI_PROVIDERS else AI_PROVIDERS[0]


def is_gemini():
    return provider().startswith("Gemini")


def secret_key(name):
    try:
        return st.secrets.get(name, "") or ""
    except Exception:   # secrets.toml 파일이 없을 때
        return ""


def key_source(name):
    """(키, 어디서 왔는지). 관리자 임시 키 → Secrets → 환경 변수 순."""
    tmp = _store()["keys"].get(name)
    if tmp:
        return tmp, "관리자 임시 키"
    if secret_key(name):
        return secret_key(name), "Secrets"
    if os.environ.get(name):
        return os.environ[name], "환경 변수"
    return "", "없음"


def get_api_key():
    return key_source("GEMINI_API_KEY" if is_gemini() else "ANTHROPIC_API_KEY")[0]


def ai_ready():
    lib_ok = (requests is not None) if is_gemini() else (anthropic is not None)
    return lib_ok and bool(get_api_key())


GEMINI_BUSY = (500, 502, 503, 504)   # 구글 서버가 붐빌 때 오는 코드


def _ask_gemini(messages, system, max_tokens):
    chosen = GEMINI_MODELS.get(get_settings().get("gemini_model"), list(GEMINI_MODELS.values())[0])
    # 고른 모델이 붐비면 나머지 모델로 차례대로 넘어감
    models = [chosen] + [m for m in GEMINI_MODELS.values() if m != chosen]
    body = {
        "systemInstruction": {"parts": [{"text": system}]},
        "contents": [{"role": "model" if m["role"] == "assistant" else "user", "parts": [{"text": m["content"]}]}
                     for m in messages],
        # 생각(thinking)에도 토큰이 쓰이므로 넉넉하게 잡음
        "generationConfig": {"maxOutputTokens": max_tokens * 4},
    }
    last_code = None
    for model in models:
        for attempt in range(2):   # 같은 모델로 한 번 더 기다렸다 재시도
            try:
                r = requests.post(GEMINI_URL.format(model=model), json=body, timeout=60,
                                  headers={"x-goog-api-key": get_api_key()})
            except requests.RequestException:
                st.session_state.ai_error = "AI 서버에 연결하지 못했어요. 인터넷 연결을 확인해 주세요."
                return None
            last_code = r.status_code
            if r.status_code == 200:
                try:
                    parts = r.json()["candidates"][0]["content"]["parts"]
                    text = "".join(p.get("text", "") for p in parts if not p.get("thought")).strip()
                except (KeyError, IndexError, ValueError):
                    text = ""
                if text:
                    st.session_state.ai_model_used = model
                    return text
                break   # 빈 답이면 다음 모델로
            if r.status_code in GEMINI_BUSY:
                if attempt == 0:
                    time.sleep(2)
                    continue
                break   # 두 번 다 붐비면 다음 모델로
            if r.status_code in (429, 404):
                break   # 이 모델의 무료 횟수 소진 또는 모델 없음 → 다음 모델로
            if "API_KEY_INVALID" in r.text or r.status_code in (400, 401, 403):
                st.session_state.ai_error = "Gemini 키가 올바르지 않아요. 사이드바 설정에서 키를 다시 확인해 주세요."
                return None
            break

    if last_code in GEMINI_BUSY:
        st.session_state.ai_error = ("지금 구글 AI 서버가 붐벼서 답을 못 받았어요 (코드 503). "
                                     "몇 분 뒤에 다시 눌러 주세요. 키나 앱 문제는 아니에요.")
    elif last_code == 429:
        st.session_state.ai_error = "오늘 무료 사용 횟수를 다 쓴 것 같아요. 1분 뒤나 내일 다시 시도해 주세요."
    elif last_code == 404:
        st.session_state.ai_error = "쓸 수 있는 Gemini 모델을 찾지 못했어요. 앱을 최신 버전으로 받아 주세요."
    else:
        st.session_state.ai_error = (f"AI 호출 중 오류가 났어요 (코드 {last_code})." if last_code
                                     else "AI가 빈 답을 보냈어요. 질문을 조금 바꿔 다시 시도해 주세요.")
    return None


def _ask_claude(messages, system, max_tokens):
    model = AI_MODELS.get(get_settings().get("claude_model"), list(AI_MODELS.values())[0])
    try:
        client = anthropic.Anthropic(api_key=get_api_key())
        res = client.messages.create(model=model, max_tokens=max_tokens, system=system, messages=messages)
        return "".join(b.text for b in res.content if b.type == "text").strip()
    except anthropic.AuthenticationError:
        st.session_state.ai_error = "API 키가 올바르지 않아요. 사이드바 설정에서 키를 다시 확인해 주세요."
    except anthropic.RateLimitError:
        st.session_state.ai_error = "요청이 너무 많거나 크레딧이 부족해요. 잠시 후 다시 시도하거나 콘솔에서 잔액을 확인하세요."
    except anthropic.APIConnectionError:
        st.session_state.ai_error = "AI 서버에 연결하지 못했어요. 인터넷 연결을 확인해 주세요."
    except Exception as e:   # 모델 이름 오류 등
        st.session_state.ai_error = f"AI 호출 중 오류가 났어요: {e}"
    return None


def ask_ai(messages, system=AI_PERSONA, max_tokens=900):
    """선택한 AI(Gemini 또는 Claude)에게 묻고 답 텍스트를 돌려준다. 실패하면 None."""
    st.session_state.ai_error = None
    if not ai_ready():
        return None
    return (_ask_gemini if is_gemini() else _ask_claude)(messages, system, max_tokens)


def ask_ai_json(prompt, max_tokens=1200):
    text = ask_ai([{"role": "user", "content": prompt + "\n\nJSON만 출력해. 설명이나 ``` 표시는 쓰지 마."}],
                  max_tokens=max_tokens)
    if not text:
        return None
    text = re.sub(r"```(?:json)?", "", text).strip()
    a, b = text.find("{"), text.rfind("}")
    try:
        return json.loads(text[a:b + 1])
    except (ValueError, json.JSONDecodeError):
        st.session_state.ai_error = "AI 답을 해석하지 못해 기본 방식으로 처리했어요."
        return None


def show_ai_error():
    if st.session_state.get("ai_error"):
        st.warning(st.session_state.ai_error)


def ai_box(text, tag="AI 피드백"):
    body = esc(text).replace("\n", "<br>")
    st.markdown(f'<div class="ai-box"><span class="ai-tag">{tag}</span><br>{body}</div>', unsafe_allow_html=True)


def ai_personalize(assignment):
    """규칙으로 정한 역할은 그대로 두고, 주제에 맞춘 할 일 3개와 한마디를 AI가 써 준다."""
    members = [{"name": r["name"], "role": r["title"], "trait": r["trait"]} for r in assignment["result"]]
    prompt = (f"프로젝트 주제: {assignment['topic']}\n결과물: {assignment['otype']}\n"
              f"팀원과 역할: {json.dumps(members, ensure_ascii=False)}\n\n"
              "각 팀원이 이 주제에서 맡은 역할로 해야 할 구체적인 할 일 3개(각 20자 이내, 체크리스트용)와 "
              "성향을 고려한 한 문장 조언을 만들어 줘. 형식: "
              '{"members":[{"name":"이름","todo":["할 일1","할 일2","할 일3"],"tip":"조언"}]}')
    data = ask_ai_json(prompt)
    if not data:
        return False
    by_name = {m.get("name"): m for m in data.get("members", [])}
    changed = False
    for r in assignment["result"]:
        m = by_name.get(r["name"])
        if m and isinstance(m.get("todo"), list) and len(m["todo"]) >= 2:
            r["todo"] = [str(t)[:40] for t in m["todo"][:3]]
            r["tip"] = str(m.get("tip", ""))[:150]
            changed = True
    return changed


def ai_review(topic, otype, text, ev):
    prompt = (f"프로젝트 주제: {topic}\n결과물 종류: {otype}\n"
              f"규칙 기반 점수: 적합성 {ev['rel']}, 신뢰도 {ev['cred']}, 분량 {ev['len']}, 종합 {ev['total']}\n\n"
              f"아래 자료를 {otype}에 쓸 자료로서 검토해 줘. '잘한 점' 1줄, '보완할 점' 2줄, "
              "'바로 할 일' 1줄로, 각 줄 앞에 그 이름을 붙여서 써. 자료에 근거 없는 주장이나 "
              "출처가 불분명한 수치가 있으면 짚어 줘.\n\n[자료]\n" + text[:6000])
    return ask_ai([{"role": "user", "content": prompt}], max_tokens=600)


# ──────────────────────────────────────────────────────────────
# 공통 UI
# ──────────────────────────────────────────────────────────────
def esc(s):
    return html.escape(str(s))


def note(message_html, warn=False):
    st.markdown(f'<div class="note{" warn" if warn else ""}">{message_html}</div>', unsafe_allow_html=True)


def page_header(title, lead):
    st.markdown(f'<p class="page-title">{title}</p><p class="page-lead">{lead}</p>', unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# 로직 1: 파트 배정
# ──────────────────────────────────────────────────────────────
def rank_roles(topic, otype):
    """주제 키워드 매칭 수 + 결과물 종류 가산점으로 역할 우선순위를 정한다."""
    t = topic.lower()
    boosts = OUTPUT_TYPES[otype]["boosts"]
    ranked = []
    for order, (task, info) in enumerate(ROLES.items()):
        hits = [k for k in info["keywords"] if k in t]
        ranked.append((task, hits, len(hits) + boosts.get(task, 0), order))
    ranked.sort(key=lambda x: (-x[2], x[3]))
    return [(task, hits) for task, hits, _, _ in ranked]


def assign_roles(topic, otype, members):
    ranked = rank_roles(topic, otype)
    hits_map = dict(ranked)
    boosts = OUTPUT_TYPES[otype]["boosts"]
    n = len(members)

    slots = []
    while len(slots) < n:
        for task, _ in ranked:
            if len(slots) >= n:
                break
            slots.append({"task": task, "sub": any(s["task"] == task for s in slots), "priority": len(slots)})

    def score(m, slot):
        task = slot["task"]
        s = 10.0 if task in m["prefs"] else 0.0
        s += TRAIT_AFFINITY.get(m["trait"], {}).get(task, 0)
        s += 0.5 * len(hits_map.get(task, [])) + 0.5 * boosts.get(task, 0)
        return s - slot["priority"] * 0.01

    pairs = sorted(((score(m, sl), mi, si) for mi, m in enumerate(members) for si, sl in enumerate(slots)),
                   key=lambda x: (-x[0], x[1], x[2]))
    used_m, used_s, result = set(), set(), {}
    for _, mi, si in pairs:
        if mi in used_m or si in used_s:
            continue
        used_m.add(mi); used_s.add(si); result[mi] = slots[si]

    out = []
    for mi, m in enumerate(members):
        slot = result[mi]
        task = slot["task"]
        reasons = []
        if task in m["prefs"]:
            reasons.append(f"선호 작업 '{task}' 일치")
        if TRAIT_AFFINITY.get(m["trait"], {}).get(task, 0) >= 2:
            reasons.append(f"{m['trait']} 성향과 잘 맞음")
        if boosts.get(task, 0) >= 2:
            reasons.append(f"{otype}에서 중요한 역할")
        if hits_map.get(task):
            reasons.append("주제 키워드: " + ", ".join(hits_map[task][:3]))
        if not reasons:
            reasons.append("남은 역할 중 가장 적합")
        info = ROLES[task]
        out.append({
            "name": m["name"], "trait": m["trait"], "task": task,
            "title": info["title"] + (" (보조)" if slot["sub"] else ""),
            "desc": info["desc"], "todo": info["todo"], "reasons": reasons,
            "pref_matched": task in m["prefs"],
        })
    return out, ranked


# ──────────────────────────────────────────────────────────────
# 로직 2: 자료 평가 (결과물 종류에 따라 기준이 달라짐)
# ──────────────────────────────────────────────────────────────
JOSA = ("으로써", "에서는", "에게서", "으로", "에서", "에게", "까지", "부터", "보다", "처럼", "에는",
        "와의", "과의", "이란", "라는", "하기", "하는", "하고", "들의", "들은", "들이",
        "은", "는", "이", "가", "을", "를", "의", "에", "와", "과", "도", "로", "만")
STOPWORDS = {"그리고", "대한", "관한", "위한", "통한", "우리", "있는", "및", "등", "그", "이번", "주제", "관련"}


def extract_keywords(text):
    out = []
    for tok in re.findall(r"[가-힣A-Za-z0-9]+", text):
        t = tok.lower()
        for j in JOSA:
            if t.endswith(j) and len(t) - len(j) >= 2:
                t = t[: -len(j)]
                break
        if len(t) >= 2 and t not in STOPWORDS and t not in out:
            out.append(t)
    return out


def score_relevance(topic, text):
    kws = extract_keywords(topic)
    if not kws:
        return 0, [], []
    low = text.lower()
    hit = [k for k in kws if k in low]
    miss = [k for k in kws if k not in low]
    return round(len(hit) / len(kws) * 100), hit, miss


RELIABILITY_RULES = [
    ("링크(URL)", r"https?://\S+|www\.\S+", 30),
    ("출처·참고문헌 표기", r"출처|참고\s*문헌|참고\s*자료|인용|references?|source", 25),
    ("괄호 인용 (저자, 연도)", r"\([^()]{0,40}(?:19|20)\d{2}[^()]{0,12}\)", 20),
    ("각주 번호 [1]", r"\[\d{1,3}\]", 15),
    ("기관·연구 근거", r"통계청|보고서|논문|연구에\s*따르면|에\s*따르면|조사\s*결과|발표한|학회|연구소", 10),
]


def score_reliability(text):
    total, found = 0, []
    for label, pat, pts in RELIABILITY_RULES:
        cnt = len(re.findall(pat, text, flags=re.IGNORECASE))
        if cnt:
            total += pts + min(cnt - 1, 2) * 3
            found.append(f"{label} {cnt}건")
    return min(total, 100), found


def score_length(text, cfg):
    n = len(re.sub(r"\s", "", text))
    if n == 0:
        return 0, n, False
    if n > cfg["max"]:   # 결과물에 비해 너무 길면 감점
        over = (n - cfg["max"]) / cfg["max"]
        return max(50, round(100 - over * 60)), n, True
    return min(100, round(n / cfg["target"] * 100)), n, False


def evaluate(topic, text, otype):
    cfg = OUTPUT_TYPES[otype]
    rel, hit, miss = score_relevance(topic, text)
    cred_raw, found = score_reliability(text)
    cred = min(100, round(cred_raw * cfg["cred_mult"]))
    length, n_chars, too_long = score_length(text, cfg)
    w = cfg["weights"]
    total = round(rel * w[0] + cred * w[1] + length * w[2])

    tips = {
        "적합성": "주제의 핵심 단어를 본문에 직접 넣어 연결성을 보여 주세요.",
        "신뢰도": "출처 링크나 (저자, 연도) 형식의 인용을 추가해 보세요.",
        "분량": ("목표보다 길어요. 핵심만 남기고 줄여 보세요." if too_long
                 else "사례나 근거를 더 보태 분량을 채워 보세요."),
    }
    # 가중치가 큰 항목의 부족분을 우선해서 지적
    gaps = {"적합성": (100 - rel) * w[0], "신뢰도": (100 - cred) * w[1], "분량": (100 - length) * w[2]}
    weakest = max(gaps, key=gaps.get)
    if total >= 80:
        comment = f"{otype} 기준 종합 {total}점. 바로 팀에 공유해도 좋은 자료예요."
    elif total >= 55:
        comment = f"{otype} 기준 종합 {total}점. 쓸 만하지만 {weakest} 보완이 필요해요. {tips[weakest]}"
    else:
        comment = f"{otype} 기준 종합 {total}점. {weakest}부터 손보는 게 좋겠어요. {tips[weakest]}"

    return {
        "rel": rel, "cred": cred, "len": length, "total": total, "n_chars": n_chars,
        "too_long": too_long, "hit": hit, "miss": miss, "found": found, "comment": comment,
        "otype": otype,
    }


def read_upload(file):
    data = file.getvalue()
    if file.name.lower().endswith(".docx"):
        if Document is None:
            st.error("docx 파일을 읽으려면 python-docx 설치가 필요합니다: py -m pip install python-docx")
            return ""
        return "\n".join(p.text for p in Document(BytesIO(data)).paragraphs)
    for enc in ("utf-8", "cp949", "euc-kr"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="ignore")


def render_scores(ev):
    cfg = OUTPUT_TYPES[ev["otype"]]
    w = cfg["weights"]
    kw_html = "".join(f'<span class="kw hit">{esc(k)}</span>' for k in ev["hit"]) + \
              "".join(f'<span class="kw miss">{esc(k)}</span>' for k in ev["miss"])
    len_note = f"공백 제외 {ev['n_chars']:,}자 · {ev['otype']} 목표 약 {cfg['target']:,}자"
    if ev["too_long"]:
        len_note += f" (최대 {cfg['max']:,}자를 넘어 감점)"
    cred_note = ("발견: " + esc(", ".join(ev["found"]))) if ev["found"] else "출처나 인용 표시를 찾지 못했어요."
    if cfg["cred_mult"] < 1:
        cred_note += f" · {ev['otype']}는 신뢰도를 더 엄격하게 봐요."
    items = [
        ("적합성", ev["rel"], w[0], f"주제 키워드 {len(ev['hit'])}/{len(ev['hit']) + len(ev['miss'])}개 포함 &nbsp; {kw_html}"),
        ("신뢰도", ev["cred"], w[1], cred_note),
        ("분량", ev["len"], w[2], len_note),
    ]
    for name, val, weight, sub in items:
        st.markdown(
            f'<div class="score-head"><span class="score-name">{name}<small>반영 {round(weight * 100)}%</small></span>'
            f'<span class="score-num">{val}<small> / 100</small></span></div>',
            unsafe_allow_html=True,
        )
        st.progress(val / 100)
        st.markdown(f'<div class="score-note">{sub}</div>', unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# 팀 데이터 · 로그인
# ──────────────────────────────────────────────────────────────
ss = st.session_state
for _k, _v in {"user": None, "draft": None, "upload_n": 0, "flash": None, "chat": [], "ai_summary": None,
               "setup_ids": [0, 1, 2, 3], "setup_next": 4, "nav": "홈", "p2_default": "자료 올리기",
               "team_code": None, "creating": False}.items():
    ss.setdefault(_k, _v)

PAGES = ["홈", "파트 분배", "자료 내 추적", "기여도 평가", "AI 도우미"]
PW_MIN = 4


def hash_pw(pw, salt=None):
    salt = salt or uuid.uuid4().hex
    return salt, hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt.encode(), 120_000).hex()


def check_pw(member, pw):
    return bool(member and member.get("salt")) and hash_pw(pw, member["salt"])[1] == member.get("pw_hash")


def set_pw(member, pw):
    member["salt"], member["pw_hash"] = hash_pw(pw)


def normalize(d):
    """예전 버전 저장 파일도 열리도록 정리."""
    t = d.get("team")
    if t:
        for i, m in enumerate(t.get("members", [])):
            if not m.get("name"):
                m["name"] = f"팀원 {i + 1}"
            m.setdefault("trait", TRAITS[i % len(TRAITS)])
            m.setdefault("prefs", [])
        t.setdefault("topic", "")
        t.setdefault("output_type", "탐구보고서")
    return d


def get_data():
    return normalize(load_data())


def enter_team(code):
    ss.team_code, ss.user, ss.creating = code, None, False
    st.query_params["team"] = code


def leave_team():
    for k in ("draft", "ai_summary", "login_who"):
        ss.pop(k, None)
    ss.team_code, ss.user, ss.chat, ss.creating = None, None, [], False
    if "team" in st.query_params:
        del st.query_params["team"]


def invite_link(code):
    try:
        url = st.context.url or ""
    except Exception:
        url = ""
    return f"{url.split('?')[0].rstrip('/')}/?team={code}" if url.startswith("http") else None


def invite_box(code):
    st.markdown(f'<div class="mini-scores">우리 팀 코드 <b style="font-size:1.15rem;letter-spacing:0.12em">{code}</b>'
                ' · 팀원은 시작 화면에서 이 코드를 넣거나, 아래 링크로 바로 들어올 수 있어요.</div>',
                unsafe_allow_html=True)
    link = invite_link(code)
    if link:
        st.code(link, language=None)


def names_of(team):
    return [m["name"] for m in team["members"]] if team else []


def find_member(team, name):
    return next((m for m in team["members"] if m["name"] == name), None) if team else None


def default_approval(n):
    return max(0, math.ceil((n - 1) / 2))   # 작성자를 뺀 나머지의 과반


def approval_needed(team):
    n = len(team["members"])
    return max(0, min(int(team.get("approval_needed", default_approval(n))), n - 1))


def deadline_of(team):
    try:
        return date.fromisoformat(team["deadline"]) if team and team.get("deadline") else None
    except ValueError:
        return None


def dday_text(team):
    dl = deadline_of(team)
    if not dl:
        return "미정"
    left = (dl - date.today()).days
    return "D-day" if left == 0 else (f"D-{left}" if left > 0 else f"D+{-left} (지남)")


def part_of(team, name):
    for r in (team.get("assignment") or {}).get("result", []):
        if r["name"] == name:
            return r
    return None


def make_assignment(team):
    members = [{"name": m["name"], "trait": m["trait"], "prefs": m["prefs"]} for m in team["members"]]
    result, ranked = assign_roles(team["topic"], team["output_type"], members)
    asg = {"topic": team["topic"], "otype": team["output_type"], "result": result, "ranked": ranked}
    if ai_ready():
        with st.spinner("AI가 주제에 맞는 할 일을 정리하고 있어요..."):
            asg["ai"] = ai_personalize(asg)
    return asg


def apply_assignment(d, asg):
    """새 배정을 저장하고, 할 일이 바뀐 팀원의 체크는 비운다."""
    old = {f"{r['name']}|{r['task']}": r["todo"] for r in (d["team"].get("assignment") or {}).get("result", [])}
    for r in asg["result"]:
        k = f"{r['name']}|{r['task']}"
        if old.get(k) != r["todo"]:
            d["progress"].pop(k, None)
    d["team"]["assignment"] = asg
    for k in [k for k in ss if str(k).startswith("chk_")]:
        del ss[k]


def go(page, view=None):
    ss.nav = page
    if view:   # 위젯 값을 직접 바꾸지 않고, 다음에 그릴 때의 기본값으로 넘김
        ss.p2_default = view
        ss.pop("p2_view", None)


def team_context(team):
    """AI 도우미가 참고할 팀 상황."""
    d = get_data()
    lines = [f"주제: {team['topic']}", f"결과물: {team['output_type']}", f"질문한 사람: {ss.user}"]
    if deadline_of(team):
        lines.append(f"마감일: {team['deadline']} (오늘 {date.today().isoformat()})")
    for r in (team.get("assignment") or {}).get("result", []):
        lines.append(f"- {r['name']}: {r['title']} / 할 일: {', '.join(r['todo'])}")
    posted = [p for p in d["posts"] if p["status"] == "posted"]
    if posted:
        lines.append("게시된 자료: " + "; ".join(f"{p['title']}({p['author']}, {p['eval']['total']}점)" for p in posted[-10:]))
    return "\n".join(lines)


# ──────────────────────────────────────────────────────────────
# 로직 3: 게시판 (동의제) — 로그인한 본인만 올리고 투표
# ──────────────────────────────────────────────────────────────
def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def refresh_status(p, names, need):
    if p["status"] != "pending":
        return
    eligible = [n for n in names if n != p["author"]]
    agree = sum(1 for v in p["votes"].values() if v)
    disagree = sum(1 for v in p["votes"].values() if not v)
    if agree >= need:
        p["status"], p["decided_at"] = "posted", now_text()
    elif disagree > len(eligible) - need:   # 남은 사람이 모두 동의해도 기준에 못 미침
        p["status"], p["decided_at"] = "rejected", now_text()


def submit_draft():
    dr = ss.draft
    if not dr or not ss.user:
        return
    d = get_data()
    need = approval_needed(d["team"])
    post = {
        "id": uuid.uuid4().hex[:8], "author": ss.user, "title": dr["title"], "content": dr["text"],
        "topic": dr["topic"], "otype": dr["eval"]["otype"],
        "eval": {k: dr["eval"][k] for k in ("rel", "cred", "len", "total", "n_chars")},
        "created_at": now_text(), "votes": {}, "need": need, "status": "pending", "ai": dr.get("ai"),
    }
    refresh_status(post, names_of(d["team"]), need)   # 필요 동의 수가 0이면 바로 게시
    d["posts"].append(post)
    save_data(d)
    ss.draft, ss.post_title, ss.post_text = None, "", ""
    ss.upload_n += 1
    ss.flash = ("바로 게시판에 올라갔어요." if post["status"] == "posted"
                else f"게시 요청을 보냈어요. 팀원 {need}명이 동의하면 게시판에 올라가요.")


def cast_vote(post_id, agree):
    d = get_data()
    for p in d["posts"]:
        if p["id"] == post_id and p["status"] == "pending" and ss.user and ss.user != p["author"]:
            p["votes"][ss.user] = agree
            refresh_status(p, names_of(d["team"]), p.get("need", approval_needed(d["team"])))
            if p["status"] == "posted":
                ss.flash = f"'{p['title']}' 자료가 동의 기준을 넘어 게시판에 올라갔어요."
            elif p["status"] == "rejected":
                ss.flash = f"'{p['title']}' 자료는 반대가 많아 반려됐어요."
    save_data(d)


def delete_post(post_id):
    d = get_data()
    d["posts"] = [p for p in d["posts"] if not (p["id"] == post_id and p["author"] == ss.user)]
    save_data(d)


def post_card(p):
    ev = p["eval"]
    st.markdown(
        f'<div class="post-head"><span class="post-title">{esc(p["title"])}</span>'
        f'<span class="total-badge">종합 {ev["total"]}점</span></div>'
        f'<div class="post-meta">{esc(p["author"])} · {esc(p["otype"])} 기준 · {esc(p["created_at"])} · {ev["n_chars"]:,}자</div>'
        f'<div class="mini-scores">적합성 {ev["rel"]} · 신뢰도 {ev["cred"]} · 분량 {ev["len"]}</div>',
        unsafe_allow_html=True,
    )
    with st.expander("내용 보기"):
        st.markdown(f'<div class="post-body">{esc(p["content"])}</div>', unsafe_allow_html=True)
        if p.get("ai"):
            ai_box(p["ai"])


def delete_control(p):
    with st.popover("삭제"):
        st.caption("내가 올린 자료를 지울까요? 되돌릴 수 없어요.")
        st.button("네, 지울게요", key=f"del_{p['id']}", on_click=delete_post, args=(p["id"],), type="primary")


# ──────────────────────────────────────────────────────────────
# 로직 4: 파트 진행률 (본인 파트를 다 끝내면 100%)
# ──────────────────────────────────────────────────────────────
GENERIC_TODO = ["맡은 부분 자료 모으기", "초안 만들기", "팀 피드백 반영하기"]
AUTO_ITEM = "게시판에 자료 1건 이상 게시"


def team_rows(d):
    team, posted = d["team"], [p for p in d["posts"] if p["status"] == "posted"]
    rows = []
    for m in team["members"]:
        a = part_of(team, m["name"])
        todo = list(a["todo"]) if a else list(GENERIC_TODO)
        key = f"{m['name']}|{a['task'] if a else '공통'}"
        checks = (d["progress"].get(key, []) + [False] * len(todo))[:len(todo)]
        auto = any(p["author"] == m["name"] for p in posted)
        done = sum(checks) + int(auto)
        rows.append({
            "name": m["name"], "role": a["title"] if a else "파트 배정 전", "desc": a["desc"] if a else "",
            "tip": (a or {}).get("tip"), "todo": todo, "key": key, "checks": checks, "auto": auto,
            "pct": round(done / (len(todo) + 1) * 100), "left": len(todo) + 1 - done,
        })
    overall = round(sum(r["pct"] for r in rows) / len(rows)) if rows else 0
    return rows, overall


def set_check(key, idx, n_items, widget_key):
    if not ss.user or not key.startswith(ss.user + "|"):
        return   # 본인 파트만 체크 가능
    d = get_data()
    cur = (d["progress"].get(key, []) + [False] * n_items)[:n_items]
    cur[idx] = bool(ss[widget_key])
    d["progress"][key] = cur
    save_data(d)


def pending_for(d, name):
    return [p for p in d["posts"] if p["status"] == "pending" and p["author"] != name and name not in p["votes"]]


# ──────────────────────────────────────────────────────────────
# 사이드바 (상태 요약은 화면을 다 그린 뒤 마지막에 채움)
# ──────────────────────────────────────────────────────────────
_qp = (st.query_params.get("team") or "").strip().upper()
if _qp and not ss.team_code:
    ss.team_code = _qp   # 초대 링크(?team=코드)로 들어온 경우
storage_error = None
try:
    if ss.team_code and not ss.creating and not team_exists(ss.team_code):
        ss.flash = f"팀 코드 {ss.team_code}인 팀을 찾지 못했어요. 코드를 다시 확인해 주세요."
        leave_team()
    if ss.team_code and st.query_params.get("team") != ss.team_code:
        st.query_params["team"] = ss.team_code   # 새로고침해도 같은 팀에 머물게
    data = get_data()
except StorageError as e:
    storage_error, data = str(e), copy.deepcopy(EMPTY_DATA)
team = data["team"]
if ss.user and ss.user not in names_of(team):
    ss.user = None   # 팀이 삭제됐거나 이름이 사라진 경우

with st.sidebar:
    st.markdown('<div class="brand">🧩 팀플 메이트</div><div class="brand-sub">팀 프로젝트 도우미</div>',
                unsafe_allow_html=True)
    page = None
    if ss.get("admin_mode"):
        st.markdown('<div class="me-line">🔧 <b>관리자 화면</b></div>', unsafe_allow_html=True)
    elif team:
        st.markdown(f'<div class="me-line">팀 코드 <b>{ss.team_code}</b></div>', unsafe_allow_html=True)
    if team and ss.user and not ss.get("admin_mode"):
        mp = part_of(team, ss.user)
        st.markdown(f'<div class="me-line">👤 <b>{esc(ss.user)}</b>님 · {esc(mp["title"] if mp else "파트 배정 전")}</div>',
                    unsafe_allow_html=True)
        if ss.nav not in PAGES:
            ss.nav = "홈"
        page = st.radio("메뉴", PAGES, key="nav", label_visibility="collapsed")
    status_box = st.empty()

    if ss.get("admin_mode"):
        if st.button("관리자 화면 나가기", width="stretch"):
            ss.admin_mode = False
            st.rerun()
    elif team and ss.user:
        c1, c2 = st.columns(2)
        if c1.button("로그아웃", width="stretch"):
            ss.user, ss.draft, ss.chat, ss.ai_summary = None, None, [], None
            st.rerun()
        if c2.button("다른 팀", width="stretch", help="다른 팀 코드로 들어가요."):
            leave_team()
            st.rerun()
    elif team:
        if st.button("다른 팀으로 가기", width="stretch"):
            leave_team()
            st.rerun()

    st.markdown(f'<div class="credit">{esc(CREDIT)}</div>', unsafe_allow_html=True)
    if not ss.get("admin_mode") and st.button("관리자", key="admin_open", type="tertiary"):
        ss.admin_mode = True
        st.rerun()


def show_flash():
    if ss.flash:
        st.success(ss.flash)
        ss.flash = None


def role_cards(team, me=None):
    result = (team.get("assignment") or {}).get("result", [])
    if not result:
        return
    n_cols = min(3, len(result))
    for start in range(0, len(result), n_cols):
        row = st.columns(n_cols, gap="medium")
        for col, r in zip(row, result[start:start + n_cols]):
            color = TRAIT_COLOR.get(r["trait"], "#3F44D9")
            chips = "".join(f'<span class="why-chip">{esc(x)}</span>' for x in r["reasons"])
            todo = "".join(f"<li>{esc(t)}</li>" for t in r["todo"])
            tip = f'<div class="ai-tip"><b>AI 한마디</b> {esc(r["tip"])}</div>' if r.get("tip") else ""
            me_tag = ' <span class="me-tag">나</span>' if r["name"] == me else ""
            col.markdown(
                f'<div class="role-card{" mine" if r["name"] == me else ""}" style="border-top-color:{color}">'
                f'<div class="who"><span class="name">{esc(r["name"])}{me_tag}</span>'
                f'<span class="pill" style="background:{color}">{esc(r["trait"])}</span></div>'
                f'<div class="role">{esc(r["title"])}</div><div class="desc">{esc(r["desc"])}</div>'
                f'<div class="why">{chips}</div><ul>{todo}</ul>{tip}</div>',
                unsafe_allow_html=True,
            )


# ──────────────────────────────────────────────────────────────
# 시작 1: 팀이 없을 때 — 팀 만들기
# ──────────────────────────────────────────────────────────────
def page_setup():
    if st.button("← 처음 화면으로"):
        ss.creating = False
        st.rerun()
    st.markdown('<p class="hero-title">팀플 메이트</p>'
                '<p class="hero-lead">역할 나누기부터 자료 모으기, 진행 확인까지 한곳에서.</p>',
                unsafe_allow_html=True)
    note("처음이시네요. 팀원이 모두 모인 자리에서 팀을 만드는 게 좋아요. "
         "<b>각자 자기 비밀번호를 직접 입력</b>하세요. 다음에 들어올 때 이름과 비밀번호로 로그인해요.")

    c1, c2 = st.columns([3, 1.3], gap="medium")
    c1.text_input("프로젝트 주제", key="setup_topic", placeholder="예: 우리 동네 플라스틱 쓰레기 줄이기 캠페인")
    c2.selectbox("만들려는 결과물", list(OUTPUT_TYPES), key="setup_otype")
    st.caption(f"{ss.setup_otype}: {OUTPUT_TYPES[ss.setup_otype]['focus']} 역할 배정과 자료 평가 기준에 반영돼요.")

    st.markdown("#### 팀원")
    ids = ss.setup_ids
    for start in range(0, len(ids), 2):
        row = st.columns(2, gap="medium")
        for col, (idx, sid) in zip(row, list(enumerate(ids))[start:start + 2]):
            ss.setdefault(f"s_trait_{sid}", TRAITS[sid % len(TRAITS)])
            with col, st.container(border=True):
                st.markdown(f'<span class="member-label">팀원 {idx + 1}</span>', unsafe_allow_html=True)
                a, b = st.columns([3, 2])
                a.text_input("이름", key=f"s_name_{sid}", placeholder="이름")
                b.selectbox("성향", TRAITS, key=f"s_trait_{sid}")
                st.multiselect("선호 작업", TASKS, key=f"s_pref_{sid}", placeholder="하나 이상 고르세요")
                st.text_input("비밀번호 (본인만 입력)", key=f"s_pw_{sid}", type="password",
                              placeholder=f"{PW_MIN}자 이상")

    b1, b2, _ = st.columns([1, 1, 4])
    if b1.button("＋ 팀원 추가", width="stretch"):
        ss.setup_ids.append(ss.setup_next)
        ss.setup_next += 1
        st.rerun()
    if b2.button("－ 팀원 삭제", width="stretch", disabled=len(ids) <= 1):
        ss.setup_ids.pop()
        st.rerun()

    n = len(ids)
    ss.setdefault("setup_need", default_approval(n))
    ss.setup_need = min(ss.setup_need, max(0, n - 1))
    with st.expander("추가 설정 (선택 · 나중에 '팀 설정 바꾸기'에서도 바꿀 수 있어요)"):
        st.date_input("마감일", key="setup_deadline", value=None, min_value=date(2020, 1, 1), format="YYYY-MM-DD")
        st.number_input("자료 게시에 필요한 동의 수 (올린 사람 제외)", min_value=0, max_value=max(0, n - 1), step=1,
                        key="setup_need", help="팀원이 자료를 올리면, 다른 팀원 중 이 수 이상이 동의해야 게시판에 올라가요. "
                                               f"기본값은 과반({default_approval(n)}명)이에요.")

    st.write("")
    if st.button("팀 만들고 파트 나누기", type="primary", width="stretch"):
        names = [ss.get(f"s_name_{i}", "").strip() for i in ids]
        pws = [ss.get(f"s_pw_{i}", "") for i in ids]
        errors = []
        if not ss.setup_topic.strip():
            errors.append("프로젝트 주제를 입력해 주세요.")
        if any(not x for x in names):
            errors.append("모든 팀원의 이름을 입력해 주세요.")
        if len(set(names)) != len(names):
            errors.append("이름이 같은 팀원이 있어요. 구분되게 바꿔 주세요 (예: 민지A, 민지B).")
        short = [nm or f"팀원 {k + 1}" for k, (nm, pw) in enumerate(zip(names, pws)) if len(pw) < PW_MIN]
        if short:
            errors.append(f"비밀번호를 {PW_MIN}자 이상으로 정해 주세요: {', '.join(short)}")
        if errors:
            for e in errors:
                st.error(e)
            return
        members = []
        for i, nm, pw in zip(ids, names, pws):
            m = {"name": nm, "trait": ss[f"s_trait_{i}"], "prefs": ss.get(f"s_pref_{i}", [])}
            set_pw(m, pw)
            members.append(m)
        t = {"topic": ss.setup_topic.strip(), "output_type": ss.setup_otype, "members": members,
             "deadline": ss.setup_deadline.isoformat() if ss.get("setup_deadline") else None,
             "approval_needed": int(ss.setup_need), "created_at": now_text()}
        t["assignment"] = make_assignment(t)
        try:
            code = new_team_code()
        except StorageError as e:
            st.error(str(e))
            return
        enter_team(code)
        save_data({"team": t, "posts": [], "progress": {}})
        ss.flash = (f"팀을 만들었어요! 우리 팀 코드는 {code}예요. 팀원들에게 코드나 초대 링크를 알려 주고, "
                    "각자 이름과 비밀번호로 들어가세요.")
        st.rerun()


# ──────────────────────────────────────────────────────────────
# 시작 0: 팀 코드 넣기 / 새 팀 만들기
# ──────────────────────────────────────────────────────────────
def page_start():
    st.markdown('<p class="hero-title">팀플 메이트</p>'
                '<p class="hero-lead">역할 나누기부터 자료 모으기, 진행 확인까지 한곳에서.</p>',
                unsafe_allow_html=True)
    show_flash()
    lc = legacy_code()
    if lc:
        note(f"업데이트 전에 만든 팀이 있어요. 그 팀의 코드는 <b>{lc}</b>예요.", warn=True)
        st.button(f"{lc} 팀으로 들어가기", on_click=enter_team, args=(lc,))

    left, right = st.columns(2, gap="large")
    with left, st.container(border=True):
        st.markdown("#### 팀 코드로 들어가기")
        st.caption("팀을 만든 친구에게 받은 6자리 코드를 넣어요.")
        with st.form("code_form", border=False):
            code = st.text_input("팀 코드", max_chars=6, placeholder="예: K7Q3MP")
            if st.form_submit_button("들어가기", type="primary", width="stretch"):
                code = code.strip().upper()
                try:
                    found = team_exists(code)
                except StorageError as e:
                    st.error(str(e))
                    found = None
                if found:
                    enter_team(code)
                    st.rerun()
                elif found is False:
                    st.error("그 코드의 팀을 찾지 못했어요. 영문 대문자와 숫자 6자리를 다시 확인해 주세요.")
    with right, st.container(border=True):
        st.markdown("#### 새 팀 만들기")
        st.caption("주제와 팀원을 정하면 역할을 나눠 주고, 우리 팀 코드를 만들어 줘요.")
        st.write("")
        if st.button("새 팀 만들기", width="stretch"):
            ss.creating = True
            st.rerun()
    st.caption("팀원에게 받은 초대 링크로 들어오면 코드를 넣지 않아도 바로 우리 팀으로 들어가져요.")


# ──────────────────────────────────────────────────────────────
# 시작 2: 팀이 있을 때 — 로그인
# ──────────────────────────────────────────────────────────────
def page_login(d):
    t = d["team"]
    rows, overall = team_rows(d)
    posted = sum(p["status"] == "posted" for p in d["posts"])
    st.markdown(f'<p class="hero-title">팀플 메이트</p><p class="hero-lead">{esc(t["topic"])} · {esc(t["output_type"])}</p>',
                unsafe_allow_html=True)
    st.markdown(
        '<div class="stat-row">'
        f'<div class="stat"><div class="stat-num">{dday_text(t)}</div><div class="stat-label">마감</div></div>'
        f'<div class="stat"><div class="stat-num">{overall}%</div><div class="stat-label">팀 전체 진행률</div></div>'
        f'<div class="stat"><div class="stat-num">{posted}건</div><div class="stat-label">게시된 자료</div></div>'
        f'<div class="stat"><div class="stat-num">{len(t["members"])}명</div><div class="stat-label">팀원</div></div>'
        '</div>', unsafe_allow_html=True)
    show_flash()

    left, right = st.columns([1.2, 1], gap="large")
    with left, st.container(border=True):
        st.markdown(f"#### 들어가기 · 팀 코드 {ss.team_code}")
        who = st.selectbox("나는 누구인가요?", names_of(t), key="login_who")
        m = find_member(t, who)
        if m and m.get("pw_hash"):
            with st.form("login_form", border=False):
                pw = st.text_input("비밀번호", type="password")
                if st.form_submit_button("들어가기", type="primary", width="stretch"):
                    if check_pw(m, pw):
                        ss.user, ss.nav = who, "홈"
                        st.rerun()
                    st.error("비밀번호가 맞지 않아요.")
        else:   # 예전 버전 팀이거나 비밀번호가 아직 없는 팀원
            st.caption("아직 비밀번호가 없어요. 지금 만들면 다음부터 이 비밀번호로 들어와요.")
            with st.form("first_pw_form", border=False):
                p1 = st.text_input("새 비밀번호", type="password", placeholder=f"{PW_MIN}자 이상")
                p2 = st.text_input("한 번 더", type="password")
                if st.form_submit_button("비밀번호 만들고 들어가기", type="primary", width="stretch"):
                    if len(p1) < PW_MIN:
                        st.error(f"{PW_MIN}자 이상으로 정해 주세요.")
                    elif p1 != p2:
                        st.error("두 비밀번호가 달라요.")
                    else:
                        d2 = get_data()
                        set_pw(find_member(d2["team"], who), p1)
                        save_data(d2)
                        ss.user, ss.nav = who, "홈"
                        st.rerun()
        with st.expander("비밀번호를 잊었어요"):
            st.caption("다른 팀원이 로그인해서 '파트 분배 → 팀 설정 바꾸기 → 팀원 관리'에서 새 비밀번호를 정해 줄 수 있어요. "
                       "바뀐 사실은 다음에 들어왔을 때 본인 화면에 표시돼요.")

        st.button("다른 팀으로 가기", on_click=leave_team)

    with right:
        invite_box(ss.team_code)
        st.markdown("#### 우리 팀")
        for r in rows:
            st.markdown(f'<div class="team-line"><b>{esc(r["name"])}</b> · {esc(r["role"])}'
                        f'<span class="team-pct">{r["pct"]}%</span></div>', unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# 화면 0: 홈 — 내 할 일과 나를 기다리는 투표
# ──────────────────────────────────────────────────────────────
def clear_notice():
    d = get_data()
    m = find_member(d["team"], ss.user)
    if m:
        m.pop("pw_notice", None)
        save_data(d)


def page_home(d):
    t = d["team"]
    rows, overall = team_rows(d)
    page_header(f"{esc(ss.user)}님, 안녕하세요", f"마감 {dday_text(t)} · 팀 전체 진행률 {overall}% · 주제: {esc(t['topic'])}")
    show_flash()

    if legacy_code() == ss.team_code:
        note(f"이 팀은 업데이트 전에 만든 팀이라 시작 화면에 팀 코드({ss.team_code})가 안내되고 있어요. "
             "팀원 모두 코드를 알게 되면 안내를 숨겨 주세요.", warn=True)
        st.button("시작 화면 안내 숨기기", on_click=hide_legacy_notice)

    m = find_member(t, ss.user)
    if m.get("pw_notice"):
        note(esc(m["pw_notice"]) + " 직접 바꾼 게 아니라면 '파트 분배 → 팀 설정 바꾸기 → 내 정보'에서 새로 정하세요.", warn=True)
        st.button("확인했어요", on_click=clear_notice)

    left, right = st.columns([1.6, 1], gap="medium")
    with left:
        my_checklist()

    with right:
        waiting = pending_for(d, ss.user)
        with st.container(border=True):
            st.markdown(f'<div class="big-count">{len(waiting)}건</div><div class="stat-label">내 투표를 기다리는 자료</div>',
                        unsafe_allow_html=True)
            st.button("투표하러 가기", width="stretch", type="primary" if waiting else "secondary",
                      on_click=go, args=("자료 내 추적", "투표하기"), disabled=not waiting)
        mine = [p for p in d["posts"] if p["author"] == ss.user]
        with st.container(border=True):
            st.markdown("**내 자료**")
            st.markdown(
                f'<div class="mini-scores">게시 {sum(p["status"] == "posted" for p in mine)}건 · '
                f'동의 대기 {sum(p["status"] == "pending" for p in mine)}건 · '
                f'반려 {sum(p["status"] == "rejected" for p in mine)}건</div>', unsafe_allow_html=True)
            st.button("자료 올리기", width="stretch", on_click=go, args=("자료 내 추적", "자료 올리기"))


@st.fragment
def my_checklist():
    """체크할 때 이 상자만 다시 그려서 빠름."""
    d = get_data()
    t = d["team"]
    rows, _ = team_rows(d)
    me = next(r for r in rows if r["name"] == ss.user)
    with st.container(border=True):
        st.markdown(f'<div class="home-role">{esc(me["role"])}</div><div class="desc">{esc(me["desc"])}</div>',
                    unsafe_allow_html=True)
        if me.get("tip"):
            st.markdown(f'<div class="ai-tip"><b>AI 한마디</b> {esc(me["tip"])}</div>', unsafe_allow_html=True)
        st.write("")
        st.progress(me["pct"] / 100, text=f"내 파트 진행률 {me['pct']}%")
        st.markdown("**내 할 일**")
        for i, item in enumerate(me["todo"]):
            wkey = f"chk_{me['key']}_{i}"
            st.checkbox(item, value=me["checks"][i], key=wkey,
                        on_change=set_check, args=(me["key"], i, len(me["todo"]), wkey))
        st.checkbox(f"{AUTO_ITEM} (게시되면 자동 체크)", value=me["auto"], disabled=True, key="auto_me")
        if not part_of(t, ss.user):
            st.caption("아직 파트가 없어 공통 할 일이 보여요. 파트 분배 화면에서 '파트 다시 나누기'를 눌러 주세요.")


# ──────────────────────────────────────────────────────────────
# 화면 1: 파트 분배 (팀 설정 포함)
# ──────────────────────────────────────────────────────────────
def page_parts(d):
    t = d["team"]
    page_header("파트 분배", "누가 어떤 역할을 맡았는지 확인하고, 필요하면 팀 설정을 바꿔요.")
    show_flash()
    show_ai_error()
    asg = t.get("assignment") or {}
    result = asg.get("result", [])
    matched = sum(r["pref_matched"] for r in result)
    note(f"<b>{esc(t['topic'])}</b> ({esc(t['output_type'])}) · 마감 {dday_text(t)} · "
         f"{len(result)}명 중 {matched}명이 원하는 작업을 맡았어요. "
         f"자료 게시에는 올린 사람을 뺀 {approval_needed(t)}명의 동의가 필요해요.")
    unassigned = [nm for nm in names_of(t) if not part_of(t, nm)]
    if unassigned:
        note(f"아직 파트가 없는 팀원: <b>{esc(', '.join(unassigned))}</b>. 아래 '파트 다시 나누기'를 눌러 주세요.", warn=True)
    role_cards(t, me=ss.user)

    st.write("")
    with st.container(border=True):
        st.markdown("**팀원 초대**")
        invite_box(ss.team_code)
    with st.expander("팀 설정 바꾸기"):
        tab_p, tab_me, tab_team = st.tabs(["프로젝트", "내 정보", "팀원 관리"])

        with tab_p, st.form("proj_form"):
            topic = st.text_input("프로젝트 주제", value=t["topic"])
            otype = st.selectbox("만들려는 결과물", list(OUTPUT_TYPES), index=list(OUTPUT_TYPES).index(t["output_type"]))
            dl = st.date_input("마감일", value=deadline_of(t), min_value=date(2020, 1, 1), format="YYYY-MM-DD")
            n = len(t["members"])
            need = st.number_input("자료 게시에 필요한 동의 수 (올린 사람 제외)", min_value=0, max_value=max(0, n - 1),
                                   value=approval_needed(t), step=1)
            redo = st.checkbox("저장하면서 파트도 다시 나누기 (주제·결과물을 바꿨다면 추천)")
            if st.form_submit_button("저장", type="primary"):
                if not topic.strip():
                    st.error("주제를 비울 수 없어요.")
                else:
                    d2 = get_data()
                    d2["team"].update({"topic": topic.strip(), "output_type": otype,
                                       "deadline": dl.isoformat() if dl else None, "approval_needed": int(need)})
                    if redo:
                        apply_assignment(d2, make_assignment(d2["team"]))
                    save_data(d2)
                    ss.flash = "프로젝트 정보를 저장했어요."
                    st.rerun()

        with tab_me:
            m = find_member(t, ss.user)
            with st.form("me_form"):
                trait = st.selectbox("내 성향", TRAITS, index=TRAITS.index(m["trait"]) if m["trait"] in TRAITS else 0)
                prefs = st.multiselect("내 선호 작업", TASKS, default=[p for p in m["prefs"] if p in TASKS])
                st.caption("성향·선호는 '파트 다시 나누기'를 할 때 반영돼요.")
                old = st.text_input("현재 비밀번호 (비밀번호를 바꿀 때만)", type="password")
                new1 = st.text_input("새 비밀번호", type="password", placeholder=f"{PW_MIN}자 이상")
                new2 = st.text_input("새 비밀번호 한 번 더", type="password")
                if st.form_submit_button("저장", type="primary"):
                    d2 = get_data()
                    m2 = find_member(d2["team"], ss.user)
                    m2["trait"], m2["prefs"] = trait, prefs
                    msg = "내 정보를 저장했어요."
                    if new1 or new2:
                        if not check_pw(m2, old):
                            st.error("현재 비밀번호가 맞지 않아요.")
                            st.stop()
                        if len(new1) < PW_MIN or new1 != new2:
                            st.error(f"새 비밀번호는 {PW_MIN}자 이상이고 두 번 똑같이 입력해야 해요.")
                            st.stop()
                        set_pw(m2, new1)
                        m2.pop("pw_notice", None)
                        msg = "내 정보와 비밀번호를 바꿨어요."
                    save_data(d2)
                    ss.flash = msg
                    st.rerun()

        with tab_team:
            if st.button("파트 다시 나누기", type="primary"):
                d2 = get_data()
                apply_assignment(d2, make_assignment(d2["team"]))
                save_data(d2)
                ss.flash = "파트를 다시 나눴어요. 할 일이 바뀐 팀원은 체크가 초기화됐어요."
                st.rerun()
            st.caption("지금 팀원 정보로 역할을 새로 정해요. 할 일이 바뀐 팀원은 진행 체크가 비워져요.")
            st.divider()

            st.markdown("**팀원 추가**")
            with st.form("add_form", clear_on_submit=True):
                a, b = st.columns([3, 2])
                nm = a.text_input("이름")
                tr = b.selectbox("성향", TRAITS)
                pf = st.multiselect("선호 작업", TASKS)
                pw = st.text_input("비밀번호 (새 팀원이 직접 입력)", type="password", placeholder=f"{PW_MIN}자 이상")
                if st.form_submit_button("추가"):
                    if not nm.strip() or nm.strip() in names_of(t):
                        st.error("이름을 입력하고, 기존 팀원과 다른 이름으로 정해 주세요.")
                    elif len(pw) < PW_MIN:
                        st.error(f"비밀번호를 {PW_MIN}자 이상으로 정해 주세요.")
                    else:
                        d2 = get_data()
                        nmem = {"name": nm.strip(), "trait": tr, "prefs": pf}
                        set_pw(nmem, pw)
                        d2["team"]["members"].append(nmem)
                        save_data(d2)
                        ss.flash = f"{nm.strip()}님을 추가했어요. '파트 다시 나누기'로 역할을 정해 주세요."
                        st.rerun()

            others = [x for x in names_of(t) if x != ss.user]
            if others:
                st.markdown("**팀원 비밀번호 다시 정하기** (잊어버린 팀원을 도울 때)")
                with st.form("reset_other_form", clear_on_submit=True):
                    who = st.selectbox("누구의 비밀번호인가요?", others)
                    npw = st.text_input("새 비밀번호", type="password", placeholder=f"{PW_MIN}자 이상")
                    if st.form_submit_button("다시 정하기"):
                        if len(npw) < PW_MIN:
                            st.error(f"{PW_MIN}자 이상으로 정해 주세요.")
                        else:
                            d2 = get_data()
                            m2 = find_member(d2["team"], who)
                            set_pw(m2, npw)
                            m2["pw_notice"] = f"{ss.user}님이 {now_text()}에 내 비밀번호를 다시 정했어요."
                            save_data(d2)
                            ss.flash = f"{who}님의 비밀번호를 바꿨어요. 본인에게 직접 알려 주세요."
                            st.rerun()

                st.markdown("**팀원 빼기**")
                with st.form("remove_form"):
                    who = st.selectbox("뺄 팀원", others, key="remove_who")
                    sure = st.checkbox("이 팀원을 팀에서 뺄게요 (올린 자료는 게시판에 남아요)")
                    if st.form_submit_button("빼기") and sure:
                        d2 = get_data()
                        d2["team"]["members"] = [x for x in d2["team"]["members"] if x["name"] != who]
                        asg2 = d2["team"].get("assignment")
                        if asg2:
                            asg2["result"] = [r for r in asg2["result"] if r["name"] != who]
                        d2["team"]["approval_needed"] = approval_needed(d2["team"])
                        save_data(d2)
                        ss.flash = f"{who}님을 팀에서 뺐어요."
                        st.rerun()


# ──────────────────────────────────────────────────────────────
# 화면 2: 자료 내 추적 (올리기 · 투표하기 · 게시판)
# ──────────────────────────────────────────────────────────────
def page_materials(d):
    page_header("자료 내 추적", "자료를 점검해서 올리고, 팀원 자료에 투표하고, 게시된 자료를 모아 봐요.")
    materials_body()


@st.fragment
def materials_body():
    """이 부분 안에서 누르는 버튼은 이 부분만 다시 그려서 빠름."""
    show_flash()
    d = get_data()
    t, posts = d["team"], d["posts"]
    waiting = pending_for(d, ss.user)
    top_l, top_r = st.columns([4, 1])
    with top_l:
        view = st.segmented_control("보기", ["자료 올리기", "투표하기", "게시판"], key="p2_view",
                                    default=ss.p2_default, label_visibility="collapsed") or ss.p2_default
        ss.p2_default = view
    if top_r.button("새로고침", width="stretch", help="다른 팀원이 올리거나 투표한 내용을 불러와요."):
        st.rerun(scope="fragment")
    st.caption(f"내 투표를 기다리는 자료 {len(waiting)}건 · 게시 기준: 올린 사람을 뺀 팀원 중 {approval_needed(t)}명 이상 동의")

    if view == "자료 올리기":
        otype, cfg = t["output_type"], OUTPUT_TYPES[t["output_type"]]
        w = cfg["weights"]
        note(f"<b>{esc(ss.user)}</b>님 이름으로 올라가요. <b>{esc(otype)}</b> 기준으로 평가해요. "
             f"적합성 {round(w[0] * 100)}% · 신뢰도 {round(w[1] * 100)}% · 분량 {round(w[2] * 100)}% 반영, "
             f"목표 분량 약 {cfg['target']:,}자. {esc(cfg['focus'])}")
        st.text_input("자료 제목", key="post_title", placeholder="예: 플라스틱 배출량 통계 정리")
        left, right = st.columns([3, 2], gap="large")
        with left:
            st.text_area("자료 붙여넣기", height=240, key="post_text", placeholder="조사한 내용을 여기에 붙여넣으세요.")
        with right:
            types = ["txt", "docx"] if Document is not None else ["txt"]
            up = st.file_uploader("또는 파일 올리기 (.txt, .docx)", type=types, key=f"upload_{ss.upload_n}")
            st.caption("붙여넣은 글과 파일을 함께 넣으면 둘을 합쳐서 평가해요.")

        if st.button("평가하기", type="primary", width="stretch"):
            text = ss.get("post_text", "") or ""
            if up is not None:
                text = (text + "\n" + read_upload(up)).strip()
            if not text.strip():
                st.warning("평가할 자료를 붙여넣거나 파일을 올려 주세요.")
                ss.draft = None
            else:
                title = ss.get("post_title", "").strip() or \
                        (up.name if up is not None else text.strip().splitlines()[0])[:40]
                ev = evaluate(t["topic"], text, otype)
                ai_text = None
                if ai_ready():
                    with st.spinner("AI가 자료를 읽고 있어요..."):
                        ai_text = ai_review(t["topic"], otype, text, ev)
                ss.draft = {"title": title, "text": text, "topic": t["topic"], "eval": ev, "ai": ai_text}

        dr = ss.draft
        if dr:
            st.divider()
            note(esc(dr["eval"]["comment"]))
            show_ai_error()
            if dr.get("ai"):
                ai_box(dr["ai"])
            elif not ai_ready():
                st.caption("AI가 연결되면 자료 내용까지 읽고 피드백해 줘요. (연결은 관리자가 해요)")
            with st.container(border=True):
                render_scores(dr["eval"])
            need = approval_needed(t)
            st.caption(f"'{dr['title']}' · " + (f"다른 팀원 {need}명이 동의하면 게시판에 올라가요." if need
                                                else "동의 없이 바로 게시돼요."))
            st.button("게시판에 올리기 요청", type="primary", width="stretch", on_click=submit_draft)

    elif view == "투표하기":
        pending = sorted([p for p in posts if p["status"] == "pending"], key=lambda x: x["created_at"], reverse=True)
        if not pending:
            st.info("동의를 기다리는 자료가 없어요.")
        for p in pending:
            need = p.get("need", approval_needed(t))
            with st.container(border=True):
                post_card(p)
                yes = [v for v, ok in p["votes"].items() if ok]
                no = [v for v, ok in p["votes"].items() if not ok]
                st.progress(min(1.0, len(yes) / need) if need else 1.0, text=f"동의 {len(yes)} / {need}명")
                st.markdown(
                    f'<div class="vote-line"><span class="vote-yes">동의</span> {esc(", ".join(yes) or "없음")} &nbsp; '
                    f'<span class="vote-no">반대</span> {esc(", ".join(no) or "없음")}</div>', unsafe_allow_html=True)
                if p["author"] == ss.user:
                    c1, c2 = st.columns([4, 1])
                    c1.caption("내가 올린 자료라 투표할 수 없어요. 팀원들의 투표를 기다리는 중이에요.")
                    with c2:
                        delete_control(p)
                else:
                    mine = p["votes"].get(ss.user)
                    c1, c2, c3 = st.columns([2, 1, 1])
                    c1.caption("아직 투표 전이에요." if mine is None else
                               f"내 투표: {'동의' if mine else '반대'} · 마감 전까지 바꿀 수 있어요.")
                    c2.button("✓ 동의함" if mine is True else "동의", key=f"yes_{p['id']}", width="stretch",
                              type="primary" if mine is True else "secondary", on_click=cast_vote, args=(p["id"], True))
                    c3.button("✓ 반대함" if mine is False else "반대", key=f"no_{p['id']}", width="stretch",
                              type="primary" if mine is False else "secondary", on_click=cast_vote, args=(p["id"], False))

    else:   # 게시판
        posted = [p for p in posts if p["status"] == "posted"]
        rejected = [p for p in posts if p["status"] == "rejected"]
        authors = names_of(t) + [a for a in dict.fromkeys(p["author"] for p in posted) if a not in names_of(t)]
        options = ["전체"] + authors
        if "board_filter" in ss and ss.board_filter not in options:
            del ss.board_filter
        f1, f2 = st.columns([4, 1.4])
        with f1:
            who = st.segmented_control("팀원", options, key="board_filter", default="전체",
                                       label_visibility="collapsed") or "전체"
        if posted:
            bundle = "\n\n".join(f"■ {p['title']} ({p['author']}, {p['created_at']})\n{p['content']}" for p in posted)
            f2.download_button("전체 받기 (.txt)", bundle.encode("utf-8"), file_name="팀플_게시자료.txt",
                               mime="text/plain", width="stretch")
        shown = sorted([p for p in posted if who == "전체" or p["author"] == who],
                       key=lambda x: x.get("decided_at", x["created_at"]), reverse=True)
        if not shown:
            st.info("게시된 자료가 아직 없어요." if who == "전체" else f"{who}님의 게시된 자료가 아직 없어요.")
        for p in shown:
            with st.container(border=True):
                post_card(p)
                agreed = [v for v, ok in p["votes"].items() if ok]
                c1, c2 = st.columns([4, 1])
                c1.caption(f"{p.get('decided_at', '')} 게시 · 동의: {', '.join(agreed) or '기준 0명으로 바로 게시'}")
                if p["author"] == ss.user:
                    with c2:
                        delete_control(p)
        if rejected:
            with st.expander(f"반려된 자료 {len(rejected)}건"):
                for p in rejected:
                    with st.container(border=True):
                        post_card(p)
                        no = [v for v, ok in p["votes"].items() if not ok]
                        c1, c2 = st.columns([4, 1])
                        c1.caption(f"반대: {', '.join(no)} · 내용을 보완해서 다시 올려 보세요.")
                        if p["author"] == ss.user:
                            with c2:
                                delete_control(p)


# ──────────────────────────────────────────────────────────────
# 화면 3: 기여도 평가 (보기 전용 — 체크는 홈에서 본인 것만)
# ──────────────────────────────────────────────────────────────
def light_chart(chart):
    """브라우저가 다크 모드여도 그래프는 흰 배경·진한 글씨로."""
    return (chart.configure(background="#FFFFFF", font="Pretendard Variable, Pretendard, Malgun Gothic, sans-serif")
            .configure_view(stroke=None)
            .configure_axis(labelColor="#1A1F36", titleColor="#5A6178", gridColor="#ECEFF5",
                            domainColor="#D5DAE5", tickColor="#D5DAE5")
            .configure_legend(labelColor="#1A1F36", titleColor="#1A1F36"))


def page_contrib(d):
    t = d["team"]
    rows, overall = team_rows(d)
    posted = [p for p in d["posts"] if p["status"] == "posted"]
    page_header("기여도 평가", "각자 맡은 파트를 다 끝내면 100%예요. 내 할 일 체크는 홈에서 해요.")
    done_parts = sum(r["pct"] == 100 for r in rows)
    left_items = sum(r["left"] for r in rows)

    chart_col, total_col = st.columns([2.6, 1], gap="medium")
    with chart_col, st.container(border=True):
        st.markdown("**팀원별 파트 진행률**")
        df = pd.DataFrame({"팀원": [r["name"] for r in rows], "진행률": [r["pct"] for r in rows],
                           "역할": [r["role"] for r in rows]})
        df["상태"] = df["진행률"].apply(lambda v: "완료" if v == 100 else ("확인 필요" if v < 40 else "진행 중"))
        base = alt.Chart(df).encode(
            y=alt.Y("팀원:N", sort="-x", title=None, axis=alt.Axis(labelFontSize=14, labelFontWeight=600)),
            x=alt.X("진행률:Q", title="본인 파트 진행률 (%)", scale=alt.Scale(domain=[0, 112]),
                    axis=alt.Axis(values=[0, 25, 50, 75, 100])),
        )
        bars = base.mark_bar(cornerRadiusEnd=6, height=24).encode(
            color=alt.Color("상태:N", scale=alt.Scale(domain=["완료", "진행 중", "확인 필요"],
                                                     range=["#12A386", "#3F44D9", "#C9822A"]),
                            legend=alt.Legend(title=None, orient="bottom")),
            tooltip=["팀원", "역할", alt.Tooltip("진행률:Q", title="진행률(%)")],
        )
        labels = base.mark_text(align="left", dx=6, fontSize=13, fontWeight=700, color="#1A1F36").encode(
            text=alt.Text("진행률:Q", format=".0f"))
        avg_rule = alt.Chart(pd.DataFrame({"x": [overall]})).mark_rule(strokeDash=[4, 4], color="#8A90A6").encode(x="x:Q")
        st.altair_chart(light_chart((bars + labels + avg_rule).properties(height=max(190, 52 * len(rows)))),
                        width="stretch", theme=None)
        st.caption(f"점선은 팀 평균({overall}%)이에요.")

    with total_col, st.container(border=True):
        st.markdown("**프로젝트 전체 진행률**")
        ring = pd.DataFrame({"구분": ["완료", "남음"], "값": [overall, 100 - overall]})
        arc = alt.Chart(ring).mark_arc(innerRadius=58, outerRadius=82).encode(
            theta=alt.Theta("값:Q", stack=True),
            color=alt.Color("구분:N", scale=alt.Scale(domain=["완료", "남음"], range=["#3F44D9", "#E3E7F0"]), legend=None),
            order=alt.Order("구분:N", sort="ascending"),
        )
        center = alt.Chart(pd.DataFrame({"t": [f"{overall}%"]})).mark_text(
            fontSize=30, fontWeight=800, color="#1A1F36").encode(text="t:N")
        st.altair_chart(light_chart((arc + center).properties(height=200)), width="stretch", theme=None)
        st.markdown(
            f'<div class="mini-scores" style="text-align:center;line-height:1.9">'
            f'완료한 파트 <b>{done_parts} / {len(rows)}</b><br>남은 할 일 <b>{left_items}개</b><br>'
            f'게시된 자료 <b>{len(posted)}건</b><br>마감 <b>{dday_text(t)}</b></div>', unsafe_allow_html=True)

    best = max(rows, key=lambda r: r["pct"])
    low = sorted([r for r in rows if r["pct"] < 40], key=lambda r: r["pct"])
    best_txt = ("아직 끝낸 할 일이 없어요. 각자 홈에서 끝낸 일을 체크해 주세요." if overall == 0
                else f"{esc(best['name'])} ({esc(best['role'])}) {best['pct']}%")
    if low:
        check_txt = (", ".join(f"{esc(r['name'])} {r['pct']}%" for r in low)
                     + " · 본인 파트 진행이 40%에 못 미쳐요. 막힌 부분이 있는지 함께 확인해 보세요.")
    else:
        r = min(rows, key=lambda x: x["pct"])
        check_txt = f"{esc(r['name'])} {r['pct']}% · 모두 40% 이상이에요. 상대적으로 낮은 팀원만 가볍게 챙겨 주세요."
    st.markdown(
        f'<div class="summary-line"><span class="tag" style="color:#3F44D9">가장 기여도가 높은 팀원</span>{best_txt}</div>'
        f'<div class="summary-line"><span class="tag" style="color:#C9822A">추가 확인이 필요한 팀원</span>{check_txt}</div>',
        unsafe_allow_html=True)

    with st.expander("팀원별 남은 할 일 보기"):
        for r in rows:
            items = [("✅ " if c else "⬜ ") + esc(x) for x, c in zip(r["todo"], r["checks"])]
            items.append(("✅ " if r["auto"] else "⬜ ") + AUTO_ITEM)
            st.markdown(f"**{esc(r['name'])}** · {esc(r['role'])} · {r['pct']}%<br>" + "<br>".join(items),
                        unsafe_allow_html=True)

    if st.button("AI에게 진행 상황 요약 받기", disabled=not ai_ready(),
                 help=None if ai_ready() else "AI가 연결되면 쓸 수 있어요. (연결은 관리자가 해요)"):
        status = "\n".join(f"- {r['name']}({r['role']}): {r['pct']}%, 남은 일: "
                           + ", ".join(x for x, c in zip(r["todo"], r["checks"]) if not c) for r in rows)
        with st.spinner("AI가 진행 상황을 살펴보고 있어요..."):
            ss.ai_summary = ask_ai([{"role": "user", "content":
                f"{team_context(t)}\n\n[진행 상황]\n{status}\n전체 {overall}%\n\n"
                "팀 상황을 3줄로 요약하고, 마감까지 남은 기간을 고려해 이번 주에 우선 할 일 3가지를 "
                "누가 할지와 함께 제안해 줘. 진행이 느린 팀원을 탓하지 말고 도울 방법을 제안해."}], max_tokens=700)
    show_ai_error()
    if ss.get("ai_summary"):
        ai_box(ss.ai_summary, tag="AI 진행 요약")


# ──────────────────────────────────────────────────────────────
# 화면 4: AI 도우미
# ──────────────────────────────────────────────────────────────
def page_ai(d):
    t = d["team"]
    page_header("AI 도우미", "우리 팀 주제·역할·게시 자료를 알고 있는 AI에게 무엇이든 물어보세요.")
    if not ai_ready():
        note("AI가 아직 연결되지 않았어요. 관리자에게 AI 연결을 부탁해 주세요. "
             "그동안 다른 기능은 모두 그대로 쓸 수 있어요.", warn=True)
        return
    if not ss.chat:
        st.markdown("**이렇게 물어볼 수 있어요**")
        cols = st.columns(3)
        for col, q in zip(cols, ["우리 주제로 발표 목차를 짜 줘", "내 파트에서 조사할 키워드를 추천해 줘",
                                 "마감까지 남은 일정표를 만들어 줘"]):
            if col.button(q, width="stretch"):
                ss.pending_q = q
                st.rerun()
    for m in ss.chat:
        with st.chat_message(m["role"], avatar="🧩" if m["role"] == "assistant" else None):
            st.markdown(m["content"])
    q = st.chat_input("질문을 입력하세요") or ss.pop("pending_q", None)
    if q:
        ss.chat.append({"role": "user", "content": q})
        with st.chat_message("user"):
            st.markdown(q)
        system = AI_PERSONA + "\n\n[우리 팀 현재 상황]\n" + team_context(t)
        with st.chat_message("assistant", avatar="🧩"):
            with st.spinner("생각하는 중..."):
                answer = ask_ai(ss.chat[-12:], system=system, max_tokens=1200)
            if answer:
                st.markdown(answer)
                ss.chat.append({"role": "assistant", "content": answer})
            else:
                ss.chat.pop()
                show_ai_error()
    if ss.chat and st.button("대화 지우기"):
        ss.chat = []
        st.rerun()


# ──────────────────────────────────────────────────────────────
# 관리자 화면 (팀 관리 · 백업 · AI · 시스템)
# ──────────────────────────────────────────────────────────────
# 관리자 비밀번호 확인 순서
#   1) 관리자 화면에서 바꾼 비밀번호 (비공개 데이터 저장소 admin/settings.json 에 해시로 저장)
#   2) Secrets 의 ADMIN_PASSWORD  ← 처음 비밀번호는 여기에 적는 것을 추천 (나만 볼 수 있음)
#   3) 둘 다 없으면 코드에 해시로 들어 있는 처음 기본값
# 한 번 바꾸면 다시 바꾸기 전까지 항상 같은 비밀번호예요.
# 바꾼 비밀번호를 잊었다면: GitHub teamplay-data → admin/settings.json 에서 "admin_auth" 부분을 지우면 2)·3)으로 돌아가요.
ADMIN_SALT = b"teamplay-admin-v1"
ADMIN_HASH = "7650994d072b1d83ba2d795c270173b05b464a5c229afa2e376fa01132715601"
ADMIN_MAX_TRIES = 5          # 한 창에서 연속으로 틀릴 수 있는 횟수
ADMIN_GLOBAL_TRIES = 10      # 모든 사람을 합쳐 10분 안에 틀릴 수 있는 횟수
ADMIN_PW_MIN = 8


def _pbkdf(pw, salt):
    return hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, 200_000).hex()


def admin_auth():
    a = get_settings().get("admin_auth") or {}
    return a if a.get("salt") and a.get("hash") else None


def admin_pw_source():
    if admin_auth():
        return "관리자 화면에서 바꾼 비밀번호"
    if secret_key("ADMIN_PASSWORD"):
        return "Secrets의 ADMIN_PASSWORD"
    return "처음 기본값 (코드에 들어 있음)"


def admin_pw_ok(pw):
    a = admin_auth()
    if a:
        return hmac.compare_digest(_pbkdf(pw, a["salt"].encode()), a["hash"])
    secret = secret_key("ADMIN_PASSWORD")
    if secret:
        return hmac.compare_digest(pw.encode("utf-8"), str(secret).encode("utf-8"))
    return hmac.compare_digest(_pbkdf(pw, ADMIN_SALT), ADMIN_HASH)


def set_admin_pw(new_pw):
    salt = uuid.uuid4().hex
    save_settings({**get_settings(), "admin_auth": {"salt": salt, "hash": _pbkdf(new_pw, salt.encode()),
                                                     "changed_at": now_text()}})


def reset_admin_pw():
    cfg = get_settings()
    cfg.pop("admin_auth", None)
    save_settings(cfg)


def admin_globally_locked():
    """여러 창을 열어 번갈아 시도하는 것도 막음 (서버 메모리 기준)."""
    s = _store()
    now = time.time()
    with s["lock"]:
        s["admin_fails"] = [t for t in s["admin_fails"] if now - t < 600]
        return len(s["admin_fails"]) >= ADMIN_GLOBAL_TRIES


def note_admin_fail():
    s = _store()
    with s["lock"]:
        s["admin_fails"].append(time.time())


def mask(key):
    return f"…{key[-4:]}" if key else ""


def team_summary(code, d):
    t = d["team"]
    rows, overall = team_rows(d)
    return {"팀 코드": code, "주제": t["topic"], "결과물": t["output_type"], "팀원": len(t["members"]),
            "게시": sum(p["status"] == "posted" for p in d["posts"]),
            "대기": sum(p["status"] == "pending" for p in d["posts"]),
            "진행률": f"{overall}%", "마감": dday_text(t), "만든 날": (t.get("created_at") or "")[:10]}


def page_admin():
    page_header("관리자", "모든 팀과 백업, AI 연결, 저장소 상태를 관리해요.")
    if not ss.get("is_admin"):
        locked = ss.get("admin_fails", 0) >= ADMIN_MAX_TRIES and time.time() < ss.get("admin_lock_until", 0)
        with st.container(border=True):
            if locked:
                st.error("비밀번호를 여러 번 틀려서 잠시 잠겼어요. 1분 뒤에 다시 시도해 주세요.")
                return
            if admin_globally_locked():
                st.error("최근 관리자 비밀번호가 여러 번 틀려서 10분 동안 잠겼어요. 잠시 뒤에 다시 시도해 주세요.")
                return
            with st.form("admin_login", border=False):
                pw = st.text_input("관리자 비밀번호", type="password")
                if st.form_submit_button("관리자로 들어가기", type="primary"):
                    if admin_pw_ok(pw):
                        ss.is_admin, ss.admin_fails = True, 0
                        st.rerun()
                    ss.admin_fails = ss.get("admin_fails", 0) + 1
                    note_admin_fail()
                    time.sleep(1)   # 빠르게 계속 찍어 보는 것 방지
                    if ss.admin_fails >= ADMIN_MAX_TRIES:
                        ss.admin_lock_until = time.time() + 60
                    st.error("비밀번호가 맞지 않아요.")
        return

    show_flash()
    if not admin_auth() and not secret_key("ADMIN_PASSWORD"):
        note("지금은 코드에 들어 있는 <b>처음 기본 비밀번호</b>로 들어와 있어요. 앱 코드가 공개 저장소에 있으니 "
             "<b>관리자 비밀번호</b> 탭에서 바로 새 비밀번호로 바꿔 주세요.", warn=True)
    tab_team, tab_backup, tab_ai, tab_sys, tab_pw = st.tabs(["팀 관리", "백업", "AI 연결", "시스템", "관리자 비밀번호"])

    # ── 팀 관리
    with tab_team:
        try:
            if "admin_codes" not in ss or st.button("목록 새로고침"):
                ss.admin_codes = list_team_codes()
            teams = {c: admin_load(c) for c in ss.admin_codes}
        except StorageError as e:
            st.error(str(e))
            teams = {}
        teams = {c: d for c, d in teams.items() if d}
        st.caption(f"팀 {len(teams)}개")
        if teams:
            st.dataframe(pd.DataFrame([team_summary(c, d) for c, d in teams.items()]),
                         hide_index=True, width="stretch")
            code = st.selectbox("자세히 볼 팀", list(teams), format_func=lambda c: f"{c} · {teams[c]['team']['topic']}")
            d = teams[code]
            t = d["team"]
            with st.container(border=True):
                st.markdown(f"**{esc(t['topic'])}** · 팀 코드 {code}")
                st.caption("팀원: " + ", ".join(names_of(t)))
                link = invite_link(code)
                if link:
                    st.code(link, language=None)
                st.download_button("이 팀 백업 받기", json.dumps(d, ensure_ascii=False, indent=2).encode("utf-8"),
                                   file_name=f"팀플메이트_{code}_{date.today().isoformat()}.json", mime="application/json")

                with st.expander("팀원 비밀번호 다시 정하기"):
                    with st.form(f"admin_pw_{code}", clear_on_submit=True):
                        who = st.selectbox("팀원", names_of(t))
                        npw = st.text_input("새 비밀번호", type="password", placeholder=f"{PW_MIN}자 이상")
                        if st.form_submit_button("다시 정하기"):
                            if len(npw) < PW_MIN:
                                st.error(f"{PW_MIN}자 이상으로 정해 주세요.")
                            else:
                                m = find_member(d["team"], who)
                                set_pw(m, npw)
                                m["pw_notice"] = f"관리자가 {now_text()}에 내 비밀번호를 다시 정했어요."
                                admin_save(code, d)
                                ss.flash = f"{code} 팀 {who}님의 비밀번호를 바꿨어요."
                                st.rerun()

                with st.expander("팀 삭제"):
                    st.caption("팀 정보·게시판·진행 체크가 모두 지워져요. 지우기 전에 백업을 받아 두세요.")
                    typed = st.text_input("확인을 위해 팀 코드를 입력하세요", key=f"del_confirm_{code}")
                    if st.button("이 팀 삭제", disabled=typed.strip().upper() != code, type="primary"):
                        admin_save(code, copy.deepcopy(EMPTY_DATA))
                        if ss.team_code == code:
                            leave_team()
                        ss.pop("admin_codes", None)
                        ss.flash = f"{code} 팀을 삭제했어요."
                        st.rerun()
        lc = legacy_code()
        if lc:
            note(f"업데이트 전에 만든 팀({lc})의 코드가 시작 화면에 안내되고 있어요.", warn=True)
            st.button("시작 화면 안내 숨기기", on_click=hide_legacy_notice, key="admin_hide_legacy")

    # ── 백업
    with tab_backup:
        st.markdown("**전체 백업**")
        st.caption("모든 팀을 파일 하나로 받아요. 평소에는 GitHub에 자동 저장되니, 큰 변경 전에 한 번씩 받아 두면 충분해요.")
        if st.button("전체 백업 파일 만들기"):
            try:
                all_teams = {c: admin_load(c) for c in list_team_codes()}
                ss.admin_backup = json.dumps({"kind": "teamplay-backup", "exported_at": now_text(),
                                              "teams": {c: d for c, d in all_teams.items() if d}},
                                             ensure_ascii=False, indent=1).encode("utf-8")
            except StorageError as e:
                st.error(str(e))
        if ss.get("admin_backup"):
            st.download_button("전체 백업 받기", ss.admin_backup, mime="application/json",
                               file_name=f"팀플메이트_전체백업_{date.today().isoformat()}.json")

        st.divider()
        st.markdown("**백업 되돌리기**")
        up = st.file_uploader("백업 파일 (.json)", type=["json"], key="admin_restore")
        if up is not None:
            try:
                obj = json.loads(up.getvalue().decode("utf-8"))
            except (ValueError, UnicodeDecodeError):
                obj = None
            if isinstance(obj, dict) and isinstance(obj.get("teams"), dict):
                st.caption(f"전체 백업 · 팀 {len(obj['teams'])}개 ({obj.get('exported_at', '')})")
                if st.button("이 백업으로 모든 팀 되돌리기", type="primary"):
                    for c, d in obj["teams"].items():
                        if len(c) == 6 and isinstance(d, dict) and d.get("team"):
                            admin_save(c, normalize(_clean(d)))
                    ss.pop("admin_codes", None)
                    ss.flash = "전체 백업을 되돌렸어요. 백업에 없는 팀은 그대로예요."
                    st.rerun()
            elif isinstance(obj, dict) and obj.get("team"):
                st.caption(f"팀 하나 백업 · {obj['team'].get('topic', '')}")
                target = st.text_input("되돌릴 팀 코드 (비우면 새 코드로 만들어요)", max_chars=6).strip().upper()
                if st.button("이 팀 되돌리기", type="primary"):
                    code = target or new_team_code()
                    admin_save(code, normalize(_clean(obj)))
                    ss.pop("admin_codes", None)
                    ss.flash = f"{code} 팀으로 되돌렸어요."
                    st.rerun()
            else:
                st.error("팀플 메이트 백업 파일이 아니에요.")

    # ── AI 연결
    with tab_ai:
        cfg = get_settings()
        with st.form("ai_form"):
            prov = st.radio("AI 종류 (모든 팀 공통)", AI_PROVIDERS, horizontal=True,
                            index=AI_PROVIDERS.index(provider()))
            gm = st.selectbox("Gemini 모델", list(GEMINI_MODELS),
                              index=list(GEMINI_MODELS).index(cfg["gemini_model"]) if cfg.get("gemini_model") in GEMINI_MODELS else 0)
            cm = st.selectbox("Claude 모델", list(AI_MODELS),
                              index=list(AI_MODELS).index(cfg["claude_model"]) if cfg.get("claude_model") in AI_MODELS else 0)
            if st.form_submit_button("저장", type="primary"):
                save_settings({**cfg, "provider": prov, "gemini_model": gm, "claude_model": cm})
                ss.flash = "AI 설정을 저장했어요. 모든 팀에 바로 적용돼요."
                st.rerun()

        st.markdown("**API 키 상태**")
        for name, label in (("GEMINI_API_KEY", "Gemini"), ("ANTHROPIC_API_KEY", "Claude")):
            key, src = key_source(name)
            st.markdown(f'<div class="team-line"><b>{label}</b> · {src}'
                        f'<span class="team-pct">{mask(key)}</span></div>', unsafe_allow_html=True)
        with st.expander("임시 키 넣기 (앱을 다시 켜면 사라짐)"):
            st.caption("키를 계속 쓰려면 share.streamlit.io → 내 앱 ⋮ → Settings → Secrets 에서 바꿔 주세요. "
                       "여기 넣은 키는 Secrets보다 먼저 쓰여요.")
            with st.form("tmp_key_form", clear_on_submit=True):
                which = st.selectbox("어떤 키", ["GEMINI_API_KEY", "ANTHROPIC_API_KEY"])
                val = st.text_input("키", type="password")
                c1, c2 = st.columns(2)
                if c1.form_submit_button("임시 키 적용") and val.strip():
                    _store()["keys"][which] = val.strip()
                    ss.flash = f"{which} 임시 키를 적용했어요."
                    st.rerun()
                if c2.form_submit_button("임시 키 지우기"):
                    _store()["keys"].pop(which, None)
                    ss.flash = f"{which} 임시 키를 지웠어요."
                    st.rerun()
        if st.button("연결 테스트", disabled=not ai_ready()):
            ok = ask_ai([{"role": "user", "content": "'연결 성공'이라고만 답해."}], max_tokens=30)
            used = f" ({ss.get('ai_model_used', '')})" if ok and is_gemini() else ""
            if ok:
                st.success(f"연결 성공!{used}")
            else:
                show_ai_error()
        if not ai_ready():
            st.caption(f"지금 선택한 {provider()}의 키가 없어서 AI 기능이 꺼져 있어요. 앱은 규칙 기반으로 계속 작동해요.")

    # ── 시스템
    with tab_sys:
        s = _store()
        gh = _gh_cfg()
        st.markdown(
            '<div class="mini-scores" style="line-height:2">'
            f'저장 위치 <b>{storage_mode()}</b>{" · " + esc(gh["repo"]) if gh else ""}<br>'
            f'저장 상태 <b>{esc(storage_status())}</b><br>'
            f'메모리에 올라온 문서 <b>{len(s["docs"])}개</b> · 저장 대기 <b>{len(s["dirty"])}개</b><br>'
            f'관리자 비밀번호 <b>{admin_pw_source()}</b><br>'
            f'Streamlit <b>{st.__version__}</b></div>', unsafe_allow_html=True)
        st.write("")
        busy = bool(s["dirty"] or s["writer"])
        if st.button("저장소에서 다시 읽기", disabled=busy,
                     help="GitHub에서 직접 고친 내용을 앱에 반영할 때 써요. 저장 중일 때는 누를 수 없어요."):
            st.cache_resource.clear()
            ss.pop("admin_codes", None)
            ss.flash = "저장소에서 다시 읽었어요."
            st.rerun()
        if st.button("관리자 로그아웃"):
            ss.is_admin, ss.admin_mode = False, False
            st.rerun()

    # ── 관리자 비밀번호
    with tab_pw:
        a = admin_auth()
        st.markdown(f'<div class="mini-scores">지금 쓰는 비밀번호 <b>{admin_pw_source()}</b>'
                    + (f' · {esc(a.get("changed_at", ""))}에 바꿈' if a else "") + '</div>', unsafe_allow_html=True)
        st.caption("한 번 바꾸면 다시 바꾸기 전까지 항상 같은 비밀번호예요. 비밀번호는 알아볼 수 없는 형태로 "
                   "비공개 데이터 저장소에만 저장돼요.")
        with st.form("admin_pw_form", clear_on_submit=True):
            cur = st.text_input("지금 비밀번호", type="password")
            n1 = st.text_input("새 비밀번호", type="password", placeholder=f"{ADMIN_PW_MIN}자 이상")
            n2 = st.text_input("새 비밀번호 한 번 더", type="password")
            if st.form_submit_button("비밀번호 바꾸기", type="primary"):
                if not admin_pw_ok(cur):
                    st.error("지금 비밀번호가 맞지 않아요.")
                elif len(n1) < ADMIN_PW_MIN:
                    st.error(f"새 비밀번호는 {ADMIN_PW_MIN}자 이상으로 정해 주세요.")
                elif n1 != n2:
                    st.error("새 비밀번호 두 개가 서로 달라요.")
                else:
                    set_admin_pw(n1)
                    ss.flash = "관리자 비밀번호를 바꿨어요. 다음부터는 새 비밀번호로 들어와요."
                    st.rerun()
        if a:
            with st.expander("처음 비밀번호로 되돌리기"):
                st.caption("앱에서 바꾼 비밀번호를 지우고, Secrets의 ADMIN_PASSWORD(없으면 처음 기본값)로 되돌려요.")
                if st.button("되돌리기"):
                    reset_admin_pw()
                    ss.flash = "관리자 비밀번호를 처음 비밀번호로 되돌렸어요."
                    st.rerun()


# ──────────────────────────────────────────────────────────────
# 화면 고르기
# ──────────────────────────────────────────────────────────────
if ss.get("admin_mode"):
    page_admin()
elif storage_error:
    st.error("저장소에 연결하지 못했어요. 기존 데이터를 지키기 위해 잠시 멈췄어요.")
    st.caption(storage_error)
    st.button("다시 연결하기", on_click=st.cache_resource.clear, type="primary")
elif not team and ss.creating:
    page_setup()
elif not team:
    page_start()
elif not ss.user:
    page_login(data)
else:
    {"홈": page_home, "파트 분배": page_parts, "자료 내 추적": page_materials,
     "기여도 평가": page_contrib, "AI 도우미": page_ai}[page](data)

# 사이드바 상태 요약: 이번 화면에서 바뀐 내용까지 반영되도록 맨 마지막에 그림
try:
    _d = get_data()
except StorageError:
    _d = copy.deepcopy(EMPTY_DATA)
_t = _d["team"]
if ss.get("admin_mode"):
    status_box.markdown(f'<div class="side-stat">저장 위치 &nbsp;<b>{storage_mode()}</b><br>'
                        f'AI &nbsp;<b>{(provider().split()[0] + " 연결됨") if ai_ready() else "꺼짐"}</b></div>',
                        unsafe_allow_html=True)
elif _t:
    _posted = sum(p["status"] == "posted" for p in _d["posts"])
    _pending = sum(p["status"] == "pending" for p in _d["posts"])
    _mine = f'내 투표 대기 &nbsp;<b>{len(pending_for(_d, ss.user))}건</b><br>' if ss.user else ""
    status_box.markdown(
        f'<div class="side-stat">주제 &nbsp;<b>{esc(_t["topic"])}</b><br>'
        f'결과물 &nbsp;<b>{esc(_t["output_type"])}</b><br>'
        f'팀원 &nbsp;<b>{len(_t["members"])}명</b><br>'
        f'마감 &nbsp;<b>{dday_text(_t)}</b><br>'
        f'AI &nbsp;<b>{(provider().split()[0] + " 연결됨") if ai_ready() else "꺼짐 (규칙 기반)"}</b><br>'
        f'게시판 &nbsp;<b>게시 {_posted}건 · 대기 {_pending}건</b><br>{_mine}</div>',
        unsafe_allow_html=True)
elif not storage_error:
    status_box.markdown(f'<div class="side-stat">팀 코드를 넣거나 새 팀을 만들어 주세요.<br>'
                        f'AI &nbsp;<b>{(provider().split()[0] + " 연결됨") if ai_ready() else "꺼짐 (규칙 기반)"}</b></div>',
                        unsafe_allow_html=True)
if _store()["error"]:
    st.sidebar.error(f"⚠ {_store()['error']}")
