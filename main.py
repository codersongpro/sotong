#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""충북 소통메신저 자동 사용자 선택 프로그램"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading, time, json, os, re, subprocess, sys, difflib

try:
    import pyautogui
    pyautogui.PAUSE = 0.15
    pyautogui.FAILSAFE = True
except ImportError:
    pyautogui = None

try:
    import pyperclip
except ImportError:
    pyperclip = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".chungbuk_auto_config.json")

# 충북교육청 소속기관 DB (모듈 로드 시 구성)
_ORG_LOOKUP: dict = {}


# ══════════════════════════════════════════════════════════════
#  파서
# ══════════════════════════════════════════════════════════════

# [2026-04-02 14:35:31.907000] 형태 타임스탬프
TIMESTAMP_RE = re.compile(r'^\[\d{4}[-/]\d{2}[-/]\d{2}[\d\s:.,\-]*\]\s*')

# 학교·기관 끝 단어
ORG_END_RE = re.compile(
    r'(학교|초등학교|중학교|고등학교|특수학교|유치원|'
    r'교육청|교육지원청|교육원|교육관|연구원|연구소|'
    r'도청|시청|군청|구청|교육부|본청|센터|지원청)$'
)
# 학교 약칭: 가경초, 세종중, 청주내곡초 등 (최대 6자 + 초/중/고)
SCHOOL_ABBR_RE = re.compile(r'^[가-힣]{1,6}(초|중|고)$')

# 쓸모없는 토큰: 순수숫자, 전화번호, URL, 특수문자 포함
JUNK_RE = re.compile(
    r'^\d+$'
    r'|\d{2,4}-\d{3,4}-\d{4}'
    r'|^https?://'
    r'|[!@#$%^*()\[\]{}<>|\\/?~`]'
)


def is_org(s: str) -> bool:
    s = s.strip()
    if bool(ORG_END_RE.search(s)) or bool(SCHOOL_ABBR_RE.match(s)):
        return True
    return s in _ORG_LOOKUP or abbreviate_school(s) in _ORG_LOOKUP


def is_person_name(s: str) -> bool:
    """2~4자 한글, 학교명 아닌 것"""
    s = s.strip()
    return bool(re.match(r'^[가-힣]{2,4}$', s)) and not is_org(s)


def abbreviate_school(name: str) -> str:
    """풀 학교명 → 검색용 약칭 (백곡초등학교 → 백곡초)"""
    name = name.strip()
    for full, short in [('초등학교', '초'), ('중학교', '중'), ('고등학교', '고')]:
        if name.endswith(full):
            return name[:-len(full)] + short
    return name


def _split_school_suffix(s: str):
    """학교명에서 접두어·종류 분리. '진천상신초등학교' 또는 '진천상신초' → ('진천상신', '초')"""
    for full, short in [('초등학교', '초'), ('중학교', '중'), ('고등학교', '고')]:
        if s.endswith(full):
            return s[:-len(full)], short
    for suf in ('초', '중', '고'):
        if s.endswith(suf):
            return s[:-1], suf
    return None, None


def lookup_org(name: str) -> str:
    """소속명 → DB 기준 표준 약칭. 오타 자동 보정."""
    name = name.strip()
    abbr = abbreviate_school(name)
    if name in _ORG_LOOKUP:
        return _ORG_LOOKUP[name]
    if abbr in _ORG_LOOKUP:
        return _ORG_LOOKUP[abbr]
    # 퍼지 매칭 (4자 이상 입력만 시도)
    keys = list(_ORG_LOOKUP.keys())
    for q in (abbr, name):
        if len(q) < 4:
            continue
        prefix, suf = _split_school_suffix(q)
        if prefix is not None:
            # 학교명: 같은 종류(초/중/고) 안에서 접두어끼리만 비교
            same_suf = [k for k in keys if k.endswith(suf)]
            prefix_map = {k[:-1]: k for k in same_suf}
            # 1) 후미 포함: "청주내곡" → "내곡" 처럼 지역 접두어가 붙은 경우
            for k_prefix, k_full in prefix_map.items():
                if len(k_prefix) >= 2 and prefix.endswith(k_prefix):
                    return _ORG_LOOKUP[k_full]
            # 2) 오타 보정: 접두어끼리 퍼지 매칭 (cutoff 높여 오매칭 방지)
            m = difflib.get_close_matches(prefix, list(prefix_map), n=1, cutoff=0.80)
            if m:
                return _ORG_LOOKUP[prefix_map[m[0]]]
        else:
            m = difflib.get_close_matches(q, keys, n=1, cutoff=0.70)
            if m:
                return _ORG_LOOKUP[m[0]]
    return abbr  # DB에 없으면 약칭 그대로


def best_org_from(tokens: list) -> str:
    """토큰 목록에서 소속명 추출"""
    for t in tokens:
        if is_org(t):
            return lookup_org(t)
    for t in tokens:
        if re.match(r'^[가-힣]{2,}$', t) and not is_person_name(t):
            return lookup_org(t)
    return ''


def extract_pair(tokens: list):
    """토큰 목록 → (소속, 이름) 또는 None"""
    tokens = [t for t in tokens if t and not JUNK_RE.search(t)]
    if not tokens:
        return None

    name_cands = [(i, t) for i, t in enumerate(tokens) if is_person_name(t)]

    if name_cands:
        name_i, name = name_cands[0]
        others = [t for i, t in enumerate(tokens) if i != name_i]
        org = best_org_from(others)
        return (org, name)

    # 이름 후보 없음 → 2~4자 한글 토큰 마지막 것
    cands = [(i, t) for i, t in enumerate(tokens) if re.match(r'^[가-힣]{2,4}$', t)]
    if cands:
        idx, name = cands[-1]
        others = [t for i, t in enumerate(tokens) if i != idx]
        org = best_org_from(others)
        return (org, name)

    return None


def _tokenize(line: str) -> list:
    """줄 → 정제된 토큰 목록"""
    if '\t' in line:
        tokens = [t.strip() for t in line.split('\t') if t.strip()]
    elif ' : ' in line:
        tokens = line.replace(' : ', ' ').split()
    else:
        tokens = line.split()
    return [t for t in tokens if t and not JUNK_RE.search(t)]


def parse_input(text: str):
    """
    스마트 파서: 표에서 복붙할 때 소속/이름이 별도 행으로 오는 경우를
    pending_org 버퍼로 연결.

    지원 패턴:
    ① 탭 구분 (Excel 표 복붙)       : 소속\t이름
    ② 행 교대 (HWP 표 복붙)         : 소속\n이름\n소속\n이름
    ③ 한 행에 소속+이름 (공백 구분)  : 백곡초등학교 김순범
    ④ 타임스탬프 채팅 로그           : [2026-04-02 ...] 이름 소속
    """
    lines = text.strip().splitlines()
    results: list = []
    skipped: list = []
    pending_org: str | None = None  # 소속행 버퍼

    for raw in lines:
        line = raw.strip()
        if not line:
            continue

        # 타임스탬프 제거
        line = TIMESTAMP_RE.sub('', line).strip()
        if not line:
            continue

        # 앞 줄번호 제거 "1. " "1) "
        line = re.sub(r'^\d+[.)]\s*', '', line).strip()

        tokens = _tokenize(line)
        if not tokens:
            continue

        all_names = [t for t in tokens if is_person_name(t)]
        all_orgs  = [t for t in tokens if is_org(t)]

        # ── 순수 소속 행 ──────────────────────────────────────
        if all_orgs and not all_names and len(tokens) <= 3:
            pending_org = best_org_from(tokens)

        # ── 순수 이름 행 (1개 이상) ───────────────────────────
        elif all_names and not all_orgs and all(is_person_name(t) for t in tokens):
            for name in all_names:
                results.append((pending_org or '', name))
            # pending_org 유지: 같은 소속에 여러 명 가능

        # ── 혼합 행 (소속 + 이름이 같은 줄) ──────────────────
        else:
            pair = extract_pair(tokens)
            if pair:
                org, name = pair
                # 이 줄에 소속이 없으면 버퍼 소속 사용
                if not org and pending_org:
                    org = pending_org
                results.append((org, name))
            else:
                skipped.append(raw)
            pending_org = None  # 혼합/미인식 행 후 버퍼 초기화

    return results, skipped


# ══════════════════════════════════════════════════════════════
#  충북교육청 소속기관 데이터베이스
#  형식: (입력될 수 있는 이름,  검색용 약칭)
# ══════════════════════════════════════════════════════════════

_CHUNGBUK_RAW: list = [
    # ── 본청 및 직속기관 ──────────────────────────────────────
    ("충청북도교육청",            "충북교육청"),
    ("충북교육청",                "충북교육청"),
    ("충청북도교육청본청",         "충북교육청"),
    ("충청북도교육연구정보원",      "도교육연구정보원"),
    ("도교육연구정보원",           "도교육연구정보원"),
    ("충청북도단재교육연수원",      "단재교육연수원"),
    ("단재교육연수원",             "단재교육연수원"),
    ("충청북도국제교육원",         "충북국제교육원"),
    ("충북국제교육원",             "충북국제교육원"),
    ("충청북도학생교육원",         "충북학생교육원"),
    ("충북학생교육원",             "충북학생교육원"),
    ("충청북도학생교육문화원",      "충북학생교육문화원"),
    ("충북학생교육문화원",          "충북학생교육문화원"),
    ("충청북도인재개발원",         "충북인재개발원"),
    ("충북인재개발원",             "충북인재개발원"),
    ("충청북도교육청인재개발원",    "충북인재개발원"),

    # ── 교육지원청 ────────────────────────────────────────────
    ("청주교육지원청",    "청주교육지원청"),
    ("충주교육지원청",    "충주교육지원청"),
    ("제천교육지원청",    "제천교육지원청"),
    ("보은교육지원청",    "보은교육지원청"),
    ("옥천교육지원청",    "옥천교육지원청"),
    ("영동교육지원청",    "영동교육지원청"),
    ("증평교육지원청",    "증평교육지원청"),
    ("진천교육지원청",    "진천교육지원청"),
    ("괴산증평교육지원청","괴산증평교육지원청"),
    ("괴산교육지원청",    "괴산증평교육지원청"),
    ("음성교육지원청",    "음성교육지원청"),
    ("단양교육지원청",    "단양교육지원청"),

    # ── 청주시 초등학교 ───────────────────────────────────────
    ("가경초등학교","가경초"), ("강내초등학교","강내초"),
    ("강서초등학교","강서초"), ("개신초등학교","개신초"),
    ("경덕초등학교","경덕초"), ("계성초등학교","계성초"),
    ("공북초등학교","공북초"), ("교동초등학교","교동초"),
    ("내덕초등학교","내덕초"), ("내곡초등학교","내곡초"),
    ("내수초등학교","내수초"), ("남성초등학교","남성초"),
    ("남일초등학교","남일초"), ("남주초등학교","남주초"),
    ("낭성초등학교","낭성초"), ("노원초등학교","노원초"),
    ("녹원초등학교","녹원초"), ("달천초등학교","달천초"),
    ("대성초등학교","대성초"), ("대원초등학교","대원초"),
    ("동막초등학교","동막초"), ("동주초등학교","동주초"),
    ("동화초등학교","동화초"), ("마산초등학교","마산초"),
    ("모충초등학교","모충초"), ("문의초등학교","문의초"),
    ("미원초등학교","미원초"), ("방서초등학교","방서초"),
    ("복대초등학교","복대초"), ("봉명초등학교","봉명초"),
    ("북일초등학교","북일초"), ("비하초등학교","비하초"),
    ("사직초등학교","사직초"), ("산남초등학교","산남초"),
    ("산성초등학교","산성초"), ("상당초등학교","상당초"),
    ("서원초등학교","서원초"), ("석교초등학교","석교초"),
    ("성화초등학교","성화초"), ("세교초등학교","세교초"),
    ("솔밭초등학교","솔밭초"), ("수곡초등학교","수곡초"),
    ("수동초등학교","수동초"), ("신봉초등학교","신봉초"),
    ("신성초등학교","신성초"), ("신전초등학교","신전초"),
    ("신흥초등학교","신흥초"), ("양성초등학교","양성초"),
    ("영운초등학교","영운초"), ("오창초등학교","오창초"),
    ("운동초등학교","운동초"), ("운천초등학교","운천초"),
    ("원봉초등학교","원봉초"), ("원평초등학교","원평초"),
    ("율량초등학교","율량초"), ("율봉초등학교","율봉초"),
    ("은행초등학교","은행초"), ("이화초등학교","이화초"),
    ("일신초등학교","일신초"), ("잠두초등학교","잠두초"),
    ("중앙초등학교","중앙초"), ("중흥초등학교","중흥초"),
    ("지북초등학교","지북초"), ("지우초등학교","지우초"),
    ("청남초등학교","청남초"), ("청당초등학교","청당초"),
    ("청룡초등학교","청룡초"), ("청원초등학교","청원초"),
    ("충일초등학교","충일초"), ("탑대성초등학교","탑대성초"),
    ("팔봉초등학교","팔봉초"), ("평촌초등학교","평촌초"),
    ("한벌초등학교","한벌초"), ("현도초등학교","현도초"),
    ("형촌초등학교","형촌초"), ("화계초등학교","화계초"),
    ("흥덕초등학교","흥덕초"),

    # ── 청주시 중학교 ─────────────────────────────────────────
    ("가경중학교","가경중"), ("강내중학교","강내중"),
    ("강서중학교","강서중"), ("경덕중학교","경덕중"),
    ("교동중학교","교동중"), ("내덕중학교","내덕중"),
    ("내수중학교","내수중"), ("남성중학교","남성중"),
    ("대성중학교","대성중"), ("동주중학교","동주중"),
    ("문의중학교","문의중"), ("미원중학교","미원중"),
    ("방서중학교","방서중"), ("봉명중학교","봉명중"),
    ("북일중학교","북일중"), ("산남중학교","산남중"),
    ("상당중학교","상당중"), ("서원중학교","서원중"),
    ("성화중학교","성화중"), ("세교중학교","세교중"),
    ("솔밭중학교","솔밭중"), ("수곡중학교","수곡중"),
    ("신봉중학교","신봉중"), ("신성중학교","신성중"),
    ("양성중학교","양성중"), ("영운중학교","영운중"),
    ("오창중학교","오창중"), ("운천중학교","운천중"),
    ("율량중학교","율량중"), ("청남중학교","청남중"),
    ("청원중학교","청원중"), ("충일중학교","충일중"),
    ("팔봉중학교","팔봉중"), ("현도중학교","현도중"),
    ("흥덕중학교","흥덕중"),

    # ── 청주시 고등학교 ───────────────────────────────────────
    ("가경고등학교","가경고"), ("경덕고등학교","경덕고"),
    ("교동고등학교","교동고"), ("내덕고등학교","내덕고"),
    ("대성고등학교","대성고"), ("동주고등학교","동주고"),
    ("봉명고등학교","봉명고"), ("북일고등학교","북일고"),
    ("산남고등학교","산남고"), ("상당고등학교","상당고"),
    ("서원고등학교","서원고"), ("세광고등학교","세광고"),
    ("성화고등학교","성화고"), ("수곡고등학교","수곡고"),
    ("신성고등학교","신성고"), ("오창고등학교","오창고"),
    ("운천고등학교","운천고"), ("율량고등학교","율량고"),
    ("청남고등학교","청남고"), ("청주고등학교","청주고"),
    ("청주대성고등학교","청주대성고"), ("청주상업고등학교","청주상업고"),
    ("청주여자고등학교","청주여고"), ("청주여고","청주여고"),
    ("청주외국어고등학교","청주외국어고"), ("충북고등학교","충북고"),
    ("충북고","충북고"), ("충일고등학교","충일고"),
    ("팔봉고등학교","팔봉고"), ("현도고등학교","현도고"),
    ("흥덕고등학교","흥덕고"),

    # ── 충주시 초등학교 ───────────────────────────────────────
    ("가금초등학교","가금초"), ("개령초등학교","개령초"),
    ("달천초등학교","달천초"), ("대림초등학교","대림초"),
    ("대소원초등학교","대소원초"), ("목벌초등학교","목벌초"),
    ("문화초등학교","문화초"), ("봉방초등학교","봉방초"),
    ("성남초등학교","성남초"), ("성내초등학교","성내초"),
    ("소태초등학교","소태초"), ("수안보초등학교","수안보초"),
    ("신니초등학교","신니초"), ("안림초등학교","안림초"),
    ("앙성초등학교","앙성초"), ("엄정초등학교","엄정초"),
    ("연수초등학교","연수초"), ("용탄초등학교","용탄초"),
    ("주덕초등학교","주덕초"), ("충원초등학교","충원초"),
    ("탄금초등학교","탄금초"),

    # ── 충주시 중학교 ─────────────────────────────────────────
    ("가금중학교","가금중"), ("달천중학교","달천중"),
    ("대소원중학교","대소원중"), ("봉방중학교","봉방중"),
    ("성남중학교","성남중"), ("성내중학교","성내중"),
    ("수안보중학교","수안보중"), ("신니중학교","신니중"),
    ("엄정중학교","엄정중"), ("연수중학교","연수중"),
    ("주덕중학교","주덕중"), ("충원중학교","충원중"),
    ("탄금중학교","탄금중"),

    # ── 충주시 고등학교 ───────────────────────────────────────
    ("달천고등학교","달천고"), ("봉방고등학교","봉방고"),
    ("성남고등학교","성남고"), ("성내고등학교","성내고"),
    ("연수고등학교","연수고"), ("충주고등학교","충주고"),
    ("충주공업고등학교","충주공고"), ("충주공고","충주공고"),
    ("충주대원고등학교","충주대원고"),
    ("충주여자고등학교","충주여고"), ("충주여고","충주여고"),
    ("충주예성여자고등학교","충주예성여고"),

    # ── 제천시 초등학교 ───────────────────────────────────────
    ("두학초등학교","두학초"), ("봉양초등학교","봉양초"),
    ("세명초등학교","세명초"), ("신월초등학교","신월초"),
    ("용두초등학교","용두초"), ("의림초등학교","의림초"),
    ("자작초등학교","자작초"), ("제천초등학교","제천초"),
    ("청전초등학교","청전초"), ("한수초등학교","한수초"),

    # ── 제천시 중학교 ─────────────────────────────────────────
    ("봉양중학교","봉양중"), ("세명중학교","세명중"),
    ("신월중학교","신월중"), ("의림중학교","의림중"),
    ("제천중학교","제천중"), ("청전중학교","청전중"),

    # ── 제천시 고등학교 ───────────────────────────────────────
    ("세명고등학교","세명고"), ("의림여자고등학교","의림여고"),
    ("제천고등학교","제천고"), ("제천여자고등학교","제천여고"),
    ("제천비즈니스고등학교","제천비즈니스고"),

    # ── 보은군 초등학교 ───────────────────────────────────────
    ("보은초등학교","보은초"), ("삼승초등학교","삼승초"),
    ("수한초등학교","수한초"), ("탄부초등학교","탄부초"),
    ("회남초등학교","회남초"),

    # ── 보은군 중·고 ─────────────────────────────────────────
    ("보은중학교","보은중"), ("삼승중학교","삼승중"),
    ("수한중학교","수한중"), ("보은고등학교","보은고"),

    # ── 옥천군 초등학교 ───────────────────────────────────────
    ("옥천초등학교","옥천초"), ("군북초등학교","군북초"),
    ("동이초등학교","동이초"), ("이원초등학교","이원초"),
    ("청산초등학교","청산초"), ("청성초등학교","청성초"),

    # ── 옥천군 중·고 ─────────────────────────────────────────
    ("옥천중학교","옥천중"), ("군북중학교","군북중"),
    ("이원중학교","이원중"), ("청산중학교","청산중"),
    ("옥천고등학교","옥천고"),

    # ── 영동군 초등학교 ───────────────────────────────────────
    ("영동초등학교","영동초"), ("매곡초등학교","매곡초"),
    ("상촌초등학교","상촌초"), ("심천초등학교","심천초"),
    ("용산초등학교","용산초"), ("용화초등학교","용화초"),
    ("학산초등학교","학산초"), ("황간초등학교","황간초"),

    # ── 영동군 중·고 ─────────────────────────────────────────
    ("영동중학교","영동중"), ("매곡중학교","매곡중"),
    ("심천중학교","심천중"), ("용화중학교","용화중"),
    ("황간중학교","황간중"),
    ("영동고등학교","영동고"), ("황간고등학교","황간고"),

    # ── 증평군 초등학교 ───────────────────────────────────────
    ("남하초등학교","남하초"), ("보강초등학교","보강초"),
    ("증평초등학교","증평초"), ("증평남초등학교","증평남초"),

    # ── 증평군 중·고 ─────────────────────────────────────────
    ("증평중학교","증평중"), ("증평남중학교","증평남중"),
    ("증평고등학교","증평고"),

    # ── 진천군 초등학교 ───────────────────────────────────────
    ("광혜원초등학교","광혜원초"), ("문백초등학교","문백초"),
    ("백곡초등학교","백곡초"), ("이월초등학교","이월초"),
    ("진천초등학교","진천초"), ("초평초등학교","초평초"),
    ("덕산초등학교","덕산초"),

    # ── 진천군 중·고 ─────────────────────────────────────────
    ("광혜원중학교","광혜원중"), ("덕산중학교","덕산중"),
    ("이월중학교","이월중"), ("진천중학교","진천중"),
    ("초평중학교","초평중"),
    ("진천고등학교","진천고"), ("진천상공고등학교","진천상공고"),

    # ── 괴산군 초등학교 ───────────────────────────────────────
    ("괴산초등학교","괴산초"), ("문광초등학교","문광초"),
    ("불정초등학교","불정초"), ("사리초등학교","사리초"),
    ("소수초등학교","소수초"), ("연풍초등학교","연풍초"),
    ("장연초등학교","장연초"), ("청천초등학교","청천초"),
    ("칠성초등학교","칠성초"),

    # ── 괴산군 중·고 ─────────────────────────────────────────
    ("괴산중학교","괴산중"), ("문광중학교","문광중"),
    ("소수중학교","소수중"), ("연풍중학교","연풍중"),
    ("청천중학교","청천중"), ("칠성중학교","칠성중"),
    ("괴산고등학교","괴산고"), ("칠성고등학교","칠성고"),

    # ── 음성군 초등학교 ───────────────────────────────────────
    ("감곡초등학교","감곡초"), ("금왕초등학교","금왕초"),
    ("대소초등학교","대소초"), ("맹동초등학교","맹동초"),
    ("삼성초등학교","삼성초"), ("생극초등학교","생극초"),
    ("소이초등학교","소이초"), ("원남초등학교","원남초"),
    ("음성초등학교","음성초"),

    # ── 음성군 중·고 ─────────────────────────────────────────
    ("감곡중학교","감곡중"), ("금왕중학교","금왕중"),
    ("대소중학교","대소중"), ("맹동중학교","맹동중"),
    ("삼성중학교","삼성중"), ("소이중학교","소이중"),
    ("음성중학교","음성중"),
    ("금왕고등학교","금왕고"), ("대소고등학교","대소고"),
    ("음성고등학교","음성고"),

    # ── 단양군 초등학교 ───────────────────────────────────────
    ("가곡초등학교","가곡초"), ("대강초등학교","대강초"),
    ("매포초등학교","매포초"), ("단양초등학교","단양초"),
    ("어상천초등학교","어상천초"), ("영춘초등학교","영춘초"),
    ("적성초등학교","적성초"),

    # ── 단양군 중·고 ─────────────────────────────────────────
    ("단양중학교","단양중"), ("대강중학교","대강중"),
    ("매포중학교","매포중"), ("영춘중학교","영춘중"),
    ("적성중학교","적성중"), ("단양고등학교","단양고"),

    # ── 특수학교 ─────────────────────────────────────────────
    ("청주성심학교","청주성심학교"),
    ("청주명암학교","청주명암학교"),
    ("충북혜림학교","충북혜림학교"),
    ("충북혜림특수학교","충북혜림학교"),
    ("충북맹학교","충북맹학교"),
    ("충북농학교","충북농학교"),
    ("충주성모학교","충주성모학교"),
    ("제천혜화학교","제천혜화학교"),
    ("청주진흥학교","청주진흥학교"),
]

# _ORG_LOOKUP 구성: 풀이름·약칭·abbreviate 형 모두 키로 등록
for _full, _abbr in _CHUNGBUK_RAW:
    _ORG_LOOKUP[_full]  = _abbr
    _ORG_LOOKUP[_abbr]  = _abbr
    _ab2 = abbreviate_school(_full)
    if _ab2 and _ab2 != _full:
        _ORG_LOOKUP[_ab2] = _abbr


# ══════════════════════════════════════════════════════════════
#  HWP 텍스트 추출
# ══════════════════════════════════════════════════════════════

def extract_hwp_text(filepath: str):
    """HWP / HWPX 파일에서 텍스트 추출 (3가지 방법 순차 시도)"""
    ext = os.path.splitext(filepath)[1].lower()

    # HWPX: ZIP 기반 XML 형식
    if ext == '.hwpx':
        t = _hwpx_zipxml(filepath)
        if t:
            return t

    # 방법 1: win32com (한/글 설치된 PC)
    t = _hwp_win32com(filepath)
    if t:
        return t

    # 방법 2: olefile 바이너리 명단 추출 (한/글 미설치도 가능)
    t = _hwp_olefile(filepath)
    if t:
        return t

    # 방법 3: hwp5txt 커맨드라인
    try:
        r = subprocess.run(
            ["hwp5txt", filepath],
            capture_output=True, text=True, encoding="utf-8", timeout=30
        )
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout
    except Exception:
        pass

    return None


def _hwp_win32com(filepath: str):
    try:
        import win32com.client
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.Open(filepath, "HWP", "forceopen:true")
        text = hwp.GetTextFile("TEXT", "")
        hwp.Quit()
        return text if text and text.strip() else None
    except Exception:
        return None


def _hwp_olefile(filepath: str):
    """
    olefile로 HWP 5.x 바이너리 직접 명단 추출.
    BodyText/Section* 스트림에서 HWPTAG_PARA_TEXT(tag=67) 레코드 추출.
    """
    try:
        import olefile, zlib, struct
    except ImportError:
        return None
    try:
        ole = olefile.OleFileIO(filepath)

        # FileHeader에서 압축 여부 확인
        is_compressed = True
        if ole.exists('FileHeader'):
            hdr = ole.openstream('FileHeader').read()
            if len(hdr) >= 36:
                flags = struct.unpack_from('<I', hdr, 32)[0]
                is_compressed = bool(flags & 0x01)

        text_parts = []
        for i in range(512):
            sname = f'BodyText/Section{i}'
            if not ole.exists(sname):
                break
            data = ole.openstream(sname).read()

            if is_compressed:
                try:
                    data = zlib.decompress(data, -15)   # raw deflate
                except Exception:
                    try:
                        data = zlib.decompress(data)
                    except Exception:
                        continue

            pos = 0
            while pos + 4 <= len(data):
                h = struct.unpack_from('<I', data, pos)[0]
                tag  = h & 0x3FF
                size = (h >> 20) & 0xFFF
                pos += 4
                if size == 0xFFF:
                    if pos + 4 > len(data):
                        break
                    size = struct.unpack_from('<I', data, pos)[0]
                    pos += 4
                if pos + size > len(data):
                    break

                if tag == 67:   # HWPTAG_PARA_TEXT
                    try:
                        t = data[pos:pos + size].decode('utf-16-le')
                        t = ''.join(c if c >= ' ' else ' ' for c in t).strip()
                        if t:
                            text_parts.append(t)
                    except Exception:
                        pass
                pos += size

        ole.close()
        return '\n'.join(text_parts) if text_parts else None
    except Exception:
        return None


def _hwpx_zipxml(filepath: str):
    """HWPX(ZIP+XML) 파일에서 텍스트 추출"""
    try:
        import zipfile
        import xml.etree.ElementTree as ET

        with zipfile.ZipFile(filepath, 'r') as z:
            sections = sorted(
                n for n in z.namelist()
                if re.search(r'[Ss]ection\d+\.xml', n)
            )
            text_parts = []
            for sname in sections:
                xml_bytes = z.read(sname)
                root = ET.fromstring(xml_bytes.decode('utf-8', errors='ignore'))
                for elem in root.iter():
                    if elem.text and elem.text.strip():
                        text_parts.append(elem.text.strip())
            return '\n'.join(text_parts) if text_parts else None
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════
#  설정
# ══════════════════════════════════════════════════════════════

class Config:
    DEFAULTS = {
        "search_field_x": None, "search_field_y": None,
        "result_first_x": None, "result_first_y": None,
        "search_delay": 1.2,
        "manual_confirm": False,
        "empty_pixel_rgb": None,
    }

    def __init__(self):
        self.data = dict(self.DEFAULTS)
        self._load()

    def _load(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.data.update(json.load(f))
        except Exception:
            pass

    def save(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def is_calibrated(self):
        return (self.data["search_field_x"] is not None
                and self.data["result_first_x"] is not None)


# ══════════════════════════════════════════════════════════════
#  위치 캡처 대화상자
# ══════════════════════════════════════════════════════════════

class CaptureDialog(tk.Toplevel):
    def __init__(self, parent, title, instruction, on_captured):
        super().__init__(parent)
        self.title("위치 캡처")
        self.geometry("420x220")
        self.resizable(False, False)
        self.grab_set()
        self.on_captured = on_captured

        tk.Label(self, text=title, font=("맑은 고딕", 11, "bold")).pack(pady=10)
        tk.Label(self, text=instruction, wraplength=370,
                 justify="center", fg="#555").pack(pady=4)
        self.status = tk.Label(self, text="", fg="blue")
        self.status.pack(pady=6)
        tk.Button(self, text="📍 3초 카운트다운 후 캡처",
                  command=self._start, bg="#FF9800", fg="white",
                  font=("맑은 고딕", 10), padx=8).pack(pady=6)
        tk.Button(self, text="취소", command=self.destroy, width=8).pack()

    def _start(self):
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        self.iconify()
        for i in range(3, 0, -1):
            self.after(0, lambda n=i: self.status.config(text=f"{n}초 후 캡처..."))
            time.sleep(1)
        x, y = pyautogui.position()
        self.after(0, lambda: self._done(x, y))

    def _done(self, x, y):
        self.deiconify()
        self.status.config(text=f"✓ 캡처 완료: ({x}, {y})", fg="green")
        self.on_captured(x, y)
        self.after(1200, self.destroy)


# ══════════════════════════════════════════════════════════════
#  메인 앱
# ══════════════════════════════════════════════════════════════

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("충북 소통메신저 자동 사용자 선택 v1.3")
        self.root.geometry("880x720")
        self.root.minsize(700, 560)

        self.config = Config()
        self.names_list: list = []   # [(org, name), ...]
        self.stop_flag = threading.Event()
        self.continue_event = threading.Event()
        self.continue_event.set()

        self._build_ui()
        self._refresh_calib_labels()
        self._check_deps()

    # ── UI ───────────────────────────────────────────────────

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        nb = ttk.Notebook(self.root)
        nb.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        f1 = ttk.Frame(nb); nb.add(f1, text="  1. 명단 입력  ")
        f2 = ttk.Frame(nb); nb.add(f2, text="  2. 위치 설정  ")
        f3 = ttk.Frame(nb); nb.add(f3, text="  3. 자동 선택  ")
        f4 = ttk.Frame(nb); nb.add(f4, text="  📖 사용 방법  ")

        self._tab_input(f1)
        self._tab_calib(f2)
        self._tab_auto(f3)
        self._tab_help(f4)

        self.status_var = tk.StringVar(value="준비")
        tk.Label(self.root, textvariable=self.status_var, bd=1,
                 relief="sunken", anchor="w", bg="#f0f0f0", fg="#333"
                 ).grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 4))

    # ── 탭1 ─────────────────────────────────────────────────

    def _tab_input(self, f):
        f.columnconfigure(0, weight=1)
        f.rowconfigure(2, weight=2)
        f.rowconfigure(5, weight=1)

        tk.Label(
            f,
            text="엑셀·한글 문서에서 명단을 복사(Ctrl+C)하여 아래에 붙여넣기 하거나,\n"
                 "파일을 직접 여세요.  소속기관명과 이름이 포함된 형식이면 됩니다.",
            justify="left", fg="#555", font=("맑은 고딕", 9)
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 2))

        bf = tk.Frame(f)
        bf.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=4)
        tk.Button(bf, text="엑셀 파일 열기 (.xlsx)",
                  command=self._open_excel, font=("맑은 고딕", 9), width=16).pack(side="left", padx=(0, 4))
        tk.Button(bf, text="HWP 파일 열기 (.hwp)",
                  command=self._open_hwp,  font=("맑은 고딕", 9), width=16).pack(side="left", padx=(0, 4))
        tk.Button(bf, text="명단 추출 →", command=self._parse,
                  bg="#2196F3", fg="white", font=("맑은 고딕", 10, "bold"),
                  width=10).pack(side="left", padx=(0, 4))
        tk.Button(bf, text="초기화", command=self._clear_input,
                  font=("맑은 고딕", 9), width=7).pack(side="left")

        self.input_text = scrolledtext.ScrolledText(f, font=("맑은 고딕", 9))
        self.input_text.grid(row=2, column=0, columnspan=2, sticky="nsew",
                             padx=10, pady=(0, 4))

        ttk.Separator(f, orient="horizontal").grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=10)

        rf = tk.Frame(f)
        rf.grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=4)
        tk.Label(rf, text="추출 결과:", font=("맑은 고딕", 9, "bold")).pack(side="left")
        self.parse_status = tk.Label(rf, text="", fg="#666", font=("맑은 고딕", 9))
        self.parse_status.pack(side="left", padx=6)

        self.parsed_list = tk.Listbox(f, font=("맑은 고딕", 9),
                                      selectmode="extended", height=8)
        sb = ttk.Scrollbar(f, orient="vertical", command=self.parsed_list.yview)
        self.parsed_list.config(yscrollcommand=sb.set)
        self.parsed_list.grid(row=5, column=0, sticky="nsew", padx=(10, 0), pady=(0, 8))
        sb.grid(row=5, column=1, sticky="ns", padx=(0, 10), pady=(0, 8))

        ctx = tk.Menu(self.root, tearoff=0)
        ctx.add_command(label="선택 항목 삭제", command=self._delete_selected)
        self.parsed_list.bind("<Button-3>",
                              lambda e: ctx.tk_popup(e.x_root, e.y_root))

    # ── 탭2 ─────────────────────────────────────────────────

    def _tab_calib(self, f):
        f.columnconfigure(0, weight=1)

        tk.Label(
            f,
            text="소통메신저 '사용자 선택' 창을 열어 둔 상태에서 아래 순서대로 위치를 설정하세요.\n"
                 "창은 항상 같은 위치에 두는 것이 좋습니다.",
            justify="center", fg="#555", font=("맑은 고딕", 9)
        ).grid(row=0, column=0, pady=(14, 6))

        for row, key, label, note in [
            (1, "search_field",
             "STEP 1 — 검색 입력창 위치",
             "소속+이름 검색 입력칸"),
            (2, "result_first",
             "STEP 2 — 검색 결과 첫 번째 항목 위치",
             "※ 먼저 임의 이름을 검색하여 결과를 띄운 후 설정하세요"),
        ]:
            lf = ttk.LabelFrame(f, text=label)
            lf.grid(row=row, column=0, sticky="ew", padx=24, pady=6)
            lf.columnconfigure(1, weight=1)
            tk.Label(lf, text=note, fg="#555",
                     font=("맑은 고딕", 9)).grid(row=0, column=0, padx=10, pady=8)
            lbl = tk.Label(lf, text="미설정", fg="red", width=20)
            lbl.grid(row=0, column=1, padx=4)
            setattr(self, f"lbl_{key}", lbl)
            tk.Button(
                lf, text="📍 위치 설정", bg="#FF9800", fg="white",
                command=lambda k=key, l=label: self._do_capture(k, l)
            ).grid(row=0, column=2, padx=10, pady=8)

        # STEP 3: 검색 결과 없음 픽셀 캡처 (선택 사항)
        lf3 = ttk.LabelFrame(f, text="STEP 3 — 결과 없음 감지 (선택)")
        lf3.grid(row=3, column=0, sticky="ew", padx=24, pady=6)
        lf3.columnconfigure(1, weight=1)
        tk.Label(
            lf3,
            text="소통메신저에서 결과가 없는 검색을 한 상태에서 캡처하세요.\n"
                 "검색 결과 없을 때 ✗ 표시 후 자동으로 다음으로 넘어갑니다.",
            fg="#555", font=("맑은 고딕", 9), justify="left"
        ).grid(row=0, column=0, padx=10, pady=8)
        self.lbl_empty_pixel = tk.Label(lf3, text="미설정 (선택)", fg="#888", width=20)
        self.lbl_empty_pixel.grid(row=0, column=1, padx=4)
        tk.Button(
            lf3, text="📍 픽셀 캡처", bg="#FF9800", fg="white",
            command=self._capture_empty_pixel
        ).grid(row=0, column=2, padx=10, pady=8)

        sg = ttk.LabelFrame(f, text="검색 설정")
        sg.grid(row=4, column=0, sticky="ew", padx=24, pady=6)

        tk.Label(sg, text="검색 후 대기 시간(초):",
                 font=("맑은 고딕", 9)).grid(row=0, column=0, padx=10, pady=8)
        self.delay_var = tk.DoubleVar(value=self.config.data["search_delay"])
        ttk.Spinbox(sg, from_=0.3, to=5.0, increment=0.1,
                    textvariable=self.delay_var, width=7
                    ).grid(row=0, column=1, padx=4, pady=8)
        tk.Label(sg, text="(느리면 값을 높이세요)",
                 fg="#888", font=("맑은 고딕", 8)).grid(row=0, column=2, padx=6)

        self.manual_var = tk.BooleanVar(
            value=self.config.data.get("manual_confirm", False))
        ttk.Checkbutton(
            sg,
            text="수동 확인 모드 — 검색 후 [계속] 버튼을 눌러야 다음으로 진행\n"
                 "(동명이인이 많을 때 사용: 검색 결과에서 사용자가 직접 더블클릭)",
            variable=self.manual_var
        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 8))

        tk.Button(f, text="✅  설정 저장", command=self._save_calib,
                  bg="#4CAF50", fg="white", font=("맑은 고딕", 10),
                  width=14).grid(row=5, column=0, pady=12)
        self.calib_msg = tk.Label(f, text="", fg="green", font=("맑은 고딕", 9))
        self.calib_msg.grid(row=6, column=0)

    # ── 탭3 ─────────────────────────────────────────────────

    def _tab_auto(self, f):
        f.columnconfigure(0, weight=1)
        f.rowconfigure(3, weight=1)

        tk.Label(
            f,
            text="소통메신저 '사용자 선택' 창 → 전체조직 탭을 열고 [시작] 버튼을 누르세요.\n"
                 "검색 형식: 소속+이름  (소속 없으면 이름만)\n"
                 "⚠  마우스를 화면 왼쪽 위 모서리로 이동하면 긴급 중지됩니다.",
            justify="center", fg="#555", font=("맑은 고딕", 9)
        ).grid(row=0, column=0, pady=(10, 6))

        bf = tk.Frame(f)
        bf.grid(row=1, column=0, pady=4)

        self.start_btn = tk.Button(
            bf, text="▶  자동 선택 시작", command=self._start,
            bg="#4CAF50", fg="white", font=("맑은 고딕", 12, "bold"), width=16)
        self.start_btn.pack(side="left", padx=6)

        self.continue_btn = tk.Button(
            bf, text="▶▶  계속", command=self._resume,
            bg="#2196F3", fg="white", font=("맑은 고딕", 11), width=8, state="disabled")
        self.continue_btn.pack(side="left", padx=6)

        self.stop_btn = tk.Button(
            bf, text="■  중지", command=self._stop,
            bg="#f44336", fg="white", font=("맑은 고딕", 11), width=8, state="disabled")
        self.stop_btn.pack(side="left", padx=6)

        tk.Label(f, text="진행 상황:", anchor="w",
                 font=("맑은 고딕", 9, "bold")).grid(
            row=2, column=0, sticky="w", padx=10)
        self.log = scrolledtext.ScrolledText(
            f, font=("맑은 고딕", 9), state="disabled", height=18)
        self.log.grid(row=3, column=0, sticky="nsew", padx=10, pady=4)

        self.progress = ttk.Progressbar(f, mode="determinate")
        self.progress.grid(row=4, column=0, sticky="ew", padx=10, pady=(0, 2))
        self.prog_label = tk.Label(f, text="0 / 0", font=("맑은 고딕", 9))
        self.prog_label.grid(row=5, column=0, pady=(0, 6))

    # ── 이벤트 ───────────────────────────────────────────────

    def _check_deps(self):
        missing = [p for p, m in [("pyautogui", pyautogui),
                                   ("pyperclip", pyperclip)] if m is None]
        if missing:
            messagebox.showerror(
                "패키지 누락",
                f"필수 패키지 미설치: {', '.join(missing)}\n\n"
                "시작.bat 을 실행하면 자동으로 설치됩니다.")

    def _open_excel(self):
        if openpyxl is None:
            messagebox.showerror("오류", "pip install openpyxl 필요")
            return
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            lines = []
            for row in ws.iter_rows(values_only=True):
                cells = [str(c).strip() if c is not None else "" for c in row]
                if any(cells):
                    lines.append("\t".join(cells))
            wb.close()
            self.input_text.delete("1.0", "end")
            self.input_text.insert("1.0", "\n".join(lines))
            self.status_var.set(f"엑셀 로드: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("오류", f"파일 읽기 실패:\n{e}")

    def _open_hwp(self):
        path = filedialog.askopenfilename(
            filetypes=[("HWP", "*.hwp *.hwpx"), ("All", "*.*")])
        if not path:
            return
        self.status_var.set("HWP 파일 읽는 중...")
        self.root.update()
        text = extract_hwp_text(path)
        if text:
            self.input_text.delete("1.0", "end")
            self.input_text.insert("1.0", text)
            self.status_var.set(f"HWP 로드: {os.path.basename(path)}")
        else:
            messagebox.showwarning(
                "HWP 읽기 실패",
                "HWP 파일을 자동으로 읽지 못했습니다.\n\n"
                "HWP에서 해당 표/목록을 직접 복사(Ctrl+C)하여\n"
                "텍스트 입력창에 붙여넣기 해주세요.")
            self.status_var.set("HWP 읽기 실패 — 직접 복사·붙여넣기 필요")

    def _parse(self):
        text = self.input_text.get("1.0", "end")
        pairs, skipped = parse_input(text)

        # 소속 없는 항목도 제외
        no_org_list = [p for p in pairs if not p[0]]
        valid = [p for p in pairs if p[0]]

        self.names_list = valid
        self.parsed_list.delete(0, "end")
        for i, (org, name) in enumerate(valid, 1):
            self.parsed_list.insert("end", f"   {i:>3}. [{org}]  {name}")

        color = "green" if valid else "red"
        msg = f"명단 추출 완료: {len(valid)}명"
        excluded = len(skipped) + len(no_org_list)
        if excluded:
            parts = []
            if skipped:
                parts.append(f"인식실패 {len(skipped)}")
            if no_org_list:
                parts.append(f"소속없음 {len(no_org_list)}")
            msg += f"  /  제외: {excluded}건 ({', '.join(parts)})"
        self.parse_status.config(text=msg, fg=color)
        self.status_var.set(msg)

    def _clear_input(self):
        self.input_text.delete("1.0", "end")
        self.parsed_list.delete(0, "end")
        self.names_list.clear()
        self.parse_status.config(text="")

    def _delete_selected(self):
        for i in reversed(self.parsed_list.curselection()):
            self.parsed_list.delete(i)
            del self.names_list[i]
        self.parse_status.config(
            text=f"명단 추출 완료: {len(self.names_list)}명", fg="green")

    def _do_capture(self, key: str, label: str):
        def on_captured(x, y):
            self.config.data[f"{key}_x"] = x
            self.config.data[f"{key}_y"] = y
            self._refresh_calib_labels()
        CaptureDialog(self.root, label,
                      "3초 카운트다운 후 마우스를 해당 위치로 이동하세요.",
                      on_captured)

    def _save_calib(self):
        self.config.data["search_delay"] = round(self.delay_var.get(), 1)
        self.config.data["manual_confirm"] = self.manual_var.get()
        self.config.save()
        self.calib_msg.config(text="✅ 설정 저장 완료")
        self.root.after(3000, lambda: self.calib_msg.config(text=""))

    def _capture_empty_pixel(self):
        rx = self.config.data.get("result_first_x")
        ry = self.config.data.get("result_first_y")
        if rx is None or ry is None:
            messagebox.showwarning("알림", "먼저 STEP 2에서 결과 위치를 설정하세요.")
            return
        if pyautogui is None:
            messagebox.showerror("오류", "pyautogui 미설치")
            return

        def run():
            self.root.iconify()
            for i in range(3, 0, -1):
                self.root.after(0, lambda n=i: self.calib_msg.config(
                    text=f"결과 없음 픽셀 캡처 {n}초 후..."))
                time.sleep(1)
            pixel = pyautogui.pixel(rx, ry)
            self.config.data["empty_pixel_rgb"] = list(pixel[:3])
            self.config.save()
            self.root.after(0, self._on_empty_pixel_captured, pixel)
            self.root.deiconify()

        threading.Thread(target=run, daemon=True).start()

    def _on_empty_pixel_captured(self, pixel):
        r, g, b = pixel[:3]
        self.calib_msg.config(text=f"✅ 결과 없음 픽셀 저장: RGB({r},{g},{b})")
        self._refresh_calib_labels()
        self.root.after(3000, lambda: self.calib_msg.config(text=""))

    def _refresh_calib_labels(self):
        for key in ("search_field", "result_first"):
            lbl = getattr(self, f"lbl_{key}", None)
            if not lbl:
                continue
            x = self.config.data.get(f"{key}_x")
            y = self.config.data.get(f"{key}_y")
            if x is not None:
                lbl.config(text=f"✓ ({x}, {y})", fg="green")
            else:
                lbl.config(text="미설정", fg="red")
        lbl3 = getattr(self, "lbl_empty_pixel", None)
        if lbl3:
            rgb = self.config.data.get("empty_pixel_rgb")
            if rgb:
                lbl3.config(text=f"✓ RGB{tuple(rgb)}", fg="green")
            else:
                lbl3.config(text="미설정 (선택)", fg="#888")

    # ── 자동화 ───────────────────────────────────────────────

    def _start(self):
        if pyautogui is None or pyperclip is None:
            messagebox.showerror("오류", "필수 패키지 미설치")
            return
        if not self.names_list:
            messagebox.showwarning("알림", "먼저 명단을 추출해 주세요.")
            return
        if not self.config.is_calibrated():
            messagebox.showwarning("알림", "위치 설정 탭에서 검색창·결과 위치를 먼저 설정하세요.")
            return

        self.stop_flag.clear()
        self.continue_event.set()
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.continue_btn.config(state="disabled")
        self.progress.config(maximum=len(self.names_list), value=0)
        self.prog_label.config(text=f"0 / {len(self.names_list)}")
        self._log_clear()
        self._log(f"자동 선택 시작 — 총 {len(self.names_list)}명\n\n")
        threading.Thread(target=self._worker, daemon=True).start()

    def _stop(self):
        self.stop_flag.set()
        self.continue_event.set()
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.continue_btn.config(state="disabled")
        self._log("\n⏹  중지\n")

    def _resume(self):
        self.continue_event.set()
        self.continue_btn.config(state="disabled")

    def _worker(self):
        total = len(self.names_list)
        ok = fail = 0
        manual = self.config.data.get("manual_confirm", False)

        for i, (org, name) in enumerate(self.names_list):
            if self.stop_flag.is_set():
                break

            self._update_progress(i, total)
            org_s = org.strip()
            search_str = f"{org_s}+{name}" if org_s else name
            self._log(f"[{i+1:>3}/{total}]  {search_str}  ... ")

            try:
                # 검색 (소속+이름 또는 이름만)
                self._do_search(org_s, name)

                if manual:
                    # 수동 모드: 검색 후 사용자가 직접 더블클릭
                    self.continue_event.clear()
                    self.root.after(0, self._show_continue, search_str)
                    self.continue_event.wait()
                    if self.stop_flag.is_set():
                        break
                    ok += 1
                    self._log("✓\n")
                elif not self._has_result():
                    # 자동 모드: 검색 결과 없음
                    fail += 1
                    self._log("✗  (검색 결과 없음)\n")
                else:
                    # 자동 모드: 첫 번째 결과 더블클릭
                    self._do_select()
                    ok += 1
                    self._log("✓\n")

            except pyautogui.FailSafeException:
                self._log("\n⚠  긴급 중지 (화면 모서리)\n")
                self.stop_flag.set()
                break
            except Exception as e:
                fail += 1
                self._log(f"✗  ({e})\n")

            time.sleep(0.2)

        self._update_progress(ok + fail, total)
        self.root.after(0, self._done, ok, fail)

    def _do_search(self, org: str, name: str):
        sx = self.config.data["search_field_x"]
        sy = self.config.data["search_field_y"]
        delay = self.config.data["search_delay"]

        search_text = f"{org}+{name}" if org else name

        pyautogui.click(sx, sy)
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        pyperclip.copy(search_text)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.15)
        pyautogui.press("enter")
        time.sleep(delay)

    def _do_select(self):
        rx = self.config.data["result_first_x"]
        ry = self.config.data["result_first_y"]
        pyautogui.doubleClick(rx, ry)
        time.sleep(0.3)

    def _has_result(self) -> bool:
        """검색 결과 위치 픽셀 색상으로 결과 유무 판단. 미설정 시 항상 True."""
        empty_rgb = self.config.data.get("empty_pixel_rgb")
        if empty_rgb is None:
            return True
        rx = self.config.data.get("result_first_x")
        ry = self.config.data.get("result_first_y")
        if rx is None or ry is None:
            return True
        try:
            pixel = pyautogui.pixel(rx, ry)
            r, g, b = empty_rgb
            tol = 15
            return not (abs(pixel[0] - r) <= tol and
                        abs(pixel[1] - g) <= tol and
                        abs(pixel[2] - b) <= tol)
        except Exception:
            return True

    def _show_continue(self, search_str: str):
        self.status_var.set(
            f"수동 선택 대기: {search_str}  →  소통메신저에서 더블클릭 후 [계속] 버튼")
        self.continue_btn.config(state="normal")

    def _done(self, ok: int, fail: int):
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.continue_btn.config(state="disabled")
        self._log(f"\n{'─'*40}\n완료  ✓ {ok}명   ✗ {fail}명\n")
        self.status_var.set(f"완료 — 성공: {ok}명, 실패: {fail}명")

    def _update_progress(self, idx: int, total: int):
        self.root.after(0, lambda: [
            self.progress.config(value=idx),
            self.prog_label.config(text=f"{idx} / {total}")
        ])

    def _log(self, msg: str):
        self.root.after(0, self._log_append, msg)

    def _log_append(self, msg: str):
        self.log.config(state="normal")
        self.log.insert("end", msg)
        self.log.see("end")
        self.log.config(state="disabled")

    def _log_clear(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    # ── 도움말 탭 ────────────────────────────────────────────

    def _tab_help(self, f):
        f.columnconfigure(0, weight=1)
        f.rowconfigure(0, weight=1)

        txt = scrolledtext.ScrolledText(
            f, font=("맑은 고딕", 9), wrap="word",
            state="normal", bg="#fafafa", relief="flat"
        )
        txt.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)

        HELP_TEXT = """\
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  충북 소통메신저  자동 사용자 선택  v1.3
  — 사용 방법 (처음 사용자도 따라할 수 있도록 작성되었습니다) —
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 이 프로그램이 하는 일
──────────────────────────────────────────────────────
  엑셀·한글 문서의 명단(소속기관 + 이름)을 붙여 넣으면
  충북 소통메신저에서 자동으로 이름을 검색하고 선택해 줍니다.
  수십~수백 명을 일일이 검색하는 반복 작업을 자동화합니다.


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 시작 전 준비 사항
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  1. 충북 소통메신저에 로그인합니다.
  2. 메시지 보내기 화면에서 [사용자 선택] 버튼을 클릭합니다.
  3. 사용자 선택 창이 열리면, 상단 탭에서 [전체조직]을 선택합니다.
  4. 소통메신저 창과 이 프로그램 창을 화면에 나란히 배치하면 편리합니다.

  ※ 프로그램 실행 중에는 마우스를 움직이지 마세요.
     (긴급 중지: 마우스를 화면 왼쪽 위 모서리로 빠르게 이동)


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ STEP 1 — 명단 입력  (탭: 1. 명단 입력)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  [ 방법 A ]  엑셀 파일 이용
    ① [엑셀 파일 열기] 버튼을 클릭합니다.
    ② 소속기관과 이름이 있는 .xlsx 파일을 선택합니다.
    ③ [명단 추출 →] 버튼을 클릭합니다.

  [ 방법 B ]  한글(HWP) 파일 이용
    ① [HWP 파일 열기] 버튼을 클릭합니다.
    ② .hwp 또는 .hwpx 파일을 선택합니다.
    ③ [명단 추출 →] 버튼을 클릭합니다.

  [ 방법 C ]  직접 복사·붙여넣기
    ① 엑셀이나 한글 문서에서 명단 표를 마우스로 드래그하여 선택합니다.
    ② Ctrl+C 로 복사합니다.
    ③ 이 프로그램의 텍스트 입력창 안을 클릭합니다.
    ④ Ctrl+V 로 붙여넣습니다.
    ⑤ [명단 추출 →] 버튼을 클릭합니다.

  ▶ 인식 가능한 형식 (모두 자동 처리)
    · 소속기관과 이름이 같은 셀/행:  "가경초등학교  홍길동"
    · 소속기관과 이름이 다른 행:     소속행 → 이름행 순서로 자동 연결
    · 탭으로 구분된 엑셀 표 복붙:    소속기관[탭]이름

  ▶ 충북교육청 소속기관명 자동 변환 (DB 내장)
    · 백곡초등학교  →  백곡초
    · 청주교육지원청  →  청주교육지원청  (그대로 유지)
    · 오타 자동 보정 (예: 백곡쵸등학교 → 백곡초)

  ▶ 명단 추출 결과 확인
    · 하단 목록에 [소속기관]  이름 형식으로 표시됩니다.
    · 소속기관이 없거나 인식에 실패한 항목은 자동으로 제외됩니다.
    · 특정 항목을 마우스 오른쪽 버튼으로 클릭하면 삭제할 수 있습니다.


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ STEP 2 — 위치 설정  (탭: 2. 위치 설정)  ※ 최초 1회만 설정
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  소통메신저 어디를 클릭해야 하는지 프로그램에 알려주는 과정입니다.
  처음 한 번만 설정하면 이후로는 자동으로 기억합니다.

  ┌─────────────────────────────────────────────────────────┐
  │ STEP 1  검색 입력창 위치                                  │
  │   ① [📍 위치 설정] 버튼을 클릭합니다.                     │
  │   ② 창이 최소화됩니다. 3초 안에 소통메신저 검색 입력칸      │
  │      위에 마우스를 올려두세요.                              │
  │   ③ 자동으로 좌표가 저장됩니다.                            │
  └─────────────────────────────────────────────────────────┘
  ┌─────────────────────────────────────────────────────────┐
  │ STEP 2  검색 결과 첫 번째 항목 위치                        │
  │   ① 소통메신저 검색창에 임의 이름(예: 홍길동)을 검색합니다. │
  │   ② 결과 목록의 첫 번째 줄이 보이는 상태에서               │
  │      [📍 위치 설정] 버튼을 클릭합니다.                     │
  │   ③ 3초 안에 결과 목록의 첫 번째 항목 위에 마우스를 올리세요.│
  └─────────────────────────────────────────────────────────┘
  ┌─────────────────────────────────────────────────────────┐
  │ STEP 3  결과 없음 감지  (선택 사항, 권장)                  │
  │   ① 소통메신저에서 존재하지 않는 이름(예: zzzzz)을 검색합니다.│
  │   ② 결과가 비어 있는 상태에서 [📍 픽셀 캡처] 버튼 클릭.   │
  │   ③ 이후 검색 결과가 없으면 자동으로 ✗ 표시 후 건너뜁니다. │
  └─────────────────────────────────────────────────────────┘
  ┌─────────────────────────────────────────────────────────┐
  │ 검색 설정                                                 │
  │   · 검색 후 대기 시간: 인터넷이 느리면 1.5~2.0초로 높이세요 │
  │   · 수동 확인 모드: 동명이인이 많을 때 체크합니다.           │
  │     (검색 후 [▶▶ 계속] 버튼을 눌러 하나씩 진행)            │
  └─────────────────────────────────────────────────────────┘

  ★ 반드시 [✅ 설정 저장] 버튼을 눌러 저장하세요!


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ STEP 3 — 자동 선택 실행  (탭: 3. 자동 선택)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  ① 소통메신저 [사용자 선택] 창을 열고 [전체조직] 탭을 선택합니다.
  ② 이 프로그램에서 [3. 자동 선택] 탭을 클릭합니다.
  ③ [▶ 자동 선택 시작] 버튼을 클릭합니다.
  ④ 프로그램이 자동으로 이름을 검색하고 선택합니다.
     진행 상황은 화면 아래 로그창에서 확인할 수 있습니다.
       ✓  → 선택 완료
       ✗  → 검색 결과 없음 (등록되지 않은 사용자)
  ⑤ 모두 완료되면 완료 메시지가 표시됩니다.

  ⚠ 긴급 중지 방법
    · 마우스를 화면의 왼쪽 위 모서리(0, 0)로 빠르게 이동합니다.
    · 또는 [■ 중지] 버튼을 클릭합니다.


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 자주 묻는 질문 (FAQ)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Q. 명단 추출 결과에 '소속없음'으로 제외된 항목이 있어요.
  A. 소속기관명을 인식하지 못한 경우입니다.
     텍스트 입력창에서 해당 줄을 직접 수정하거나,
     소속기관[탭]이름 형식으로 입력 후 다시 명단 추출해 보세요.

  Q. 프로그램이 엉뚱한 위치를 클릭해요.
  A. 소통메신저 창의 위치나 크기가 바뀌었을 수 있습니다.
     [2. 위치 설정] 탭에서 위치를 다시 설정하고 저장하세요.

  Q. 검색은 됐는데 빈 화면이 나와요.
  A. 소통메신저에 해당 사용자가 등록되지 않은 경우입니다.
     STEP 3(결과 없음 감지)를 설정하면 자동으로 ✗ 처리합니다.

  Q. 속도가 너무 빠르거나 느려요.
  A. [2. 위치 설정] 탭의 '검색 후 대기 시간'을 조절하세요.
     느린 PC나 인터넷: 1.5~2.5초 / 빠른 환경: 0.8~1.2초

  Q. 동명이인이 있어서 잘못 선택될까 봐 걱정돼요.
  A. '수동 확인 모드'를 체크하세요.
     검색 후 [▶▶ 계속] 버튼이 활성화되면, 소통메신저 결과에서
     직접 더블클릭하여 선택한 뒤 [계속]을 누릅니다.

  Q. HWP 파일을 열었는데 내용이 안 나와요.
  A. 한/글 프로그램이 설치되지 않은 경우 일부 파일이 열리지 않을 수 있습니다.
     한글 문서를 직접 열어 해당 표를 Ctrl+C 로 복사 후
     텍스트 입력창에 Ctrl+V 로 붙여넣어 사용하세요.


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 개발자 정보
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Developed by  송동석
  Teacher  |  Data Analytics  |  App Developer

  협업 및 피드백:  dungst.me@gmail.com

  충청북도교육청 소통메신저 자동화 도구  |  버전 v1.3
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
        txt.insert("1.0", HELP_TEXT)
        txt.config(state="disabled")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = App()
    app.run()
