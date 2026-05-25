"""Parsing and organization lookup helpers for SotongPick."""

import difflib
import json
import os
import re
import sys


def _resource_path(name: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(__file__))
    return os.path.join(base, name)


def load_org_lookup(path: str | None = None) -> dict[str, str]:
    db_path = path or _resource_path("org_db.json")
    with open(db_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return {str(k): str(v) for k, v in data.items()}


_ORG_LOOKUP: dict[str, str] = load_org_lookup()


ORG_END_RE = re.compile(
    r'(학교|초등학교|중학교|고등학교|특수학교|유치원|교육청|교육지원청|'
    r'교육원|교육관|연구원|연구소|도청|시청|군청|구청|교육부|본청|센터|지원청)$'
)
SCHOOL_ABBR_RE = re.compile(r'^[가-힣]{1,6}(초|중|고)$')
TIMESTAMP_RE = re.compile(r'^\[\d{4}[-/]\d{2}[-/]\d{2}[\d\s:.,\-]*\]\s*')
JUNK_RE = re.compile(r'^\d+$|\d{2,4}-\d{3,4}-\d{4}|^https?://|[!@#$%^*()\[\]{}<>|\\/?~`]')


def abbreviate_school(s: str) -> str:
    """풀 학교명 → 검색용 약칭 (백곡초등학교 → 백곡초)"""
    s = s.strip()
    for suffix, short in [("초등학교", "초"), ("중학교", "중"), ("고등학교", "고")]:
        if s.endswith(suffix) and len(s) > len(suffix):
            return s[:-len(suffix)] + short
    return s


def is_org(s: str) -> bool:
    s = s.strip()
    if not s:
        return False
    if ORG_END_RE.search(s):
        return True
    if SCHOOL_ABBR_RE.match(s):
        return True
    if s in _ORG_LOOKUP:
        return True
    abbr = abbreviate_school(s)
    return abbr in _ORG_LOOKUP


def is_person_name(s: str) -> bool:
    """2~4자 한글, 학교명 아닌 것"""
    s = s.strip()
    if not re.match(r'^[가-힣]{2,4}$', s):
        return False
    return not is_org(s)


def _split_school_suffix(s: str):
    """학교명에서 접두어·종류 분리. '진천상신초등학교' → ('진천상신', '초')"""
    for suffix, short in [("초등학교", "초"), ("중학교", "중"), ("고등학교", "고")]:
        if s.endswith(suffix) and len(s) > len(suffix):
            return s[:-len(suffix)], short
    for short in ["초", "중", "고"]:
        if s.endswith(short) and len(s) > 1:
            return s[:-len(short)], short
    return s, ""


def lookup_org(s: str):
    """소속명 → DB 기준 표준 약칭. 오타 자동 보정."""
    s = s.strip()
    if not s:
        return None
    if s in _ORG_LOOKUP:
        return _ORG_LOOKUP[s]
    abbr = abbreviate_school(s)
    if abbr in _ORG_LOOKUP:
        return _ORG_LOOKUP[abbr]
    keys = list(_ORG_LOOKUP.keys())
    prefix, suffix = _split_school_suffix(s)
    if suffix:
        # Exact prefix+type match first
        for k, v in _ORG_LOOKUP.items():
            kpre, ksuf = _split_school_suffix(k)
            if kpre == prefix and ksuf == suffix:
                return v
        # Type-aware fuzzy: only match keys with same school type to prevent
        # 초/중/고 names from resolving to 유치원 entries (e.g. 오송솔미초 → 오송솔미초병설유치원)
        type_keys = [k for k in keys if _split_school_suffix(k)[1] == suffix]
        matches = difflib.get_close_matches(s, type_keys, n=1, cutoff=0.75)
        if not matches and len(abbr) > 1:
            matches = difflib.get_close_matches(abbr, type_keys, n=1, cutoff=0.75)
    else:
        matches = difflib.get_close_matches(s, keys, n=1, cutoff=0.75)
        if not matches and len(abbr) > 1:
            matches = difflib.get_close_matches(abbr, keys, n=1, cutoff=0.75)
    if matches:
        return _ORG_LOOKUP[matches[0]]
    return None


def best_org_from(tokens: list) -> str:
    """토큰 목록에서 소속명 추출"""
    for tok in tokens:
        if is_org(tok):
            result = lookup_org(tok)
            if result:
                return result
    for tok in tokens:
        if re.match(r'^[가-힣]{2,}$', tok) and not is_person_name(tok):
            result = lookup_org(tok)
            if result:
                return result
    return ''


def extract_pair(tokens: list):
    """토큰 목록 → (소속, 이름) 또는 None"""
    tokens = [t for t in tokens if not JUNK_RE.search(t)]
    name_idx = None
    for i, tok in enumerate(tokens):
        if is_person_name(tok):
            name_idx = i
            break
    if name_idx is None:
        return None
    other = [t for i, t in enumerate(tokens) if i != name_idx]
    org = best_org_from(other) if other else ''
    return (org, tokens[name_idx])


def _tokenize(line: str) -> list:
    """줄 → 정제된 토큰 목록"""
    line = line.strip().replace(' : ', '\t')
    parts = line.split('\t')
    tokens = [t.strip() for t in parts if t.strip()]
    if len(tokens) <= 1:
        tokens = [t.strip() for t in line.split(' ') if t.strip()]
    return [t for t in tokens if not JUNK_RE.search(t)]


def parse_input(text: str) -> list:
    """
    스마트 파서: 표에서 복붙할 때 소속/이름이 별도 행으로 오는 경우를
    pending_org 버퍼로 연결.

    지원 패턴:
    ① 탭 구분 (Excel 표 복붙)       : 소속      이름
    ② 행 교대 (HWP 표 복붙)         : 소속\n이름\n소속\n이름
    ③ 한 행에 소속+이름 (공백 구분)  : 백곡초등학교 김순범
    ④ 타임스탬프 채팅 로그           : [2026-04-02 ...] 이름 소속
    """
    results = []
    pending_org = ''
    for raw_line in text.strip().splitlines():
        line = TIMESTAMP_RE.sub('', raw_line).strip()
        line = re.sub(r'^\d+[.)]\s*', '', line)
        if not line:
            continue
        tokens = _tokenize(line)
        if not tokens:
            continue
        if len(tokens) == 1:
            tok = tokens[0]
            if is_person_name(tok):
                results.append({'org': pending_org, 'name': tok})
                pending_org = ''
                continue
            if is_org(tok):
                resolved = lookup_org(tok)
                pending_org = resolved if resolved else tok
                continue
        if all(is_person_name(t) for t in tokens):
            for t in tokens:
                results.append({'org': pending_org, 'name': t})
            pending_org = ''
            continue
        pair = extract_pair(tokens)
        if pair:
            org, name = pair
            if not org and pending_org:
                org = pending_org
            pending_org = ''
            results.append({'org': org, 'name': name})
            continue
        org = best_org_from(tokens)
        names = [t for t in tokens if is_person_name(t)]
        if org and names:
            for n in names:
                results.append({'org': org, 'name': n})
            pending_org = ''
        elif org:
            pending_org = org
        elif names:
            for n in names:
                results.append({'org': pending_org, 'name': n})
    return results
