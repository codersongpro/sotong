"""충북 소통메신저 자동 사용자 선택 v1.5"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import time
import json
import os
import re
import subprocess
import sys
import difflib

try:
    import pyautogui
    pyautogui.PAUSE = 0.3
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

# ─────────────────────────────────────────────
#  설정 파일 경로
# ─────────────────────────────────────────────
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".chungbuk_auto_config.json")

# ─────────────────────────────────────────────
#  정규식
# ─────────────────────────────────────────────
ORG_END_RE = re.compile(
    r'(학교|초등학교|중학교|고등학교|특수학교|유치원|교육청|교육지원청|'
    r'교육원|교육관|연구원|연구소|도청|시청|군청|구청|교육부|본청|센터|지원청)$'
)
SCHOOL_ABBR_RE = re.compile(r'^[가-힣]{1,6}(초|중|고)$')
TIMESTAMP_RE = re.compile(r'^\[\d{4}[-/]\d{2}[-/]\d{2}[\d\s:.,\-]*\]\s*')
JUNK_RE = re.compile(r'^\d+$|\d{2,4}-\d{3,4}-\d{4}|^https?://|[!@#$%^*()\[\]{}<>|\\/?~`]')

# ─────────────────────────────────────────────
#  소속기관 DB  (입력명 → 소통메신저 검색명)
# ─────────────────────────────────────────────
_ORG_LOOKUP: dict = dict((
    # 교육청 본청·지원청
    ("충청북도교육청", "충북교육청"),
    ("충북교육청", "충북교육청"),
    ("충청북도교육청본청", "충북교육청"),
    ("청주교육지원청", "청주교육지원청"),
    ("충주교육지원청", "충주교육지원청"),
    ("제천교육지원청", "제천교육지원청"),
    ("보은교육지원청", "보은교육지원청"),
    ("옥천교육지원청", "옥천교육지원청"),
    ("영동교육지원청", "영동교육지원청"),
    ("증평교육지원청", "증평교육지원청"),
    ("진천교육지원청", "진천교육지원청"),
    ("괴산증평교육지원청", "괴산증평교육지원청"),
    ("괴산교육지원청", "괴산증평교육지원청"),
    ("음성교육지원청", "음성교육지원청"),
    ("단양교육지원청", "단양교육지원청"),
    # 직속기관
    ("충청북도자연과학교육원", "충북자연과학교육원"),
    ("충북자연과학교육원", "충북자연과학교육원"),
    ("충청북도단재교육연수원", "단재교육연수원"),
    ("단재교육연수원", "단재교육연수원"),
    ("충청북도교육도서관", "충북교육도서관"),
    ("충북교육도서관", "충북교육도서관"),
    ("충청북도교육문화원", "충북교육문화원"),
    ("충북교육문화원", "충북교육문화원"),
    ("충청북도학생수련원", "충북학생수련원"),
    ("충북학생수련원", "충북학생수련원"),
    ("충청북도국제교육원", "충북국제교육원"),
    ("충북국제교육원", "충북국제교육원"),
    ("충청북도교육연구정보원", "도교육연구정보원"),
    ("도교육연구정보원", "도교육연구정보원"),
    ("충청북도중원교육문화원", "충북중원교육문화원"),
    ("충북중원교육문화원", "충북중원교육문화원"),
    ("충청북도해양교육원", "충북해양교육원"),
    ("충북해양교육원", "충북해양교육원"),
    ("충청북도유아교육진흥원", "충북유아교육진흥원"),
    ("충북유아교육진흥원", "충북유아교육진흥원"),
    ("충청북도진로교육원", "충북진로교육원"),
    ("충북진로교육원", "충북진로교육원"),
    ("충청북도특수교육원", "충북특수교육원"),
    ("충북특수교육원", "충북특수교육원"),
    ("충청북도학생교육원", "충북학생교육원"),
    ("충북학생교육원", "충북학생교육원"),
    ("충청북도학생교육문화원", "충북학생교육문화원"),
    ("충북학생교육문화원", "충북학생교육문화원"),
    ("충청북도인재개발원", "충북인재개발원"),
    ("충북인재개발원", "충북인재개발원"),
    ("충청북도교육청인재개발원", "충북인재개발원"),
    # ── 고등학교 ─────────────────────────────
    ("가경고등학교", "가경고"), ("괴산고등학교", "괴산고"),
    ("국원고등학교", "국원고"), ("금천고등학교", "금천고"),
    ("광혜원고등학교", "광혜원고"), ("단양고등학교", "단양고"),
    ("대금고등학교", "대금고"), ("대성여자상업고등학교", "대성여자상업고"),
    ("동성고등학교", "동성고"), ("매괴고등학교", "매괴고"),
    ("보은고등학교", "보은고"), ("보은여자고등학교", "보은여고"),
    ("봉명고등학교", "봉명고"), ("산남고등학교", "산남고"),
    ("상당고등학교", "상당고"), ("서전고등학교", "서전고"),
    ("서원고등학교", "서원고"), ("세광고등학교", "세광고"),
    ("세명고등학교", "세명고"), ("세명컴퓨터고등학교", "세명컴퓨터고"),
    ("양업고등학교", "양업고"), ("양청고등학교", "양청고"),
    ("영동고등학교", "영동고"), ("영동미래고등학교", "영동미래고"),
    ("영동산업과학고등학교", "영동산업과학고"),
    ("오송고등학교", "오송고"), ("오창고등학교", "오창고"),
    ("옥천고등학교", "옥천고"), ("운호고등학교", "운호고"),
    ("음성고등학교", "음성고"), ("일신여자고등학교", "일신여고"),
    ("제천고등학교", "제천고"), ("제천디지털전자고등학교", "제천디지털전자고"),
    ("제천산업고등학교", "제천산업고"), ("제천상업고등학교", "제천상업고"),
    ("제천여자고등학교", "제천여고"), ("제천제일고등학교", "제천제일고"),
    ("주성고등학교", "주성고"), ("증평공업고등학교", "증평공업고"),
    ("증평정보고등학교", "증평정보고"), ("중앙탑고등학교", "중앙탑고"),
    ("진천고등학교", "진천고"), ("진천상업고등학교", "진천상업고"),
    ("청산고등학교", "청산고"), ("청석고등학교", "청석고"),
    ("청원고등학교", "청원고"), ("청주고등학교", "청주고"),
    ("청주공업고등학교", "청주공업고"), ("청주농업고등학교", "청주농업고"),
    ("청주대성고등학교", "청주대성고"), ("청주신흥고등학교", "청주신흥고"),
    ("청주여자고등학교", "청주여고"), ("청주여고", "청주여고"),
    ("청주여자상업고등학교", "청주여자상업고"),
    ("청주외국어고등학교", "청주외국어고"),
    ("청주IT과학고등학교", "청주IT과학고"),
    ("청주중앙여자고등학교", "청주중앙여고"),
    ("청주하이텍고등학교", "청주하이텍고"),
    ("충북고등학교", "충북고"), ("충북고", "충북고"),
    ("충북과학고등학교", "충북과학고"), ("충북공업고등학교", "충북공업고"),
    ("충북반도체고등학교", "충북반도체고"), ("충북비즈니스고등학교", "충북비즈니스고"),
    ("충북산업과학고등학교", "충북산업과학고"),
    ("충북상업정보고등학교", "충북상업정보고"),
    ("충북생명산업고등학교", "충북생명산업고"),
    ("충북에너지고등학교", "충북에너지고"),
    ("충북여자고등학교", "충북여고"), ("충북외국어고등학교", "충북외국어고"),
    ("충북예술고등학교", "충북예술고"), ("충북체육고등학교", "충북체육고"),
    ("충주고등학교", "충주고"), ("충주공업고등학교", "충주공고"),
    ("충주공고", "충주공고"), ("충주대원고등학교", "충주대원고"),
    ("충주상업고등학교", "충주상업고"),
    ("충주예성여자고등학교", "충주예성여고"),
    ("충주여자고등학교", "충주여고"), ("충주여고", "충주여고"),
    ("충주중산고등학교", "충주중산고"), ("충원고등학교", "충원고"),
    ("학산고등학교", "학산고"),
    ("한국바이오마이스터고등학교", "한국바이오마이스터고"),
    ("한국교원대학교부설고등학교", "한국교원대부설고"),
    ("한국호텔관광고등학교", "한국호텔관광고"),
    ("한림디자인고등학교", "한림디자인고"),
    ("형석고등학교", "형석고"), ("황간고등학교", "황간고"),
    ("흥덕고등학교", "흥덕고"),
    ("충북대학교사범대학부설고등학교", "충북대사범대부설고"),
    ("경덕고등학교", "경덕고"), ("교동고등학교", "교동고"),
    ("내덕고등학교", "내덕고"), ("대성고등학교", "대성고"),
    ("동주고등학교", "동주고"), ("북일고등학교", "북일고"),
    ("성화고등학교", "성화고"), ("수곡고등학교", "수곡고"),
    ("신성고등학교", "신성고"), ("운천고등학교", "운천고"),
    ("율량고등학교", "율량고"), ("청남고등학교", "청남고"),
    ("충일고등학교", "충일고"), ("팔봉고등학교", "팔봉고"),
    ("현도고등학교", "현도고"), ("달천고등학교", "달천고"),
    ("봉방고등학교", "봉방고"), ("성남고등학교", "성남고"),
    ("성내고등학교", "성내고"), ("연수고등학교", "연수고"),
    ("은여울고등학교", "은여울고"), ("의림여자고등학교", "의림여고"),
    ("제천비즈니스고등학교", "제천비즈니스고"),
    ("금왕고등학교", "금왕고"), ("대소고등학교", "대소고"),
    ("증평고등학교", "증평고"), ("진천상공고등학교", "진천상공고"),
    ("칠성고등학교", "칠성고"),
    # ── 중학교 ─────────────────────────────
    ("가경중학교", "가경중"), ("각리중학교", "각리중"),
    ("감곡중학교", "감곡중"), ("강서중학교", "강서중"),
    ("경덕중학교", "경덕중"), ("괴산중학교", "괴산중"),
    ("괴산오성중학교", "괴산오성중"), ("괴산중앙중학교", "괴산중앙중"),
    ("교동중학교", "교동중"), ("구룡중학교", "구룡중"),
    ("국원중학교", "국원중"), ("금가중학교", "금가중"),
    ("금천중학교", "금천중"), ("남성중학교", "남성중"),
    ("남이중학교", "남이중"), ("내수중학교", "내수중"),
    ("내토중학교", "내토중"), ("노은중학교", "노은중"),
    ("단양중학교", "단양중"), ("단성중학교", "단성중"),
    ("단양소백산중학교", "단양소백산중"), ("대금중학교", "대금중"),
    ("대성중학교", "대성중"), ("대성여자중학교", "대성여중"),
    ("대소중학교", "대소중"), ("덕산중학교", "덕산중"),
    ("동성중학교", "동성중"), ("매괴중학교", "매괴중"),
    ("매포중학교", "매포중"), ("무극중학교", "무극중"),
    ("문의중학교", "문의중"), ("미덕중학교", "미덕중"),
    ("보덕중학교", "보덕중"), ("보은중학교", "보은중"),
    ("보은여자중학교", "보은여중"), ("봉명중학교", "봉명중"),
    ("부강중학교", "부강중"), ("산남중학교", "산남중"),
    ("살미중학교", "살미중"), ("상당중학교", "상당중"),
    ("서원중학교", "서원중"), ("서현중학교", "서현중"),
    ("세광중학교", "세광중"), ("세명중학교", "세명중"),
    ("성화중학교", "성화중"), ("소이중학교", "소이중"),
    ("솔밭중학교", "솔밭중"), ("송절중학교", "송절중"),
    ("수곡중학교", "수곡중"), ("수안보중학교", "수안보중"),
    ("속리산중학교", "속리산중"), ("신니중학교", "신니중"),
    ("안내중학교", "안내중"), ("안남중학교", "안남중"),
    ("앙성중학교", "앙성중"), ("양청중학교", "양청중"),
    ("엄정중학교", "엄정중"), ("영동중학교", "영동중"),
    ("영동정수중학교", "영동정수중"), ("오송중학교", "오송중"),
    ("오창중학교", "오창중"), ("옥산중학교", "옥산중"),
    ("옥천중학교", "옥천중"), ("옥천여자중학교", "옥천여중"),
    ("용문중학교", "용문중"), ("용성중학교", "용성중"),
    ("용암중학교", "용암중"), ("용아중학교", "용아중"),
    ("운동중학교", "운동중"), ("운호중학교", "운호중"),
    ("원평중학교", "원평중"), ("음성중학교", "음성중"),
    ("음성여자중학교", "음성여중"), ("이원중학교", "이원중"),
    ("이월중학교", "이월중"), ("일신여자중학교", "일신여중"),
    ("연수중학교", "연수중"), ("연풍중학교", "연풍중"),
    ("제천중학교", "제천중"), ("제천동중학교", "제천동중"),
    ("제천여자중학교", "제천여중"), ("주덕중학교", "주덕중"),
    ("주성중학교", "주성중"), ("증평중학교", "증평중"),
    ("증평여자중학교", "증평여중"), ("진천중학교", "진천중"),
    ("진천여자중학교", "진천여중"), ("중앙탑중학교", "중앙탑중"),
    ("광혜원중학교", "광혜원중"), ("청산중학교", "청산중"),
    ("청안중학교", "청안중"), ("청주중학교", "청주중"),
    ("청주남중학교", "청주남중"), ("청주동중학교", "청주동중"),
    ("청주여자중학교", "청주여중"), ("청주중앙중학교", "청주중앙중"),
    ("청주중앙여자중학교", "청주중앙여중"), ("청천중학교", "청천중"),
    ("칠금중학교", "칠금중"), ("탄금중학교", "탄금중"),
    ("학산중학교", "학산중"), ("한송중학교", "한송중"),
    ("한천중학교", "한천중"), ("형석중학교", "형석중"),
    ("혜화중학교", "혜화중"), ("황간중학교", "황간중"),
    ("회인중학교", "회인중"), ("흥덕중학교", "흥덕중"),
    ("한국교원대학교부설미호중학교", "한국교원대부설미호중"),
    ("한국교원대학교부설경덕중학교", "한국교원대부설경덕중"),
    ("충주중학교", "충주중"), ("충주여자중학교", "충주여중"),
    ("충원중학교", "충원중"), ("충북중학교", "충북중"),
    ("충북여자중학교", "충북여중"), ("한빛중학교", "한빛중"),
    ("은여울중학교", "은여울중"),
    ("강내중학교", "강내중"), ("내덕중학교", "내덕중"),
    ("낭성중학교", "낭성중"), ("동주중학교", "동주중"),
    ("방서중학교", "방서중"), ("북일중학교", "북일중"),
    ("수동중학교", "수동중"), ("신봉중학교", "신봉중"),
    ("신성중학교", "신성중"), ("양성중학교", "양성중"),
    ("영운중학교", "영운중"), ("율량중학교", "율량중"),
    ("청남중학교", "청남중"), ("청원중학교", "청원중"),
    ("충일중학교", "충일중"), ("팔봉중학교", "팔봉중"),
    ("현도중학교", "현도중"), ("가금중학교", "가금중"),
    ("달천중학교", "달천중"), ("대소원중학교", "대소원중"),
    ("봉방중학교", "봉방중"), ("성남중학교", "성남중"),
    ("성내중학교", "성내중"), ("봉양중학교", "봉양중"),
    ("신월중학교", "신월중"), ("의림중학교", "의림중"),
    ("청전중학교", "청전중"), ("삼승중학교", "삼승중"),
    ("수한중학교", "수한중"), ("군북중학교", "군북중"),
    ("매곡중학교", "매곡중"), ("심천중학교", "심천중"),
    ("용화중학교", "용화중"), ("증평남중학교", "증평남중"),
    ("초평중학교", "초평중"), ("문광중학교", "문광중"),
    ("소수중학교", "소수중"), ("칠성중학교", "칠성중"),
    ("금왕중학교", "금왕중"), ("맹동중학교", "맹동중"),
    ("삼성중학교", "삼성중"), ("대강중학교", "대강중"),
    ("영춘중학교", "영춘중"), ("적성중학교", "적성중"),
    # ── 초등학교 ────────────────────────────
    ("가경초등학교", "가경초"), ("가곡초등학교", "가곡초"),
    ("가덕초등학교", "가덕초"), ("가수초등학교", "가수초"),
    ("가엽초등학교", "가엽초"), ("가평초등학교", "가평초"),
    ("가흥초등학교", "가흥초"), ("각리초등학교", "각리초"),
    ("간디초등학교", "간디초"), ("갈원초등학교", "갈원초"),
    ("감곡초등학교", "감곡초"), ("감물초등학교", "감물초"),
    ("강내초등학교", "강내초"), ("강서초등학교", "강서초"),
    ("강외초등학교", "강외초"), ("강천초등학교", "강천초"),
    ("개신초등학교", "개신초"), ("경덕초등학교", "경덕초"),
    ("경산초등학교", "경산초"), ("계산초등학교", "계산초"),
    ("관기초등학교", "관기초"), ("광혜원초등학교", "광혜원초"),
    ("괴산명덕초등학교", "괴산명덕초"), ("교동초등학교", "교동초"),
    ("구룡초등학교", "구룡초"), ("구정초등학교", "구정초"),
    ("국원초등학교", "국원초"), ("군남초등학교", "군남초"),
    ("군서초등학교", "군서초"), ("금가초등학교", "금가초"),
    ("금구초등학교", "금구초"), ("금당초등학교", "금당초"),
    ("금성초등학교", "금성초"), ("금천초등학교", "금천초"),
    ("길상초등학교", "길상초"), ("남당초등학교", "남당초"),
    ("남성초등학교", "남성초"), ("남신초등학교", "남신초"),
    ("남이초등학교", "남이초"), ("남일초등학교", "남일초"),
    ("남천초등학교", "남천초"), ("남평초등학교", "남평초"),
    ("남한강초등학교", "남한강초"), ("낭성초등학교", "낭성초"),
    ("내북초등학교", "내북초"), ("내수초등학교", "내수초"),
    ("내토초등학교", "내토초"), ("노은초등학교", "노은초"),
    ("능산초등학교", "능산초"), ("단양초등학교", "단양초"),
    ("단월초등학교", "단월초"), ("단천초등학교", "단천초"),
    ("달천초등학교", "달천초"), ("대가초등학교", "대가초"),
    ("대강초등학교", "대강초"), ("대길초등학교", "대길초"),
    ("대미초등학교", "대미초"), ("대소원초등학교", "대소원초"),
    ("대소초등학교", "대소초"), ("대장초등학교", "대장초"),
    ("덕벌초등학교", "덕벌초"), ("덕성초등학교", "덕성초"),
    ("덕신초등학교", "덕신초"), ("덕산초등학교", "덕산초"),
    ("도안초등학교", "도안초"), ("동광초등학교", "동광초"),
    ("동락초등학교", "동락초"), ("동량초등학교", "동량초"),
    ("동명초등학교", "동명초"), ("동산초등학교", "동산초"),
    ("동성초등학교", "동성초"), ("동이초등학교", "동이초"),
    ("동인초등학교", "동인초"), ("동주초등학교", "동주초"),
    ("동화초등학교", "동화초"), ("두학초등학교", "두학초"),
    ("만승초등학교", "만승초"), ("매곡초등학교", "매곡초"),
    ("매포초등학교", "매포초"), ("맹동초등학교", "맹동초"),
    ("명지초등학교", "명지초"), ("모충초등학교", "모충초"),
    ("목도초등학교", "목도초"), ("목행초등학교", "목행초"),
    ("무극초등학교", "무극초"), ("문광초등학교", "문광초"),
    ("문백초등학교", "문백초"), ("문상초등학교", "문상초"),
    ("문의초등학교", "문의초"), ("미원초등학교", "미원초"),
    ("백곡초등학교", "백곡초"), ("백봉초등학교", "백봉초"),
    ("백운초등학교", "백운초"), ("보덕초등학교", "보덕초"),
    ("보은초등학교", "보은초"), ("봉덕초등학교", "봉덕초"),
    ("봉명초등학교", "봉명초"), ("봉양초등학교", "봉양초"),
    ("봉정초등학교", "봉정초"), ("봉학초등학교", "봉학초"),
    ("부강초등학교", "부강초"), ("부윤초등학교", "부윤초"),
    ("북이초등학교", "북이초"), ("분평초등학교", "분평초"),
    ("비봉초등학교", "비봉초"), ("사직초등학교", "사직초"),
    ("사천초등학교", "사천초"), ("산남초등학교", "산남초"),
    ("산외초등학교", "산외초"), ("산척초등학교", "산척초"),
    ("살미초등학교", "살미초"), ("삼보초등학교", "삼보초"),
    ("삼산초등학교", "삼산초"), ("삼양초등학교", "삼양초"),
    ("상당초등학교", "상당초"), ("상모초등학교", "상모초"),
    ("상신초등학교", "상신초"), ("상촌초등학교", "상촌초"),
    ("새터초등학교", "새터초"), ("생극초등학교", "생극초"),
    ("생명초등학교", "생명초"), ("서경초등학교", "서경초"),
    ("서원초등학교", "서원초"), ("서촌초등학교", "서촌초"),
    ("서현초등학교", "서현초"), ("석교초등학교", "석교초"),
    ("석성초등학교", "석성초"), ("성화초등학교", "성화초"),
    ("세명초등학교", "세명초"), ("소수초등학교", "소수초"),
    ("소이초등학교", "소이초"), ("속리초등학교", "속리초"),
    ("솔밭초등학교", "솔밭초"), ("송면초등학교", "송면초"),
    ("송절초등학교", "송절초"), ("송학초등학교", "송학초"),
    ("수곡초등학교", "수곡초"), ("수안보초등학교", "수안보초"),
    ("수회초등학교", "수회초"), ("숭덕초등학교", "숭덕초"),
    ("신니초등학교", "신니초"), ("신백초등학교", "신백초"),
    ("신항초등학교", "신항초"), ("심천초등학교", "심천초"),
    ("쌍봉초등학교", "쌍봉초"), ("안남초등학교", "안남초"),
    ("안내초등학교", "안내초"), ("앙성초등학교", "앙성초"),
    ("양강초등학교", "양강초"), ("양산초등학교", "양산초"),
    ("양청초등학교", "양청초"), ("어상천초등학교", "어상천초"),
    ("엄정초등학교", "엄정초"), ("연풍초등학교", "연풍초"),
    ("영동초등학교", "영동초"), ("영춘초등학교", "영춘초"),
    ("오갑초등학교", "오갑초"), ("오선초등학교", "오선초"),
    ("오송초등학교", "오송초"), ("오창초등학교", "오창초"),
    ("옥동초등학교", "옥동초"), ("옥산초등학교", "옥산초"),
    ("옥천초등학교", "옥천초"), ("용담초등학교", "용담초"),
    ("용두초등학교", "용두초"), ("용성초등학교", "용성초"),
    ("용암초등학교", "용암초"), ("용원초등학교", "용원초"),
    ("용천초등학교", "용천초"), ("용화초등학교", "용화초"),
    ("우암초등학교", "우암초"), ("운천초등학교", "운천초"),
    ("원남초등학교", "원남초"), ("원평초등학교", "원평초"),
    ("월곡초등학교", "월곡초"), ("율량초등학교", "율량초"),
    ("이수초등학교", "이수초"), ("이원초등학교", "이원초"),
    ("이월초등학교", "이월초"), ("이평초등학교", "이평초"),
    ("입석초등학교", "입석초"), ("장야초등학교", "장야초"),
    ("장연초등학교", "장연초"), ("장풍초등학교", "장풍초"),
    ("적성초등학교", "적성초"), ("제천초등학교", "제천초"),
    ("제천중앙초등학교", "제천중앙초"), ("죽리초등학교", "죽리초"),
    ("죽향초등학교", "죽향초"), ("주덕초등학교", "주덕초"),
    ("주성초등학교", "주성초"), ("중앙탑초등학교", "중앙탑초"),
    ("증평초등학교", "증평초"), ("진천초등학교", "진천초"),
    ("진천상산초등학교", "진천상산초"), ("진흥초등학교", "진흥초"),
    ("창리초등학교", "창리초"), ("창신초등학교", "창신초"),
    ("초강초등학교", "초강초"), ("초평초등학교", "초평초"),
    ("청산초등학교", "청산초"), ("청안초등학교", "청안초"),
    ("청주초등학교", "청주초"), ("청주중앙초등학교", "청주중앙초"),
    ("청천초등학교", "청천초"), ("청룡초등학교", "청룡초"),
    ("충주초등학교", "충주초"), ("충주교현초등학교", "충주교현초"),
    ("충주남산초등학교", "충주남산초"), ("충주대림초등학교", "충주대림초"),
    ("충주성남초등학교", "충주성남초"), ("충주중앙초등학교", "충주중앙초"),
    ("칠금초등학교", "칠금초"), ("탄금초등학교", "탄금초"),
    ("탄부초등학교", "탄부초"), ("하당초등학교", "하당초"),
    ("학산초등학교", "학산초"), ("학성초등학교", "학성초"),
    ("한벌초등학교", "한벌초"), ("한솔초등학교", "한솔초"),
    ("한송초등학교", "한송초"), ("한천초등학교", "한천초"),
    ("행정초등학교", "행정초"), ("현도초등학교", "현도초"),
    ("홍광초등학교", "홍광초"), ("화당초등학교", "화당초"),
    ("화산초등학교", "화산초"), ("황간초등학교", "황간초"),
    ("회남초등학교", "회남초"), ("회인초등학교", "회인초"),
    ("흥덕초등학교", "흥덕초"),
    ("한국교원대학교부설월곡초등학교", "한국교원대부설월곡초"),
    ("가금초등학교", "가금초"), ("개령초등학교", "개령초"),
    ("대림초등학교", "대림초"), ("목벌초등학교", "목벌초"),
    ("문화초등학교", "문화초"), ("봉방초등학교", "봉방초"),
    ("성남초등학교", "성남초"), ("성내초등학교", "성내초"),
    ("소태초등학교", "소태초"), ("안림초등학교", "안림초"),
    ("용탄초등학교", "용탄초"), ("충원초등학교", "충원초"),
    ("신월초등학교", "신월초"), ("의림초등학교", "의림초"),
    ("자작초등학교", "자작초"), ("청전초등학교", "청전초"),
    ("한수초등학교", "한수초"), ("삼승초등학교", "삼승초"),
    ("수한초등학교", "수한초"), ("군북초등학교", "군북초"),
    ("청성초등학교", "청성초"), ("용산초등학교", "용산초"),
    ("남하초등학교", "남하초"), ("보강초등학교", "보강초"),
    ("증평남초등학교", "증평남초"), ("불정초등학교", "불정초"),
    ("사리초등학교", "사리초"), ("금왕초등학교", "금왕초"),
    ("삼성초등학교", "삼성초"),
    # ── 특수 및 각종학교 ─────────────────────
    ("청주혜화학교", "청주혜화학교"), ("청주혜원학교", "청주혜원학교"),
    ("숭덕학교", "숭덕학교"), ("충주성심학교", "충주성심학교"),
    ("충주성모학교", "충주성모학교"), ("제천청암학교", "제천청암학교"),
    ("꽃동네학교", "꽃동네학교"),
    ("청주성심학교", "청주성심학교"), ("청주명암학교", "청주명암학교"),
    ("충북혜림학교", "충북혜림학교"), ("충북혜림특수학교", "충북혜림학교"),
    ("충북맹학교", "충북맹학교"), ("충북농학교", "충북농학교"),
    ("제천혜화학교", "제천혜화학교"), ("청주진흥학교", "청주진흥학교"),
))

# ─────────────────────────────────────────────
#  파싱 함수
# ─────────────────────────────────────────────

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
        for k, v in _ORG_LOOKUP.items():
            kpre, ksuf = _split_school_suffix(k)
            if kpre == prefix and ksuf == suffix:
                return v
    matches = difflib.get_close_matches(s, keys, n=1, cutoff=0.6)
    if not matches and len(abbr) > 1:
        matches = difflib.get_close_matches(abbr, keys, n=1, cutoff=0.6)
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
        pair = extract_pair(tokens)
        if pair:
            org, name = pair
            if not org and pending_org:
                org = pending_org
            pending_org = ''
            results.append({'org': org, 'name': name})
            continue
        if all(is_person_name(t) for t in tokens):
            for t in tokens:
                results.append({'org': pending_org, 'name': t})
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


# ─────────────────────────────────────────────
#  HWP 텍스트 추출
# ─────────────────────────────────────────────

def _hwpx_zipxml(path: str) -> str:
    """HWPX(ZIP+XML) 파일에서 텍스트 추출"""
    import zipfile
    import xml.etree.ElementTree as ET
    try:
        with zipfile.ZipFile(path, 'r') as z:
            texts = []
            for name in sorted(z.namelist()):
                if re.search(r'[Ss]ection\d+\.xml', name):
                    root = ET.fromstring(z.read(name).decode('utf-8', 'ignore'))
                    for el in root.iter():
                        if el.text and el.text.strip():
                            texts.append(el.text.strip())
            return '\n'.join(texts)
    except Exception:
        return ''


def _hwp_win32com(path: str) -> str:
    try:
        import win32com.client as client
        hwp = client.gencache.EnsureDispatch('HWPFrame.HwpObject')
        hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
        hwp.Open(path, 'HWP', 'forceopen:true')
        text = hwp.GetTextFile('TEXT', '').strip()
        hwp.Quit()
        return text
    except Exception:
        return ''


def _hwp_olefile(path: str) -> str:
    """
    olefile로 HWP 5.x 바이너리 직접 명단 추출.
    BodyText/Section* 스트림에서 HWPTAG_PARA_TEXT(tag=67) 레코드 추출.
    """
    try:
        import olefile
        import zlib
        import struct
    except ImportError:
        return ''
    try:
        ole = olefile.OleFileIO(path)
        header = ole.openstream('FileHeader').read()
        compressed = bool(struct.unpack_from('<I', header, 36)[0] & 1)
        texts = []
        for i in range(100):
            sname = f'BodyText/Section{i}'
            if not ole.exists(sname):
                break
            data = ole.openstream(sname).read()
            if compressed:
                try:
                    data = zlib.decompress(data, -15)
                except Exception:
                    pass
            offset = 0
            while offset + 4 <= len(data):
                hw = struct.unpack_from('<I', data, offset)[0]
                tag = hw & 0x3FF
                size = (hw >> 20) & 0xFFF
                if size == 0xFFF:
                    size = struct.unpack_from('<I', data, offset + 4)[0]
                    offset += 4
                offset += 4
                if tag == 67 and offset + size <= len(data):
                    chunk = data[offset:offset + size]
                    txt = ''.join(
                        chr(struct.unpack_from('<H', chunk, j)[0])
                        for j in range(0, len(chunk) - 1, 2)
                        if struct.unpack_from('<H', chunk, j)[0] >= 0x20
                    ).strip()
                    if txt:
                        texts.append(txt)
                offset += size
        ole.close()
        return '\n'.join(texts)
    except Exception:
        return ''


def extract_hwp_text(filepath: str) -> str:
    """HWP / HWPX 파일에서 텍스트 추출 (3가지 방법 순차 시도)"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.hwpx':
        return _hwpx_zipxml(filepath)
    text = _hwp_win32com(filepath)
    if text:
        return text
    text = _hwp_olefile(filepath)
    if text:
        return text
    try:
        result = subprocess.run(
            ['hwp5txt', filepath], capture_output=True, text=True, encoding='utf-8'
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return ''


# ─────────────────────────────────────────────
#  Config
# ─────────────────────────────────────────────

class Config:
    DEFAULTS = {
        'search_field_x': None, 'search_field_y': None,
        'result_first_x': None, 'result_first_y': None,
        'empty_pixel_rgb': None,
        'search_delay': 0.5,
        'manual_confirm': False,
    }

    def __init__(self):
        self.data = dict(self.DEFAULTS)
        self._load()

    def _load(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    self.data.update(json.load(f))
        except Exception:
            pass

    def save(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.data, f)

    def is_calibrated(self):
        return (self.data.get('search_field_x') is not None and
                self.data.get('result_first_x') is not None)


# ─────────────────────────────────────────────
#  CaptureDialog
# ─────────────────────────────────────────────

class CaptureDialog(tk.Toplevel):
    def __init__(self, parent, on_captured):
        super().__init__(parent)
        self.on_captured = on_captured
        self.title('위치 캡처')
        self.geometry('420x220')
        self.resizable(False, False)
        self.grab_set()
        tk.Label(
            self,
            text='3초 카운트다운 후 마우스를 해당 위치로 이동하세요.',
            bg='#555', fg='white', font=('맑은 고딕', 11), pady=18
        ).pack(fill='x')
        self.status = tk.Label(
            self, text='📍 3초 카운트다운 후 캡처',
            font=('맑은 고딕', 13, 'bold'), fg='#FF9800'
        )
        self.status.pack(pady=14)
        tk.Button(self, text='취소', command=self.destroy).pack()
        self._start()

    def _start(self):
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        self.iconify()
        for i in range(3, 0, -1):
            self.after(0, lambda i=i: self.status.config(text=f'{i}초 후 캡처...'))
            time.sleep(1)
        pos = pyautogui.position()
        self.after(0, lambda: self._done(pos))

    def _done(self, pos):
        self.deiconify()
        self.status.config(text=f'✓ 캡처 완료: ({pos.x}, {pos.y})', fg='green')
        self.on_captured(pos.x, pos.y)
        self.after(1200, self.destroy)


# ─────────────────────────────────────────────
#  도움말 텍스트
# ─────────────────────────────────────────────
_HELP_TEXT = """━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  충북 소통메신저  자동 사용자 선택  v1.5
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
  2. 오른쪽 위의 편지 버튼(✉)을 클릭한 뒤,
     '쪽지 작성' 또는 '대화하기' 메뉴를 선택합니다.
  3. 메시지 보내기 화면에서 [사용자 선택] 버튼을 클릭합니다.
  4. 사용자 선택 창이 열리면, 상단 탭에서 [전체조직]을 선택합니다.
  5. 소통메신저 창과 이 프로그램 창을 화면에 나란히 배치하면 편리합니다.

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
    · 소속기관이 없거나 인식에 실패한 항목은 목록에 표시되며 제외 안내됩니다.
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

  충청북도교육청 소통메신저 자동화 도구  |  버전 v1.5
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"""


# ─────────────────────────────────────────────
#  App
# ─────────────────────────────────────────────

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title('충북 소통메신저 자동 사용자 선택 v1.5')
        self.root.geometry('900x720')
        self.root.minsize(880, 700)

        self.config = Config()
        self.names_list: list = []
        self.stop_flag = threading.Event()
        self.continue_event = threading.Event()
        self.continue_event.set()

        self._apply_theme()
        self._build_ui()
        self._refresh_calib_labels()
        self._check_deps()

    # ── 테마 (충북교육청 블루) ─────────────────
    def _apply_theme(self):
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure('TNotebook', background='#E8EAF6', tabmargins=[2, 5, 2, 0])
        style.configure('TNotebook.Tab', padding=[14, 7],
                        font=('맑은 고딕', 10, 'bold'),
                        background='#C5CAE9', foreground='#37474F')
        style.map('TNotebook.Tab',
                  background=[('selected', '#1565C0')],
                  foreground=[('selected', 'white')])
        style.configure('TFrame', background='#F5F7FA')
        style.configure('TLabelframe', background='#F5F7FA')
        style.configure('TLabelframe.Label',
                        font=('맑은 고딕', 10, 'bold'), foreground='#1565C0')

    # ── UI 빌드 ────────────────────────────────
    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # 개발자 정보 바 (항상 상단 표시)
        dev_bar = tk.Label(
            self.root,
            text='  Developed by 송동석  |  Teacher · Data Analytics · App Developer'
                 '  |  협업: dungst.me@gmail.com  |  v1.5  ',
            bg='#1565C0', fg='white',
            font=('맑은 고딕', 9), anchor='w', pady=5
        )
        dev_bar.grid(row=0, column=0, sticky='ew')

        # 탭 노트북
        nb = ttk.Notebook(self.root)
        nb.grid(row=1, column=0, sticky='nsew', padx=4, pady=(0, 4))

        f1 = ttk.Frame(nb); nb.add(f1, text='  1. 명단 입력  ')
        f2 = ttk.Frame(nb); nb.add(f2, text='  2. 위치 설정  ')
        f3 = ttk.Frame(nb); nb.add(f3, text='  3. 자동 선택  ')
        f4 = ttk.Frame(nb); nb.add(f4, text='  📖 사용 방법  ')

        self._tab_input(f1)
        self._tab_calib(f2)
        self._tab_auto(f3)
        self._tab_help(f4)

        # 상태바
        self.status_var = tk.StringVar(value='준비')
        tk.Label(
            self.root, textvariable=self.status_var,
            relief='sunken', anchor='w', bg='#f0f0f0', fg='#333',
            font=('맑은 고딕', 9), pady=3
        ).grid(row=2, column=0, sticky='ew')

    # ── 탭 1: 명단 입력 ────────────────────────
    def _tab_input(self, frame: ttk.Frame):
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        frame.rowconfigure(5, weight=2)

        # ① 입력 형식 안내 박스
        guide = tk.Frame(frame, bg='#E3F2FD', bd=1, relief='solid')
        guide.grid(row=0, column=0, sticky='ew', padx=10, pady=(10, 4))
        guide.columnconfigure(0, weight=1)

        tk.Label(
            guide,
            text='📋  명단 붙여넣기 형식 안내',
            bg='#E3F2FD', fg='#0D47A1', font=('맑은 고딕', 9, 'bold'), anchor='w'
        ).grid(row=0, column=0, sticky='w', padx=10, pady=(6, 2))

        tk.Label(
            guide,
            text='엑셀에서 소속기관(A열)과 이름(B열)을 함께 선택한 뒤 복사(Ctrl+C)하고 아래에 붙여넣기 하세요.\n'
                 '예)  충주중학교    홍길동\n'
                 '     청주고등학교  김철수\n'
                 '소속기관이 없는 경우 이름만 입력해도 됩니다. 엑셀·HWP 파일을 직접 열어도 됩니다.',
            bg='#E3F2FD', fg='#333', font=('맑은 고딕', 9), justify='left', anchor='w'
        ).grid(row=1, column=0, sticky='w', padx=10, pady=(0, 8))

        # ② 파일 열기 버튼 행
        file_btn_frame = tk.Frame(frame, bg='#F5F7FA')
        file_btn_frame.grid(row=1, column=0, sticky='ew', padx=8, pady=(0, 2))

        for text, cmd, bg in [
            ('엑셀 파일 열기 (.xlsx)', self._open_excel, '#607D8B'),
            ('HWP 파일 열기 (.hwp)',   self._open_hwp,   '#607D8B'),
        ]:
            tk.Button(
                file_btn_frame, text=text, command=cmd,
                bg=bg, fg='white', activebackground=bg,
                relief='flat', font=('맑은 고딕', 9), padx=8, pady=4, cursor='hand2'
            ).pack(side='left', padx=3)

        # ③ 텍스트 입력 영역
        self.input_text = scrolledtext.ScrolledText(
            frame, height=8, font=('맑은 고딕', 9), wrap='none'
        )
        self.input_text.grid(row=2, column=0, sticky='nsew', padx=8, pady=4)

        # ④ 추출 버튼 행
        action_frame = tk.Frame(frame, bg='#F5F7FA')
        action_frame.grid(row=3, column=0, sticky='ew', padx=8, pady=(0, 4))

        for text, cmd, bg in [
            ('명단 추출 →', self._parse,       '#1565C0'),
            ('초기화',       self._clear_input, '#9E9E9E'),
        ]:
            tk.Button(
                action_frame, text=text, command=cmd,
                bg=bg, fg='white', activebackground=bg,
                relief='flat', font=('맑은 고딕', 9, 'bold'), padx=12, pady=5, cursor='hand2'
            ).pack(side='left', padx=3)

        # ⑤ 추출 결과 상태 라벨
        self.parse_status = tk.Label(
            frame, text='', fg='#555', font=('맑은 고딕', 9), anchor='w'
        )
        self.parse_status.grid(row=4, column=0, sticky='w', padx=12, pady=(0, 2))

        # ⑥ 추출된 명단 리스트
        list_frame = tk.Frame(frame)
        list_frame.grid(row=5, column=0, sticky='nsew', padx=8, pady=(0, 4))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.parsed_list = tk.Listbox(
            list_frame, font=('맑은 고딕', 9), selectmode='extended',
            activestyle='none', selectbackground='#1565C0', selectforeground='white'
        )
        self.parsed_list.grid(row=0, column=0, sticky='nsew')
        sb = ttk.Scrollbar(list_frame, orient='vertical',
                           command=self.parsed_list.yview)
        sb.grid(row=0, column=1, sticky='ns')
        self.parsed_list.config(yscrollcommand=sb.set)

        # ⑦ 하단 버튼 + 범례
        bottom_frame = tk.Frame(frame, bg='#F5F7FA')
        bottom_frame.grid(row=6, column=0, sticky='ew', padx=8, pady=(0, 6))
        bottom_frame.columnconfigure(1, weight=1)

        tk.Button(
            bottom_frame, text='선택 항목 삭제', command=self._delete_selected,
            bg='#9E9E9E', fg='white', activebackground='#9E9E9E',
            relief='flat', font=('맑은 고딕', 9), padx=8, pady=3, cursor='hand2'
        ).grid(row=0, column=0, sticky='w')

        tk.Label(
            bottom_frame, text='● 빨간색 항목 = 자동 선택 실패 (검색 결과 없음)',
            fg='#B71C1C', font=('맑은 고딕', 8), bg='#F5F7FA', anchor='e'
        ).grid(row=0, column=1, sticky='e', padx=(0, 4))

        ctx = tk.Menu(self.root, tearoff=0)
        ctx.add_command(label='삭제', command=self._delete_selected)
        self.parsed_list.bind('<Button-3>', lambda e: ctx.tk_popup(e.x_root, e.y_root))

    # ── 탭 2: 위치 설정 ────────────────────────
    def _tab_calib(self, frame: ttk.Frame):
        frame.columnconfigure(0, weight=1)

        tk.Label(
            frame,
            text="소통메신저 '사용자 선택' 창을 열어 둔 상태에서 아래 순서대로 위치를 설정하세요.\n"
                 "창은 항상 같은 위치에 두는 것이 좋습니다.",
            fg='#555', font=('맑은 고딕', 9), justify='center'
        ).grid(row=0, column=0, sticky='ew', padx=10, pady=(10, 6))

        # STEP 1·2: 위치 설정
        pos_frame = ttk.LabelFrame(frame, text='STEP 1 · 2 — 검색창 및 결과 위치 설정')
        pos_frame.grid(row=1, column=0, sticky='ew', padx=10, pady=4)
        pos_frame.columnconfigure(1, weight=1)

        for row_i, (label_text, key) in enumerate([
            ('STEP 1  검색 입력창:', 'search_field'),
            ('STEP 2  결과 첫 번째:', 'result_first'),
        ]):
            tk.Label(pos_frame, text=label_text,
                     font=('맑은 고딕', 9, 'bold')).grid(
                row=row_i, column=0, padx=8, pady=6, sticky='w')

            lbl = tk.Label(pos_frame, text='미설정', fg='red',
                           font=('맑은 고딕', 9))
            lbl.grid(row=row_i, column=1, sticky='w', padx=4)
            setattr(self, f'lbl_{key}', lbl)

            tk.Button(
                pos_frame, text='📍 위치 설정',
                bg='#FF9800', fg='white', activebackground='#FF9800',
                relief='flat', font=('맑은 고딕', 9), padx=6, pady=3,
                cursor='hand2',
                command=lambda k=key: self._do_capture(k)
            ).grid(row=row_i, column=2, padx=8, pady=4)

        # 검색 설정
        setting_frame = ttk.LabelFrame(frame, text='검색 설정')
        setting_frame.grid(row=3, column=0, sticky='ew', padx=10, pady=4)

        tk.Label(setting_frame, text='검색 후 대기 시간(초):',
                 font=('맑은 고딕', 9)).grid(row=0, column=0, padx=8, pady=6, sticky='w')

        self.delay_var = tk.DoubleVar(value=self.config.data.get('search_delay', 0.5))
        ttk.Spinbox(
            setting_frame, from_=0.3, to=5.0, increment=0.1,
            textvariable=self.delay_var, width=6,
            font=('맑은 고딕', 9)
        ).grid(row=0, column=1, padx=4, sticky='w')

        tk.Label(setting_frame, text='(느리면 값을 높이세요)',
                 fg='#888', font=('맑은 고딕', 9)).grid(
            row=0, column=2, padx=4, sticky='w')

        self.manual_var = tk.BooleanVar(
            value=self.config.data.get('manual_confirm', False)
        )
        tk.Checkbutton(
            setting_frame,
            text='수동 확인 모드 — 검색 후 [계속] 버튼을 눌러야 다음으로 진행\n'
                 '(동명이인이 많을 때 사용: 검색 결과에서 사용자가 직접 더블클릭)',
            variable=self.manual_var,
            font=('맑은 고딕', 9), justify='left', anchor='w'
        ).grid(row=1, column=0, columnspan=3, sticky='w', padx=8, pady=4)

        tk.Button(
            frame, text='✅  설정 저장',
            bg='#4CAF50', fg='white', activebackground='#4CAF50',
            relief='flat', font=('맑은 고딕', 10, 'bold'), padx=12, pady=6,
            cursor='hand2', command=self._save_calib
        ).grid(row=4, column=0, pady=8)

        self.calib_msg = tk.Label(
            frame, text='', fg='green', font=('맑은 고딕', 9)
        )
        self.calib_msg.grid(row=5, column=0)

    # ── 탭 3: 자동 선택 ────────────────────────
    def _tab_auto(self, frame: ttk.Frame):
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        tk.Label(
            frame,
            text="소통메신저 '사용자 선택' 창 → 전체조직 탭을 열고 [시작] 버튼을 누르세요.\n"
                 "검색 형식: 소속+이름  (소속 없으면 이름만)\n"
                 "⚠  마우스를 화면 왼쪽 위 모서리로 이동하면 긴급 중지됩니다.",
            fg='#555', font=('맑은 고딕', 9), justify='center'
        ).grid(row=0, column=0, sticky='ew', padx=10, pady=(10, 6))

        btn_frame = tk.Frame(frame, bg='#F5F7FA')
        btn_frame.grid(row=1, column=0, sticky='ew', padx=8, pady=4)

        self.start_btn = tk.Button(
            btn_frame, text='▶  자동 선택 시작',
            bg='#4CAF50', fg='white', activebackground='#388E3C',
            relief='flat', font=('맑은 고딕', 10, 'bold'), padx=10, pady=6,
            cursor='hand2', command=self._start
        )
        self.start_btn.pack(side='left', padx=4)

        self.continue_btn = tk.Button(
            btn_frame, text='▶▶  계속',
            bg='#2196F3', fg='white', activebackground='#1565C0',
            relief='flat', font=('맑은 고딕', 10, 'bold'), padx=10, pady=6,
            cursor='hand2', state='disabled', command=self._resume
        )
        self.continue_btn.pack(side='left', padx=4)

        self.stop_btn = tk.Button(
            btn_frame, text='■  중지',
            bg='#F44336', fg='white', activebackground='#C62828',
            relief='flat', font=('맑은 고딕', 10, 'bold'), padx=10, pady=6,
            cursor='hand2', state='disabled', command=self._stop
        )
        self.stop_btn.pack(side='left', padx=4)

        tk.Label(frame, text='진행 상황:', anchor='w',
                 font=('맑은 고딕', 9)).grid(
            row=3, column=0, sticky='w', padx=10, pady=(6, 2))

        self.log = scrolledtext.ScrolledText(
            frame, height=14, state='disabled',
            font=('맑은 고딕', 9), wrap='none'
        )
        self.log.tag_config('ok', foreground='#1B5E20')
        self.log.tag_config('fail', foreground='#B71C1C')
        self.log.grid(row=2, column=0, sticky='nsew', padx=8, pady=4)

        prog_row = tk.Frame(frame, bg='#F5F7FA')
        prog_row.grid(row=4, column=0, sticky='ew', padx=8, pady=(2, 8))
        prog_row.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(prog_row, mode='determinate')
        self.progress.grid(row=0, column=0, sticky='ew')

        self.prog_label = tk.Label(
            prog_row, text='0 / 0', font=('맑은 고딕', 9), bg='#F5F7FA'
        )
        self.prog_label.grid(row=0, column=1, padx=8)

    # ── 탭 4: 사용 방법 ────────────────────────
    def _tab_help(self, frame: ttk.Frame):
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        txt = scrolledtext.ScrolledText(
            frame, font=('맑은 고딕', 10), wrap='word', state='normal'
        )
        txt.grid(row=0, column=0, sticky='nsew', padx=4, pady=4)
        txt.insert('1.0', _HELP_TEXT)
        txt.config(state='disabled')

    # ── 의존성 확인 ────────────────────────────
    def _check_deps(self):
        missing = []
        if pyautogui is None:
            missing.append('pyautogui')
        if pyperclip is None:
            missing.append('pyperclip')
        if missing:
            messagebox.showerror(
                '패키지 누락',
                f'필수 패키지 미설치: {", ".join(missing)}\n\n'
                '시작.bat 을 실행하면 자동으로 설치됩니다.'
            )

    # ── 엑셀 열기 ──────────────────────────────
    def _open_excel(self):
        if openpyxl is None:
            messagebox.showerror('오류', 'pip install openpyxl 필요')
            return
        path = filedialog.askopenfilename(
            filetypes=[('Excel', '*.xlsx'), ('모든 파일', '*.*')]
        )
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            lines = []
            for row in ws.iter_rows():
                cells = [str(c.value).strip() if c.value is not None else '' for c in row]
                if any(cells):
                    lines.append('\t'.join(cells))
            wb.close()
            self.input_text.delete('1.0', 'end')
            self.input_text.insert('1.0', '\n'.join(lines))
            self.status_var.set(f'엑셀 로드: {os.path.basename(path)}')
        except Exception as e:
            messagebox.showerror('파일 읽기 실패', f'파일 읽기 실패:\n{e}')

    # ── HWP 열기 ───────────────────────────────
    def _open_hwp(self):
        path = filedialog.askopenfilename(
            filetypes=[('HWP', '*.hwp *.hwpx'), ('모든 파일', '*.*')]
        )
        if not path:
            return
        self.status_var.set('HWP 파일 읽는 중...')
        self.root.update()
        text = extract_hwp_text(path)
        if text:
            self.input_text.delete('1.0', 'end')
            self.input_text.insert('1.0', text)
            self.status_var.set(f'HWP 로드: {os.path.basename(path)}')
        else:
            messagebox.showwarning(
                'HWP 읽기 실패',
                'HWP 파일을 자동으로 읽지 못했습니다.\n\n'
                'HWP에서 해당 표/목록을 직접 복사(Ctrl+C)하여\n'
                '텍스트 입력창에 붙여넣기 해주세요.'
            )
            self.status_var.set('HWP 읽기 실패 — 직접 복사·붙여넣기 필요')

    # ── 명단 추출 ──────────────────────────────
    def _parse(self):
        raw = self.input_text.get('1.0', 'end')
        parsed = parse_input(raw)
        self.names_list.clear()
        self.parsed_list.delete(0, 'end')

        ok = fail = no_org = 0
        warn_items = []
        for item in parsed:
            org = item.get('org', '')
            name = item.get('name', '')
            if not name:
                fail += 1
                continue
            if not org:
                no_org += 1
                warn_items.append(name)
            self.names_list.append({'org': org, 'name': name})
            label = f'[{org}]  {name}' if org else f'{name}  (소속없음)'
            self.parsed_list.insert('end', label)
            ok += 1

        color = 'green' if ok > 0 else 'red'
        parts = [f'명단 추출 완료: {ok}명']
        if fail:
            parts.append(f'인식실패 {fail}')
        if no_org:
            parts.append(f'소속없음 {no_org}')
            if warn_items:
                parts.append(f'({", ".join(warn_items[:5])}{"..." if len(warn_items) > 5 else ""})')
        self.parse_status.config(
            text='  /  제외: '.join(parts) if (fail or no_org) else parts[0],
            fg=color
        )
        self.status_var.set(f'명단 추출 완료: {ok}명')

    def _clear_input(self):
        self.input_text.delete('1.0', 'end')
        self.parsed_list.delete(0, 'end')
        self.names_list.clear()
        self.parse_status.config(text='')

    def _delete_selected(self):
        for i in reversed(self.parsed_list.curselection()):
            self.parsed_list.delete(i)
            del self.names_list[i]
        total = len(self.names_list)
        self.parse_status.config(
            text=f'명단 추출 완료: {total}명', fg='green'
        )

    # ── 위치 캡처 ──────────────────────────────
    def _do_capture(self, key: str):
        if pyautogui is None:
            messagebox.showerror('오류', 'pyautogui 미설치')
            return

        def on_captured(x, y):
            self.config.data[key + '_x'] = x
            self.config.data[key + '_y'] = y
            self._refresh_calib_labels()

        CaptureDialog(self.root, on_captured)

    def _save_calib(self):
        self.config.data['search_delay'] = round(self.delay_var.get(), 1)
        self.config.data['manual_confirm'] = self.manual_var.get()
        self.config.save()
        self.calib_msg.config(text='✅ 설정 저장 완료')
        self.root.after(2000, lambda: self.calib_msg.config(text=''))

    def _refresh_calib_labels(self):
        for key in ('search_field', 'result_first'):
            x = self.config.data.get(key + '_x')
            y = self.config.data.get(key + '_y')
            lbl = getattr(self, f'lbl_{key}', None)
            if lbl:
                if x is not None and y is not None:
                    lbl.config(text=f'✓ ({x}, {y})', fg='green')
                else:
                    lbl.config(text='미설정', fg='red')

    # ── 자동 선택 시작/중지/계속 ───────────────
    def _start(self):
        if pyautogui is None or pyperclip is None:
            messagebox.showerror('오류', '필수 패키지 미설치')
            return
        if not self.names_list:
            messagebox.showwarning('알림', '먼저 명단을 추출해 주세요.')
            return
        if not self.config.is_calibrated():
            messagebox.showwarning(
                '알림', '위치 설정 탭에서 검색창·결과 위치를 먼저 설정하세요.'
            )
            return
        self.stop_flag.clear()
        self.continue_event.set()
        self.start_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.continue_btn.config(state='disabled')
        total = len(self.names_list)
        self.progress.config(maximum=total, value=0)
        self.prog_label.config(text=f'0 / {total}')
        self._log_clear()
        self._log(f'자동 선택 시작 — 총 {total}명\n\n')
        threading.Thread(target=self._worker, daemon=True).start()

    def _stop(self):
        self.stop_flag.set()
        self.continue_event.set()
        self.start_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.continue_btn.config(state='disabled')
        self._log('\n⏹  중지\n')

    def _resume(self):
        self.continue_event.set()
        self.continue_btn.config(state='disabled')

    # ── 자동화 워커 ────────────────────────────
    def _worker(self):
        ok = fail = 0
        manual = self.config.data.get('manual_confirm', False)
        total = len(self.names_list)

        for idx, item in enumerate(self.names_list):
            if self.stop_flag.is_set():
                break
            org = item.get('org', '')
            name = item.get('name', '')
            search_str = (org + '+' + name) if org else name
            prefix = f'[{idx + 1}/{total}]  '

            self._log(f'{prefix}{search_str}  ... ')
            try:
                self._do_search(search_str)
                time.sleep(self.config.data.get('search_delay', 0.5))

                if not self._has_result():
                    fail += 1
                    self._log('✗  (검색 결과 없음)\n')
                    self._mark_failed(idx)
                    self._update_progress(idx + 1, total)
                    continue

                if manual:
                    self.continue_event.clear()
                    self.root.after(0, self._show_continue)
                    self.continue_event.wait()
                    if self.stop_flag.is_set():
                        break
                    ok += 1
                    self._log('✓\n')
                else:
                    result = self._do_select()
                    if result == 'duplicate':
                        fail += 1
                        self._log('⚠  (이미 선택된 사용자)\n')
                        self._mark_failed(idx)
                        self._update_progress(idx + 1, total)
                        continue
                    ok += 1
                    self._log('✓\n')
            except pyautogui.FailSafeException:
                self._log('\n⚠  긴급 중지 (화면 모서리)\n')
                self.stop_flag.set()
                break
            except Exception as e:
                fail += 1
                self._log(f'✗  ({e})\n')
                self._mark_failed(idx)

            self._update_progress(idx + 1, total)
            time.sleep(0.1)

        self.root.after(0, lambda: self._done(ok, fail))

    def _do_search(self, search_str: str):
        x = self.config.data['search_field_x']
        y = self.config.data['search_field_y']
        delay = self.config.data.get('search_delay', 0.5) * 0.3
        pyautogui.click(x, y)
        time.sleep(delay)
        pyautogui.hotkey('ctrl', 'a')
        pyperclip.copy(search_str)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(delay)
        pyautogui.press('enter')

    def _do_select(self) -> str:
        """더블클릭 후 팝업 여부로 결과 반환: 'ok' | 'duplicate'"""
        x = self.config.data['result_first_x']
        y = self.config.data['result_first_y']
        time.sleep(0.2)
        pyautogui.doubleClick(x, y)
        time.sleep(0.4)
        if self._popup_appeared():
            pyautogui.press('enter')
            time.sleep(0.15)
            return 'duplicate'
        return 'ok'

    def _popup_appeared(self) -> bool:
        """win32gui로 모달 다이얼로그 출현 감지."""
        try:
            import win32gui
            found = []
            def cb(hwnd, _):
                if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                    if win32gui.GetClassName(hwnd) == '#32770':
                        found.append(hwnd)
            win32gui.EnumWindows(cb, None)
            return bool(found)
        except Exception:
            return False

    def _has_result(self) -> bool:
        """검색 결과 첫 항목 위치의 픽셀 밝기로 결과 유무 자동 판단."""
        x = self.config.data.get('result_first_x')
        y = self.config.data.get('result_first_y')
        if x is None or y is None:
            return True
        try:
            r, g, b = pyautogui.pixel(x, y)[:3]
            # 결과 없으면 배경이 흰색(240+) → 결과 있으면 텍스트로 인해 더 어두움
            return not (r > 240 and g > 240 and b > 240)
        except Exception:
            return True

    def _show_continue(self):
        name = self.names_list[0].get('name', '') if self.names_list else ''
        self.status_var.set(
            f'수동 선택 대기: {name}  →  소통메신저에서 더블클릭 후 [계속] 버튼'
        )
        self.continue_btn.config(state='normal')

    def _done(self, ok: int, fail: int):
        self.start_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.continue_btn.config(state='disabled')
        sep = '─' * 44
        self._log(f'\n{sep}\n완료  ✓ {ok}명   ✗ {fail}명\n')
        if fail:
            self.status_var.set(f'완료 — 성공: {ok}명, 실패: {fail}명  ← 빨간색 항목 확인')
            messagebox.showwarning(
                '추가 실패 알림',
                f'받는 사람에 추가되지 않은 인원이 있습니다.\n\n'
                f'  ✓ 성공: {ok}명\n'
                f'  ✗ 실패: {fail}명\n\n'
                f'[1. 명단 입력] 탭에서 빨간색 항목을 확인하세요.\n'
                f'(검색 결과 없음 또는 이미 선택된 사용자)'
            )
        else:
            self.status_var.set(f'완료 — 성공: {ok}명')

    def _update_progress(self, idx: int, total: int):
        self.root.after(
            0,
            lambda: (
                self.progress.config(value=idx),
                self.prog_label.config(text=f'{idx} / {total}'),
            )
        )

    def _mark_failed(self, idx: int):
        self.root.after(
            0,
            lambda: self.parsed_list.itemconfig(
                idx, {'bg': '#FFCDD2', 'fg': '#B71C1C'}
            )
        )

    def _log(self, msg: str):
        self.root.after(0, lambda: self._log_append(msg))

    def _log_append(self, msg: str):
        self.log.config(state='normal')
        if '✓' in msg:
            self.log.insert('end', msg, 'ok')
        elif '✗' in msg or '⚠' in msg:
            self.log.insert('end', msg, 'fail')
        else:
            self.log.insert('end', msg)
        self.log.see('end')
        self.log.config(state='disabled')

    def _log_clear(self):
        self.log.config(state='normal')
        self.log.delete('1.0', 'end')
        self.log.config(state='disabled')


# ─────────────────────────────────────────────
#  진입점
# ─────────────────────────────────────────────
if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
