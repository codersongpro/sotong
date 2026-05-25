"""소통픽 — 충북 소통메신저 자동 사용자 선택"""

APP_NAME    = '소통픽'
APP_VERSION = '1.7.1'

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import time
import json
import logging
import os
import webbrowser
import urllib.request

from app_config import Config
from automation import (
    FAIL_DUPLICATE,
    FAIL_MANUAL_STOP,
    FAIL_NO_USER,
    failure_reason_from_error,
)
from hwp_extract import extract_hwp_text
from sotong_parser import parse_input
from ui_helpers import format_item_label

LOG_FILE = os.path.join(os.path.expanduser("~"), ".chungbuk_auto.log")
LATEST_RELEASE_API = 'https://api.github.com/repos/codersongpro/sotong/releases/latest'
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, encoding='utf-8')

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
#  CaptureDialog
# ─────────────────────────────────────────────

class CaptureDialog(tk.Toplevel):
    def __init__(self, parent, on_captured, label='위치'):
        super().__init__(parent)
        self.on_captured = on_captured
        self.title('위치 캡처')
        self.geometry('420x270')
        self.resizable(False, False)
        self.grab_set()

        tk.Label(
            self, text=f'📍  캡처 대상: {label}',
            bg='#1565C0', fg='white', font=('맑은 고딕', 12, 'bold'), pady=12
        ).pack(fill='x')

        tk.Label(
            self,
            text='[캡처 시작] 버튼을 클릭한 뒤 소통메신저의 대상 위치로\n'
                 '마우스를 이동하세요. Enter로 확정, Esc로 취소합니다.',
            font=('맑은 고딕', 10), justify='center', pady=12
        ).pack()

        self.status = tk.Label(
            self, text='아래 버튼을 클릭하여 캡처를 시작하세요.',
            font=('맑은 고딕', 11, 'bold'), fg='#FF9800'
        )
        self.status.pack(pady=6)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        self.start_btn = tk.Button(
            btn_frame, text='캡처 시작', command=self._begin,
            bg='#FF9800', fg='white', font=('맑은 고딕', 11, 'bold'),
            relief='flat', padx=16, pady=6, cursor='hand2'
        )
        self.start_btn.pack(side='left', padx=6)
        tk.Button(
            btn_frame, text='취소', command=self.destroy,
            bg='#9E9E9E', fg='white', font=('맑은 고딕', 10),
            relief='flat', padx=12, pady=6
        ).pack(side='left', padx=6)

    def _begin(self):
        self.start_btn.config(state='disabled')
        self.status.config(text='마우스 위치 확인 중...  Enter 확정 / Esc 취소')
        self.bind('<Return>', lambda _e: self._confirm())
        self.bind('<Escape>', lambda _e: self.destroy())
        self.focus_set()
        self._poll_position()

    def _poll_position(self):
        if not self.winfo_exists():
            return
        pos = pyautogui.position()
        self.status.config(text=f'현재 위치: ({pos.x}, {pos.y})  Enter 확정 / Esc 취소')
        self.after(120, self._poll_position)

    def _confirm(self):
        self._done(pyautogui.position())

    def _done(self, pos):
        self.status.config(text=f'✓ 캡처 완료: ({pos.x}, {pos.y})', fg='green')
        self.on_captured(pos.x, pos.y)
        self.after(1200, self.destroy)


# ─────────────────────────────────────────────
#  도움말 텍스트
# ─────────────────────────────────────────────
_HELP_TEXT = f"""━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  {APP_NAME}  v{APP_VERSION}  —  충북 소통메신저 자동 사용자 선택
  처음 사용자도 따라할 수 있도록 작성되었습니다.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 이 프로그램이 하는 일
──────────────────────────────────────────────────────
  명단(소속기관 + 이름)을 붙여넣거나 파일로 불러오면
  소통메신저에서 아래 3단계를 자동으로 반복합니다.

    1단계: 검색 입력창에 이름 입력 → 검색
    2단계: 검색 결과 첫 번째 항목 클릭 (선택)
    3단계: [사용자 선택] 버튼 클릭 (추가 완료)

  수십~수백 명을 일일이 처리하는 반복 작업을 자동화합니다.


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 시작 전 준비 사항
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  1. 소통메신저에 로그인합니다.
  2. 오른쪽 위의 편지 버튼(✉) 클릭 →
     '쪽지 작성' 또는 '대화하기'를 선택합니다.
  3. 메시지 작성 화면에서 [사용자 선택] 버튼을 클릭합니다.
  4. 사용자 선택 창 상단 탭에서 [전체조직]을 선택합니다.
  5. 소통메신저 창과 {APP_NAME} 창을 나란히 배치하면 편리합니다.

  ※ 실행 중에는 마우스를 움직이지 마세요.
     긴급 중지: 마우스를 화면 왼쪽 위 모서리(0,0)로 빠르게 이동


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 탭 1 — 명단 입력
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  명단 입력 방법은 두 가지입니다.

  [ 방법 A ]  파일 직접 열기
    ① [엑셀 파일 열기] 또는 [HWP 파일 열기] 버튼을 클릭합니다.
    ② 파일을 선택하면 자동으로 입력창에 불러옵니다.
    ③ [명단 추출 →] 버튼을 클릭합니다.

  [ 방법 B ]  복사·붙여넣기
    ① 엑셀·한글에서 소속기관(A열)과 이름(B열)을 선택합니다.
    ② Ctrl+C 로 복사합니다.
    ③ 입력창을 클릭하고 Ctrl+V 로 붙여넣습니다.
    ④ [명단 추출 →] 버튼을 클릭합니다.

  ▶ 추출 결과 목록 활용
    · 항목 더블클릭:  소속기관·이름 직접 수정
    · Delete 키 또는 [선택 항목 삭제 (Del)]:  선택 항목 제거
    · 빨간색 항목:  자동 추가에 실패한 항목


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 탭 2 — 위치 설정  ※ 최초 1회만 설정
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  소통메신저의 클릭 위치 3곳을 {APP_NAME}에 알려주는 과정입니다.
  한 번만 설정하면 이후에는 자동으로 기억합니다.

  공통 캡처 방법:
    [📍 위치 설정] 클릭 → [캡처 시작] 클릭 →
    소통메신저의 해당 위치로 마우스를 이동한 뒤 Enter 키로 확정하세요.
    잘못 눌렀다면 Esc 키로 취소할 수 있습니다.

  [ STEP 1 ]  검색 입력창 위치
    소통메신저 '이름 검색' 입력칸 위로 마우스를 이동한 뒤 Enter.

  [ STEP 2 ]  결과 첫 번째 항목 위치
    임의 이름(예: 홍길동)을 검색한 뒤
    결과 목록의 첫 번째 줄 위로 마우스를 이동한 뒤 Enter.

  [ STEP 3 ]  사용자 선택 버튼 위치
    결과가 보이는 상태에서
    [사용자 선택] 또는 [추가] 버튼 위로 마우스를 이동한 뒤 Enter.

  [ 검색 설정 ]
    · 검색 후 대기 시간: 기본 0.5초.
      인터넷이 느리면 1.0~2.0초로 높이세요.
    · 수동 확인 모드: 동명이인이 있을 때 체크합니다.
      (검색 결과를 직접 확인 후 [▶▶ 계속] 클릭)

  ★ 반드시 [✅ 설정 저장] 버튼을 눌러 저장하세요!


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 탭 3 — 자동 선택 실행
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  ① 소통메신저 [사용자 선택] 창 → [전체조직] 탭을 열어둡니다.
  ② {APP_NAME}에서 [3. 자동 선택] 탭을 클릭합니다.
  ③ [▶ 자동 선택 시작] 버튼을 클릭합니다.
  ④ 명단의 각 이름마다 아래 3단계가 자동으로 실행됩니다.
       1단계: 이름 검색
       2단계: 결과 첫 번째 항목 클릭
       3단계: [사용자 선택] 버튼 클릭
     진행 상황은 로그창에서 확인할 수 있습니다.
       ✓  → 추가 완료
       ✗  → 검색 결과 없음
  ⑤ 완료 시 성공·실패 건수가 표시됩니다.

  ⚠ 긴급 중지
    · 마우스를 화면 왼쪽 위 모서리(0, 0)로 빠르게 이동
    · 또는 [■ 중지] 버튼 클릭


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 자주 묻는 질문
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Q. 소속기관이 '소속없음'으로 표시돼요.
  A. 소속기관명을 인식하지 못한 경우입니다.
     목록에서 해당 항목을 더블클릭하여 직접 수정하세요.

  Q. 프로그램이 엉뚱한 위치를 클릭해요.
  A. 소통메신저 창 위치가 바뀌었을 수 있습니다.
     [2. 위치 설정] 탭에서 3곳을 다시 설정하고 저장하세요.

  Q. 검색은 됐는데 추가가 안 돼요.
  A. [사용자 선택] 버튼 위치(STEP 3)가 잘못 설정되었을 수 있습니다.
     위치를 다시 캡처하고 저장하세요.

  Q. 검색 결과가 아예 없어요.
  A. 소통메신저에 등록되지 않은 사용자입니다.
     해당 항목은 자동으로 ✗ 처리됩니다.

  Q. 너무 빠르게 진행돼서 오류가 생겨요.
  A. [2. 위치 설정] 탭의 '검색 후 대기 시간'을 늘리세요.
     느린 환경: 1.0~2.0초 권장

  Q. 동명이인이 있어서 걱정돼요.
  A. '수동 확인 모드'를 체크하세요.
     검색 후 결과를 직접 확인하고 [▶▶ 계속]을 눌러 진행합니다.

  Q. HWP 파일이 안 열려요.
  A. 한/글이 설치되어 있지 않으면 일부 파일이 열리지 않습니다.
     한글에서 표를 Ctrl+C로 복사 후 입력창에 Ctrl+V로 붙여넣으세요.


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 개발자 정보
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Developed by  송동석
  Teacher  |  Data Analytics  |  App Developer
  협업 및 피드백:  dungst.me@gmail.com

  {APP_NAME}  |  버전 v{APP_VERSION}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"""


# ─────────────────────────────────────────────
#  App
# ─────────────────────────────────────────────


def _version_parts(version: str) -> tuple:
    version = (version or '').strip().lstrip('vV')
    parts = []
    for part in version.split('.'):
        digits = ''.join(ch for ch in part if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts or [0])


def _is_newer_version(latest: str, current: str) -> bool:
    latest_parts = _version_parts(latest)
    current_parts = _version_parts(current)
    size = max(len(latest_parts), len(current_parts))
    latest_parts += (0,) * (size - len(latest_parts))
    current_parts += (0,) * (size - len(current_parts))
    return latest_parts > current_parts

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f'{APP_NAME}  v{APP_VERSION}')
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
        self._check_for_update_async()

    # ── 테마 (충북교육청 블루) ─────────────────
    def _apply_theme(self):
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception as exc:
            logging.info("Tk 테마 적용 실패: %s", exc)
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
            text=f'  {APP_NAME} v{APP_VERSION}  |  Developed by 송동석'
                 '  |  Teacher · Data Analytics · App Developer'
                 '  |  협업: dungst.me@gmail.com  ',
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
        self._refresh_ready_status()

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
            text='📋  명단을 입력하는 방법 (두 가지 중 하나 선택)',
            bg='#E3F2FD', fg='#0D47A1', font=('맑은 고딕', 9, 'bold'), anchor='w'
        ).grid(row=0, column=0, sticky='w', padx=10, pady=(6, 2))

        tk.Label(
            guide,
            text='방법 ①  복사·붙여넣기 — 엑셀·HWP에서 소속기관·이름 범위를 선택 후 복사(Ctrl+C),\n'
                 '                              아래 입력창에 붙여넣기(Ctrl+V) → [명단 추출 →] 클릭\n'
                 '방법 ②  파일 직접 열기  — 아래 [엑셀 파일 열기] 또는 [HWP 파일 열기] 버튼 클릭\n'
                 '\n'
                 '형식 예)  충주중학교    홍길동        (소속기관  이름  순서)',
            bg='#E3F2FD', fg='#333', font=('맑은 고딕', 9), justify='left', anchor='w'
        ).grid(row=1, column=0, sticky='w', padx=10, pady=(0, 8))

        self.ready_status = tk.Label(
            guide, text='', bg='#E3F2FD', fg='#0D47A1',
            font=('맑은 고딕', 9, 'bold'), anchor='w'
        )
        self.ready_status.grid(row=2, column=0, sticky='ew', padx=10, pady=(0, 8))

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
            ('초기화',       self._clear_input, '#E53935'),
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
            bottom_frame, text='선택 항목 삭제 (Del)', command=self._delete_selected,
            bg='#EF6C00', fg='white', activebackground='#E65100',
            relief='flat', font=('맑은 고딕', 9), padx=8, pady=3, cursor='hand2'
        ).grid(row=0, column=0, sticky='w')

        tk.Label(
            bottom_frame,
            text='더블클릭으로 수정  ·  ● 빨간색 = 자동 선택 실패',
            fg='#555', font=('맑은 고딕', 8), bg='#F5F7FA', anchor='e'
        ).grid(row=0, column=1, sticky='e', padx=(0, 4))

        ctx = tk.Menu(self.root, tearoff=0)
        ctx.add_command(label='수정', command=self._edit_item)
        ctx.add_command(label='삭제', command=self._delete_selected)
        self.parsed_list.bind('<Button-3>', lambda e: ctx.tk_popup(e.x_root, e.y_root))
        self.parsed_list.bind('<Double-Button-1>', self._edit_item)
        self.parsed_list.bind('<Delete>', lambda e: self._delete_selected())

    # ── 탭 2: 위치 설정 ────────────────────────
    def _tab_calib(self, frame: ttk.Frame):
        frame.columnconfigure(0, weight=1)

        tk.Label(
            frame,
            text="소통메신저 '사용자 선택' 창을 열어 둔 상태에서 아래 3곳의 위치를 순서대로 설정하세요.\n"
                 "STEP 1: 검색 입력창  ·  STEP 2: 결과 첫 번째 항목  ·  STEP 3: 사용자 선택 버튼\n"
                 "[캡처 시작] 후 마우스를 대상 위치로 이동하고 Enter로 확정합니다. Esc로 취소할 수 있습니다.",
            fg='#555', font=('맑은 고딕', 9), justify='center'
        ).grid(row=0, column=0, sticky='ew', padx=10, pady=(10, 6))

        # STEP 1·2·3: 위치 설정
        pos_frame = ttk.LabelFrame(frame, text='STEP 1 · 2 · 3 — 위치 설정 (순서대로)')
        pos_frame.grid(row=1, column=0, sticky='ew', padx=10, pady=4)
        pos_frame.columnconfigure(1, weight=1)

        for row_i, (label_text, key) in enumerate([
            ('STEP 1  검색 입력창:', 'search_field'),
            ('STEP 2  결과 첫 번째:', 'result_first'),
            ('STEP 3  사용자 선택 버튼:', 'add_button'),
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
            text='수동 확인 모드 권장 — 검색 후 [계속] 버튼을 눌러야 다음으로 진행\n'
                 '(동명이인·검색 결과 오탐이 있을 때 안전: 결과 클릭 → 선택 버튼 클릭을 직접 수행)',
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
            text="소통메신저 '사용자 선택' 창 → 전체조직 탭을 열고 아래 버튼을 누르세요.\n"
                 "각 이름마다 ① 검색 입력  ② 결과 첫 번째 클릭  ③ 선택 버튼 클릭  이 자동으로 반복됩니다.\n"
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

        self.retry_failed_btn = tk.Button(
            btn_frame, text='↻  실패 항목만 다시 실행',
            bg='#795548', fg='white', activebackground='#5D4037',
            relief='flat', font=('맑은 고딕', 10, 'bold'), padx=10, pady=6,
            cursor='hand2', state='disabled', command=self._retry_failed
        )
        self.retry_failed_btn.pack(side='left', padx=4)

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
        frame.rowconfigure(1, weight=1)

        tk.Button(
            frame,
            text='▶  영상으로 사용 방법 보기 (YouTube)',
            bg='#FF0000', fg='white', activebackground='#CC0000',
            relief='flat', font=('맑은 고딕', 10, 'bold'), padx=10, pady=6,
            cursor='hand2',
            command=lambda: webbrowser.open('https://youtu.be/shZnB5NRN5g')
        ).grid(row=0, column=0, pady=(10, 4))

        txt = scrolledtext.ScrolledText(
            frame, font=('맑은 고딕', 10), wrap='word', state='normal'
        )
        txt.grid(row=1, column=0, sticky='nsew', padx=4, pady=4)
        txt.insert('1.0', _HELP_TEXT)
        txt.config(state='disabled')

    def _format_item_label(self, item: dict) -> str:
        return format_item_label(item)

    def _refresh_ready_status(self):
        label = getattr(self, 'ready_status', None)
        if not label:
            return
        names = len(self.names_list)
        positions = '완료' if self.config.is_calibrated() else '미설정'
        manual = '켜짐' if self.config.data.get('manual_confirm', False) else '꺼짐'
        delay = self.config.data.get('search_delay', 0.5)
        label.config(
            text=f'준비 상태  |  명단 {names}명  ·  위치 {positions}  ·  수동 확인 {manual}  ·  대기 {delay}초'
        )

    def _refresh_failed_retry_state(self):
        btn = getattr(self, 'retry_failed_btn', None)
        if not btn:
            return
        has_failed = any(item.get('failure_reason') for item in self.names_list)
        btn.config(state='normal' if has_failed else 'disabled')

    def _rebuild_parsed_list(self):
        self.parsed_list.delete(0, 'end')
        for item in self.names_list:
            self.parsed_list.insert('end', self._format_item_label(item))
            if item.get('failure_reason'):
                self.parsed_list.itemconfig('end', {'bg': '#FFCDD2', 'fg': '#B71C1C'})
            elif not item.get('org'):
                self.parsed_list.itemconfig('end', {'bg': '#E0E0E0', 'fg': '#757575'})
        self._refresh_ready_status()
        self._refresh_failed_retry_state()

    # ── 의존성 확인 ────────────────────────────

    # update check
    def _check_for_update_async(self):
        threading.Thread(target=self._check_for_update_worker, daemon=True).start()

    def _check_for_update_worker(self):
        try:
            req = urllib.request.Request(
                LATEST_RELEASE_API,
                headers={'User-Agent': f'{APP_NAME}/{APP_VERSION}'}
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode('utf-8'))
            latest = (data.get('tag_name') or data.get('name') or '').strip().lstrip('vV')
            url = data.get('html_url') or 'https://github.com/codersongpro/sotong/releases/latest'
            if latest and _is_newer_version(latest, APP_VERSION):
                self.root.after(0, lambda: self._show_update_notice(latest, url))
        except Exception as exc:
            logging.info("업데이트 확인 실패: %s", exc)

    def _show_update_notice(self, latest: str, url: str):
        open_page = messagebox.askyesno(
            '새 버전 알림',
            f'{APP_NAME} 새 버전이 나왔습니다.\n\n'
            f'현재 버전: {APP_VERSION}\n'
            f'최신 버전: {latest}\n\n'
            '다운로드 페이지를 열까요?'
        )
        if open_page:
            webbrowser.open(url)

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
        valid = []
        for item in parsed:
            org = item.get('org', '').replace(' ', '')
            name = item.get('name', '').replace(' ', '')
            if not name:
                fail += 1
                continue
            valid.append({'org': org, 'name': name})

        valid.sort(key=lambda x: x['name'])

        warn_items = []
        for entry in valid:
            org, name = entry['org'], entry['name']
            if not org:
                no_org += 1
                warn_items.append(name)
            self.names_list.append({'org': org, 'name': name})
            ok += 1
        self._rebuild_parsed_list()

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
        self._refresh_ready_status()

    def _clear_input(self):
        self.input_text.delete('1.0', 'end')
        self.parsed_list.delete(0, 'end')
        self.names_list.clear()
        self.parse_status.config(text='')
        self._refresh_ready_status()
        self._refresh_failed_retry_state()

    def _delete_selected(self):
        for i in reversed(self.parsed_list.curselection()):
            del self.names_list[i]
        self._rebuild_parsed_list()
        total = len(self.names_list)
        self.parse_status.config(
            text=f'명단 추출 완료: {total}명', fg='green'
        )

    def _edit_item(self, event=None):
        sel = self.parsed_list.curselection()
        if not sel:
            return
        idx = sel[0]
        item = self.names_list[idx]

        dlg = tk.Toplevel(self.root)
        dlg.title('항목 수정')
        dlg.geometry('360x190')
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.transient(self.root)

        tk.Label(dlg, text='소속기관:', font=('맑은 고딕', 10)).grid(
            row=0, column=0, padx=14, pady=(18, 6), sticky='e')
        org_var = tk.StringVar(value=item.get('org', ''))
        org_entry = tk.Entry(dlg, textvariable=org_var, font=('맑은 고딕', 10), width=22)
        org_entry.grid(row=0, column=1, padx=8, pady=(18, 6), sticky='w')

        tk.Label(dlg, text='이름:', font=('맑은 고딕', 10)).grid(
            row=1, column=0, padx=14, pady=6, sticky='e')
        name_var = tk.StringVar(value=item.get('name', ''))
        tk.Entry(dlg, textvariable=name_var, font=('맑은 고딕', 10), width=22).grid(
            row=1, column=1, padx=8, pady=6, sticky='w')

        def apply():
            new_org = org_var.get().strip().replace(' ', '')
            new_name = name_var.get().strip().replace(' ', '')
            if not new_name:
                messagebox.showwarning('알림', '이름을 입력하세요.', parent=dlg)
                return
            self.names_list[idx] = {'org': new_org, 'name': new_name}
            self._rebuild_parsed_list()
            dlg.destroy()

        btn_frame = tk.Frame(dlg)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=14)
        tk.Button(btn_frame, text='저장', command=apply,
                  bg='#1565C0', fg='white', relief='flat',
                  font=('맑은 고딕', 10), padx=14, pady=4).pack(side='left', padx=6)
        tk.Button(btn_frame, text='취소', command=dlg.destroy,
                  bg='#9E9E9E', fg='white', relief='flat',
                  font=('맑은 고딕', 10), padx=14, pady=4).pack(side='left', padx=6)

        dlg.bind('<Return>', lambda e: apply())
        dlg.bind('<Escape>', lambda e: dlg.destroy())
        org_entry.focus_set()

    # ── 위치 캡처 ──────────────────────────────
    def _do_capture(self, key: str):
        if pyautogui is None:
            messagebox.showerror('오류', 'pyautogui 미설치')
            return

        def on_captured(x, y):
            self.config.data[key + '_x'] = x
            self.config.data[key + '_y'] = y
            self._refresh_calib_labels()

        labels = {
            'search_field': '검색 입력창',
            'result_first': '결과 첫 번째 행',
            'add_button':   '사용자 선택 버튼',
        }
        CaptureDialog(self.root, on_captured, label=labels.get(key, key))

    def _save_calib(self):
        self.config.data['search_delay'] = round(self.delay_var.get(), 1)
        self.config.data['manual_confirm'] = self.manual_var.get()
        self.config.save()
        self.calib_msg.config(text='✅ 설정 저장 완료')
        self.root.after(2000, lambda: self.calib_msg.config(text=''))
        self._refresh_ready_status()

    def _refresh_calib_labels(self):
        for key in ('search_field', 'result_first', 'add_button'):
            x = self.config.data.get(key + '_x')
            y = self.config.data.get(key + '_y')
            lbl = getattr(self, f'lbl_{key}', None)
            if lbl:
                if x is not None and y is not None:
                    lbl.config(text=f'✓ ({x}, {y})', fg='green')
                else:
                    lbl.config(text='미설정', fg='red')
        self._refresh_ready_status()

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
                '알림', '위치 설정 탭에서 검색창·결과·선택 버튼 위치를 모두 설정하세요.'
            )
            return
        self.stop_flag.clear()
        self.continue_event.set()
        for item in self.names_list:
            item.pop('failure_reason', None)
        self._rebuild_parsed_list()
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

    def _retry_failed(self):
        failed = [
            {'org': item.get('org', ''), 'name': item.get('name', '')}
            for item in self.names_list
            if item.get('failure_reason')
        ]
        if not failed:
            messagebox.showinfo('알림', '다시 실행할 실패 항목이 없습니다.')
            return
        self.names_list = failed
        self._rebuild_parsed_list()
        self.parse_status.config(text=f'실패 항목 재실행 준비: {len(failed)}명', fg='green')
        self._start()

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
                    self._log('—  (사용자 없음)\n')
                    self._mark_failed(idx, FAIL_NO_USER)
                    self._update_progress(idx + 1, total)
                    continue

                if manual:
                    self.continue_event.clear()
                    self.root.after(0, lambda n=name: self._show_continue(n))
                    self.continue_event.wait()
                    if self.stop_flag.is_set():
                        self._mark_failed(idx, FAIL_MANUAL_STOP)
                        break
                    ok += 1
                    self._log('✓\n')
                else:
                    result = self._do_select()
                    if result == 'duplicate':
                        fail += 1
                        self._log('⚠  (이미 선택된 사용자)\n')
                        self._mark_failed(idx, FAIL_DUPLICATE)
                        self._update_progress(idx + 1, total)
                        continue
                    ok += 1
                    self._log('✓\n')
            except pyautogui.FailSafeException:
                self._log('\n⚠  긴급 중지 (화면 모서리)\n')
                self._mark_failed(idx, FAIL_MANUAL_STOP)
                self.stop_flag.set()
                break
            except Exception as e:
                fail += 1
                self._log(f'✗  ({e})\n')
                self._mark_failed(idx, failure_reason_from_error(e))

            self._update_progress(idx + 1, total)
            time.sleep(0.1)

        self.root.after(0, lambda: self._done(ok, fail))

    def _do_search(self, search_str: str):
        x = self.config.data['search_field_x']
        y = self.config.data['search_field_y']
        if x is None or y is None:
            raise RuntimeError('좌표 오류: 검색 입력창 위치 미설정')
        delay = self.config.data.get('search_delay', 0.5) * 0.3
        pyautogui.click(x, y)
        time.sleep(delay)
        pyautogui.hotkey('ctrl', 'a')
        pyperclip.copy(search_str)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(delay)
        pyautogui.press('enter')

    def _do_select(self) -> str:
        """3단계: 결과 클릭 → 선택 버튼 클릭 → 팝업 감지"""
        rx = self.config.data['result_first_x']
        ry = self.config.data['result_first_y']
        ax = self.config.data['add_button_x']
        ay = self.config.data['add_button_y']
        if None in (rx, ry, ax, ay):
            raise RuntimeError('좌표 오류: 결과 또는 선택 버튼 위치 미설정')
        time.sleep(0.2)
        pyautogui.click(rx, ry)
        time.sleep(0.15)
        before = self._snapshot_dialogs()
        pyautogui.click(ax, ay)
        time.sleep(0.4)
        new_hwnds = self._snapshot_dialogs() - before
        if new_hwnds and self._is_duplicate_popup(new_hwnds):
            pyautogui.press('enter')
            time.sleep(0.15)
            return 'duplicate'
        return 'ok'

    def _snapshot_dialogs(self) -> set:
        """현재 열려있는 #32770 다이얼로그 핸들 집합 반환."""
        try:
            import win32gui
            found = set()
            def cb(hwnd, _):
                if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                    if win32gui.GetClassName(hwnd) == '#32770':
                        found.add(hwnd)
            win32gui.EnumWindows(cb, None)
            return found
        except Exception as exc:
            logging.debug("다이얼로그 목록 확인 실패: %s", exc)
            return set()

    def _is_duplicate_popup(self, hwnds: set) -> bool:
        """새 다이얼로그의 텍스트에 '이미'가 포함되면 중복 팝업으로 판정."""
        try:
            import win32gui
            for hwnd in hwnds:
                texts = [win32gui.GetWindowText(hwnd)]
                def collect(h, _):
                    t = win32gui.GetWindowText(h)
                    if t:
                        texts.append(t)
                win32gui.EnumChildWindows(hwnd, collect, None)
                if any('이미' in t for t in texts):
                    return True
        except Exception as exc:
            logging.debug("중복 팝업 확인 실패: %s", exc)
        return False

    def _has_result(self) -> bool:
        """검색 결과 첫 항목 위치의 픽셀 밝기로 결과 유무 자동 판단."""
        x = self.config.data.get('result_first_x')
        y = self.config.data.get('result_first_y')
        if x is None or y is None:
            raise RuntimeError('좌표 오류: 결과 위치 미설정')
        try:
            r, g, b = pyautogui.pixel(x, y)[:3]
            # 결과 없으면 배경이 흰색(240+) → 결과 있으면 텍스트로 인해 더 어두움
            return not (r > 240 and g > 240 and b > 240)
        except Exception as exc:
            raise RuntimeError(f'좌표 오류: 결과 위치 확인 실패 ({exc})') from exc

    def _show_continue(self, name: str):
        name = name or ''
        self.status_var.set(
            f'수동 선택 대기: {name}  →  소통메신저에서 결과 클릭 → 선택 버튼 클릭 후 [계속] 버튼'
        )
        self.continue_btn.config(state='normal')

    def _done(self, ok: int, fail: int):
        self.start_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.continue_btn.config(state='disabled')
        self._refresh_failed_retry_state()
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

    def _mark_failed(self, idx: int, reason: str):
        def apply():
            if idx >= len(self.names_list):
                return
            self.names_list[idx]['failure_reason'] = reason
            self.parsed_list.delete(idx)
            self.parsed_list.insert(idx, self._format_item_label(self.names_list[idx]))
            self.parsed_list.itemconfig(idx, {'bg': '#FFCDD2', 'fg': '#B71C1C'})
            self._refresh_failed_retry_state()
        self.root.after(0, apply)

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

