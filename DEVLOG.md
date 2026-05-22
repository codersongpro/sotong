# 소통픽 개발 대화 로그

> **프로젝트**: 소통픽 (충북 소통메신저 자동 사용자 선택)  
> **브랜치**: `claude/improve-ui-developer-info-4Z7b4`  
> **PR**: https://github.com/codersongpro/sotong/pull/9  
> **개발자**: 송동석 · dungst.me@gmail.com

---

## 목차

1. [주요 상수 및 구조](#1-주요-상수-및-구조)
2. [명단 수정·공백제거·학교급 매칭·캡처 UI 개선](#2-명단-수정공백제거학교급-매칭캡처-ui-개선)
3. [버튼 색상·Delete 키·입력 방법 안내](#3-버튼-색상delete-키입력-방법-안내)
4. [프로그램명 '소통픽'으로 변경](#4-프로그램명-소통픽으로-변경)
5. [index.html Cursor 디자인 시스템 적용](#5-indexhtml-cursor-디자인-시스템-적용)
6. [build.bat 오류 수정](#6-buildbat-오류-수정)
7. [3단계 자동 선택 알고리즘](#7-3단계-자동-선택-알고리즘)
8. [안내 문구 전면 수정](#8-안내-문구-전면-수정)
9. [명단 가나다 오름차순 정렬](#9-명단-가나다-오름차순-정렬)
10. [DB에 없는 기관명 처리](#10-db에-없는-기관명-처리)
11. [중복 감지 오탐 수정 (1차)](#11-중복-감지-오탐-수정-1차)
12. [솔강초등학교 DB 추가](#12-솔강초등학교-db-추가)
13. [중복 감지 오탐 수정 (2차 — 텍스트 검증)](#13-중복-감지-오탐-수정-2차--텍스트-검증)
14. [소속없음 항목 음영 표시](#14-소속없음-항목-음영-표시)
15. [YouTube 튜토리얼 링크 추가](#15-youtube-튜토리얼-링크-추가)

---

## 1. 주요 상수 및 구조

**요청**: 프로그램 이름과 버전을 상수로 관리하고 자동 표기

**변경 파일**: `main.py`

```python
APP_NAME    = '소통픽'
APP_VERSION = '1.5'
```

- 윈도우 타이틀, 개발자 표시줄, 도움말 텍스트에 f-string으로 자동 반영
- 개발자 표시줄: `소통픽 v1.5 | Developed by 송동석 | Teacher · Data Analytics · App Developer | 협업: dungst.me@gmail.com`

---

## 2. 명단 수정·공백제거·학교급 매칭·캡처 UI 개선

**요청 (6가지)**

1. 추출된 명단을 사용자가 직접 수정할 수 있게 (더블클릭)
2. 추출된 소속기관·이름에 공백 제거
3. 초/중/고 학교급을 명확히 확인해 소속기관 검색 (오송솔미초 → 유치원 방지)
4. 위치 캡처 안내 팝업 개선 (마우스 이동 시간 확보)
5. 캡처 후 3초 카운트다운
6. 더블클릭 추가 신뢰성 개선

**주요 코드 변경**

### 더블클릭 수정 다이얼로그 (`_edit_item`)
```python
def _edit_item(self, event=None):
    sel = self.parsed_list.curselection()
    if not sel: return
    idx = sel[0]
    item = self.names_list[idx]
    dlg = tk.Toplevel(self.root)
    dlg.title('항목 수정')
    dlg.geometry('360x190')
    # org Entry, name Entry, 저장/취소 버튼
    # 저장 시 공백 제거, names_list 및 listbox 갱신
```

### 공백 제거 (`_parse`)
```python
org = item.get('org', '').replace(' ', '')
name = item.get('name', '').replace(' ', '')
```

### 학교급 구분 퍼지 매칭 (`lookup_org`, cutoff=0.75)
```python
# 초/중/고 suffix가 있으면 같은 학교급 키만 후보로 제한
type_keys = [k for k in keys if _split_school_suffix(k)[1] == suffix]
matches = difflib.get_close_matches(s, type_keys, n=1, cutoff=0.75)
```

### CaptureDialog — 3초 카운트다운
```python
class CaptureDialog(tk.Toplevel):
    def _run(self):
        self.after(0, self.iconify)
        for i in range(3, 0, -1):
            self.after(0, lambda i=i: self.status.config(text=f'{i}초 후 캡처...'))
            time.sleep(1)
        pos = pyautogui.position()
        self.after(0, lambda: self._done(pos))
```

---

## 3. 버튼 색상·Delete 키·입력 방법 안내

**요청**

- 초기화 버튼(빨간)과 선택 항목 삭제 버튼(주황) 색상 구분
- Delete 키로 추출 명단 삭제
- 명단 입력 두 가지 방법(파일/복붙) 프로그램 내 안내

**변경**

```python
('초기화', self._clear_input, '#E53935'),       # 빨간
# 선택 항목 삭제: bg='#EF6C00'                  # 주황
```

```python
self.parsed_list.bind('<Delete>', lambda e: self._delete_selected())
```

---

## 4. 프로그램명 '소통픽'으로 변경

**요청**: 프로그램 이름을 '소통픽'으로, 빌드 파일명도 변경, 버전 자동 표기

**변경 파일**: `main.py`, `build.bat`, `ChungbukMessenger_AutoSelect.spec`

- `APP_NAME = '소통픽'`
- build.bat: `--name "sotongpick"` (ASCII — CP949 인코딩 문제 방지)
- spec: `name='sotongpick'`

---

## 5. index.html Cursor 디자인 시스템 적용

**요청**: Cursor 디자인 시스템 스펙으로 index.html 재설계

**디자인 토큰**

```css
:root {
  --canvas:    #f7f7f4;   /* warm cream */
  --ink:       #26251e;
  --primary:   #f54e00;   /* Cursor Orange */
  --hairline:  #e6e5e0;
}
```

**주요 섹션**

- Nav: sticky, 헤어라인 하단 경계
- Hero: `소통픽` 대형 타이포, 다운로드 CTA
- Features: 4-카드 그리드 (엑셀, HWP, 복붙, 오타보정)
- How-to: 탭 UI (①명단입력 / ②위치설정 / ③자동선택)
- FAQ: `<details>/<summary>` 아코디언
- Developer 카드: 송동석 · v1.5
- Font: Inter + JetBrains Mono (Google Fonts)

---

## 6. build.bat 오류 수정

**증상**: CMD 실행 시 `'N'은(는) 내부...`, `FIND: 매개 변수 형식이 틀립니다`, `do은(는) 예상되지 않았습니다`

**원인 (2가지)**

| 원인 | 증상 | 수정 |
|------|------|------|
| LF 줄바꿈 | Windows CMD가 파싱 실패 | Python binary write로 `\n` → `\r\n` 변환 |
| `--name "소통픽"` (한글 UTF-8) | CP949 CMD가 바이트 손상 | `--name "sotongpick"` (ASCII)으로 변경 |

---

## 7. 3단계 자동 선택 알고리즘

**요청**

> 위치 설정 시 캡처 시간을 3초로 설정하고, 프로그램 동작 알고리즘을 3단계로:
> 1단계: 검색 입력창에서 검색
> 2단계: 검색된 결과를 한 번 클릭
> 3단계: 사용자 선택 버튼 한 번 클릭
> 검색 후 대기 시간 기본 0.5초

**알고리즘 이력**

| 버전 | 방식 |
|------|------|
| 초기 | `pyautogui.doubleClick(x, y)` — 불안정 |
| v2 | 단일 클릭 × 2 + 0.1s sleep |
| v3 | 추가 버튼만 클릭 |
| v4 | 더블클릭 복원 |
| **현재** | result_first 클릭 → add_button 클릭 (3단계) |

**핵심 코드**

```python
# Config
class Config:
    DEFAULTS = {
        'search_field_x': None, 'search_field_y': None,
        'result_first_x': None, 'result_first_y': None,
        'add_button_x':   None, 'add_button_y':   None,
        'search_delay': 0.5,
        'manual_confirm': False,
    }

# _do_select
def _do_select(self) -> str:
    time.sleep(0.2)
    pyautogui.click(rx, ry)    # STEP 2: 결과 첫 번째 클릭
    time.sleep(0.15)
    before = self._snapshot_dialogs()
    pyautogui.click(ax, ay)    # STEP 3: 사용자 선택 버튼 클릭
    time.sleep(0.4)
    new_hwnds = self._snapshot_dialogs() - before
    if new_hwnds and self._is_duplicate_popup(new_hwnds):
        pyautogui.press('enter')
        return 'duplicate'
    return 'ok'
```

---

## 8. 안내 문구 전면 수정

**요청**: 안내 문구도 전부 수정, 웹페이지 포함해서

**변경 위치**

| 위치 | 변경 내용 |
|------|-----------|
| `_HELP_TEXT` | 3단계 알고리즘 기준으로 전면 재작성 |
| 탭 2 헤더 | "STEP 1·2·3 — 3곳 위치 순서대로 설정" |
| 탭 3 헤더 | "① 검색 입력 ② 결과 클릭 ③ 선택 버튼 자동 반복" |
| 수동 확인 체크박스 | "더블클릭" → "결과 클릭 → 선택 버튼 클릭" |
| `_show_continue` 상태 | "더블클릭 후" → "결과 클릭 → 선택 버튼 클릭 후 [계속]" |
| index.html 탭 3 설명 | 3단계 흐름 명시 |
| index.html FAQ 동명이인 | "직접 선택" → "결과 클릭 → 선택 버튼 → [계속]" |

---

## 9. 명단 가나다 오름차순 정렬

**요청**: 명단 추출 시 사용자 이름 기준 가나다 한글 오름차순 자동 정렬

**변경** (`_parse`)

```python
valid.sort(key=lambda x: x['name'])
```

> Python 문자열 정렬은 유니코드 코드포인트 순서를 따르며, 한글(AC00–D7A3)은 연속 블록이므로 별도 라이브러리 없이 가나다순 정렬이 적용됩니다.

---

## 10. DB에 없는 기관명 처리

**요청**: 입력한 기관명이 DB에 없을 경우 사용자가 직접 입력한 기관명 그대로 반영

**분석**: `parse_input` 내 `pending_org` 처리에서 이미 DB 미매칭 시 원본 토큰을 사용하고 있었음

```python
# parse_input 내 — 이미 구현된 로직
resolved = lookup_org(tok)
pending_org = resolved if resolved else tok   # DB 없으면 원본 그대로
```

`best_org_from`은 DB 매칭 실패 시 빈 문자열을 반환하는 구조였으므로, 원본 보존 경로 재확인 후 종료.

---

## 11. 중복 감지 오탐 수정 (1차)

**증상**: 정상 추가된 사용자도 "이미 선택된 사용자"로 표시

**원인**: 기존 `_popup_appeared()`가 소통메신저 자체 다이얼로그(`#32770`)를 항상 감지

**1차 수정**: 클릭 전/후 핸들 집합 비교 (스냅샷 방식)

```python
def _snapshot_dialogs(self) -> set:
    import win32gui
    found = set()
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            if win32gui.GetClassName(hwnd) == '#32770':
                found.add(hwnd)
    win32gui.EnumWindows(cb, None)
    return found
```

**로그 텍스트 변경**: `'✗  (검색 결과 없음)'` → `'—  (사용자 없음)'`

---

## 12. 솔강초등학교 DB 추가

**요청**: DB에 솔강초등학교 추가

```python
("솔강초등학교", "솔강초등학교"),
("솔강초",       "솔강초등학교"),
```

> 기존에 `솔강초등학교병설유치원`, `솔강중학교`는 DB에 있었으나 `솔강초등학교` 본교가 누락되어 있었음

---

## 13. 중복 감지 오탐 수정 (2차 — 텍스트 검증)

**증상** (스크린샷 확인): 3~6번째 사용자가 오른쪽 "선택된 사용자" 목록에 정상 추가됐으나 "이미 선택된 사용자"로 표시

**원인**: 소통메신저가 사용자 추가 후 내부적으로 다이얼로그 핸들을 재생성 → 새 핸들이 스냅샷 diff에 잡혀 오탐

**2차 수정**: 새 다이얼로그의 자식 컨트롤 텍스트에 `"이미"` 포함 여부로 최종 판정

```python
def _is_duplicate_popup(self, hwnds: set) -> bool:
    import win32gui
    for hwnd in hwnds:
        texts = [win32gui.GetWindowText(hwnd)]
        def collect(h, _):
            t = win32gui.GetWindowText(h)
            if t: texts.append(t)
        win32gui.EnumChildWindows(hwnd, collect, None)
        if any('이미' in t for t in texts):
            return True
    return False
```

---

## 14. 소속없음 항목 음영 표시

**요청**: 명단 추출 시 소속이 없는 항목을 음영으로 표시

**변경** (`_parse`)

```python
if not org:
    self.parsed_list.itemconfig('end', {'bg': '#E0E0E0', 'fg': '#757575'})
```

- 배경: `#E0E0E0` (회색)
- 글자: `#757575` (중간 회색)
- 자동 선택 실패 시 `_mark_failed`의 빨간 음영으로 덮어씌워짐 (우선순위 정상)

---

## 15. YouTube 튜토리얼 링크 추가

**요청**: 프로그램 안과 웹페이지에 사용 방법 영상 안내 (`https://youtu.be/shZnB5NRN5g`)

**main.py** — 사용 방법 탭 상단에 버튼 추가

```python
import webbrowser

tk.Button(
    frame,
    text='▶  영상으로 사용 방법 보기 (YouTube)',
    bg='#FF0000', fg='white',
    command=lambda: webbrowser.open('https://youtu.be/shZnB5NRN5g')
).grid(row=0, column=0, pady=(10, 4))
```

**index.html** — how-to 섹션 상단에 링크 버튼 추가

```html
<a href="https://youtu.be/shZnB5NRN5g" target="_blank" rel="noopener"
   style="background:#FF0000;color:#fff;padding:10px 20px;border-radius:6px;font-weight:600;">
  ▶ 영상으로 사용 방법 보기 (YouTube)
</a>
```

---

## 커밋 히스토리

```
80eb819  Add YouTube tutorial link to help tab and webpage
d9f9bab  Shade no-org entries in parsed list (grey bg+fg)
19c5458  Fix false-positive duplicate: verify popup text before flagging as duplicate
7b24d1c  Add 솔강초등학교 to org DB
bc1bb9b  Fix false-positive duplicate detection; show '사용자 없음' for no results
2368903  Sort parsed name list in ascending Korean order by name
f951d0f  Update all guidance text for 3-step algorithm (검색→결과 클릭→선택 버튼)
15d643b  3단계 자동 선택 알고리즘 구현 (검색→결과클릭→선택버튼)
a693a0c  build.bat: CRLF 변환, exe 이름 ASCII화 (sotongpick)
2805ac0  사용방법 도움말 전면 개선: 박스 제거, 문구 정확화
3343b48  프로그램명 '소통픽'으로 변경, 버전 상수화, 기관 수 안내 제거
c8ad4a5  캡처 3초로 단축, STEP2 result_first 복원, index.html Cursor 디자인 시스템 적용
516874c  추가 버튼 클릭 방식으로 변경 (더블클릭 → 단일클릭)
3356685  UI 개선: 버튼 색상 구분, Delete 키 삭제, 명단 입력 안내 개선
6ff63be  Add .gitignore to exclude Python bytecode and build artifacts
2f6b417  6가지 기능 개선: 명단 수정, 공백제거, 학교급 매칭, 캡처UI, 더블클릭 수정
```

---

## 파일 구조

```
sotong/
├── main.py                          # 메인 애플리케이션 (tkinter GUI + 자동화)
├── index.html                       # 안내 웹페이지 (Cursor 디자인 시스템)
├── build.bat                        # Windows 빌드 스크립트 (PyInstaller)
├── ChungbukMessenger_AutoSelect.spec # PyInstaller spec
├── .gitignore
└── DEVLOG.md                        # 이 파일
```

## 빌드 방법

1. `build.bat` 더블클릭 실행
2. Python 자동 탐지 및 설치 (없으면 Python 3.11 자동 다운로드)
3. 필요 패키지 자동 설치 (`pyautogui`, `pyperclip`, `openpyxl`, `pywin32`, `pyinstaller`)
4. `dist/sotongpick.exe` 생성
