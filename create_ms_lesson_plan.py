"""
Microsoft 트랙 수업계획서 생성 스크립트 (완전 새 버전)

설계 원칙:
- 30분 성공 경험: 매 회차 시작 30분 안에 수강생 화면에 결과물
- 회차당 핵심 도구 1~2개 (5개 도구 X -> 1개 깊이 O)
- 체험 먼저, 이름은 나중에
- 웹 브라우저 우선 (Copilot CLI -> 웹 UI)
- 80% 완성 템플릿 제공
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# STYLING HELPERS
# ============================================================

MS_BLUE = '2B579A'
MS_DARK = '1B3A62'
ACCENT_GREEN = '107C10'
ACCENT_ORANGE = 'D83B01'
GRAY = '666666'


def create_doc():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style.font.size = Pt(9)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    return doc


def styled_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = 'Malgun Gothic'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    return h


def para(doc, text, bold=False, size=9, color=None, italic=False, align=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = 'Malgun Gothic'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(1)
    if align:
        p.alignment = align
    return p


def bullet(doc, text, bold_prefix='', size=9, indent_level=0):
    p = doc.add_paragraph(style='List Bullet')
    if bold_prefix:
        run_b = p.add_run(bold_prefix)
        run_b.bold = True
        run_b.font.name = 'Malgun Gothic'
        run_b.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
        run_b.font.size = Pt(size)
    run = p.add_run(text)
    run.font.name = 'Malgun Gothic'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    run.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(1)
    return p


def make_table(doc, headers, rows, header_color=MS_BLUE, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(8)
        run.font.name = 'Malgun Gothic'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), header_color)
        shading.set(qn('w:val'), 'clear')
        cell._tc.get_or_add_tcPr().append(shading)
        run.font.color.rgb = RGBColor(255, 255, 255)

    # Data rows
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(8)
            run.font.name = 'Malgun Gothic'
            run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')

    # Column widths
    if col_widths:
        for row in table.rows:
            for i, width in enumerate(col_widths):
                if i < len(row.cells):
                    row.cells[i].width = Cm(width)

    return table


def info_box(doc, title, items, color_rgb=(43, 87, 154)):
    """색상 강조 정보 박스 (제목 + 불릿 리스트)"""
    para(doc, title, bold=True, size=10, color=color_rgb)
    for item in items:
        bullet(doc, item, size=8)


def section_divider(doc):
    para(doc, '─' * 60, size=7, color=(180, 180, 180))


# ============================================================
# SESSION DATA - Microsoft 트랙 (완전 새 버전)
# ============================================================

sessions = [
    # ── 1회차 ──────────────────────────────────────────
    {
        'num': 1,
        'title': 'AI 첫 만남 + 보안 — 설치 없이 바로 체험',
        'subtitle': 'ChatGPT/Copilot 웹에서 바로 체험 + 기업 데이터 보안 기초',
        'goal': '웹 브라우저만으로 AI를 직접 체험하고, 회사 데이터를 안전하게 다루는 기준을 세운다',
        'tools': ['Copilot 웹 (copilot.microsoft.com)', 'ChatGPT 웹 (chat.openai.com)'],
        'tools_note': '설치 없음! 웹 브라우저만 사용',
        'success_30min': '"AI가 이메일 초안을 1분 만에 써줬다" — 수강생 화면에 즉시 결과물',
        'killer_demo': '완성된 자금일보 자동화 + 대시보드를 먼저 시연 → "6주 뒤 여러분이 만들 결과물입니다"',
        'timetable': [
            ['0:00~0:15', '킬러 데모 + 과정 소개',
             '완성된 자금일보 자동화 시연 (버튼 하나로 엑셀 완성)\n'
             '완성된 HTML 대시보드 시연 (팀원이 브라우저로 접속)\n'
             '"6주 뒤 여러분이 직접 만들 결과물입니다"\n'
             '6회차 로드맵 한눈에 보기',
             '강의 + 시연'],
            ['0:15~0:50', '[30분 성공] AI에게 이메일 써달라고 하기',
             'Copilot 웹 접속 (copilot.microsoft.com)\n'
             '"거래처에 보낼 입금 요청 이메일 써줘" 입력\n'
             '→ 30초 만에 완성된 이메일 확인!\n'
             '같은 요청을 ChatGPT에도 해보기\n'
             '→ 두 결과 비교 ("어디가 더 좋아요?")',
             '미니 실습'],
            ['0:50~1:40', '메인 실습: AI로 업무 문서 3종 만들기',
             '[시연 10분] 회의록 요약 / 보고서 초안 / 데이터 정리 시연\n'
             '[따라하기 40분]\n'
             '(1) 회의록 초안: "다음 회의 내용을 요약해줘" + 샘플 텍스트 붙여넣기\n'
             '(2) 보고서 초안: "월간 실적 보고서 초안 써줘" + 가상 데이터 제공\n'
             '(3) 데이터 정리: "이 데이터를 표로 정리해줘" + CSV 붙여넣기',
             '시연 + 실습'],
            ['1:40~2:00', '휴식', '', ''],
            ['2:00~2:50', '심화: 기업 데이터 보안 기초',
             '[시연 10분] 삼성 소스코드 유출 사건 + AI에 넣으면 안 되는 것\n'
             '[따라하기 40분]\n'
             '데이터 4등급 분류 실습 (80% 완성 템플릿 제공):\n'
             '  (1) 자유 — 공개 정보, 뉴스\n'
             '  (2) 클라우드 OK — 일반 업무 문서\n'
             '  (3) 로컬 AI만 — 내부 재무 데이터\n'
             '  (4) AI 금지 — 개인정보, 영업비밀\n'
             '우리 부서 데이터를 4등급에 직접 분류하기',
             '시연 + 실습'],
            ['2:50~3:40', '응용: AI 도구별 보안 지도 만들기',
             'Copilot 웹: 무료 vs Business 학습 정책 차이 정리\n'
             'ChatGPT: Free vs Plus vs Team 데이터 처리 정책\n'
             'Ollama(5회차 예고): 인터넷 없이 쓰는 AI\n'
             '→ 80% 완성 "AI 보안 지도" 템플릿에 우리 회사 기준 채우기\n'
             'AI 코드 검증 4단계 소개: 말하기→받기→검증하기→실행하기',
             '자유 실습'],
            ['3:40~4:00', '정리 + 월요일 액션 플랜',
             '오늘 배운 것 정리 (1분 복습)\n'
             '월요일 액션: "출근해서 Copilot 웹으로 이메일 1통 써보기"\n'
             '다음 회차 예고: Excel + VBA로 자금일보 자동화',
             '정리'],
        ],
        'output': [
            'AI가 작성한 업무 문서 3종 (이메일, 회의록, 보고서 초안)',
            '데이터 4등급 분류 기준표 (우리 부서 버전)',
            'AI 보안 지도 (도구별 보안 조건 정리)',
        ],
        'monday_action': '출근해서 Copilot 웹(copilot.microsoft.com)으로 이메일 1통 작성해보기',
        'plan_b': [
            'Copilot 웹 접속 불가 시 → ChatGPT 웹 사용 (동일 실습 가능)',
            '회사 네트워크에서 두 사이트 모두 차단 시 → 강사가 핫스팟 제공 or 시연 모드',
        ],
        'it_checklist': [
            'copilot.microsoft.com 접속 가능 여부',
            'chat.openai.com 접속 가능 여부',
            '수강생 PC에 Chrome/Edge 브라우저 설치 여부',
        ],
    },

    # ── 2회차 ──────────────────────────────────────────
    {
        'num': 2,
        'title': '자금일보 자동화 — Excel VBA 하나만 깊이',
        'subtitle': '80% 완성 VBA 템플릿 + AI가 나머지 20% 채우기',
        'goal': 'AI 웹에서 VBA 코드를 생성하고, Excel 버튼 하나로 자금일보를 완성한다',
        'tools': ['Excel + VBA (AI가 웹에서 코드 생성 → 복붙)'],
        'tools_note': '코드는 Copilot/ChatGPT 웹에서 생성 → Excel에 복사-붙여넣기 (CLI 불필요)',
        'success_30min': '"버튼 하나로 자금일보가 완성됐다" — 강사 제공 80% 완성 VBA + AI가 나머지 작성',
        'killer_demo': None,
        'timetable': [
            ['0:00~0:15', '도입: VBA가 뭔가요?',
             'VBA = "엑셀 안에 숨어 있는 자동화 도구"\n'
             '버튼 하나로 자금일보 완성 시연 (킬러 데모 복습)\n'
             '"오늘 이걸 직접 만듭니다"',
             '강의 + 시연'],
            ['0:15~0:50', '[30분 성공] 80% 완성 VBA로 자금일보 1분 만에',
             '강사 제공 Excel 파일 열기 (80% 완성 VBA 포함)\n'
             '→ "매크로 실행" 버튼 클릭\n'
             '→ 은행 CSV 데이터가 자금일보 양식에 자동 채워짐!\n'
             '"방금 실행한 게 VBA 매크로입니다"\n'
             '코드 열어보기 (Alt+F11) — "이 코드를 AI가 써준 겁니다"',
             '미니 실습'],
            ['0:50~1:40', '메인 실습: AI에게 VBA 코드 요청하기',
             '[시연 10분] Copilot 웹에서 VBA 코드 요청하는 법\n'
             '[따라하기 40분]\n'
             '(1) Copilot 웹 열기\n'
             '(2) "은행 거래내역 CSV를 읽어서 자금일보 양식에 채우는\n'
             '     Excel VBA 코드 만들어줘" 입력\n'
             '(3) 생성된 코드 복사 → Excel VBA 편집기(Alt+F11)에 붙여넣기\n'
             '(4) 실행 → 결과 확인\n'
             '(5) 안 되면? AI에게 "이런 에러가 났어" 다시 요청\n'
             'AI 코드 검증: XMLHTTP/외부 호출 없는지 확인하기',
             '시연 + 실습'],
            ['1:40~2:00', '휴식', '', ''],
            ['2:00~2:50', '심화: 채권연령표 자동화',
             '[시연 10분] 채권연령표 완성본 시연\n'
             '[따라하기 40분]\n'
             'AI에게 "매출채권 경과일수 계산 + 연령 구간 분류\n'
             'VBA 매크로 만들어줘" 요청\n'
             '→ 코드 복사 → 실행 → 결과 확인\n'
             '분류 기준 변경해보기 (30/60/90/120일)\n'
             'Power Query 맛보기: 여러 CSV 한번에 합치기 (시연 수준)',
             '시연 + 실습'],
            ['2:50~3:40', '응용: Document AI 맛보기 + 자유 실습',
             '업스테이지 Document Parse 데모 시연 (10분)\n'
             '→ PDF 영수증을 자동으로 표로 변환하는 과정 보기\n'
             '나머지 시간: 본인 업무에 맞는 VBA 매크로 만들기\n'
             '→ AI에게 요청 → 코드 복사 → 실행 → 수정 반복',
             '시연 + 자유 실습'],
            ['3:40~4:00', '정리 + 월요일 액션 플랜',
             '오늘 배운 것: "AI가 코드를 써주고, 나는 실행만 한다"\n'
             'VBA 코드 검증 체크리스트 배포\n'
             '월요일 액션: "자금일보 VBA 파일을 내 업무 폴더에 저장하고 1번 실행해보기"\n'
             '다음 회차 예고: Power Automate로 자동화 확장',
             '정리'],
        ],
        'output': [
            '자금일보 VBA 매크로 (버튼 하나로 실행)',
            '채권연령표 VBA 매크로',
            'VBA 코드 검증 체크리스트',
        ],
        'monday_action': '자금일보 VBA 파일을 업무 폴더에 저장하고 실제 데이터로 1번 실행해보기',
        'plan_b': [
            'VBA 매크로 보안 설정으로 실행 불가 시 → IT팀 매크로 허용 요청 가이드 제공',
            'Alt+F11 진입 불가 시 → 강사가 USB로 매크로 포함 파일 배포',
            'Copilot 웹 접속 불가 시 → ChatGPT 웹 사용',
        ],
        'it_checklist': [
            'Excel 매크로(VBA) 실행 가능 여부 (보안 설정 확인)',
            'Alt+F11로 VBA 편집기 접근 가능 여부',
            'copilot.microsoft.com 또는 chat.openai.com 접속 가능 여부',
        ],
    },

    # ── 3회차 ──────────────────────────────────────────
    {
        'num': 3,
        'title': 'AI가 혼자 판단하게 만들기 — 노코드 자동화',
        'subtitle': 'Power Automate Desktop 실습 먼저 → "방금 한 게 Agent입니다"',
        'goal': '노코드 드래그앤드롭으로 자동화 플로우를 만들고, Agent/MCP 개념을 체험으로 이해한다',
        'tools': ['Power Automate Desktop (노코드 드래그앤드롭)'],
        'tools_note': 'Windows 10/11에 기본 포함. 추가 설치 불필요 (Microsoft 계정 필요)',
        'success_30min': '"이메일이 오면 자동으로 폴더에 저장됐다" — 드래그앤드롭만으로 완성',
        'killer_demo': None,
        'timetable': [
            ['0:00~0:15', '도입: 반복 업무의 비용',
             '"매일 30분 × 250일 = 연간 125시간"\n'
             '반복 업무 예시: 이메일 첨부파일 저장, 파일 이름 변경, 보고서 복사\n'
             'Power Automate Desktop 시연: 이메일 → 폴더 자동 저장',
             '강의 + 시연'],
            ['0:15~0:50', '[30분 성공] 이메일 첨부파일 자동 저장',
             'Power Automate Desktop 실행 (시작 메뉴에서 검색)\n'
             '→ "새 플로우" 만들기\n'
             '→ "Outlook에서 메일 가져오기" 블록 드래그\n'
             '→ "첨부파일 저장" 블록 드래그\n'
             '→ 실행! 이메일 첨부파일이 지정 폴더에 자동 저장!\n'
             '"방금 마우스만으로 자동화를 만들었습니다"',
             '미니 실습'],
            ['0:50~1:40', '메인 실습: 파일 정리 자동화',
             '[시연 10분] 폴더 감시 → 파일 분류 자동화 시연\n'
             '[따라하기 40분]\n'
             '(1) "폴더의 파일 가져오기" 블록\n'
             '(2) "조건 분기" 블록: 파일명에 "견적"이 포함되면\n'
             '(3) "파일 이동" 블록: → 견적서 폴더로 이동\n'
             '(4) 실행 → 파일이 자동으로 분류!\n'
             '조건 추가: "계약서", "송장" 등 확장하기',
             '시연 + 실습'],
            ['1:40~2:00', '휴식', '', ''],
            ['2:00~2:50', '심화: Excel 데이터 자동 처리',
             '[시연 10분] Excel 읽기 → 가공 → 저장 시연\n'
             '[따라하기 40분]\n'
             'Power Automate에서 Excel 열기\n'
             '→ 특정 열 데이터 읽기\n'
             '→ 조건 분기 (금액이 100만원 이상이면)\n'
             '→ 결과를 새 Excel 파일에 저장\n'
             '→ Teams 웹훅으로 "처리 완료" 알림 보내기 (맛보기)',
             '시연 + 실습'],
            ['2:50~3:40', '응용: "방금 한 게 Agent입니다"',
             '개념 소개 (10분):\n'
             '  Agent = "스스로 판단하는 AI 도우미"\n'
             '  MCP = "AI 도구끼리 대화하는 규칙"\n'
             '  "방금 만든 플로우가 바로 Agent의 시작입니다"\n'
             '자유 실습 (40분):\n'
             '  본인 업무에 맞는 플로우 만들기\n'
             '  예시: 매일 아침 보고서 파일 자동 복사 + 이름 변경\n'
             '  예시: 특정 폴더에 파일 들어오면 → Teams 알림',
             '강의 + 자유 실습'],
            ['3:40~4:00', '정리 + 월요일 액션 플랜',
             '오늘 배운 것: "마우스 드래그만으로 자동화를 만든다"\n'
             'Agent의 핵심: 입력→판단→실행→알림\n'
             '월요일 액션: "매일 하는 반복 업무 1개를 Power Automate로 자동화해보기"\n'
             '다음 회차 예고: AI로 대시보드 만들고 팀에 공유하기',
             '정리'],
        ],
        'output': [
            '이메일 첨부파일 자동 저장 플로우',
            '파일 자동 분류 플로우',
            'Excel 데이터 자동 처리 플로우',
        ],
        'monday_action': '매일 하는 반복 업무 1개를 Power Automate Desktop으로 자동화해보기',
        'plan_b': [
            'Power Automate Desktop 설치 불가 시 → Power Automate 웹 버전(make.powerautomate.com) 사용',
            'Microsoft 계정 없는 경우 → 강사가 체험 계정 제공',
            'IT 정책으로 Power Automate 차단 시 → VBA로 동일 자동화 구현 (2회차 복습 확장)',
        ],
        'it_checklist': [
            'Power Automate Desktop 설치/실행 가능 여부 (Windows 10/11 기본 포함)',
            'Microsoft 계정 로그인 가능 여부',
            'Outlook 데스크톱 앱 사용 가능 여부 (이메일 자동화 실습용)',
        ],
    },

    # ── 4회차 ──────────────────────────────────────────
    {
        'num': 4,
        'title': '대시보드 만들고 팀에 공유 — AI로 HTML 1개만 제대로',
        'subtitle': 'ChatGPT/Copilot 웹에서 HTML 생성 → OneDrive/Teams로 공유',
        'goal': 'AI가 만든 대시보드를 팀원이 브라우저로 볼 수 있게 공유한다',
        'tools': ['Copilot/ChatGPT 웹 (HTML 코드 생성)', 'OneDrive 또는 Teams (파일 공유)'],
        'tools_note': 'python -m http.server 불필요! HTML 파일을 OneDrive/Teams에 올리면 끝',
        'success_30min': '"내가 만든 대시보드를 팀원이 봤다" — AI가 만든 HTML을 OneDrive 링크로 공유',
        'killer_demo': None,
        'timetable': [
            ['0:00~0:15', '도입: 바이브코딩이란?',
             '바이브코딩 = "분위기로 코딩하기"\n'
             '"이런 느낌의 대시보드 만들어줘" → AI가 코드 작성\n'
             '완성된 대시보드 3종 시연 (생산실적, 매출현황, KPI)',
             '강의 + 시연'],
            ['0:15~0:50', '[30분 성공] AI에게 대시보드 만들어달라고 하기',
             'Copilot 웹 열기\n'
             '"월별 매출 데이터로 차트가 있는 대시보드\n'
             ' HTML 페이지 만들어줘. 샘플 데이터 포함해서." 입력\n'
             '→ 생성된 HTML 코드 전체 복사\n'
             '→ 메모장에 붙여넣기 → "dashboard.html"로 저장\n'
             '→ 더블클릭으로 브라우저에서 열기\n'
             '"내가 코드를 하나도 안 썼는데 대시보드가 나왔다!"',
             '미니 실습'],
            ['0:50~1:40', '메인 실습: 우리 데이터로 대시보드 만들기',
             '[시연 10분] 실제 Excel 데이터를 대시보드에 넣는 법\n'
             '[따라하기 40분]\n'
             '(1) 강사 제공 샘플 Excel 데이터 열기\n'
             '(2) 데이터를 복사하여 AI에게 전달:\n'
             '    "이 데이터로 막대차트+원형차트 대시보드 HTML 만들어줘"\n'
             '(3) HTML 저장 → 브라우저에서 확인\n'
             '(4) "색상을 파란색 계열로 바꿔줘" 수정 요청\n'
             '(5) "필터 기능 추가해줘" 추가 요청',
             '시연 + 실습'],
            ['1:40~2:00', '휴식', '', ''],
            ['2:00~2:50', '심화: 팀에 공유하기 (OneDrive/Teams)',
             '[시연 10분] 공유 방법 3가지 시연\n'
             '[따라하기 40분]\n'
             '방법 1: OneDrive에 HTML 업로드 → 공유 링크 생성\n'
             '방법 2: Teams 채널에 HTML 파일 업로드\n'
             '방법 3: Teams 웹훅으로 "새 대시보드 올렸습니다" 알림\n'
             '→ 옆 사람에게 링크 보내서 진짜 열리는지 확인!\n'
             '"Teams 웹훅" 맛보기: 자동 알림 보내기',
             '시연 + 실습'],
            ['2:50~3:40', '응용: 나만의 업무 도구 만들기',
             'Power BI 데모 (10분): "이런 것도 있다" 수준 시연\n'
             '자유 실습 (40분):\n'
             '  AI에게 요청하여 본인 업무 도구 만들기\n'
             '  예시: "자금일보 입력폼 웹페이지 만들어줘"\n'
             '  예시: "채권 연령 계산기 웹페이지 만들어줘"\n'
             '  예시: "재고 현황판 만들어줘"\n'
             '→ 만든 도구를 팀원과 공유하기',
             '시연 + 자유 실습'],
            ['3:40~4:00', '정리 + 월요일 액션 플랜',
             '오늘 배운 것: "AI가 코드를 쓰고, 나는 공유만 한다"\n'
             '월요일 액션: "이번 주 팀 미팅에서 대시보드 보여주기"\n'
             '다음 회차 예고: 인터넷 없이도 AI 쓰기 (로컬 AI)',
             '정리'],
        ],
        'output': [
            'AI가 만든 HTML 대시보드 (차트 포함)',
            'OneDrive/Teams 공유 링크 (팀원 접속 확인 완료)',
            '나만의 업무 도구 HTML 1종',
        ],
        'monday_action': '이번 주 팀 미팅에서 AI로 만든 대시보드를 보여주기',
        'plan_b': [
            'OneDrive 미사용 환경 → 공유 폴더에 HTML 파일 복사 (탐색기 네트워크 드라이브)',
            'Teams 미사용 환경 → 이메일로 HTML 파일 첨부 공유',
            'HTML 저장이 어려운 경우 → AI가 만든 코드를 강사가 USB로 배포',
        ],
        'it_checklist': [
            'OneDrive 또는 Teams 사용 가능 여부',
            'HTML 파일을 PC에 저장 가능 여부',
            '웹 브라우저에서 로컬 HTML 파일 열기 가능 여부',
        ],
    },

    # ── 5회차 ──────────────────────────────────────────
    {
        'num': 5,
        'title': '로컬 AI — 인터넷 없이도 AI가 답한다',
        'subtitle': 'LM Studio GUI 먼저 체험 → Ollama CLI는 선택',
        'goal': '인터넷 연결 없이 AI를 사용할 수 있음을 체험하고, 보안 등급별 AI 선택 기준을 세운다',
        'tools': ['LM Studio (GUI — 클릭만으로 실행)', 'Ollama (CLI — 선택 심화)'],
        'tools_note': 'LM Studio 우선 사용 (설치 파일 강사 USB 배포). Ollama는 원하는 사람만.',
        'success_30min': '"인터넷 끊어도 AI가 답한다" — LM Studio에서 한국어 질문에 답변 받기',
        'killer_demo': None,
        'timetable': [
            ['0:00~0:15', '도입: 왜 로컬 AI가 필요한가?',
             '클라우드 AI의 한계: 보안 등급 3~4 데이터는 못 넣음\n'
             '"이런 데이터도 AI에게 물어보고 싶다면?"\n'
             '→ 내 컴퓨터 안에서만 도는 AI: 로컬 AI\n'
             'LM Studio 시연: 클릭 3번으로 AI 실행',
             '강의 + 시연'],
            ['0:15~0:50', '[30분 성공] LM Studio에서 AI 실행하기',
             'LM Studio 설치 (강사 USB 배포 또는 lmstudio.ai)\n'
             '→ Solar Mini 모델 다운로드 (강사가 미리 준비한 파일)\n'
             '→ "채팅" 탭에서 "안녕하세요" 입력\n'
             '→ AI가 답한다! (인터넷 없이!)\n'
             '→ "자금일보 양식 만들어줘" 요청해보기\n'
             '"이 대화는 인터넷으로 안 나갑니다"',
             '미니 실습'],
            ['0:50~1:40', '메인 실습: Solar vs Gemma4 비교',
             '[시연 10분] 두 모델의 차이점 소개\n'
             '[따라하기 40분]\n'
             '같은 질문을 Solar와 Gemma4에게 각각:\n'
             '(1) 한국어 이메일 작성 → 어디가 더 자연스러운가?\n'
             '(2) VBA 코드 작성 → 어디가 더 정확한가?\n'
             '(3) 데이터 분석 요청 → 어디가 더 유용한가?\n'
             '→ 비교표 작성 (80% 완성 템플릿 제공)',
             '시연 + 실습'],
            ['1:40~2:00', '휴식', '', ''],
            ['2:00~2:50', '심화: Ollama CLI 체험 (선택)',
             '[시연 10분] Ollama 설치 + 명령어 소개\n'
             '[따라하기 40분]\n'
             '원하는 사람: Ollama 설치 → ollama run solar-mini\n'
             '→ 터미널에서 AI와 대화\n'
             '→ "이걸 왜 쓰나요?" — 자동화 스크립트 연동 가능\n'
             '선택하지 않은 사람: LM Studio 심화 (모델 설정 조정, 파라미터 변경)',
             '시연 + 실습'],
            ['2:50~3:40', '응용: 보안 등급별 AI 선택 가이드',
             '종합 정리 (80% 완성 템플릿 제공):\n'
             '  등급 1 (자유) → Copilot 웹 / ChatGPT\n'
             '  등급 2 (클라우드 OK) → Copilot Business / ChatGPT Team\n'
             '  등급 3 (로컬 AI만) → LM Studio / Ollama\n'
             '  등급 4 (AI 금지) → 사람이 직접\n'
             '우리 부서 데이터에 맞는 AI 선택표 완성\n'
             '→ 1회차 보안 지도와 연결하여 완성본 만들기',
             '자유 실습'],
            ['3:40~4:00', '정리 + 월요일 액션 플랜',
             '오늘 배운 것: "보안 등급에 따라 AI를 골라 쓴다"\n'
             '월요일 액션: "LM Studio로 보안 등급 3 데이터에 대해 AI에게 1번 물어보기"\n'
             '다음 회차 예고: 지금까지 배운 모든 것을 연결하기',
             '정리'],
        ],
        'output': [
            'LM Studio 설치 + Solar 모델 실행 완료',
            'Solar vs Gemma4 비교표',
            '보안 등급별 AI 선택 가이드 (우리 부서 버전)',
        ],
        'monday_action': 'LM Studio로 보안 등급 3 데이터(내부 재무 데이터 등)에 대해 AI에게 1번 물어보기',
        'plan_b': [
            'LM Studio 설치 불가 시 (관리자 권한 없음) → 강사 PC에서 시연 + 수강생은 화면 공유로 체험',
            '모델 다운로드 느린 경우 → 강사가 USB로 모델 파일 배포',
            'PC 사양 부족 시 (RAM 8GB 미만) → 소형 모델(Phi-3 mini) 사용 또는 강사 시연',
        ],
        'it_checklist': [
            'LM Studio 설치 가능 여부 (관리자 권한 필요 여부 확인)',
            'PC RAM 8GB 이상 여부 (16GB 권장)',
            'USB 드라이브 연결 가능 여부 (모델 파일 배포용)',
        ],
    },

    # ── 6회차 ──────────────────────────────────────────
    {
        'num': 6,
        'title': '전부 연결하기 — 나의 AI 업무 파이프라인 완성',
        'subtitle': '단계별 가이드 + 페어 학습으로 종합 프로젝트',
        'goal': '1~5회차에서 배운 도구들을 연결하여 본인 업무의 end-to-end 자동화 파이프라인을 완성한다',
        'tools': ['지금까지 배운 모든 도구 (자유 선택)'],
        'tools_note': 'Copilot/ChatGPT 웹 + Excel VBA + Power Automate + OneDrive/Teams + LM Studio',
        'success_30min': '"내 업무가 자동화됐다!" — 페어와 함께 자동화 파이프라인 1단계 완성',
        'killer_demo': None,
        'timetable': [
            ['0:00~0:15', '도입: 5주간의 여정 복습',
             '1회차: AI 체험 + 보안 → "AI를 안전하게 쓸 줄 안다"\n'
             '2회차: VBA 자동화 → "버튼 하나로 문서 완성"\n'
             '3회차: 노코드 자동화 → "마우스만으로 자동화"\n'
             '4회차: 대시보드 → "만들고 공유까지"\n'
             '5회차: 로컬 AI → "보안 데이터도 AI에게"\n'
             '"오늘은 이 모든 걸 연결합니다"',
             '강의'],
            ['0:15~0:50', '[30분 성공] 페어 매칭 + 파이프라인 1단계',
             '2인 1조 페어 매칭 (비슷한 업무끼리)\n'
             '→ 강사 제공 "파이프라인 설계 워크시트" (80% 완성)\n'
             '→ 입력/처리/출력/공유 4단계 빈칸 채우기\n'
             '→ 1단계(입력) 바로 실습:\n'
             '  예: "매일 아침 이메일 첨부파일 → 폴더 자동 저장" 완성\n'
             '"15분 만에 파이프라인 첫 단계가 돌아갑니다!"',
             '미니 실습'],
            ['0:50~1:40', '메인 실습: OneDrive 연동 + 파이프라인 2~3단계',
             '[시연 10분] OneDrive 동기화 + 파이프라인 연결 시연\n'
             '[따라하기 40분]\n'
             'OneDrive PC 동기화 설정 (탐색기에서 바로 접근)\n'
             '파이프라인 2단계(처리): VBA 매크로 또는 Power Automate로 데이터 가공\n'
             '파이프라인 3단계(출력): AI가 만든 보고서/대시보드 자동 생성\n'
             '→ 페어와 서로 코드 리뷰 + 도움',
             '시연 + 실습'],
            ['1:40~2:00', '휴식', '', ''],
            ['2:00~2:50', '심화: 파이프라인 4단계(공유) + 통합',
             '[시연 10분] 완성된 파이프라인 전체 시연\n'
             '[따라하기 40분]\n'
             '4단계(공유): OneDrive/Teams로 결과물 자동 공유\n'
             '→ Teams 웹훅 알림 연동 (선택)\n'
             '→ 파이프라인 전체 흐름 테스트\n'
             '→ 안 되는 부분은 페어/강사와 함께 해결',
             '시연 + 실습'],
            ['2:50~3:40', '발표 준비 + 페어 발표',
             '발표 준비 (15분):\n'
             '  "내가 만든 파이프라인" 1분 발표 준비\n'
             '  구조: 문제 → 해결 → 데모 → 효과\n'
             '페어 발표 (35분):\n'
             '  2인 1조로 발표 (조당 3분)\n'
             '  → 청중 질문 + 강사 피드백\n'
             '  → 베스트 파이프라인 투표',
             '발표 + 토론'],
            ['3:40~4:00', '수료 + AI 시대 생존 가이드',
             '6주 성과 정리: "여러분은 이제 AI를 업무에 쓸 줄 아는 사람입니다"\n'
             'AI 시대 핵심 역량 3가지:\n'
             '  (1) AI에게 잘 시키기 (프롬프트)\n'
             '  (2) AI 결과를 검증하기 (보안 + 정확성)\n'
             '  (3) AI와 사람의 역할 나누기\n'
             '추천 리소스 목록 배포 + 수료',
             '정리 + 수료'],
        ],
        'output': [
            '본인 업무의 end-to-end AI 자동화 파이프라인',
            '파이프라인 설계 워크시트 (완성본)',
            '1분 발표 자료',
        ],
        'monday_action': '완성한 파이프라인을 실제 업무에 1주일간 적용하고 결과 기록하기',
        'plan_b': [
            '파이프라인 전체 완성이 어려운 경우 → 2단계까지만 완성하고 나머지는 설계만',
            'OneDrive 동기화 불가 시 → 공유 폴더 또는 이메일 공유로 대체',
            '발표가 부담스러운 경우 → 워크시트 제출로 대체 가능',
        ],
        'it_checklist': [
            'OneDrive PC 동기화 설정 가능 여부',
            '이전 회차 도구들 (Excel VBA, Power Automate, LM Studio) 정상 작동 여부',
            'Teams 접속 가능 여부 (웹훅 알림 실습용)',
        ],
    },
]


# ============================================================
# GENERATE DOCUMENT
# ============================================================

def generate_document():
    doc = create_doc()

    # ── Cover / Title ──
    title = doc.add_heading('수업계획서 — Microsoft 트랙', level=0)
    for run in title.runs:
        run.font.name = 'Malgun Gothic'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
        run.font.size = Pt(18)

    para(doc, 'AI로 쉬워지는 직장인 업무 자동화', bold=True, size=12)
    para(doc, 'Copilot 웹 · Excel VBA · Power Automate Desktop · OneDrive/Teams · LM Studio',
         size=10, color=(100, 100, 100))

    # ── 과정 정보 ──
    styled_heading(doc, '과정 정보', level=2)
    make_table(doc,
        ['항목', '내용'],
        [
            ['과정명', 'AI로 쉬워지는 직장인 업무 자동화'],
            ['트랙', 'Microsoft 트랙'],
            ['교육 대상', '외부 AI 사용 제한 기업 재직자 (비개발자)'],
            ['총 교육 시간', '24시간 (6회 x 4시간)'],
            ['교육 방식', '이론 20% + 실습 80% (80% 완성 템플릿 기반)'],
            ['핵심 원칙', '30분 안에 성공 경험 / 회차당 도구 1~2개만 / 웹 브라우저 우선'],
            ['주요 AI 도구', 'Copilot 웹, ChatGPT 웹, LM Studio (로컬 AI)'],
            ['자동화 도구', 'Excel VBA, Power Automate Desktop'],
            ['공유 도구', 'OneDrive, Teams (웹훅 알림)'],
            ['로컬 AI', 'LM Studio + Upstage Solar / Gemma 4'],
        ],
        header_color=MS_BLUE
    )

    # ── 수업 설계 원칙 ──
    styled_heading(doc, '수업 설계 원칙', level=2)

    principles = [
        ('30분 성공 경험: ', '매 회차 시작 30분 안에 수강생 화면에 결과물이 나와야 합니다.'),
        ('회차당 도구 1~2개: ', '5개 도구를 훑는 것이 아니라, 1개를 깊이 있게 다룹니다.'),
        ('체험 먼저, 이름은 나중에: ', '실습을 먼저 하고, "방금 한 게 Agent입니다" 식으로 개념을 소개합니다.'),
        ('웹 브라우저 우선: ', '설치가 필요한 도구는 뒤로 미루고, 웹에서 바로 시작합니다.'),
        ('80% 완성 템플릿: ', '빈 칸에서 시작하지 않습니다. 80% 완성된 템플릿을 제공합니다.'),
    ]
    for prefix, text in principles:
        bullet(doc, text, bold_prefix=prefix, size=9)

    # ── 4시간 시간 블록 구조 ──
    styled_heading(doc, '4시간 시간 블록 구조 (매 회차 동일)', level=2)
    make_table(doc,
        ['블록', '시간', '내용', '구성'],
        [
            ['블록1', '0:00~0:50', '도입 + 30분 성공 경험', '강의 15분 + 미니실습 35분'],
            ['블록2', '0:50~1:40', '메인 실습', '시연 10분 + 따라하기 40분'],
            ['', '1:40~2:00', '휴식', '20분'],
            ['블록3', '2:00~2:50', '심화 실습', '시연 10분 + 따라하기 40분'],
            ['블록4', '2:50~3:40', '응용/자유 실습', '50분'],
            ['블록5', '3:40~4:00', '정리 + 월요일 액션 플랜', '20분'],
        ],
        header_color=MS_BLUE
    )

    # ── 6회차 로드맵 ──
    styled_heading(doc, '6회차 로드맵', level=2)
    make_table(doc,
        ['회차', '제목', '핵심 도구', '30분 성공 경험'],
        [
            ['1회차', 'AI 첫 만남 + 보안', 'Copilot/ChatGPT 웹', '"AI가 이메일 초안을 1분에 써줬다"'],
            ['2회차', '자금일보 자동화', 'Excel + VBA', '"버튼 하나로 자금일보 완성"'],
            ['3회차', 'AI가 혼자 판단하게', 'Power Automate Desktop', '"이메일 오면 자동 폴더 저장"'],
            ['4회차', '대시보드 + 팀 공유', 'AI HTML 생성 + OneDrive', '"내 대시보드를 팀원이 봤다"'],
            ['5회차', '로컬 AI (오프라인)', 'LM Studio + Ollama', '"인터넷 끊어도 AI가 답한다"'],
            ['6회차', '전부 연결하기', '전체 도구 통합', '"내 업무가 자동화됐다!"'],
        ],
        header_color=MS_BLUE
    )

    doc.add_page_break()

    # ── 각 회차 상세 ──
    for session in sessions:
        num = session['num']

        # Page break between sessions (except first)
        if num > 1:
            doc.add_page_break()

        # Session header
        styled_heading(doc, f'{num}회차: {session["title"]}', level=1)
        para(doc, session['subtitle'], size=10, color=(100, 100, 100), italic=True)

        # 학습 목표
        para(doc, f'학습 목표: {session["goal"]}', bold=True, size=10)

        # 핵심 도구
        para(doc, '핵심 도구:', bold=True, size=9, color=(43, 87, 154))
        for tool in session['tools']:
            bullet(doc, tool, size=9)
        if session.get('tools_note'):
            para(doc, f'  * {session["tools_note"]}', size=8, italic=True, color=(150, 100, 0))

        # 30분 성공 경험
        para(doc, '', size=4)
        para(doc, f'30분 성공 경험', bold=True, size=10, color=(16, 124, 16))
        para(doc, session['success_30min'], size=9)

        # 킬러 데모 (1회차만)
        if session.get('killer_demo'):
            para(doc, f'킬러 데모:', bold=True, size=9, color=(216, 59, 1))
            para(doc, session['killer_demo'], size=9)

        # 시간표
        para(doc, '', size=4)
        styled_heading(doc, '상세 시간표', level=3)
        make_table(doc,
            ['시간', '주제', '상세 내용', '형태'],
            session['timetable'],
            header_color=MS_BLUE
        )

        # 산출물
        para(doc, '', size=4)
        para(doc, '산출물:', bold=True, size=10, color=(0, 100, 0))
        for item in session['output']:
            bullet(doc, item, size=9)

        # 월요일 액션 플랜
        para(doc, '', size=4)
        para(doc, '월요일 액션 플랜:', bold=True, size=10, color=(43, 87, 154))
        para(doc, session['monday_action'], size=9)

        # 플랜B
        para(doc, '', size=4)
        para(doc, '플랜B (대비 시나리오):', bold=True, size=9, color=(216, 59, 1))
        for item in session['plan_b']:
            bullet(doc, item, size=8)

        # IT 사전 체크리스트
        para(doc, '', size=4)
        para(doc, 'IT 사전 체크리스트:', bold=True, size=9, color=(100, 100, 100))
        for item in session['it_checklist']:
            bullet(doc, f'[ ] {item}', size=8)

    # ── 부록: 전체 IT 사전 협조 체크리스트 ──
    doc.add_page_break()
    styled_heading(doc, '부록: IT 사전 협조 체크리스트 (종합)', level=1)
    para(doc, '교육 시작 최소 1주 전에 IT 부서와 확인이 필요한 항목입니다.', size=9, italic=True)

    all_checks = [
        ('네트워크', [
            'copilot.microsoft.com 접속 허용',
            'chat.openai.com 접속 허용',
            'lmstudio.ai 접속 허용 (또는 USB 설치 대비)',
        ]),
        ('소프트웨어', [
            'Excel 매크로(VBA) 실행 허용 (보안 설정)',
            'Power Automate Desktop 실행 허용 (Windows 기본 포함)',
            'LM Studio 설치 허용 (또는 강사 시연 모드)',
        ]),
        ('계정', [
            'Microsoft 계정 로그인 가능 (Power Automate용)',
            'OneDrive 동기화 가능',
            'Teams 접속 가능',
        ]),
        ('하드웨어', [
            'PC RAM 8GB 이상 (로컬 AI용, 16GB 권장)',
            'USB 드라이브 연결 가능 (모델 파일 배포용)',
            '듀얼 모니터 또는 화면 분할 가능 (권장)',
        ]),
    ]

    for category, checks in all_checks:
        para(doc, category, bold=True, size=10)
        for check in checks:
            bullet(doc, f'[ ] {check}', size=9)

    # ── 부록: 인지 부하 체크 ──
    para(doc, '', size=6)
    styled_heading(doc, '부록: 회차별 인지 부하 체크', level=2)
    make_table(doc,
        ['회차', '새 도구', '새 개념', '인지 부하 수준'],
        [
            ['1회차', 'Copilot 웹, ChatGPT 웹', '데이터 4등급 분류', '낮음 (웹만 사용)'],
            ['2회차', 'Excel VBA', 'VBA 매크로, AI 코드 검증', '중간 (코드 복붙)'],
            ['3회차', 'Power Automate Desktop', 'Agent, MCP (체험 후 소개)', '중간 (노코드)'],
            ['4회차', '없음 (기존 도구 조합)', '바이브코딩, HTML', '낮음~중간'],
            ['5회차', 'LM Studio, Ollama(선택)', '로컬 AI, 모델 비교', '중간~높음'],
            ['6회차', '없음 (전체 통합)', '파이프라인 설계', '높음 (종합)'],
        ],
        header_color=MS_BLUE
    )

    para(doc, '* 새 도구: 회차당 최대 2개 / 새 개념: 회차당 최대 2개 / 새 용어: "쉬운 말" 병기 필수',
         size=8, italic=True, color=(100, 100, 100))

    # Save
    output_path = r'C:\Users\moonk\OneDrive\AI강사준비\수업계획서_Microsoft_트랙.docx'
    doc.save(output_path)
    print(f'Saved: {output_path}')
    return output_path


if __name__ == '__main__':
    generate_document()
