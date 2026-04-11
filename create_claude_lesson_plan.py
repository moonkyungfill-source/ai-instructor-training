"""Claude 트랙 수업계획서 생성 스크립트 (완전 개편)

개선사항 반영:
- Artifacts = 킬러 콘텐츠 (대화 중 웹앱/차트 즉시 생성)
- 200K 컨텍스트 (100페이지 문서 통분석)
- MCP (Claude Desktop에서 파일시스템/Slack/DB 직접 연결)
- Projects (부서별 AI 어시스턴트 구축)
- Pro Plan 호환 (Teams Plan 필수가 아님)
- 30분 성공 경험 / 4시간 블록 구조
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# DOCUMENT HELPERS
# ============================================================

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


def para(doc, text, bold=False, size=9, color=None, italic=False):
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
    return p


def make_table(doc, headers, rows, header_color='D97706'):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
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
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(8)
            run.font.name = 'Malgun Gothic'
            run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    return table


def add_bullet_list(doc, items, size=9, indent_cm=1.0):
    """불릿 리스트 추가"""
    for item in items:
        p = doc.add_paragraph(style='List Bullet')
        p.clear()
        run = p.add_run(item)
        run.font.name = 'Malgun Gothic'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
        run.font.size = Pt(size)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(1)


# ============================================================
# SESSION DATA — Claude 트랙 (완전 개편)
# ============================================================

sessions = [
    # ── 1회차 ──────────────────────────────────────
    {
        'num': 1,
        'title': 'AI 첫 만남 + 보안 — claude.ai 웹에서 바로 체험',
        'goal': 'AI 보안 리터러시 확보 + claude.ai 200K 컨텍스트로 "AI가 진짜 되네!" 체감',
        'success_30min': '"AI가 이 계약서 20페이지를 1분에 요약했다" — 200K 컨텍스트 체감',
        'tools': 'claude.ai 웹 (설치 없음, Pro Plan OK)',
        'new_concepts': 'AI 보안 4등급 / 프롬프트 기초',
        'timetable': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '(강의 15분) 왜 ChatGPT가 막혔을까?\n'
             '삼성 소스코드 유출 사건 → 기업 AI 정책 현실\n'
             'Claude 소개: 환각 최소화, 200K 컨텍스트\n\n'
             '(미니실습 35분) ★30분 성공 경험★\n'
             '준비된 20페이지 계약서 PDF 업로드\n'
             '→ "핵심 조항 5개 요약해줘" 입력\n'
             '→ 1분 안에 요약 결과 확인\n'
             '→ "3조 위약금 조항을 쉬운 말로 설명해줘"\n'
             '→ "을에게 불리한 조항을 찾아줘"',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습:\n데이터 보안 분류',
             '(시연 10분) 데이터 4등급 분류 기준 설명\n'
             '① 자유 사용 ② 클라우드 OK ③ 로컬만 ④ 절대불가\n'
             'Claude 보안 정책: Pro Plan 학습 차단 옵션\n\n'
             '(따라하기 40분) 80% 완성 템플릿 배포\n'
             '→ 우리 회사 데이터 목록 작성\n'
             '→ 등급별 분류 + AI 도구 매핑\n'
             '→ Claude에게 "이 분류표 빠진 거 없어?" 검증',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화 실습:\nClaude 200K 긴 문서 분석',
             '(시연 10분) 100페이지 PDF 업로드 시연\n'
             '→ 요약 / 특정 조항 검색 / 비교 분석\n\n'
             '(따라하기 40분) 준비된 감사보고서(50+페이지) 업로드\n'
             '→ "경영진에게 보고할 핵심 이슈 3개 추출"\n'
             '→ "전년 대비 변동이 큰 항목 찾아줘"\n'
             '→ "감사의견에서 주의할 점을 표로 정리해줘"\n'
             '* Claude 200K = 경쟁사(GPT 128K, Gemini 1M이지만 정확도) 대비 강점',
             '시연 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용/자유 실습',
             '본인 업무 문서(계약서/보고서/매뉴얼 등) 업로드\n'
             '→ Claude에게 분석·요약·비교 요청\n'
             '→ "이 보고서에서 수치 오류 있어?" 등 검증\n'
             '→ 결과를 팀원과 공유하며 비교\n\n'
             '※ 민감 문서 사용 시 데이터 등급 확인 후 진행',
             '자유 실습\n50분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 월요일 액션 플랜',
             '오늘 배운 것 정리:\n'
             '① claude.ai 접속 + 프롬프트 기초\n'
             '② 200K 긴 문서 분석\n'
             '③ 데이터 보안 4등급 분류\n\n'
             '★ 월요일 액션 플랜 ★\n'
             '→ "내 업무에서 AI에 맡길 문서 3개 선정"\n'
             '→ "데이터 등급 확인 후 claude.ai에 업로드"',
             '정리 20분'],
        ],
        'output': '데이터 분류 기준표 완성 + claude.ai 긴 문서 분석 체험',
        'plan_b': '인터넷 불안정 시: 강사 화면 미러링 + 스크린샷 자료 배포',
        'monday_action': '내 업무에서 AI에 맡길 문서 3개 선정 → 데이터 등급 확인 → claude.ai에 업로드해 분석',
    },

    # ── 2회차 ──────────────────────────────────────
    {
        'num': 2,
        'title': '자금일보 자동화 — Artifacts로 즉시 결과',
        'goal': 'Claude Artifacts로 코드/수식/문서를 즉시 생성+실행 체험',
        'success_30min': '"Claude가 만든 수식으로 자금일보가 완성됐다" — Artifacts 즉시 생성 체감',
        'tools': 'claude.ai Artifacts (코드/수식 즉시 생성+실행)',
        'new_concepts': 'Artifacts(즉시 실행 가능한 결과물) / 프롬프트 엔지니어링',
        'timetable': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '(강의 15분) 프롬프트 엔지니어링 핵심 3원칙\n'
             '① 역할 부여 ② 구체적 지시 ③ 예시 제공\n'
             'Artifacts란? → 대화 중 바로 실행되는 결과물\n\n'
             '(미니실습 35분) ★30분 성공 경험★\n'
             '80% 완성 자금일보 양식(엑셀) 배포\n'
             '→ Claude에게 "이 양식에 맞는 엑셀 수식 만들어줘"\n'
             '→ Artifacts에서 수식 코드 즉시 확인\n'
             '→ 엑셀에 붙여넣기 → 자금일보 완성\n'
             '→ "입출금 합계가 안 맞을 때 빨간색 표시" 추가 요청',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습:\n채권연령표 자동화',
             '(시연 10분) Claude에게 복잡한 엑셀 수식 요청 시연\n'
             '→ 경과일수 계산 + 구간 분류 수식 생성\n\n'
             '(따라하기 40분) 80% 완성 채권연령표 템플릿 배포\n'
             '→ "매출채권 경과일수 계산해서 30/60/90/120일 구간 분류"\n'
             '→ "연체 금액 자동 합계 + 비율 계산"\n'
             '→ "부서별 채권현황 요약 피벗 만들어줘"\n'
             '→ 결과를 엑셀에 적용',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화 실습:\n100페이지 감사보고서 분석',
             '(시연 10분) 긴 문서 업로드 + Artifacts 활용 시연\n'
             '→ 보고서 핵심 추출 → 표/차트로 즉시 시각화\n\n'
             '(따라하기 40분) 준비된 감사보고서(100p) 업로드\n'
             '→ "재무상태 변동 요약을 표로 만들어줘"\n'
             '→ "전기 대비 증감이 큰 계정과목 Top 10"\n'
             '→ "감사의견 핵심을 경영진 보고용 1페이지로"\n'
             '→ Artifacts로 보고서 초안 즉시 생성',
             '시연 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용/자유 실습:\nDocument AI 맛보기',
             '본인 업무 엑셀에 Claude 수식 적용 자유 실습\n\n'
             'Document AI 맛보기 (선택):\n'
             '→ 업스테이지 Document Parse API 소개\n'
             '→ PDF 영수증/계약서 → 텍스트 자동 추출\n'
             '→ Claude에 붙여넣기 → 정리된 표 완성\n\n'
             '※ API 키 없으면 강사 시연으로 대체',
             '자유 실습\n50분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 월요일 액션 플랜',
             '오늘 배운 것 정리:\n'
             '① Artifacts로 수식/코드 즉시 생성\n'
             '② 긴 문서 분석 → 보고서 초안 자동 생성\n'
             '③ 프롬프트 3원칙 (역할·구체·예시)\n\n'
             '★ 월요일 액션 플랜 ★\n'
             '→ "내 반복 엑셀 작업 1개를 Claude로 자동화"\n'
             '→ "만든 수식을 팀원에게 공유"',
             '정리 20분'],
        ],
        'output': '자금일보 수식 완성 + 채권연령표 자동화 + 감사보고서 요약',
        'plan_b': 'Artifacts 미지원 시: Claude 텍스트 출력 → 수동 엑셀 붙여넣기',
        'monday_action': '내 반복 엑셀 작업 1개를 Claude Artifacts로 자동화 → 팀원에게 공유',
    },

    # ── 3회차 ──────────────────────────────────────
    {
        'num': 3,
        'title': 'AI가 혼자 판단하게 만들기 — Claude Desktop + MCP',
        'goal': 'Claude Desktop + MCP로 파일시스템 직접 연결 → Agent 개념 체험 먼저, 이름은 나중에',
        'success_30min': '"Claude가 내 폴더를 직접 읽고 정리했다" — MCP 파일시스템 연결 체감',
        'tools': 'Claude Desktop + MCP (파일시스템 연결)',
        'new_concepts': 'MCP(모델이 도구를 직접 사용하는 프로토콜) / Agent(스스로 판단하는 AI)',
        'timetable': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '(강의 15분) Claude Desktop 소개\n'
             '"웹 버전과 뭐가 달라?" → 내 PC와 직접 연결\n'
             'MCP란 이름은 아직! 일단 해보자\n\n'
             '(미니실습 35분) ★30분 성공 경험★\n'
             'Claude Desktop 설치 (사전 배포 완료)\n'
             '→ 파일시스템 MCP 설정 (config 파일 복붙)\n'
             '→ "내 문서 폴더에 뭐가 있어?" 입력\n'
             '→ Claude가 폴더 목록을 직접 읽어서 보여줌\n'
             '→ "txt 파일만 골라서 내용 요약해줘"\n'
             '→ "파일 이름을 날짜_제목 형식으로 바꿔줘"',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습:\n파일 자동 정리 Agent',
             '(시연 10분) "방금 한 게 Agent입니다"\n'
             'Agent = AI가 스스로 판단해서 도구를 사용하는 것\n'
             'MCP = Model Context Protocol (AI가 도구를 쓰는 규격)\n\n'
             '(따라하기 40분) 연습용 폴더(30개 파일) 배포\n'
             '→ "이 폴더를 종류별로 분류해줘"\n'
             '→ "계약서는 /계약서 폴더로, 견적서는 /견적서로"\n'
             '→ "각 폴더에 요약 README 만들어줘"\n'
             '→ Claude가 직접 파일 이동+생성하는 모습 확인',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화 실습:\nMCP 확장 연결',
             '(시연 10분) MCP 생태계 소개\n'
             '파일시스템 외: Slack / DB / 웹 검색 등\n'
             '기업에서 쓸 수 있는 MCP 서버 목록\n\n'
             '(따라하기 40분) 추가 MCP 실습\n'
             '→ 파일 읽기 → Claude가 분석 → 결과 파일로 저장\n'
             '→ "이 CSV 파일 읽어서 월별 매출 요약 만들어줘"\n'
             '→ "결과를 summary.md 파일로 저장해줘"\n'
             '→ 입력→분석→출력 전 과정을 AI가 직접 수행',
             '시연 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용/자유 실습',
             '본인 업무 폴더로 MCP Agent 실습\n'
             '→ "이번 달 보고서 파일 모아서 요약해줘"\n'
             '→ "거래처별로 파일 분류해줘"\n'
             '→ 본인 업무에 맞는 Agent 시나리오 설계\n\n'
             '※ 보안 주의: 연습용 폴더만 연결\n'
             '실제 업무 폴더 연결은 데이터 등급 확인 후',
             '자유 실습\n50분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 월요일 액션 플랜',
             '오늘 배운 것 정리:\n'
             '① Claude Desktop + MCP = AI가 내 PC와 직접 연결\n'
             '② Agent = AI가 스스로 판단해서 도구 사용\n'
             '③ 파일 읽기/이동/생성을 AI가 직접 수행\n\n'
             '★ 월요일 액션 플랜 ★\n'
             '→ "내 PC에서 매주 반복하는 파일 정리 1개를\n'
             '   Claude Desktop Agent에게 맡기기"',
             '정리 20분'],
        ],
        'output': 'Claude Desktop + MCP 파일시스템 Agent 구성 완료',
        'plan_b': 'Claude Desktop 설치 불가(관리자 권한 없음) 시:\nclaude.ai 웹에서 파일 업로드 방식으로 동일 실습\n→ 수동 업로드이지만 분석·요약·정리 결과는 동일',
        'monday_action': 'Claude Desktop에서 매주 반복하는 파일 정리 작업 1개를 Agent에게 맡겨보기',
    },

    # ── 4회차 ──────────────────────────────────────
    {
        'num': 4,
        'title': '대시보드 만들고 팀에 공유 — Artifacts 웹앱 즉시 생성',
        'goal': 'Claude Artifacts로 대화만으로 웹앱/대시보드 즉시 제작 + 팀 공유',
        'success_30min': '"대화로 만든 대시보드가 바로 동작한다" — Artifacts HTML/React 실행 체감',
        'tools': 'Claude Artifacts (HTML/React 즉시 실행)',
        'new_concepts': '바이브코딩(대화로 코딩) / Artifacts 웹앱',
        'timetable': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '(강의 15분) 바이브코딩이란?\n'
             '"코드를 모르는 사람이 대화로 앱을 만드는 것"\n'
             'Artifacts = Claude만의 차별점\n'
             '(Copilot, Gemini에는 없는 기능)\n\n'
             '(미니실습 35분) ★30분 성공 경험★\n'
             '→ Claude에게 "생산 실적 데이터로 차트 대시보드 만들어줘" 입력\n'
             '→ Artifacts에서 즉시 실행되는 대시보드 확인\n'
             '→ "막대 차트를 원형 차트로 바꿔줘"\n'
             '→ "날짜 필터 추가해줘"\n'
             '→ 대화만으로 대시보드 완성',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습:\n업무 도구 웹앱 제작',
             '(시연 10분) Artifacts로 만들 수 있는 것들:\n'
             '→ 계산기, 입력폼, 데이터 뷰어, 게임까지\n\n'
             '(따라하기 40분) 업무 도구 3개 중 택 1 제작:\n'
             '① "자금일보 입력 → 자동 계산 → PDF 출력 웹앱"\n'
             '② "채권 연령 자동 계산기"\n'
             '③ "재고 현황 대시보드"\n'
             '→ 80% 완성 프롬프트 템플릿 배포\n'
             '→ 빈칸만 채워서 Claude에게 요청\n'
             '→ Artifacts에서 즉시 결과 확인 + 수정 반복',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화 실습:\nKPI 대시보드 고도화',
             '(시연 10분) 실제 회사 데이터로 대시보드 고도화 시연\n'
             '→ 여러 차트 + 필터 + 인쇄 기능\n\n'
             '(따라하기 40분) 본인 업무 데이터로 대시보드 제작\n'
             '→ "부서별 매출 현황 대시보드"\n'
             '→ "월별 추이 라인 차트 + 목표 대비 달성률"\n'
             '→ "인쇄 버튼 추가해줘"\n'
             '→ "회사 로고 + 부서명 헤더 넣어줘"',
             '시연 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용/자유 실습:\n공유하기',
             'Artifacts 결과물 공유 방법:\n'
             '① Artifacts 공유 링크 생성 → 팀원에게 전달\n'
             '② HTML 다운로드 → 사내 파일서버에 업로드\n'
             '③ python -m http.server → 같은 네트워크에서 접속\n\n'
             '자유 제작:\n'
             '→ 본인이 원하는 업무 도구 자유 제작\n'
             '→ 팀원과 결과물 교환 + 상호 피드백\n'
             '→ "공장 태블릿용 생산현황판" 등',
             '자유 실습\n50분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 월요일 액션 플랜',
             '오늘 배운 것 정리:\n'
             '① Artifacts = 대화로 앱 즉시 생성 (Claude만의 차별점)\n'
             '② 바이브코딩으로 대시보드/도구 제작\n'
             '③ 공유: 링크 / HTML 다운로드 / 로컬서버\n\n'
             '★ 월요일 액션 플랜 ★\n'
             '→ "팀에서 엑셀로 보던 현황판 1개를\n'
             '   Artifacts 대시보드로 만들어서 공유"',
             '정리 20분'],
        ],
        'output': '업무용 대시보드 + 웹앱 제작 완료 + 팀 공유',
        'plan_b': 'Artifacts 실행 불가 시:\nClaude가 생성한 HTML 코드를 복사 → 메모장에 저장(.html) → 브라우저로 열기',
        'monday_action': '팀에서 엑셀로 보던 현황판 1개를 Artifacts 대시보드로 만들어 팀원에게 공유',
    },

    # ── 5회차 ──────────────────────────────────────
    {
        'num': 5,
        'title': '로컬 AI — 인터넷 없이도 AI가 답한다',
        'goal': 'LM Studio GUI로 로컬 AI 첫 체험 + 하이브리드(로컬+클라우드) 전략 수립',
        'success_30min': '"인터넷 끊어도 AI가 답한다" — LM Studio 로컬 AI 실행 체감',
        'tools': 'LM Studio GUI → Ollama (선택)',
        'new_concepts': '로컬 AI / 하이브리드 AI 전략',
        'timetable': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '(강의 15분) 로컬 AI vs 클라우드 AI\n'
             '→ 왜 로컬 AI가 필요한가? (보안, 비용, 오프라인)\n'
             '→ LM Studio: 클릭만으로 AI 실행\n\n'
             '(미니실습 35분) ★30분 성공 경험★\n'
             'LM Studio 설치 (사전 배포 완료)\n'
             '→ Solar Mini 모델 다운로드 (사전 다운로드 완료)\n'
             '→ "안녕하세요. 저는 로컬 AI입니다" 대화\n'
             '→ 와이파이 끄기 → 동일 질문 → "인터넷 없이도 된다!"\n'
             '→ "이 계약서 요약해줘" (파일 드래그 앤 드롭)',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습:\nSolar vs Gemma4 비교',
             '(시연 10분) 모델 비교 기준 설명\n'
             '→ 한국어 성능 / 코드 생성 / 문서 처리 / 속도\n\n'
             '(따라하기 40분) 동일 프롬프트로 모델 비교\n'
             '→ Solar Mini: 한국어 특화 (업스테이지)\n'
             '→ Gemma4: Google 경량 모델\n'
             '→ 비교 항목: 한국어 자연스러움, 업무 문서 이해도\n'
             '→ "어떤 모델이 내 업무에 맞을까?" 판단\n'
             '→ 비교표 직접 작성',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화 실습:\nOllama CLI (선택) +\n하이브리드 전략',
             '(시연 10분) Ollama 소개 (명령줄 기반 로컬 AI)\n'
             '→ LM Studio GUI vs Ollama CLI 장단점\n'
             '→ 하이브리드 전략: 민감→로컬, 복잡→Claude\n\n'
             '(따라하기 40분)\n'
             'A코스(Ollama 도전): 설치 → Solar 실행 → 대화\n'
             'B코스(LM Studio 심화): 파일 일괄 처리 실습\n'
             '→ 폴더 안 파일 10개를 로컬 AI로 요약\n'
             '→ "민감 데이터는 로컬 AI, 복잡 분석은 Claude"\n'
             '→ 하이브리드 판단 기준표 작성',
             '시연 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용/자유 실습',
             '보안 등급별 AI 선택 가이드 작성:\n'
             '→ 1등급(자유): Claude Pro 사용\n'
             '→ 2등급(클라우드OK): Claude Pro (학습차단 ON)\n'
             '→ 3등급(로컬만): LM Studio / Ollama\n'
             '→ 4등급(절대불가): AI 사용 금지\n\n'
             '본인 업무별 AI 도구 매핑표 완성\n'
             '→ "이 업무는 Claude, 저 업무는 로컬 AI"',
             '자유 실습\n50분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 월요일 액션 플랜',
             '오늘 배운 것 정리:\n'
             '① LM Studio로 로컬 AI 실행 (인터넷 불필요)\n'
             '② Solar vs Gemma4 한국어 성능 비교\n'
             '③ 하이브리드 전략: 보안 등급별 AI 선택\n\n'
             '★ 월요일 액션 플랜 ★\n'
             '→ "3등급 데이터 업무 1개를 로컬 AI로 처리해보기"\n'
             '→ "Claude와 비교해서 어떤 게 나은지 판단"',
             '정리 20분'],
        ],
        'output': 'LM Studio 로컬 AI 실행 + Solar/Gemma4 비교표 + 하이브리드 전략표',
        'plan_b': 'PC 사양 부족으로 로컬 AI 설치 불가 시:\n강사 PC로 시연 + 수강생은 claude.ai에서 동일 프롬프트 실습',
        'monday_action': '3등급(로컬만) 데이터 업무 1개를 LM Studio 로컬 AI로 처리 → Claude 결과와 비교',
    },

    # ── 6회차 ──────────────────────────────────────
    {
        'num': 6,
        'title': '전부 연결하기 + 부서 AI 어시스턴트 — Claude Projects',
        'goal': 'Claude Projects에 부서 문서 업로드 → 교육 후에도 계속 쓰는 맞춤 AI 완성',
        'success_30min': '"우리 부서 전용 AI가 완성됐다" — Projects에 문서 업로드 즉시 체감',
        'tools': 'Claude Projects (부서 지식 베이스)',
        'new_concepts': 'Projects(맞춤 AI 지식 베이스) / 파이프라인 설계',
        'timetable': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '(강의 15분) Claude Projects 소개\n'
             '→ "우리 부서 규정/양식/매뉴얼을 AI에게 가르치자"\n'
             '→ Projects = 부서 전용 AI 어시스턴트\n'
             '→ 교육 끝나고도 계속 사용 가능!\n\n'
             '(미니실습 35분) ★30분 성공 경험★\n'
             'Claude Projects 생성\n'
             '→ "재무팀 AI 어시스턴트" 프로젝트 생성\n'
             '→ 부서 규정 PDF 3개 업로드\n'
             '→ "출장비 정산 기준이 뭐야?" 질문\n'
             '→ 업로드한 규정에서 정확히 답변!\n'
             '→ "이번 달 마감 체크리스트 만들어줘"',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습:\n부서 AI 어시스턴트 구축',
             '(시연 10분) 효과적인 Projects 구성법\n'
             '→ 시스템 프롬프트 작성 요령\n'
             '→ 문서 업로드 전략 (규정 > 양식 > 사례)\n\n'
             '(따라하기 40분) 본인 부서 Projects 구축\n'
             '→ 부서명 + 역할 정의 시스템 프롬프트 작성\n'
             '→ 규정/양식/FAQ 문서 업로드 (최소 5개)\n'
             '→ 테스트 질문 5개로 답변 품질 확인\n'
             '→ 시스템 프롬프트 개선 반복\n'
             '※ 80% 완성 시스템 프롬프트 템플릿 배포',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화 실습:\n6회차 총정리 파이프라인',
             '(시연 10분) 1~6회차 도구 총정리 시연\n'
             '→ claude.ai(분석) + Artifacts(시각화) +\n'
             '   Desktop+MCP(Agent) + Projects(지식)\n\n'
             '(따라하기 40분) 종합 파이프라인 설계\n'
             '→ 본인 업무에서 가장 반복적인 작업 1개 선정\n'
             '→ "데이터 입력 → AI 분석 → 보고서 생성 → 공유"\n'
             '→ 각 단계에 어떤 Claude 기능을 쓸지 설계\n'
             '→ 실제로 1회 실행해보기',
             '시연 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용/자유 실습:\n종합 프로젝트',
             '본인 업무 파이프라인 완성:\n'
             '→ Projects에 부서 문서 추가 업로드\n'
             '→ Artifacts로 결과물 시각화\n'
             '→ 파이프라인 문서화 (다른 팀원도 따라할 수 있게)\n\n'
             '동료 피드백:\n'
             '→ 2인 1조로 상대방 Projects 테스트\n'
             '→ "이런 질문도 답해줄까?" 경계 테스트',
             '자유 실습\n50분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 수료',
             '6회차 전체 정리:\n'
             '① claude.ai — 긴 문서 분석 (200K 컨텍스트)\n'
             '② Artifacts — 수식/코드/대시보드 즉시 생성\n'
             '③ Desktop+MCP — AI가 내 PC를 직접 제어 (Agent)\n'
             '④ 로컬 AI — 인터넷 없이 민감 데이터 처리\n'
             '⑤ Projects — 부서 전용 AI 어시스턴트\n\n'
             '★ 교육 후 액션 플랜 ★\n'
             '→ "부서 Projects를 팀원에게 공유"\n'
             '→ "매주 1개씩 반복 업무를 AI로 전환"\n\n'
             '수료식 + Q&A',
             '정리 20분'],
        ],
        'output': '부서 전용 Claude Projects AI 어시스턴트 완성 + 업무 파이프라인 문서',
        'plan_b': 'Projects 미지원(Free Plan) 시:\nclaude.ai 대화에서 매번 문서 업로드 방식으로 유사 경험 제공\n→ "Projects는 이걸 저장해두는 것" 설명',
        'monday_action': '부서 Projects를 팀원에게 공유 + 매주 1개씩 반복 업무를 Claude로 전환',
    },
]


# ============================================================
# GENERATE DOCUMENT
# ============================================================

def generate_claude_lesson_plan():
    doc = create_doc()
    color = 'D97706'

    # ── 표지 ──
    title = doc.add_heading('수업계획서 — Claude 트랙', level=0)
    for run in title.runs:
        run.font.name = 'Malgun Gothic'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
        run.font.size = Pt(18)

    para(doc, 'AI로 쉬워지는 직장인 업무 자동화', bold=True, size=12)
    para(doc, 'claude.ai · Artifacts · Claude Desktop + MCP · Projects · LM Studio',
         size=10, color=(100, 100, 100))

    # ── 과정 정보 ──
    styled_heading(doc, '과정 정보', level=2)
    make_table(doc,
        ['항목', '내용'],
        [
            ['과정명', 'AI로 쉬워지는 직장인 업무 자동화'],
            ['트랙', 'Claude 트랙 (claude.ai · Artifacts · Desktop + MCP · Projects)'],
            ['교육 대상', '외부 AI 사용 제한 기업 재직자 (비개발자)'],
            ['총 교육 시간', '24시간 (6회 x 4시간)'],
            ['교육 방식', '이론 20% + 실습 80%'],
            ['플랜 호환', 'Claude Pro Plan 이상 (Teams Plan 필수 아님)'],
            ['핵심 도구', 'claude.ai 웹, Artifacts, Claude Desktop, MCP, Projects'],
            ['로컬 AI', 'LM Studio + Ollama (5회차) / Upstage Solar, Gemma4'],
            ['Claude 차별점', '200K 컨텍스트 · Artifacts 즉시실행 · MCP Agent · Projects 지식베이스'],
        ],
        header_color=color
    )

    # ── 수업 설계 원칙 ──
    styled_heading(doc, '수업 설계 원칙', level=2)
    principles = [
        '30분 성공 경험: 매 회차 시작 30분 안에 수강생 화면에 결과물',
        '회차당 핵심 도구 1~2개: 도구 나열 아닌 하나에 집중',
        '체험 먼저, 이름은 나중에: 실습 후 → "방금 한 게 Agent입니다"',
        '웹 브라우저 우선: 1~2회차는 설치 없이 claude.ai만 사용',
        '80% 완성 템플릿: 빈 칸 시작 금지, 거의 완성된 템플릿에서 수정',
    ]
    add_bullet_list(doc, principles)

    # ── 4시간 시간 블록 구조 ──
    styled_heading(doc, '4시간 시간 블록 구조 (매 회차 동일)', level=2)
    make_table(doc,
        ['블록', '시간', '구성', '비중'],
        [
            ['블록1', '0:00~0:50', '도입 + 30분 성공 경험 (강의 15분 + 미니실습 35분)', '50분'],
            ['블록2', '0:50~1:40', '메인 실습 (시연 10분 + 따라하기 40분)', '50분'],
            ['휴식', '1:40~2:00', '휴식', '20분'],
            ['블록3', '2:00~2:50', '심화 실습 (시연 10분 + 따라하기 40분)', '50분'],
            ['블록4', '2:50~3:40', '응용/자유 실습', '50분'],
            ['블록5', '3:40~4:00', '정리 + 월요일 액션 플랜', '20분'],
        ],
        header_color=color
    )

    # ── 6회차 커리큘럼 요약 ──
    styled_heading(doc, '6회차 커리큘럼 요약', level=2)
    make_table(doc,
        ['회차', '제목', '30분 성공 경험', '핵심 도구', '산출물'],
        [
            ['1', 'AI 첫 만남 + 보안', '20페이지 계약서 1분 요약', 'claude.ai 웹', '데이터 분류표 + 문서 분석 체험'],
            ['2', '자금일보 자동화', 'Claude가 만든 수식으로 자금일보 완성', 'Artifacts', '자금일보 + 채권연령표 + 감사보고서 요약'],
            ['3', 'AI가 혼자 판단하게', 'Claude가 내 폴더를 직접 읽고 정리', 'Desktop + MCP', 'MCP Agent 파일 정리 체험'],
            ['4', '대시보드 + 팀 공유', '대화로 만든 대시보드가 바로 동작', 'Artifacts 웹앱', '업무 대시보드 + 팀 공유'],
            ['5', '로컬 AI', '인터넷 끊어도 AI가 답한다', 'LM Studio', '로컬 AI + 하이브리드 전략표'],
            ['6', '전부 연결하기', '우리 부서 전용 AI 완성', 'Projects', '부서 AI 어시스턴트 + 파이프라인'],
        ],
        header_color=color
    )

    doc.add_page_break()

    # ── 각 회차 상세 ──
    for session in sessions:
        num = session['num']

        styled_heading(doc, f'{num}회차: {session["title"]}', level=1)

        # 메타 정보
        para(doc, f'학습 목표: {session["goal"]}', bold=True, size=9)
        para(doc, f'30분 성공 경험: {session["success_30min"]}',
             bold=True, size=9, color=(0, 100, 180))
        para(doc, f'핵심 도구: {session["tools"]}', size=9)
        para(doc, f'새 개념: {session["new_concepts"]}', size=9)

        # 시간표
        styled_heading(doc, '상세 시간표', level=2)
        make_table(doc,
            ['시간', '주제', '실습 내용', '수업 형태'],
            session['timetable'],
            header_color=color
        )

        # 산출물
        para(doc, '')
        para(doc, f'산출물: {session["output"]}',
             bold=True, size=9, color=(0, 100, 0))

        # 플랜B
        para(doc, f'플랜B (트러블슈팅):', bold=True, size=9)
        para(doc, session['plan_b'], size=8, italic=True, color=(150, 80, 0))

        # 월요일 액션 플랜
        para(doc, f'월요일 액션 플랜:', bold=True, size=9)
        para(doc, session['monday_action'], size=9, color=(0, 80, 0))

        # 페이지 나누기 (마지막 회차 제외)
        if num < 6:
            doc.add_page_break()

    # ── 사전 IT 협조 체크리스트 ──
    doc.add_page_break()
    styled_heading(doc, '[부록] 사전 IT 협조 체크리스트', level=1)
    make_table(doc,
        ['회차', '필요 사항', '사전 준비', '플랜B'],
        [
            ['1회차', 'claude.ai 웹 접속', '사내 방화벽에서 claude.ai 허용', '강사 핫스팟 + 미러링'],
            ['2회차', 'claude.ai + 엑셀', '동일 + 엑셀 설치 확인', 'Google Sheets 대체'],
            ['3회차', 'Claude Desktop 설치', '관리자 권한 또는 사전 설치', 'claude.ai 웹 파일 업로드 방식'],
            ['4회차', 'claude.ai Artifacts', '별도 설치 불필요', 'HTML 코드 복사 → 브라우저 열기'],
            ['5회차', 'LM Studio 설치\n+ 모델 다운로드', '관리자 권한 + 8GB RAM 이상\n모델 파일 USB 배포', '강사 PC 시연 + claude.ai 실습'],
            ['6회차', 'Claude Projects', 'Pro Plan 이상 구독 확인', '대화 내 문서 업로드 방식'],
        ],
        header_color=color
    )

    # ── 보안 가이드라인 ──
    styled_heading(doc, '[부록] 보안 가이드라인', level=1)
    make_table(doc,
        ['데이터 등급', '정의', '사용 가능 도구', '예시'],
        [
            ['1등급 (자유)', '공개 정보, 일반 질문', 'Claude Pro, 모든 AI', '일반 업무 질문, 공개 데이터'],
            ['2등급 (클라우드OK)', '사내 일반 문서', 'Claude Pro (학습차단 ON)', '내부 보고서, 양식, 절차서'],
            ['3등급 (로컬만)', '민감 데이터', 'LM Studio / Ollama만', '인사정보, 급여, 미공개 재무'],
            ['4등급 (절대불가)', '극비 데이터', 'AI 사용 금지', '영업비밀, 개인정보, 법률문서'],
        ],
        header_color=color
    )

    para(doc, '')
    para(doc, '주의사항:', bold=True, size=9)
    security_notes = [
        'Claude Pro Plan: 설정에서 "학습에 사용하지 않음" 옵션 활성화 가능',
        'Claude Teams Plan: 기본적으로 학습 차단 (기업용)',
        'MCP 파일시스템 연결 시: 연습용 폴더만 지정, 실제 업무 폴더는 등급 확인 후',
        'AI가 생성한 코드: 반드시 외부 URL/API 키/삭제 명령 체크 후 실행',
        'Document AI API 키: 환경변수로 관리, 코드에 직접 입력 금지',
    ]
    add_bullet_list(doc, security_notes, size=8)

    # ── 저장 ──
    path = r'C:\Users\moonk\OneDrive\AI강사준비\수업계획서_Claude_트랙.docx'
    doc.save(path)
    print(f'Saved: {path}')
    return path


if __name__ == '__main__':
    generate_claude_lesson_plan()
