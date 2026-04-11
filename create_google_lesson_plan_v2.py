"""
Google 트랙 수업계획서 v2 생성 스크립트
- 수업 설계 원칙 반영: 30분 성공 경험, 회차당 도구 1~2개, 체험 먼저 이름은 나중에
- 4시간 시간 블록 구조 적용
- 80% 완성 템플릿 제공 원칙
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# STYLE HELPERS
# ============================================================

GOOGLE_GREEN = '0F9D58'
GOOGLE_BLUE = '4285F4'
DARK_GREEN = (0, 100, 0)
GRAY = (100, 100, 100)
ACCENT_ORANGE = (211, 84, 0)


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


def para_multi(doc, segments):
    """여러 스타일 세그먼트를 한 문단에 넣기.
    segments: list of dict with keys: text, bold, size, color, italic
    """
    p = doc.add_paragraph()
    for seg in segments:
        run = p.add_run(seg.get('text', ''))
        run.font.name = 'Malgun Gothic'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
        run.font.size = Pt(seg.get('size', 9))
        run.bold = seg.get('bold', False)
        run.italic = seg.get('italic', False)
        c = seg.get('color')
        if c:
            run.font.color.rgb = RGBColor(*c)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(1)
    return p


def set_cell_shading(cell, color_hex):
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color_hex)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)


def make_table(doc, headers, rows, header_color=GOOGLE_GREEN):
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
        set_cell_shading(cell, header_color)
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
    return table


def info_box(doc, label, text, label_color=ACCENT_ORANGE):
    """라벨 + 텍스트를 한 문단으로 출력 (강조 박스 느낌)"""
    para_multi(doc, [
        {'text': f'[{label}] ', 'bold': True, 'size': 9, 'color': label_color},
        {'text': text, 'size': 9},
    ])


# ============================================================
# SESSION DATA (완전 새로 작성)
# ============================================================

sessions = [
    {
        'num': 1,
        'title': 'AI 첫 만남 + 보안',
        'subtitle': 'Gemini 웹에서 바로 체험 + 데이터 분류 (설치 없음!)',
        'goal': '웹 브라우저만으로 AI를 체험하고, 회사 데이터를 안전하게 다루는 기준을 세운다',
        'success_30min': 'AI가 보고서 초안을 3분에 써줬다',
        'tools': ['Gemini 웹 (gemini.google.com)'],
        'tool_note': '설치 없음 - 웹 브라우저만 필요',
        'rows': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '[보안 안내] 오늘 사용하는 Gemini 무료는 학습에 사용됩니다. 더미 데이터만 사용하세요.\n'
             '[강의 15분] AI란 무엇인가 (1분 시연: Gemini에게 보고서 초안 요청)\n'
             '왜 회사에서 ChatGPT를 막았을까? (삼성 사례)\n'
             '[미니실습 35분] Gemini 웹 접속 → "우리 팀 주간보고 초안 써줘"\n'
             '⚠️ 교육용 더미 데이터만 사용\n'
             '→ 3분 안에 결과 확인 → "AI가 써줬다!" 성공 경험',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습: 데이터 분류 기준표',
             '[시연 10분] 데이터 4등급 분류 기준 설명\n'
             '(1)공개 가능 (2)클라우드OK (3)사내만 (4)절대불가\n'
             '[따라하기 40분] 80% 완성 템플릿에 우리 회사 데이터 분류\n'
             'Gemini에게 "이 데이터는 몇 등급?" 질문하며 분류 완성\n'
             '⚠️ 교육용 더미 데이터만 사용',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화: AI 도구 보안 지도',
             '[시연 10분] Gemini 무료 vs 유료 학습 정책 차이\n'
             'AI에 넣으면 안 되는 것 vs 되는 것 정리\n'
             '[따라하기 40분] AI 코드 검증 4단계 실습:\n'
             '(1)말하기→(2)받기→(3)검증하기→(4)실행하기\n'
             'Gemini에게 코드 요청 → 외부URL 체크 → 안전 확인\n'
             '⚠️ 교육용 더미 데이터만 사용',
             '시연 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용: 업무 자동화 지도 작성',
             '내 반복 업무 리스트업 (Gemini가 도와줌)\n'
             '데이터 등급별 AI 도구 매핑표 작성\n'
             '"이 업무에 AI를 쓸 수 있을까?" Gemini에게 물어보기\n'
             '동료와 비교 → 추가 아이디어 발견',
             '자유 실습\n50분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 월요일 액션 플랜',
             '오늘 배운 것 3줄 정리\n'
             '월요일 액션: "내 업무 1개를 Gemini에게 시켜보기"\n'
             '다음 시간 예고: "자금일보를 버튼 하나로 만든다"',
             '정리 20분'],
        ],
        'output': '데이터 분류 기준표 + 업무 자동화 지도 + AI 보안 체크리스트',
        'plan_b': '인터넷 불가 시: 강사 PC 시연 + 수강생은 종이 워크시트로 데이터 분류',
        'monday_action': '내 반복 업무 1개를 더미 데이터 또는 데이터 등급 확인 후 Gemini 웹에서 시켜보고, 결과를 캡처해서 가져오기',
    },
    {
        'num': 2,
        'title': '자금일보 자동화',
        'subtitle': 'Apps Script 하나만 깊이 + Document AI로 PDF 파싱',
        'goal': 'AI가 생성한 코드로 자금일보를 자동 완성하고, 문서 자동 파싱을 체험한다',
        'success_30min': '버튼 하나로 자금일보 완성',
        'tools': ['Google Sheets + Apps Script (AI가 코드 생성)'],
        'tool_note': '80% 완성 템플릿 제공 - 빈 칸 시작 아님',
        'rows': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '[보안 안내] 오늘 사용하는 AI Studio 무료는 학습에 사용됩니다. 더미 데이터만 사용하세요.\n'
             '[강의 15분] 지난주 월요일 액션 공유 + 자금일보 자동화 시연\n'
             '"이 버튼 하나면 자금일보 끝!" 결과 먼저 보여주기\n'
             '[미니실습 35분] 80% 완성 스프레드시트 열기\n'
             '→ AI Studio에게 "남은 부분 채워줘" → 버튼 클릭 → 완성!\n'
             '⚠️ 교육용 더미 데이터만 사용',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습: Apps Script로 자금일보',
             '[시연 10분] Apps Script 편집기 여는 법 + AI가 코드 써주는 과정\n'
             '[따라하기 40분] AI Studio에게 "은행 거래내역 CSV를 읽어서\n'
             '자금일보 양식에 채우는 Apps Script 만들어줘"\n'
             '⚠️ 교육용 더미 데이터만 사용\n'
             '→ 코드 붙여넣기 → 실행 → 결과 확인\n'
             '→ AI 코드 검증 4단계 적용',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화: Document AI로 PDF 파싱 → Sheets 자동 정리',
             '[시연 10분] Document AI (업스테이지 Document Parse) 시연\n'
             'PDF 영수증 1장 → AI가 자동으로 표 추출 데모\n'
             '[따라하기 40분] Document AI로 PDF 문서 파싱\n'
             '→ 추출된 데이터를 Google Sheets에 자동 정리\n'
             '→ 정리 결과 확인 + 다양한 PDF 문서 테스트\n'
             '⚠️ 교육용 더미 데이터만 사용',
             '시연 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용: 채권연령표 자동화 + 자유 실습',
             '[시연 10분] 채권연령표 완성본 시연\n'
             'AI에게 "매출채권 경과일수 계산 +\n'
             '연령 구간별 분류 Apps Script 만들어줘"\n'
             '→ 코드 적용 → 실행 → 보안 검증\n'
             '⚠️ 교육용 더미 데이터만 사용\n'
             '여유 시간: 자금일보 코드 커스터마이즈',
             '시연 10분\n+ 자유 실습 40분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 월요일 액션 플랜',
             '오늘 만든 것: 자금일보 + Document AI 파싱 + 채권연령표 자동화\n'
             '월요일 액션: "더미 데이터 또는 데이터 등급 확인 후 자금일보 돌려보기"\n'
             '다음 시간 예고: "AI가 스스로 판단하게 만든다"',
             '정리 20분'],
        ],
        'output': '자금일보 자동 생성 Apps Script + Document AI PDF 파싱 → Sheets 자동 정리 + 채권연령표 Apps Script',
        'plan_b': 'Apps Script 실행 불가 시: AI Studio에서 코드 생성까지만 + 강사 PC에서 실행 시연',
        'monday_action': '더미 데이터 또는 데이터 등급 확인 후 은행 거래내역(또는 샘플)으로 자금일보 Apps Script 실행해보기',
    },
    {
        'num': 3,
        'title': 'AI가 혼자 판단하게 만들기',
        'subtitle': 'Opal Agent 실습 먼저 → "방금 한 게 Agent입니다"',
        'goal': 'AI가 데이터를 보고 스스로 판단하는 경험을 한 뒤, Agent 개념을 이해한다',
        'success_30min': '데이터 넣으면 AI가 판단까지',
        'tools': ['Opal Agent (노코드)'],
        'tool_note': '체험 먼저, 용어는 나중에 (MCP/Agentic은 체험 후 소개)',
        'rows': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '[보안 안내] 오늘 사용하는 Opal(Labs)은 학습에 사용됩니다. 더미 데이터만 사용하세요.\n'
             '[강의 15분] 지난주 액션 공유 + "AI가 혼자 판단?" 질문 던지기\n'
             '강사가 Opal Agent 시연: 데이터 입력 → AI가 분류+판단+결과\n'
             '[미니실습 35분] Opal 접속 → 80% 완성 Agent 템플릿 열기\n'
             '→ 샘플 데이터 입력 → AI가 판단 결과 출력\n'
             '⚠️ 교육용 더미 데이터만 사용\n'
             '→ "방금 여러분이 한 게 AI Agent입니다"',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습: 나만의 Agent 만들기',
             '[시연 10분] Opal Agent 스텝 구성법 시연\n'
             '[따라하기 40분] "생산일보 데이터 → AI가 이상치 판단\n'
             '→ 정상/경고/위험으로 분류 → 결과 표시"\n'
             'Agent 3스텝 구성: 입력 → 판단 → 출력\n'
             '(노코드 — 코딩 없이 자연어로 지시)\n'
             '⚠️ 교육용 더미 데이터만 사용',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화: Agent 개념 정리 + 워크플로우',
             '[강의 10분] "방금 한 게 Agent입니다"\n'
             '→ Agent, MCP, Agentic Workflow 용어 설명\n'
             '→ 도구를 연결하는 AI = Agent\n'
             '[따라하기 40분] Opal Agent 워크플로우 확장:\n'
             'Google Forms 입력 → AI 데이터 검증\n'
             '→ Sheets 자동 기록 → 조건별 알림\n'
             '⚠️ 교육용 더미 데이터만 사용',
             '강의 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용: 내 업무에 Agent 적용',
             '본인 업무에 맞는 Agent 시나리오 설계\n'
             'Opal에서 자유롭게 Agent 제작\n'
             '⚠️ 교육용 더미 데이터만 사용\n'
             '동료와 서로의 Agent 테스트 + 피드백\n'
             'Agent가 틀릴 때: 사람이 확인해야 하는 지점 논의',
             '자유 실습\n50분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 월요일 액션 플랜',
             'Agent의 핵심: "AI가 판단까지" (도구 연결 + 자율 판단)\n'
             '월요일 액션: "내 업무 1개에 Opal Agent 적용해보기"\n'
             '다음 시간 예고: "대시보드 만들어서 팀에 공유"',
             '정리 20분'],
        ],
        'output': 'Opal Agent 워크플로우 (노코드) + Agent 개념 이해 + 내 업무 Agent 설계서',
        'plan_b': 'Opal 접속 불가 시: AI Studio에서 "판단 프롬프트" 실습 (수동 Agent 체험)',
        'monday_action': 'Opal Agent로 내 업무 데이터 1건을 자동 판단시켜 보기',
    },
    {
        'num': 4,
        'title': '대시보드 만들고 팀에 공유',
        'subtitle': 'AI Studio로 HTML 대시보드 1개만 제대로',
        'goal': 'AI로 대시보드를 만들고, 팀원에게 실제로 공유하는 경험을 한다',
        'success_30min': '내가 만든 대시보드를 팀원이 봤다',
        'tools': ['AI Studio (바이브코딩)'],
        'tool_note': 'Mixboard/Pomely는 "이런 것도 있다" 수준 데모만',
        'rows': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '[보안 안내] 오늘 사용하는 AI Studio/Opal은 학습에 사용됩니다. 더미 데이터만 사용하세요.\n'
             '[강의 15분] 바이브코딩이란? (자연어로 앱 만들기)\n'
             '강사 시연: AI Studio에게 "매출 차트 대시보드 만들어줘"\n'
             '→ 30초 만에 대시보드 완성 → "이걸 여러분도 합니다"\n'
             '[미니실습 35분] AI Studio 접속 → 80% 완성 프롬프트 제공\n'
             '→ "생산 실적 대시보드 HTML 만들어줘"\n'
             '⚠️ 교육용 더미 데이터만 사용\n'
             '→ 결과 다운로드 → 브라우저에서 열기 → 성공!',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습: 대시보드 커스터마이즈',
             '[시연 10분] 대시보드에 내 데이터 넣는 법\n'
             '[따라하기 40분] Google Sheets에서 데이터 복사\n'
             '→ AI Studio에게 "이 데이터로 차트 대시보드 업데이트해줘"\n'
             '⚠️ 교육용 더미 데이터만 사용\n'
             '→ 차트 종류 변경, 색상 변경, 필터 추가\n'
             '→ "월별 매출 추이 + 부서별 비교" 대시보드 완성',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화: 팀에 공유하기',
             '[시연 10분] Opal 링크 공유 방법 시연\n'
             '[따라하기 40분]\n'
             '방법1: Opal에 대시보드 올리기 → 링크 공유\n'
             '⚠️ 교육용 더미 데이터만 사용\n'
             '[ ] 접근 통제 체크리스트:\n'
             '  [ ] 공유 대상: 특정 사용자만\n'
             '  [ ] 만료일 설정\n'
             '  [ ] Labs 서비스 = 실제 데이터 금지\n'
             '→ 옆 사람이 링크로 접속 → "팀원이 봤다!"\n'
             '방법2 (플랜B): HTML 파일을 Google Drive에 올리기\n'
             '→ 공유 링크 생성 → 팀원 접속 확인',
             '시연 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용: 자유 제작 + Google Labs 데모',
             '[데모 15분] Mixboard(컨셉보드), Pomely(마케팅) 소개\n'
             '⚠️ 교육용 더미 데이터만 사용\n'
             '→ "이런 것도 있다" 수준, 깊이 실습은 안 함\n'
             '[자유 실습 35분] 본인 업무에 맞는 대시보드/미니앱 제작\n'
             '동료와 상호 접속 + 피드백',
             '데모 15분\n+ 자유 실습 35분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 월요일 액션 플랜',
             '오늘 만든 것: 나만의 대시보드 + 팀 공유 완료\n'
             '월요일 액션: "더미 데이터 또는 데이터 등급 확인 후 대시보드 만들어서 팀장에게 보여주기"\n'
             '다음 시간 예고: "인터넷 없이도 AI를 쓴다"',
             '정리 20분'],
        ],
        'output': 'HTML 대시보드 + Opal/Drive 링크로 팀 공유 완료',
        'plan_b': 'Opal 링크 불가 시: HTML 파일을 Google Drive 공유 또는 USB로 직접 전달',
        'monday_action': '더미 데이터 또는 데이터 등급 확인 후 대시보드 만들어서 팀장/동료 1명에게 공유하기',
    },
    {
        'num': 5,
        'title': '로컬 AI - 인터넷 없이도',
        'subtitle': 'LM Studio GUI 먼저 (클릭) → Ollama CLI 나중에',
        'goal': '인터넷 없이 AI를 실행하는 경험을 하고, 보안 등급별 AI 선택 기준을 세운다',
        'success_30min': '인터넷 끊어도 AI가 답한다',
        'tools': ['LM Studio (GUI)', 'Ollama (선택, CLI)'],
        'tool_note': 'GUI 먼저 → CLI 나중에 / 설치 불가 시 강사 PC 시연',
        'rows': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '[보안 안내] 오늘 사용하는 로컬 AI(LM Studio)는 인터넷 없이 동작합니다. 3등급 데이터까지 사용 가능합니다.\n'
             '[강의 15분] "인터넷 끊어도 AI가 된다?" → 로컬 AI 개념\n'
             '강사 시연: 와이파이 끄기 → LM Studio에서 질문 → 답변!\n'
             '[미니실습 35분] LM Studio 설치 (사전 배포 USB 또는 다운로드)\n'
             '→ Solar 모델 로드 → 질문 입력 → 오프라인 답변 확인\n'
             '→ "인터넷 없이 AI가 답했다!" 성공 경험',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습: Solar vs Gemma4 비교',
             '[시연 10분] 두 모델 비교 기준 설명 (한국어, 코드, 속도)\n'
             '[따라하기 40분] LM Studio에서 모델 교체:\n'
             '같은 질문을 Solar → Gemma4로 비교\n'
             '한국어 자연스러움 / 코드 생성 품질 / 응답 속도\n'
             '비교표 작성 (80% 완성 템플릿 제공)',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화: Ollama CLI 체험 (선택)',
             '[시연 10분] Ollama 설치 + 명령어 1개로 AI 실행\n'
             '"ollama run solar" → 터미널에서 AI 대화\n'
             '[따라하기 40분]\n'
             '선택A: Ollama 설치 → solar 실행 → 업무 질문\n'
             '선택B (비개발자): LM Studio에서 심화 실습 계속\n'
             '→ 파일 요약, 보고서 초안 등 실무 활용',
             '시연 10분\n+ 실습 40분'],
            ['[블록4]\n2:50~3:40',
             '응용: 보안 등급별 AI 선택 가이드',
             '도구별 정확한 보안 조건 총정리:\n'
             '- 1등급(공개): Gemini 웹 OK\n'
             '- 2등급(클라우드OK): AI Studio/Opal OK\n'
             '- 3등급(사내만): LM Studio/Ollama (로컬)\n'
             '- 4등급(절대불가): AI 사용 불가\n'
             '우리 회사에 맞는 AI 도구 조합 설계',
             '자유 실습\n50분'],
            ['[블록5]\n3:40~4:00',
             '정리 + 월요일 액션 플랜',
             '클라우드 AI vs 로컬 AI 선택 기준 정리\n'
             '월요일 액션: "3등급 데이터 1건을 LM Studio로 처리해보기"\n'
             '다음 시간 예고: "지금까지 배운 전부를 연결한다"',
             '정리 20분'],
        ],
        'output': 'LM Studio 로컬 AI 실행 + Solar vs Gemma4 비교표 + 보안 등급별 AI 선택 가이드',
        'plan_b': '설치 불가 시: 강사 PC를 프로젝터로 시연 + 수강생은 Gemini 웹에서 동일 질문 비교',
        'monday_action': '보안 3등급 데이터 1건을 LM Studio(로컬 AI)로 처리해보기',
    },
    {
        'num': 6,
        'title': '전부 연결하기',
        'subtitle': '단계별 가이드 + 페어 학습',
        'goal': '1~5회차에서 배운 도구를 연결해 본인 업무의 자동화 파이프라인을 완성한다',
        'success_30min': '내 업무가 자동화됐다!',
        'tools': ['1~5회차 도구 전체 (선택 조합)'],
        'tool_note': '조별 프로젝트 + 발표',
        'rows': [
            ['[블록1]\n0:00~0:50',
             '도입 + 30분 성공 경험',
             '[보안 안내] 오늘은 종합 실습입니다. 각 도구의 데이터 등급을 확인하고 더미 데이터만 사용하세요.\n'
             '[강의 15분] 1~5회차 총정리 (도구 지도 한 장)\n'
             '"오늘 목표: 내 업무가 자동화된다"\n'
             '강사가 end-to-end 시연: 데이터 입력 → AI 판단\n'
             '→ 보고서 생성 → 대시보드 → 팀 공유\n'
             '[미니실습 35분] 80% 완성 파이프라인 템플릿 제공\n'
             '→ 본인 업무 데이터 넣기 → "내 업무가 자동화됐다!"',
             '강의 15분\n+ 실습 35분'],
            ['[블록2]\n0:50~1:40',
             '메인 실습: 조별 프로젝트 설계',
             '[시연 10분] 프로젝트 진행 방법 + 역할 분담\n'
             '[따라하기 40분] 2~3인 조 편성 → 업무 시나리오 선택\n'
             '파이프라인 설계: 어떤 도구를 어떤 순서로?\n'
             '- 데이터 수집 (Gemini/Sheets)\n'
             '- AI 처리 (AI Studio/Opal Agent)\n'
             '- 결과 공유 (대시보드/Drive)',
             '시연 10분\n+ 실습 40분'],
            ['', '휴식', '1:40~2:00 (20분)', ''],
            ['[블록3]\n2:00~2:50',
             '심화: 조별 프로젝트 구현',
             '페어 학습: 잘하는 사람 + 배우는 사람 짝꿍\n'
             '강사가 조별 순회하며 트러블슈팅\n'
             '중간 체크포인트: "여기까지 되면 성공"\n'
             '막히는 조에게 단계별 가이드 카드 제공',
             '프로젝트\n50분'],
            ['[블록4]\n2:50~3:40',
             '발표 준비 + 발표',
             '[준비 20분] 조별 발표 자료 정리\n'
             '- 어떤 업무를 자동화했나?\n'
             '- 어떤 도구 조합을 썼나?\n'
             '- 얼마나 시간이 절약되나?\n'
             '[발표 30분] 조별 5분 발표 + 상호 피드백',
             '준비 20분\n+ 발표 30분'],
            ['[블록5]\n3:40~4:00',
             '수료 + 지속 학습 가이드 + 데이터 정리',
             '6회차 전체 요약: "이제 AI로 업무를 바꿀 수 있다"\n'
             '지속 학습 로드맵 제공\n'
             '사내 AI 활용 커뮤니티 안내\n'
             '[교육 후 데이터 정리 가이드]\n'
             '  [ ] Gemini 대화 삭제\n'
             '  [ ] Sheets 실습파일 삭제\n'
             '  [ ] Opal Agent 정리\n'
             '  [ ] AI Studio 프로젝트 정리\n'
             '수료',
             '정리 20분'],
        ],
        'output': '본인 업무 end-to-end AI 자동화 파이프라인 + 조별 발표 자료',
        'plan_b': '프로젝트 진행 어려운 조: 강사 제공 완성 예제를 따라하며 커스터마이즈',
        'monday_action': '완성한 파이프라인을 실제 업무에 적용하고, 1주일 뒤 결과를 팀에 공유하기',
    },
]


# ============================================================
# GENERATE DOCUMENT
# ============================================================

def build_document():
    doc = create_doc()

    # ----- Title -----
    title = doc.add_heading('수업계획서 - Google 트랙 (v2)', level=0)
    for run in title.runs:
        run.font.name = 'Malgun Gothic'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
        run.font.size = Pt(18)

    para(doc, 'AI로 쉬워지는 직장인 업무 자동화', bold=True, size=12)
    para(doc, 'Google AI Studio / Gemini / Apps Script / Opal Agent / LM Studio',
         size=10, color=GRAY)

    # ----- 수업 설계 원칙 -----
    styled_heading(doc, '수업 설계 원칙', level=2)
    principles = [
        '30분 성공 경험: 매 회차 시작 30분 안에 수강생 화면에 결과물이 나온다',
        '회차당 핵심 도구 1~2개: 5개 도구가 아니라 1개를 깊이 있게',
        '체험 먼저, 이름은 나중에: 실습을 먼저 하고, "방금 한 게 Agent입니다" 식으로 개념 소개',
        '웹 브라우저 우선: 1회차는 설치 없이 웹 AI만 사용, 설치형은 5회차로 이동',
        '80% 완성 템플릿: 빈 칸에서 시작하지 않음 - 거의 다 된 것에서 나머지를 채우는 방식',
    ]
    for pr in principles:
        para(doc, f'  {pr}', size=9)

    # ----- 4시간 시간 블록 -----
    styled_heading(doc, '4시간 시간 블록 (매 회차 동일 구조)', level=2)
    make_table(doc,
        ['블록', '시간', '내용', '구성'],
        [
            ['블록1', '0:00~0:50', '도입 + 30분 성공 경험', '강의 15분 + 미니실습 35분'],
            ['블록2', '0:50~1:40', '메인 실습', '시연 10분 + 따라하기 40분'],
            ['휴식', '1:40~2:00', '20분 휴식', '-'],
            ['블록3', '2:00~2:50', '심화 실습', '시연 10분 + 따라하기 40분'],
            ['블록4', '2:50~3:40', '응용/자유 실습', '50분'],
            ['블록5', '3:40~4:00', '정리 + 월요일 액션 플랜', '20분'],
        ],
        header_color=GOOGLE_GREEN
    )

    # ----- 과정 정보 -----
    styled_heading(doc, '과정 정보', level=2)
    make_table(doc,
        ['항목', '내용'],
        [
            ['과정명', 'AI로 쉬워지는 직장인 업무 자동화'],
            ['트랙', 'Google 트랙'],
            ['교육 대상', '외부 AI 사용 제한 기업 재직자 (비개발자)'],
            ['총 교육 시간', '24시간 (6회 x 4시간)'],
            ['교육 방식', '이론 20% + 실습 80%'],
            ['핵심 원칙', '30분 성공 경험 / 회차당 도구 1~2개 / 80% 완성 템플릿'],
            ['주요 AI 도구', 'Gemini 웹, Google AI Studio, Apps Script, Opal Agent, LM Studio'],
            ['보조 도구', 'Document AI(2회차 심화), Mixboard/Pomely(데모), Ollama(선택)'],
            ['로컬 AI', 'LM Studio(GUI) + Ollama(CLI 선택) / Solar, Gemma4'],
        ],
        header_color=GOOGLE_GREEN
    )

    # ----- 6회차 개요 -----
    styled_heading(doc, '6회차 개요', level=2)
    overview_rows = []
    for s in sessions:
        overview_rows.append([
            f'{s["num"]}회차',
            s['title'],
            s['success_30min'],
            ', '.join(s['tools']),
        ])
    make_table(doc,
        ['회차', '주제', '30분 성공 경험', '핵심 도구'],
        overview_rows,
        header_color=GOOGLE_GREEN
    )

    # ----- 각 회차 상세 -----
    for s in sessions:
        doc.add_page_break()

        styled_heading(doc, f'{s["num"]}회차: {s["title"]}', level=1)
        para(doc, s['subtitle'], size=10, color=GRAY, italic=True)

        # 학습 목표
        para_multi(doc, [
            {'text': '학습 목표: ', 'bold': True, 'size': 9},
            {'text': s['goal'], 'size': 9},
        ])

        # 30분 성공 경험
        info_box(doc, '30분 성공 경험', f'"{s["success_30min"]}"')

        # 핵심 도구
        info_box(doc, '핵심 도구', ' / '.join(s['tools']))
        if s.get('tool_note'):
            para(doc, f'  * {s["tool_note"]}', size=8, color=GRAY, italic=True)

        # 시간표 테이블
        para(doc, '')  # spacer
        make_table(doc,
            ['시간', '주제', '실습 내용', '수업 형태'],
            s['rows'],
            header_color=GOOGLE_GREEN
        )

        # 산출물
        para(doc, '')  # spacer
        para_multi(doc, [
            {'text': '산출물: ', 'bold': True, 'size': 9, 'color': DARK_GREEN},
            {'text': s['output'], 'size': 9, 'color': DARK_GREEN},
        ])

        # 플랜B
        info_box(doc, '플랜B', s['plan_b'], label_color=(180, 0, 0))

        # 월요일 액션 플랜
        info_box(doc, '월요일 액션 플랜', s['monday_action'], label_color=(0, 0, 160))

    # ----- 부록: 사전 준비 체크리스트 -----
    doc.add_page_break()
    styled_heading(doc, '부록: 사전 준비 체크리스트 (IT 부서 협조)', level=1)

    make_table(doc,
        ['항목', '필요 시점', '상세'],
        [
            ['Google 계정', '1회차 전', 'Gemini 웹 + AI Studio + Google Labs(Opal) 접속용'],
            ['웹 브라우저', '1회차', 'Chrome 또는 Edge (최신 버전)'],
            ['Google Sheets 접근', '2회차 전', 'Apps Script 실행을 위한 Google Workspace 접근'],
            ['LM Studio 설치 파일', '5회차 전', 'USB 또는 사전 다운로드 (약 500MB)'],
            ['Ollama 설치 파일 (선택)', '5회차', 'CLI 선호자용, 선택 사항'],
            ['AI 모델 파일', '5회차 전', 'Solar/Gemma4 모델 (약 2~4GB) USB 배포 또는 다운로드'],
            ['인터넷 접속', '1~4회차', '웹 AI 실습용 (5회차는 오프라인 가능)'],
            ['프로젝터/스크린', '매 회차', '강사 시연용'],
        ],
        header_color=GOOGLE_BLUE
    )

    para(doc, '')
    styled_heading(doc, '부록: 보안 원칙 요약', level=2)
    make_table(doc,
        ['데이터 등급', '설명', '사용 가능한 AI 도구'],
        [
            ['1등급 (공개)', '외부 공개 가능한 데이터', 'Gemini 웹, AI Studio, Opal, 모든 도구 OK'],
            ['2등급 (클라우드OK)', '사내용이지만 클라우드 OK', 'AI Studio(유료), Google Workspace 연동'],
            ['3등급 (사내만)', '사내에서만 처리 가능', 'LM Studio, Ollama (로컬 AI만)'],
            ['4등급 (절대불가)', 'AI 처리 자체 불가', 'AI 사용 불가 - 수동 처리'],
        ],
        header_color='D93025'
    )

    para(doc, '')
    para(doc, 'AI 코드 검증 4단계:', bold=True, size=9)
    steps = [
        '1단계 - 말하기: AI에게 원하는 것을 명확히 요청',
        '2단계 - 받기: AI가 생성한 코드/결과를 받음',
        '3단계 - 검증하기: 외부 URL, API 키, 삭제 명령 등 위험 요소 확인',
        '4단계 - 실행하기: 검증 후에만 실행',
    ]
    for step in steps:
        para(doc, f'  {step}', size=9)

    # ----- Save -----
    output_path = r'C:\Users\moonk\OneDrive\AI강사준비\수업계획서_Google_트랙.docx'
    doc.save(output_path)
    print(f'Saved: {output_path}')
    return output_path


if __name__ == '__main__':
    build_document()
