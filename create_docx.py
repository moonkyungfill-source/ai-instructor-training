from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

style = doc.styles['Normal']
font = style.font
font.name = 'Malgun Gothic'
font.size = Pt(9)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')

# Helper functions
def add_heading_styled(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = 'Malgun Gothic'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    return h

def add_para(text, bold=False, size=9, align=None, color=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = 'Malgun Gothic'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(1)
    return p

def add_table(headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Header
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
        from docx.oxml import OxmlElement
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), '2B579A')
        shading_elm.set(qn('w:val'), 'clear')
        cell._tc.get_or_add_tcPr().append(shading_elm)
        run.font.color.rgb = RGBColor(255, 255, 255)
    # Rows
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

# ===== TITLE =====
title = doc.add_heading('AI로 쉬워지는 직장인 업무 자동화', level=0)
for run in title.runs:
    run.font.name = 'Malgun Gothic'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    run.font.size = Pt(18)

add_para('ChatGPT 없어도 괜찮아 — 회사가 허용한 AI Agent가 VBA도, Script도, 대시보드도 대신 만들어준다', size=10, color=(80, 80, 80))

# ===== 1. 과정 개요 =====
add_heading_styled('1. 과정 개요', level=2)
add_table(
    ['항목', '내용'],
    [
        ['교육 대상', '외부 AI 사용이 제한된 기업 재직자 (비개발자). 대학생·청년 대상 변형 가능'],
        ['총 교육 시간', '24시간 (6회 × 4시간), 이론 20% + 실습 80%'],
        ['핵심 구조', '기업별 AI 환경에 맞춘 Google / Microsoft / Claude 3트랙 + 로컬 AI(Ollama) 통합'],
    ]
)

# ===== 2. 기획 배경 =====
add_heading_styled('2. 기획 배경 — AI Agent 시대, 그러나 회사에선 AI가 막혀있다', level=2)

add_para('2026년은 AI Agent의 시대입니다. MCP(Model Context Protocol)로 AI가 도구를 연결하고, Agentic Workflow로 업무를 자동 처리하며, 바이브코딩으로 비개발자도 앱을 만듭니다. 그러나 국내 대기업·금융권·공공기관의 70% 이상이 외부 AI를 차단하고 있습니다.', size=9)

add_para('진짜 문제는 도구 자체가 어려운 것입니다. VBA, Apps Script, Python은 회사에 있지만 활용하지 못합니다. 본 과정은 "코드를 배우는 게 아니라, AI Agent에게 시키는 법을 배운다"는 철학으로, 회사가 허용한 AI만으로 업무 자동화를 실현합니다. 업스테이지 Solar LLM(로컬 AI)과 Document AI(문서 파싱)를 커리큘럼에 직접 통합합니다.', size=9)

add_para('기업마다 허용된 AI가 다르기 때문에, 교육 전 기업 환경을 파악하고 해당 트랙으로 강의합니다:', bold=True, size=9)

add_table(
    ['Google Workspace 기업', 'Microsoft 365 기업', 'Claude 도입 기업', '모든 기업 (공통)'],
    [
        ['AI Studio, Gemini\nApps Script', 'Copilot CLI, VBA\nPower Query', 'Claude in Chrome\nExcel·Word', 'Ollama + Solar/Gemma4\n(오프라인)'],
    ]
)

# ===== 3. 커리큘럼 =====
add_heading_styled('3. 커리큘럼 (6회 × 4시간 = 24시간)', level=2)

add_table(
    ['회차', '주제', '핵심 실습', '가져가는 것'],
    [
        ['1', '사내에서 AI를 안전하게\n쓰는 법', '데이터 4등급 분류, 도구별 보안 조건,\nAI 보안 세팅, 코드 검증 체크리스트', '데이터 분류표\n+ 보안 세팅 완료'],
        ['2', 'AI + 스프레드시트\n자동화', '프롬프트 엔지니어링, 자금일보·채권연령표\n자동화, Upstage Document AI 문서 파싱', '자금일보 자동 생성\n매크로'],
        ['3', 'AI Agent 워크플로우', 'AI Agent·MCP·Agentic Workflow 개념,\n생산일보, 데이터→보고서→이메일 자동화', 'AI Agent 보고서\n파이프라인'],
        ['4', '바이브코딩\n+ 로컬 공유', 'HTML 대시보드 제작, 업무 도구 제작,\n로컬 네트워크로 팀원에게 공유', '팀 대시보드\n+ 로컬 공유'],
        ['5', '로컬 AI 심화', 'Solar LLM vs Gemma 4 비교, 문서 일괄\n처리, LM Studio GUI, 오프라인 파이프라인', '로컬 AI 실무\n환경 완성'],
        ['6', '종합 프로젝트', '전체 연결: 수집→Document AI→AI Agent\n→보고서→대시보드→공유→발송', 'end-to-end\n자동화 파이프라인'],
    ]
)

add_para('')
add_para('매 회차 공통 패턴: 말하기(프롬프트) → 받기(코드) → 검증하기(보안 체크) → 실행하기', bold=True, size=9)
add_para('STEP 3 "검증"이 이 과정의 차별점. AI가 만든 코드의 외부 호출·키 노출·삭제 명령을 체크합니다.', size=8, color=(100, 100, 100))

# ===== PAGE BREAK =====
doc.add_page_break()

# ===== 4. 비즈니스 활용 케이스 =====
add_heading_styled('4. 비즈니스 활용 케이스', level=2)

add_table(
    ['부서', 'AI가 만들어주는 것', '공유 방식', '효과'],
    [
        ['재무팀', '자금일보 + 자금 현황 대시보드', '팀장 PC → 로컬 네트워크', '매일 1시간 → 5분'],
        ['생산관리', '생산일보 + 실적 모니터', '공장 태블릿 실시간 조회', '수기 기록 제거'],
        ['회계팀', '채권연령표 + 연령 분석 차트', '회의실 KPI 보드', '월말 2일 → 2시간'],
        ['영업팀', '일일 매출 현황판', '팀원 브라우저 접속', '실시간 파악'],
        ['총무', 'PDF 영수증 파싱 (Document AI)', '스프레드시트 자동 정리', '수작업 입력 제거'],
    ]
)

add_para('모든 도구는 클라우드 배포 없이 같은 와이파이 내 팀원과 공유 가능', size=8, color=(100, 100, 100))

# ===== 5. 2026 AI 트렌드 반영 =====
add_heading_styled('5. 2026 AI 트렌드 반영 근거', level=2)

add_table(
    ['트렌드', '커리큘럼 반영'],
    [
        ['AI Agent / Agentic Workflow', '3회차: AI Agent 기반 워크플로우 자동화, 도구 연결 패턴'],
        ['MCP (Model Context Protocol)', '3회차: AI가 스프레드시트·이메일·파일시스템을 직접 연결'],
        ['바이브코딩', '4회차: 대화만으로 HTML 대시보드·업무 도구 제작 + 로컬 공유'],
        ['로컬/온디바이스 AI', '전 회차: Ollama + Solar/Gemma4 오프라인 AI 통합'],
        ['한국어 특화 LLM', '5회차: 업스테이지 Solar LLM 한국어 성능 비교 실습'],
        ['Document AI', '2회차: 업스테이지 Document Parse로 PDF 자동 파싱'],
        ['AI 거버넌스', '1회차+전 회차: 보안 등급별 도구 분류, 코드 보안 검증 4단계'],
    ]
)

# ===== 6. 기대 성과 =====
add_heading_styled('6. 기대 성과', level=2)

add_table(
    ['지표', '목표'],
    [
        ['반복 문서 작업 시간 절감', '주당 5~10시간'],
        ['자동화 결과물', '수강생 1인당 5개 이상'],
        ['보안 규정 준수율', '100% (기업 승인 도구 + 코드 검증 4단계)'],
        ['AI Agent 활용 역량', '프롬프트 → 코드 생성 → 검증 → 실행 사이클 체득'],
    ]
)

# ===== 7. 차별점 =====
add_heading_styled('7. 타 과정과의 차별점', level=2)

add_table(
    ['기존 AI 교육', '이 과정'],
    [
        ['ChatGPT 사용법 → 회사에서 못 씀', '회사 승인된 AI로만 진행'],
        ['VBA 문법 배우기 → 포기', 'AI Agent가 코드를 짜줌, 나는 실행만'],
        ['AI Agent 이론만', '실제 업무 문서(자금일보, 채권연령표)로 실습'],
        ['도구 1개만 다룸', 'Google / MS / Claude 3트랙 + 로컬 AI'],
        ['보안 언급 없음', '데이터 분류 + AI 코드 보안 검증 절차 내장'],
        ['업스테이지 기술 미활용', 'Solar LLM + Document AI 커리큘럼 포함'],
    ]
)

# ===== Footer =====
add_para('')
add_para('본 과정은 업스테이지의 "Making AI Beneficial" 비전 아래, AI Agent·Agentic Workflow·바이브코딩 등 2026년 핵심 AI 트렌드를 비개발자도 활용할 수 있도록 풀어낸 실전 교육입니다. 업스테이지 Solar LLM과 Document AI를 커리큘럼에 직접 통합하여, AI 기술이 보안 환경에서도 실질적 업무 가치를 만드는 과정을 보여줍니다.', size=8, color=(80, 80, 80))

# Save
output_path = r'C:\Users\moonk\OneDrive\AI강사준비\강의기획서_제출용.docx'
doc.save(output_path)
print(f'Saved: {output_path}')
