"""
Microsoft + Claude 트랙 수업계획서 v2 생성
수업 설계 원칙 완전 적용: 30분 성공, 도구 1~2개, 체험→이름, 웹 우선
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Style helpers (identical to Google v2)
def create_doc():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(2); s.bottom_margin = Cm(1.5); s.left_margin = Cm(2); s.right_margin = Cm(2)
    st = doc.styles['Normal']; st.font.name = 'Malgun Gothic'; st.font.size = Pt(9)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    return doc

def styled_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs: r.font.name = 'Malgun Gothic'; r.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    return h

def para(doc, text, bold=False, size=9, color=None, italic=False):
    p = doc.add_paragraph(); run = p.add_run(text)
    run.font.name = 'Malgun Gothic'; run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    run.font.size = Pt(size); run.bold = bold; run.italic = italic
    if color: run.font.color.rgb = RGBColor(*color)
    p.paragraph_format.space_after = Pt(2); p.paragraph_format.space_before = Pt(1)
    return p

def para_multi(doc, segments):
    p = doc.add_paragraph()
    for seg in segments:
        run = p.add_run(seg.get('text', ''))
        run.font.name = 'Malgun Gothic'; run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
        run.font.size = Pt(seg.get('size', 9)); run.bold = seg.get('bold', False); run.italic = seg.get('italic', False)
        c = seg.get('color')
        if c: run.font.color.rgb = RGBColor(*c)
    p.paragraph_format.space_after = Pt(2); p.paragraph_format.space_before = Pt(1)
    return p

def set_cell_shading(cell, color_hex):
    sh = OxmlElement('w:shd'); sh.set(qn('w:fill'), color_hex); sh.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(sh)

def make_table(doc, headers, rows, header_color='2B579A'):
    table = doc.add_table(rows=1+len(rows), cols=len(headers))
    table.style = 'Table Grid'; table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]; cell.text = ''; p = cell.paragraphs[0]
        run = p.add_run(h); run.bold = True; run.font.size = Pt(8)
        run.font.name = 'Malgun Gothic'; run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER; set_cell_shading(cell, header_color)
        run.font.color.rgb = RGBColor(255,255,255)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.rows[ri+1].cells[ci]; cell.text = ''; p = cell.paragraphs[0]
            run = p.add_run(str(val)); run.font.size = Pt(8)
            run.font.name = 'Malgun Gothic'; run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    return table

def info_box(doc, label, text, label_color=(211,84,0)):
    para_multi(doc, [{'text': f'[{label}] ', 'bold': True, 'size': 9, 'color': label_color}, {'text': text, 'size': 9}])

def build_track(track_name, color, subtitle, ai_tool, extra, sessions, output_path):
    doc = create_doc()
    t = doc.add_heading(f'수업계획서 - {track_name} 트랙', level=0)
    for r in t.runs: r.font.name='Malgun Gothic'; r.element.rPr.rFonts.set(qn('w:eastAsia'),'Malgun Gothic'); r.font.size=Pt(18)
    para(doc, 'AI로 쉬워지는 직장인 업무 자동화', bold=True, size=12)
    para(doc, subtitle, size=10, color=(100,100,100))

    styled_heading(doc, '과정 정보', level=2)
    make_table(doc, ['항목','내용'], [
        ['과정명','AI로 쉬워지는 직장인 업무 자동화'],
        ['트랙', f'{track_name} 트랙 ({subtitle})'],
        ['교육 대상','외부 AI 사용 제한 기업 재직자 (비개발자)'],
        ['총 교육 시간','24시간 (6회 x 4시간)'],
        ['교육 방식','이론 20% + 실습 80%'],
        ['주요 AI 도구', ai_tool],
        ['추가 도구', extra],
        ['로컬 AI (공통)','Ollama + Upstage Solar / Gemma 4 (5회차)'],
    ], header_color=color)

    styled_heading(doc, '수업 설계 원칙', level=2)
    for principle in [
        '30분 성공 경험: 매 회차 시작 30분 안에 수강생 화면에 결과물',
        '회차당 핵심 도구 1~2개: 하나를 제대로 > 다섯을 어렴풋이',
        '체험 먼저, 이름은 나중에: 실습 후 "방금 한 게 Agent입니다"',
        '웹 브라우저 우선: 설치형 도구에 반드시 웹 대안(플랜B) 준비',
        '80% 완성 템플릿: 빈 칸에서 시작하지 않음',
        'AI 코드 검증 4단계: 말하기 -> 받기 -> 검증하기 -> 실행하기',
    ]:
        para(doc, f'  {principle}', size=8)

    for i, s in enumerate(sessions):
        if i > 0 and i % 2 == 0: doc.add_page_break()
        styled_heading(doc, f'{s["num"]}회차: {s["title"]}', level=2)
        para(doc, s['subtitle'], size=9, color=(100,100,100), italic=True)
        info_box(doc, '학습 목표', s['goal'])
        info_box(doc, '30분 성공 경험', s['success_30min'], label_color=(0,100,0))
        info_box(doc, '핵심 도구', ', '.join(s['tools']))
        if s.get('planb'): info_box(doc, '플랜B', s['planb'], label_color=(180,0,0))
        para(doc, '')
        make_table(doc, ['시간','주제','실습 내용','수업 형태'], s['rows'], header_color=color)
        para(doc, '')
        info_box(doc, '산출물', s['output'], label_color=(0,100,0))
        info_box(doc, '월요일 액션', s['monday'], label_color=(0,0,150))
        para(doc, '')

    # Appendix
    doc.add_page_break()
    styled_heading(doc, '부록: 사전 준비 체크리스트', level=2)
    make_table(doc, ['항목','담당','시기'], [
        ['수강생 AI 도구 계정 확보','강사+수강생','교육 1주 전'],
        ['도메인 화이트리스트 등록','IT팀','교육 2주 전'],
        ['VBA 매크로 실행 허용 여부 확인','IT팀','교육 2주 전'],
        ['교육용 샘플 데이터 준비 (더미)','강사','교육 1주 전'],
        ['80% 완성 템플릿 배포','강사','교육 전일'],
        ['5회차 Ollama/LM Studio 설치 가능 여부','IT팀','교육 2주 전'],
        ['용어 사전 카드 인쇄','강사','교육 전일'],
    ], header_color=color)

    doc.save(output_path)
    print(f'Saved: {output_path}')

# ============================================================
# MICROSOFT TRACK
# ============================================================
# MS 보안 브리핑 문구 (매 회차 블록1에 삽입)
MS_SECURITY_BRIEFING = '\n[보안] Copilot Free는 학습 사용. Business 플랜 확인! 더미 데이터만 사용하세요.'
# MS 더미 데이터 경고 (Copilot Free 실습에 추가)
MS_DUMMY_WARNING = '\n⚠️ Business 플랜 확인! Free는 학습 사용. 더미 데이터만 사용'

ms_sessions = [
    {'num':1, 'title':'AI 첫 만남 + 보안',
     'subtitle':'GitHub Copilot에서 바로 체험 + 데이터 분류 (설치 없음!)',
     'goal':'GitHub Copilot 웹으로 즉시 업무 체험 + 회사 데이터 보안 기준 수립',
     'success_30min':'GitHub Copilot이 VBA 코드를 10초에 써줬다',
     'tools':['GitHub Copilot (github.com/copilot)'],
     'planb':'Copilot 접속 불가 시 ChatGPT(chat.openai.com) 또는 Copilot 웹(copilot.microsoft.com)',
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 30분 성공 경험',
          '[강의 15분] AI란 무엇인가 + 킬러 데모\n(완성된 자금일보 자동화 + 대시보드 시연)\n왜 ChatGPT가 막혔을까? (삼성 사례)\nGitHub Copilot = 코딩 전문 AI'
          + MS_SECURITY_BRIEFING +
          '\n[미니실습 35분] GitHub Copilot 웹 접속\n→ "엑셀 날짜별 정리하는 VBA 코드 써줘"\n→ 10초 만에 완성 → "AI가 코드를 짰다!"'
          + MS_DUMMY_WARNING,'강의 15분\n+ 실습 35분'],
         ['[블록2]\n0:50~1:40','메인 실습: 데이터 분류 기준표',
          '[시연 10분] 데이터 4등급 분류\n(1)공개 (2)클라우드OK (3)사내만 (4)절대불가\nGitHub Copilot Business vs Free 학습 정책 차이\n[따라하기 40분] 80% 완성 템플릿에 우리 회사 데이터 분류\nCopilot에게 "이 데이터는 몇 등급?" 물으며 완성'
          + MS_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','심화: Copilot 웹에서 VBA/수식 코드 더 만들어보기',
          '[자유 실습 50분] Copilot 웹에서 다양한 VBA/수식 코드 요청\n→ "IF 중첩 수식 만들어줘", "피벗 매크로 만들어줘"\n→ 코드 검증(4단계) 반복 연습\n→ 외부URL/API키/삭제명령 체크 실습'
          + MS_DUMMY_WARNING,'자유 실습\n50분'],
         ['[블록4]\n2:50~3:40','응용: 업무 자동화 지도 작성 (Copilot 웹 활용)',
          'Copilot 웹에서 우리 부서 업무 분석 요청\n내 반복 업무 리스트업 + AI 도구 매핑표\n"어떤 업무에 어떤 AI를?" 전략 수립'
          + MS_DUMMY_WARNING,'자유 실습\n50분'],
         ['[블록5]\n3:40~4:00','정리 + 월요일 액션',
          '오늘 배운 것 3줄 정리\n월요일 액션 공유\n다음 시간 예고: "자금일보가 버튼 하나로"','정리 20분'],
     ],
     'output':'데이터 분류 기준표 + Copilot 웹 VBA/수식 실습 + 코드 검증 체크리스트',
     'monday':'Copilot 웹에서 더미 데이터로 업무 수식/VBA 1개 만들어보기'},

    {'num':2, 'title':'자금일보 자동화',
     'subtitle':'Excel VBA 하나만 깊이 — 80% 완성 템플릿 + AI가 나머지',
     'goal':'AI에게 시켜서 자금일보 VBA 매크로를 완성하고 실행한다',
     'success_30min':'버튼 하나로 자금일보가 완성됐다',
     'tools':['GitHub Copilot + Excel VBA (Copilot이 코드 생성 → 복붙)'],
     'planb':'VBA 매크로 차단 시 → Excel 수식+함수만으로 자동화 / Power Query 대안',
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 30분 성공 경험',
          '[강의 10분] 프롬프트 엔지니어링 핵심 3가지\n(역할 지정, 구체적 요청, 예시 포함)'
          + MS_SECURITY_BRIEFING +
          '\n[미니실습 40분] 80% 완성 자금일보 템플릿 열기\n→ GitHub Copilot에게 "나머지 VBA 완성해줘"\n→ 코드 복붙 → 버튼 클릭 → 자금일보 완성!'
          + MS_DUMMY_WARNING,'강의 10분\n+ 실습 40분'],
         ['[블록2]\n0:50~1:40','메인 실습: 채권연령표 자동화',
          '[시연 10분] 채권연령표 양식과 자동화 포인트\n경과일수 계산 → 연령 구간 분류 → 요약\n[따라하기 40분] AI에게 "채권연령표 VBA 만들어줘"\n→ 코드 검증(4단계) → 붙여넣기 → 실행'
          + MS_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','심화: Document AI 메인 실습',
          '[시연 10분] 업스테이지 Document AI 소개\nPDF 영수증/계약서를 AI가 읽는다 — 이것이 핵심!\n[따라하기 40분] Document Parse API로\nPDF 영수증 3개 파싱 → Excel로 자동 정리\n날짜/금액/항목 추출 → 깔끔한 표 완성'
          + MS_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['[블록4]\n2:50~3:40','응용: 본인 업무 자동화 (자유 실습)',
          '본인 업무에 맞는 VBA 요청 작성\n또는 채권연령표 코드 커스터마이즈\nAI에게 코드 받기 → 검증 → 실행\n안 되면 에러를 AI에게 보여주고 수정 요청'
          + MS_DUMMY_WARNING,'자유 실습\n50분'],
         ['[블록5]\n3:40~4:00','정리 + 월요일 액션',
          '오늘 만든 매크로 USB/메일 저장\n월요일 액션 공유\n다음: "AI가 매일 자동으로 보고서를 만든다"','정리 20분'],
     ],
     'output':'자금일보 VBA 매크로 + 채권연령표 매크로 + Document AI PDF→Excel 변환',
     'monday':'만든 자금일보 매크로를 더미 데이터 또는 등급 확인 후 사용'},

    {'num':3, 'title':'AI가 혼자 판단하게 만들기',
     'subtitle':'Copilot + VBA로 생산일보 자동화 + Copilot CLI로 PowerShell 생성 → "방금 한 게 Agent"',
     'goal':'Copilot이 만든 VBA+PowerShell로 업무를 자동화하고, AI Agent 개념을 체험으로 이해한다',
     'success_30min':'생산일보가 매일 아침 자동으로 만들어진다!',
     'tools':['GitHub Copilot + Excel VBA + 작업 스케줄러', 'PowerShell (Windows 기본 내장)', 'gh copilot CLI'],
     'planb':'VBA 차단 시 → Excel 수식 자동화 / PowerShell 단독으로 진행',
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 30분 성공 경험',
          '[강의 10분] "AI가 혼자 일하는 것" 데모\n(폴더에 파일 넣기→자동 집계→보고서→알림)'
          + MS_SECURITY_BRIEFING +
          '\n[미니실습 40분] Copilot에게\n"생산 데이터 엑셀을 읽어서 라인별 집계하는 VBA"\n요청 → 80% 완성 템플릿에 붙여넣기 → 실행\n→ 생산일보 자동 완성!'
          + MS_DUMMY_WARNING,'강의 10분\n+ 실습 40분'],
         ['[블록2]\n0:50~1:40','메인 실습: 작업 스케줄러로 매일 자동 실행',
          '[시연 10분] Windows 작업 스케줄러란?\n"매일 아침 8시에 알아서 돌아가게 하는 것"\n[따라하기 40분] 만든 VBA 매크로를\n작업 스케줄러에 등록\n→ "매일 오전 8시 자동 실행" 설정\n→ 내일 아침 출근하면 보고서가 이미 완성!','시연 10분\n+ 실습 40분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','심화: Copilot CLI + PowerShell',
          '[시연 10분] 터미널이란? "컴퓨터에게 문자로 말하는 창"\ngh copilot suggest / explain 소개\n"자연어로 명령하면 코드가 나온다"\n[따라하기 40분] gh copilot suggest:\n"OneDrive 폴더 안 파일을 날짜별로 분류"\n→ PowerShell 스크립트 자동 생성 → 실행\n→ 파일이 자동 정리됨!\n+ "이름 규칙 일괄 변경" 스크립트도 제작'
          + MS_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['[블록4]\n2:50~3:40','응용: 개념 정리 + 본인 워크플로우',
          '"방금 한 게 AI Agent입니다"\nCopilot CLI가 코드를 만들고\nVBA가 데이터를 처리하고\nPowerShell이 파일을 정리하고\n스케줄러가 매일 실행 = Agentic Workflow\n본인 부서에 적용할 자동화 흐름 설계','자유 실습\n50분'],
         ['[블록5]\n3:40~4:00','정리 + 월요일 액션',
          '만든 VBA + PowerShell 저장\nAgent = "조건에 따라 스스로 판단하는 자동화"\n다음: "대시보드를 만들고 팀원에게 공유한다"','정리 20분'],
     ],
     'output':'생산일보 VBA 자동화 + 작업 스케줄러 예약 + Copilot CLI + PowerShell 스크립트',
     'monday':'작업 스케줄러에 등록한 VBA가 내일 아침 자동 실행되는지 확인'},

    {'num':4, 'title':'대시보드 만들고 팀에 공유',
     'subtitle':'AI로 HTML 대시보드 1개 제대로 만들고 팀원이 바로 쓰게 하기',
     'goal':'AI에게 시켜서 대시보드를 만들고, 파일 공유로 팀원이 접속하게 한다',
     'success_30min':'내가 만든 대시보드를 팀원이 봤다',
     'tools':['GitHub Copilot (HTML 코드 생성)'],
     'planb':'Copilot 접속 불가 → ChatGPT 웹 / Power BI 무료 데모',
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 30분 성공 경험',
          '[강의 10분] 바이브코딩이란?\n"말로 하면 AI가 코드를 짠다" 시연'
          + MS_SECURITY_BRIEFING +
          '\n[미니실습 40분] gh copilot suggest:\n"월별 매출 차트 대시보드 HTML 만들어줘"\n→ 코드 복사 → OneDrive 폴더에 .html 저장\n→ 브라우저 열기 → "내가 만든 대시보드!"'
          + MS_DUMMY_WARNING,'강의 10분\n+ 실습 40분'],
         ['[블록2]\n0:50~1:40','메인 실습: 업무 도구 제작',
          '[시연 10분] 다양한 업무 도구 예시\n(자금현황판, 채권 계산기, 재고 모니터)\n[따라하기 40분] Copilot CLI/웹에게\n본인 업무 도구 요청 → 코드 검증\n→ OneDrive에 .html 저장'
          + MS_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','심화: OneDrive로 팀에 공유 + 접근 통제',
          '[시연 10분] OneDrive 공유의 힘:\n(1) OneDrive 공유 링크 생성 → 팀원 클릭만으로 접속\n(2) Teams 채널에 링크 게시\n(3) 공유 폴더에 배치 → 자동 동기화\n\n'
          '🔒 접근 통제 체크리스트:\n[ ] 공유 대상: 특정 사용자만 (모든 사람 = 금지)\n[ ] 권한: 보기만\n[ ] 만료일: 교육 종료+7일\n\n'
          '[따라하기 40분] 만든 대시보드를 OneDrive에 업로드\n→ 공유 링크 생성(체크리스트 확인!) → 팀원에게 전달 → 접속 확인','시연 10분\n+ 실습 40분'],
         ['[블록4]\n2:50~3:40','비즈니스 케이스 + 자유 제작',
          '공장 태블릿 → 생산 대시보드\n영업팀 일일 현황판\n회의실 KPI 보드\n본인 도구 자유 제작 + 동료 접속 테스트'
          + MS_DUMMY_WARNING,'자유 실습\n50분'],
         ['[블록5]\n3:40~4:00','정리 + 월요일 액션',
          '만든 대시보드 저장 확인\n"코드 못 해도 AI가 만들어준다"\n다음: "인터넷 없이도 AI를 쓴다"','정리 20분'],
     ],
     'output':'HTML 대시보드 + 업무 도구 + 팀 공유 완료 (접근 통제 적용)',
     'monday':'만든 대시보드를 팀 Teams/OneDrive에 공유하고 팀원 반응 확인'},

    {'num':5, 'title':'로컬 AI - 인터넷 없이도',
     'subtitle':'LM Studio GUI로 먼저 성공 → Ollama CLI는 선택',
     'goal':'로컬 AI를 설치하고, 인터넷 없이 AI를 사용하는 경험을 한다',
     'success_30min':'인터넷 끊어도 AI가 답한다!',
     'tools':['LM Studio (GUI — 클릭으로 사용)'],
     'planb':'설치 불가 → 강사 PC 화면 공유로 시연 + 수강생은 웹 AI에서 동일 실습',
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 30분 성공 경험',
          '[강의 15분] 왜 로컬 AI인가?\n1회차 데이터 분류 복습 → "3등급 데이터는 로컬만"\nLM Studio 소개: "클릭만으로 AI 챗봇"'
          + MS_SECURITY_BRIEFING +
          '\n[미니실습 35분] LM Studio 실행(사전 설치)\n→ 모델 선택(Solar Mini) → 질문\n→ "인터넷 끊어도 답한다!" 체험','강의 15분\n+ 실습 35분'],
         ['[블록2]\n0:50~1:40','메인 실습: Solar vs Gemma4 비교',
          '[시연 10분] 모델 비교 기준: 한국어 / 코드 / 속도\n[따라하기 40분] 같은 질문을 Solar와 Gemma4에게\n→ 결과 비교 → "어떤 모델이 우리 업무에 맞나?"','시연 10분\n+ 실습 40분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','심화: Ollama CLI 체험 (선택)',
          '[시연 10분] 터미널이란? "컴퓨터에게 문자로 말하는 창"\n[따라하기 40분] ollama run solar-mini\n→ 같은 질문 → LM Studio와 결과 비교\n(설치 안 되면 강사 화면 함께 보기)','시연 10분\n+ 실습 40분'],
         ['[블록4]\n2:50~3:40','응용: 로컬 AI로 문서 처리',
          '로컬 AI에게 "이 파일 요약해줘" 시키기\n파일 분류, 이름 변경 스크립트 요청\n"민감 데이터도 안전하게 AI 활용" 체험','자유 실습\n50분'],
         ['[블록5]\n3:40~4:00','보안 등급별 도구 선택 가이드',
          '1~5회차 도구 총정리\n"어떤 데이터에 어떤 AI?" 최종 매핑표\n다음: "전부 연결한다"','정리 20분'],
     ],
     'output':'로컬 AI 환경 + Solar vs Gemma 비교 + 보안 등급별 도구 매핑표',
     'monday':'LM Studio에서 업무 질문 3개 해보고 결과 정리'},

    {'num':6, 'title':'전부 연결하기 - 종합 프로젝트',
     'subtitle':'1~5회차를 조합하여 본인 업무 end-to-end 자동화 파이프라인 완성',
     'goal':'배운 도구를 조합해 본인 업무의 전체 흐름을 자동화한다',
     'success_30min':'내 업무가 자동화됐다!',
     'tools':['1~5회차 전체 도구 (선택 조합)'],
     'planb':None,
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 전체 파이프라인 설계',
          '[강의 10분] 전 과정 복습 (5분 요약)\n프로젝트: 전체를 하나로 연결한다'
          + MS_SECURITY_BRIEFING +
          '\n[미니실습 40분] OneDrive 연동 실습\nOneDrive → Windows 탐색기 마운트\n자동 동기화 설정 + "OneDrive 폴더를 자동화 대상으로" 매핑\ngh copilot suggest로 OneDrive 폴더 구조 자동 생성 스크립트\n→ 프로젝트 폴더 체계 완성'
          + MS_DUMMY_WARNING,'강의 10분\n+ 실습 40분'],
         ['[블록2]\n0:50~1:40','페어 프로젝트 시작',
          '2인 1조 구성\n본인 업무 파이프라인 설계:\nOneDrive 폴더 감시 → Copilot CLI로 스크립트\n→ VBA 보고서 → OneDrive 자동 저장\n→ 공유 링크로 팀 공유','페어 실습\n50분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','프로젝트 구현',
          '설계한 파이프라인을 실제로 구축\nAI에게 코드 요청 → 검증 → 실행\n막히면 페어 파트너 + AI에게 도움 요청','페어 실습\n50분'],
         ['[블록4]\n2:50~3:40','프로젝트 완성 + 테스트',
          '파이프라인 테스트 실행\n결과물 대시보드 확인\n발표 준비 (2분 시연)','실습\n50분'],
         ['[블록5]\n3:40~4:00','발표 + 수료 + 데이터 정리',
          '조별 2분 시연 발표\nAI Agent 시대의 업무 역량 정리\n"AI에게 시키는 4단계" 최종 복습\n\n'
          '🧹 교육 데이터 정리 체크리스트:\n[ ] Copilot 대화 삭제\n[ ] OneDrive 실습파일 삭제 + 휴지통 비우기 (93일 보관 주의!)\n[ ] VBA 매크로 정리 (불필요한 매크로 삭제)\n[ ] 작업 스케줄러 비활성화\n\n수료','발표 20분'],
     ],
     'output':'본인 업무 end-to-end AI 자동화 파이프라인',
     'monday':'완성한 자동화를 실제 업무에서 1주일간 사용해보기'},
]

# ============================================================
# CLAUDE TRACK
# ============================================================
# Claude 보안 브리핑 문구 (매 회차 블록1에 삽입)
CL_SECURITY_BRIEFING = '\n[보안] Claude Free/Pro는 학습 사용. Teams 플랜만 안전! 더미 데이터만 사용하세요.'
# Claude 더미 데이터 경고
CL_DUMMY_WARNING = '\n⚠️ Teams 플랜만 안전! Free/Pro는 학습 사용. 더미 데이터만 사용'

claude_sessions = [
    {'num':1, 'title':'AI 첫 만남 + 보안',
     'subtitle':'claude.ai 웹에서 바로 체험 + 200K 컨텍스트 위력 + 데이터 분류',
     'goal':'Claude의 강점(긴 문서 분석, 정직한 답변)을 체험하고 보안 기준을 수립한다',
     'success_30min':'AI가 20페이지 계약서를 1분에 요약했다!',
     'tools':['claude.ai 웹 (Pro Plan 이상)'],
     'planb':'Claude 접속 불가 → ChatGPT / Gemini 웹',
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 30분 성공 경험',
          '[강의 15분] AI란 무엇인가 + Claude 차별점\n"가장 긴 문서를 읽는 AI" (200K 토큰)\n"모르면 모른다고 말하는 AI" (환각 최소화)'
          + CL_SECURITY_BRIEFING +
          '\n[미니실습 35분] claude.ai 접속\n→ 20페이지 PDF 업로드 → "핵심 3줄 요약해줘"\n→ 1분 만에 완성 → "이게 된다고?"'
          + CL_DUMMY_WARNING,'강의 15분\n+ 실습 35분'],
         ['[블록2]\n0:50~1:40','메인 실습: 데이터 분류 기준표',
          '[시연 10분] 데이터 4등급 분류\nClaude Teams Plan 보안 정책 설명\n외부 파일 프롬프트 인젝션 주의사항\n[따라하기 40분] 80% 완성 템플릿에\n우리 회사 데이터 분류 + Claude로 검증'
          + CL_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','심화: AI 코드 검증 4단계',
          '[시연 10분] Claude에게 코드를 받았을 때\n검증 체크리스트 (외부URL, API키, 삭제명령)\n[따라하기 40분] Claude에게 엑셀 수식 요청\n→ 받은 코드 검증 실습\n→ Claude 스스로 "이 코드 안전해?" 확인'
          + CL_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['[블록4]\n2:50~3:40','응용: 업무 자동화 지도',
          '반복 업무 리스트업 + AI 매핑표 작성\nClaude에게 업무 분석 요청\n"어떤 업무에 어떤 AI를?" 전략 수립'
          + CL_DUMMY_WARNING,'자유 실습\n50분'],
         ['[블록5]\n3:40~4:00','정리 + 월요일 액션',
          'Claude = "가장 긴 문서를 가장 정직하게 읽는 AI"\n다음: "자금일보를 Claude가 만들어준다"','정리 20분'],
     ],
     'output':'데이터 분류 기준표 + 코드 검증 체크리스트 + 업무 자동화 지도',
     'monday':'더미 데이터 또는 등급 확인 후 긴 문서(보고서/계약서) 1개를 Claude에게 요약 시켜보기'},

    {'num':2, 'title':'자금일보 자동화 + Artifacts',
     'subtitle':'Claude Artifacts로 수식과 코드를 즉시 생성하고 실행까지',
     'goal':'Artifacts로 코드를 바로 실행해보고, 자금일보를 자동화한다',
     'success_30min':'Claude가 만든 수식으로 자금일보가 완성됐다',
     'tools':['claude.ai Artifacts (코드/수식 즉시 실행)'],
     'planb':'Artifacts 미지원 → Claude에서 코드 생성 후 복붙',
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 30분 성공 경험',
          '[강의 10분] Artifacts란?\n"Claude가 만든 코드가 바로 실행된다"\n프롬프트 엔지니어링 핵심 3가지'
          + CL_SECURITY_BRIEFING +
          '\n[미니실습 40분] 80% 완성 자금일보 데이터 준비\n→ Claude에게 "자금일보 수식 만들어줘"\n→ Artifacts에서 바로 결과 확인!'
          + CL_DUMMY_WARNING,'강의 10분\n+ 실습 40분'],
         ['[블록2]\n0:50~1:40','메인 실습: 채권연령표 자동화',
          '[시연 10분] 채권연령표 양식 + Artifacts 활용법\n[따라하기 40분] Claude에게 경과일수 분류 요청\n→ Artifacts에서 코드 실행\n→ 결과 다운로드 → Excel에 적용'
          + CL_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','심화: Document AI 메인 실습',
          '[시연 10분] 업스테이지 Document AI 소개\nPDF 영수증/계약서를 AI가 읽는다 — 이것이 핵심!\n[따라하기 40분] Document Parse API로\nPDF 영수증 3개 파싱 → Claude에게 결과 정리 요청\n→ 날짜/금액/항목 추출 → Excel 스프레드시트 완성'
          + CL_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['[블록4]\n2:50~3:40','응용: 본인 업무 자동화 (자유 실습)',
          'Claude Artifacts로 본인 업무 수식/코드 생성\n또는 채권연령표 코드 커스터마이즈\n안 되면 에러를 Claude에게 보여주고 수정\n"Artifacts = 나만의 코딩 놀이터"'
          + CL_DUMMY_WARNING,'자유 실습\n50분'],
         ['[블록5]\n3:40~4:00','정리 + 월요일 액션',
          'Artifacts 활용 팁 정리\n만든 코드 저장\n다음: "Claude가 내 폴더를 직접 읽는다"','정리 20분'],
     ],
     'output':'자금일보 수식/코드 + 채권연령표 자동화 + Document AI PDF→Excel 변환',
     'monday':'더미 데이터 또는 등급 확인 후 Artifacts에서 업무 수식 1개를 더 만들어보기'},

    {'num':3, 'title':'AI가 혼자 판단하게 만들기',
     'subtitle':'Claude Desktop + MCP로 파일을 직접 읽고 정리 → "방금 한 게 Agent"',
     'goal':'Claude가 내 파일시스템에 직접 접근하여 작업하는 체험 + Agent 개념 이해',
     'success_30min':'Claude가 내 폴더를 직접 읽고 정리했다!',
     'tools':['Claude Desktop + MCP (파일시스템 연결)'],
     'planb':'Claude Desktop 설치 불가 → claude.ai 웹에서 파일 업로드 방식으로 동일 워크플로우',
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 30분 성공 경험',
          '[강의 10분] "AI가 내 컴퓨터를 직접 본다" 데모\n(Claude Desktop이 폴더를 읽고 파일 정리)'
          + CL_SECURITY_BRIEFING +
          '\n⚠️ MCP 연결 폴더는 연습용만! 파일 내용이 Anthropic 서버로 전송됩니다'
          '\n[미니실습 40분] Claude Desktop 실행(사전설치)\n→ MCP 파일시스템 연결\n→ "이 폴더 파일을 날짜별로 정리해줘"\n→ Claude가 직접 실행!'
          + CL_DUMMY_WARNING,'강의 10분\n+ 실습 40분'],
         ['[블록2]\n0:50~1:40','메인 실습: 생산일보 Agent',
          '[시연 10분] Claude가 여러 도구를 연결하는 원리\n[따라하기 40분] Claude Desktop + MCP로\n데이터 파일 읽기 → 집계 → 보고서 생성\n→ "AI가 혼자서 다 했다!"'
          + CL_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','심화: 개념 정리 + 워크플로우 설계',
          '"방금 한 게 AI Agent입니다"\nAgent, MCP, Agentic Workflow 이름 붙이기\n체험 후 개념 정리 → 훨씬 쉽게 이해!\n본인 부서에 적용할 워크플로우 설계','강의 20분\n+ 실습 30분'],
         ['[블록4]\n2:50~3:40','응용: 로컬 AI와 연계',
          'Ollama(Solar) 맛보기 시연\n"민감 데이터는 로컬, 복잡 분석은 Claude"\n하이브리드 아키텍처 개념','자유 실습\n50분'],
         ['[블록5]\n3:40~4:00','정리 + 월요일 액션',
          'Agent = "AI가 도구를 써서 스스로 일하는 것"\n다음: "대화로 웹앱을 만든다"','정리 20분'],
     ],
     'output':'Claude Desktop MCP 설정 + Agent 워크플로우 체험 + 개념 이해',
     'monday':'더미 데이터 또는 등급 확인 후 Claude Desktop에서 파일 정리 1가지를 시켜보기'},

    {'num':4, 'title':'대시보드 만들고 팀에 공유',
     'subtitle':'Claude Artifacts = 킬러 콘텐츠. 대화로 웹앱을 만들고 공유한다',
     'goal':'Artifacts로 인터랙티브 대시보드/웹앱을 만들고 팀에 공유한다',
     'success_30min':'대화로 만든 대시보드가 바로 동작한다!',
     'tools':['Claude Artifacts (HTML/React 즉시 실행)'],
     'planb':'Artifacts 미지원 → Claude에서 HTML 생성 후 파일로 저장/공유',
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 30분 성공 경험',
          '[강의 10분] 바이브코딩이란?\nArtifacts 시연: "말하면 앱이 만들어진다"'
          + CL_SECURITY_BRIEFING +
          '\n[미니실습 40분] Claude에게\n"매출 차트 대시보드 만들어줘"\n→ Artifacts에서 바로 실행!\n→ 필터/차트가 동작하는 인터랙티브 대시보드'
          + CL_DUMMY_WARNING,'강의 10분\n+ 실습 40분'],
         ['[블록2]\n0:50~1:40','메인 실습: 업무 도구 제작',
          '[시연 10분] 다양한 Artifacts 예시\n(계산기, 퀴즈, 폼, 대시보드)\n[따라하기 40분] 본인 업무 도구 제작:\n"자금일보 입력 → PDF 출력 앱"\n"채권 연령 계산기" 등'
          + CL_DUMMY_WARNING,'시연 10분\n+ 실습 40분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','심화: 팀에 공유하기 + 접근 통제',
          '⚠️ Publish 버튼 = 전세계 공개! 절대 금지. HTML 다운로드 또는 링크만\n\n'
          '[시연 10분] Artifacts 공유 방법:\n(1) Artifacts 링크 공유 (가장 쉬움)\n(2) HTML 다운로드 → 파일 공유\n(3) OneDrive/Google Drive 업로드\n\n'
          '🔒 접근 통제 체크리스트:\n[ ] 공유 대상: 특정 사용자만 (모든 사람 = 금지)\n[ ] 권한: 보기만\n[ ] 만료일: 교육 종료+7일\n\n'
          '[따라하기 40분] 실제 공유 + 팀원 접속 테스트','시연 10분\n+ 실습 40분'],
         ['[블록4]\n2:50~3:40','비즈니스 케이스 + 자유 제작',
          '공장 모니터 / 영업 현황판 / KPI 보드\n본인 도구 자유 제작\nArtifacts = Copilot/Gemini에 없는 킬러 기능'
          + CL_DUMMY_WARNING,'자유 실습\n50분'],
         ['[블록5]\n3:40~4:00','정리 + 월요일 액션',
          '"코드 못 해도 웹앱을 만든다 = Artifacts"\n만든 도구 저장\n다음: "인터넷 없이도 AI를 쓴다"','정리 20분'],
     ],
     'output':'Artifacts 대시보드 + 업무 도구 + 팀 공유 완료 (접근 통제 적용)',
     'monday':'더미 데이터 또는 등급 확인 후 Artifacts로 업무 도구 1개를 더 만들고 팀에 공유'},

    {'num':5, 'title':'로컬 AI - 인터넷 없이도',
     'subtitle':'LM Studio GUI 먼저 → Ollama CLI 선택 + 하이브리드 아키텍처',
     'goal':'로컬 AI 환경 구축 + "민감 데이터는 로컬, 복잡 분석은 Claude" 하이브리드',
     'success_30min':'인터넷 끊어도 AI가 답한다!',
     'tools':['LM Studio (GUI)'],
     'planb':'설치 불가 → 강사 PC 시연 + 수강생은 claude.ai에서 동일 실습',
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + 30분 성공 경험',
          '[강의 15분] 왜 로컬 AI인가?\n"3등급 데이터는 로컬만" (1회차 복습)\n하이브리드: 민감→로컬, 복잡→Claude'
          + CL_SECURITY_BRIEFING +
          '\n[미니실습 35분] LM Studio 실행\n→ Solar Mini 선택 → 질문 → 답변!','강의 15분\n+ 실습 35분'],
         ['[블록2]\n0:50~1:40','메인 실습: Solar vs Gemma4 비교',
          '[시연 10분] 비교 기준: 한국어 / 코드 / 속도\n[따라하기 40분] 같은 질문 비교\n→ "우리 업무에는 어떤 모델?"','시연 10분\n+ 실습 40분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','심화: 하이브리드 실습',
          '[시연 10분] 하이브리드 아키텍처 설계\n[따라하기 40분] 민감 질문은 LM Studio(로컬)\n복잡 분석은 Claude(클라우드)\n→ 결과 비교 → 최적 조합 찾기','시연 10분\n+ 실습 40분'],
         ['[블록4]\n2:50~3:40','응용: Ollama CLI 체험 (선택)',
          'ollama run solar-mini 실행\n터미널 체험 (선택사항)\nClaude Desktop MCP + 로컬 모델 연계 가능성','자유 실습\n50분'],
         ['[블록5]\n3:40~4:00','도구 총정리',
          '1~5회차 도구 전체 매핑표\n보안 등급별 최적 도구 조합\n다음: "전부 연결한다"','정리 20분'],
     ],
     'output':'로컬 AI 환경 + Solar vs Gemma + 하이브리드 아키텍처',
     'monday':'민감 업무 질문 1개를 로컬 AI로, 복잡 분석 1개를 Claude로'},

    {'num':6, 'title':'전부 연결하기 + 부서 AI 어시스턴트',
     'subtitle':'Claude Projects로 우리 부서 전용 AI를 만들고, 전체 파이프라인 완성',
     'goal':'Claude Projects로 지속 사용 가능한 부서 AI를 만들고 전체 자동화를 완성한다',
     'success_30min':'우리 부서 전용 AI 어시스턴트가 완성됐다!',
     'tools':['Claude Projects + 1~5회차 전체 도구'],
     'planb':None,
     'rows':[
         ['[블록1]\n0:00~0:50','도입 + Claude Projects',
          '[강의 10분] Claude Projects란?\n"부서 문서를 AI에게 가르치면 전문가가 된다"'
          + CL_SECURITY_BRIEFING +
          '\n[미니실습 40분] Projects 생성\n→ 부서 규정/양식/FAQ 업로드\n→ "이 규정에서 출장비 한도가?" → 즉시 답변!'
          + CL_DUMMY_WARNING,'강의 10분\n+ 실습 40분'],
         ['[블록2]\n0:50~1:40','페어 프로젝트 시작',
          '2인 1조 구성\n본인 업무 파이프라인 설계:\n데이터 수집 → Claude 분석 → Artifacts 대시보드\n→ 클라우드 저장 → 공유','페어 실습\n50분'],
         ['','휴식','1:40~2:00 (20분)',''],
         ['[블록3]\n2:00~2:50','프로젝트 구현',
          '설계한 파이프라인 구축\nClaude Projects + Artifacts + MCP 활용\n막히면 페어 파트너 + Claude에게 도움','페어 실습\n50분'],
         ['[블록4]\n2:50~3:40','프로젝트 완성 + 테스트',
          '파이프라인 테스트 실행\nProjects에 Custom Instructions 설정\n발표 준비 (2분 시연)','실습\n50분'],
         ['[블록5]\n3:40~4:00','발표 + 수료 + 데이터 정리',
          '조별 2분 시연 발표\n"AI에게 시키는 4단계" 최종 복습\nProjects = 교육 끝나도 계속 쓸 수 있는 AI\n\n'
          '🧹 교육 데이터 정리 체크리스트:\n[ ] Claude 대화 삭제\n[ ] Artifacts 삭제\n[ ] Projects 문서 삭제\n[ ] MCP config 제거\n[ ] LM Studio 히스토리 정리\n\n수료','발표 20분'],
     ],
     'output':'Claude Projects 부서 AI 어시스턴트 + end-to-end 자동화 파이프라인',
     'monday':'만든 Projects를 팀원 1명에게 공유하고 사용법 알려주기'},
]

# ============================================================
# GENERATE
# ============================================================
build_track('Microsoft', '2B579A',
    'GitHub Copilot · Excel VBA · PowerShell · 작업 스케줄러 · OneDrive',
    'GitHub Copilot (github.com/copilot)',
    'Excel VBA, PowerShell (Windows 내장), 작업 스케줄러, Power Query, gh copilot CLI',
    ms_sessions,
    r'C:\Users\moonk\OneDrive\AI강사준비\수업계획서_Microsoft_트랙.docx')

build_track('Claude', 'D97706',
    'Claude.ai · Artifacts · Claude Desktop + MCP · Projects',
    'Claude (claude.ai, Pro/Teams Plan)',
    'Artifacts, Claude Desktop, MCP, Projects',
    claude_sessions,
    r'C:\Users\moonk\OneDrive\AI강사준비\수업계획서_Claude_트랙.docx')

print('Done! MS + Claude tracks created.')
