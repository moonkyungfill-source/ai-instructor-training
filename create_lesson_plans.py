from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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

def make_table(doc, headers, rows, header_color='2B579A'):
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


# ============================================================
# TRACK DATA
# ============================================================

tracks = {
    'Google': {
        'color': '0F9D58',
        'subtitle': 'Google AI Studio · Apps Script · Gemini · Google Labs (Opal·Mixboard·Pomely·Flow)',
        'ai_tool': 'Google AI Studio / Gemini',
        'extra_tools': 'Google Labs — Opal(미니앱), Mixboard(컨셉보드), Pomely(마케팅), Flow(영상)',
        'code_tool': 'Google Apps Script',
        'spreadsheet': 'Google Sheets',
        'doc_tool': 'Google Docs',
        'cloud': 'Google Drive',
        'dashboard': 'Looker Studio',
        'sessions': [
            {   # 1회차
                'title': '사내에서 AI를 안전하게 쓰는 법',
                'goal': 'AI 보안 리터러시 확보 + Google 트랙 환경 세팅',
                'rows': [
                    ['0:00-0:50', '왜 ChatGPT가 막혔을까?', '삼성 소스코드 유출 사건 분석\n기업 데이터 유출 경로 이해\nAI에 넣으면 안 되는 것 vs 되는 것', '강의 + 토론'],
                    ['0:50-1:30', '실습: 데이터 분류 기준표', '우리 회사 데이터를 4등급 분류\n(1)자유 (2)클라우드OK (3)로컬만 (4)절대불가', '개인 실습'],
                    ['1:30-2:10', 'AI 도구 보안 지도', 'Google AI Studio: 무료 vs 유료 학습 정책 차이\nOllama: 로컬 AI의 진짜 보안 범위\n도구별 정확한 보안 조건', '강의'],
                    ['2:10-3:00', '실습: AI 도구 세팅', 'Google AI Studio 접속 + Billing 계정 확인\nGoogle Labs 가입 (Opal, Mixboard, Pomely, Flow)\nOllama 설치 → 보안 설정 → Solar/Gemma4 실행', '실습 (트러블슈팅 포함)'],
                    ['3:00-3:40', '실습: AI 코드 보안 검증', 'AI에게 코드 요청 → 외부 URL/API 키/삭제 명령 체크\nUrlFetchApp 등 외부 호출 확인법\n"이 코드 안전해?" AI에게 재확인', '실습'],
                    ['3:40-4:00', '업무 자동화 지도 작성', '반복 업무 리스트업\n데이터 등급별 AI 도구 매핑', '개인 작업'],
                ],
                'output': '데이터 분류 기준표 + Google AI Studio/Ollama 세팅 완료 + 코드 검증 체크리스트',
            },
            {   # 2회차
                'title': 'AI에게 시키기 + 스프레드시트 자동화',
                'goal': '프롬프트 엔지니어링 습득 + 자금일보/채권연령표 자동화',
                'rows': [
                    ['0:00-0:30', '프롬프트 엔지니어링', '같은 요청을 3가지로 → 결과 비교\nAI Studio vs Ollama(Solar) 비교\n"잘 시키면 잘 나온다" 체감', '실습'],
                    ['0:30-1:30', '실습: 자금일보 자동화', 'AI Studio에게 "은행 거래내역 CSV를 읽어서\n자금일보 양식에 채우는 Apps Script 만들어줘"\n→ Sheets에 적용 → 실행\n안 되면 Ollama(Solar)로 동일 요청', '실습'],
                    ['1:30-2:20', '실습: 채권연령표 자동화', 'AI에게 "매출채권 경과일수 계산 + 연령 구간 분류\nApps Script 만들어줘"\n→ 클라우드 AI / 로컬 AI 결과 비교', '실습'],
                    ['2:20-3:10', '실습: Document AI 문서 파싱', '업스테이지 Document Parse API로\nPDF 영수증·계약서 자동 파싱\n→ Google Sheets로 자동 정리', '실습'],
                    ['3:10-4:00', 'AI 코드 보안 검증 실습', 'AI가 만든 Apps Script 코드에서\nUrlFetchApp 외부 호출 여부 확인\n수정 요청 → 안전한 코드로 변환', '실습 + 정리'],
                ],
                'output': '자금일보/채권연령표 자동 생성 Apps Script + Document AI 체험',
            },
            {   # 3회차
                'title': 'AI Agent 워크플로우 & 보고서 자동화',
                'goal': 'AI Agent 개념 이해 + Opal Agent 워크플로우 + 보고서 파이프라인',
                'rows': [
                    ['0:00-0:40', 'AI Agent 개념', 'AI Agent란? — 도구를 연결하는 AI\nMCP(Model Context Protocol) 소개\nAgentic Workflow 개념 + Google Opal Agent 소개', '강의 + 시연'],
                    ['0:40-1:40', '실습: 생산일보 Agent', 'AI Studio에게 Apps Script 요청\n→ 트리거로 매일 자동 실행\n+ Opal로 생산일보 Agent 미니앱 제작\n→ Ollama(Solar)로 로컬 대안', '실습'],
                    ['1:40-2:40', '실습: Opal Agentic 워크플로우', 'Opal Agent Step 활용:\nGoogle Forms 입력 → AI가 데이터 검증\n→ Sheets 자동 기록 → 조건별 이메일 발송\n(노코드로 Agent 워크플로우 구축)', '실습'],
                    ['2:40-3:20', '실습: 로컬 AI Agent', 'Ollama(Solar/Gemma4)로\n동일 워크플로우를 오프라인 구현\n→ 클라우드 vs 로컬 비교', '실습'],
                    ['3:20-4:00', 'AI Agent 한계와 주의점', '오류 대응 + Agent 검증 포인트\n사람이 확인해야 하는 지점', '강의 + Q&A'],
                ],
                'output': 'Opal Agent 워크플로우 + Apps Script 자동 보고서 (클라우드 + 로컬)',
            },
            {   # 4회차
                'title': '바이브코딩 + Google Labs: 만들고, 공유하고, 쓰게 하기',
                'goal': 'Opal 미니앱 + HTML 대시보드 + Mixboard 시각화 + 로컬 공유',
                'rows': [
                    ['0:00-0:30', '바이브코딩 + Google Labs 소개', '바이브코딩 개념\nGoogle Labs 도구 소개:\n- Opal: 자연어로 미니앱 제작\n- Mixboard: AI 컨셉보드\n- Pomely: 브랜드 마케팅 자동화\n- Flow: AI 영상 제작', '강의 + 시연'],
                    ['0:30-1:20', '실습: Opal로 업무 미니앱 제작', 'Opal에서 "자금일보 입력 → 자동 계산 →\nPDF 출력하는 앱 만들어줘" 입력\n→ 5분 만에 미니앱 완성\n+ "채권 연령 계산기" 앱 제작', '실습'],
                    ['1:20-2:00', '실습: AI Studio HTML 대시보드', 'AI Studio에게 "생산 실적 차트 대시보드\nHTML 만들어줘" → Sheets 데이터 연동\n+ Mixboard로 대시보드 디자인 구상', '실습'],
                    ['2:00-2:30', '실습: Pomely 마케팅 활용', 'Pomely로 브랜드 DNA 분석\n→ 보고서 표지/프레젠테이션 자료 자동 생성\n(마케팅/기획 부서 활용)', '실습'],
                    ['2:30-3:10', '실습: 로컬 네트워크 공유', 'python -m http.server 실행\n→ 같은 와이파이 팀원이 접속\n+ Opal 앱은 링크 공유로 바로 접속\n→ 클라우드 없이 팀 공유 완성', '실습'],
                    ['3:10-4:00', '비즈니스 케이스 + 자유 제작', '공장 태블릿 / 영업 현황판 / KPI 보드\n본인 도구 자유 제작 + 상호 접속', '실습 + 토론'],
                ],
                'output': 'Opal 미니앱 + HTML 대시보드 + 로컬 네트워크 공유',
            },
            {   # 5회차
                'title': '로컬 AI 심화: Solar LLM vs Gemma 4',
                'goal': '모델 비교 + 오프라인 실무 파이프라인 완성',
                'rows': [
                    ['0:00-0:40', '로컬 AI 모델 비교', 'Upstage Solar Mini vs Gemma 4 vs Llama\n한국어 성능 / 코드 생성 / 문서 처리\n직접 비교 실습', '실습'],
                    ['0:40-1:30', '실습: Solar로 문서 일괄 처리', '폴더 안 파일 100개를\nSolar가 자동 분류·이름 변경·요약\n(완전 오프라인)', '실습'],
                    ['1:30-2:20', '실습: 로컬 AI + 엑셀', '"이 폴더 CSV → 월간 보고서"\n자동화 스크립트 → 작업 스케줄러 연동', '실습'],
                    ['2:20-3:10', '실습: LM Studio GUI', '클릭만으로 AI 챗봇 실행\n비개발자도 사용 가능한 로컬 AI', '실습'],
                    ['3:10-4:00', '보안 등급별 AI 선택 가이드', '도구별 정확한 보안 조건 총정리\n우리 회사에 맞는 조합 설계', '강의 + 개인 작업'],
                ],
                'output': 'Solar vs Gemma 비교 체험 + 로컬 AI 실무 파이프라인',
            },
            {   # 6회차
                'title': '전부 연결하기: 종합 프로젝트',
                'goal': 'end-to-end AI 자동화 파이프라인 완성',
                'rows': [
                    ['0:00-0:40', '클라우드 드라이브 PC 연동', 'Google Drive → Windows 탐색기 마운트\n자동 동기화 설정', '실습'],
                    ['0:40-1:20', 'AI로 파일 자동 정리', 'AI에게 "폴더 구조 설계 → 파일 분류\n→ 이름 규칙" 스크립트 요청', '실습'],
                    ['1:20-3:20', '종합 프로젝트', '각자의 업무 파이프라인 완성:\n데이터 수집 → Document AI 파싱\n→ AI Agent가 정리 → 보고서 생성\n→ Drive 저장 + 대시보드 팀 공유\n→ 이메일 자동 발송', '프로젝트'],
                    ['3:20-4:00', '발표 및 수료', '자동화 결과 공유\nAI Agent 시대의 업무 역량 정리\n수료', '발표'],
                ],
                'output': '본인 업무의 end-to-end AI Agent 자동화 파이프라인',
            },
        ],
    },
    'Microsoft': {
        'color': '2B579A',
        'subtitle': 'GitHub Copilot CLI · Excel VBA · Power Query · OneDrive',
        'ai_tool': 'GitHub Copilot CLI',
        'extra_tools': 'Power Query, Outlook VBA, Windows 작업 스케줄러',
        'code_tool': 'Excel VBA + Power Query',
        'spreadsheet': 'Microsoft Excel',
        'doc_tool': 'Microsoft Word',
        'cloud': 'OneDrive',
        'dashboard': 'Power BI',
        'sessions': [
            {   # 1회차
                'title': '사내에서 AI를 안전하게 쓰는 법',
                'goal': 'AI 보안 리터러시 확보 + Microsoft 트랙 환경 세팅',
                'rows': [
                    ['0:00-0:50', '왜 ChatGPT가 막혔을까?', '삼성 소스코드 유출 사건 분석\n기업 데이터 유출 경로 이해\nAI에 넣으면 안 되는 것 vs 되는 것', '강의 + 토론'],
                    ['0:50-1:30', '실습: 데이터 분류 기준표', '우리 회사 데이터를 4등급 분류\n(1)자유 (2)클라우드OK (3)로컬만 (4)절대불가', '개인 실습'],
                    ['1:30-2:10', 'AI 도구 보안 지도', 'Copilot CLI: Business vs Free 학습 정책 차이\nOllama: 로컬 AI의 진짜 보안 범위\n도구별 정확한 보안 조건', '강의'],
                    ['2:10-3:00', '실습: AI 도구 세팅', 'GitHub Copilot CLI 설치 + Business 플랜 확인\ngh copilot suggest / explain 첫 실행\nOllama 설치 → 보안 설정 → Gemma4 실행', '실습'],
                    ['3:00-3:40', '실습: AI 코드 보안 검증', 'AI에게 VBA 코드 요청\n→ XMLHTTP 외부 호출 여부 확인\n→ "이 코드 안전해?" AI에게 재확인', '실습'],
                    ['3:40-4:00', '업무 자동화 지도 작성', '반복 업무 리스트업\n데이터 등급별 AI 도구 매핑', '개인 작업'],
                ],
                'output': '데이터 분류 기준표 + Copilot CLI/Ollama 세팅 완료 + 코드 검증 체크리스트',
            },
            {   # 2회차
                'title': 'AI에게 시키기 + 스프레드시트 자동화',
                'goal': '프롬프트 엔지니어링 습득 + 자금일보/채권연령표 자동화',
                'rows': [
                    ['0:00-0:30', '프롬프트 엔지니어링', '같은 요청을 3가지로 → 결과 비교\nCopilot CLI vs Ollama(Gemma4) 비교', '실습'],
                    ['0:30-1:30', '실습: 자금일보 자동화', 'Copilot CLI에게 "은행 거래내역 CSV를\n자금일보 양식에 채우는 VBA 만들어줘"\n→ Excel에 적용 → 실행\n안 되면 Ollama(Gemma4)로 동일 요청', '실습'],
                    ['1:30-2:20', '실습: 채권연령표 자동화', '"매출채권 경과일수 계산 + 연령 분류\nVBA 매크로 만들어줘"\n+ Power Query로 다중 파일 병합', '실습'],
                    ['2:20-3:10', '실습: Document AI 문서 파싱', '업스테이지 Document Parse API로\nPDF 영수증·계약서 자동 파싱\n→ Excel로 자동 정리', '실습'],
                    ['3:10-4:00', 'AI 코드 보안 검증 실습', 'VBA 코드에서 XMLHTTP 외부 호출 확인\ngh copilot explain으로 코드 이해\n수정 요청 → 안전한 코드로 변환', '실습 + 정리'],
                ],
                'output': '자금일보/채권연령표 VBA 매크로 + Document AI 체험',
            },
            {   # 3회차
                'title': 'AI Agent 워크플로우 & 보고서 자동화',
                'goal': 'AI Agent 개념 이해 + Agentic 자동 보고서 파이프라인 구축',
                'rows': [
                    ['0:00-0:40', 'AI Agent 개념', 'AI Agent란? — 도구를 연결하는 AI\nMCP(Model Context Protocol) 소개\nAgentic Workflow 비즈니스 적용', '강의 + 시연'],
                    ['0:40-1:40', '실습: 생산일보 Agent', 'Copilot CLI에게 "일일 생산 실적 집계\nVBA 매크로 만들어줘" 요청\n→ Windows 작업 스케줄러로 매일 자동\n→ Ollama(Gemma4)로 로컬 대안', '실습'],
                    ['1:40-2:40', '실습: Agentic 워크플로우', '폴더 감시 → 새 파일 감지\n→ AI가 데이터 분류·정리\n→ 보고서 자동 생성\n→ Outlook VBA로 메일 자동 발송', '실습'],
                    ['2:40-3:20', '실습: 로컬 AI Agent', 'Ollama(Gemma4)로\n동일한 워크플로우를 오프라인에서 구현', '실습'],
                    ['3:20-4:00', 'AI Agent 한계와 주의점', '오류 대응 + Agent 검증 포인트\n사람이 확인해야 하는 지점', '강의 + Q&A'],
                ],
                'output': 'AI Agent 기반 자동 보고서 파이프라인 (Copilot + Ollama)',
            },
            {   # 4회차
                'title': '바이브코딩: 만들고, 공유하고, 쓰게 하기',
                'goal': 'HTML 대시보드 제작 + 로컬 네트워크 팀 공유',
                'rows': [
                    ['0:00-0:30', '바이브코딩 소개', '비즈니스 활용 사례 시연', '강의 + 시연'],
                    ['0:30-1:20', '실습: HTML 대시보드', 'Copilot에게 "생산 실적 차트 대시보드\nHTML 만들어줘" → Excel 데이터 연동', '실습'],
                    ['1:20-2:00', '실습: 업무 도구 제작', '"자금일보 입력폼 웹앱"\n"채권 연령 계산기" 등\n+ Excel 템플릿+매크로 일체형', '실습'],
                    ['2:00-2:40', '실습: 로컬 네트워크 공유', 'python -m http.server 실행\n→ 팀원이 브라우저로 접속\n→ 클라우드 없이 팀 공유', '실습'],
                    ['2:40-3:20', '비즈니스 활용 케이스', '공장 태블릿 / 영업 현황판 / KPI 보드', '토론 + 실습'],
                    ['3:20-4:00', '자유 제작 + 상호 접속', '본인 도구 제작 + 팀원끼리 접속', '실습'],
                ],
                'output': 'HTML 대시보드 + 로컬 네트워크 공유',
            },
            {   # 5회차
                'title': '로컬 AI 심화: Solar LLM vs Gemma 4',
                'goal': '모델 비교 + 오프라인 실무 파이프라인 완성',
                'rows': [
                    ['0:00-0:40', '로컬 AI 모델 비교', 'Solar Mini vs Gemma 4 vs Llama\n한국어/코드/문서 성능 직접 비교', '실습'],
                    ['0:40-1:30', '실습: Solar로 문서 일괄 처리', '파일 100개 자동 분류·이름 변경·요약', '실습'],
                    ['1:30-2:20', '실습: 로컬 AI + 엑셀', 'CSV → 월간 보고서 자동화\n→ 작업 스케줄러 연동', '실습'],
                    ['2:20-3:10', '실습: LM Studio GUI', '클릭만으로 AI 챗봇 실행', '실습'],
                    ['3:10-4:00', '보안 등급별 AI 선택 가이드', '도구별 보안 조건 총정리', '강의 + 개인 작업'],
                ],
                'output': 'Solar vs Gemma 비교 + 로컬 AI 실무 파이프라인',
            },
            {   # 6회차
                'title': '전부 연결하기: 종합 프로젝트',
                'goal': 'end-to-end AI 자동화 파이프라인 완성',
                'rows': [
                    ['0:00-0:40', '클라우드 드라이브 PC 연동', 'OneDrive → Windows 탐색기 마운트\n자동 동기화 설정', '실습'],
                    ['0:40-1:20', 'AI로 파일 자동 정리', 'Copilot에게 파일 정리 스크립트 요청', '실습'],
                    ['1:20-3:20', '종합 프로젝트', '데이터 수집 → Document AI 파싱\n→ AI Agent 정리 → VBA 보고서\n→ OneDrive 저장 + 대시보드 공유\n→ Outlook 자동 발송', '프로젝트'],
                    ['3:20-4:00', '발표 및 수료', '결과 공유 + 수료', '발표'],
                ],
                'output': 'end-to-end AI Agent 자동화 파이프라인',
            },
        ],
    },
    'Claude': {
        'color': 'D97706',
        'subtitle': 'Claude in Chrome · Claude in Excel · Claude in Word',
        'ai_tool': 'Claude (Teams Plan)',
        'extra_tools': 'Claude in Chrome·Excel·Word 확장',
        'code_tool': 'Claude in Excel / Word',
        'spreadsheet': 'Excel (Claude 확장)',
        'doc_tool': 'Word (Claude 확장)',
        'cloud': 'OneDrive / Google Drive',
        'dashboard': 'Claude HTML 생성',
        'sessions': [
            {   # 1회차
                'title': '사내에서 AI를 안전하게 쓰는 법',
                'goal': 'AI 보안 리터러시 확보 + Claude 트랙 환경 세팅',
                'rows': [
                    ['0:00-0:50', '왜 ChatGPT가 막혔을까?', '삼성 소스코드 유출 사건 분석\n기업 데이터 유출 경로 이해', '강의 + 토론'],
                    ['0:50-1:30', '실습: 데이터 분류 기준표', '4등급 분류 + Claude 사용 시 주의사항\n외부 출처 파일 프롬프트 인젝션 이해', '개인 실습'],
                    ['1:30-2:10', 'AI 도구 보안 지도', 'Claude Teams Plan 보안 정책\nChrome 확장 사용 시 주의점\nOllama 로컬 AI 보안 범위', '강의'],
                    ['2:10-3:00', '실습: AI 도구 세팅', 'Claude in Chrome·Excel·Word 설치\nTeams Plan 확인\nOllama 설치 → 보안 설정 → Solar 실행', '실습'],
                    ['3:00-3:40', '실습: AI 코드 보안 검증', 'Claude에게 코드 요청\n→ 외부 URL/API 키 체크\n→ 신뢰할 수 있는 파일만 사용 원칙', '실습'],
                    ['3:40-4:00', '업무 자동화 지도 작성', '반복 업무 리스트업 + AI 매핑', '개인 작업'],
                ],
                'output': '데이터 분류 기준표 + Claude/Ollama 세팅 완료 + 코드 검증 체크리스트',
            },
            {   # 2회차
                'title': 'AI에게 시키기 + 스프레드시트 자동화',
                'goal': '프롬프트 엔지니어링 습득 + 자금일보/채권연령표 자동화',
                'rows': [
                    ['0:00-0:30', '프롬프트 엔지니어링', '같은 요청을 3가지로 비교\nClaude vs Ollama(Solar) 비교', '실습'],
                    ['0:30-1:30', '실습: 자금일보 자동화', 'Claude in Excel에서 직접\n"자금일보 자동화해줘" → 수식+매크로 생성\n안 되면 Ollama(Solar)로 동일 요청', '실습'],
                    ['1:30-2:20', '실습: 채권연령표 자동화', 'Claude에게 경과일수 분류 요청\n→ 수식 + VBA 매크로 동시 생성', '실습'],
                    ['2:20-3:10', '실습: Document AI 문서 파싱', '업스테이지 Document Parse로\nPDF 자동 파싱 → 스프레드시트 정리', '실습'],
                    ['3:10-4:00', 'AI 코드 보안 검증 실습', 'Claude가 만든 코드 검증\n외부 호출 확인 → 수정 요청', '실습 + 정리'],
                ],
                'output': '자금일보/채권연령표 자동화 + Document AI 체험',
            },
            {   # 3회차
                'title': 'AI Agent 워크플로우 & 보고서 자동화',
                'goal': 'AI Agent 개념 이해 + Claude 멀티앱 워크플로우',
                'rows': [
                    ['0:00-0:40', 'AI Agent 개념', 'AI Agent / MCP / Agentic Workflow\n도구를 연결하는 AI 소개', '강의 + 시연'],
                    ['0:40-1:40', '실습: 생산일보 Agent', 'Claude in Excel에서 데이터 분석\n→ Claude in Word에서 보고서 양식 채우기\n→ Ollama 로컬 대안 실습', '실습'],
                    ['1:40-2:40', '실습: Agentic 워크플로우', 'Claude in Chrome으로 웹 데이터 수집\n→ Claude in Excel로 자동 정리\n→ Claude in Word로 보고서 생성\n→ 이메일 발송', '실습'],
                    ['2:40-3:20', '실습: 로컬 AI Agent', 'Ollama(Solar)로 동일 워크플로우\n오프라인 구현', '실습'],
                    ['3:20-4:00', 'AI Agent 한계와 주의점', '오류 대응 + 검증 포인트', '강의 + Q&A'],
                ],
                'output': 'Claude 멀티앱 연계 AI Agent 파이프라인',
            },
            {   # 4회차
                'title': '바이브코딩: 만들고, 공유하고, 쓰게 하기',
                'goal': 'HTML 대시보드 제작 + 로컬 네트워크 팀 공유',
                'rows': [
                    ['0:00-0:30', '바이브코딩 소개', '비즈니스 활용 사례 시연', '강의 + 시연'],
                    ['0:30-1:20', '실습: HTML 대시보드', 'Claude에게 "생산 실적 대시보드\nHTML 만들어줘" → 데이터 연동', '실습'],
                    ['1:20-2:00', '실습: 업무 도구 제작', '"자금일보 웹앱"\n"채권 계산기" 등 Claude로 제작', '실습'],
                    ['2:00-2:40', '실습: 로컬 네트워크 공유', 'python -m http.server\n→ 팀원이 브라우저로 접속', '실습'],
                    ['2:40-3:20', '비즈니스 활용 케이스', '공장 태블릿 / 영업 현황판 / KPI 보드', '토론 + 실습'],
                    ['3:20-4:00', '자유 제작 + 상호 접속', '본인 도구 + 팀원 접속', '실습'],
                ],
                'output': 'HTML 대시보드 + 로컬 네트워크 공유',
            },
            {   # 5회차
                'title': '로컬 AI 심화: Solar LLM vs Gemma 4',
                'goal': '모델 비교 + 오프라인 실무 파이프라인 완성',
                'rows': [
                    ['0:00-0:40', '로컬 AI 모델 비교', 'Solar Mini vs Gemma 4 vs Llama\n한국어/코드/문서 성능 직접 비교', '실습'],
                    ['0:40-1:30', '실습: Solar로 문서 일괄 처리', '파일 100개 자동 분류·이름 변경·요약', '실습'],
                    ['1:30-2:20', '실습: 로컬 AI + 엑셀', 'CSV → 월간 보고서 자동화', '실습'],
                    ['2:20-3:10', '실습: LM Studio GUI', '클릭만으로 AI 챗봇 실행', '실습'],
                    ['3:10-4:00', '보안 등급별 AI 선택 가이드', '도구별 보안 조건 총정리', '강의 + 개인 작업'],
                ],
                'output': 'Solar vs Gemma 비교 + 로컬 AI 실무 파이프라인',
            },
            {   # 6회차
                'title': '전부 연결하기: 종합 프로젝트',
                'goal': 'end-to-end AI 자동화 파이프라인 완성',
                'rows': [
                    ['0:00-0:40', '클라우드 드라이브 PC 연동', 'OneDrive/Google Drive 마운트\n자동 동기화', '실습'],
                    ['0:40-1:20', 'AI로 파일 자동 정리', 'Claude에게 파일 정리 스크립트 요청', '실습'],
                    ['1:20-3:20', '종합 프로젝트', '데이터 수집 → Document AI\n→ Claude Agent 정리 → 보고서\n→ 클라우드 저장 + 대시보드 공유\n→ 자동 발송', '프로젝트'],
                    ['3:20-4:00', '발표 및 수료', '결과 공유 + 수료', '발표'],
                ],
                'output': 'end-to-end Claude AI Agent 자동화 파이프라인',
            },
        ],
    },
}

# ============================================================
# GENERATE DOCUMENTS
# ============================================================

for track_name, track in tracks.items():
    doc = create_doc()
    color = track['color']

    # Title
    title = doc.add_heading(f'수업계획서 — {track_name} 트랙', level=0)
    for run in title.runs:
        run.font.name = 'Malgun Gothic'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
        run.font.size = Pt(18)

    para(doc, f'AI로 쉬워지는 직장인 업무 자동화', bold=True, size=12)
    para(doc, track['subtitle'], size=10, color=(100, 100, 100))

    # Overview table
    styled_heading(doc, '과정 정보', level=2)
    make_table(doc,
        ['항목', '내용'],
        [
            ['과정명', 'AI로 쉬워지는 직장인 업무 자동화'],
            ['트랙', f'{track_name} 트랙 ({track["subtitle"]})'],
            ['교육 대상', '외부 AI 사용 제한 기업 재직자 (비개발자)'],
            ['총 교육 시간', '24시간 (6회 × 4시간)'],
            ['교육 방식', '이론 20% + 실습 80%'],
            ['주요 AI 도구', track['ai_tool']],
            ['추가 도구', track.get('extra_tools', '-')],
            ['로컬 AI (공통)', 'Ollama + Upstage Solar / Gemma 4'],
        ],
        header_color=color
    )

    # Sessions
    for i, session in enumerate(track['sessions']):
        if i > 0 and i % 2 == 0:
            doc.add_page_break()

        styled_heading(doc, f'{i+1}회차: {session["title"]}', level=2)
        para(doc, f'학습 목표: {session["goal"]}', bold=True, size=9)

        make_table(doc,
            ['시간', '주제', '실습 내용', '수업 형태'],
            session['rows'],
            header_color=color
        )

        para(doc, f'산출물: {session["output"]}', bold=True, size=9, color=(0, 100, 0))
        para(doc, '')

    # Save
    path = rf'C:\Users\moonk\OneDrive\AI강사준비\수업계획서_{track_name}_트랙.docx'
    doc.save(path)
    print(f'Saved: {path}')

print('Done! 3 files created.')
