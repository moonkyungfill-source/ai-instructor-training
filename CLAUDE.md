# Upstage AI 교육 전문 강사 지원 프로젝트

## 프로젝트 개요

Upstage "AI 교육 전문 강사 및 멘토풀" 모집에 **TRACK 1: 강사**로 지원하기 위한 서류 준비 프로젝트.

- **회사**: 업스테이지 (Upstage) - "Making AI Beneficial", "Building intelligence for the future of work"
- **지원 트랙**: TRACK 1 강사 - AI 기술 및 서비스 개발 관련 이론/실습 강의, 커리큘럼 설계 및 교안 제작
- **과제 참고**: https://go.upstage.ai/4l9naSJ

## 제출 서류 (필수)

| # | 서류 | 형식 | 상태 |
|---|------|------|------|
| 1 | 강의 이력서 및 경력 기술서 | PDF | 미완 |
| 2 | 강의 시연 영상 | 영상 (5~10분) | 미완 |
| 3 | 사전 과제: 강의 기획서 (커리큘럼 포함) | PDF, A4 2페이지 | **완료** (강의기획서_제출용.docx) |

## 과정명: AI로 쉬워지는 직장인 업무 자동화

### 3트랙 구조
- **Google 트랙**: AI Studio, Gemini, Apps Script, Google Labs (Opal, Mixboard, Pomely, Flow), NotebookLM
- **Microsoft 트랙**: GitHub Copilot (웹 + CLI), VBA, PowerShell, 작업 스케줄러, OneDrive
- **Claude 트랙**: Claude in Chrome·Excel·Word, Artifacts, MCP, Projects
- **공통**: Ollama + Upstage Solar / Gemma 4 (로컬 AI), Document AI

## ⚠️ 수업 설계 원칙 (검토 반영)

### 비개발자 따라하기 원칙

1. **30분 성공 경험**: 매 회차 시작 30분 안에 수강생 화면에 결과물이 나와야 한다
2. **하나의 회차 = 하나의 큰 성공 경험**: 회차당 핵심 도구 1~2개, 새 개념 2개 이하
3. **체험 먼저, 이름은 나중에**: "Agentic Workflow를 배우겠습니다" ❌ → "AI가 혼자 판단하게 해봅시다" ✅ → 실습 후 "방금 한 게 Agent입니다"
4. **웹 브라우저 우선**: 설치형 도구에는 반드시 웹 기반 대안(플랜B) 준비
5. **에러는 정상**: "에러 메시지를 복사해서 AI에게 물어보세요" 패턴을 1회차에서 가르침
6. **80% 완성 템플릿**: 빈 칸에서 시작하지 않고, 거의 완성된 템플릿에 AI로 나머지 채우기

### 보안 원칙

1. **데이터 4등급 분류**: (1) AI 자유 (2) 클라우드 OK (3) 로컬만 (4) 절대 불가
2. **AI 코드 검증 4단계**: 말하기 → 받기 → **검증하기** → 실행하기
3. **도구별 정확한 보안 조건 명시**: "최강 보안" 같은 과장 표현 금지
4. **플랜별 차이 명시**: 무료 vs 유료, Business vs Free 학습 정책 구분

### 🔴 라이센스/학습정책 경고 원칙 (CRITICAL)

> **무료 플랜 = 기업 내부 자료가 모델 학습에 쓰일 수 있다!**

수업계획서의 **모든 도구**에 대해 아래 3가지를 반드시 명시해야 한다:

1. **학습정책 태그**: 해당 플랜에서 사용자 데이터가 AI 모델 학습에 사용되는지 여부
   - 🔴 학습 사용 (위험) / 🟡 조건부 안전 / 🟢 학습 미사용 (안전)
2. **강사 경고 문구**: 실습 시작 전 수강생에게 반드시 말해야 할 한 줄
3. **실습 데이터 지침**: 실제 업무 데이터 사용 가능 여부, 더미 데이터 필요 여부

#### 도구별 학습정책 요약

| 도구 | 무료 | 유료 (안전 플랜) | 강사 경고 |
|------|------|-----------------|----------|
| Google Gemini | 🔴 학습 사용 | 🟢 Workspace 유료 | "무료 Gemini에 회사 자료 넣지 마세요" |
| Google AI Studio | 🔴 학습 사용 | 🟡 Billing 계정 시 미사용 | "유료 계정 확인 후 사용하세요" |
| GitHub Copilot | 🔴 Free/Pro 학습 사용 (2026.4~) | 🟢 Business/Enterprise | "반드시 Business 플랜인지 확인!" |
| Claude | 🔴 Free 학습 사용 | 🟢 Pro/Teams/Enterprise 미사용 | "Pro 이상만 안전합니다" |
| Claude in Excel/Word | 🟡 Teams 미학습이나 CellShock 취약점 | 🟢 Enterprise | "외부 파일과 함께 사용 주의" |
| Ollama (로컬) | 🟢 데이터 외부 미전송 | 🟢 동일 | "보안 설정 필수 (localhost 바인딩)" |
| LM Studio (로컬) | 🟢 데이터 외부 미전송 | 🟢 동일 | "클로즈드 소스, IT 확인 권장" |
| Document AI (Upstage) | 🟡 API 전송 | 🟢 기업용 미학습 | "API로 전송됨, 더미 데이터 권장" |
| Google Labs (Opal 등) | 🔴 실험적, 학습+인간리뷰 | - | "절대 실제 업무 데이터 넣지 마세요" |
| AI Studio 무료 | 🔴 학습+인간리뷰어 열람 | 🟢 Billing 계정 | "인간 리뷰어가 볼 수 있습니다!" |
| Drive+Gemini 연동 | 🔴 Gemini가 파일 읽으면 학습 | 🟢 Workspace DPA | "AI 연결 시 실습 전용 계정 사용" |

#### 수업계획서 작성 시 필수 규칙

- 모든 실습 항목에 🔴/🟡/🟢 태그 표시
- 🔴 도구 사용 시 **더미 데이터 필수** 안내
- 강사 경고 문구를 수업계획서 Word에 명시
- 교육 자료에 **"이 도구에 회사 자료를 넣어도 되나요?" 판단 플로차트** 포함

### 📋 ISMS-P 준수 원칙

수업계획서가 ISMS-P 인증 기업에서 정보보호 교육으로 인정받으려면:

1. **2.2.4 인식제고 및 교육훈련**: 교육 계획서, 출석부, 결과 보고서, 효과 평가 양식 필수
2. **2.1.3 정보자산 관리**: 데이터 보안등급 분류 실습이 1회차에 포함됨 ✅
3. **2.3.1 외부자 보안**: 외부 AI 서비스 사용 시 기업의 승인된 서비스 목록 확인 필수
4. **2.6 접근통제**: AI 도구 계정은 개인별 사용, 공유 계정 금지
5. **2.9.7 정보자산 폐기**: 교육 후 AI 대화 기록, 캐시, 더미 데이터 정리 가이드 필요

### ⚠️ 개인정보보호 원칙

1. **실습 데이터에 개인정보 포함 금지**: 자금일보, 채권연령표 실습 시 더미 데이터 사용
2. **AI에 개인정보 입력 금지 교육**: 1회차에서 "4등급 = AI에 절대 넣지 않는다" 강조
3. **Document AI 파싱 시**: 샘플 PDF에 실제 개인정보 없는지 사전 확인
4. **교육 후 데이터 정리**: AI 대화 기록, 업로드 파일, 캐시 삭제 가이드 제공
5. **개인정보보호법 제29조**: 안전조치의무 — AI 도구 사용도 안전조치에 해당

### 🔒 문서보안 DRM 대응 전략

> 대부분의 한국 기업은 Fasoo, SoftCamp, MarkAny 등 문서보안 DRM을 사용.
> DRM 걸린 파일은 **승인된 프로그램(Excel, Word 등)에서만** 열림.
> Python openpyxl, pandas 등으로 직접 읽기 **불가**.

#### DRM 환경에서 AI 도구 호환성

| 도구/방법 | DRM 파일 접근 | 이유 | 실습 가능 |
|-----------|-------------|------|----------|
| **Excel VBA** | ✅ 가능 | Excel 내부에서 실행 → DRM 통과 | ✅ |
| **Claude in Excel** | ✅ 가능 | Excel 애드인 → DRM 통과 | ✅ |
| **xlwings (Python)** | ✅ 가능 | Excel 앱을 제어 → DRM 통과 | ✅ |
| **Power Query** | ✅ 가능 | Excel 내부 기능 → DRM 통과 | ✅ |
| **openpyxl/pandas** | ❌ 불가 | 파일 직접 읽기 → DRM 차단 | ❌ |
| **Apps Script** | ⚠️ 해당없음 | Google Sheets는 별도 생태계 | ✅ (DRM 미적용) |
| **Document AI** | ❌ 불가 | 파일을 외부 전송 → DRM 차단 | ❌ (DRM 해제 필요) |
| **Ollama/LM Studio** | ❌ 불가 | 파일 직접 읽기 불가 | ❌ (텍스트 복붙은 가능) |

#### DRM 환경 실습 전략

1. **Excel 내부에서 해결**: VBA + Claude in Excel + xlwings = DRM 파일 직접 처리 가능
2. **텍스트 복사 방식**: DRM 파일을 Excel에서 열고 → 텍스트/데이터를 복사 → AI에 붙여넣기
3. **더미 데이터 우회**: DRM 없는 교육용 더미 엑셀 파일을 사전 준비 (가장 안전)
4. **xlwings 활용**: Python에서 Excel 앱을 제어하므로 DRM 통과. `import xlwings as xw; wb = xw.Book('drm_file.xlsx')` 가능

#### 수업계획서 반영 규칙

- **MS 트랙**: VBA + xlwings 중심으로 설계 (DRM 호환)
- **Claude 트랙**: Claude in Excel 애드인 (DRM 호환)
- **Google 트랙**: DRM 미적용 환경 (Google Sheets)
- 모든 트랙: **"파일을 직접 읽는 코드"보다 "Excel 앱 내에서 동작하는 코드" 우선**

### 기업 환경 현실 대응

1. **관리자 권한 없는 PC 전제**: Ollama/LM Studio는 사전 IT 협조 또는 웹 대안
2. **도메인 화이트리스트**: 교육 2주 전 IT팀에 URL 목록 전달
3. **Google 계정 필요 시**: 사전 안내 필수
4. **플랜B 매핑**: 모든 설치형 도구에 웹 기반 대안 준비
5. **라이센스 사전 확인**: 교육 전 기업의 AI 도구 플랜(무료/유료) 반드시 확인
6. **문서보안 DRM 확인**: 교육 전 기업의 DRM 솔루션 확인 → VBA/xlwings/Claude in Excel로 대응

## 커스텀 명령어

| 명령어 | 용도 |
|--------|------|
| `/curriculum-plan` | 강의 기획서 (사전 과제) 작성 |
| `/career-summary` | 경력 기술서 초안 작성 |
| `/application-answers` | 사전 질문 답변 작성 |
| `/research-trends` | 2026 AI 트렌드 리서치 |
| `/teaching-demo` | 강의 시연 준비 가이드 |
| `/lesson-plan` | 트랙별 수업계획서 작성/수정 |
| `/review-ux` | 비개발자 관점 수업 따라하기 검토 |
| `/review-security` | ISMS-P + 라이센스/학습정책 + 개인정보 종합 검토 |
| `/check-licenses` | Playwright로 AI 서비스 정책 페이지 자동 수집 (출처 보존) |
| `/summary-html` | 트랙별 HTML 1페이지 요약본 생성 (기업 담당자/수강생 배포용) |
| `/check-beginner` | 비개발자 따라하기 18항목 자동 체크 (품질 게이트) |
| `/pre-training-checklist` | 교육 전 기업 담당자 조율 35항목 체크리스트 생성 |
| `/pre-training-infra-test` | IT 인프라 환경 실측 테스트 (네트워크/서비스접속/설치/QR) |
| `/edu-tools` | 에듀테크 도구(Slido/Mentimeter) + QR 배포 + 실시간 Q&A 계획 |
| `/biz-cases` | 부서별 비즈니스 활용 케이스 개발 |

## 작업 원칙

1. **업스테이지 제품 연계**: Solar LLM, Document AI 등 업스테이지 기술을 커리큘럼에 자연스럽게 반영
2. **비개발자 최우선**: 코딩 경험 없는 직장인이 30분 안에 성공 경험을 느끼도록 설계
3. **최신 트렌드**: 2026년 AI Agent, Agentic Workflow, MCP, 바이브코딩 반영
4. **보안 현실 반영**: 도구별 정확한 보안 조건 명시, 과장 금지
5. **웹 우선, 설치 대안**: 기업 PC 제약 고려, 모든 실습에 웹 기반 플랜B

## 생성된 산출물

| 파일 | 용도 |
|------|------|
| `강의기획서_초안.md` | 상세 기획서 (면접/강의 준비용) |
| `강의기획서_제출용.md` | A4 2페이지 압축 (사전 과제) |
| `강의기획서_제출용.docx` | Word 변환본 |
| `수업계획서_Google_트랙.docx` | Google 트랙 상세 수업계획서 |
| `수업계획서_Microsoft_트랙.docx` | Microsoft 트랙 상세 수업계획서 |
| `수업계획서_Claude_트랙.docx` | Claude 트랙 상세 수업계획서 |
| `research/` | 트렌드 리서치, 도구 브레인스톰 |
| `review/` | 검토 리포트 (실현가능성, 보안, UX 등) |
