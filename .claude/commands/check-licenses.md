# AI 서비스 라이센스/정보보호 정책 자동 수집 (playwright-cli)

## 목적

`playwright-cli`로 각 AI 서비스의 공식 정보보호/개인정보보호/학습정책 페이지를 방문하여 핵심 내용을 수집하고, **출처 URL과 함께** 라이센스 검토 근거를 남긴다.

## 반드시 playwright-cli를 사용할 것!

> **Python playwright 라이브러리 사용 금지.** 반드시 `npx playwright-cli` 또는 `playwright-cli` CLI 명령을 Bash 도구로 실행한다.

## 수집 대상 URL (22개)

### Google 트랙
1. `https://policies.google.com/terms/generative-ai` — Gemini 생성형 AI 약관
2. `https://ai.google.dev/gemini-api/terms` — AI Studio API 약관 (유료/무료 차이)
3. `https://workspace.google.com/terms/dpa_terms.html` — Workspace DPA
4. `https://workspace.google.com/terms/gemini-ai/` — Gemini for Workspace AI 약관
5. `https://labs.google/terms` — Google Labs 실험 서비스 약관
6. `https://policies.google.com/privacy` — Google 개인정보처리방침

### Microsoft 트랙
7. `https://docs.github.com/en/copilot/responsible-use-of-github-copilot-features/github-copilot-general-privacy-faq` — Copilot 개인정보 FAQ
8. `https://docs.github.com/en/site-policy/github-terms/github-terms-for-additional-products-and-features` — GitHub 제품 약관
9. `https://learn.microsoft.com/en-us/compliance/assurance/assurance-data-protection` — M365 데이터 보호
10. `https://learn.microsoft.com/en-us/sharepoint/safeguarding-your-data` — OneDrive 보안

### Claude 트랙
11. `https://www.anthropic.com/policies/consumer-terms` — Anthropic 이용약관
12. `https://www.anthropic.com/policies/privacy` — Anthropic 개인정보처리방침
13. `https://www.anthropic.com/policies/usage-policy` — Anthropic 사용 정책
14. `https://www.anthropic.com/claude/teams` — Claude Teams/Enterprise 정보

### 로컬 AI / 공통
15. `https://github.com/ollama/ollama/blob/main/LICENSE` — Ollama MIT 라이센스
16. `https://lmstudio.ai/terms` — LM Studio 이용약관
17. `https://huggingface.co/upstage/solar-pro-preview-instruct` — Solar 모델 라이센스
18. `https://ai.google.dev/gemma/terms` — Gemma 이용약관
19. `https://www.upstage.ai/policies/terms-of-service` — Upstage Document AI 약관

## 수집 절차 (playwright-cli)

$ARGUMENTS에 트랙명이 있으면 해당 트랙만, 없으면 전체 수집.

각 URL에 대해 아래 단계를 반복:

### STEP 1: 페이지 열기
```bash
playwright-cli goto "<URL>"
```

### STEP 2: 페이지 스냅샷 (텍스트 추출)
```bash
playwright-cli snapshot --raw
```

### STEP 3: 핵심 키워드 검색 (JavaScript eval)
```bash
playwright-cli eval "document.body.innerText" --raw
```

### STEP 4: 특정 섹션 추출 (training, data 관련)
```bash
playwright-cli eval "(() => {
  const text = document.body.innerText;
  const keywords = ['training', 'train', 'improve', 'human review', 'data retention', 'customer data', 'learning'];
  const lines = text.split('\\n');
  const matches = [];
  lines.forEach((line, i) => {
    keywords.forEach(kw => {
      if (line.toLowerCase().includes(kw)) {
        const context = lines.slice(Math.max(0,i-1), Math.min(lines.length,i+3)).join(' ');
        matches.push({keyword: kw, context: context.substring(0, 300)});
      }
    });
  });
  return JSON.stringify(matches.slice(0, 20));
})()" --raw
```

### STEP 5: 결과를 licenses/ 폴더에 마크다운으로 저장

각 URL의 수집 결과를 아래 형식으로 `licenses/라이센스_검토_YYYYMMDD.md`에 기록:

```markdown
## [트랙명] 서비스명

- **URL**: [출처 URL](URL)
- **수집 일시**: YYYY-MM-DD HH:MM
- **학습정책 태그**: 🔴/🟡/🟢

### 핵심 발견
> (training/data 관련 원문 인용)

### 강사 경고 문구
> "실습 시작 전 이 말을 하세요"

### ISMS-P 관련 조항
- 2.X.X 항목명: 준수/위반 리스크
```

## 주의사항

- **반드시 playwright-cli CLI 명령을 사용**할 것. Python playwright 라이브러리 금지.
- 각 페이지에서 2~3초 대기 (렌더링 완료 대기)
- 로그인 필요 페이지는 스냅샷만 취하고 "로그인 필요" 표시
- 수집 결과에 **반드시 원본 URL(출처)** 포함
- 정책은 수시 변경되므로 **수업 1주 전에 반드시 재실행**
