#!/bin/bash
# AI 서비스 라이센스/정보보호 정책 자동 수집 (playwright-cli 전용)
# 사용법: bash tools/collect_licenses.sh [google|microsoft|claude|local|all]

set -e

OUTPUT_DIR="C:/Users/moonk/OneDrive/AI강사준비/licenses"
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
REPORT="$OUTPUT_DIR/라이센스_검토_${TIMESTAMP}.md"

mkdir -p "$OUTPUT_DIR"

# 키워드 추출 JS 함수
EXTRACT_JS='(() => {
  const text = document.body.innerText;
  const keywords = ["training","train ","improve","human review","data retention","customer data","model","opt-out","opt out","학습","개인정보"];
  const lines = text.split("\\n").filter(l => l.trim().length > 10);
  const matches = [];
  lines.forEach((line, i) => {
    keywords.forEach(kw => {
      if (line.toLowerCase().includes(kw.toLowerCase())) {
        matches.push(kw + " >>> " + line.substring(0, 300));
      }
    });
  });
  return matches.slice(0, 10).join("\\n---\\n");
})()'

# 리포트 헤더
cat > "$REPORT" << 'HEADER'
# AI 서비스 라이센스/정보보호 정책 수집 결과

> **이 문서는 `playwright-cli`로 자동 생성되었습니다.**
> 정책은 수시 변경되므로 수업 전 반드시 재실행하세요.

HEADER

echo "**수집 일시**: $(date '+%Y-%m-%d %H:%M:%S')" >> "$REPORT"
echo "" >> "$REPORT"
echo "---" >> "$REPORT"
echo "" >> "$REPORT"

# 수집 함수
collect_page() {
  local TRACK="$1"
  local NAME="$2"
  local URL="$3"

  echo "[$TRACK] $NAME..."
  echo "  URL: $URL"

  echo "## [$TRACK] $NAME" >> "$REPORT"
  echo "" >> "$REPORT"
  echo "- **URL**: [$URL]($URL)" >> "$REPORT"
  echo "- **수집 일시**: $(date '+%Y-%m-%d %H:%M:%S')" >> "$REPORT"
  echo "" >> "$REPORT"

  # STEP 1: 페이지 방문
  GOTO_RESULT=$(playwright-cli goto "$URL" 2>&1) || true

  # 페이지 타이틀 추출
  TITLE=$(echo "$GOTO_RESULT" | grep "Page Title:" | head -1 | sed 's/.*Page Title: //')
  echo "- **페이지 제목**: $TITLE" >> "$REPORT"
  echo "" >> "$REPORT"

  # 잠시 대기 (렌더링)
  sleep 2

  # STEP 2: 키워드 추출
  echo "### 핵심 키워드 매칭 (training/data/학습 관련)" >> "$REPORT"
  echo "" >> "$REPORT"

  KEYWORDS=$(playwright-cli eval "$EXTRACT_JS" --raw 2>&1) || true

  if [ -n "$KEYWORDS" ] && [ "$KEYWORDS" != "\"\"" ] && [ "$KEYWORDS" != "null" ]; then
    # 키워드 결과를 인용 형식으로 저장
    echo "$KEYWORDS" | tr -d '"' | sed 's/\\n---\\n/\n\n/g' | while IFS= read -r line; do
      if [ -n "$line" ]; then
        echo "> $line" >> "$REPORT"
        echo "" >> "$REPORT"
      fi
    done
    echo "  ✅ 키워드 매칭 완료"
  else
    echo "> (키워드 매칭 없음 또는 접근 실패)" >> "$REPORT"
    echo "" >> "$REPORT"
    echo "  ⚠️ 키워드 매칭 없음"
  fi

  echo "---" >> "$REPORT"
  echo "" >> "$REPORT"
}

# 브라우저 열기
echo "=========================================="
echo "  playwright-cli 브라우저 시작"
echo "=========================================="
playwright-cli open 2>&1 | head -3

TRACK_ARG="${1:-all}"

# Google 트랙
if [ "$TRACK_ARG" = "all" ] || [ "$TRACK_ARG" = "google" ]; then
  echo ""
  echo "=========================================="
  echo "  GOOGLE 트랙 수집"
  echo "=========================================="
  collect_page "Google" "Gemini 생성형 AI 약관" "https://policies.google.com/terms/generative-ai"
  collect_page "Google" "AI Studio API 약관" "https://ai.google.dev/gemini-api/terms"
  collect_page "Google" "Workspace DPA" "https://workspace.google.com/terms/dpa_terms.html"
  collect_page "Google" "Gemini for Workspace AI 약관" "https://workspace.google.com/intl/en/terms/gemini-ai/"
  collect_page "Google" "Google Labs 약관" "https://labs.google/terms"
  collect_page "Google" "Google 개인정보처리방침" "https://policies.google.com/privacy"
fi

# Microsoft 트랙
if [ "$TRACK_ARG" = "all" ] || [ "$TRACK_ARG" = "microsoft" ]; then
  echo ""
  echo "=========================================="
  echo "  MICROSOFT 트랙 수집"
  echo "=========================================="
  collect_page "Microsoft" "Copilot 개인정보 FAQ" "https://docs.github.com/en/copilot/responsible-use-of-github-copilot-features/github-copilot-general-privacy-faq"
  collect_page "Microsoft" "GitHub 제품 약관" "https://docs.github.com/en/site-policy/github-terms/github-terms-for-additional-products-and-features"
  collect_page "Microsoft" "M365 데이터 보호" "https://learn.microsoft.com/en-us/compliance/assurance/assurance-data-protection"
  collect_page "Microsoft" "OneDrive 보안" "https://learn.microsoft.com/en-us/sharepoint/safeguarding-your-data"
fi

# Claude 트랙
if [ "$TRACK_ARG" = "all" ] || [ "$TRACK_ARG" = "claude" ]; then
  echo ""
  echo "=========================================="
  echo "  CLAUDE 트랙 수집"
  echo "=========================================="
  collect_page "Claude" "Anthropic 이용약관" "https://www.anthropic.com/legal/consumer-terms"
  collect_page "Claude" "Anthropic 개인정보처리방침" "https://www.anthropic.com/legal/privacy"
  collect_page "Claude" "Anthropic 사용 정책" "https://www.anthropic.com/legal/aup"
  collect_page "Claude" "Claude Teams/Enterprise" "https://www.anthropic.com/claude/teams"
fi

# 로컬 AI
if [ "$TRACK_ARG" = "all" ] || [ "$TRACK_ARG" = "local" ]; then
  echo ""
  echo "=========================================="
  echo "  로컬 AI / 공통 수집"
  echo "=========================================="
  collect_page "로컬" "Ollama 라이센스 (MIT)" "https://github.com/ollama/ollama/blob/main/LICENSE"
  collect_page "로컬" "LM Studio 이용약관" "https://lmstudio.ai/terms"
  collect_page "로컬" "Gemma 이용약관" "https://ai.google.dev/gemma/terms"
  collect_page "공통" "Upstage 이용약관" "https://www.upstage.ai/policies/terms-of-service"
fi

# 브라우저 닫기
playwright-cli close 2>&1 | head -1

echo ""
echo "=========================================="
echo "  수집 완료!"
echo "  리포트: $REPORT"
echo "=========================================="
