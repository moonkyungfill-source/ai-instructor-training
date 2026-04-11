# 문서보안 DRM 환경에서 AI 활용 가이드

## DRM이란?

한국 기업 대부분은 Fasoo DRM, SoftCamp SHIELDEX, MarkAny 등 문서보안 솔루션을 사용.
DRM이 걸린 파일은 **승인된 프로그램(Excel, Word, HWP 등)에서만** 열 수 있고,
외부 프로그램이 직접 파일을 읽으면 **암호화된 바이너리**만 보임.

## 핵심 원리: "앱 안에서" vs "앱 밖에서"

```
[Excel 앱 내부]                    [Excel 앱 외부]
VBA ✅                             openpyxl ❌
Claude in Excel ✅                  pandas ❌
xlwings ✅ (Excel 앱 제어)          csv.reader ❌
Power Query ✅                     Document AI ❌
수식/함수 ✅                       Ollama 파일 읽기 ❌
```

**규칙**: DRM 파일을 다루려면 **Excel 앱 안에서 동작하는 방법**만 사용.

---

## 방법 1: VBA (가장 확실)

DRM 파일이 Excel에서 열리면 VBA는 정상 동작.

```vba
' AI에게 이 코드를 만들어달라고 요청:
' "열려있는 엑셀에서 A열 데이터를 읽어서 합계를 구하는 VBA 만들어줘"

Sub 자금일보_자동화()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' DRM 파일이어도 VBA로 데이터 읽기/쓰기 가능
    Dim total As Double
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        total = total + ws.Cells(i, 3).Value  ' C열 금액 합계
    Next i
    
    ws.Cells(lastRow + 1, 3).Value = total
    MsgBox "자금일보 합계: " & Format(total, "#,##0")
End Sub
```

**강사 안내**: "DRM 걸린 엑셀이어도 VBA는 됩니다. AI에게 코드만 받아서 붙여넣으세요."

---

## 방법 2: xlwings (Python에서 Excel 제어)

xlwings는 Python에서 **Excel 앱 자체를 원격 제어**하므로 DRM 통과.

```python
# AI에게 이 코드를 만들어달라고 요청:
# "xlwings로 열려있는 엑셀에서 데이터를 읽어서 자금일보 요약하는 Python 만들어줘"

import xlwings as xw

# 현재 열려있는 Excel 파일에 연결 (DRM 파일이어도 OK!)
wb = xw.books.active
ws = wb.sheets[0]

# 데이터 읽기
data = ws.range('A1').expand().value

# 처리 (pandas와 조합 가능)
import pandas as pd
df = pd.DataFrame(data[1:], columns=data[0])
summary = df.groupby('계정과목')['금액'].sum()

# 결과를 새 시트에 쓰기
new_sheet = wb.sheets.add('자금일보_요약')
new_sheet.range('A1').value = summary.reset_index().values.tolist()
```

**핵심**: `xw.books.active`는 이미 Excel에서 열려있는 파일에 연결하므로 DRM 우회가 아닌 **정상 접근**.

### xlwings 설치
```bash
pip install xlwings
```

### xlwings vs openpyxl 차이

| | xlwings | openpyxl |
|---|---------|----------|
| 작동 방식 | Excel 앱을 제어 | 파일을 직접 읽기 |
| DRM 파일 | ✅ 가능 | ❌ 불가 |
| Excel 필요 | 필요 (설치되어 있어야 함) | 불필요 |
| 속도 | 느림 (앱 통신) | 빠름 (파일 직접) |
| 서식/차트 | 완벽 지원 | 일부 지원 |

---

## 방법 3: Claude in Excel (가장 쉬움)

Claude in Excel 애드인은 **Excel 안에서 동작**하므로 DRM 파일에서도 사용 가능.

```
[수강생 실습 흐름]
1. DRM 걸린 자금일보.xlsx를 Excel에서 연다
2. Claude in Excel 애드인 열기
3. "이 시트의 데이터로 자금일보 요약을 만들어줘" 입력
4. Claude가 Excel 안에서 바로 결과 생성
```

**단점**: Claude in Excel은 **데이터를 Anthropic 서버로 전송**하므로, DRM 파일이라도 데이터 내용은 외부로 나감. → **더미 데이터로만 실습!**

---

## 방법 4: 텍스트 복사-붙여넣기 (가장 범용)

DRM 파일이라도 **Excel에서 열고 → 셀을 복사 → AI에 붙여넣기**는 가능한 경우가 많음.
(단, 일부 강력한 DRM은 복사도 차단)

```
[흐름]
1. DRM 엑셀 파일을 Excel에서 연다
2. 필요한 범위 선택 → Ctrl+C (복사가 되는 DRM인 경우)
3. ChatGPT/Gemini/Claude 웹에 Ctrl+V
4. "이 데이터로 자금일보 요약해줘" 요청
```

**주의**: 복사한 데이터가 외부 AI로 나가므로 **더미 데이터 원칙** 동일 적용.

---

## 수업계획서 트랙별 DRM 대응

### Microsoft 트랙 (DRM 최적)

| 회차 | DRM 환경 대응 |
|------|-------------|
| 2회차 자금일보 | **VBA** (DRM 통과) + Copilot이 코드 생성 |
| 3회차 생산일보 | **VBA + xlwings** (DRM 통과) + PowerShell은 **텍스트 파일만** |
| 4회차 대시보드 | HTML은 DRM 무관 (새로 생성하는 파일) |

### Claude 트랙 (DRM 호환)

| 회차 | DRM 환경 대응 |
|------|-------------|
| 2회차 자금일보 | **Claude in Excel** (DRM 통과, 단 데이터 외부 전송) |
| 3회차 MCP | MCP filesystem은 **DRM 파일 읽기 불가** → 텍스트 복사 방식 |
| 4회차 Artifacts | 새로 생성하는 파일이므로 DRM 무관 |

### Google 트랙 (DRM 해당 없음)

| 회차 | DRM 환경 대응 |
|------|-------------|
| 전체 | Google Sheets는 별도 생태계 → **DRM 미적용** |
| | 단, 기업이 Google Workspace를 안 쓰면 이 트랙 자체가 불가 |

---

## 교육 전 DRM 체크리스트

```
□ 기업 DRM 솔루션 확인 (Fasoo? SoftCamp? MarkAny?)
□ DRM이 복사-붙여넣기를 차단하는지 확인
□ Excel VBA 매크로 실행이 DRM과 공존 가능한지 확인
□ Claude in Excel 애드인이 DRM과 충돌하는지 확인
□ xlwings 설치 가능 여부 확인 (Python + pip)
□ DRM 없는 더미 데이터 사전 준비 (교육용)
□ DRM 파일 반출 정책 확인 (USB, 메일, 클라우드)
```

## 강사 멘트

> "우리 회사 엑셀에 DRM이 걸려있죠? 걱정 마세요. VBA는 Excel 안에서 돌아가니까 DRM 상관없이 됩니다. AI에게 VBA 코드를 받아서 Excel에 붙여넣기만 하면 DRM 파일도 자동화할 수 있습니다. 다만 오늘 실습은 DRM 없는 더미 데이터로 합니다."
