# 사용법

## 설치

Python 3.10 이상.

```bash
pip install pandas matplotlib plotly python-docx adjustText pywin32
```

`pywin32`는 HWPX 변환(Windows + 한컴오피스 한글 전용) 기능에만 필요합니다. Windows가 아니거나 한컴오피스가 없어도 나머지 기능은 동일하게 동작합니다.

## 입력 데이터 배치

```
input/seoul_expenses.csv
```

원본 CSV는 Git에 포함하지 않습니다(개인정보 가능 데이터, `.gitignore` 처리). 로컬에 직접 배치해야 합니다.

## 실행

### 전체 파이프라인 (권장)

```bash
python scripts/run_pipeline.py --input input/seoul_expenses.csv --formats html docx hwpx
```

원하는 형식만 선택할 수도 있습니다.

```bash
python scripts/run_pipeline.py --input input/seoul_expenses.csv --formats html docx
```

```bash
python scripts/run_pipeline.py --input input/seoul_expenses.csv --formats html
```

인자 없이 실행하면 사용법을 출력하고 종료합니다.

```bash
python scripts/run_pipeline.py
```

### 개별 스크립트 실행

```bash
# 1. 집계·차트·분석 JSON
python scripts/generate_charts.py

# 2. 인터랙티브 대시보드
python scripts/generate_dashboard.py

# 3. DOCX 보고서
python scripts/generate_docx_report.py

# 4. HWPX 변환 (Windows + 한컴오피스 전용)
python scripts/convert_docx_to_hwpx.py

# 5. 검증
python scripts/validate_data_and_privacy.py
python scripts/validate_docx_layout.py
python scripts/validate_hwpx_output.py
```

## 산출물 위치

| 경로 | 내용 |
|---|---|
| `workspace/*.json`, `*.md` | 중간 분석 산출물 |
| `output/charts/*.png` | 정적 차트 이미지 |
| `output/dashboard.html` | 인터랙티브 대시보드 |
| `output/*.docx`, `output/*.hwpx` | 최종 보고서 |
| `output/*_report.md` | 검증 보고서 |

## 문제 해결

| 증상 | 원인 | 대응 |
|---|---|---|
| HWPX가 생성되지 않음 | Windows가 아니거나 한컴오피스 미설치 | 정상 동작 — HTML·DOCX는 그대로 생성됨. 콘솔에 건너뛴 사유가 출력됨 |
| `ModuleNotFoundError: win32com` | pywin32 미설치 | `pip install pywin32` (Windows 전용) |
| 대시보드 차트가 비어 있음 | `input/seoul_expenses.csv` 컬럼명이 원본과 다름 | 원본 CSV 컬럼명(`해당월`, `비목`, `전체부서명`, `집행금액` 등)을 변경하지 않았는지 확인 |
| DOCX 표가 잘려 보임 | 사용 중인 Word/한글 버전의 렌더링 차이 | `docs/output_formats.md`의 레이아웃 규칙은 구조적으로 적용되어 있으나, 실제 렌더링은 프로그램마다 다를 수 있음 — 직접 열어 확인 필요 |

## GitHub Pages로 대시보드 공개하기

1. 저장소 Settings → Pages → Source를 `main` 브랜치의 `/docs` 폴더로 설정합니다.
2. `docs/index.html`(최종 대시보드 사본)이 공개 URL에서 자동으로 열립니다.
3. 대시보드를 갱신했다면 `output/dashboard.html`을 다시 `docs/index.html`로 복사한 뒤 커밋합니다(자동 동기화 없음).
