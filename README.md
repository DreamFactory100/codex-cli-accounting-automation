# ABC Company 회계분석 자동화 실습

Codex CLI, Windows PowerShell, Python을 연결해 가상 회계 거래를 분석하고, 단위 테스트와 독립 검증을 거친 뒤 AI 요약 보고서를 생성하는 교육용 프로젝트입니다. 실제 기업의 장부나 감사 자료가 아니며, 이상 항목 탐지와 검증 자동화의 역할을 실습하는 데 목적이 있습니다.

## 주요 기능

- UTF-8 CSV의 필수 열, 날짜, 정수 금액, 차변·대변 규칙과 식별자 중복 검사
- 월별·계정별 차변/대변 집계와 전표별 대차 검증
- 전표 전체 복제 및 전표 내부 중복 행 후보 탐지
- 5개 시트로 구성된 서식 적용 Excel 보고서 생성
- 17개 단위 테스트 실행
- 분석 구현을 import하지 않는 별도 검증기로 저장된 Excel과 정답표 비교
- PowerShell에서 분석, 테스트, 검증, Codex AI 보고서 생성을 순차 실행하고 단계별 로그 보존

중복 항목은 자동으로 오류 확정하지 않고 `검토 필요` 후보로 표시합니다.

## 사용 환경

검증에 사용한 환경은 다음과 같습니다.

- Windows 11
- Windows PowerShell 5.1
- Python 3.12
- Codex CLI (`codex.cmd`가 `PATH`에서 검색 가능해야 함)

명령은 프로젝트 루트의 Windows PowerShell 5.1에서 실행하는 것을 기준으로 합니다.

## Codex CLI 설치 및 ChatGPT 인증

Gate 8의 마지막 단계는 `codex exec`를 호출하므로 Codex CLI가 필요합니다. Node.js와 npm이 설치되어 있다면 PowerShell에서 다음과 같이 설치할 수 있습니다.

```powershell
npm install -g @openai/codex
codex --version
```

프로젝트 폴더에서 아래 명령을 실행하고 브라우저에서 ChatGPT 로그인을 완료합니다.

```powershell
codex login
codex login status
```

첫 `codex` 실행 화면에서 `Sign in with ChatGPT`를 선택해도 됩니다. 인증 캐시나 토큰은 비밀정보이므로 저장소에 추가하지 마세요. 설치·인증 방식은 변경될 수 있으므로 [Codex CLI 공식 안내](https://learn.chatgpt.com/docs/codex/cli)와 [공식 인증 안내](https://learn.chatgpt.com/docs/auth?surface=cli)를 함께 확인하세요.

## 프로젝트별 Python 가상환경 준비

Python 3.12가 설치되어 있는지 확인한 뒤 프로젝트 루트에서 전용 가상환경을 만듭니다.

```powershell
py -3.12 --version
py -3.12 -m venv .venv
```

이 문서의 명령은 활성화 여부에 영향을 받지 않도록 `.venv`의 Python을 직접 지정합니다. 활성화해서 사용하려면 다음을 실행할 수도 있습니다.

```powershell
.\.venv\Scripts\Activate.ps1
```

실행 정책 때문에 활성화가 차단되더라도 아래의 직접 경로 방식은 사용할 수 있습니다.

## 패키지 설치

```powershell
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

`requirements.txt`에는 검증 당시 사용한 `pandas`, `openpyxl`, `numpy` 등의 버전이 고정되어 있습니다.

## 개별 실행

### 1. Python 분석

기존 파일은 덮어쓰지 않으므로 아직 존재하지 않는 출력 파일명을 지정해야 합니다.

```powershell
.\.venv\Scripts\python.exe -B src\analyze_transactions.py `
  --input data\abc_transactions_2026Q1.csv `
  --output outputs\abc_accounting_analysis_reproduced.xlsx
```

성공하면 표준 출력에 `status: PASS`, 출력 경로, 시트별 행 수가 JSON으로 표시됩니다.

### 2. 17개 단위 테스트

```powershell
.\.venv\Scripts\python.exe -B -m unittest discover -s tests
```

### 3. 독립 검증

분석 단계에서 방금 만든 Excel을 검증합니다.

```powershell
.\.venv\Scripts\python.exe -B src\validate_results.py `
  --workbook outputs\abc_accounting_analysis_reproduced.xlsx `
  --answer-key data\validation_answer_key.csv
```

검증기는 분석 모듈을 import하지 않고 저장된 Excel을 정답 CSV와 비교합니다. 월별·계정별 집계, 전체 전표 차액, 중복 후보, 정상 사례의 오탐, 요약 교차합계, 시트 순서, 헤더 서식, 열 너비, 숫자 표시 형식과 저장값을 검사합니다.

## Gate 8 통합 자동화

Python 환경과 Codex 로그인이 준비된 뒤 프로젝트 루트에서 실행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\run_gate8.ps1
```

`ExecutionPolicy Bypass`는 이 명령으로 시작한 새 PowerShell 프로세스에만 적용되며 시스템 전체의 실행 정책을 변경하지 않습니다.

스크립트는 다음 순서로 실행됩니다.

1. 거래 CSV 분석 및 Excel 생성
2. 17개 단위 테스트
3. 정답표를 이용한 독립 검증과 `PASS` 확인
4. 테스트 로그와 검증 JSON만 근거로 Codex 한국어 Markdown 보고서 생성

각 실행은 `runs\yyyyMMdd-HHmmss-fff-고유값\` 폴더를 새로 만들며 기존 실행 결과를 덮어쓰지 않습니다. 어느 단계든 실패하면 즉시 중단하고 이미 만든 산출물, 단계 로그 및 `failure.log`를 보존합니다. 4단계에는 네트워크 연결과 유효한 Codex 인증이 필요합니다.

## 결과 폴더와 주요 산출물

### `runs/<실행 ID>/`

- `accounting_analysis.xlsx`: 해당 실행에서 새로 생성한 정상 Excel 보고서
- `analysis.log`: 분석 프로세스 종료 코드, 표준 출력과 표준 오류
- `tests.log`: 17개 테스트 결과
- `validation.json`: 독립 검증의 구조화된 결과
- `validation.log`: 독립 검증 종료 코드와 전체 출력
- `ai_report.md`: 검증 결과를 요약한 Codex 보고서
- `codex.log`: Codex 실행 로그
- `failure.log`: 통합 실행 실패 시 실패 사유

### `outputs/`의 두 기존 Excel에 대한 주의

- `outputs/abc_accounting_analysis.xlsx`는 이전 단계에서 만든 파일입니다. 계산값과 시트 구조는 같지만 일부 열의 명시적 너비 설정이 빠져 있어 현재의 독립 검증 기준을 통과하는 정상본으로 취급하면 안 됩니다.
- `outputs/abc_accounting_analysis_gate6.xlsx`는 Gate 6에서 새로 생성되어 열 너비를 포함한 저장값·서식 검증을 통과한 정상 Excel입니다.
- Gate 8의 `runs/<실행 ID>/accounting_analysis.xlsx`도 수정된 현재 소스로 매번 새로 생성되고 독립 검증을 통과한 파일입니다.

`outputs/`와 `runs/`는 재생성 가능한 실행 산출물이므로 Git 추적 대상에서 제외됩니다.

## 실제 확인된 검증 결과

저장된 성공 실행 기록에서 다음을 확인했습니다.

- 단위 테스트: `Ran 17 tests` / `OK` / 종료 코드 0
- 독립 검증 JSON: `status = PASS`
- 5개 시트 데이터 행 수: 분석요약 17, 월별계정집계 33, 전표대차검증 62, 중복의심전표 2, 중복의심행 1
- 월별·계정별 일치: 33건
- 대차 불일치: 3건 (`ABC-202601-019` 1,000원, `ABC-202602-019` 2,000원, `ABC-202603-020` 30,000원)
- 전체 복제 의심 전표: 2쌍
- 전표 내부 중복 행 그룹: 1개 (`ABC-202603-020`, 분개순번 1과 3)
- 정상 동일 금액 사례 3건과 정상 월별 반복 사례 9건의 오탐: 0건
- 총 차변 31,027,000원, 총 대변 30,994,000원
- 스타일 및 저장값 검사: `PASS`

이는 저장된 실행 로그와 검증 JSON에서 확인된 결과이며, 다른 데이터나 수정된 코드에 그대로 적용되는 보장은 아닙니다.

## 교육용 자료와 의도된 오류

`data/abc_transactions_2026Q1.csv`는 ABC Company를 가정해 만든 2026년 1분기 교육용 자료입니다. 실제 기업·거래처·거래가 아니며 세무신고나 완전한 재무제표 작성용 자료도 아닙니다.

학습과 검증을 위해 다음 오류가 의도적으로 포함되어 있습니다.

- 대차 불일치 전표 3건: 1,000원, 2,000원, 30,000원 차이
- 실제 오류로 설계한 전표 5건: 전체 복제 2건과 행 복제 1건을 포함
- 정상인데 반복되거나 금액이 같은 비교 사례를 함께 포함해 단순 규칙의 오탐 여부 확인

중복 후보 3개가 모두 별개의 오류 3건이라는 뜻은 아닙니다. 예를 들어 `ABC-202603-020`은 대차 불일치와 내부 행 복제가 겹친 하나의 오류 전표입니다. 전체 복제 쌍에서는 원본 전표는 정상이고 추가된 복제 전표가 의도된 오류입니다.

## AI 보고서와 기계 검증의 역할 및 한계

기계적 검증은 미리 정한 규칙과 정답표에 따라 결과값, 구조, 서식과 정상 사례 오탐을 재현 가능하게 판정합니다. 다만 정답표 자체가 잘못되었거나 규칙에 없는 업무 위험이 있으면 발견하지 못할 수 있습니다. 현재 검증기는 원본 거래 CSV를 다시 계산하는 대신 저장된 Excel과 별도 정답표를 비교하므로, 서로 같은 방식으로 잘못 작성된 두 자료까지 독립적으로 보증하지는 못합니다.

Codex AI 보고서는 `validation.json`과 `tests.log`의 기록을 읽어 사람이 보기 쉬운 한국어 설명으로 요약합니다. AI는 Gate 8에서 계산을 다시 수행하거나 원본의 진실성, 회계처리의 적법성, 감사의견을 보증하지 않습니다. AI 문장은 보조 설명이며, 합격 여부의 근거는 테스트와 독립 검증 결과입니다.

## 다른 PC에서 재현할 때

- 프로젝트 전체 경로가 달라도 되지만 모든 명령은 프로젝트 루트에서 실행하세요.
- Windows PowerShell 5.1과 Python 3.12를 사용하고 프로젝트마다 `.venv`를 새로 만드세요. 다른 PC의 `.venv`를 복사하지 마세요.
- `requirements.txt`의 고정 버전을 그대로 설치하세요.
- CSV의 UTF-8/BOM, 한글 파일 내용과 줄바꿈이 변환되지 않도록 주의하세요.
- 분석기는 프로젝트 밖 경로를 거부하고 기존 Excel을 덮어쓰지 않습니다. 재실행 시 새 출력 이름을 쓰세요.
- Gate 8 전에 `codex.cmd`가 `PATH`에 있고 `codex login status`가 성공하는지 확인하세요.
- `runs/`, `outputs/`, `.venv/`는 배포되지 않으므로 새 PC에서 직접 생성해야 합니다.
- `data/validation_answer_key.csv`는 독립 검증용이며 분석기는 이 파일을 읽지 않습니다. 분석 입력으로 사용하거나 분석 로직에 결합하지 마세요.
- 실제 회사 데이터로 바꾸려면 개인정보·영업비밀 보호, 계정 체계, 중요성 기준, 회계정책과 법규를 별도로 검토해야 합니다.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
