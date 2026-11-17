# 일일 금융지표 사이트 (GitHub Pages)

매 영업일 아침 7:30(한국시간)쯤 GitHub 서버가 자동으로 값을 수집해서 사이트와 엑셀 파일을 갱신합니다.
개인 PC는 꺼져 있어도 됩니다.

## 구성 (파일 설명)
- `.github/workflows/daily.yml` : 매일 자동 실행 설정 (월~금 7:30 한국시간)
- `run_daily_v6.py` : 수집 실행 → `docs/` 아래 파일 생성
- `collectors_v6.py` : 사이트별 수집 로직
  (금융투자협회 KOFIA, 한국자금중개, 네이버금융, global-rates.com의 Term SOFR,
  인증키가 있으면 한국은행 ECOS도 대체 경로로 사용)
- `holidays.csv` : 공휴일 목록 (연말에 다음 해 날짜를 추가해 주세요)
- `row_labels.json` : 채권 표 43행의 종류/종류명/신용등급 이름
- `docs/index.html` : 사이트 화면

자동으로 만들어지는 파일 (`docs/` 안, 직접 고치지 마세요)
- `data.json` : 가장 최근 수집 결과와 실패 내역
- `history.json` : 날짜별로 쌓이는 공개 데이터
- `private.json` : 날짜별 Term SOFR (비밀번호로 잠김)
- `daily_data.xlsx` : 내려받는 엑셀 파일
- `backfill_log.json` : 밀린 날짜 채우기 실행 기록
- `debug/` : 수집에 실패한 순간의 화면 캡처와 페이지 내용

## 설정값 (Secrets)
`Settings` → `Secrets and variables` → `Actions` → `New repository secret` 에서 등록합니다.
- `XLSX_PASSWORD` : Term SOFR 값을 잠그는 비밀번호. 없으면 Term SOFR는 저장하지 않습니다.
  **한 번 정하면 바꾸지 마세요.** 바꾸면 예전 기록을 지키려고 새 값 저장을 멈춥니다.
- `ECOS_API_KEY` (선택) : 한국은행 ECOS 무료 인증키. 있으면 CD·CP·KOSPI·환율을
  다른 곳에서 못 가져왔을 때 ECOS에서 대신 가져옵니다.

## 설치 순서 (최초 1회, 약 15분)

1. github.com 가입 (개인 이메일, 개인 장비 권장)
2. 우측 상단 `+` → `New repository`
   - 이름: 예) `daily-market`
   - **Public** 선택 (무료 계정은 Pages를 쓰려면 Public이어야 합니다)
   - `Create repository`
3. 압축을 푼 폴더 안의 **모든 파일과 폴더(`.github`, `docs` 포함)** 를 저장소 화면에 끌어다 놓기
   (`uploading an existing file` 링크 → 드래그 → 아래 `Commit changes`)
   - `.github` 폴더가 안 올라가면: `Add file` → `Create new file` → 이름칸에
     `.github/workflows/daily.yml` 입력 후 파일 내용을 붙여넣기
4. `Settings` → `Pages` → Source를 `Deploy from a branch`, Branch를 `main` / `/docs` 로 선택 → `Save`
5. `Actions` 탭 → 왼쪽 `daily-market-data` → `Run workflow` 버튼으로 첫 실행 (2~5분 소요)
6. 사이트 주소: `https://본인아이디.github.io/저장소이름/`
   (처음 한두 분은 404가 보일 수 있습니다)

## 매일 사용법
1. 외부망 PC에서 사이트 접속 → `엑셀 파일 내려받기`
2. 자료전송시스템으로 업무망에 전달
3. 업무망에서 받은 파일의 `전일` 시트 **E2:V44** 를 복사 → 원본 엑셀 `전일` 시트 **E2** 에 붙여넣기
   (`전전일`도 동일). 노란색 칸 / "미수집"은 직접 확인해서 채워 주세요.
4. `민평 금리`!A1 날짜 갱신, `그래프1` 값 붙여넣기는 기존처럼 진행

## 밀린 날짜 채우기
자동 실행이 며칠 멈췄거나 빠진 날이 있으면:
1. `Actions` 탭 → 왼쪽 `daily-market-data` → `Run workflow`
2. 시작일(예 `2026-11-03`)과 종료일을 입력 (종료일을 비우면 전 영업일까지)
3. 날짜를 비워 두고 실행하면 평소처럼 전일·전전일만 수집합니다
4. 실행 결과는 `docs/backfill_log.json` 에 남고, 사이트 상단 문구와 내려받는 엑셀도
   채운 기록을 바탕으로 전일·전전일 기준으로 다시 만들어집니다

## 과거 자료·틀린 값 직접 넣기
`docs/history.json` 은 직접 고치지 말고 아래 파일에 적어 주세요. 다음 실행 때 기록에 합쳐지고,
**직접 넣은 값이 자동 수집 값보다 우선**합니다(밀린 날짜 채우기를 다시 돌려도 유지).
- `manual_data.csv` : 지표 값. 한 줄에 `날짜,항목,값` (예 `2026-07-31,CP(91일),3.80`)
  - 항목: 환율, KOSPI, 다우지수, CP(91일), CD(91일), IRS(3Y), CRS(3년)
  - Term SOFR는 저장소가 공개라서 적지 마세요
- `manual_bond/날짜.txt` : 채권 표. 원본 엑셀 E2:T44를 복사해 붙여넣기 (예 `manual_bond/2026-07-31.txt`)
- 형식이 틀린 줄이나 파일은 건너뛰고, 사이트 상단에 `[수동입력]` 문구로 이유가 표시됩니다.
- 저장소 화면에서 파일을 열고 연필(✏️) 버튼으로 고친 뒤 `Commit changes` 하면 됩니다.
- 기록은 최근 1000영업일(약 4년)까지만 보관합니다.

## 문제가 생기면
- 사이트 상단 빨간 박스에 **수집에 실패한 항목과 사유**가 표시됩니다. 그 문구를 그대로 복사해서 알려 주세요.
- 자세한 로그: `Actions` 탭 → 해당 실행 클릭 → `Collect data and build site files` 펼치기
- 실패한 순간의 화면 캡처(png)와 페이지 내용(txt)이 `docs/debug/` 에 저장됩니다.
- 예약 실행은 GitHub 사정으로 몇 분~1시간 늦어질 수 있고, 저장소에 오래 활동이 없으면 멈출 수 있습니다.
  멈췄다면 `Actions` 탭에서 다시 켜 주세요.
- 이 저장소와 사이트는 누구나 볼 수 있습니다. 회사 서식이나 내부 정보는 올리지 마세요.
