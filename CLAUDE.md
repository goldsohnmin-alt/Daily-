# CLAUDE.md

일일 금융지표 자동 수집 사이트. 평일 아침 GitHub Actions가 금리·지수·환율 등을 수집해
`docs/` 아래 JSON·엑셀을 만들고, GitHub Pages(`main` 브랜치 `/docs`)로 공개한다.
사용자는 사이트에서 엑셀을 받아 업무망 원본 엑셀에 붙여넣는다(README "매일 사용법" 참고).

## 실제로 쓰이는 파일 (현재 버전 = v6)
- `.github/workflows/daily.yml` : 실행 설정. cron `30 22 * * 0-4`(UTC) = 한국시간 월~금 07:30.
  수동 실행 시 `start_date`/`end_date`를 넣으면 밀린 날짜를 채운다(backfill).
- `run_daily_v6.py` : 진입점. 수집 → `docs/` 파일 생성. 워크플로가 이 파일을 직접 실행한다.
- `collectors_v6.py` : 사이트별 수집 로직(Selenium + requests/pandas/BeautifulSoup).
  KOFIA 채권 금리표(43행×16열)·CP·CD, 한국자금중개 IRS/CRS, 네이버금융 KOSPI·환율·다우, Term SOFR(global-rates.com, CME는 서버 접속을 막아서 사용하지 않음).
  `ECOS_API_KEY`가 있으면 한국은행 ECOS API를 대체 경로로 쓴다.
- `holidays.csv` : 공휴일 목록. 전 영업일 계산에 쓰인다. 매년 말 다음 해 날짜를 추가해야 한다.
- `row_labels.json` : 채권 표 43행의 종류/종류명/신용등급 이름.
- `manual_data.csv`, `manual_bond/YYYY-MM-DD.txt` : 사용자가 직접 넣는 과거 자료·정정 값.
  매 실행마다 `load_manual`/`apply_manual`(run_daily_v6.py)이 history.json 에 덮어쓴다(자동 수집보다 우선).
  잘못된 줄은 건너뛰고 `[수동입력]` 문구를 errors 에 남긴다. Term SOFR는 공개라서 받지 않는다.
- `docs/index.html` : 사이트 화면. `data.json`, `history.json`, `private.json`(선택)을 읽는다.

## 생성되는 파일 (`docs/`, 워크플로가 자동 커밋)
- `data.json` : 최신 수집 결과 + `errors`(사이트 상단 빨간 상자에 표시)
- `history.json` : 날짜별 공개 데이터 (채권 표 + Term SOFR 제외 지표)
- `private.json` : 날짜별 Term SOFR, `XLSX_PASSWORD`로 암호화
- `daily_data.xlsx` : 내려받는 엑셀 (`전일`/`전전일` 시트, E2:V44 영역)
- `backfill_log.json` : backfill 실행 기록
- `debug/` : 수집 실패 시 화면(png)과 페이지 내용(txt)

## Secrets (GitHub 저장소 설정)
- `ECOS_API_KEY` : 선택. 없으면 ECOS 경로는 건너뛴다.
- `XLSX_PASSWORD` : Term SOFR 암호화용. 바꾸면 기존 `private.json`을 보호하려고 저장을 건너뛴다.

## 지켜야 할 것
- 저장소와 사이트는 Public이다. 회사 서식, 내부 정보, 비밀값을 커밋하지 않는다.
- 항목 하나가 실패해도 전체 실행은 계속돼야 한다. 실패는 `errors`에 사람이 읽을 수 있는 문장으로 남긴다.
- 채권 표는 43×16 모양이 맞을 때만 정상으로 게시한다.
- `docs/` 안의 생성 파일은 워크플로가 매일 덮어쓰므로 손으로 고치지 않는다.

## 옛 버전 파일
옛 버전 파일(`build_site*.py`, `scrapers*.py`, `run_daily*.py`·`collectors*.py`의 v1~v5,
루트 `daily.yml`)은 2026-11-17에 삭제했다. 필요하면 git 기록에서 꺼낸다.
새 버전을 만들 때 파일을 복사해 `_v7`로 늘리지 말고 v6 파일을 직접 고친다.
버전을 바꾼다면 워크플로의 `run:` 줄도 함께 바꿔야 한다.

## 로컬 실행
```bash
pip install -r requirements.txt cryptography xlrd
python run_daily_v6.py
```
Chrome이 필요하다(Selenium). 결과는 `docs/`에 쓰인다.
