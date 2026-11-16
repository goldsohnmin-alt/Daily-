# -*- coding: utf-8 -*-
"""
사이트별 스크래핑 함수.
전부 '최선의 추정'으로 작성되었습니다. 실행 중 요소를 못 찾으면(TimeoutException 등)
F12 개발자도구로 실제 id/class를 확인해서 아래 SELECTOR 상수들을 고쳐주세요.
"""

import time
import datetime as dt
import re

import requests
import pandas as pd
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


# =====================================================================
# 공통: 브라우저, 영업일 계산
# =====================================================================

def make_driver(headless: bool = True):
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1400,1000")
    opts.add_argument("--lang=ko-KR")
    # 서버(리눅스) 실행용 옵션
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    try:
        # Selenium 4.6+ 내장 드라이버 관리자 사용
        driver = webdriver.Chrome(options=opts)
    except Exception:
        from webdriver_manager.chrome import ChromeDriverManager
        from selenium.webdriver.chrome.service import Service
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=opts)
    return driver


def get_holiday_set(wb) -> set:
    """워크북('공휴일' 시트)에서 공휴일 날짜 목록을 set(date)로 읽어옴 (업무망, win32com 워크북 객체용)"""
    ws = wb.Sheets("공휴일")
    holidays = set()
    r = 2
    while True:
        val = ws.Cells(r, 1).Value
        if val is None:
            break
        if isinstance(val, dt.datetime):
            holidays.add(val.date())
        r += 1
    return holidays


def get_holiday_set_from_csv(path="holidays.csv") -> set:
    """
    외부망 PC용: Excel/win32com 없이 별도 관리하는 holidays.csv에서 공휴일 목록을 읽음.
    csv 형식: 한 줄에 YYYY-MM-DD 날짜 하나씩 (헤더 줄 '공휴일날짜' 포함 가능).
    최초 1회는 기존 엑셀 '공휴일' 시트 내용을 그대로 복사해서 만들어두고,
    이후 연말에 다음 해 공휴일을 추가해주면 됨.
    """
    holidays = set()
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip().split(",")[0]
            if not line or not re.match(r"\d{4}-\d{2}-\d{2}", line):
                continue
            holidays.add(dt.datetime.strptime(line, "%Y-%m-%d").date())
    return holidays


def get_prev_business_day(base_date: dt.date, holidays: set) -> dt.date:
    d = base_date - dt.timedelta(days=1)
    while d.weekday() >= 5 or d in holidays:  # 5=토, 6=일
        d -= dt.timedelta(days=1)
    return d


def _date_str(d: dt.date, sep="-") -> str:
    return d.strftime(f"%Y{sep}%m{sep}%d")


# =====================================================================
# 1. KOFIA 채권시가평가수익률 (E2:T44 매트릭스)
#    화면: 시가평가 > 채권시가평가수익률 > 일자별 탭
#    URL은 고정 URL이 없다고 하셨으므로, 홈 -> 상단메뉴 클릭으로 진입합니다.
#    아래 URL은 사이트맵 상 가장 가능성 높은 경로를 추정한 것이므로,
#    안 열리면 실제 홈 URL로 바꾸고 메뉴 클릭 코드를 추가해야 합니다.
# =====================================================================

KOFIA_HOME = "https://www.kofiabond.or.kr"
# KOFIA는 전체가 하나의 도메인(SPA/WebSquare)으로 되어 있어 화면별 URL이 없음(사용자 확인 완료).
# 따라서 홈 접속 -> 상단메뉴/좌측메뉴/탭을 '보이는 텍스트' 기준으로 클릭해서 이동한다.


def _kofia_goto(driver, top_menu: str, sub_menu: str = None, tab: str = None, wait_sec: int = 20):
    """
    KOFIA 홈에 접속한 뒤 상단메뉴(top_menu) -> (선택)좌측 서브메뉴(sub_menu) -> (선택)탭(tab) 순으로 클릭.
    실제로는 페이지 이동 없이 AJAX로 내용만 바뀌는 구조이므로, 클릭 후 sleep으로 렌더링을 기다린다.
    """
    wait = WebDriverWait(driver, wait_sec)
    driver.get(KOFIA_HOME)
    time.sleep(2)

    top = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, top_menu)))
    top.click()
    time.sleep(2)

    if sub_menu:
        sub = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, sub_menu)))
        sub.click()
        time.sleep(2)

    if tab:
        tab_el = wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//*[contains(normalize-space(text()),'{tab}')]")
        ))
        tab_el.click()
        time.sleep(1.5)

    return wait


def scrape_kofia_bond_matrix(driver, target_date: dt.date) -> pd.DataFrame:
    """
    상단메뉴 '시가평가' -> 좌측 '채권시가평가수익률' (기본이 '일자별' 탭으로 열림, 스크린샷 확인).
    43행 x 16열(3월~50년) 표를 읽어와 DataFrame으로 반환.
    반환 DataFrame의 행 순서는 화면 표시 순서와 동일 (엑셀 '전일'!E2:T44 순서와 일치, 확인 완료).
    """
    wait = _kofia_goto(driver, top_menu="시가평가", sub_menu="채권시가평가수익률")

    # 날짜 입력 필드 - 스크린샷 상 '조회일' 옆 텍스트박스
    # TODO: 여러 input 중 실제 어떤 것인지 다를 수 있음. 첫 실행 시 headless=False로 눈으로 확인 필요.
    date_input = None
    for sel in ["input[name='searchDate']", "input.date-input", "input[type='text']"]:
        els = driver.find_elements(By.CSS_SELECTOR, sel)
        if els:
            date_input = els[0]
            break
    if date_input is None:
        raise RuntimeError("KOFIA 날짜 입력창을 찾지 못했습니다. F12로 확인 후 선택자를 수정하세요.")

    date_input.clear()
    date_input.send_keys(_date_str(target_date))

    # 5개 평가사 체크박스는 기본 체크되어 있는 것으로 화면에서 확인됨(스크린샷) -> 별도 조작 불필요
    # 혹시 안 되어 있다면 아래 라벨 텍스트로 체크박스를 찾아 클릭하는 로직 추가:
    # for label in ["나이스피앤아이", "한국자산평가", "KIS자산평가", "에프앤자산평가", "이지자산평가"]:
    #     cb = driver.find_element(By.XPATH, f"//label[contains(text(),'{label}')]/preceding-sibling::input")
    #     if not cb.is_selected():
    #         cb.click()

    # 조회 버튼 클릭
    search_btn = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//a[contains(text(),'조회')] | //button[contains(text(),'조회')]")
    ))
    search_btn.click()
    time.sleep(2)  # 렌더링 대기 (필요시 명시적 wait로 교체)

    # 결과 표 파싱
    tables = pd.read_html(driver.page_source)
    # 가장 열 개수가 많은(=본 표) 테이블을 자동 선택
    target_table = max(tables, key=lambda t: t.shape[1])

    # 헤더/합계 행 등을 제외하고 순수 데이터 43행만 추출해야 함.
    # 실제 파싱 후 target_table.shape 를 print해서 43행 16열(+종류/종류명/신용등급 열)인지 확인 필요.
    df = target_table.copy()

    # TODO: 화면 캡처 기준 컬럼 = [종류, 종류명, 신용등급, 3월,6월,9월,1년,1년6월,2년,2년6월,3년,4년,5년,7년,10년,15년,20년,30년,50년]
    # 앞의 종류/종류명/신용등급 3개 열을 제외한 만기 컬럼만 잘라내어 반환
    numeric_cols = df.columns[3:19]  # 필요시 인덱스 조정
    bond_df = df[numeric_cols].apply(pd.to_numeric, errors="coerce")

    if bond_df.shape[0] != 43 or bond_df.shape[1] != 16:
        print(f"[경고] 예상과 다른 shape: {bond_df.shape}. 표 파싱 로직을 다시 확인하세요.")

    return bond_df.reset_index(drop=True)


def scrape_kofia_cp(driver, query_date: dt.date) -> float:
    """
    상단메뉴 '채권금리' (기본이 '최종호가수익률' 탭으로 열림, 스크린샷 확인)에서
    query_date로 조회하면 '전일' 컬럼에 전영업일 CP(91일) 값이 뜸 (사용자 확인 완료).
    """
    wait = _kofia_goto(driver, top_menu="채권금리")

    date_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']")))
    date_input.clear()
    date_input.send_keys(_date_str(query_date))

    search_btn = driver.find_element(By.XPATH, "//a[contains(text(),'조회')] | //button[contains(text(),'조회')]")
    search_btn.click()
    time.sleep(2)

    soup = BeautifulSoup(driver.page_source, "lxml")
    # 'CP(91일)' 행을 찾아 '전일' 컬럼(스크린샷 상 4번째 값 컬럼) 값을 추출
    row = soup.find(string=re.compile("CP\\(91일\\)"))
    if row is None:
        raise RuntimeError("CP(91일) 행을 찾지 못했습니다.")
    tr = row.find_parent("tr")
    cells = [td.get_text(strip=True) for td in tr.find_all("td")]
    # 스크린샷 기준 컬럼 순서: 종류명, 잔존기간, 오전, 오후, 전일대비, 전일, 연중최고, 연중최저
    # '전일' 값은 인덱스 5 (0-based) 근처 - 실제 파싱 후 인덱스 확인 필요
    prev_val = float(cells[5].replace(",", ""))
    return prev_val


def scrape_kofia_cd(driver, target_date: dt.date) -> float:
    """
    KOFIA 상단메뉴 'CD수익률' > '일자별,기간별 CD수익률' 화면.
    표 구조(스크린샷 확인 완료): 거래일 | CD수익률 | 대표수익률(80~100일)
    조회기간은 기본이 최근 1년으로 넓게 잡혀 있어, 한 번 조회해서 나온 리스트에서
    target_date에 해당하는 행의 'CD수익률' 컬럼 값을 그대로 가져오면 됨.
    """
    wait = _kofia_goto(driver, top_menu="CD수익률", sub_menu="일자별,기간별 CD수익률")

    # '일자별,기간별' 탭이 기본 선택되어 있는 것으로 화면에서 확인됨 -> 별도 클릭 불필요
    # 조회기간(시작일~종료일) 입력창 - 기본값이 이미 넓은 범위(예: 1년)로 잡혀 있으므로
    # 굳이 재설정하지 않고 기본 조회 버튼만 눌러도 target_date가 포함될 가능성이 높음.
    try:
        search_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(),'조회')] | //button[contains(text(),'조회')]")
        ))
        search_btn.click()
        time.sleep(2)
    except Exception:
        pass  # 페이지가 처음부터 기본 데이터를 보여주는 경우

    tables = pd.read_html(driver.page_source)
    # '거래일' 컬럼이 있는 표를 자동 탐색
    df = None
    for t in tables:
        if any("거래일" in str(c) for c in t.columns):
            df = t
            break
    if df is None:
        raise RuntimeError("CD수익률 표를 찾지 못했습니다. (컬럼명 '거래일' 기준 탐색 실패)")

    date_col = [c for c in df.columns if "거래일" in str(c)][0]
    cd_col = [c for c in df.columns if str(c).strip() == "CD수익률"][0]

    df[date_col] = df[date_col].astype(str)
    target_str = target_date.strftime("%Y-%m-%d")
    match = df[df[date_col] == target_str]
    if match.empty:
        raise RuntimeError(f"CD수익률에서 {target_date} 데이터를 찾지 못했습니다. "
                            f"(조회기간 범위를 벗어났을 수 있음 - 필요시 시작일을 더 과거로 넓혀서 재조회)")
    val = str(match.iloc[0][cd_col]).strip()
    if val in ("-", ""):
        raise RuntimeError(f"{target_date}의 CD수익률 값이 '-'(휴장/미고시)입니다.")
    return float(val)


# =====================================================================
# 2. 한국자금중개 IRS / CRS (3년물)
# =====================================================================

KMB_URL = "https://www.kmbco.com/kor/rate/deri_rate.do"


def scrape_kmb_irs_crs(driver, target_date: dt.date):
    driver.get(KMB_URL)
    wait = WebDriverWait(driver, 20)

    def _get_value_for_tab(tab_label: str) -> float:
        tab = wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(),'{tab_label}')]")))
        tab.click()
        time.sleep(1)

        date_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']")))
        date_input.clear()
        date_input.send_keys(_date_str(target_date))

        # 날짜 입력 후 자동조회이거나 별도 조회버튼이 있을 수 있음
        try:
            btn = driver.find_element(By.XPATH, "//a[contains(text(),'조회')] | //button[contains(text(),'조회')]")
            btn.click()
        except Exception:
            date_input.send_keys("\n")
        time.sleep(2)

        tables = pd.read_html(driver.page_source)
        df = max(tables, key=lambda t: t.shape[0])
        # '3년' 만기 행 찾기 - 컬럼명에 '3년' 또는 행 라벨에 '3년' 포함
        for col in df.columns:
            if "3년" in str(col):
                return float(df[col].iloc[0])
        # 만기가 행 방향일 경우
        row = df[df.apply(lambda r: r.astype(str).str.contains("3년").any(), axis=1)]
        if not row.empty:
            return float(row.iloc[0, -1])
        raise RuntimeError(f"{tab_label} 3년물 값을 표에서 찾지 못했습니다.")

    irs = _get_value_for_tab("IRS")
    crs = _get_value_for_tab("CRS")
    return irs, crs


# =====================================================================
# 3. 네이버금융: KOSPI, 다우지수, 원/달러 매매기준율
# =====================================================================

def scrape_naver_kospi_dow_fx(target_date: dt.date):
    """
    target_date의 종가/기준율을 조회. 네이버금융은 '일별시세'에서 특정일자 값을 바로 조회하기보다
    최근 값 리스트를 가져와서 target_date에 해당하는 값을 찾는 방식이 안정적입니다.
    """
    headers = {"User-Agent": "Mozilla/5.0"}

    # --- KOSPI ---
    kospi_url = "https://finance.naver.com/sise/sise_index_day.naver?code=KOSPI"
    kospi_val = _naver_index_value_for_date(kospi_url, target_date, headers)

    # --- 다우지수 (world index) ---
    dow_url = "https://finance.naver.com/world/worldDayListJson.naver?symbol=DJI@DJI&fdtc=0"
    dow_val = _naver_world_index_value_for_date(dow_url, target_date, headers)

    # --- 원/달러 매매기준율 ---
    fx_url = "https://finance.naver.com/marketindex/exchangeDailyQuote.naver?marketindexCd=FX_USDKRW"
    fx_val = _naver_fx_value_for_date(fx_url, target_date, headers)

    return fx_val, kospi_val, dow_val


def _naver_index_value_for_date(url, target_date, headers):
    date_str = target_date.strftime("%Y.%m.%d")
    for page in range(1, 6):
        resp = requests.get(f"{url}&page={page}", headers=headers, timeout=10)
        soup = BeautifulSoup(resp.text, "lxml")
        rows = soup.select("table.type_1 tr")
        for row in rows:
            tds = row.find_all("td")
            if len(tds) >= 2 and tds[0].get_text(strip=True) == date_str:
                return float(tds[1].get_text(strip=True).replace(",", ""))
    raise RuntimeError(f"KOSPI에서 {target_date} 값을 찾지 못했습니다. (네이버 페이지 구조 변경 가능성)")


def _naver_world_index_value_for_date(url, target_date, headers):
    resp = requests.get(url, headers=headers, timeout=10)
    data = resp.json()
    date_str = target_date.strftime("%Y.%m.%d")
    for item in data:
        if item.get("xymd") == target_date.strftime("%Y%m%d") or item.get("date") == date_str:
            return float(str(item.get("clos", item.get("close", ""))).replace(",", ""))
    raise RuntimeError(f"다우지수에서 {target_date} 값을 찾지 못했습니다. (JSON 구조 확인 필요)")


def _naver_fx_value_for_date(url, target_date, headers):
    date_str = target_date.strftime("%Y.%m.%d")
    for page in range(1, 6):
        resp = requests.get(f"{url}&page={page}", headers=headers, timeout=10)
        soup = BeautifulSoup(resp.text, "lxml")
        rows = soup.select("table.tbl_exchange tr")
        for row in rows:
            tds = row.find_all("td")
            if len(tds) >= 2 and tds[0].get_text(strip=True) == date_str:
                return float(tds[1].get_text(strip=True).replace(",", ""))
    raise RuntimeError(f"환율에서 {target_date} 값을 찾지 못했습니다.")


# =====================================================================
# 4. CME Term SOFR (3개월물)
# =====================================================================

TERM_SOFR_URL = "https://www.cmegroup.com/market-data/cme-group-benchmark-administration/term-sofr.html"


def scrape_term_sofr(driver, target_date: dt.date) -> float:
    """
    CME 페이지는 라이선스/JS 렌더링 이슈로 접근이 까다로울 수 있습니다.
    우선 페이지에서 값 추출을 시도하고, 실패하면 수동 입력을 요청합니다.
    """
    try:
        driver.get(TERM_SOFR_URL)
        time.sleep(3)
        soup = BeautifulSoup(driver.page_source, "lxml")
        # TODO: 실제 페이지에서 '3 Month' 값이 표시되는 위치를 찾아 선택자 지정
        candidates = soup.find_all(string=re.compile(r"3[\s-]?Month", re.I))
        if candidates:
            # 이 부분은 실제 페이지 구조를 보고 정교화 필요
            parent = candidates[0].find_parent()
            text = parent.get_text(" ", strip=True)
            m = re.search(r"(\d+\.\d+)", text)
            if m:
                return float(m.group(1))
    except Exception as e:
        print(f"[Term SOFR 자동조회 실패] {e}")

    import sys
    if not sys.stdin or not sys.stdin.isatty():
        # 서버 실행 중에는 입력받을 수 없으므로 실패로 처리 (사이트에 '미수집'으로 표시됨)
        raise RuntimeError("Term SOFR 자동조회 실패")
    val = input(f"Term SOFR(3M, {target_date}) 값을 CME 사이트에서 확인하여 직접 입력하세요: ")
    return float(val)
