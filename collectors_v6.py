# -*- coding: utf-8 -*-
"""
사이트별 스크래핑 함수.
전부 '최선의 추정'으로 작성되었습니다. 실행 중 요소를 못 찾으면(TimeoutException 등)
F12 개발자도구로 실제 id/class를 확인해서 아래 SELECTOR 상수들을 고쳐주세요.
"""

import os
import json
import time
import datetime as dt
import re
from io import StringIO

import requests
import pandas as pd
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    TimeoutException, StaleElementReferenceException, WebDriverException,
)


# =====================================================================
# 공통: 브라우저, 영업일 계산
# =====================================================================

_LAST_DRIVER = None


def make_driver(headless: bool = True):
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1400,1000")
    opts.add_argument("--lang=ko-KR")
    opts.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36")
    # 서버(리눅스) 실행용 옵션
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.set_capability("goog:loggingPrefs", {"performance": "ALL"})  # 사이트가 주고받는 요청 기록(진단용)
    try:
        # Selenium 4.6+ 내장 드라이버 관리자 사용
        driver = webdriver.Chrome(options=opts)
    except Exception:
        from webdriver_manager.chrome import ChromeDriverManager
        from selenium.webdriver.chrome.service import Service
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=opts)
    try:
        driver.execute_cdp_cmd("Network.enable", {})
    except Exception:  # noqa: BLE001
        pass
    global _LAST_DRIVER
    _LAST_DRIVER = driver
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
# 프레임(액자) 안까지 뒤져서 요소를 찾는 도구
#   KOFIA 는 바깥이 빈 껍데기이고 실제 화면이 프레임 안에 들어 있다.
# =====================================================================

MATURITIES = ["3월", "6월", "9월", "1년", "1년6월", "2년", "2년6월", "3년",
              "4년", "5년", "7년", "10년", "15년", "20년", "30년", "50년"]


def _search(driver, locators, depth=0):
    """현재 문서 -> 하위 프레임 순서로 탐색. 찾으면 그 프레임 안에 머문 채 요소를 반환."""
    for by, val in locators:
        try:
            els = driver.find_elements(by, val)
        except WebDriverException:
            continue
        for el in els:
            try:
                if el.tag_name.lower() == "area" or el.is_displayed():
                    return el
            except StaleElementReferenceException:
                continue
    if depth >= 3:
        return None
    n = len(driver.find_elements(By.CSS_SELECTOR, "iframe, frame"))
    for i in range(n):
        try:
            driver.switch_to.frame(i)
        except WebDriverException:
            continue
        found = _search(driver, locators, depth + 1)
        if found is not None:
            return found
        driver.switch_to.parent_frame()
    return None


def _find_in_frames(driver, locators, timeout=20):
    end = time.time() + timeout
    while time.time() < end:
        driver.switch_to.default_content()
        el = _search(driver, locators)
        if el is not None:
            return el
        time.sleep(0.7)
    raise TimeoutException(f"요소를 찾지 못함: {locators[0][1][:80]}")


def _text_locators(t):
    return [
        (By.XPATH, f"//a[normalize-space(.)='{t}']"),
        (By.XPATH, f"//*[self::span or self::li or self::td or self::div or self::button][normalize-space(text())='{t}']"),
        (By.XPATH, f"//img[@alt='{t}' or @title='{t}']"),
        (By.XPATH, f"//area[@alt='{t}' or @title='{t}']"),
        (By.XPATH, f"//a[contains(normalize-space(.),'{t}')]"),
        (By.XPATH, f"//*[contains(@alt,'{t}') or contains(@title,'{t}')]"),
    ]


def _click(driver, el):
    try:
        el.click()
    except WebDriverException:
        driver.execute_script("arguments[0].click();", el)


def _click_text(driver, text, timeout=20):
    el = _find_in_frames(driver, _text_locators(text), timeout)
    _click(driver, el)
    time.sleep(2.5)


SEARCH_LOCATORS = [
    (By.XPATH, "//a[normalize-space(.)='조회']"),
    (By.XPATH, "//input[(@type='button' or @type='image' or @type='submit') and (@value='조회' or @alt='조회')]"),
    (By.XPATH, "//button[normalize-space(.)='조회']"),
    (By.XPATH, "//img[@alt='조회']"),
    (By.XPATH, "//a[contains(normalize-space(.),'조회')]"),
    (By.XPATH, "//*[contains(@class,'btn') and contains(normalize-space(.),'조회')]"),
]

def _walk_frames(driver, probe, depth=0):
    """현재 문서 -> 하위 프레임 순서로 probe(driver)를 돌려 처음 참인 결과를 반환(그 프레임 안에 머문다)."""
    r = probe(driver)
    if r:
        return r
    if depth >= 3:
        return None
    n = len(driver.find_elements(By.CSS_SELECTOR, "iframe, frame"))
    for i in range(n):
        try:
            driver.switch_to.frame(i)
        except WebDriverException:
            continue
        r = _walk_frames(driver, probe, depth + 1)
        if r:
            return r
        driver.switch_to.parent_frame()
    return None


def _wait_probe(driver, probe, timeout=20, what="요소"):
    end = time.time() + timeout
    while time.time() < end:
        driver.switch_to.default_content()
        r = _walk_frames(driver, probe)
        if r:
            return r
        time.sleep(0.7)
    raise TimeoutException(f"{what}를 찾지 못함")


_JS_FIND_DATE = """
var re = /^\\d{4}-\\d{2}-\\d{2}$/;
var els = document.querySelectorAll('input');
for (var i = 0; i < els.length; i++) {
  var e = els[i];
  if (re.test((e.value || '').trim()) && (e.offsetWidth || e.offsetHeight)) return e;
}
return null;
"""


def _date_probe(driver):
    try:
        return driver.execute_script(_JS_FIND_DATE)
    except WebDriverException:
        return None


_POPUP_LOG = []


def _reset_popups():
    del _POPUP_LOG[:]


def _raise_if_date_popup():
    """'YYYYMMDD 형식으로 입력' 안내창이 떴으면 날짜가 안 들어간 것이므로 조용히 넘어가지 않고 실패로 처리."""
    for t in _POPUP_LOG:
        if "YYYYMMDD" in t:
            raise RuntimeError("날짜 형식 안내창이 떴음: " + " ".join(t.split())[:90])


POPUP_XPATHS = (
    "//*[@id='popup_ok']",
    "//input[normalize-space(@value)='Yes' or normalize-space(@value)='확인' or normalize-space(@value)='OK']",
    "//*[self::button or self::span or self::a or self::div][normalize-space(text())='Yes' or normalize-space(text())='확인']",
)


def _dismiss_popups(driver):
    """'Yes'/'확인' 버튼이 있는 안내창이 떠 있으면 문구를 기록하고 닫는다 (없으면 아무 일도 안 함).
    이 사이트의 버튼은 value=' Yes '(앞뒤 공백)이고 id='popup_ok' 이다."""
    def probe(d):
        for xp in POPUP_XPATHS:
            try:
                for el in d.find_elements(By.XPATH, xp):
                    if el.is_displayed():
                        return el
            except WebDriverException:
                pass
        return None
    for _ in range(4):
        driver.switch_to.default_content()
        el = _walk_frames(driver, probe)
        if not el:
            break
        try:
            txt = driver.execute_script(
                "var e=arguments[0];for(var i=0;i<6&&e.parentElement;i++){e=e.parentElement;"
                "if((e.innerText||'').trim().length>12) return e.innerText;} return '';", el)
            _POPUP_LOG.append(txt or "")
        except WebDriverException:
            pass
        _click(driver, el)
        time.sleep(1)
    driver.switch_to.default_content()


def _read_val(driver, el):
    return (driver.execute_script("return arguments[0].value;", el) or "").strip()


# ---------- 진단 도구: 날짜 칸의 구조, 달력 팝업, 조회할 때 오가는 요청을 docs/debug 에 저장 ----------
_DIAG_DONE = set()
DEBUG_OUT = os.path.join("docs", "debug")

_JS_ELEMENTS = r"""
return Array.from(document.querySelectorAll('input,select,button,a,img,td,th,span,i')).filter(function(e){
  return (e.offsetWidth || e.offsetHeight) && ((e.innerText||'').trim().length<=12);
}).slice(0, 220).map(function(e){
  return [e.tagName, 'id='+e.id, 'class='+(e.className||'').toString().slice(0,50), 'type='+(e.type||''),
          'value='+(e.value||'').toString().slice(0,20), 'title='+(e.title||''), 'alt='+(e.alt||''),
          'onclick='+(e.getAttribute('onclick')||'').slice(0,50), 'text='+(e.innerText||'').trim().slice(0,12)].join(' | ');
});
"""


def _diag_write(name, text):
    try:
        os.makedirs(DEBUG_OUT, exist_ok=True)
        with open(os.path.join(DEBUG_OUT, name), "w", encoding="utf-8") as f:
            f.write(text)
    except Exception:  # noqa: BLE001
        pass


def _dump_date_input(driver, el):
    if "date_input" in _DIAG_DONE:
        return
    _DIAG_DONE.add("date_input")
    try:
        info = driver.execute_script(
            "var e=arguments[0],p=e.parentElement,g=p?p.parentElement:null;"
            "return {html:e.outerHTML.slice(0,700), readonly:e.readOnly, disabled:e.disabled, maxlength:e.maxLength,"
            "parent:p?p.outerHTML.slice(0,1500):'', grand:g?g.outerHTML.slice(0,2500):''};", el)
        _diag_write("date_input.txt", json.dumps(info, ensure_ascii=False, indent=1))
    except Exception as e:  # noqa: BLE001
        _diag_write("date_input.txt", f"실패 {type(e).__name__} {e}")


def _dump_calendar(driver):
    if "calendar" in _DIAG_DONE:
        return
    _DIAG_DONE.add("calendar")
    try:
        _diag_write("date_popup.txt", "\n".join(driver.execute_script(_JS_ELEMENTS) or []))
    except Exception as e:  # noqa: BLE001
        _diag_write("date_popup.txt", f"실패 {type(e).__name__} {e}")


def _net_mark(driver):
    """이전 기록을 비워 둔다(조회 버튼을 누른 직후의 요청만 남기려고)."""
    try:
        driver.get_log("performance")
    except Exception:  # noqa: BLE001
        pass


def _net_save_once(driver, name):
    if name in _DIAG_DONE:
        return
    _DIAG_DONE.add(name)
    lines = []
    try:
        raw = driver.get_log("performance")
        reqs = []
        for entry in raw:
            msg = json.loads(entry["message"])["message"]
            if msg.get("method") != "Network.requestWillBeSent":
                continue
            p = msg["params"]
            r = p["request"]
            if re.search(r"\.(js|css|png|gif|jpe?g|ico|woff2?|svg)(\?|$)", r["url"], re.I):
                continue
            reqs.append((p["requestId"], r["method"], r["url"], r.get("postData", "")))
        for rid, method, url, post in reqs[-10:]:
            lines.append(f"{method} {url}\n  본문: {post[:1500]}")
            try:
                body = driver.execute_cdp_cmd("Network.getResponseBody", {"requestId": rid}).get("body", "")
                lines.append("  응답: " + body[:2500].replace("\n", " "))
            except Exception as e:  # noqa: BLE001
                lines.append(f"  응답: (읽기 실패 {type(e).__name__})")
    except Exception as e:  # noqa: BLE001
        lines.append(f"기록 읽기 실패 {type(e).__name__} {e}")
    _diag_write(f"net_{name}.txt", "\n".join(lines) or "(요청 없음)")


def _diag_search_capture(driver):
    """날짜 입력에 실패했을 때: 기본 날짜로라도 조회해 보고 그때 오가는 요청을 기록한다."""
    if "bond_default" in _DIAG_DONE:
        return
    try:
        _check_all_boxes(driver)
        _net_mark(driver)
        _press_search(driver)
        _dismiss_popups(driver)
        _net_save_once(driver, "bond_default")
    except Exception:  # noqa: BLE001
        pass
    finally:
        driver.switch_to.default_content()


# ---------- 날짜 칸 조작: (1) 타이핑 (2) 천천히 타이핑 (3) 달력 버튼 ----------
_JS_WS_SET = r"""
var el = arguments[0], v = arguments[1];
var id = (el.id || '').replace(/_input$/, '');
var comp = null;
try { comp = WebSquare.util.getComponentById(id); } catch (e) {}
if (!comp) { try { comp = $p.getComponentById(id); } catch (e) {} }
if (!comp) { try { comp = $p.getComponentById('mf_wfm_contents_' + id); } catch (e) {} }
if (!comp) return 'nocomp';
try { comp.setValue(v); } catch (e) { return 'seterr:' + e; }
try { return String(comp.getValue()); } catch (e) { return 'ok'; }
"""


def _calendar_pick(driver, el, value):
    """입력칸 옆의 달력 버튼을 눌러 연/월/일을 고른다(구조를 모르는 상태의 최선 시도 + 진단 기록)."""
    y, m, d = (int(x) for x in value.split("-"))
    btn = driver.execute_script(
        "var e=arguments[0], p=e.parentElement;"
        "for (var i=0;i<3 && p;i++, p=p.parentElement){"
        "  var c=p.querySelectorAll('img,button,a,span,i,input[type=button],input[type=image]');"
        "  for (var k=0;k<c.length;k++){var x=c[k]; if(x===e) continue; if(x.offsetWidth&&x.offsetHeight) return x;}}"
        "return null;", el)
    if not btn:
        raise RuntimeError("달력 버튼을 못 찾음")
    _click(driver, btn)
    time.sleep(1)
    _dump_calendar(driver)
    # 연/월 선택칸(select)이 있으면 맞춘다
    driver.execute_script(
        "var y=arguments[0], m=arguments[1];"
        "document.querySelectorAll('select').forEach(function(s){"
        "  if(!(s.offsetWidth||s.offsetHeight)) return;"
        "  Array.from(s.options).forEach(function(o){"
        "    var t=(o.text||'')+' '+(o.value||'');"
        "    var isYear = /^\\s*\\d{4}/.test(o.text||'') || /^\\d{4}$/.test(o.value||'');"
        "    if(isYear && String(parseInt(t))===String(y)){ s.value=o.value; s.dispatchEvent(new Event('change',{bubbles:true})); }"
        "    else if(!isYear && s.options.length>=12 && parseInt(o.value||o.text,10)===m){ s.value=o.value; s.dispatchEvent(new Event('change',{bubbles:true})); }"
        "  });"
        "});", y, m)
    time.sleep(0.8)
    # 날짜 칸(숫자 d) 클릭 - 지난달/다음달 흐린 칸은 제외
    clicked = driver.execute_script(
        "var d=String(arguments[0]); var cand=Array.from(document.querySelectorAll('td,a,span,button'))"
        ".filter(function(x){return (x.offsetWidth||x.offsetHeight) && (x.innerText||'').trim()===d;});"
        "cand=cand.filter(function(x){var c=((x.className||'')+' '+(x.parentElement?x.parentElement.className:'')).toLowerCase();"
        "  return !/(other|prev|next|dim|disabled|out|gray)/.test(c);});"
        "if(!cand.length) return false; cand[0].click(); return true;", d)
    if not clicked:
        raise RuntimeError("달력에서 날짜 칸을 못 찾음")
    time.sleep(1)


def _set_first_date(driver, value, alt=False):
    """조회일 칸을 value(YYYY-MM-DD)로 맞춘다. 사이트가 날짜를 정상 형식으로 받아들인 경우에만 성공으로 본다
    (다른 날짜 값이 조용히 섞이는 걸 막기 위해). 순서: 부품 지정 -> 타이핑 -> 천천히 타이핑 -> 달력 버튼.
    달력을 건드리면 칸이 새로 그려져 이전 참조가 사라질 수 있어, 읽을 때마다 필요하면 다시 찾는다."""
    digits = value.replace("-", "")
    state = {"el": _wait_probe(driver, _date_probe, 25, "날짜 입력칸")}
    tried = []

    def get_el():
        try:
            state["el"].is_displayed()
        except WebDriverException:
            state["el"] = _wait_probe(driver, _date_probe, 8, "날짜 입력칸")
        return state["el"]

    def cur():
        for _ in range(3):
            try:
                return _read_val(driver, get_el())
            except (WebDriverException, TimeoutException):
                time.sleep(0.5)
        return ""

    if cur() == value:
        return
    try:
        _dump_date_input(driver, get_el())
    except (WebDriverException, TimeoutException):
        pass

    def attempt(name, fn):
        try:
            fn()
        except (WebDriverException, TimeoutException, RuntimeError) as e:
            tried.append(f"{name}:{type(e).__name__}:{str(e)[:30]}")
        time.sleep(0.8)
        if cur() == value:
            return True
        tried.append(f"{name}→'{cur()}'")
        return False

    def m_component():
        res = driver.execute_script(_JS_WS_SET, get_el(), digits)
        if res == "nocomp" or str(res).startswith("seterr"):
            tried.append(f"부품:{res}")

    def m_type():
        e = get_el()
        e.click()
        e.send_keys(Keys.CONTROL, "a")
        e.send_keys(digits)
        e.send_keys(Keys.TAB)
        driver.execute_script("arguments[0].blur();", get_el())

    def m_type_slow():
        e = get_el()
        e.click()
        e.send_keys(Keys.CONTROL, "a")
        e.send_keys(Keys.BACKSPACE)
        for ch in digits:
            get_el().send_keys(ch)
            time.sleep(0.15)
        get_el().send_keys(Keys.ENTER)
        get_el().send_keys(Keys.TAB)

    def m_calendar():
        _calendar_pick(driver, get_el(), value)

    for name, fn in (("부품지정", m_component), ("타이핑", m_type), ("천천히", m_type_slow), ("달력", m_calendar)):
        if attempt(name, fn):
            return
    raise RuntimeError(f"날짜 입력 실패(칸에 남은 값 '{cur()}', 원하는 값 {value}; 시도 {tried})")


def _verify_date_after_search(driver, value):
    """조회 뒤에도 조회일 칸이 원하는 날짜로 남아 있는지 확인(다른 날짜 값이 조용히 섞이는 걸 막는다)."""
    try:
        el = _wait_probe(driver, _date_probe, 8, "날짜 입력칸")
        cur = _read_val(driver, el)
    except TimeoutException:
        return
    if cur.replace("-", "") != value.replace("-", ""):
        raise RuntimeError(f"조회 뒤 조회일이 '{cur}' 로 바뀌어 있음(원하는 값 {value})")


def _check_all_boxes(driver):
    """현재 프레임에서 보이는 체크박스를 모두 체크한다 (평가사 5개)."""
    boxes = driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
    vis = [b for b in boxes if b.is_displayed()] or boxes
    for b in vis:
        try:
            if not b.is_selected():
                _click(driver, b)
        except StaleElementReferenceException:
            pass
    return len(vis)


def _press_search(driver):
    el = _find_in_frames(driver, SEARCH_LOCATORS, 20)
    _click(driver, el)
    time.sleep(3)


# =====================================================================
# 1. KOFIA 채권정보센터
# =====================================================================

KOFIA_HOME = "https://www.kofiabond.or.kr"


def _kofia_goto(driver, top_menu, sub_menu=None):
    driver.get(KOFIA_HOME)
    time.sleep(3)
    _click_text(driver, top_menu, timeout=25)
    if sub_menu:
        try:
            _click_text(driver, sub_menu, timeout=10)
        except TimeoutException:
            pass  # 상단메뉴를 누르면 이미 해당 화면이 열려 있는 경우


def _norm(c):
    return re.sub(r"\s+", "", str(c))


def _flat_columns(df):
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [_norm(c[-1]) for c in df.columns]
    else:
        df.columns = [_norm(c) for c in df.columns]
    return df


def parse_bond_table(html):
    """채권시가평가수익률 표(종류/종류명/신용등급 + 3월~50년)에서 만기 16열만 뽑는다."""
    best = None
    for t in pd.read_html(StringIO(html)):
        t = _flat_columns(t.copy())
        if sum(1 for c in t.columns if c in MATURITIES) >= 14:
            best = t
            break
    if best is None:
        raise RuntimeError("채권 표(3월~50년 열)를 찾지 못함")
    cols = [m for m in MATURITIES if m in best.columns]
    df = best[cols].apply(pd.to_numeric, errors="coerce")
    for m in MATURITIES:
        if m not in df.columns:
            df[m] = float("nan")
    return df[MATURITIES].reset_index(drop=True)


def _bond_once(driver, target_date, alt):
    _reset_popups()
    _kofia_goto(driver, "시가평가", "채권시가평가수익률")
    time.sleep(2)
    _dismiss_popups(driver)
    try:
        _set_first_date(driver, _date_str(target_date), alt=alt)
    except (RuntimeError, WebDriverException, TimeoutException):
        _diag_search_capture(driver)  # 다음 수정을 위해 기본 날짜로 조회한 요청을 기록
        raise
    n = _check_all_boxes(driver)
    if n < 2:
        raise RuntimeError(f"평가사 체크박스를 못 찾음(찾은 수 {n})")
    _net_mark(driver)
    _press_search(driver)
    _net_save_once(driver, "bond")
    _dismiss_popups(driver)
    _raise_if_date_popup()
    _find_in_frames(driver, [(By.XPATH, "//th[contains(.,'3월')]")], 25)
    _verify_date_after_search(driver, _date_str(target_date))
    return parse_bond_table(driver.page_source)


def scrape_kofia_bond_matrix(driver, target_date):
    """상단 '시가평가' -> 채권시가평가수익률. 평가사 5개 체크 후 조회 -> 5개사 평균 표.
    첫 시도가 실패하면 날짜를 넣는 방식을 바꿔 한 번 더 시도한다."""
    last = None
    for attempt in (1, 2):
        try:
            return _bond_once(driver, target_date, alt=(attempt == 2))
        except Exception as e:  # noqa: BLE001
            last = e
            driver.switch_to.default_content()
    raise last


def _cp_once(driver, query_date, alt):
    _reset_popups()
    _kofia_goto(driver, "채권금리")
    time.sleep(2)
    _dismiss_popups(driver)
    _set_first_date(driver, _date_str(query_date), alt=alt)
    _press_search(driver)
    _dismiss_popups(driver)
    _raise_if_date_popup()
    row = _find_in_frames(driver, [(By.XPATH, "//tr[td[contains(normalize-space(.),'CP(91일)')]]")], 20)
    _verify_date_after_search(driver, _date_str(query_date))
    row = _find_in_frames(driver, [(By.XPATH, "//tr[td[contains(normalize-space(.),'CP(91일)')]]")], 20)
    cells = [td.text.strip() for td in row.find_elements(By.TAG_NAME, "td")]
    # 열 순서: 종류명, 잔존기간, 오전, 오후, 전일대비, 전일, 연중최고, 연중최저 (+ 끝에 빈 칸이 붙기도 함)
    # 끝의 빈 칸 때문에 뒤에서 세면 밀리므로 앞에서 센다 -> 4번째(cells[3])가 당일 '오후'
    if len(cells) < 4 or "CP" not in cells[0]:
        raise RuntimeError(f"CP(91일) 행 모양이 이상함: {cells}")
    val = cells[3].replace(",", "")
    if not val or val == "-":
        raise RuntimeError(f"CP(91일) {query_date} 오후 값이 아직 없음: {cells}")
    return float(val)


def scrape_kofia_cp(driver, query_date):
    """상단 '채권금리'(최종호가수익률)에서 query_date(=대상 날짜) 로 조회 -> CP(91일)의 당일 '오후' 값.
    다음 날짜로 조회해 '전일' 칸을 읽던 방식은 아침 실행 때 그날 자료(11:30/16:00 게시)가
    아직 없어 표가 비므로 쓰지 않는다.
    실패하면 날짜 입력 방식을 바꿔 한 번 더 시도한다."""
    last = None
    for attempt in (1, 2):
        try:
            return _cp_once(driver, query_date, alt=(attempt == 2))
        except Exception as e:  # noqa: BLE001
            last = e
            driver.switch_to.default_content()
    raise last


def scrape_kofia_cd(driver, target_date):
    """상단 'CD수익률' -> 일자별,기간별 CD수익률 표에서 target_date 행의 CD수익률."""
    _kofia_goto(driver, "CD수익률", "일자별,기간별 CD수익률")
    try:
        _press_search(driver)
    except TimeoutException:
        pass
    _find_in_frames(driver, [(By.XPATH, "//th[contains(.,'거래일')]")], 20)
    df = None
    for t in pd.read_html(StringIO(driver.page_source)):
        t = _flat_columns(t.copy())
        if any("거래일" in c for c in t.columns):
            df = t
            break
    if df is None:
        raise RuntimeError("CD수익률 표를 찾지 못함")
    date_col = [c for c in df.columns if "거래일" in c][0]
    cd_col = [c for c in df.columns if c == "CD수익률"][0]
    key = target_date.strftime("%Y-%m-%d")
    match = df[df[date_col].astype(str).str.strip() == key]
    if match.empty:
        raise RuntimeError(f"CD수익률 표에 {key} 행이 없음")
    val = str(match.iloc[0][cd_col]).strip()
    if val in ("-", "", "nan"):
        raise RuntimeError(f"{key} CD수익률이 '-'(미고시)")
    return float(val)


# =====================================================================
# 2. 한국자금중개 IRS / CRS (3년물)
#    페이지에 최근 한 달치 일별 표(전송일 / 1Y ... 10Y)가 이미 들어 있어서
#    클릭 없이 날짜 행만 읽으면 된다. CRS 는 탭 버튼만 한 번 누른다.
# =====================================================================

KMB_URL = "https://www.kmbco.com/kor/rate/deri_rate.do"


def parse_kmb_history(html, target_date, tenor="3Y"):
    key = target_date.strftime("%Y/%m/%d")
    seen = []
    for t in pd.read_html(StringIO(html)):
        t = _flat_columns(t.copy())
        first = t.columns[0]
        if "전송일" not in first or tenor not in t.columns:
            continue
        dates = t[first].astype(str).str.strip()
        seen = dates.tolist()[:3]
        row = t[dates == key]
        if not row.empty:
            return float(row.iloc[0][tenor])
    raise RuntimeError(f"한국자금중개 표에 {key} 행이 없음 (보이는 날짜 앞 3건={seen})")


def _kmb_active_tab(driver):
    try:
        els = driver.find_elements(By.CSS_SELECTOR, "nav.chartTab button.on")
        return els[0].text.strip() if els else None
    except WebDriverException:
        return None


KMB_EXCEL = "https://www.kmbco.com/kor/rate/deri_rate_excel.do"


def _kmb_excel_value(target_date, dtype, tenor="3Y"):
    """화면 표(최근 9영업일)에 없는 날짜용: 기간을 지정해 받는 엑셀 다운로드 주소에서 값을 읽는다."""
    s = (target_date - dt.timedelta(days=7)).strftime("%Y-%m-%d")
    e = (target_date + dt.timedelta(days=3)).strftime("%Y-%m-%d")
    r = requests.get(f"{KMB_EXCEL}?sDate={s}&eDate={e}&dType={dtype}", headers={
        "User-Agent": NAVER_HEADERS["User-Agent"], "Referer": KMB_URL}, timeout=40)
    if r.status_code != 200:
        raise RuntimeError(f"엑셀 주소 http {r.status_code}")
    tables = []
    try:
        tables = pd.read_html(StringIO(r.content.decode("utf-8", errors="ignore")))
    except Exception:  # noqa: BLE001
        try:
            import io
            tables = [pd.read_excel(io.BytesIO(r.content))]
        except Exception:  # noqa: BLE001
            raise RuntimeError(f"엑셀 주소 내용을 못 읽음(content-type={r.headers.get('content-type')}, 앞부분={r.text[:60]!r})")
    want = target_date.strftime("%Y%m%d")
    for t in tables:
        t = _flat_columns(t.copy())
        if tenor not in t.columns:
            continue
        first = t.columns[0]
        digits = t[first].astype(str).str.replace(r"\D", "", regex=True)
        row = t[digits == want]
        if not row.empty:
            return float(row.iloc[0][tenor])
    raise RuntimeError(f"엑셀 주소 표에 {target_date} 행이 없음")


def scrape_kmb_irs_crs(driver, target_date):
    driver.get(KMB_URL)
    WebDriverWait(driver, 30).until(lambda d: _kmb_active_tab(d) is not None)
    time.sleep(1)
    if _kmb_active_tab(driver) != "IRS":
        driver.execute_script("selType('IRS');")
        WebDriverWait(driver, 20).until(lambda d: _kmb_active_tab(d) == "IRS")
        time.sleep(2)
    try:
        irs = parse_kmb_history(driver.page_source, target_date)
    except RuntimeError as e1:
        try:
            irs = _kmb_excel_value(target_date, "IRS")
        except Exception as e2:  # noqa: BLE001
            raise RuntimeError(f"IRS 화면 표 실패({str(e1)[:80]}) / 엑셀 주소도 실패({str(e2)[:120]})")

    btn = driver.find_element(By.XPATH, "//nav[contains(@class,'chartTab')]//button[normalize-space(.)='CRS']")
    _click(driver, btn)
    WebDriverWait(driver, 20).until(lambda d: _kmb_active_tab(d) == "CRS")
    time.sleep(2)
    try:
        crs = parse_kmb_history(driver.page_source, target_date)
    except RuntimeError as e1:
        try:
            crs = _kmb_excel_value(target_date, "CRS")
        except Exception as e2:  # noqa: BLE001
            raise RuntimeError(f"CRS 화면 표 실패({str(e1)[:80]}) / 엑셀 주소도 실패({str(e2)[:120]})")
    return irs, crs


# =====================================================================
# 한국은행 ECOS (공식 통계 API) - 무료 인증키가 있을 때만 사용, 없으면 건너뜀
#   환경변수 ECOS_API_KEY 로 키를 받는다 (GitHub Secrets 에 저장).
#   표/항목 코드는 실행 때 항목 목록을 조회해서 이름으로 찾는다(코드 오류 방지).
# =====================================================================

ECOS_BASE = "https://ecos.bok.or.kr/api"


def _ecos_key():
    k = os.environ.get("ECOS_API_KEY", "").strip()
    if not k:
        raise RuntimeError("ECOS 인증키 없음(선택 기능)")
    return k


def _ecos_json(url):
    r = requests.get(url, timeout=25)
    if r.status_code != 200:
        raise RuntimeError(f"ECOS http {r.status_code}")
    return r.json()


def _ecos_value(stat_code, name_pred, target_date, label):
    """stat_code 표에서 name_pred(항목이름)이 참인 항목의 target_date 일별 값."""
    key = _ecos_key()
    items = _ecos_json(f"{ECOS_BASE}/StatisticItemList/{key}/json/kr/1/300/{stat_code}")
    rows = items.get("StatisticItemList", {}).get("row", [])
    if not rows:
        raise RuntimeError(f"ECOS 항목목록 없음 {str(items)[:100]}")
    hit = next((r for r in rows if name_pred(str(r.get("ITEM_NAME", "")))), None)
    if hit is None:
        names = [str(r.get("ITEM_NAME", "")) for r in rows][:10]
        raise RuntimeError(f"ECOS {label} 항목을 못 찾음, 목록 앞부분={names}")
    ymd = target_date.strftime("%Y%m%d")
    url = f"{ECOS_BASE}/StatisticSearch/{key}/json/kr/1/5/{stat_code}/D/{ymd}/{ymd}/{hit['ITEM_CODE']}"
    data = _ecos_json(url)
    got = data.get("StatisticSearch", {}).get("row", [])
    if not got:
        msg = data.get("RESULT", {}).get("MESSAGE", str(data)[:100])
        raise RuntimeError(f"ECOS {label} {ymd} 값 없음({msg})")
    return _num(got[0]["DATA_VALUE"])


def ecos_cd91(target_date):
    return _ecos_value("817Y002", lambda n: n.strip().upper().startswith("CD") and "91" in n, target_date, "CD(91일)")


def ecos_cp91(target_date):
    return _ecos_value("817Y002", lambda n: n.strip().upper().startswith("CP") and "91" in n, target_date, "CP(91일)")


def ecos_kospi(target_date):
    return _ecos_value("802Y001", lambda n: n.strip().upper().replace(" ", "") in ("KOSPI", "KOSPI지수", "KOSPI종가"), target_date, "KOSPI")


def ecos_usdkrw(target_date):
    return _ecos_value("731Y001", lambda n: "미국달러" in n and "매매기준율" in n, target_date, "원/달러 매매기준율")


def get_cd91(driver, target_date):
    return _chain("CD(91일)", [
        ("ECOS", lambda: ecos_cd91(target_date)),
        ("KOFIA화면", lambda: scrape_kofia_cd(driver, target_date)),
    ])


def get_cp91(driver, target_date):
    return _chain("CP(91일)", [
        ("ECOS", lambda: ecos_cp91(target_date)),
        ("KOFIA화면", lambda: scrape_kofia_cp(driver, target_date)),
    ])


# =====================================================================
# 3. 네이버금융 (KOSPI / 다우지수 / 원달러 매매기준율)
#    옛 주소가 410(없는 주소)로 막혀서 여러 경로를 순서대로 시도한다.
# =====================================================================

NAVER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Referer": "https://finance.naver.com/",
    "Accept-Language": "ko-KR,ko;q=0.9",
}


def _num(text):
    return float(str(text).replace(",", "").strip())


def _chain(label, attempts):
    errs = []
    for name, fn in attempts:
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            errs.append(f"{name}: {type(e).__name__} {str(e).split('Stacktrace')[0].strip()[:110]}")
    raise RuntimeError(f"{label} 모든 경로 실패 | " + " | ".join(errs))


def _parse_legacy_table(html, date_str):
    soup = BeautifulSoup(html, "lxml")
    seen = []
    for row in soup.select("tr"):
        tds = row.find_all("td")
        if len(tds) >= 2 and re.match(r"\d{4}\.\d{2}\.\d{2}", tds[0].get_text(strip=True)):
            seen.append(tds[0].get_text(strip=True))
            if tds[0].get_text(strip=True) == date_str:
                return _num(tds[1].get_text(strip=True)), seen
    return None, seen


def _legacy_via_requests(url, target_date):
    date_str = target_date.strftime("%Y.%m.%d")
    status = None
    for page in range(1, 6):
        resp = requests.get(f"{url}&page={page}", headers=NAVER_HEADERS, timeout=15)
        status = resp.status_code
        if status != 200:
            raise RuntimeError(f"http {status}")
        val, _ = _parse_legacy_table(resp.content.decode(resp.apparent_encoding or "euc-kr", errors="ignore"), date_str)
        if val is not None:
            return val
    raise RuntimeError(f"{date_str} 행 없음")


def _legacy_via_browser(driver, url, target_date):
    if driver is None:
        raise RuntimeError("브라우저 없음")
    date_str = target_date.strftime("%Y.%m.%d")
    for page in range(1, 4):
        driver.switch_to.default_content()
        driver.get(f"{url}&page={page}")
        time.sleep(1.5)
        val, seen = _parse_legacy_table(driver.page_source, date_str)
        if val is not None:
            return val
    raise RuntimeError(f"{date_str} 행 없음 (읽은 날짜 {len(seen)}건)")


def _m_naver_json(url_tpl, target_date, pages=4):
    """url_tpl 안의 {page} 를 1,2,3.. 으로 바꿔 가며 target_date 행을 찾는다(쪽당 약 30거래일)."""
    key = target_date.strftime("%Y-%m-%d")
    seen = []
    for page in range(1, pages + 1):
        resp = requests.get(url_tpl.format(page=page), headers=NAVER_HEADERS, timeout=15)
        if resp.status_code != 200:
            raise RuntimeError(f"http {resp.status_code}")
        data = resp.json()
        if isinstance(data, dict):
            data = data.get("result", data.get("priceInfos", []))
        if not data:
            break
        for item in data:
            d = str(item.get("localTradedAt", item.get("tradedAt", "")))[:10]
            seen.append(d)
            if d == key:
                return _num(item.get("closePrice", item.get("close", "")))
    raise RuntimeError(f"{key} 항목 없음 (읽은 날짜 {len(seen)}건, 앞 3건={seen[:3]})")


def _yahoo_close(symbol, target_date):
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?range=3mo&interval=1d"
    resp = requests.get(url, headers={"User-Agent": NAVER_HEADERS["User-Agent"]}, timeout=15)
    if resp.status_code != 200:
        raise RuntimeError(f"http {resp.status_code}")
    r = resp.json()["chart"]["result"][0]
    off = r["meta"].get("gmtoffset", 0)
    closes = r["indicators"]["quote"][0]["close"]
    for ts, c in zip(r["timestamp"], closes):
        if c is not None and dt.datetime.utcfromtimestamp(ts + off).date() == target_date:
            return round(c, 2)
    raise RuntimeError(f"{target_date} 종가 없음")


def scrape_naver_kospi(target_date, driver=None):
    legacy = "https://finance.naver.com/sise/sise_index_day.naver?code=KOSPI"
    return _chain("KOSPI", [
        ("ECOS", lambda: ecos_kospi(target_date)),
        ("옛주소(requests)", lambda: _legacy_via_requests(legacy, target_date)),
        ("옛주소(브라우저)", lambda: _legacy_via_browser(driver, legacy, target_date)),
        ("네이버모바일", lambda: _m_naver_json("https://m.stock.naver.com/api/index/KOSPI/price?pageSize=30&page={page}", target_date)),
        ("야후", lambda: _yahoo_close("%5EKS11", target_date)),
    ])


def scrape_naver_fx(target_date, driver=None):
    legacy = "https://finance.naver.com/marketindex/exchangeDailyQuote.naver?marketindexCd=FX_USDKRW"
    return _chain("환율", [
        ("ECOS", lambda: ecos_usdkrw(target_date)),
        ("옛주소(requests)", lambda: _legacy_via_requests(legacy, target_date)),
        ("옛주소(브라우저)", lambda: _legacy_via_browser(driver, legacy, target_date)),
        ("네이버모바일", lambda: _m_naver_json("https://m.stock.naver.com/front-api/marketIndex/prices?category=exchange&reutersCode=FX_USDKRW&page={page}&pageSize=30", target_date)),
    ])


def scrape_naver_dow(target_date, driver=None):
    base = "https://finance.naver.com/world/worldDayListJson.naver?symbol=DJI@DJI&fdtc=0"
    ymd = target_date.strftime("%Y%m%d")

    def naver_world():
        seen, status = [], None
        for page in range(1, 4):
            resp = requests.get(f"{base}&page={page}", headers=NAVER_HEADERS, timeout=15)
            status = resp.status_code
            data = resp.json()
            for item in data:
                seen.append(str(item.get("xymd", "")))
                if str(item.get("xymd", "")) == ymd:
                    for key in ("close_price", "clos", "close", "last"):
                        if item.get(key) not in (None, ""):
                            return _num(item[key])
        raise RuntimeError(f"{ymd} 없음 (http={status}, 앞 3건={seen[:3]})")

    return _chain("다우", [("네이버해외", naver_world), ("야후", lambda: _yahoo_close("%5EDJI", target_date))])


def scrape_naver_kospi_dow_fx(target_date):
    """(로컬 실행용) 세 값을 한 번에 반환."""
    return scrape_naver_fx(target_date), scrape_naver_kospi(target_date), scrape_naver_dow(target_date)


# =====================================================================
# 4. Term SOFR (3개월)
#    CME 사이트가 클라우드 서버 접속을 차단(ERR_HTTP2_PROTOCOL_ERROR)해서 자동 수집 불가.
# =====================================================================

def scrape_term_sofr(driver, target_date):
    import sys
    if sys.stdin and sys.stdin.isatty():
        val = input(f"Term SOFR(3M, {target_date}) 값을 CME 사이트에서 확인해 직접 입력하세요: ")
        return float(val)
    raise RuntimeError("CME 사이트가 서버 접속을 차단해 자동 수집 불가 - 직접 입력 필요")


# =====================================================================
# 5. Term SOFR (3개월) - global-rates.com 의 공개 표
#    CME 사이트는 서버 접속을 막아서 이쪽을 쓴다. 값은 공개 페이지에 싣지 않고
#    비밀번호로 잠근 엑셀에만 저장한다(run_daily_v4.py 참고).
# =====================================================================

GLOBAL_RATES_URL = "https://www.global-rates.com/en/interest-rates/cme-term-sofr/2/term-sofr-interest-3-months/"


_GR_ROWS = None


def _parse_gr(text):
    rows = {}
    for d, v in re.findall(r"(\d{2}-\d{2}-\d{4})\s+(\d+\.\d+)\s*%", text):
        rows[dt.datetime.strptime(d, "%m-%d-%Y").date()] = float(v)
    return rows


def _gr_rows(driver=None):
    """global-rates 표를 한 번만 읽어 {날짜: 값} 으로 보관한다. 이 사이트는 인증서 체인이 불완전해서
    파이썬 요청이 인증서 오류로 막힐 수 있어, 그땐 실제 브라우저(크롬)로 대신 읽는다."""
    global _GR_ROWS
    if _GR_ROWS is not None:
        return _GR_ROWS
    text, err = None, None
    try:
        resp = requests.get(GLOBAL_RATES_URL, headers={
            "User-Agent": NAVER_HEADERS["User-Agent"], "Accept-Language": "en-US,en;q=0.9"}, timeout=25)
        if resp.status_code != 200:
            raise RuntimeError(f"http {resp.status_code}")
        text = BeautifulSoup(resp.text, "lxml").get_text(" ", strip=True)
    except Exception as e:  # noqa: BLE001
        err = e
    if text is None or not _parse_gr(text):
        # 이 사이트는 인증서가 정상이 아니어서(체인 불완전/오류) 검증이 막힌다. 공개된 숫자를 읽기만 하므로
        # 인증서 검증 없이 한 번 더 시도한다(값은 사람이 붙여넣을 때 눈으로 확인하는 전제).
        try:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            r2 = requests.get(GLOBAL_RATES_URL, headers={
                "User-Agent": NAVER_HEADERS["User-Agent"], "Accept-Language": "en-US,en;q=0.9"},
                timeout=25, verify=False)
            if r2.status_code == 200:
                t2 = BeautifulSoup(r2.text, "lxml").get_text(" ", strip=True)
                if _parse_gr(t2):
                    text = t2
        except Exception as e2:  # noqa: BLE001
            err = e2
    if text is None or not _parse_gr(text):
        drv = driver or _LAST_DRIVER
        if drv is None:
            raise RuntimeError(f"global-rates 접속 실패({type(err).__name__ if err else '표 없음'}), 브라우저도 없음")
        try:
            drv.switch_to.default_content()
            drv.get(GLOBAL_RATES_URL)
            time.sleep(4)
            text = BeautifulSoup(drv.page_source, "lxml").get_text(" ", strip=True)
            title = drv.title
        except Exception as e2:  # noqa: BLE001
            raise RuntimeError(f"global-rates 접속 실패(요청: {type(err).__name__ if err else '-'}, 브라우저: {type(e2).__name__})")
        if not _parse_gr(text):
            raise RuntimeError(f"global-rates 브라우저로도 표를 못 읽음(제목 '{title}', 앞부분 '{text[:60]}')")
    rows = _parse_gr(text)
    if not rows:
        raise RuntimeError(f"global-rates 표를 못 읽음(길이 {len(text)})")
    _GR_ROWS = rows
    return rows


def scrape_term_sofr_gr(target_date, driver=None):
    """(값, 안내문구 또는 None) 반환. 이 페이지는 최근 5영업일 정도만 보여 주므로 그보다 오래된 날짜는 실패한다.
    미국 휴일 등으로 그 날짜가 없으면 5일 이내 가장 가까운 이전 날짜 값을 쓰고 안내문구를 돌려준다."""
    rows = _gr_rows(driver)
    if target_date in rows:
        return rows[target_date], None
    earlier = [d for d in rows if d <= target_date and (target_date - d).days <= 5]
    if earlier:
        d = max(earlier)
        return rows[d], f"{target_date} 값이 없어 {d} 값을 사용"
    seen = [d.isoformat() for d in sorted(rows)[-3:]]
    raise RuntimeError(f"global-rates 에 {target_date} 값 없음(이 페이지는 최근 며칠만 보여 줌, 최근 날짜 {seen})")
