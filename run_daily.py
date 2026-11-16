# -*- coding: utf-8 -*-
"""
GitHub Actions에서 매일 실행되는 스크립트 (run_daily.py).
수집한 값을 docs/data.json 과 docs/daily_data.xlsx 로 저장하고,
docs/index.html 이 data.json 을 읽어 화면에 보여줍니다.

- 항목별로 실패해도 나머지는 계속 진행하고, 실패 내용은 errors 에 기록되어 사이트 상단에 표시됩니다.
- 표(43행 x 16열)는 모양이 맞을 때만 '정상'으로 게시합니다.
"""

import json
import math
import re
import datetime as dt
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl
from openpyxl.styles import Font, PatternFill

from collectors import (
    get_holiday_set_from_csv,
    get_prev_business_day,
    scrape_kofia_bond_matrix,
    get_cp91,
    get_cd91,
    scrape_kmb_irs_crs,
    scrape_naver_kospi,
    scrape_naver_dow,
    scrape_naver_fx,
    scrape_term_sofr,
    make_driver,
)

KST = ZoneInfo("Asia/Seoul")
OUT = Path("docs")
MATURITIES = ["3월", "6월", "9월", "1년", "1년6월", "2년", "2년6월", "3년",
              "4년", "5년", "7년", "10년", "15년", "20년", "30년", "50년"]
REF_LABELS = ["환율", "KOSPI", "다우지수", "CP(91일)", "CD(91일)", "IRS(3Y)", "CRS(3년)", "Term SOFR"]


def clean(v):
    """NaN/None -> None, 숫자는 float 로"""
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(f) else f


DEBUG_DIR = OUT / "debug"


def _dump_frames(driver, f, depth=0, prefix=""):
    """현재 문서 안의 프레임들을 순서대로 들어가 소스를 파일에 적는다 (최대 3단계)."""
    if depth >= 3:
        return
    frames = driver.find_elements("css selector", "iframe, frame")
    for i in range(len(frames)):
        try:
            driver.switch_to.frame(i)
        except Exception:  # noqa: BLE001
            continue
        try:
            src = driver.page_source or ""
            f.write(f"\n\n===== 프레임 {prefix}{i} | url={driver.current_url} | 길이={len(src)} =====\n")
            f.write(src[:60000])
            _dump_frames(driver, f, depth + 1, f"{prefix}{i}.")
        finally:
            driver.switch_to.parent_frame()


def save_debug(driver, dbg_id):
    """실패한 순간의 화면(png)과 페이지 내용(txt)을 docs/debug 에 저장 (원인 파악용)."""
    try:
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        driver.switch_to.default_content()
        driver.save_screenshot(str(DEBUG_DIR / f"{dbg_id}.png"))
        src = driver.page_source or ""
        with open(DEBUG_DIR / f"{dbg_id}.txt", "w", encoding="utf-8") as f:
            f.write(f"url: {driver.current_url}\ntitle: {driver.title}\n")
            f.write(f"page_source_length: {len(src)}\n\n")
            f.write(src[:60000])
            _dump_frames(driver, f)
        driver.switch_to.default_content()
    except Exception as e:  # noqa: BLE001
        print("debug 저장 실패:", e)


def collect_day(driver, target, cp_query_date, errors, tag, dbg_prefix=None):
    """하루치(전일 또는 전전일) 데이터를 수집. 실패 항목은 None 으로 두고 errors 에 기록."""
    ref = {k: None for k in REF_LABELS}
    matrix = None

    def run(name, fn, dbg_id=None, browser=True):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            msg = str(e).split("Stacktrace")[0].strip()
            errors.append(f"[{tag}] {name}: {type(e).__name__} {msg[:300]}")
            if browser and dbg_prefix and dbg_id:
                save_debug(driver, f"{dbg_prefix}_{dbg_id}")
            return None

    df = run("채권시가평가수익률", lambda: scrape_kofia_bond_matrix(driver, target), "bond")
    if df is not None:
        if df.shape == (43, 16):
            matrix = [[clean(v) for v in row] for row in df.values.tolist()]
        else:
            errors.append(f"[{tag}] 채권시가평가수익률: 표 크기가 예상(43x16)과 다름 {df.shape}")

    ref["CP(91일)"] = run("CP(91일)", lambda: get_cp91(driver, cp_query_date, target), "cp")
    ref["CD(91일)"] = run("CD(91일)", lambda: get_cd91(driver, target), "cd")

    irs_crs = run("IRS/CRS", lambda: scrape_kmb_irs_crs(driver, target), "kmb")
    if irs_crs:
        ref["IRS(3Y)"], ref["CRS(3년)"] = irs_crs

    ref["Term SOFR"] = run("Term SOFR", lambda: scrape_term_sofr(driver, target), browser=False)

    ref["환율"] = run("네이버 환율", lambda: scrape_naver_fx(target, driver), browser=False)
    ref["KOSPI"] = run("네이버 KOSPI", lambda: scrape_naver_kospi(target, driver), browser=False)
    ref["다우지수"] = run("네이버 다우지수", lambda: scrape_naver_dow(target, driver), browser=False)

    ref = {k: clean(v) for k, v in ref.items()}
    return {"date": target.isoformat(), "matrix": matrix, "ref": ref}


def write_excel(path, days, row_labels, generated_at, errors):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    bold = Font(bold=True)
    warn = PatternFill("solid", fgColor="FFF2CC")

    for sheet_name, key in (("전일", "prev"), ("전전일", "prevprev")):
        ws = wb.create_sheet(sheet_name)
        day = days[key]
        ws.cell(row=1, column=1, value="종류").font = bold
        ws.cell(row=1, column=2, value="종류명").font = bold
        ws.cell(row=1, column=3, value="신용등급").font = bold
        ws.cell(row=1, column=4, value="고시기관").font = bold
        for i, m in enumerate(MATURITIES):
            ws.cell(row=1, column=5 + i, value=m).font = bold  # E1:T1
        for r in range(43):
            for c in range(3):
                ws.cell(row=2 + r, column=1 + c, value=row_labels[r][c])
            ws.cell(row=2 + r, column=4, value="민평평균")
        if day["matrix"]:
            for r, row in enumerate(day["matrix"]):
                for c, v in enumerate(row):
                    ws.cell(row=2 + r, column=5 + c, value=v if v is not None else "-")  # E2:T44
        else:
            ws.cell(row=2, column=5, value="채권 표 수집 실패").fill = warn
        for i, label in enumerate(REF_LABELS):
            ws.cell(row=2 + i, column=21, value=label)  # U2:U9
            v = day["ref"][label]
            cell = ws.cell(row=2 + i, column=22, value=v)  # V2:V9
            if v is None:
                cell.fill = warn
        ws.cell(row=1, column=21, value=f"기준일 {day['date']}").font = bold

    info = wb.create_sheet("정보")
    info["A1"] = "생성 시각(KST)"
    info["B1"] = generated_at
    info["A2"] = "사용법"
    info["B2"] = "전일/전전일 시트의 E2:V44를 복사해서 원본 파일의 같은 시트 E2에 붙여넣기"
    for i, e in enumerate(errors):
        info.cell(row=4 + i, column=1, value="수집 실패")
        info.cell(row=4 + i, column=2, value=e)
    wb.save(path)


def main():
    now = dt.datetime.now(KST)
    today = now.date()
    holidays = get_holiday_set_from_csv("holidays.csv")
    prev = get_prev_business_day(today, holidays)
    prevprev = get_prev_business_day(prev, holidays)
    print(f"오늘 {today} / 전영업일 {prev} / 전전영업일 {prevprev}")

    row_labels = json.load(open("row_labels.json", encoding="utf-8"))
    errors = []

    driver = None
    try:
        driver = make_driver(headless=True)
    except Exception as e:  # noqa: BLE001
        errors.append(f"브라우저 시작 실패: {str(e)[:200]}")

    days = {}
    for key, target, cp_q in (("prev", prev, today), ("prevprev", prevprev, prev)):
        if driver is None:
            days[key] = {"date": target.isoformat(), "matrix": None,
                         "ref": {k: None for k in REF_LABELS}}
            continue
        days[key] = collect_day(driver, target, cp_q, errors, "전일" if key == "prev" else "전전일",
                                dbg_prefix="prev" if key == "prev" else None)
    if driver:
        driver.quit()

    OUT.mkdir(exist_ok=True)
    generated_at = now.strftime("%Y-%m-%d %H:%M")
    data = {
        "generated_at": generated_at,
        "today": today.isoformat(),
        "maturities": MATURITIES,
        "ref_labels": REF_LABELS,
        "row_labels": row_labels,
        "days": days,
        "errors": errors,
    }
    with open(OUT / "data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    write_excel(OUT / "daily_data.xlsx", days, row_labels, generated_at, errors)
    print(f"완료. 실패 항목 {len(errors)}건")
    for e in errors:
        print(" -", e)


if __name__ == "__main__":
    main()
