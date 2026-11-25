# -*- coding: utf-8 -*-
"""
GitHub Actions에서 매일 실행되는 스크립트 (run_daily.py).
수집한 값을 docs/data.json 과 docs/daily_data.xlsx 로 저장하고,
docs/index.html 이 data.json 을 읽어 화면에 보여줍니다.

- 항목별로 실패해도 나머지는 계속 진행하고, 실패 내용은 errors 에 기록되어 사이트 상단에 표시됩니다.
- 표(43행 x 16열)는 모양이 맞을 때만 '정상'으로 게시합니다.
"""

import base64
import copy
import json
import math
import os
import re
import datetime as dt
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl
from openpyxl.styles import Font, PatternFill

from collectors_v6 import (
    get_holiday_set_from_csv,
    get_prev_business_day,
    scrape_kofia_bond_matrix,
    get_cp91,
    get_cd91,
    scrape_kmb_irs_crs,
    scrape_naver_kospi,
    scrape_naver_dow,
    scrape_naver_fx,
    scrape_term_sofr_gr,
    TermSofrNotYet,
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


_JS_ELEMENT_MAP = r"""
return Array.from(document.querySelectorAll('input,select,button,a,img,area,th')).filter(function(e){
  return (e.offsetWidth || e.offsetHeight);
}).slice(0, 150).map(function(e){
  return [e.tagName, 'id=' + e.id, 'name=' + e.name, 'class=' + (e.className || '').toString().slice(0, 40),
          'type=' + (e.type || ''), 'value=' + (e.value || '').toString().slice(0, 20),
          'title=' + (e.title || ''), 'alt=' + (e.alt || ''),
          'onclick=' + (e.getAttribute('onclick') || '').slice(0, 40),
          'text=' + (e.innerText || '').trim().slice(0, 20),
          'checked=' + e.checked].join(' | ');
});
"""


def _dump_frames(driver, f, depth=0, prefix=""):
    """현재 문서 안의 프레임들을 순서대로 들어가 '보이는 요소 목록'과 소스 앞부분을 적는다 (최대 3단계)."""
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
            f.write(f"\n\n===== 프레임 {prefix}{i} | 길이={len(src)} =====\n")
            try:
                for line in driver.execute_script(_JS_ELEMENT_MAP) or []:
                    f.write("  " + line + "\n")
            except Exception as e:  # noqa: BLE001
                f.write(f"  (요소 목록 실패: {e})\n")
            f.write("--- 소스 앞부분 ---\n" + src[:3000])
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
            f.write(src[:3000])
            _dump_frames(driver, f)
        driver.switch_to.default_content()
    except Exception as e:  # noqa: BLE001
        print("debug 저장 실패:", e)


def collect_day(driver, target, errors, tag, dbg_prefix=None, quiet_sofr=False):
    """하루치(전일 또는 전전일) 데이터를 수집. 실패 항목은 None 으로 두고 errors 에 기록."""
    ref = {k: None for k in REF_LABELS}
    matrix = None

    def run(name, fn, dbg_id=None, browser=True, quiet=False):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            msg = str(e).split("Stacktrace")[0].strip()
            if not quiet:
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

    ref["CP(91일)"] = run("CP(91일)", lambda: get_cp91(driver, target), "cp")
    ref["CD(91일)"] = run("CD(91일)", lambda: get_cd91(driver, target), "cd")

    irs_crs = run("IRS/CRS", lambda: scrape_kmb_irs_crs(driver, target), "kmb")
    if irs_crs:
        ref["IRS(3Y)"], ref["CRS(3년)"] = irs_crs

    sofr_tmp = None

    def get_sofr():
        nonlocal sofr_tmp
        try:
            return scrape_term_sofr_gr(target)
        except TermSofrNotYet as e:
            # 아직 발표 전: 비워 두지 않고 마지막(전 영업일) 값을 임시로 넣고 표시한다(사용자 2026-11-25).
            # 다음 날 이 날짜를 '전전일'로 다시 받을 때 실제 값으로 바뀐다.
            if quiet_sofr or e.last_value is None:
                return None
            sofr_tmp = e.last_date.isoformat()
            errors.append(f"[안내] [{tag}] Term SOFR: {target} 값이 아직 발표 전이라 {sofr_tmp} 값을 임시로 넣음"
                          " - 다음 날 실제 값으로 바뀜")
            return e.last_value, None

    tsofr = run("Term SOFR", get_sofr, browser=False, quiet=quiet_sofr)
    if tsofr:
        ref["Term SOFR"], note = tsofr
        if note:
            errors.append(f"[{tag}] Term SOFR 참고: {note}")

    ref["환율"] = run("네이버 환율", lambda: scrape_naver_fx(target, driver), browser=False)
    ref["KOSPI"] = run("네이버 KOSPI", lambda: scrape_naver_kospi(target, driver), browser=False)
    ref["다우지수"] = run("네이버 다우지수", lambda: scrape_naver_dow(target, driver), browser=False)

    ref = {k: clean(v) for k, v in ref.items()}
    return {"date": target.isoformat(), "matrix": matrix, "ref": ref, "sofr_tmp": sofr_tmp}


HISTORY_PATH = OUT / "history.json"      # 날짜별 공개 데이터(채권 표 + Term SOFR 제외 지표)
PRIVATE_PATH = OUT / "private.json"      # 날짜별 Term SOFR (비밀번호로 잠금)
HISTORY_KEEP = 1000
PUBLIC_REF_LABELS = [l for l in REF_LABELS if l != "Term SOFR"]
# 채권 표가 통째로 없는 과거 날짜용: 표의 몇 칸만 직접 넣을 때 쓰는 이름 (history 의 "pts" 에 담는다)
POINT_LABELS = ["회사채(AA-) 3Y", "회사채(AAA) 3Y", "국채(10Y)", "국채(5Y)", "국채(3Y)"]
PBKDF2_ITER = 200_000


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:  # noqa: BLE001
        return default


def save_history(hist):
    keep = sorted(hist.keys())[-HISTORY_KEEP:]
    with open(HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump({k: hist[k] for k in keep}, f, ensure_ascii=False, separators=(",", ":"))


def _derive_key(password, salt):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=PBKDF2_ITER)
    return kdf.derive(password.encode("utf-8"))


def encrypt_json(obj, password):
    """AES-256-GCM + PBKDF2. 브라우저(WebCrypto)에서 같은 비밀번호로 풀 수 있는 형식."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    salt, iv = os.urandom(16), os.urandom(12)
    ct = AESGCM(_derive_key(password, salt)).encrypt(iv, json.dumps(obj).encode("utf-8"), None)
    b64 = lambda b: base64.b64encode(b).decode("ascii")  # noqa: E731
    return {"v": 1, "iter": PBKDF2_ITER, "salt": b64(salt), "iv": b64(iv), "ct": b64(ct)}


def decrypt_json(blob, password):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    salt, iv, ct = (base64.b64decode(blob[k]) for k in ("salt", "iv", "ct"))
    return json.loads(AESGCM(_derive_key(password, salt)).decrypt(iv, ct, None).decode("utf-8"))


def business_days(start, end, holidays):
    out, d = [], start
    while d <= end:
        if d.weekday() < 5 and d not in holidays:
            out.append(d)
        d += dt.timedelta(days=1)
    return out


def merge_into_history(hist, day):
    """새로 얻은 값만 덮어쓴다(이미 저장된 좋은 값을 실패한 날의 빈 값으로 지우지 않는다)."""
    has_ref = any(day["ref"].get(k) is not None for k in PUBLIC_REF_LABELS)
    if not day["matrix"] and not has_ref:
        return
    e = hist.setdefault(day["date"], {"matrix": None, "ref": {}})
    if day["matrix"]:
        e["matrix"] = day["matrix"]
    for k in PUBLIC_REF_LABELS:
        if day["ref"].get(k) is not None:
            e["ref"][k] = day["ref"][k]


MANUAL_CSV = Path("manual_data.csv")    # 날짜,항목,값 (지표 직접 입력)
MANUAL_BOND_DIR = Path("manual_bond")   # YYYY-MM-DD.txt (채권 표 43행 x 16열 직접 입력)


def _read_text(path):
    """엑셀에서 저장한 CSV(cp949)와 웹에서 고친 파일(UTF-8)을 모두 읽는다."""
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    raise ValueError("글자 인코딩을 알 수 없음(UTF-8 또는 CP949로 저장해 주세요)")


def _manual_number(s):
    s = s.strip().replace(",", "")
    if s in ("", "-"):
        return None
    return float(s)


def load_manual(errors):
    """직접 입력한 값을 {날짜: {"ref": {...}, "matrix": [[...]] 또는 None}} 로 읽는다.
    잘못된 줄은 건너뛰고 errors 에 '[수동입력]' 문구로 남긴다(사이트가 멈추지 않도록)."""
    manual = {}

    def entry(d):
        return manual.setdefault(d, {"ref": {}, "pts": {}, "matrix": None})

    if MANUAL_CSV.exists():
        try:
            lines = _read_text(MANUAL_CSV).splitlines()
        except ValueError as e:
            errors.append(f"[수동입력] {MANUAL_CSV}: {e}")
            lines = []
        for n, line in enumerate(lines, 1):
            line = line.strip()
            if not line or line.startswith("#") or line.replace(" ", "").startswith("날짜,"):
                continue
            parts = [p.strip() for p in line.split(",", 2)]
            try:
                if len(parts) != 3:
                    raise ValueError("'날짜,항목,값' 세 칸이 아님")
                d = dt.date.fromisoformat(parts[0]).isoformat()
                label = parts[1]
                if label == "Term SOFR":
                    raise ValueError("Term SOFR는 공개 저장소에 올라가므로 여기에 적지 마세요")
                if label not in PUBLIC_REF_LABELS + POINT_LABELS:
                    raise ValueError(f"항목 이름은 {', '.join(PUBLIC_REF_LABELS + POINT_LABELS)} 중 하나")
                v = _manual_number(parts[2])
                if v is None:
                    raise ValueError("값이 비어 있음")
                entry(d)["pts" if label in POINT_LABELS else "ref"][label] = v
            except ValueError as e:
                errors.append(f"[수동입력] {MANUAL_CSV} {n}번째 줄 건너뜀: {e} ({line[:60]})")

    if MANUAL_BOND_DIR.is_dir():
        for path in sorted(MANUAL_BOND_DIR.glob("*.txt")):
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", path.stem):
                continue  # 사용법.txt 같은 파일은 무시
            try:
                d = dt.date.fromisoformat(path.stem).isoformat()
                # 탭만 있는 줄(값이 모두 빈 행)은 살리고, 완전히 빈 줄만 버린다
                rows = [ln for ln in _read_text(path).splitlines() if ln.strip(" \r")]
                if len(rows) != 43:
                    raise ValueError(f"줄 수가 43이 아님({len(rows)}줄)")
                matrix = []
                for i, ln in enumerate(rows, 1):
                    cells = ln.split("\t") if "\t" in ln else ln.split(",")
                    if len(cells) != 16:
                        raise ValueError(f"{i}번째 줄 칸 수가 16이 아님({len(cells)}칸)")
                    matrix.append([_manual_number(c) for c in cells])
                entry(d)["matrix"] = matrix
            except ValueError as e:
                errors.append(f"[수동입력] {path} 건너뜀: {e}")
    return manual


def apply_manual(target, manual_day):
    """target({"matrix", "ref"}) 에 직접 입력한 값을 덮어쓴다. 바뀐 항목 이름 목록을 돌려준다."""
    changed = []
    if manual_day["matrix"]:
        target["matrix"] = manual_day["matrix"]
        changed.append("채권 표")
    for k, v in manual_day["ref"].items():
        target.setdefault("ref", {})[k] = v
        changed.append(k)
    for k, v in manual_day.get("pts", {}).items():
        target.setdefault("pts", {})[k] = v  # 채권 표와 별도로 저장(표가 없는 날도 기간 표에 나오도록)
        changed.append(k)
    return changed


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

    start_s = os.environ.get("START_DATE", "").strip()
    end_s = os.environ.get("END_DATE", "").strip()
    backfill = bool(start_s)
    errors = []

    # 예약 실행은 GitHub 사정으로 늦거나 빠질 수 있어 아침에 여러 번 걸어 둔다.
    # 오늘 이미 전일 기준으로 문제없이 끝났으면 뒤의 예약 실행은 아무것도 바꾸지 않고 끝낸다.
    # PC 자동 실행(pc_run.py)은 SKIP_IF_DONE=1 로 같은 판단을 쓴다.
    if not backfill and (os.environ.get("GITHUB_EVENT_NAME") == "schedule"
                         or os.environ.get("SKIP_IF_DONE") == "1"):
        last = load_json(OUT / "data.json", {})
        problems = [e for e in last.get("errors", []) if not e.startswith("[안내]")]
        if (last.get("today") == today.isoformat()
                and last.get("days", {}).get("prev", {}).get("date") == prev.isoformat()
                and not problems):
            print("오늘 수집이 이미 정상적으로 끝나 이번 예약 실행은 건너뜀")
            return
        if last.get("today") == today.isoformat():
            print(f"오늘 앞선 실행에 문제가 있어 다시 수집: {problems[:3]}")

    if backfill:
        start = dt.date.fromisoformat(start_s)
        end = min(dt.date.fromisoformat(end_s) if end_s else prev, prev)
        targets = business_days(start, end, holidays)
        if len(targets) > 25:
            errors.append(f"[안내] 한 번에 최대 25영업일까지만 채워서 {targets[-25]} ~ {targets[-1]} 만 처리함")
            targets = targets[-25:]
        plan = [(d.isoformat(), d) for d in targets]
        print(f"채우기 모드: {len(plan)}영업일 ({start_s} ~ {end})")
    else:
        plan = [("prev", prev), ("prevprev", prevprev)]

    row_labels = json.load(open("row_labels.json", encoding="utf-8"))

    driver = None
    try:
        driver = make_driver(headless=True)
    except Exception as e:  # noqa: BLE001
        errors.append(f"브라우저 시작 실패: {str(e)[:200]}")

    results = {}
    for key, target in plan:
        tag = {"prev": "전일", "prevprev": "전전일"}.get(key, key)
        if driver is None:
            results[key] = {"date": target.isoformat(), "matrix": None,
                            "ref": {k: None for k in REF_LABELS}}
            continue
        results[key] = collect_day(driver, target, errors, tag,
                                   dbg_prefix={"prev": "prev", "prevprev": "prevprev"}.get(key),
                                   quiet_sofr=backfill)
    if driver:
        driver.quit()

    OUT.mkdir(exist_ok=True)
    generated_at = now.strftime("%Y-%m-%d %H:%M")

    # 날짜별 기록에 합치기 (평소 실행에서 채권 표 조회가 실패하면 저장해 둔 표로 대신)
    hist = load_json(HISTORY_PATH, {})
    if not backfill:
        for key, label in (("prev", "전일"), ("prevprev", "전전일")):
            day = results[key]
            old = hist.get(day["date"], {}).get("matrix")
            if not day["matrix"] and old:
                day["matrix"] = old
                errors.append(f"[{label}] 채권표: 직접 조회에 실패해 이전에 저장해 둔 값을 사용함")
    for day in results.values():
        merge_into_history(hist, day)
    # 직접 입력한 값은 자동 수집 값보다 우선한다 (매 실행마다 다시 덮어써서 항상 유지됨)
    manual = load_manual(errors)
    for d, mday in manual.items():
        apply_manual(hist.setdefault(d, {"matrix": None, "ref": {}}), mday)
    save_history(hist)

    # Term SOFR 는 공개 파일에 싣지 않고, 비밀번호로 잠근 private.json 에만 날짜별로 쌓는다
    (OUT / "daily_data_private.xlsx").unlink(missing_ok=True)  # 예전 버전 잔재 정리
    sofr_new = {d["date"]: d["ref"]["Term SOFR"] for d in results.values() if d["ref"]["Term SOFR"] is not None}
    pw = os.environ.get("XLSX_PASSWORD", "").strip()
    if sofr_new and not pw:
        errors.append("[안내] Term SOFR: 비밀번호 설정(XLSX_PASSWORD)이 없어 값을 저장하지 않음")
    elif sofr_new:
        priv, can_save = {}, True
        if PRIVATE_PATH.exists():
            try:
                priv = decrypt_json(load_json(PRIVATE_PATH, {}), pw)
            except Exception:  # noqa: BLE001
                can_save = False  # 기존 기록을 열 수 없으면 덮어쓰지 않고 그대로 보존한다
                errors.append("[Term SOFR 저장 보류] 이전 기록을 열지 못했어요(비밀번호가 바뀌었나요?). "
                              "기존 기록을 지키려고 이번에는 저장하지 않았습니다 - XLSX_PASSWORD 를 원래 값으로 되돌려 주세요")
        if can_save:
            # "_임시": {날짜: 대신 넣은 날짜} - 아직 발표 전이라 앞 날짜 값을 넣은 날(사이트에 '임시'로 표시)
            tmp = priv.pop("_임시", {})
            for d in results.values():
                if d["ref"]["Term SOFR"] is None:
                    continue
                if d.get("sofr_tmp"):
                    tmp[d["date"]] = d["sofr_tmp"]
                else:
                    tmp.pop(d["date"], None)      # 실제 값을 받았으면 임시 표시를 지운다
            priv.update(sofr_new)
            priv = {k: priv[k] for k in sorted(priv)[-HISTORY_KEEP:]}
            tmp = {k: v for k, v in tmp.items() if k in priv}
            if tmp:
                priv["_임시"] = tmp
            with open(PRIVATE_PATH, "w", encoding="utf-8") as f:
                json.dump(encrypt_json(priv, pw), f)
            errors.append("[안내] Term SOFR: 공개 페이지에는 싣지 않고 비밀번호로 잠가 보관 - 사이트에서 비밀번호를 입력하면 표시")
    if backfill and results and not sofr_new:
        errors.append("[안내] Term SOFR: global-rates 는 최근 며칠만 보여 줘 오래된 날짜는 비어 있음(직접 입력 필요)")

    if backfill:
        with open(OUT / "backfill_log.json", "w", encoding="utf-8") as f:
            json.dump({"generated_at": generated_at, "start": start_s, "end": end_s,
                       "processed": [d["date"] for d in results.values()], "errors": errors}, f, ensure_ascii=False)
        # 채우기 뒤에도 화면 상단(data.json)과 내려받는 엑셀이 옛 결과로 남지 않도록,
        # 방금 갱신한 날짜별 기록에서 전일/전전일을 다시 만든다.
        public_days, page_errors = {}, []
        for key, label, d in (("prev", "전일", prev), ("prevprev", "전전일", prevprev)):
            saved = hist.get(d.isoformat(), {})
            ref = {k: saved.get("ref", {}).get(k) for k in REF_LABELS}
            ref["Term SOFR"] = None
            public_days[key] = {"date": d.isoformat(), "matrix": saved.get("matrix"), "ref": ref}
            page_errors += [e for e in errors if e.startswith(f"[{d.isoformat()}]")]
            missing = [k for k in PUBLIC_REF_LABELS if ref[k] is None]
            if not saved.get("matrix"):
                missing.insert(0, "채권시가평가수익률")
            if missing:
                page_errors.append(f"[{label}] 저장된 값 없음: {', '.join(missing)}")
        page_errors += [e for e in errors if not e.startswith("[20")]  # 날짜 표시 없는 안내·전체 문구
        page_errors.append(f"[안내] 밀린 날짜 채우기({start_s} ~ {end}) 결과로 다시 만든 화면입니다")
        errors = page_errors
    else:
        public_days = copy.deepcopy(results)
        for k in public_days:
            public_days[k]["ref"]["Term SOFR"] = None
            changed = apply_manual(public_days[k], manual.get(public_days[k]["date"], {"ref": {}, "pts": {}, "matrix": None}))
            if changed:
                tag = {"prev": "전일", "prevprev": "전전일"}[k]
                errors.append(f"[안내] {tag} 직접 입력 값 사용: {', '.join(changed)}")

    data = {
        "generated_at": generated_at,
        "today": today.isoformat(),
        "maturities": MATURITIES,
        "ref_labels": REF_LABELS,
        "row_labels": row_labels,
        "days": public_days,
        "errors": errors,
    }
    with open(OUT / "data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    write_excel(OUT / "daily_data.xlsx", public_days, row_labels, generated_at, errors)

    print(f"완료. 안내·실패 {len(errors)}건")
    for e in errors:
        print(" -", e)


if __name__ == "__main__":
    main()
