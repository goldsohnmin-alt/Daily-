"""PC 자동 실행: GitHub 예약 실행이 늦을 때를 대비해 이 PC에서 오늘 자료를 수집해 올린다.

왜: GitHub 무료 예약 실행은 몇 시간씩 늦게 시작되곤 한다(2026-11-18·19 약 5시간).
언제: 작업 스케줄러가 평일 07:00과 윈도우 로그인 5분 뒤에 run_pc.bat 으로 실행한다.
하는 일:
  1) 주말·공휴일이면 아무것도 하지 않는다.
  2) GitHub 실행이 돌고 있으면 끝날 때까지 기다린다(최대 WAIT_MAX 분).
  3) 최신 결과를 받고(git pull), 오늘 이미 정상으로 끝났으면 그대로 끝낸다(SKIP_IF_DONE).
  4) 아니면 run_daily_v6.py 를 이 PC에서 돌리고 docs/ 만 커밋해 올린다(코드는 건드리지 않는다).
비밀번호: Term SOFR 저장용 XLSX_PASSWORD 는 윈도우 '사용자 환경 변수'에 넣는다(저장소에 넣지 않는다).
실행:  .venv\\Scripts\\python pc_run.py        기록: output\\pc_run.log (bat 로 실행할 때)
"""
from __future__ import annotations

import datetime as dt
import json
import os
import subprocess
import sys
import socket
import time
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
REPO = "goldsohnmin-alt/Daily-"
WORKFLOW = "daily-market-data"
WAIT_MAX = 60      # 분
POLL = 120         # 초
NET_MAX = 30       # 분. 인터넷이 안 될 때 기다리는 최대 시간(작업 스케줄러 시간 제한 안에서)


def log(*a):
    print(dt.datetime.now().strftime("%Y-%m-%d %H:%M"), *a, flush=True)


def git(*args, check=True) -> subprocess.CompletedProcess:
    r = subprocess.run(["git", *args], cwd=BASE, capture_output=True, text=True, encoding="utf-8")
    if check and r.returncode:
        raise RuntimeError(f"git {' '.join(args)} 실패: {r.stderr.strip() or r.stdout.strip()}")
    return r


def actions_running() -> bool:
    """이 저장소의 GitHub 실행 중 돌고 있거나 대기 중인 것이 있는지(공개 API, 로그인 불필요)."""
    url = f"https://api.github.com/repos/{REPO}/actions/runs?per_page=10"
    req = urllib.request.Request(url, headers={"User-Agent": "daily-site-pc-run",
                                               "Accept": "application/vnd.github+json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        runs = json.load(r)["workflow_runs"]
    return any(x["status"] in ("in_progress", "queued", "waiting") and x["name"] == WORKFLOW
               for x in runs)


def online() -> bool:
    """인터넷이 되는가(GitHub 주소를 찾고 접속까지 되는지 10초 안에 확인)."""
    try:
        socket.create_connection(("api.github.com", 443), timeout=10).close()
        return True
    except OSError:
        return False


def wait_online() -> bool:
    """인터넷이 될 때까지 1분마다 다시 확인한다(최대 NET_MAX 분, PC가 잠든 시간도 포함).
    2026-11-24: 로그인 직후 인터넷이 안 잡혀 'GitHub 실행 중'으로 잘못 알고 기다리다 끝난 일이 있었다."""
    start = time.time()
    while not online():
        if time.time() - start >= NET_MAX * 60:
            return False
        log(f"인터넷 연결이 안 되어 1분 뒤 다시 확인합니다 ({int(time.time() - start) // 60}분)")
        time.sleep(60)
    if time.time() - start > 5:
        log("인터넷이 연결되었습니다.")
    return True


def wait_for_actions() -> bool:
    start = time.time()
    while True:
        if not wait_online():
            log(f"인터넷이 {NET_MAX}분 넘게 연결되지 않아 이번에는 건너뜁니다.")
            return False
        try:
            if not actions_running():
                return True
            log(f"GitHub 실행이 돌고 있어 기다립니다 ({int(time.time() - start) // 60}분)")
        except Exception as e:      # 인터넷은 되는데 확인을 못 하면(GitHub 응답 문제) 안전하게 기다린다
            log("GitHub 실행 상태 확인 실패:", e)
        if time.time() - start >= WAIT_MAX * 60:
            log(f"GitHub 실행이 {WAIT_MAX}분 넘게 끝나지 않아 이번에는 건너뜁니다.")
            return False
        time.sleep(POLL)


def main() -> int:
    os.chdir(BASE)
    sys.path.insert(0, BASE)
    import run_daily_v6 as rd

    today = dt.datetime.now(rd.KST).date()
    if today.weekday() >= 5 or today in rd.get_holiday_set_from_csv("holidays.csv"):
        log(f"{today} 은 영업일이 아니라 실행하지 않습니다.")
        return 0

    log("PC 자동 실행 시작")
    if not os.environ.get("XLSX_PASSWORD", "").strip():
        log("XLSX_PASSWORD 환경 변수가 없어 이번 Term SOFR 는 저장되지 않습니다.")
    if not wait_for_actions():
        return 1
    git("pull", "--rebase", "--autostash", "-q")

    os.environ["SKIP_IF_DONE"] = "1"
    for k in ("START_DATE", "END_DATE"):
        os.environ.pop(k, None)
    rd.main()

    if not git("status", "--porcelain", "--", "docs").stdout.strip():
        log("바뀐 결과가 없습니다.")
        return 0
    git("add", "--", "docs")
    git("commit", "-q", "-m", f"PC update {today.isoformat()}")
    if not wait_for_actions():
        log("올리기 직전 GitHub 실행이 끝나지 않아, 커밋만 해 두었습니다. 다음 실행 때 함께 올라갑니다.")
        return 1
    git("pull", "--rebase", "--autostash", "-q")
    git("push", "-q")
    log("올렸습니다.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        log("실패:", e)
        # 끝나지 않은 rebase 가 남지 않게 정리
        if any(os.path.isdir(os.path.join(BASE, ".git", d)) for d in ("rebase-merge", "rebase-apply")):
            git("rebase", "--abort", check=False)
        sys.exit(1)
