"""전기넷(jungi.net) 공고별 '참여업체 엑셀' 다운로드 스크립트 (사용자 PC에서 실행).

공고마다: 상세페이지 열기 → 공고번호/종목(전기) 확인 → 참여업체보기 클릭
→ 업체목록 창에서 엑셀 클릭 → 기본 파일명 그대로 다운로드 폴더에 저장.

사용:
    pip install playwright openpyxl && playwright install chromium
    python scripts/jungi_download.py --xlsx "<목록.xlsx>" --limit 3        # 먼저 3건 시험
    python scripts/jungi_download.py --xlsx "<목록.xlsx>"                   # 전체 (끊겨도 재실행하면 이어서)

기본 저장 폴더는 사용자 Downloads(C:\\Users\\admin\\Downloads). 진행 기록은 --log (CSV).
로그인이 필요하면 창이 뜬 상태(기본 headed)에서 직접 로그인한 뒤 Enter 를 누른다.
ID/PW 는 코드·파일에 저장하지 않는다.
"""
import argparse
import csv
import random
import re
import sys
import time
from pathlib import Path

import openpyxl
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

HEADER_ROW = 6          # 엑셀 6행이 헤더, 7행부터 데이터
COL_NO, COL_URL = 3, 10  # 0-based: 공고번호(D), 상세페이지(K)
MAX_CONSECUTIVE_FAIL = 5


def read_targets(xlsx):
    ws = openpyxl.load_workbook(xlsx, read_only=True).active
    out = []
    for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
        if row[COL_NO] and row[COL_URL]:
            out.append((str(row[COL_NO]).strip(), str(row[COL_URL]).strip()))
    return out


def load_done(log):
    if not log.exists():
        return set()
    with open(log, encoding="utf-8-sig") as f:
        return {r["공고번호"] for r in csv.DictReader(f) if r["상태"] == "OK"}


def unique_path(folder, name):
    """브라우저처럼 같은 이름이 있으면 'name (1).ext' 로 저장한다 (덮어쓰지 않음)."""
    p = folder / name
    n = 1
    while p.exists():
        p = folder / f"{Path(name).stem} ({n}){Path(name).suffix}"
        n += 1
    return p


def find_click(ctx_pages, text, timeout=8000):
    """열려 있는 모든 페이지/프레임에서 text 가 보이는 요소를 찾아 반환."""
    end = time.time() + timeout / 1000
    while time.time() < end:
        for pg in ctx_pages():
            for fr in pg.frames:
                for loc in (fr.get_by_role("button", name=re.compile(text)),
                            fr.get_by_role("link", name=re.compile(text)),
                            fr.get_by_text(text, exact=False)):
                    try:
                        if loc.count() and loc.first.is_visible():
                            return loc.first
                    except Exception:
                        pass
        time.sleep(0.3)
    return None


def process(ctx, page, no, url, outdir):
    page.goto(url, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_load_state("networkidle", timeout=15000)
    body = page.inner_text("body")

    if no not in body:
        return "SKIP", "공고번호 불일치"
    m = re.search(r"종목\s*[:：]?\s*([^\n\t]+)", body)
    kind = m.group(1).strip() if m else ""
    if "전기" not in kind:
        return "SKIP", f"종목={kind or '확인불가'}"

    before = set(ctx.pages)
    btn = find_click(lambda: ctx.pages, "참여업체보기")
    if not btn:
        return "FAIL", "참여업체보기 버튼 없음"
    btn.click()
    page.wait_for_timeout(1500)  # 새 창 또는 레이어 팝업이 뜰 시간

    for p in set(ctx.pages) - before:
        try:
            p.wait_for_load_state("domcontentloaded", timeout=10000)
        except PWTimeout:
            pass

    xl = find_click(lambda: ctx.pages, r"엑셀|Excel")
    if not xl:
        return "FAIL", "엑셀 버튼 없음"
    # 다운로드는 클릭한 페이지(업체목록 창일 수 있음)에서 발생한다 — 모든 페이지에 리스너를 건다
    # 엑셀 버튼이 새 탭을 열어 다운로드하는 경우도 있으므로, 이후 생기는 페이지에도 리스너를 건다
    got = []

    def attach(pg):
        pg.on("download", lambda d: got.append(d))

    for p in ctx.pages:
        attach(p)
    ctx.on("page", attach)
    try:
        xl.click()
        end = time.time() + 60
        while not got and time.time() < end:
            page.wait_for_timeout(300)  # time.sleep 은 Playwright 이벤트를 처리하지 못한다
    finally:
        ctx.remove_listener("page", attach)
    if not got:
        return "FAIL", "다운로드 이벤트 없음"
    d = got[0]
    dest = unique_path(outdir, d.suggested_filename)
    d.save_as(dest)

    for p in set(ctx.pages) - before:   # 업체목록 창 닫기
        try:
            p.close()
        except Exception:
            pass
    return "OK", dest.name


def safe_accept(d):
    """alert/confirm 을 닫는다. 이미 닫힌 대화상자면 오류가 나므로 무시한다."""
    try:
        d.accept()
    except Exception:
        pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--out", default=str(Path.home() / "Downloads"))
    ap.add_argument("--log", default="jungi_download_log.csv")
    ap.add_argument("--start", type=int, default=0, help="목록에서 시작할 0-based 위치")
    ap.add_argument("--limit", type=int, default=0, help="이번 실행에서 처리할 최대 건수 (0=전체)")
    ap.add_argument("--delay", type=float, default=8.0, help="공고 사이 기본 대기(초), ±50%% 지터")
    ap.add_argument("--batch", type=int, default=50, help="이 건수마다 길게 쉰다 (0=쉬지 않음)")
    ap.add_argument("--batch-pause", type=float, default=120.0, help="배치마다 쉬는 시간(초)")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--profile", default=".jungi_profile", help="로그인 유지용 전용 브라우저 프로필 폴더")
    a = ap.parse_args()

    outdir = Path(a.out)
    outdir.mkdir(parents=True, exist_ok=True)
    log = Path(a.log)
    done = load_done(log)
    targets = [t for t in read_targets(a.xlsx)[a.start:] if t[0] not in done]
    if a.limit:
        targets = targets[: a.limit]
    print(f"대상 {len(targets)}건 (이미 완료 {len(done)}건 제외) → {outdir}")

    new_log = not log.exists()
    fails = 0
    t0 = time.time()
    with sync_playwright() as pw, open(log, "a", newline="", encoding="utf-8-sig") as lf:
        w = csv.writer(lf)
        if new_log:
            w.writerow(["공고번호", "상태", "상세", "시각"])
        ctx = pw.chromium.launch_persistent_context(
            a.profile, headless=a.headless, accept_downloads=True,
            viewport={"width": 1280, "height": 900})
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.on("dialog", safe_accept)
        page.goto("https://www.jungi.net/")
        if not a.headless:
            input("로그인이 필요하면 창에서 직접 로그인한 뒤 Enter 를 누르세요 (불필요하면 바로 Enter): ")

        for i, (no, url) in enumerate(targets, 1):
            try:
                status, detail = process(ctx, page, no, url, outdir)
            except Exception as e:
                status, detail = "FAIL", f"{type(e).__name__}: {str(e)[:120]}"
            # 남은 팝업 정리
            for p in ctx.pages:
                if p is not page:
                    try:
                        p.close()
                    except Exception:
                        pass
            w.writerow([no, status, detail, time.strftime("%Y-%m-%d %H:%M:%S")])
            lf.flush()
            fails = fails + 1 if status == "FAIL" else 0
            el = time.time() - t0
            eta = el / i * (len(targets) - i)
            print(f"[{i}/{len(targets)}] {no} {status} {detail}  (남은 예상 {eta/60:.0f}분)")
            if fails >= MAX_CONSECUTIVE_FAIL:
                print(f"연속 {fails}건 실패 — 중단합니다. 사이트 구조/차단 여부를 확인하세요.")
                break
            time.sleep(a.delay * random.uniform(0.5, 1.5))
            if a.batch and i % a.batch == 0 and i < len(targets):
                print(f"{i}건마다 휴식 — {a.batch_pause:.0f}초 대기 (서버 부하 완화)")
                time.sleep(a.batch_pause)
        ctx.close()


if __name__ == "__main__":
    sys.exit(main())
