"""JavaScript で描画されるページをヘッドレス Chrome（--dump-dom）で読み、CSV に保存する。

    python3 fetch_dynamic_pages.py hoshino     # 星野リゾート・リート物件の所在地・客室数（物件入替時のみ）
    python3 fetch_dynamic_pages.py moritrust   # 森トラストリートのホテル運営実績（毎月）

出力：
    data/hoshino_hotels_master.csv
    data/jreit_moritrust_hotels.csv
"""
from __future__ import annotations

import csv
import os
import re
import signal
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import requests

DATA = Path(__file__).parent / "data"
RETRIEVED = "2026-10-26"
HOSHINO = "https://www.hoshinoresorts-reit.com"
MORITRUST_URL = "https://www.mt-reit.jp/ja/portfolio/result.html"
PREF_RE = re.compile(r"^(北海道|東京都|京都府|大阪府|.{2,3}県)")


def dump_dom(url: str) -> str:
    """Chrome は DOM を出力したあとも終了しないことがあるので、時間切れでプロセスグループごと止める。"""
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as prof:
        p = subprocess.Popen(
            ["google-chrome", "--headless=new", "--no-sandbox", "--disable-gpu",
             f"--user-data-dir={prof}", "--virtual-time-budget=5000", "--dump-dom", url],
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, start_new_session=True,
        )
        try:
            out, _ = p.communicate(timeout=40)
        except subprocess.TimeoutExpired:
            os.killpg(p.pid, signal.SIGKILL)
            out, _ = p.communicate()
        return out.decode("utf-8", "ignore")


def _text(html: str) -> str:
    text = re.sub(r"<script.*?</script>|<style.*?</style>", "", html, flags=re.S)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", text))


def _write(name: str, rows: list[dict]) -> None:
    DATA.mkdir(exist_ok=True)
    with (DATA / name).open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print(f"{len(rows)} 行を data/{name} に保存")


def parse_hoshino(detail_id: str, html: str) -> dict | None:
    text = _text(html)
    m_name = re.search(r"ポートフォリオ一覧 (\S[^|｜]{0,40}?) ([A-Z]-\d+) \1", text)
    m_rooms = re.search(r"客室数 (?:（注\d*） ?)?([\d,]+) ?室", text)
    m_addr = re.search(r"所在地 (\S+)", text)
    if not (m_name and m_addr):
        return None
    pref = PREF_RE.match(m_addr.group(1))
    return {
        "物件番号": m_name.group(2),
        "物件名": m_name.group(1).strip(),
        "都道府県": pref.group(1) if pref else "",
        "所在地": m_addr.group(1),
        "客室数": int(m_rooms.group(1).replace(",", "")) if m_rooms else "",
        "出典URL": f"{HOSHINO}/ja/portfolio/detail.html?id={detail_id}",
        "取得日": RETRIEVED,
    }


def hoshino() -> None:
    html = requests.get(f"{HOSHINO}/ja/portfolio/list.html", headers={"User-Agent": "Mozilla/5.0"}, timeout=30).text
    ids = sorted(set(re.findall(r"detail\.html\?id=(\d+)", html)))
    urls = [f"{HOSHINO}/ja/portfolio/detail.html?id={i}" for i in ids]
    with ThreadPoolExecutor(max_workers=8) as ex:
        pages = list(ex.map(dump_dom, urls))
    _write("hoshino_hotels_master.csv", [r for i, p in zip(ids, pages) if (r := parse_hoshino(i, p))])


def parse_moritrust(html: str) -> list[dict]:
    """ホテル見出し（h2/h3）と表の行を出現順にたどり、月ごとの OCC・ADR・RevPAR を取り出す。"""
    body = re.sub(r"<script.*?</script>|<style.*?</style>", "", html, flags=re.S)
    hotel, months, out = None, [], {}
    for m in re.finditer(r"<h(\d)[^>]*>(.*?)</h\1>|<tr[^>]*>(.*?)</tr>", body, re.S):
        if m.group(2) is not None:
            title = re.sub(r"<[^>]+>|\s+", " ", m.group(2)).strip()
            if not re.match(r"第\d+期", title) and title not in ("森トラストリート投資法人", "ホテル運営実績"):
                hotel = title
            continue
        cells = [re.sub(r"<[^>]+>|\s+", " ", c).strip() for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", m.group(3), re.S)]
        if not cells or not hotel:
            continue
        if cells[0] == "賃料算出基準月":
            months = cells[1:]
        key = {"客室稼働率（％）": "OCC(%)", "ADR（円）": "ADR(円)", "RevPAR（円）": "RevPAR(円)"}.get(cells[0])
        if key:
            for mon, v in zip(months, cells[1:]):
                if v not in ("-", ""):
                    out.setdefault((hotel, mon), {})[key] = float(v.replace(",", ""))
    rows = []
    for (hotel, mon), vals in out.items():
        ym = re.match(r"(\d{4})年(\d{1,2})月", mon)
        rows.append({"REIT": "森トラストリート", "ホテル": hotel, "対象月": f"{ym.group(1)}-{int(ym.group(2)):02d}",
                     **vals, "前年比": "(未)（サイトに前年同月値の表示なし）",
                     "出典URL": MORITRUST_URL, "取得日": RETRIEVED})
    return sorted(rows, key=lambda r: (r["ホテル"], r["対象月"]))


def moritrust() -> None:
    _write("jreit_moritrust_hotels.csv", parse_moritrust(dump_dom(MORITRUST_URL)))


if __name__ == "__main__":
    targets = sys.argv[1:] or ["hoshino", "moritrust"]
    for t in targets:
        {"hoshino": hoshino, "moritrust": moritrust}[t]()
