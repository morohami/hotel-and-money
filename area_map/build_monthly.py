"""月次版：観光庁の都道府県×宿泊施設タイプ別の月次値と、REIT・事業会社・観光協会の月次公表値を
1本の縦持ちCSVにまとめ、月を切り替えられる1ファイルHTMLとタイプ別地図（PNG）を作る。

    python3 build_monthly.py            # raw/ にキャッシュがあれば再利用
    python3 build_monthly.py --refresh  # 元ファイルを取り直す

新しい月を足すときは MONTHS と各ソースのURL表（MLIT_SOKUHO2・MLIT_PRESS・ICHIGO_PDFS など）に1行足す（README参照）。
森トラストリートは JavaScript 描画のため fetch_dynamic_pages.py moritrust で先に CSV を更新する。
"""
from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path

import matplotlib.pyplot as plt
import openpyxl
import pdfplumber
from matplotlib.colors import TwoSlopeNorm

import build_area_map as B

HERE, RAW, DATA = B.HERE, B.RAW, B.DATA
RETRIEVED = B.RETRIEVED
MLIT = "https://www.mlit.go.jp/kankocho/content/"

# 表示する月（観光庁は都道府県別の第2次速報が出ている最新月まで、REIT等は最新の開示月まで）
MONTHS = [f"{y}-{m:02d}" for y, m in [(2025, 7), (2025, 8), (2025, 9), (2025, 10), (2025, 11), (2025, 12),
                                       (2026, 1), (2026, 2), (2026, 3), (2026, 4), (2026, 5), (2026, 6),
                                       (2026, 7), (2026, 8)]]
MLIT_LATEST = "2026-07"

# 観光庁：月ごとの「その月の人泊数（第4表）」と「前年同月の人泊数（第4表）」の出どころ
MLIT_SUII = MLIT + "002026697.xlsx"           # 推移表：都道府県×タイプ別の客室稼働率（2011年～）
MLIT_KAKUTEI = {2024: MLIT + "002026764.xlsx", 2025: MLIT + "002026762.xlsx"}  # 年の確定値
MLIT_SOKUHO2 = {  # 第2次速報 集計結果 Excel
    "2026-01": MLIT + "002026744.xlsx", "2026-02": MLIT + "002026742.xlsx", "2026-03": MLIT + "002026740.xlsx",
    "2026-04": MLIT + "002026733.xlsx", "2026-05": MLIT + "002026728.xlsx", "2026-06": MLIT + "002026747.xlsx",
    "2026-07": MLIT + "002025568.xlsx",
}
MLIT_PRESS = {  # 第2次速報 報道発表資料（検算用：観光庁自身が計算した前年同月差・比率）
    "2026-01": MLIT + "002026745.pdf", "2026-02": MLIT + "002026743.pdf", "2026-03": MLIT + "002026741.pdf",
    "2026-04": MLIT + "002026735.pdf", "2026-05": MLIT + "002026730.pdf", "2026-06": MLIT + "002026726.pdf",
    "2026-07": MLIT + "002025565.pdf",
}
MLIT_SOKUHO1 = {"2026-08": MLIT + "002025567.xlsx"}  # 第1次速報（全国値のみ）

ICHIGO_PDFS = {  # いちごホテルリート「ホテル運営状況」（月次）
    m: f"https://www.ichigo-hotel.co.jp/ir/news/news_file/file/IchigoHotel_{d}_Monthly_Operating_Results_JPN.pdf"
    for m, d in [("2025-07", "20250825"), ("2025-08", "20250925"), ("2025-09", "20251027"), ("2025-10", "20251125"),
                 ("2025-11", "20251225"), ("2025-12", "20260126"), ("2026-01", "20260225"), ("2026-02", "20260325"),
                 ("2026-03", "20260427"), ("2026-04", "20260525"), ("2026-05", "20260625"), ("2026-06", "20260727"),
                 ("2026-07", "20260825"), ("2026-08", "20260925")]
}
HOSHINO_XLSX = B.CONFIG["reit_files"]["hoshino_xlsx"]
JHR_XLSX = B.CONFIG["reit_files"]["jhr_xlsx"]
INV_PDFS = {"2026-08": B.CONFIG["reit_files"]["inv_pdf"]}
POLARIS_XLSX = "https://xml.irpocket.com/C3010/2026/performance_jp.xlsx"
POLARIS_PAGE = "https://www.polaris-holdings.com/ir_performance/"
KYOTO_PAGE = "https://www.kyokanko.or.jp/report/hotel{ym}"
MORITRUST_CSV = DATA / "jreit_moritrust_hotels.csv"

TYPES = {"計": "全体", "旅館": "旅館", "リゾート\nホテル": "リゾート", "ビジネス\nホテル": "ビジネス", "シティ\nホテル": "シティ"}
FOCUS = ["全国", "東京都", "大阪府", "京都府", "北海道", "福岡県", "沖縄県", "愛知県", "神奈川県"]

COLS = ["出典キー", "公表者", "資料", "対象月", "比較対象月", "エリア区分", "エリア", "都道府県", "公表区分", "タイプ",
        "指標", "単位", "当月", "前年同月", "前年差・比", "差の単位", "差の根拠", "施設数", "客室数", "出典URL", "取得日", "備考"]


def rec(**kw) -> dict:
    r = {c: "" for c in COLS}
    r.update({k: ("" if v is None else v) for k, v in kw.items()})
    r["取得日"] = RETRIEVED
    return r


def prev_month(ym: str) -> str:
    y, m = map(int, ym.split("-"))
    return f"{y - 1}-{m:02d}"


def raw_key(url: str) -> str:
    stem = re.sub(r"\.(xlsx|xls|pdf|html)$", "", url.split("//", 1)[1])
    return re.sub(r"[^\w-]", "_", stem)[-90:]


def get(url: str, refresh: bool) -> Path:
    return B.fetch(raw_key(url), url, refresh)


def get_html(url: str, refresh: bool) -> str:
    RAW.mkdir(exist_ok=True)
    path = RAW / f"{raw_key(url)}.html"
    if refresh or not path.exists():
        import requests
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
        r.raise_for_status()
        path.write_bytes(r.content)
    return path.read_text(encoding="utf-8", errors="ignore")


def rnd(v, d=1):
    return None if v is None else round(v, d)


# ---------------------------------------------------------------- 観光庁

def _sheet_rows(path: Path, prefix: str, month: int):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    name = next(s for s in wb.sheetnames if s.startswith(prefix) and s.endswith(f"({month}月)"))
    return list(wb[name].iter_rows(values_only=True))


def _label(v) -> str | None:
    s = str(v or "")
    if re.match(r"^[\s\u3000]*\d{2}\D", s):
        return B.pref_name(s)
    if re.match(r"^令和\d+年\d+月$", s.strip()):
        return "全国"
    return None


def read_table4(path: Path, month: int) -> dict:
    """第4表：都道府県×タイプ別の延べ宿泊者数・外国人延べ宿泊者数（人泊）。"""
    rows = _sheet_rows(path, "第4表", month)
    head = rows[4]
    fcol = next(i for i, v in enumerate(rows[3]) if v and "外国人" in str(v))
    out = {}
    for r in rows[6:]:
        lab = _label(r[0])
        if not lab:
            continue
        d = {("全体", "total"): B.num(r[1]), ("全体", "foreign"): B.num(r[fcol])}
        for i, h in enumerate(head):
            if h in TYPES and i != 1:
                d[(TYPES[h], "foreign" if i > fcol else "total")] = B.num(r[i])
        out[lab] = d
    return out


def read_table8_types(path: Path, month: int) -> dict:
    """第8表：都道府県×タイプ別の客室稼働率（％）。推移表の検算に使う。"""
    rows = _sheet_rows(path, "第8表", month)
    head = rows[4]
    tcol = {TYPES[h]: i for i, h in enumerate(head) if h in TYPES}
    out = {}
    for r in rows[6:]:
        lab = _label(r[0])
        if lab:
            out[lab] = {"全体": B.num(r[1]), **{t: B.num(r[i]) for t, i in tcol.items()}}
    return out


def read_suii_occ(path: Path) -> dict:
    """推移表 4-1（2026年～）と 旧4-2（2011～2025年）から {(都道府県, タイプ, 'YYYY-MM'): OCC}。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for sheet, year_row, month_row in (("4-1", 2, 3), ("旧4-2", 2, 3)):
        rows = list(wb[sheet].iter_rows(values_only=True))
        cols, cur = {}, None
        for i, m in enumerate(rows[month_row]):
            if rows[year_row][i]:
                cur = rows[year_row][i]
            mm = re.match(r"(\d+)月", str(m or ""))
            era = re.match(r"令和(\d+)年", str(cur or ""))
            if mm and era:
                cols[i] = f"{2018 + int(era.group(1))}-{int(mm.group(1)):02d}"
        pref = None
        for r in rows[4:]:
            if r[0] and str(r[0]).strip():
                pref = "全国" if "全" in str(r[0]) and "国" in str(r[0]) else (
                    B.pref_name(r[0]) if re.match(r"^[\s\u3000]*\d{2}", str(r[0])) else None)
            if pref and r[1] in TYPES:
                for i, ym in cols.items():
                    out.setdefault((pref, TYPES[r[1]], ym), B.num(r[i]))
    return out


def people_source(ym: str) -> str | None:
    y = int(ym[:4])
    return MLIT_SOKUHO2.get(ym) or MLIT_KAKUTEI.get(y)


def build_mlit_monthly(refresh: bool) -> tuple[list[dict], list[dict]]:
    occ = read_suii_occ(get(MLIT_SUII, refresh))
    t4 = {}
    for ym in MONTHS:
        for m_ in (ym, prev_month(ym)):
            url = people_source(m_)
            if url and m_ not in t4:
                t4[m_] = (read_table4(get(url, refresh), int(m_[5:])), url)
    wide, recs = [], []
    for ym in MONTHS:
        py = prev_month(ym)
        if ym not in t4:
            continue
        cur, cur_url = t4[ym]
        prv, prv_url = t4[py]
        for pref in cur:
            for t in TYPES.values():
                a, b = occ.get((pref, t, ym)), occ.get((pref, t, py))
                tn, tp = cur[pref].get((t, "total")), prv.get(pref, {}).get((t, "total"))
                fn, fp = cur[pref].get((t, "foreign")), prv.get(pref, {}).get((t, "foreign"))
                fr_n = rnd(fn / tn * 100) if fn is not None and tn else None
                fr_p = rnd(fp / tp * 100) if fp is not None and tp else None
                d_occ = rnd(a - b) if a is not None and b is not None else None
                d_fr = rnd(fr_n - fr_p) if fr_n is not None and fr_p is not None else None
                wide.append({
                    "対象月": ym, "前年同月": py, "都道府県": pref, "タイプ": t,
                    "OCC_当月(%)": a, "OCC_前年同月(%)": b, "OCC_前年差(pt)": d_occ,
                    "延べ宿泊者数_当月(人泊)": tn, "延べ宿泊者数_前年同月(人泊)": tp,
                    "外国人延べ_当月(人泊)": fn, "外国人延べ_前年同月(人泊)": fp,
                    "外国人比率_当月(%)": fr_n, "外国人比率_前年同月(%)": fr_p, "外国人比率_前年差(pt)": d_fr,
                    "出典_OCC": MLIT_SUII, "出典_人泊_当月": cur_url, "出典_人泊_前年同月": prv_url,
                    "取得日": RETRIEVED,
                    "備考": ("2026年1月分から層化基準を変更（前年差に見直しの影響を含み得る）" if ym >= "2026-01" else ""),
                })
                common = dict(出典キー="M", 公表者="観光庁", 資料="宿泊旅行統計調査", 対象月=ym, 比較対象月=py,
                              エリア区分="全国" if pref == "全国" else "都道府県", エリア=pref,
                              都道府県="" if pref == "全国" else pref, 公表区分=t, タイプ=t)
                note = "2026年1月分から層化基準を変更" if ym >= "2026-01" else ""
                recs.append(rec(**common, 指標="OCC", 単位="%", 当月=a, 前年同月=b, **{"前年差・比": d_occ},
                                差の単位="pt", 差の根拠="当月・前年同月の公表値の差", 出典URL=MLIT_SUII,
                                備考=note if a is not None else "(未)：公表値なし（'-'）"))
                recs.append(rec(**common, 指標="外国人比率", 単位="%", 当月=fr_n, 前年同月=fr_p,
                                **{"前年差・比": d_fr}, 差の単位="pt",
                                差の根拠="公表の外国人延べ÷延べ宿泊者数（人泊）から算出",
                                出典URL=cur_url + " / " + prv_url,
                                備考=note if fr_n is not None else "(未)：公表人泊数なし（'-'）"))
    # 第1次速報（全国のみ）
    for ym, url in MLIT_SOKUHO1.items():
        nat = B.json_national_aug(get(url, refresh))
        b = occ.get(("全国", "全体", prev_month(ym)))
        a = B.num(nat["OCC全体"])
        recs.append(rec(出典キー="M", 公表者="観光庁", 資料="宿泊旅行統計調査（第1次速報）", 対象月=ym,
                        比較対象月=prev_month(ym), エリア区分="全国", エリア="全国", 公表区分="全体", タイプ="全体",
                        指標="OCC", 単位="%", 当月=a, 前年同月=b, **{"前年差・比": rnd(a - b)}, 差の単位="pt",
                        差の根拠="当月（第1次速報）・前年同月（推移表）の公表値の差", 出典URL=url + " / " + MLIT_SUII,
                        備考="第1次速報は全国値のみ"))
        for pref in t4[MLIT_LATEST][0]:
            if pref != "全国":
                recs.append(rec(出典キー="M", 公表者="観光庁", 資料="宿泊旅行統計調査", 対象月=ym,
                                エリア区分="都道府県", エリア=pref, 都道府県=pref, 公表区分="全体", タイプ="全体",
                                指標="OCC", 単位="%", 出典URL=B.CONFIG["mlit_page"],
                                備考="(未)：都道府県別は第2次速報（2026-10-30公表予定）で公表"))
    return wide, recs


# ---------------------------------------------------------------- ホテル別（いちご・星野・森トラスト）

def hotel_recs(reit, key, h, ym, src, category="", note="") -> list[dict]:
    pct = lambda a, b: rnd((a / b - 1) * 100) if a and b else None
    base = dict(出典キー=key, 公表者=reit, 資料="ホテル別月次運営実績", 対象月=ym, 比較対象月=prev_month(ym),
                エリア区分="ホテル", エリア=h.name, 都道府県=h.pref, 公表区分=category, 施設数=1, 客室数=h.rooms,
                出典URL=src)
    closed = "（当月0＝休館・改装等）" if h.occ == 0 else ""
    miss = lambda v: "" if v else f"(未)：当月値なし{closed}"
    return [
        rec(**base, 指標="OCC", 単位="%", 当月=h.occ or None, 前年同月=h.occ_py or None,
            **{"前年差・比": rnd(h.occ - h.occ_py) if h.occ and h.occ_py else None}, 差の単位="pt",
            差の根拠="当月・前年同月の公表値の差", 備考=miss(h.occ) or note),
        rec(**base, 指標="ADR", 単位="円", 当月=h.adr or None, 前年同月=h.adr_py or None,
            **{"前年差・比": pct(h.adr, h.adr_py)}, 差の単位="%", 差の根拠="当月・前年同月の公表値の比",
            備考=miss(h.adr) or note),
        rec(**base, 指標="RevPAR", 単位="円", 当月=h.revpar or None, 前年同月=h.revpar_py or None,
            **{"前年差・比": pct(h.revpar, h.revpar_py)}, 差の単位="%", 差の根拠="当月・前年同月の公表値の比",
            備考=miss(h.revpar) or note),
    ]


def ichigo_monthly(refresh: bool) -> tuple[dict, list[dict]]:
    lookup = B.ichigo_lookup(refresh)
    by_month, recs = {}, []
    for ym, url in ICHIGO_PDFS.items():
        path = get(url, refresh)
        with pdfplumber.open(path) as pdf:
            head = pdf.pages[0].extract_text() or ""
        m = re.search(r"(\d{4})年(\d{1,2})月度", head)
        if not m or f"{m.group(1)}-{int(m.group(2)):02d}" != ym:
            raise ValueError(f"いちご {url} の対象月が {ym} と一致しない")
        hs = B.parse_ichigo_pdf(path, url, lookup)
        by_month[ym] = hs
        for h in hs:
            recs += hotel_recs("いちごホテルリート", "I", h, ym, url)
    return by_month, recs


def hoshino_monthly(refresh: bool) -> tuple[dict, list[dict]]:
    """運営実績データ（期ごとのシート）を全期分読み、同じ月が複数期にあれば新しい期を採る。"""
    wb = openpyxl.load_workbook(get(HOSHINO_XLSX, refresh), read_only=True, data_only=True)
    vals, meta = {}, {}
    for sheet in [s for s in wb.sheetnames if re.match(r"第\d+期開示", s.strip())]:
        rows = list(wb[sheet].iter_rows(values_only=True))
        cols, cur = {}, None
        for i, mm in enumerate(rows[4]):
            if rows[3][i] and re.match(r"\d{4}年", str(rows[3][i])):
                cur = str(rows[3][i])[:4]
            m = re.match(r"(\d+)月", str(mm or ""))
            if m and cur:
                cols[i] = f"{cur}-{int(m.group(1)):02d}"
        code = name = cat = None
        for r in rows[5:]:
            if r[3]:
                code, name = str(r[3]).strip(), str(r[4]).strip()
                cat = str(r[2]).strip() if r[2] else cat
                meta.setdefault(code, (name, cat))
            if code and r[5]:
                for i, ym in cols.items():
                    vals.setdefault((code, ym, r[5]), (B.num(r[i]), sheet))
    master = {}
    with (DATA / "hoshino_hotels_master.csv").open(encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            master[r["物件番号"]] = (r["都道府県"], int(r["客室数"]) if r["客室数"] else None)
    by_month, recs = {}, []
    for ym in MONTHS:
        py = prev_month(ym)
        for code, (name, cat) in meta.items():
            g = lambda m_, k: vals.get((code, m_, k), (None, None))[0]
            if all(g(m_, k) is None for m_ in (ym, py) for k in ("客室稼働率", "ADR（円）")):
                continue
            pref, rooms = master.get(code, ("", None))
            o, op = g(ym, "客室稼働率"), g(py, "客室稼働率")
            h = B.Hotel("星野リゾート・リート", name, pref, rooms,
                        rnd(o * 100) if o is not None else None, rnd(op * 100) if op is not None else None,
                        g(ym, "ADR（円）"), g(py, "ADR（円）"), g(ym, "RevPAR（円）"), g(py, "RevPAR（円）"), HOSHINO_XLSX)
            by_month.setdefault(ym, []).append(h)
            sheet = vals.get((code, ym, "客室稼働率"), (None, ""))[1]
            recs += hotel_recs("星野リゾート・リート", "H", h, ym, HOSHINO_XLSX, category=cat,
                               note=f"シート「{sheet.strip()}」")
    return by_month, recs


def moritrust_monthly() -> list[dict]:
    recs = []
    with MORITRUST_CSV.open(encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    pref = {"新大阪": "大阪府", "東京": "東京都"}
    for r in rows:
        p = next((v for k, v in pref.items() if k in r["ホテル"]), "")
        for metric, col, unit in (("OCC", "OCC(%)", "%"), ("ADR", "ADR(円)", "円"), ("RevPAR", "RevPAR(円)", "円")):
            recs.append(rec(出典キー="T", 公表者="森トラストリート", 資料="ホテル運営実績（ウェブページ）",
                            対象月=r["対象月"], エリア区分="ホテル", エリア=r["ホテル"], 都道府県=p, 施設数=1,
                            指標=metric, 単位=unit, 当月=B.num(r[col]), 出典URL=r["出典URL"],
                            備考="前年同月比は(未)：サイトに前年同月値の表示なし"))
    have = {r["対象月"] for r in rows}
    for ym in MONTHS:
        if ym not in have:
            recs.append(rec(出典キー="T", 公表者="森トラストリート", 資料="ホテル運営実績（ウェブページ）", 対象月=ym,
                            エリア区分="ホテル", エリア="（全ホテル）", 指標="OCC", 単位="%",
                            出典URL="https://www.mt-reit.jp/ja/portfolio/result.html",
                            備考="(未)：サイトの表示期間（2025-12～2026-07）外"))
    return recs


def pref_aggregate(by_src: dict[str, dict], ym: str) -> list[dict]:
    """ホテル別の公表値を所在都道府県で客室数加重。前年同月と比べられるホテルだけを使う。"""
    groups = {}
    for src, by_month in by_src.items():
        for h in by_month.get(ym, []):
            if h.ok() and h.pref and h.rooms:
                groups.setdefault(h.pref, []).append((src, h))
    out = []
    for pref, items in groups.items():
        hs = [h for _, h in items]
        rooms = sum(h.rooms for h in hs)
        w = lambda a: sum(getattr(h, a) * h.rooms for h in hs) / rooms
        rev, rev_py, occ, occ_py = w("revpar"), w("revpar_py"), w("occ"), w("occ_py")
        adr, adr_py = rev / (occ / 100), rev_py / (occ_py / 100)
        urls = sorted({h.src for h in hs})
        base = dict(出典キー="R", 公表者="いちご・星野（ホテル別開示を集計）", 資料="都道府県別集計（客室数加重）",
                    対象月=ym, 比較対象月=prev_month(ym), エリア区分="都道府県", エリア=pref, 都道府県=pref,
                    施設数=len(hs), 客室数=rooms, 出典URL=" / ".join(urls),
                    備考="内訳：" + "、".join(h.name for h in hs))
        out += [
            rec(**base, 指標="OCC", 単位="%", 当月=rnd(occ), 前年同月=rnd(occ_py), **{"前年差・比": rnd(occ - occ_py)},
                差の単位="pt", 差の根拠="ホテル別OCCを客室数で加重"),
            rec(**base, 指標="ADR", 単位="円", 当月=round(adr), 前年同月=round(adr_py),
                **{"前年差・比": rnd((adr / adr_py - 1) * 100)}, 差の単位="%",
                差の根拠="加重RevPAR÷加重OCC（客室数加重）"),
            rec(**base, 指標="RevPAR", 単位="円", 当月=round(rev), 前年同月=round(rev_py),
                **{"前年差・比": rnd((rev / rev_py - 1) * 100)}, 差の単位="%", 差の根拠="ホテル別RevPARを客室数で加重"),
        ]
    return out


# ---------------------------------------------------------------- 地域別の公表値（JHR・インヴィンシブル・ポラリス・京都市）

def jhr_monthly(refresh: bool) -> list[dict]:
    ws = openpyxl.load_workbook(get(JHR_XLSX, refresh), data_only=True)["変動賃料導入29ホテル（地域別）"]
    data, region, metric = {}, None, None
    for r in ws.iter_rows(values_only=True):
        if r[1] and r[3] is None and all(c is None for c in r[4:16]) and "主要指標" not in str(r[1]):
            region = r[1]
        if r[1] in ("客室稼働率（％）", "ADR（円）", "RevPAR（円）"):
            metric = r[1]
        elif r[1] and "売上高" in str(r[1]):
            metric = None
        if region and metric and r[3] in ("2026年", "増減", "増減率"):
            data.setdefault(region, {}).setdefault(metric, {})[r[3]] = r[4:16]
    count = {"北海道": 2, "東京": 3, "関東（東京除く）": 5, "大阪": 4, "関西（大阪除く）": 4,
             "中国": 2, "九州（沖縄除く）": 5, "沖縄": 4}
    recs = []
    for ym in MONTHS:
        y, m = int(ym[:4]), int(ym[5:])
        for reg, d in data.items():
            base = dict(出典キー="J", 公表者="ジャパン・ホテル・リート", 資料="ホテル運営実績Excel（変動賃料等導入29ホテル・地域別）",
                        対象月=ym, 比較対象月=prev_month(ym), エリア区分="公表地域", エリア=reg, 公表区分=reg,
                        施設数=count.get(reg), 出典URL=JHR_XLSX)
            vals = {k: (v["2026年"][m - 1] if y == 2026 else None) for k, v in d.items()}
            if y != 2026 or vals["客室稼働率（％）"] in (None, ""):
                recs.append(rec(**base, 指標="OCC", 単位="%", 備考="(未)：公表Excelは2026年1月分以降のみ掲載"))
                continue
            o, a, p = d["客室稼働率（％）"], d["ADR（円）"], d["RevPAR（円）"]
            recs += [
                rec(**base, 指標="OCC", 単位="%", 当月=rnd(o["2026年"][m - 1] * 100), **{"前年差・比": o["増減"][m - 1]},
                    差の単位="pt", 差の根拠="REIT公表の増減"),
                rec(**base, 指標="ADR", 単位="円", 当月=a["2026年"][m - 1], **{"前年差・比": rnd(a["増減率"][m - 1] * 100)},
                    差の単位="%", 差の根拠="REIT公表の増減率"),
                rec(**base, 指標="RevPAR", 単位="円", 当月=p["2026年"][m - 1], **{"前年差・比": rnd(p["増減率"][m - 1] * 100)},
                    差の単位="%", 差の根拠="REIT公表の増減率"),
            ]
    return recs


def invincible_monthly(refresh: bool) -> list[dict]:
    recs = []
    for ym in MONTHS:
        base = dict(出典キー="V", 公表者="インヴィンシブル投資法人", 資料="月次運用実績（国内ホテル101物件・地域別）",
                    対象月=ym, 比較対象月=prev_month(ym), エリア区分="公表地域")
        if ym not in INV_PDFS:
            recs.append(rec(**base, エリア="（全地域）", 指標="OCC", 単位="%", 出典URL="https://www.invincible-inv.co.jp/",
                            備考="(未)：公式サイトが緊急メンテナンス（503）でTDnetは直近31日分のみのため取得できず"))
            continue
        B.CONFIG["reit_files"]["inv_pdf"] = INV_PDFS[ym]
        for r in B.invincible_regions(refresh):
            b = dict(base, エリア=r["地域"], 公表区分=r["地域"], 施設数=r["ホテル数"], 出典URL=INV_PDFS[ym],
                     備考="" if r["ホテル数"] else "地域別の物件数は非記載")
            recs += [
                rec(**b, 指標="OCC", 単位="%", 当月=r["OCC_当月(%)"], **{"前年差・比": r["OCC_前年差(pt)"]},
                    差の単位="pt", 差の根拠="REIT公表の前年同月差"),
                rec(**b, 指標="ADR", 単位="円", 当月=r["ADR_当月(円)"], **{"前年差・比": r["ADR_前年比(%)"]},
                    差の単位="%", 差の根拠="REIT公表の前年同月比"),
                rec(**b, 指標="RevPAR", 単位="円", 当月=r["RevPAR_当月(円)"], **{"前年差・比": r["RevPAR_前年比(%)"]},
                    差の単位="%", 差の根拠="REIT公表の前年同月比"),
            ]
    return recs


def polaris_monthly(refresh: bool) -> list[dict]:
    """ポラリス・ホールディングス 月次経営レポート Excel（2026年度国内比較可能ホテル68軒）。
    エリア別・全体とも、同じ68軒の前年度（2025年4月～）の同月値も掲載されている。"""
    wb = openpyxl.load_workbook(get(POLARIS_XLSX, refresh), data_only=True)
    metrics = {"客室稼働率": ("OCC", "%"), "ADR": ("ADR", "円"), "RevPAR": ("RevPAR", "円"), "インバウンド比率": ("外国人比率", "%")}
    recs = []
    rows = list(wb["エリア別_単月"].iter_rows(values_only=True))
    yrow = next(i for i, r in enumerate(rows) if r[2] == "単月")
    ycur, cols = None, {}
    for i, mm in enumerate(rows[yrow + 1]):
        if isinstance(rows[yrow][i], int):
            ycur = rows[yrow][i]
        m = re.match(r"(\d+)月", str(mm or ""))
        if m and ycur:
            cols[i] = f"{ycur}-{int(m.group(1)):02d}"
    area = n = None
    table = {}
    for r in rows[yrow + 2:]:
        if r[1] and "ホテル" in str(r[1]):
            area, n = re.sub(r"：.*", "", str(r[1])), r[3]
        if area and r[2] in metrics:
            for i, ym in cols.items():
                if B.num(r[i]) is not None:
                    table[(area, metrics[r[2]], ym)] = (B.num(r[i]), n)
    got = set()
    for (area, (metric, unit), ym), (v, n) in table.items():
        got.add(ym)
        pv = table.get((area, (metric, unit), prev_month(ym)), (None, None))[0]
        pct = unit == "%"
        cur = rnd(v * 100) if pct else v
        prv = (rnd(pv * 100) if pct else pv) if pv is not None else None
        chg = (rnd(cur - prv) if pct else rnd((v / pv - 1) * 100)) if pv else None
        recs.append(rec(出典キー="P", 公表者="ポラリス・ホールディングス", 資料="月次経営レポート（エリア別・単月・修正値）",
                        対象月=ym, 比較対象月=prev_month(ym), エリア区分="公表地域", エリア=area, 公表区分=area, 施設数=n,
                        指標=metric, 単位=unit, 当月=cur, 前年同月=prv, **{"前年差・比": chg}, 差の単位="pt" if pct else "%",
                        差の根拠="当年度・前年度の同月公表値の差・比" if pv is not None else "",
                        出典URL=POLARIS_XLSX,
                        備考="" if pv is not None else "2026年度比較可能ホテルの前年度値として掲載（その前年同月は(未)）"))
    rows = list(wb["全体"].iter_rows(values_only=True))
    hdr = next(r for r in rows if r[3] == "単月")
    mcols = {i: int(re.match(r"(\d+)月", str(v)).group(1)) for i, v in enumerate(hdr) if re.match(r"\d+月$", str(v or ""))}
    metric = None
    block = {}
    for r in rows[rows.index(hdr) + 1:]:
        if r[3] == "累計":
            break
        if r[1] in metrics:
            metric = metrics[r[1]]
        if metric and r[2] in ("2026年3月期", "2027年3月期"):
            block.setdefault(metric, {})[r[2]] = {m: B.num(r[i]) for i, m in mcols.items()}
    n_all = rows[0][4]
    for (metric, unit), d in block.items():
        for m, pv in d["2026年3月期"].items():
            ym = f"{2025 if m >= 4 else 2026}-{m:02d}"
            if pv is None or ym not in MONTHS:
                continue
            recs.append(rec(出典キー="P", 公表者="ポラリス・ホールディングス", 資料="月次経営レポート（国内比較可能ホテル全体・単月）",
                            対象月=ym, エリア区分="全国", エリア="国内比較可能ホテル全体", 公表区分="全体", 施設数=n_all,
                            指標=metric, 単位=unit, 当月=rnd(pv * 100) if unit == "%" else round(pv), 出典URL=POLARIS_XLSX,
                            備考="2026年度比較可能ホテルの前年度値として掲載（その前年同月は(未)）"))
        for m, v in d["2027年3月期"].items():
            if v is None:
                continue
            ym = f"{2026 if m >= 4 else 2027}-{m:02d}"
            pv = d["2026年3月期"].get(m)
            cur = rnd(v * 100) if unit == "%" else round(v)
            prv = rnd(pv * 100) if unit == "%" and pv is not None else (round(pv) if pv is not None else None)
            chg = (rnd(cur - prv) if unit == "%" else rnd((v / pv - 1) * 100)) if pv else None
            recs.append(rec(出典キー="P", 公表者="ポラリス・ホールディングス", 資料="月次経営レポート（国内比較可能ホテル全体・単月）",
                            対象月=ym, 比較対象月=prev_month(ym), エリア区分="全国", エリア="国内比較可能ホテル全体",
                            公表区分="全体", 施設数=n_all, 指標=metric, 単位=unit, 当月=cur, 前年同月=prv,
                            **{"前年差・比": chg}, 差の単位="pt" if unit == "%" else "%",
                            差の根拠="当期・前期の同月公表値の差・比", 出典URL=POLARIS_XLSX))
    for ym in MONTHS:
        if ym not in got:
            recs.append(rec(出典キー="P", 公表者="ポラリス・ホールディングス", 資料="月次経営レポート", 対象月=ym,
                            エリア区分="公表地域", エリア="（全エリア）", 指標="OCC", 単位="%", 出典URL=POLARIS_PAGE,
                            備考="(未)：掲載Excelは2026年度分と、その比較用の前年度4～8月分のみ（2025年度版のURLは404）"))
    return recs


def _kyoto_text(html: str) -> str:
    t = re.sub(r"<script.*?</script>|<style.*?</style>", "", html, flags=re.S)
    t = re.sub(r"<[^>]+>", " ", t).replace("&nbsp;", " ")
    return re.sub(r"\s+", " ", t)


def _kyoto_signed(seg: str, mag: float, cur: float, prv: float | None, ratio: bool) -> float | None:
    """カードの数値は符号なしなので、前年値（あれば）か本文の増減の語で符号を決める。"""
    if prv is not None:
        calc = (cur / prv - 1) * 100 if ratio else cur - prv
        return mag if calc > 0 else -mag if calc < 0 else 0.0
    if re.search(r"増|上回", seg):
        return mag
    if re.search(r"減|下回", seg):
        return -mag
    return None


def parse_kyoto(html: str) -> dict:
    t = _kyoto_text(html)
    t = t[t.index("調査結果のポイント"):]
    out = {}
    specs = [("OCC", r"客室稼働率 ([\d.]+) ?[％%] ([\d.]+) ?pts", r"主要ホテル ?(\d+) ?施設", False),
             ("ADR", r"平均客室単価 ([\d,]+) ?円 ([\d.]+) ?[％%]", None, True),
             ("RevPAR", r"客室収益指数 ([\d,]+) ?円 ([\d.]+) ?[％%]", None, True),
             ("旅館OCC", r"旅館の客室稼働率 ([\d.]+) ?[％%] ([\d.]+) ?pts", r"旅館 ?(\d+) ?施設", False)]
    for key, pat, npat, ratio in specs:
        m = re.search(pat, t)
        seg = t[m.end():m.end() + 300]
        end = re.search(r"た(?=\s)", seg)  # カード直後の1文（「…となった」「…更新した」）だけを見る
        seg = seg[:end.end()] if end else seg[:150]
        cur, mag = B.num(m.group(1)), B.num(m.group(2))
        pm = re.search(r"前年同月の ?([\d,.]+) ?(?:[％%]|円)", seg)
        prv = B.num(pm.group(1)) if pm else None
        n = re.search(npat, seg) if npat else None
        out[key] = {"当月": cur, "前年同月": prv, "差": _kyoto_signed(seg, mag, cur, prv, ratio),
                    "差の大きさ": mag, "施設数": int(n.group(1)) if n else None}
    m = re.search(r"今回のホテル統計の対象施設数は ?(\d+) ?軒", t) or re.search(r"主要ホテル ?(\d+) ?施設", t)
    for k in ("ADR", "RevPAR"):
        out[k]["施設数"] = out["OCC"]["施設数"] or (int(m.group(1)) if m else None)
    return out


def kyoto_monthly(refresh: bool) -> list[dict]:
    recs = []
    for ym in MONTHS:
        url = KYOTO_PAGE.format(ym=ym.replace("-", ""))
        k = parse_kyoto(get_html(url, refresh))
        base = dict(出典キー="K", 公表者="京都市観光協会", 資料="京都市観光協会データ月報", 対象月=ym,
                    比較対象月=prev_month(ym), エリア区分="市", エリア="京都市", 都道府県="京都府", 出典URL=url)
        for key, metric, cat, unit, du in (("OCC", "OCC", "ホテル（市内主要ホテル）", "%", "pt"),
                                            ("ADR", "ADR", "ホテル（市内主要ホテル）", "円", "%"),
                                            ("RevPAR", "RevPAR", "ホテル（市内主要ホテル）", "円", "%"),
                                            ("旅館OCC", "OCC", "旅館（市内主要旅館）", "%", "pt")):
            v = k[key]
            recs.append(rec(**base, 公表区分=cat, 施設数=v["施設数"], 指標=metric, 単位=unit, 当月=v["当月"],
                            前年同月=v["前年同月"], **{"前年差・比": v["差"]}, 差の単位=du,
                            差の根拠="協会公表の前年同月差・比（符号は本文の増減表記）",
                            備考="" if v["差"] is not None else
                            f"(未)：差の大きさ{v['差の大きさ']}{du}のみ記載で、本文に増減の方向の記載なし"))
        recs.append(rec(**base, 公表区分="ホテル（市内主要ホテル）", 指標="外国人比率", 単位="%",
                        備考="(未)：月報の要約ページに外国人比率の記載なし（延べ宿泊数のみ）"))
    return recs


# ---------------------------------------------------------------- 出力

def write_rows(path: Path, rows: list[dict], cols=None) -> None:
    path.parent.mkdir(exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols or list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)


def fmt_signed(v, unit="") -> str:
    return "(未)" if v in (None, "") else f"{float(v):+.1f}{unit}"


def map_occ_types(wide: list[dict], ym: str, out: Path, prefs) -> list[dict]:
    """4タイプの小地図＋主要都市圏のタイプ別表。図中に書いた数値を返す（検算用）。"""
    by = {(r["都道府県"], r["タイプ"]): r for r in wide if r["対象月"] == ym}
    types = ["旅館", "リゾート", "ビジネス", "シティ"]
    vals = [by[(p, t)]["OCC_前年差(pt)"] for (p, t) in by if t in types and p != "全国"]
    lim = min(15.0, max(10.0, round(max(abs(v) for v in vals if v is not None) + 0.5)))
    norm = TwoSlopeNorm(vmin=-lim, vcenter=0, vmax=lim)
    cmap = plt.get_cmap("RdYlGn")
    fig = plt.figure(figsize=(8, 10.4), dpi=100)
    y, m = ym.split("-")
    fig.text(0.02, 0.975, f"宿泊施設タイプ別 客室稼働率の前年同月差（{y}年{int(m)}月・第2次速報 vs {int(y) - 1}年{int(m)}月）",
             fontsize=11.5, weight="bold")
    shown = []
    for k, t in enumerate(types):
        cx, cy = (k % 2) * 0.5, 0.735 - (k // 2) * 0.24
        nat = by[("全国", t)]
        B.draw_japan(fig, [cx + 0.03, cy, 0.47, 0.22], prefs,
                     lambda n, t=t: "#dddddd" if by.get((n, t), {}).get("OCC_前年差(pt)") is None
                     else cmap(norm(by[(n, t)]["OCC_前年差(pt)"])))
        fig.text(cx + 0.02, cy + 0.205, f"{t}", fontsize=10.5, weight="bold")
        fig.text(cx + 0.02, cy + 0.187, f"全国 {nat['OCC_当月(%)']:.1f}%（{fmt_signed(nat['OCC_前年差(pt)'], 'pt')}）",
                 fontsize=7.5)
        shown.append({"図": out.name, "対象月": ym, "都道府県": "全国", "タイプ": t, "項目": "OCC_当月(%)",
                      "表示": f"{nat['OCC_当月(%)']:.1f}"})
        shown.append({"図": out.name, "対象月": ym, "都道府県": "全国", "タイプ": t, "項目": "OCC_前年差(pt)",
                      "表示": fmt_signed(nat["OCC_前年差(pt)"])})
    sm = plt.cm.ScalarMappable(norm=norm, cmap=cmap)
    cax = fig.add_axes([0.1, 0.478, 0.8, 0.011])
    cb = fig.colorbar(sm, cax=cax, orientation="horizontal", extend="both")
    cb.ax.tick_params(labelsize=7)
    fig.text(0.1, 0.447, f"色：客室稼働率の前年同月差（pt、±{lim:.0f}pt超は端の色）　緑=上昇 / 赤=低下 / 灰色=公表値なし（'-'）",
             fontsize=7.5)
    fig.text(0.02, 0.418, "主要都市圏のタイプ別 客室稼働率（当月%／前年同月差pt）", fontsize=8.5, weight="bold")
    rows = []
    for p in FOCUS:
        row = [p]
        for t in ["全体"] + types:
            r = by[(p, t)]
            a, d = r["OCC_当月(%)"], r["OCC_前年差(pt)"]
            txt = "(未)" if a is None else f"{a:.1f}  {fmt_signed(d)}"
            row.append(txt)
            shown.append({"図": out.name, "対象月": ym, "都道府県": p, "タイプ": t, "項目": "OCC_当月(%)",
                          "表示": "(未)" if a is None else f"{a:.1f}"})
            shown.append({"図": out.name, "対象月": ym, "都道府県": p, "タイプ": t, "項目": "OCC_前年差(pt)",
                          "表示": fmt_signed(d)})
        rows.append(row)
    B.draw_table(fig, [0.02, 0.105, 0.96, 0.305], ["", "全体", "旅館", "リゾート", "ビジネス", "シティ"], rows,
                 fontsize=7.6, col_widths=[0.15, 0.17, 0.17, 0.17, 0.17, 0.17])
    ax = fig.axes[-1]
    for (r, c), cell in ax.tables[0].get_celld().items():
        txt = cell.get_text().get_text()
        if r > 0 and c > 0:
            cell.get_text().set_color("#1a7f37" if "+" in txt and not txt.endswith("+0.0") else
                                      "#c62828" if re.search(r"\s-\d", txt) else "#333333")
    fig.text(0.02, 0.012,
             "出典：観光庁「宿泊旅行統計調査」推移表 " + MLIT_SUII + "（取得 " + RETRIEVED + "）。タイプは観光庁の区分のまま。\n"
             "注：2026年1月分から層化基準を従業者数→客室数に変更（前年差に見直しの影響を含み得る）。\n"
             "都道府県・タイプによっては標本が少なく誤差が大きい（例：山形県は標準誤差率20%超）。境界：dataofjapan/land",
             fontsize=6.2, color="#333333")
    fig.savefig(out, dpi=100)
    plt.close(fig)
    return shown


def svg_paths(prefs) -> dict:
    """都道府県の境界を簡略化して SVG パスにする。本土は 128.6～146.2E / 30.6～45.7N、沖縄は左上の枠に描く。"""
    k = 32.0
    W, H = (146.2 - 128.6) * k, (45.7 - 30.6) * k * 1.2
    ox, oy, ow = 8.0, 140.0, 165.0
    ks = ow / (128.6 - 122.8)

    def proj(lon, lat, okinawa):
        if okinawa:
            return ox + (lon - 122.8) * ks, oy + (27.2 - lat) * ks * 1.1
        return (lon - 128.6) * k, (45.7 - lat) * k * 1.2

    out = {}
    for name, geom in prefs.items():
        ok = name == "沖縄県"
        g = geom.simplify(0.01 if not ok else 0.004, preserve_topology=True)
        polys = [g] if g.geom_type == "Polygon" else list(g.geoms)
        d = []
        for p in polys:
            if p.area < (0.0008 if not ok else 0.0003):
                continue
            if ok and p.bounds[0] > 128.6:
                continue
            pts = [proj(x, y, ok) for x, y in p.exterior.coords]
            d.append("M" + "L".join(f"{x:.1f},{y:.1f}" for x, y in pts) + "Z")
        out[name] = "".join(d)
    return {"paths": out, "w": round(W), "h": round(H),
            "inset": [ox, oy, ow, round((27.2 - 23.9) * ks * 1.1, 1)]}


SOURCES = [
    # キー, 公表者, 資料, URL, 区分, 指標, 月の範囲, 状況, 理由
    ("M", "観光庁", "宿泊旅行統計調査（推移表・第2次速報・年の確定値・第1次速報）", B.CONFIG["mlit_page"],
     "都道府県×タイプ（旅館・リゾート・ビジネス・シティ・簡易宿所・会社団体）", "OCC、延べ・外国人延べ宿泊者数",
     "2025-07～2026-07（都道府県）、2026-08（全国のみ）", "取得", "ADR・RevPARは調査項目にない"),
    ("I", "いちごホテルリート", "ホテル運営状況（月次PDF）", "https://www.ichigo-hotel.co.jp/ir/news/",
     "ホテル別（27軒）", "OCC・ADR・RevPAR（当月・前年同月）", "2025-07～2026-08", "取得", ""),
    ("H", "星野リゾート・リート", "運営実績データ（Excel）", HOSHINO_XLSX,
     "ホテル別・投資区分（星のや・リゾナーレ・界・都市観光ほか）", "OCC・ADR・RevPAR", "2025-07～2026-08", "取得", ""),
    ("J", "ジャパン・ホテル・リート", "ホテル運営実績Excel（変動賃料等導入29ホテル・地域別）", JHR_XLSX,
     "地域別（8地域）", "OCC・ADR・RevPAR（当月と前年増減）", "2026-01～2026-08", "取得（2025年分は未）",
     "掲載Excelは2026年分のみ。ホテル別のOCC・ADRは非開示"),
    ("V", "インヴィンシブル投資法人", "月次運用実績（TDnet PDF）", B.CONFIG["reit_files"]["inv_pdf"],
     "地域別（7地域＋合計）", "OCC・ADR・RevPAR（前年同月比）", "2026-08のみ", "一部取得",
     "公式サイトが緊急メンテナンス（503）で過去分・ホテル別が取れない。TDnetは直近31日分のみ"),
    ("T", "森トラストリート", "ホテル運営実績（ウェブページ）", "https://www.mt-reit.jp/ja/portfolio/result.html",
     "ホテル別（2軒）", "OCC・ADR・RevPAR（当月のみ）", "2025-12～2026-07", "取得（前年比は未）", "前年同月値の掲載なし"),
    ("P", "ポラリス・ホールディングス", "月次経営レポート（Excel）", POLARIS_XLSX,
     "エリア別（6エリア・計68軒）と全体", "OCC・ADR・RevPAR・インバウンド比率（当年度と前年度の同月値）",
     "2026-04～2026-08（前年比あり）、2025-07・08（値のみ）", "取得",
     "2025年9月～2026年3月は掲載なし（2025年度版ExcelのURLは404）"),
    ("K", "京都市観光協会", "京都市観光協会データ月報", "https://www.kyokanko.or.jp/report/hotel202608",
     "京都市・ホテル（主要116～118施設）／旅館（主要21～23施設）", "ホテルOCC・ADR・RevPAR、旅館OCC（前年同月差・比）",
     "2025-07～2026-08", "取得（外国人比率は未）", "外国人比率は要約ページに記載なし（延べ宿泊数のみ）"),
    ("R", "いちご・星野（ホテル別開示を集計）", "都道府県別集計（客室数加重）", "（I・H の各URL）",
     "都道府県", "OCC・ADR・RevPAR", "2025-07～2026-08", "集計値", "前年同月と比べられるホテルのみ"),
    ("-", "共立メンテナンス（ドーミーイン）", "月次レポート", "https://www.kyoritsugroup.co.jp/Portals/0/pdf/ir/library/2026_8_monthly-report_jp.pdf",
     "事業セグメント", "売上高のみ", "", "未取得", "OCC・ADR・RevPARの記載なし（売上高と前年比のみ）"),
    ("-", "藤田観光", "IRライブラリ", "https://www.fujita-kanko.co.jp/ir/",
     "", "", "", "未取得", "月次のOCC・ADRの開示が見つからない（四半期決算資料のみ）"),
    ("-", "沖縄振興開発金融公庫", "県内主要ホテルの稼働状況", "https://www.okinawakouko.go.jp/report/research/re002/index.html",
     "沖縄県・ホテルタイプ別", "", "年度・半期", "未取得",
     "月次ではなく年度・半期の報告。PDFの文字がCIDフォントで抽出できず数値を確認できない"),
    ("-", "霞ヶ関ホテルリート", "", "https://www.kasumigaseki-hotelreit.co.jp/",
     "", "", "", "未取得", "ホスト名が解決できず接続不可"),
]


def build_html(recs: list[dict], svg: dict, out: Path) -> None:
    keys = ["出典キー", "対象月", "エリア区分", "エリア", "都道府県", "公表区分", "タイプ", "指標", "単位", "当月",
            "前年同月", "前年差・比", "差の単位", "差の根拠", "施設数", "客室数", "出典URL", "取得日", "備考", "公表者", "資料"]
    numeric = {"当月", "前年同月", "前年差・比", "施設数", "客室数"}
    strs: dict[str, int] = {}
    rows = []
    for r in recs:
        rows.append([r[k] if k in numeric else strs.setdefault(str(r[k]), len(strs)) for k in keys])
    payload = {"keys": keys, "rows": rows, "strs": list(strs), "months": MONTHS, "focus": FOCUS, "svg": svg,
               "sources": [dict(zip(["キー", "公表者", "資料", "URL", "区分", "指標", "月", "状況", "理由"], s)) for s in SOURCES],
               "retrieved": RETRIEVED}
    html = (HERE / "viewer_template.html").read_text(encoding="utf-8")
    html = html.replace("/*__DATA__*/null", json.dumps(payload, ensure_ascii=False, separators=(",", ":")))
    out.write_text(html, encoding="utf-8")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--refresh", action="store_true")
    args = ap.parse_args()
    r = args.refresh

    wide, recs = build_mlit_monthly(r)
    write_rows(DATA / "mlit_pref_type_monthly.csv", wide)
    ich, ich_recs = ichigo_monthly(r)
    hos, hos_recs = hoshino_monthly(r)
    agg = [x for ym in MONTHS for x in pref_aggregate({"I": ich, "H": hos}, ym)]
    recs += ich_recs + hos_recs + moritrust_monthly() + agg
    recs += jhr_monthly(r) + invincible_monthly(r) + polaris_monthly(r) + kyoto_monthly(r)
    write_rows(DATA / "monthly_all_sources.csv", recs, COLS)
    write_rows(DATA / "sources_catalog.csv",
               [dict(zip(["キー", "公表者", "資料", "URL", "区分", "指標", "対象月の範囲", "取得状況", "理由・制約"], s),
                     取得日=RETRIEVED) for s in SOURCES])

    metro = [w for w in wide if w["対象月"] == MLIT_LATEST and w["都道府県"] in FOCUS]
    metro.sort(key=lambda w: (FOCUS.index(w["都道府県"]), list(TYPES.values()).index(w["タイプ"])))
    write_rows(DATA / f"mlit_type_metro_{MLIT_LATEST}.csv", metro)

    B.setup_font()
    prefs = B.load_prefs(r)
    shown = map_occ_types(wide, MLIT_LATEST, HERE / f"map_occ_type_yoy_{MLIT_LATEST}.png", prefs)
    write_rows(DATA / f"png_values_map_occ_type_yoy_{MLIT_LATEST}.csv", shown)
    build_html(recs, svg_paths(prefs), HERE / "area_map_viewer.html")
    print(f"{len(recs)} 行（monthly_all_sources.csv）, {len(wide)} 行（mlit_pref_type_monthly.csv）")


if __name__ == "__main__":
    main()
