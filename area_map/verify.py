"""検算：地図・表・HTML・CSVの数値を、一次資料（raw/ にキャッシュした元ファイル）から別の読み方で取り直して突き合わせる。

    python3 verify.py      # build_area_map.py と build_monthly.py を実行したあとに使う

出力：data/verification_summary.csv（検査ごとの件数）、data/verification_mismatches.csv（不一致の明細）
一致の判定：転記値は完全一致。こちらで計算した差・比は、一次資料に公表の差・比があればそれと比べ、
小数第1位の丸め差（±0.1）は「丸め差」として別に数える。
"""
from __future__ import annotations

import csv
import json
import re
from collections import defaultdict
from pathlib import Path

import openpyxl
import pdfplumber

import build_monthly as M

DATA = M.DATA
results: dict[str, list] = defaultdict(lambda: [0, 0, 0])  # 件数, 一致, 丸め差
mismatches: list[dict] = []


def f(v):
    if v in (None, "", "-", "－"):
        return None
    try:
        return float(str(v).replace(",", "").replace("+", "").replace("％", "").replace("%", ""))
    except ValueError:
        return None


def check(name, label, got, expected, rounding_ok=False):
    g, e = f(got), f(expected)
    results[name][0] += 1
    if g is None and e is None or (g is not None and e is not None and abs(g - e) < 1e-9):
        results[name][1] += 1
    elif rounding_ok and g is not None and e is not None and abs(g - e) <= 0.1 + 1e-9:
        results[name][2] += 1
    else:
        mismatches.append({"検査": name, "対象": label, "CSV等の値": got, "一次資料の値": expected})


def load_csv(name):
    with (DATA / name).open(encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


ALL = load_csv("monthly_all_sources.csv")
WIDE = load_csv("mlit_pref_type_monthly.csv")
TYPE_COL = {"旅館": "旅館", "リゾート\nホテル": "リゾート", "ビジネス\nホテル": "ビジネス", "シティ\nホテル": "シティ"}


# ------------------------------------------------ 観光庁
def table8(url: str, month: int) -> dict:
    """第8表を見出しの文字列で列を探して読む（build_monthly とは別実装）。"""
    wb = openpyxl.load_workbook(M.get(url, False), read_only=True, data_only=True)
    ws = wb[[s for s in wb.sheetnames if re.fullmatch(rf"第8表\({month}月\)", s)][0]]
    rows = list(ws.iter_rows(values_only=True))
    hdr = next(r for r in rows if r and "旅館" in [str(x) for x in r])
    cols = {TYPE_COL[str(v)]: i for i, v in enumerate(hdr) if str(v) in TYPE_COL}
    out = {}
    for r in rows:
        s = str(r[0] or "")
        m = re.match(r"^[\s\u3000]*\d{2}(\S+)$", s)
        name = m.group(1) if m else ("全国" if re.fullmatch(r"令和\d+年\d+月", s.strip()) else None)
        if name:
            out[name] = {"全体": r[1], **{t: r[i] for t, i in cols.items()}}
    return out


def check_mlit():
    cache = {}
    for w in WIDE:
        for ym, col in ((w["対象月"], "OCC_当月(%)"), (w["前年同月"], "OCC_前年同月(%)")):
            url = M.people_source(ym)
            key = (url, int(ym[5:]))
            if key not in cache:
                cache[key] = table8(url, int(ym[5:]))
            check("観光庁OCC：推移表→第8表（別ファイル）", f"{ym} {w['都道府県']} {w['タイプ']}", w[col],
                  cache[key].get(w["都道府県"], {}).get(w["タイプ"]))

    suii = openpyxl.load_workbook(M.get(M.MLIT_SUII, False), read_only=True, data_only=True)

    def suii_series(sheet):
        rows = list(suii[sheet].iter_rows(values_only=True))
        years, months = rows[2], rows[3]
        cols, cur = {}, None
        for i, mm in enumerate(months):
            cur = years[i] or cur
            if cur and mm and re.match(r"令和\d+年", str(cur)):
                cols[f"{2018 + int(re.search(r'\d+', str(cur)).group())}-{int(str(mm)[:-1]):02d}"] = i
        out = {}
        for r in rows[4:]:
            s = str(r[0] or "")
            m = re.match(r"^[\s\u3000]*\d{2}(\S+)$", s)
            name = m.group(1) if m else ("全国" if "全" in s and "国" in s else None)
            if name:
                out[name] = {ym: r[i] for ym, i in cols.items()}
        return out

    tot = {**{k: v for k, v in suii_series("旧1-2").items()}}
    tot_new = suii_series("1-1")
    fr, fr_new = suii_series("旧3-2"), suii_series("3-1")
    for w in WIDE:
        if w["タイプ"] != "全体":
            continue
        p = w["都道府県"]
        for ym, a, b in ((w["対象月"], "延べ宿泊者数_当月(人泊)", "外国人延べ_当月(人泊)"),
                         (w["前年同月"], "延べ宿泊者数_前年同月(人泊)", "外国人延べ_前年同月(人泊)")):
            src_t, src_f = (tot_new, fr_new) if ym >= "2026-01" else (tot, fr)
            check("観光庁人泊数：第4表→推移表（別ファイル）", f"{ym} {p} 延べ", w[a], src_t.get(p, {}).get(ym))
            check("観光庁人泊数：第4表→推移表（別ファイル）", f"{ym} {p} 外国人", w[b], src_f.get(p, {}).get(ym))

    # 報道発表資料（観光庁自身が載せた当月値・前年同月差・外国人比率）
    sign = {"＋": 1, "+": 1, "－": -1, "-": -1, "±": 0}
    for ym, url in M.MLIT_PRESS.items():
        with pdfplumber.open(M.get(url, False)) as pdf:
            pages = [p.extract_text() or "" for p in pdf.pages]
        text = "\n".join(pages)
        prelim = re.search(r"2025年1月～2026年\d+月の数値は第2次速報値", text) is not None
        occ = {}
        for t in pages:
            if "都道府県別宿泊施設タイプ別客室稼働率（" in t:
                for m in re.finditer(r"^(\S+?[都道府県]|全国) ([\d.]+) (?:\d+|-) ([＋－±+\-])([\d.]+)", t, re.M):
                    occ[m.group(1)] = (float(m.group(2)), sign[m.group(3)] * float(m.group(4)))
        foreign, _ = M.B._press_official(M.get(url, False))
        diff_name = ("（参考）観光庁：前年差→報道発表資料（前年値が2025年第2次速報だった時点の公表）" if prelim
                     else "観光庁：前年差→報道発表資料の公表値")
        for w in WIDE:
            if w["対象月"] != ym or w["タイプ"] != "全体":
                continue
            p = w["都道府県"]
            if p in occ:
                check("観光庁：当月OCC・外国人比率→報道発表資料の公表値", f"{ym} {p} OCC当月", w["OCC_当月(%)"], occ[p][0])
                check(diff_name, f"{ym} {p} OCC前年差", w["OCC_前年差(pt)"], occ[p][1], rounding_ok=True)
            if p in foreign:
                check("観光庁：当月OCC・外国人比率→報道発表資料の公表値", f"{ym} {p} 外国人比率", w["外国人比率_当月(%)"],
                      foreign[p][1], rounding_ok=True)
                a, b = f(w["外国人延べ_当月(人泊)"]), f(w["外国人延べ_前年同月(人泊)"])
                check(diff_name, f"{ym} {p} 外国人前年比", round((a / b - 1) * 100, 1) if a and b else None,
                      foreign[p][0], rounding_ok=True)


# ------------------------------------------------ いちご（PDF本文の行から読み直す）
def check_ichigo():
    dash = "[－–—]"
    pat = {k: re.compile(rf"{k}\s*（[^）]*）\s*([\d,.]+|{dash})\s+([\d,.]+|{dash})\s+([+\-]?[\d,.]+|{dash})") for k in ("RevPAR", "客室稼働率", "ADR")}
    for ym, url in M.ICHIGO_PDFS.items():
        with pdfplumber.open(M.get(url, False)) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)
        body = text[text.index("ホテル／項目"):]
        seq = {k: [m.groups() for m in p.finditer(body)] for k, p in pat.items()}
        recs = [r for r in ALL if r["出典キー"] == "I" and r["対象月"] == ym]
        hotels = list(dict.fromkeys(r["エリア"] for r in recs))
        for k, metric in (("RevPAR", "RevPAR"), ("客室稼働率", "OCC"), ("ADR", "ADR")):
            for i, h in enumerate(hotels):
                r = next(x for x in recs if x["エリア"] == h and x["指標"] == metric)
                a, b, d = seq[k][i] if i < len(seq[k]) else (None, None, None)
                a, b = (None if re.fullmatch(dash, v or "") else v for v in (a, b))
                # PDFは当月と前年同月が同値のとき差額欄を「－」で載せる
                if d and re.fullmatch(dash, d):
                    d = "0" if a is not None and a == b else None
                cur = r["当月"] or ("0" if f(a) == 0 else "")
                prv = r["前年同月"] or ("0" if f(b) == 0 else "")
                check("いちご：表の読み取り→PDF本文の行", f"{ym} {h} {metric} 当月", cur, a)
                check("いちご：表の読み取り→PDF本文の行", f"{ym} {h} {metric} 前年同月", prv, b)
                if metric == "OCC" and r["前年差・比"]:
                    check("いちご：計算した前年差→PDF記載の差額", f"{ym} {h} OCC差", r["前年差・比"], d, rounding_ok=True)


# ------------------------------------------------ 星野（セル座標で読み直す）
def check_hoshino():
    wb = openpyxl.load_workbook(M.get(M.HOSHINO_XLSX, False), data_only=True)
    grid = {}
    for ws in wb.worksheets:
        if not re.match(r"第\d+期開示", ws.title.strip()):
            continue
        year_cells = {c.column: str(c.value)[:4] for c in ws[4] if c.value and re.match(r"\d{4}年", str(c.value))}
        col_month, cur = {}, None
        for c in ws[5]:
            cur = year_cells.get(c.column, cur)
            if c.value and re.match(r"\d+月$", str(c.value)) and cur:
                col_month[c.column] = f"{cur}-{int(str(c.value)[:-1]):02d}"
        code = None
        for row in ws.iter_rows(min_row=6):
            code = str(row[3].value).strip() if row[3].value else code
            metric = row[5].value
            for c in row:
                if c.column in col_month and code and metric:
                    grid.setdefault((code, col_month[c.column], metric), (c.value, ws.title.strip()))
    name_to_code = {}
    for ws in wb.worksheets:
        if re.match(r"第\d+期開示", ws.title.strip()):
            for row in ws.iter_rows(min_row=6, values_only=True):
                if row[3] and row[4]:
                    name_to_code.setdefault(str(row[4]).strip(), str(row[3]).strip())
    key = {"OCC": "客室稼働率", "ADR": "ADR（円）", "RevPAR": "RevPAR（円）"}
    for r in (x for x in ALL if x["出典キー"] == "H"):
        code = name_to_code[r["エリア"]]
        for col, ym in (("当月", r["対象月"]), ("前年同月", r["比較対象月"])):
            v = grid.get((code, ym, key[r["指標"]]), (None, None))[0]
            v = f(v)
            if v is not None and r["指標"] == "OCC":
                v = round(v * 100, 1)
            if v == 0:
                v = None
            check("星野：Excel→セル座標で再読", f"{r['対象月']} {r['エリア']} {r['指標']} {col}", r[col], v)


# ------------------------------------------------ JHR・インヴィンシブル・ポラリス・京都
def check_jhr():
    ws = openpyxl.load_workbook(M.get(M.JHR_XLSX, False), data_only=True)["変動賃料導入29ホテル（地域別）"]
    cells = {}
    region = metric = None
    for row in ws.iter_rows():
        b, d = row[1].value, row[3].value
        if b and d is None and not any(c.value not in (None, "") for c in row[4:16]):
            region = b
        if b in ("客室稼働率（％）", "ADR（円）", "RevPAR（円）"):
            metric = b
        elif b:
            metric = None
        if metric and d in ("2026年", "増減", "増減率"):
            for c in row[4:16]:
                cells[(region, metric, d, c.column - 4)] = c.value
    mk = {"OCC": "客室稼働率（％）", "ADR": "ADR（円）", "RevPAR": "RevPAR（円）"}
    for r in (x for x in ALL if x["出典キー"] == "J" and x["当月"] != ""):
        m = int(r["対象月"][5:])
        cur = cells[(r["エリア"], mk[r["指標"]], "2026年", m)]
        chg = cells[(r["エリア"], mk[r["指標"]], "増減" if r["指標"] == "OCC" else "増減率", m)]
        check("JHR：Excel→セル座標で再読", f"{r['対象月']} {r['エリア']} {r['指標']} 当月", r["当月"],
              round(cur * 100, 1) if r["指標"] == "OCC" else cur)
        check("JHR：Excel→セル座標で再読", f"{r['対象月']} {r['エリア']} {r['指標']} 前年差・比", r["前年差・比"],
              chg if r["指標"] == "OCC" else round(chg * 100, 1))


def check_invincible():
    for ym, url in M.INV_PDFS.items():
        with pdfplumber.open(M.B.fetch("invincible_monthly", url, False)) as pdf:
            lines = "\n".join(p.extract_text() or "" for p in pdf.pages).splitlines()
        for r in (x for x in ALL if x["出典キー"] == "V" and x["対象月"] == ym and x["当月"] != ""):
            line = next((l for l in lines if l.startswith(r["エリア"] + " ")), "")[len(r["エリア"]) + 1:]
            toks = re.findall(r"[+△]?[\d,]+(?:\.\d+)?", line)
            nums = [(-1 if t.startswith("△") else 1) * float(t.lstrip("+△").replace(",", "")) for t in toks]
            pos = {"OCC": 0, "ADR": 2, "RevPAR": 4}[r["指標"]]
            check("インヴィンシブル：PDF本文の地域行と照合", f"{ym} {r['エリア']} {r['指標']} 当月", r["当月"],
                  nums[pos] if len(nums) > pos else None)
            check("インヴィンシブル：PDF本文の地域行と照合", f"{ym} {r['エリア']} {r['指標']} 前年差・比", r["前年差・比"],
                  nums[pos + 1] if len(nums) > pos + 1 else None)


def check_polaris():
    wb = openpyxl.load_workbook(M.get(M.POLARIS_XLSX, False), data_only=True)
    ws = wb["エリア別_単月"]
    yr = next(c.row for c in ws["C"] if c.value == "単月")
    colym, cur = {}, None
    for c in ws[yr + 1]:
        y = ws.cell(yr, c.column).value
        cur = y if isinstance(y, int) else cur
        if c.value and re.match(r"\d+月$", str(c.value)) and cur:
            colym[f"{cur}-{int(str(c.value)[:-1]):02d}"] = c.column
    area_row = {}
    for c in ws["B"]:
        if c.value and "ホテル" in str(c.value):
            area_row[str(c.value).split("：")[0]] = c.row
    off = {"OCC": 1, "ADR": 2, "RevPAR": 3, "外国人比率": 4}
    for r in (x for x in ALL if x["出典キー"] == "P" and x["エリア区分"] == "公表地域" and x["当月"] != ""):
        row = area_row[r["エリア"]] + off[r["指標"]]
        for col, ym in (("当月", r["対象月"]), ("前年同月", r["比較対象月"])):
            v = ws.cell(row, colym[ym]).value if ym in colym else None
            if v is not None and r["単位"] == "%":
                v = round(v * 100, 1)
            check("ポラリス：Excel→セル座標で再読", f"{r['対象月']} {r['エリア']} {r['指標']} {col}", r[col], v)
    ws = wb["全体"]
    lab = {"OCC": "客室稼働率", "ADR": "ADR", "RevPAR": "RevPAR", "外国人比率": "インバウンド比率"}
    hdr = next(c.row for c in ws["D"] if c.value == "単月")
    mcol = {int(str(c.value)[:-1]): c.column for c in ws[hdr] if c.value and re.match(r"\d+月$", str(c.value))}
    for r in (x for x in ALL if x["出典キー"] == "P" and x["エリア区分"] == "全国" and x["当月"] != ""):
        mrow = next(c.row for c in ws["B"] if c.row > hdr and c.value == lab[r["指標"]])
        y, m = int(r["対象月"][:4]), int(r["対象月"][5:])
        fy = y + 1 if m >= 4 else y
        frow = next(c.row for c in ws["C"] if c.row > mrow and c.value == f"{fy}年3月期")
        v = ws.cell(frow, mcol[m]).value
        v = round(v * 100, 1) if r["単位"] == "%" else round(v)
        check("ポラリス：Excel→セル座標で再読", f"{r['対象月']} 全体 {r['指標']} 当月", r["当月"], v)


def check_kyoto():
    """カードの数値ではなく本文の文（「…は17,272円となり、前年同月の17,325円から…」）から読み直す。"""
    for ym in M.MONTHS:
        url = M.KYOTO_PAGE.format(ym=ym.replace("-", ""))
        t = M._kyoto_text(M.get_html(url, False))
        t = t[t.index("調査結果のポイント"):]
        sent = {
            "OCC|ホテル": re.search(r"主要ホテル ?\d+ ?施設(?:における|の) ?\d{4}年\d+月の客室稼働率は ?([\d.]+) ?[％%]", t),
            "ADR|ホテル": re.search(r"平均客室単価※?は ?([\d,]+) ?円となり", t),
            "RevPAR|ホテル": re.search(r"客室収益指数(?:※|（[^）]*）)?は ?([\d,]+) ?円となり", t),
            "OCC|旅館": re.search(r"主要旅館 ?\d+ ?施設における ?\d+月の稼働率は ?([\d.]+) ?[％%]", t),
        }
        for r in (x for x in ALL if x["出典キー"] == "K" and x["対象月"] == ym and x["当月"] != ""):
            k = f"{r['指標']}|{'旅館' if r['公表区分'].startswith('旅館') else 'ホテル'}"
            m = sent.get(k)
            check("京都市観光協会：カード→本文の文で再読", f"{ym} {k} 当月", r["当月"], m.group(1) if m else None)


# ------------------------------------------------ 集計・HTML・PNG・旧ファイル
def check_aggregate():
    hot = defaultdict(dict)
    for r in ALL:
        if r["出典キー"] in ("I", "H") and r["当月"] and r["前年同月"] and r["都道府県"] and r["客室数"]:
            hot[(r["対象月"], r["出典キー"], r["エリア"])][r["指標"]] = r
    groups = defaultdict(list)
    for (ym, _, _), v in hot.items():
        if len(v) == 3:
            groups[(ym, v["OCC"]["都道府県"])].append(v)
    for r in (x for x in ALL if x["出典キー"] == "R"):
        hs = groups[(r["対象月"], r["都道府県"])]
        rooms = sum(float(h["OCC"]["客室数"]) for h in hs)
        w = lambda met, col: sum(float(h[met][col]) * float(h["OCC"]["客室数"]) for h in hs) / rooms
        occ, occp, rev, revp = w("OCC", "当月"), w("OCC", "前年同月"), w("RevPAR", "当月"), w("RevPAR", "前年同月")
        exp = {"OCC": (round(occ, 1), round(occ - occp, 1)),
               "RevPAR": (round(rev), round((rev / revp - 1) * 100, 1)),
               "ADR": (round(rev / (occ / 100)), round(((rev / (occ / 100)) / (revp / (occp / 100)) - 1) * 100, 1))}
        check("都道府県集計：ホテル別CSVから再計算", f"{r['対象月']} {r['都道府県']} {r['指標']} 軒数", r["施設数"], len(hs))
        check("都道府県集計：ホテル別CSVから再計算", f"{r['対象月']} {r['都道府県']} {r['指標']} 当月", r["当月"], exp[r["指標"]][0])
        check("都道府県集計：ホテル別CSVから再計算", f"{r['対象月']} {r['都道府県']} {r['指標']} 差・比", r["前年差・比"],
              exp[r["指標"]][1])


def check_html():
    html = (M.HERE / "area_map_viewer.html").read_text(encoding="utf-8")
    payload = json.loads(re.search(r"const D = (\{.*?\});\n", html, re.S).group(1))
    keys, strs = payload["keys"], payload["strs"]
    numeric = {"当月", "前年同月", "前年差・比", "施設数", "客室数"}
    rows = [{k: (v if k in numeric else strs[v]) for k, v in zip(keys, row)} for row in payload["rows"]]
    results["HTML埋め込みデータ→CSV"]  # 行数の確認も兼ねる
    if len(rows) != len(ALL):
        mismatches.append({"検査": "HTML埋め込みデータ→CSV", "対象": "行数", "CSV等の値": len(rows), "一次資料の値": len(ALL)})
    for h, c in zip(rows, ALL):
        for k in ("当月", "前年同月", "前年差・比", "施設数", "客室数"):
            if c[k] != "" or h[k] != "":
                check("HTML埋め込みデータ→CSV", f"{c['出典キー']} {c['対象月']} {c['エリア']} {c['指標']} {k}", h[k], c[k])
        for k in ("出典URL", "取得日", "対象月", "エリア", "指標"):
            results["HTML埋め込みデータ→CSV（文字列）"][0] += 1
            if str(h[k]) == c[k]:
                results["HTML埋め込みデータ→CSV（文字列）"][1] += 1
            else:
                mismatches.append({"検査": "HTML埋め込みデータ→CSV（文字列）", "対象": f"{c['エリア']} {k}", "CSV等の値": h[k],
                                   "一次資料の値": c[k]})


def check_png():
    wide = {(w["対象月"], w["都道府県"], w["タイプ"]): w for w in WIDE}
    t8 = table8(M.MLIT_SOKUHO2[M.MLIT_LATEST], int(M.MLIT_LATEST[5:]))
    t8p = table8(M.people_source(M.prev_month(M.MLIT_LATEST)), int(M.MLIT_LATEST[5:]))
    for r in load_csv(f"png_values_map_occ_type_yoy_{M.MLIT_LATEST}.csv"):
        w = wide[(r["対象月"], r["都道府県"], r["タイプ"])]
        check("タイプ別地図PNGの表示値→CSV", f"{r['都道府県']} {r['タイプ']} {r['項目']}", r["表示"], w[r["項目"]])
        a, b = f(t8[r["都道府県"]][r["タイプ"]]), f(t8p[r["都道府県"]][r["タイプ"]])
        exp = a if r["項目"] == "OCC_当月(%)" else (round(a - b, 1) if a is not None and b is not None else None)
        check("タイプ別地図PNGの表示値→第8表（一次資料）", f"{r['都道府県']} {r['タイプ']} {r['項目']}", r["表示"], exp)
    old = {r["都道府県"]: r for r in load_csv(f"mlit_prefecture_{M.MLIT_LATEST}.csv")}
    for p, r in old.items():
        w = wide[(M.MLIT_LATEST, p, "全体")]
        check("前回の地図CSV（2026-07）→今回の月次CSV", f"{p} OCC前年差", r["OCC_全体_前年差(pt)"], w["OCC_前年差(pt)"])
        check("前回の地図CSV（2026-07）→今回の月次CSV", f"{p} 外国人比率", r["外国人比率_当月(%)"], w["外国人比率_当月(%)"])
        if r["検算_公表_OCC前年同月差(pt)"]:
            check("前回の地図CSV（2026-07）→報道発表資料", f"{p} OCC前年差", r["OCC_全体_前年差(pt)"],
                  r["検算_公表_OCC前年同月差(pt)"], rounding_ok=True)
        if r["検算_公表_外国人比率(%)"]:
            check("前回の地図CSV（2026-07）→報道発表資料", f"{p} 外国人比率", r["外国人比率_当月(%)"], r["検算_公表_外国人比率(%)"],
                  rounding_ok=True)
            check("前回の地図CSV（2026-07）→報道発表資料", f"{p} 外国人前年比", r["外国人延べ_前年同月比(%)"],
                  r["検算_公表_外国人前年同月比(%)"], rounding_ok=True)
    new = {(r["エリア"], r["指標"]): r for r in ALL if r["出典キー"] in ("I", "H") and r["対象月"] == "2026-08"}
    for r in load_csv("jreit_hotels_2026-08.csv"):
        for met, col, ccol in (("OCC", "OCC_当月(%)", "当月"), ("ADR", "ADR_当月(円)", "当月"), ("RevPAR", "RevPAR_当月(円)", "当月"),
                               ("OCC", "OCC_前年同月(%)", "前年同月"), ("ADR", "ADR_前年同月(円)", "前年同月")):
            v = r[col] if f(r[col]) else ""
            check("前回のJ-REIT CSV（TDnet版PDF）→今回の月次CSV（自社サイト版PDF）", f"{r['ホテル']} {col}", v,
                  new.get((r["ホテル"], met), {}).get(ccol, "(行なし)"))
    agg = {(r["都道府県"], r["指標"]): r for r in ALL if r["出典キー"] == "R" and r["対象月"] == "2026-08"}
    for r in load_csv("jreit_area_reference_2026-08.csv"):
        check("前回の都道府県集計CSV→今回の集計", f"{r['都道府県']} RevPAR比", r["加重_RevPAR_前年比(%)"],
              agg[(r["都道府県"], "RevPAR")]["前年差・比"])
        check("前回の都道府県集計CSV→今回の集計", f"{r['都道府県']} ADR比", r["加重_ADR_前年比(%)"], agg[(r["都道府県"], "ADR")]["前年差・比"])
    reg = {(r["出典キー"], r["対象月"], r["エリア"], r["指標"]): r for r in ALL if r["出典キー"] in ("J", "V")}
    for r in load_csv("jreit_regions_reported.csv"):
        k = "J" if r["REIT"].startswith("ジャパン") else "V"
        for met, col in (("OCC", "OCC_前年差(pt)"), ("ADR", "ADR_前年比(%)"), ("RevPAR", "RevPAR_前年比(%)")):
            check("前回の地域別CSV→今回の月次CSV", f"{k} {r['対象月']} {r['地域']} {met}", r[col],
                  reg[(k, r["対象月"], r["地域"], met)]["前年差・比"])


def main():
    for fn in (check_mlit, check_ichigo, check_hoshino, check_jhr, check_invincible, check_polaris, check_kyoto,
               check_aggregate, check_html, check_png):
        fn()
    summary = [{"検査": k, "比較件数": v[0], "完全一致": v[1], "丸め差(±0.1)": v[2], "不一致": v[0] - v[1] - v[2]}
               for k, v in results.items()]
    M.write_rows(DATA / "verification_summary.csv", summary)
    M.write_rows(DATA / "verification_mismatches.csv",
                 mismatches or [{"検査": "", "対象": "（不一致なし）", "CSV等の値": "", "一次資料の値": ""}])
    for s in summary:
        print(f"{s['検査']}: {s['比較件数']}件 一致{s['完全一致']} 丸め差{s['丸め差(±0.1)']} 不一致{s['不一致']}")
    tot = [sum(s[k] for s in summary) for k in ("比較件数", "完全一致", "丸め差(±0.1)", "不一致")]
    print(f"合計: {tot[0]}件 一致{tot[1]} 丸め差{tot[2]} 不一致{tot[3]}")


if __name__ == "__main__":
    main()
