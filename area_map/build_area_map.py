"""エリア別ホテル指標（観光庁 宿泊旅行統計調査 + ホテル系J-REIT月次開示）を集計し、
CSV と都道府県別コロプレス地図（PNG）を作る。

    python3 build_area_map.py            # raw/ にキャッシュがあれば再利用
    python3 build_area_map.py --refresh  # 元ファイルを取り直す

対象月や元ファイルのURLは下の CONFIG を書き換える（README の更新手順を参照）。
"""
from __future__ import annotations

import argparse
import csv
import json
import re
from dataclasses import dataclass
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl
import pdfplumber
import requests
from matplotlib import font_manager
from matplotlib.colors import TwoSlopeNorm
from PIL import Image
from shapely.geometry import shape

HERE = Path(__file__).parent
RAW = HERE / "raw"
DATA = HERE / "data"

RETRIEVED = "2026-10-26"  # 取得日（JST）

CONFIG = {
    # 観光庁：都道府県別の最新月（第2次速報）と比較する前年同月
    "mlit_month": (2026, 7),
    "mlit_files": {
        "suii": "https://www.mlit.go.jp/kankocho/content/002026697.xlsx",      # 推移表
        "press": "https://www.mlit.go.jp/kankocho/content/002025565.pdf",      # 報道発表資料（7月2次・8月1次）
        "sokuho1": "https://www.mlit.go.jp/kankocho/content/002025567.xlsx",   # 8月1次速報 集計結果
    },
    "mlit_page": "https://www.mlit.go.jp/kankocho/tokei_hakusyo/shukuhakutokei.html",
    # J-REIT：最新の共通月
    "reit_month": (2026, 8),
    "reit_files": {
        "jhr_xlsx": "https://www.jhrth.co.jp/file/review_exceldl-973246b76834e8803f8659c458779d6d41f9cd10.xlsx",
        "inv_pdf": "https://www.release.tdnet.info/inbs/140120260925540060.pdf",
        "ichigo_pdf": "https://www.release.tdnet.info/inbs/140120260925540031.pdf",
        "ichigo_xlsx": "https://www.ichigo-hotel.co.jp/pdf/ir/library/2026/IchigoHotel_20260917_July_2026_Fiscal_Period_Hotel_Data_JPN.xlsx",
        "hoshino_xlsx": "https://ssl4.eir-parts.net/doc/3287/ir_material_for_fiscal_ym5/195377/00.xlsx",
        "geojson": "https://raw.githubusercontent.com/dataofjapan/land/master/japan.geojson",
    },
    "hoshino_sheets": {"current": "第27期開示", "prior": "第26期開示"},
}

ERA = {2025: "令和7年", 2026: "令和8年"}
TYPES = ["計", "旅館", "リゾート\nホテル", "ビジネス\nホテル", "シティ\nホテル", "簡易宿所"]
TYPE_LABEL = {"計": "全体", "旅館": "旅館", "リゾート\nホテル": "リゾート", "ビジネス\nホテル": "ビジネス",
              "シティ\nホテル": "シティ", "簡易宿所": "簡易宿所"}


def fetch(key: str, url: str, refresh: bool) -> Path:
    RAW.mkdir(exist_ok=True)
    path = RAW / f"{key}{Path(url).suffix}"
    if refresh or not path.exists():
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=120)
        r.raise_for_status()
        path.write_bytes(r.content)
    return path


def write_csv(name: str, rows: list[dict]) -> None:
    DATA.mkdir(exist_ok=True)
    with (DATA / name).open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)


def num(v):
    if v is None or v == "" or v == "-" or v == "－":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("％", "").replace("%", "").replace("＋", "+").replace("－", "-").replace("△", "-")
    try:
        return float(s)
    except ValueError:
        return None


def pref_name(label: str) -> str:
    return re.sub(r"^[\s\u3000]*\d{2}", "", str(label)).strip()


# ---------------------------------------------------------------- 観光庁

def _month_col(ws, year: int, month: int, header_row: int, month_row: int) -> int:
    rows = list(ws.iter_rows(min_row=1, max_row=month_row + 1, values_only=True))
    years, months = rows[header_row], rows[month_row]
    cur = None
    for i, m in enumerate(months):
        if years[i]:
            cur = years[i]
        if cur == ERA[year] and m == f"{month}月":
            return i
    raise ValueError(f"{ws.title}: {year}年{month}月の列が見つからない")


def _read_simple(ws, col: int) -> dict[str, float]:
    out = {}
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r[0] and re.match(r"^[\s\u3000]*(\d{2}|全)", str(r[0])):
            label = "全国" if "全" in str(r[0]) and "国" in str(r[0]) else pref_name(r[0])
            out[label] = num(r[col])
    return out


def _read_occ(ws, col: int) -> dict[tuple[str, str], float]:
    out, cur = {}, None
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r[0] and str(r[0]).strip():
            cur = "全国" if "全" in str(r[0]) and "国" in str(r[0]) else pref_name(r[0])
        if cur and r[1] in TYPES:
            out[(cur, r[1])] = num(r[col])
    return out


def _press_official(pdf_path: Path) -> tuple[dict, dict]:
    """報道発表資料の都道府県表（観光庁自身が計算した前年同月比・差）を読む。検算用。"""
    foreign, occ = {}, {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "都道府県別外国人延べ宿泊者数（" in text:
                for m in re.finditer(r"^(\S+?[都道府県]|全国) ([\d,]+) ([+\-][\d.]+)% ([\d.]+)%", text, re.M):
                    foreign[m.group(1)] = (num(m.group(3)), num(m.group(4)))
            if "都道府県別宿泊施設タイプ別客室稼働率（" in text:
                for m in re.finditer(r"^(\S+?[都道府県]) ([\d.]+) \d+ ([＋－+\-][\d.]+)", text, re.M):
                    occ[m.group(1)] = num(m.group(3))
                m = re.search(r"^全国 ([\d.]+) - ([＋－+\-][\d.]+)", text, re.M)
                if m:
                    occ["全国"] = num(m.group(2))
    return foreign, occ


def build_mlit(refresh: bool) -> list[dict]:
    files = {k: fetch(f"mlit_{k}", u, refresh) for k, u in CONFIG["mlit_files"].items()}
    y, m = CONFIG["mlit_month"]
    wb = openpyxl.load_workbook(files["suii"], read_only=True, data_only=True)

    def new_col(ws):  # 2026年以降のシート（年見出し1行目=行3、月=行4）
        rows = list(ws.iter_rows(min_row=4, max_row=4, values_only=True))[0]
        return list(rows).index(f"{m}月")

    occ_now = _read_occ(wb["4-1"], new_col(wb["4-1"]))
    occ_prev = _read_occ(wb["旧4-2"], _month_col(wb["旧4-2"], y - 1, m, 2, 3))
    tot_now = _read_simple(wb["1-1"], new_col(wb["1-1"]))
    tot_prev = _read_simple(wb["旧1-2"], _month_col(wb["旧1-2"], y - 1, m, 2, 3))
    fr_now = _read_simple(wb["3-1"], new_col(wb["3-1"]))
    fr_prev = _read_simple(wb["旧3-2"], _month_col(wb["旧3-2"], y - 1, m, 2, 3))
    off_foreign, off_occ = _press_official(files["press"])

    src_x = CONFIG["mlit_files"]["suii"]
    src_p = CONFIG["mlit_files"]["press"]
    rows = []
    for pref in tot_now:
        row = {"都道府県": pref, "対象月": f"{y}-{m:02d}", "前年同月": f"{y - 1}-{m:02d}"}
        for t in TYPES:
            a, b = occ_now.get((pref, t)), occ_prev.get((pref, t))
            lab = TYPE_LABEL[t]
            row[f"OCC_{lab}_当月(%)"] = a
            row[f"OCC_{lab}_前年同月(%)"] = b
            row[f"OCC_{lab}_前年差(pt)"] = round(a - b, 1) if a is not None and b is not None else None
        tn, tp, fn, fp = tot_now.get(pref), tot_prev.get(pref), fr_now.get(pref), fr_prev.get(pref)
        row["延べ宿泊者数_当月(人泊)"] = tn
        row["延べ宿泊者数_前年同月(人泊)"] = tp
        row["外国人延べ宿泊者数_当月(人泊)"] = fn
        row["外国人延べ宿泊者数_前年同月(人泊)"] = fp
        row["外国人延べ_前年同月比(%)"] = round((fn / fp - 1) * 100, 1) if fn and fp else None
        row["外国人比率_当月(%)"] = round(fn / tn * 100, 1) if fn and tn else None
        row["外国人比率_前年同月(%)"] = round(fp / tp * 100, 1) if fp and tp else None
        row["外国人比率_前年差(pt)"] = (round(row["外国人比率_当月(%)"] - row["外国人比率_前年同月(%)"], 1)
                                if row["外国人比率_当月(%)"] is not None and row["外国人比率_前年同月(%)"] is not None else None)
        of = off_foreign.get(pref)
        row["検算_公表_OCC前年同月差(pt)"] = off_occ.get(pref)
        row["検算_公表_外国人前年同月比(%)"] = of[0] if of else None
        row["検算_公表_外国人比率(%)"] = of[1] if of else None
        row["出典_数値"] = src_x
        row["出典_検算"] = src_p
        row["取得日"] = RETRIEVED
        rows.append(row)

    nat = json_national_aug(files["sokuho1"])
    write_csv(f"mlit_prefecture_{y}-{m:02d}.csv", rows)
    return rows, nat


def json_national_aug(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = next(s for s in wb.sheetnames if s.startswith("第5表"))
    vals = list(wb[sheet].iter_rows(min_row=7, max_row=7, values_only=True))[0]
    return {"月": vals[0], "OCC全体": vals[1]}


# ---------------------------------------------------------------- J-REIT

@dataclass
class Hotel:
    reit: str
    name: str
    pref: str
    rooms: int | None
    occ: float | None
    occ_py: float | None
    adr: float | None
    adr_py: float | None
    revpar: float | None
    revpar_py: float | None
    src: str

    def ok(self) -> bool:
        return all(v for v in (self.occ, self.occ_py, self.adr, self.adr_py, self.revpar, self.revpar_py))

    def row(self, month: str) -> dict:
        pct = lambda a, b: round((a / b - 1) * 100, 1) if a and b else None
        return {
            "REIT": self.reit, "ホテル": self.name, "都道府県": self.pref, "客室数": self.rooms, "対象月": month,
            "OCC_当月(%)": self.occ, "OCC_前年同月(%)": self.occ_py,
            "OCC_前年差(pt)": round(self.occ - self.occ_py, 1) if self.occ is not None and self.occ_py is not None else None,
            "ADR_当月(円)": self.adr, "ADR_前年同月(円)": self.adr_py, "ADR_前年比(%)": pct(self.adr, self.adr_py),
            "RevPAR_当月(円)": self.revpar, "RevPAR_前年同月(円)": self.revpar_py, "RevPAR_前年比(%)": pct(self.revpar, self.revpar_py),
            "集計対象": "対象" if self.ok() else "対象外（当月か前年同月の値が無い・0＝休館）",
            "出典URL": self.src, "取得日": RETRIEVED,
        }


def _norm(s: str) -> str:
    s = re.sub(r"※\d+", "", s or "")
    return re.sub(r"[\s\u3000（）()]", "", s)


def ichigo_hotels(refresh: bool) -> list[Hotel]:
    pdf_path = fetch("ichigo_monthly", CONFIG["reit_files"]["ichigo_pdf"], refresh)
    lookup = ichigo_lookup(refresh)
    return parse_ichigo_pdf(pdf_path, CONFIG["reit_files"]["ichigo_pdf"], lookup)


def ichigo_lookup(refresh: bool):
    """物件一覧（Excel）からホテル名→（都道府県, 客室数）を引く関数を返す。"""
    xlsx = fetch("ichigo_hoteldata", CONFIG["reit_files"]["ichigo_xlsx"], refresh)
    master = {}
    ws = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)["(1)物件一覧"]
    for r in ws.iter_rows(min_row=6, values_only=True):
        if r[3] and r[4] and isinstance(r[10], (int, float)):
            pref = re.match(r"(北海道|東京都|京都府|大阪府|.{2,3}県)", str(r[4]).replace("\n", ""))
            master[_norm(str(r[3]))] = (pref.group(1) if pref else "", int(r[10]))

    def lookup(name):
        key = _norm(name)
        for k, v in master.items():
            if k.startswith(key) or key.startswith(k):
                return v
        return ("", None)

    return lookup


def parse_ichigo_pdf(pdf_path: Path, src: str, lookup) -> list[Hotel]:
    hotels = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                cur, vals = None, {}
                for r in table + [["__END__"]]:
                    if r[0] and r[0] not in ("", "ホテル／項目") and (len(r) > 1 and r[1] and "売上高" in str(r[1]) or r[0] == "__END__"):
                        if cur and len(vals) == 3:
                            pref, rooms = lookup(cur)
                            hotels.append(Hotel("いちごホテルリート", cur, pref, rooms,
                                                vals["occ"][0], vals["occ"][1], vals["adr"][0], vals["adr"][1],
                                                vals["revpar"][0], vals["revpar"][1], src))
                        name = re.sub(r"[（(]\s*[）)]", "", re.sub(r"\s+", "", re.sub(r"※\d+", "", r[0])))
                        cur, vals = (name if r[0] != "__END__" else None), {}
                    if cur and len(r) >= 4:
                        metric = str(r[1] or "")
                        key = "revpar" if "RevPAR" in metric else "occ" if "稼働率" in metric else "adr" if "ADR" in metric else None
                        if key:
                            vals[key] = (num(r[2]), num(r[3]))
    return hotels


def hoshino_hotels(refresh: bool) -> list[Hotel]:
    xlsx = fetch("hoshino_ops", CONFIG["reit_files"]["hoshino_xlsx"], refresh)
    y, m = CONFIG["reit_month"]
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)

    def read(sheet, year):
        ws = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        ycur, col = None, None
        for i, mm in enumerate(rows[4]):
            if rows[3][i]:
                ycur = rows[3][i]
            if ycur == f"{year}年" and mm == f"{m}月":
                col = i
        out, code, name = {}, None, None
        for r in rows[5:]:
            if r[3]:
                code, name = r[3], r[4]
            if code and r[5]:
                out.setdefault(code, {"name": name})[r[5]] = num(r[col])
        return out

    cur = read(CONFIG["hoshino_sheets"]["current"], y)
    prev = read(CONFIG["hoshino_sheets"]["prior"], y - 1)
    master = {}
    with (DATA / "hoshino_hotels_master.csv").open(encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            master[r["物件番号"]] = (r["都道府県"], int(r["客室数"]) if r["客室数"] else None)
    hotels = []
    for code, v in cur.items():
        p = prev.get(code, {})
        pref, rooms = master.get(code, ("", None))
        occ = v.get("客室稼働率")
        occ_py = p.get("客室稼働率")
        hotels.append(Hotel("星野リゾート・リート", str(v["name"]).strip(), pref, rooms,
                            round(occ * 100, 1) if occ else None, round(occ_py * 100, 1) if occ_py else None,
                            v.get("ADR（円）"), p.get("ADR（円）"), v.get("RevPAR（円）"), p.get("RevPAR（円）"),
                            CONFIG["reit_files"]["hoshino_xlsx"]))
    return hotels


def jhr_regions(refresh: bool) -> list[dict]:
    xlsx = fetch("jhr_review", CONFIG["reit_files"]["jhr_xlsx"], refresh)
    ws = openpyxl.load_workbook(xlsx, data_only=True)["変動賃料導入29ホテル（地域別）"]
    rows = list(ws.iter_rows(values_only=True))
    out, region, metric = [], None, None
    data = {}
    for r in rows:
        if r[1] and r[3] is None and all(c is None for c in r[4:16]) and "主要指標" not in str(r[1]):
            region = r[1]
        if r[1] in ("客室稼働率（％）", "ADR（円）", "RevPAR（円）"):
            metric = r[1]
        elif r[1] and "売上高" in str(r[1]):
            metric = None
        if region and metric and r[3] in ("2026年", "増減", "増減率"):
            data.setdefault(region, {}).setdefault(metric, {})[r[3]] = r[4:16]
    hotel_count = {"北海道": 2, "東京": 3, "関東（東京除く）": 5, "大阪": 4, "関西（大阪除く）": 4,
                   "中国": 2, "九州（沖縄除く）": 5, "沖縄": 4}
    for month in (7, 8):
        for reg, d in data.items():
            o, a, p = d["客室稼働率（％）"], d["ADR（円）"], d["RevPAR（円）"]
            i = month - 1
            out.append({
                "REIT": "ジャパン・ホテル・リート", "地域": reg, "ホテル数": hotel_count.get(reg), "対象月": f"2026-{month:02d}",
                "OCC_当月(%)": round(o["2026年"][i] * 100, 1), "OCC_前年差(pt)": o["増減"][i],
                "ADR_当月(円)": a["2026年"][i], "ADR_前年比(%)": round(a["増減率"][i] * 100, 1),
                "RevPAR_当月(円)": p["2026年"][i], "RevPAR_前年比(%)": round(p["増減率"][i] * 100, 1),
                "集計方法": "REIT公表の地域合算値（販売客室数・販売可能客室数ベース）",
                "出典URL": CONFIG["reit_files"]["jhr_xlsx"], "取得日": RETRIEVED,
            })
    return out


def invincible_regions(refresh: bool) -> list[dict]:
    pdf_path = fetch("invincible_monthly", CONFIG["reit_files"]["inv_pdf"], refresh)
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    out = []
    pat = re.compile(r"^(東京23区|首都圏（東京23区除く）|中部|関西|九州|北海道|その他|合計) ([\d.]+)% （([+△][\d.]+)pt） ([\d,]+) （([+△][\d.]+)%） ([\d,]+) （([+△][\d.]+)%）", re.M)
    for mm in pat.finditer(text):
        out.append({
            "REIT": "インヴィンシブル", "地域": mm.group(1), "ホテル数": 101 if mm.group(1) == "合計" else None,
            "対象月": "2026-08", "OCC_当月(%)": num(mm.group(2)), "OCC_前年差(pt)": num(mm.group(3)),
            "ADR_当月(円)": num(mm.group(4)), "ADR_前年比(%)": num(mm.group(5)),
            "RevPAR_当月(円)": num(mm.group(6)), "RevPAR_前年比(%)": num(mm.group(7)),
            "集計方法": "REIT公表の地域合算値（国内101物件、地域別ホテル数は非記載）",
            "出典URL": CONFIG["reit_files"]["inv_pdf"], "取得日": RETRIEVED,
        })
    return out


def aggregate_by_pref(hotels: list[Hotel], month: str) -> list[dict]:
    groups: dict[str, list[Hotel]] = {}
    for h in hotels:
        if h.ok() and h.pref and h.rooms:
            groups.setdefault(h.pref, []).append(h)
    rows = []
    for pref, hs in sorted(groups.items(), key=lambda kv: -sum(h.rooms for h in kv[1])):
        rooms = sum(h.rooms for h in hs)
        w = lambda attr: sum(getattr(h, attr) * h.rooms for h in hs) / rooms
        rev, rev_py = w("revpar"), w("revpar_py")
        occ, occ_py = w("occ"), w("occ_py")
        adr, adr_py = rev / (occ / 100), rev_py / (occ_py / 100)
        avg = lambda f: round(sum(f(h) for h in hs) / len(hs), 1)
        rows.append({
            "都道府県": pref, "対象月": month, "ホテル数": len(hs), "客室数計": rooms,
            "内訳": " / ".join(f"{h.reit[:2]}:{h.name}" for h in hs),
            "加重_OCC_前年差(pt)": round(occ - occ_py, 1),
            "加重_ADR_前年比(%)": round((adr / adr_py - 1) * 100, 1),
            "加重_RevPAR_前年比(%)": round((rev / rev_py - 1) * 100, 1),
            "単純平均_ADR_前年比(%)": avg(lambda h: (h.adr / h.adr_py - 1) * 100),
            "単純平均_RevPAR_前年比(%)": avg(lambda h: (h.revpar / h.revpar_py - 1) * 100),
            "加重_ADR_当月(円)": round(adr), "加重_RevPAR_当月(円)": round(rev),
            "集計方法": "客室数加重（RevPAR・OCCを客室数で加重、ADR=加重RevPAR÷加重OCC）と単純平均の併記",
            "出典": "いちご: " + CONFIG["reit_files"]["ichigo_pdf"] + " / 星野: " + CONFIG["reit_files"]["hoshino_xlsx"],
            "取得日": RETRIEVED,
        })
    return rows


# ---------------------------------------------------------------- 地図

def setup_font() -> None:
    for path in ["/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"]:
        if Path(path).exists():
            font_manager.fontManager.addfont(path)
    plt.rcParams["font.family"] = ["Noto Sans CJK JP"]


def load_prefs(refresh: bool):
    gj = json.loads(fetch("japan", CONFIG["reit_files"]["geojson"], refresh).read_text(encoding="utf-8"))
    return {ft["properties"]["nam_ja"]: shape(ft["geometry"]) for ft in gj["features"]}


def _draw_geom(ax, geom, color, lw=0.3):
    polys = [geom] if geom.geom_type == "Polygon" else list(geom.geoms)
    for p in polys:
        x, y = p.exterior.xy
        ax.fill(x, y, facecolor=color, edgecolor="#555555", linewidth=lw)


def draw_japan(fig, rect, prefs, color_of) -> None:
    """本土を rect に、沖縄県（先島諸島を含み、大東諸島は除く）を左上の枠内に描く。"""
    ax = fig.add_axes(rect)
    ax.set_xlim(128.6, 146.2)
    ax.set_ylim(30.6, 45.7)
    ax.set_aspect(1.2)
    ax.axis("off")
    for name, geom in prefs.items():
        if name != "沖縄県":
            _draw_geom(ax, geom, color_of(name))
    x0, y0, w, h = rect
    inset = fig.add_axes([x0 + w * 0.02, y0 + h * 0.52, w * 0.30, h * 0.2])
    inset.set_xlim(122.8, 128.6)
    inset.set_ylim(23.9, 27.2)
    inset.set_aspect(1.1)
    inset.set_xticks([])
    inset.set_yticks([])
    for s in inset.spines.values():
        s.set_color("#999999")
        s.set_linewidth(0.6)
    _draw_geom(inset, prefs["沖縄県"], color_of("沖縄県"), lw=0.2)
    inset.text(122.95, 27.05, "沖縄県", fontsize=6, color="#555555", va="top")


def draw_table(fig, rect, header, rows, fontsize=7.0, col_widths=None, value_cols=()) -> None:
    """値の列は前年比の符号で文字色を変える（緑=上昇 / 赤=低下）。"""
    ax = fig.add_axes(rect)
    ax.axis("off")
    tbl = ax.table(cellText=rows, colLabels=header, loc="upper left", cellLoc="right", colLoc="center",
                   colWidths=col_widths, bbox=[0, 0, 1, 1])
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(fontsize)
    for (r, c), cell in tbl.get_celld().items():
        cell.set_linewidth(0.3)
        cell.set_edgecolor("#bbbbbb")
        if r == 0:
            cell.set_facecolor("#f0f0f0")
        elif c == 0:
            cell.get_text().set_horizontalalignment("left")
            cell.PAD = 0.04
        if r > 0 and c in value_cols:
            t = cell.get_text().get_text()
            cell.get_text().set_color("#1a7f37" if t.startswith("+") and t.strip("+0.%pt") else
                                      "#c62828" if t.startswith("-") else "#333333")


def _fmt(v, unit="", digits=1):
    return "(未)" if v is None else f"{v:+.{digits}f}{unit}"


def map_occ(rows: list[dict], nat_aug: dict, out: Path, prefs) -> None:
    vals = {r["都道府県"]: r["OCC_全体_前年差(pt)"] for r in rows if r["都道府県"] != "全国"}
    nat = next(r for r in rows if r["都道府県"] == "全国")
    lim = max(abs(v) for v in vals.values() if v is not None)
    lim = max(5.0, round(lim + 0.5))
    norm = TwoSlopeNorm(vmin=-lim, vcenter=0, vmax=lim)
    cmap = plt.get_cmap("RdYlGn")

    fig = plt.figure(figsize=(8, 7.6), dpi=100)
    draw_japan(fig, [0.0, 0.14, 0.66, 0.78], prefs,
               lambda n: "#dddddd" if vals.get(n) is None else cmap(norm(vals[n])))

    sm = plt.cm.ScalarMappable(norm=norm, cmap=cmap)
    cax = fig.add_axes([0.08, 0.105, 0.5, 0.02])
    cb = fig.colorbar(sm, cax=cax, orientation="horizontal")
    cb.set_label("客室稼働率（全体）の前年同月差（ポイント）  緑=上昇 / 赤=低下", fontsize=8)
    cb.ax.tick_params(labelsize=7)

    y, m = CONFIG["mlit_month"]
    fig.text(0.02, 0.965, f"都道府県別 客室稼働率の前年同月差（{y}年{m}月・第2次速報 vs {y - 1}年{m}月・確定値）",
             fontsize=11.5, weight="bold")
    fig.text(0.02, 0.937, f"全国 {nat['OCC_全体_当月(%)']:.1f}%（{_fmt(nat['OCC_全体_前年差(pt)'], 'pt')}）"
             f"  外国人比率 {nat['外国人比率_当月(%)']:.1f}%（{_fmt(nat['外国人比率_前年差(pt)'], 'pt')}）"
             f"  ※8月1次速報の全国OCCは{nat_aug['OCC全体']}%（都道府県別は未公表）", fontsize=8)

    ranked = sorted(((v, k) for k, v in vals.items() if v is not None), reverse=True)
    fig.text(0.665, 0.905, "OCC前年差 上昇・低下 上位6", fontsize=8, weight="bold")
    draw_table(fig, [0.665, 0.70, 0.32, 0.195], ["上昇", "pt", "低下", "pt"],
               [[k1, f"{v1:+.1f}", k2, f"{v2:+.1f}"] for (v1, k1), (v2, k2) in zip(ranked[:6], ranked[::-1][:6])],
               fontsize=7.2, col_widths=[0.3, 0.2, 0.3, 0.2], value_cols=(1, 3))
    by = {r["都道府県"]: r for r in rows}
    focus = ["全国", "東京都", "大阪府", "京都府", "北海道", "福岡県", "沖縄県", "愛知県", "神奈川県"]
    fig.text(0.665, 0.665, "主要都市圏（全体／シティ／ビジネス、外国人比率）", fontsize=8, weight="bold")
    draw_table(fig, [0.665, 0.30, 0.32, 0.355], ["", "OCC差\n全体", "シティ", "ビジ\nネス", "外国人\n比率", "比率\n差"],
               [[k, _fmt(by[k]["OCC_全体_前年差(pt)"]), _fmt(by[k]["OCC_シティ_前年差(pt)"]),
                 _fmt(by[k]["OCC_ビジネス_前年差(pt)"]), f"{by[k]['外国人比率_当月(%)']:.1f}%",
                 _fmt(by[k]["外国人比率_前年差(pt)"])] for k in focus],
               fontsize=6.8, col_widths=[0.22, 0.16, 0.15, 0.15, 0.17, 0.15], value_cols=(1, 2, 3, 5))
    fig.text(0.665, 0.285, "単位：OCC差・比率差はpt（前年同月差）", fontsize=6.3, color="#555555")

    fig.text(0.02, 0.012,
             "出典：観光庁「宿泊旅行統計調査」推移表 https://www.mlit.go.jp/kankocho/content/002026697.xlsx（取得 " + RETRIEVED + "）\n"
             "注：2026年1月分から層化基準を従業者数→客室数に変更。前年差には見直しの影響が含まれ得る（観光庁注記）。"
             "境界：dataofjapan/land",
             fontsize=6.3, color="#333333")
    fig.savefig(out, dpi=100)
    plt.close(fig)


def map_reit(pref_rows: list[dict], region_rows: list[dict], out: Path, prefs) -> None:
    vals = {r["都道府県"]: r for r in pref_rows}
    lim = max(abs(r["加重_RevPAR_前年比(%)"]) for r in pref_rows)
    lim = min(max(10.0, round(lim + 1)), 30.0)
    norm = TwoSlopeNorm(vmin=-lim, vcenter=0, vmax=lim)
    cmap = plt.get_cmap("RdYlGn")

    fig = plt.figure(figsize=(8, 9.6), dpi=100)
    draw_japan(fig, [0.0, 0.40, 0.6, 0.53], prefs,
               lambda n: "#eeeeee" if n not in vals else cmap(norm(vals[n]["加重_RevPAR_前年比(%)"])))

    sm = plt.cm.ScalarMappable(norm=norm, cmap=cmap)
    cax = fig.add_axes([0.06, 0.405, 0.48, 0.014])
    cb = fig.colorbar(sm, cax=cax, orientation="horizontal", extend="both")
    cb.ax.tick_params(labelsize=7)
    fig.text(0.06, 0.375, "色：RevPAR前年同月比（%、客室数加重）　灰色：データなし", fontsize=7.5)

    y, m = CONFIG["reit_month"]
    fig.text(0.02, 0.972, f"J-REIT保有ホテルの都道府県別 ADR・RevPAR 前年同月比（{y}年{m}月）", fontsize=11.5, weight="bold")
    fig.text(0.02, 0.952, "地図・右表：いちごホテルリート＋星野リゾート・リートのホテル別開示を所在都道府県で集計（客室数加重）",
             fontsize=7.8)

    fig.text(0.615, 0.935, "都道府県別（客室数の多い順）", fontsize=7.5, weight="bold")
    draw_table(fig, [0.615, 0.37, 0.37, 0.56], ["都道府県", "軒数", "客室数", "ADR", "RevPAR"],
               [[r["都道府県"], r["ホテル数"], f"{r['客室数計']:,}", f"{r['加重_ADR_前年比(%)']:+.1f}%",
                 f"{r['加重_RevPAR_前年比(%)']:+.1f}%"] for r in pref_rows],
               fontsize=6.4, col_widths=[0.26, 0.14, 0.2, 0.2, 0.2], value_cols=(3, 4))

    reg = [r for r in region_rows if r["対象月"] == f"{y}-{m:02d}"]
    fig.text(0.02, 0.352, "参考：REIT公表の地域別合算値（前年同月比。都道府県への割り振りはしていない）", fontsize=7.5, weight="bold")
    half = (len(reg) + 1) // 2
    for i, part in enumerate((reg[:half], reg[half:])):
        draw_table(fig, [0.02 + i * 0.49, 0.345 - 0.0215 * (len(part) + 1), 0.47, 0.0215 * (len(part) + 1)],
                   ["REIT・地域", "OCC差", "ADR", "RevPAR"],
                   [[f"{'JHR' if r['REIT'].startswith('ジャパン') else 'INV'} {r['地域']}"
                     + (f"（{r['ホテル数']}軒）" if r["ホテル数"] else ""),
                     f"{r['OCC_前年差(pt)']:+.1f}pt", f"{r['ADR_前年比(%)']:+.1f}%", f"{r['RevPAR_前年比(%)']:+.1f}%"]
                    for r in part],
                   fontsize=6.4, col_widths=[0.5, 0.16, 0.17, 0.17], value_cols=(1, 2, 3))
    fig.text(0.02, 0.135, "JHR＝ジャパン・ホテル・リート（変動賃料等導入29ホテル）、INV＝インヴィンシブル（国内101物件、地域別の軒数は非開示）",
             fontsize=6.3, color="#555555")

    fig.text(0.02, 0.068,
             "出典：いちごホテルリート 2026年8月度ホテル運営状況 https://www.release.tdnet.info/inbs/140120260925540031.pdf / 星野リゾート・リート 運営実績データ\n"
             "https://ssl4.eir-parts.net/doc/3287/ir_material_for_fiscal_ym5/195377/00.xlsx / JHR ホテル運営実績Excel https://www.jhrth.co.jp/ja/portfolio/review.html /\n"
             "インヴィンシブル 2026年8月運用実績 https://www.release.tdnet.info/inbs/140120260925540060.pdf（いずれも取得 " + RETRIEVED + "）\n"
             "注：ADR・RevPARの定義（サービス料の扱い等）はREITごとに異なる。休館・新規取得で前年比較できないホテルは除外。境界：dataofjapan/land",
             fontsize=6.0, color="#333333")
    fig.savefig(out, dpi=100)
    plt.close(fig)
    img = Image.open(out)
    img.crop((0, 0, img.width, img.height - 55)).save(out)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--refresh", action="store_true")
    args = ap.parse_args()

    mlit_rows, nat_aug = build_mlit(args.refresh)
    y, m = CONFIG["reit_month"]
    month = f"{y}-{m:02d}"
    hotels = ichigo_hotels(args.refresh) + hoshino_hotels(args.refresh)
    write_csv(f"jreit_hotels_{month}.csv", [h.row(month) for h in hotels])
    pref_rows = aggregate_by_pref(hotels, month)
    write_csv(f"jreit_area_reference_{month}.csv", pref_rows)
    regions = jhr_regions(args.refresh) + invincible_regions(args.refresh)
    write_csv("jreit_regions_reported.csv", regions)

    setup_font()
    prefs = load_prefs(args.refresh)
    my, mm = CONFIG["mlit_month"]
    map_occ(mlit_rows, nat_aug, HERE / f"map_occ_yoy_{my}-{mm:02d}.png", prefs)
    map_reit(pref_rows, regions, HERE / f"map_jreit_adr_revpar_yoy_{month}.png", prefs)
    print("done")


if __name__ == "__main__":
    main()
