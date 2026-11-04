# -*- coding: utf-8 -*-
"""
班次空檔分析：讀運務通知 .docx（內嵌班次表 Excel）或 .xlsx，
找出每台車頭（曳引車／全拖車）調整後班表中，兩班之間停在同一站的空檔，
並依班表反推的站間車程，列出「來回＋作業緩衝」放得進空檔的鄰近站所。

用法：
  python analyze.py <檔案1.docx|xlsx> [<檔案2> ...] -o 輸出.xlsx [--min 60] [--buffer 30]
"""
import sys, os, re, io, json, argparse, zipfile, statistics, collections, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from parse_shift import parse_sheet, norm

MIN_LEG = 5      # 站間車程下限（分）：彰化↔秀水 班表記 0 分
MAX_LEG = 600    # 超過 10 小時的「車程」視為異常不採計


def extract_xlsx(path):
    """docx → 內嵌 xlsx bytes；xlsx → 原檔 bytes"""
    if path.lower().endswith(".xlsx"):
        return open(path, "rb").read()
    z = zipfile.ZipFile(path)
    names = z.namelist()
    for n in names:
        if n.startswith("word/embeddings/") and n.lower().endswith(".xlsx"):
            return z.read(n)
    for n in names:
        if n.startswith("word/embeddings/") and n.lower().endswith(".bin"):
            b = z.read(n)
            try:
                import olefile
                ole = olefile.OleFileIO(io.BytesIO(b))
                for s in ole.listdir():
                    if s[-1].lower() == "package":
                        return ole.openstream(s).read()
            except Exception:
                pass
            i = b.find(b"PK\x03\x04")
            if i >= 0:
                return b[i:]
    raise RuntimeError(f"{path}: 找不到內嵌的班次表 Excel")


def label_of(path):
    m = re.search(r"(1[0-2]|0?[1-9])(0[1-9]|[12]\d|3[01])(?!\d)", os.path.basename(path))
    return m.group(0) if m else os.path.splitext(os.path.basename(path))[0]


def hhmm(m):
    if m is None:
        return ""
    m = int(m) % 1440
    return f"{m // 60:02d}{m % 60:02d}"


def opt(m):
    """營運時間：以 12:00 為一天起點，讓跨午夜排序一致（12:00→720、03:10→1630）"""
    return (int(m) % 1440 - 720) % 1440 + 720


def load_day(path):
    wb = openpyxl.load_workbook(io.BytesIO(extract_xlsx(path)), data_only=True)
    day = {"label": label_of(path), "file": os.path.basename(path), "tractors": []}
    for sheet in ["曳引車", "全拖車"]:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        rcol = None
        for c in range(1, ws.max_column + 1):
            if norm(ws.cell(3, c).value) == "備註":
                rcol = get_column_letter(c)
        trs = parse_sheet(ws, rcol)
        unit = None
        for t in trs:
            if t["unit"]:
                unit = t["unit"]
            t["unit"] = unit
            t["kind"] = sheet
            build_timeline(t)
            day["tractors"].append(t)
    return day


def build_timeline(t):
    """調整後班表（排除綠色）→ 依序的 (站, 絕對分鐘) 點、班次段、空檔"""
    pts = [(b, p) for b in t["blocks"] for p in b["points"] if p["color"] != "green"]
    prev = None
    dayoff = 0
    seq = []
    for b, p in pts:
        q = dict(p)
        q["block_row"] = b["row"]
        if p["time"] is not None and not isinstance(p["time"], list):
            tm = p["time"]
            if prev is not None and tm < prev - 60:
                dayoff += 1
            q["abs"] = tm + dayoff * 1440
            prev = tm
        else:
            q["abs"] = None
        q["shift_code"] = p["shift"]
        q["trailer"] = None
        q["note"] = None
        q["shift_color"] = None
        seq.append(q)
    # 每個點對應到所屬班次（代號列合併儲存格的起始欄）
    for b in t["blocks"]:
        cols = sorted(b["shifts"], key=lambda s: column_index_from_string(s["col"]))
        for q in [x for x in seq if x["block_row"] == b["row"]]:
            ci = column_index_from_string(q["col"])
            meta = None
            for s in cols:
                if column_index_from_string(s["col"]) <= ci:
                    meta = s
            if meta:
                q["trailer"] = meta["trailer"]
                q["note"] = meta["note"]
                q["shift_color"] = meta["color"]
                if q["shift_code"] is None:
                    q["shift_code"] = meta["code"]
    legs, gaps = [], []
    for i, (a, b) in enumerate(zip(seq, seq[1:])):
        if a["abs"] is None or b["abs"] is None:
            continue
        dur = b["abs"] - a["abs"]
        if a["station"] == b["station"]:
            gaps.append({"station": a["station"], "start": a["abs"], "end": b["abs"], "dur": dur,
                         "prev": a, "next": b, "prev_leg": prev_leg(seq, i), "next_leg": next_leg(seq, i + 1)})
        else:
            legs.append({"from": a["station"], "to": b["station"], "start": a["abs"], "end": b["abs"], "dur": dur,
                         "code": a["shift_code"], "trailer": a["trailer"], "note": a["note"],
                         "color": "yellow" if "yellow" in (a["color"], b["color"], a["shift_color"]) else "white"})
    t["seq"] = seq
    t["legs"] = legs
    t["gaps"] = gaps
    t["active"] = t["color"] != "green" and len(seq) > 0
    t["double"] = bool(t.get("remark") and "雙駕駛" in str(t["remark"]))


def prev_leg(seq, i):
    """空檔起點所屬的班次（到站的那一班）；無班次代號＝日初待命"""
    a = seq[i]
    if a["shift_code"] is None and a["trailer"] is None:
        return None
    frm = None
    j = i
    while j > 0:
        if seq[j - 1]["station"] != seq[j]["station"]:
            frm = seq[j - 1]["station"]
            break
        j -= 1
    return {"from": frm, "code": a["shift_code"], "trailer": a["trailer"], "note": a["note"]}


def next_leg(seq, i):
    """空檔終點所屬的班次（離站的那一班）"""
    b = seq[i]
    if b["shift_code"] is None and b["trailer"] is None:
        return None
    to = None
    j = i
    while j < len(seq) - 1:
        if seq[j + 1]["station"] != seq[j]["station"]:
            to = seq[j + 1]["station"]
            break
        j += 1
    return {"to": to, "code": b["shift_code"], "trailer": b["trailer"], "note": b["note"]}


def travel_matrix(days):
    samples = collections.defaultdict(list)
    for d in days:
        for t in d["tractors"]:
            if not t["active"]:
                continue
            for l in t["legs"]:
                if 0 <= l["dur"] <= MAX_LEG:
                    samples[(l["from"], l["to"])].append(l["dur"])
    mat = {}
    for k, v in samples.items():
        mat[k] = {"n": len(v), "med": max(MIN_LEG, statistics.median(v)), "min": min(v), "max": max(v)}
    return mat


def travel(mat, a, b):
    if (a, b) in mat:
        return mat[(a, b)]["med"], mat[(a, b)]["n"]
    if (b, a) in mat:
        return mat[(b, a)]["med"], mat[(b, a)]["n"]
    return None, 0


def candidates(mat, station, gap_min, buffer):
    """來回車程＋作業緩衝 放得進空檔的站所"""
    others = set()
    for (a, b) in mat:
        if a == station:
            others.add(b)
        if b == station:
            others.add(a)
    out = []
    for T in others:
        t1, n1 = travel(mat, station, T)
        t2, n2 = travel(mat, T, station)
        if t1 is None or t2 is None:
            continue
        rt = t1 + t2 + buffer
        if rt <= gap_min:
            out.append({"station": T, "go": t1, "back": t2, "round": rt, "slack": gap_min - rt, "n": min(n1, n2)})
    out.sort(key=lambda x: (x["round"], x["station"]))
    return out


def analyze(days, min_gap, buffer):
    mat = travel_matrix(days)
    rows = []
    for d in days:
        for t in d["tractors"]:
            if not t["active"]:
                continue
            for g in t["gaps"]:
                if g["dur"] < min_gap:
                    continue
                cands = candidates(mat, g["station"], g["dur"], buffer)
                flags = []
                if g["prev_leg"] is None:
                    flags.append("日初待命(前無班次)")
                if g["next_leg"] is None:
                    flags.append("日末(後無班次)")
                if t["double"]:
                    flags.append("雙駕駛")
                for x in (g["prev_leg"], g["next_leg"]):
                    if x and x.get("note"):
                        flags.append(f"備註:{x['note']}")
                if "yellow" in (g["prev"]["color"], g["next"]["color"]):
                    flags.append("前後班次有調整")
                rows.append({"day": d["label"], "kind": t["kind"], "unit": t["unit"], "code": t["code"], "name": t["name"],
                             "remark": t["remark"], "station": g["station"], "start": g["start"], "end": g["end"],
                             "dur": g["dur"], "prev": g["prev_leg"], "next": g["next_leg"], "cands": cands, "flags": flags})
    return mat, rows


# ---------- 加班需求清單 → 配車 ----------
SAME_STATION_RT = 60   # 本站加班（起迄同站，去客戶取貨）車程未知時假設來回 60 分


def parse_hhmm(s):
    s = str(s or "").strip().replace(":", "").replace("：", "")
    if not re.fullmatch(r"\d{3,4}", s):
        return None
    s = s.zfill(4)
    h, m = int(s[:2]), int(s[2:])
    if h > 24 or m > 59:
        return None
    return (h % 24) * 60 + m


def read_needs(path):
    """CSV（UTF-8/BOM/cp950）或 xlsx 第一張表；欄位：起站、迄站 必填，日期/時間中位/時間窗起/時間窗迄/機率/等級/預估櫃數/常見班次 選填"""
    import csv
    rows = []
    if path.lower().endswith(".xlsx"):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        hdr = [str(h or "").strip() for h in next(it)]
        for r in it:
            rows.append({hdr[i]: (r[i] if i < len(r) else None) for i in range(len(hdr))})
    else:
        raw = open(path, "rb").read()
        for enc in ("utf-8-sig", "utf-8", "cp950"):
            try:
                txt = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        rows = list(csv.DictReader(io.StringIO(txt)))
    needs = []
    for r in rows:
        g = lambda *ks: next((str(r[k]).strip() for k in ks if k in r and r[k] not in (None, "")), "")
        a, b = g("起站", "from"), g("迄站", "to")
        if not a or not b:
            continue
        tm = parse_hhmm(g("時間中位", "時間", "time"))
        t1 = parse_hhmm(g("時間窗起", "t1"))
        t2 = parse_hhmm(g("時間窗迄", "t2"))
        if t1 is None and tm is not None:
            t1 = tm - 30
        if t2 is None and tm is not None:
            t2 = tm + 30
        needs.append({"day": re.sub(r"\D", "", g("日期", "day"))[-4:] or None, "from": a, "to": b,
                      "tmed": opt(tm) if tm is not None else None, "t1": opt(t1) if t1 is not None else 720, "t2": opt(t2) if t2 is not None else 2159,
                      "prob": g("機率", "prob"), "tier": g("等級", "tier"), "boxes": g("預估櫃數", "boxes"), "names": g("常見班次", "names"),
                      "note": g("備註", "note")})
    return needs


def match_needs(days, mat, needs, min_gap, buffer):
    """每條需求：找起站（A）或迄站（B）在站空等、時間窗內放得下來回的車頭"""
    out = []
    for d in days:
        dn = [x for x in needs if not x["day"] or x["day"] == d["label"][-4:]]
        for nd in dn:
            S, T = nd["from"], nd["to"]
            if S == T:
                tST = tTS = SAME_STATION_RT / 2
                known = False
            else:
                tST, _ = travel(mat, S, T)
                tTS, _ = travel(mat, T, S)
                known = tST is not None and tTS is not None
                if tST is None:
                    tST = tTS
                if tTS is None:
                    tTS = tST
            cands = []
            if tST is not None:
                for t in d["tractors"]:
                    if not t["active"]:
                        continue
                    for g in t["gaps"]:
                        if g["dur"] < min_gap or g["station"] not in (S, T):
                            continue
                        gs = opt(g["start"])
                        ge = gs + g["dur"]
                        if g["station"] == S:
                            lo, hi, pat = max(gs, nd["t1"]), min(nd["t2"], ge - (tST + tTS + buffer)), "A 起站在站"
                        else:
                            lo, hi, pat = max(gs + tTS, nd["t1"]), min(nd["t2"], ge - (tST + buffer)), "B 迄站在站"
                        if S == T and g["station"] == S:
                            pat = "本站加班"
                        if lo > hi:
                            continue
                        cands.append({"pattern": pat, "tractor": t, "gap": g, "lo": lo, "hi": hi, "slack": ge - lo - (tST + tTS + buffer),
                                      "rt": tST + tTS + buffer})
            cands.sort(key=lambda c: (0 if c["pattern"].startswith(("A", "本")) else 1, -c["slack"]))
            out.append({"day": d["label"], "need": nd, "known": known, "tST": tST, "tTS": tTS, "cands": cands})
    return out


def write_needs_sheets(wb, matched, days, buffer):
    for d in days:
        ms = [m for m in matched if m["day"] == d["label"]]
        if not ms:
            continue
        ws = wb.create_sheet(f"加班配車_{d['label']}")
        data = []
        for m in ms:
            nd = m["need"]
            win = f"{hhmm(nd['t1'])}–{hhmm(nd['t2'])}" if nd["tmed"] is not None else "全日"
            base = [nd["from"], nd["to"], nd["tier"], nd["prob"], nd["boxes"], win, nd["names"],
                    (f"{int(m['tST'])}/{int(m['tTS'])}" if m["tST"] is not None else "無車程資料") + ("" if m["known"] else "（估）")]
            if not m["cands"]:
                data.append(base + ["—", None, None, None, None, None, None, None, None, "無合適車頭" if m["tST"] is not None else "無車程資料"])
            for c in m["cands"][:12]:
                t, g = c["tractor"], c["gap"]
                flags = []
                if t["double"]:
                    flags.append("雙駕駛")
                if g["prev_leg"] is None:
                    flags.append("日初待命")
                for x in (g["prev_leg"], g["next_leg"]):
                    if x and x.get("note"):
                        flags.append(f"備註:{x['note']}")
                data.append(base + [c["pattern"], t["kind"], t["unit"], t["code"], t["name"], g["station"], f"{hhmm(g['start'])}–{hhmm(g['end'])}（{g['dur']}分）",
                                    f"{hhmm(c['lo'])}–{hhmm(c['hi'])}", int(c["slack"]), "；".join(flags)])
        write_table(ws, ["起站", "迄站", "等級", "機率", "預估櫃數", "發車時間窗", "常見班次", "去/回車程(分)",
                         "跑法", "車種", "單位", "車頭代號", "車頭名稱", "空檔站", "空檔時段", "可出發區間", "餘裕(分)", "標記"],
                    data, [8, 8, 9, 6, 8, 11, 26, 12, 11, 8, 7, 9, 11, 8, 18, 12, 8, 22])
        for row in ws.iter_rows(min_row=2):
            tier = row[2].value
            if tier == "幾乎每週":
                row[2].fill = GRN
            elif tier == "很可能":
                row[2].fill = YEL
            if row[8].value == "—":
                row[8].fill = RED
        # 需求總表
        ws2 = wb.create_sheet(f"加班需求_{d['label']}")
        data = [[m["need"]["from"], m["need"]["to"], m["need"]["tier"], m["need"]["prob"], m["need"]["boxes"],
                 (f"{hhmm(m['need']['t1'])}–{hhmm(m['need']['t2'])}" if m["need"]["tmed"] is not None else "全日"), m["need"]["names"],
                 len(m["cands"]), sum(1 for c in m["cands"] if c["pattern"][0] in "A本"), sum(1 for c in m["cands"] if c["pattern"][0] == "B"),
                 "、".join(list(dict.fromkeys(f"{c['tractor']['code']}({c['pattern'][0]})" for c in m["cands"]))[:6])] for m in ms]
        write_table(ws2, ["起站", "迄站", "等級", "機率", "預估櫃數", "發車時間窗", "常見班次", "候選車頭數", "A 起站", "B 迄站", "候選車頭（前 6）"],
                    data, [8, 8, 9, 6, 8, 11, 26, 9, 7, 7, 50])


# ---------- Excel ----------
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", name="微軟正黑體")
BODY_FONT = Font(name="微軟正黑體", size=10)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
YEL = PatternFill("solid", fgColor="FFF2CC")
GRN = PatternFill("solid", fgColor="E2EFDA")
RED = PatternFill("solid", fgColor="FCE4D6")
UNIT_SEQ = ["北區", "新竹", "中區", "南區", "羅東", "棧板部"]


def unit_rank(u):
    return UNIT_SEQ.index(u) if u in UNIT_SEQ else 99


def kind_rank(k):
    return {"曳引車": 0, "全拖車": 1}.get(k, 2)


def write_table(ws, headers, rows, widths=None, freeze="A2"):
    ws.append(headers)
    for c in ws[1]:
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in row:
            c.font = BODY_FONT
            c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 30


def cand_text(cands):
    return "、".join(f"{c['station']}({int(c['round'])}分)" for c in cands)


def write_excel(days, mat, rows, out, min_gap, buffer, matched=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("說明")
    lines = [["班次空檔分析（可排加班的空檔與鄰近站所）"], [],
             ["產生時間", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")],
             ["來源檔案", "；".join(d["file"] for d in days)],
             ["空檔門檻", f"{min_gap} 分鐘以上（同一站停留，兩班之間）"],
             ["作業緩衝", f"{buffer} 分鐘（掛／卸車廂等作業，加在來回車程之外）"],
             ["站間車程", "由調整後班表（白＋黃）實際班次反推，取中位數，下限 5 分；表「站間車程」有樣本數可參考"],
             ["可跑站所", "去程＋回程＋作業緩衝 ≤ 空檔分鐘，依來回時間由近到遠排列；括號內為來回合計分鐘"],
             ["底色規則", "白＝原班次、黃＝新增／調整班次、綠＝刪減（已排除不計）"],
             ["注意", "全拖車部分班次備註「含休息時間」，空檔若為司機休息請自行判斷；雙駕駛車頭已在「標記」欄註明"],
             [], ["工作表", "內容"],
             ["空檔_MMDD", "每台車頭的空檔清單（可篩選單位／站所／分鐘）"],
             ["站所別_MMDD", "依站所、時間排序，看某站某時段有哪些車頭在站"],
             ["時間軸_MMDD", "每台車頭完整行程（班次段＋空檔）"],
             ["站間車程", "反推的站間車程（樣本數／中位／最小／最長）"],
             ["加班需求_MMDD", "（有給需求清單時）每條加班路線的候選車頭數"],
             ["加班配車_MMDD", "（有給需求清單時）每條加班路線的候選車頭：A＝起站在站空等去送再空回、B＝迄站在站空等先空去起站載回來；可出發區間＝時間窗與空檔都放得下的發車時間"]]
    for l in lines:
        ws.append(l)
    ws["A1"].font = Font(bold=True, size=14, name="微軟正黑體")
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = BODY_FONT
            c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 90
    for d in days:
        lab = d["label"]
        drows = [r for r in rows if r["day"] == lab]
        ws = wb.create_sheet(f"空檔_{lab}")
        data = []
        for r in sorted(drows, key=lambda x: (kind_rank(x["kind"]), unit_rank(x["unit"]), x["code"], opt(x["start"]))):
            data.append([r["kind"], r["unit"], r["code"], r["name"], r["remark"], r["station"], hhmm(r["start"]), hhmm(r["end"]), r["dur"],
                         (r["prev"] or {}).get("code"), (r["prev"] or {}).get("trailer"), (r["prev"] or {}).get("from"),
                         (r["next"] or {}).get("code"), (r["next"] or {}).get("trailer"), (r["next"] or {}).get("to"),
                         len(r["cands"]), cand_text(r["cands"]), "；".join(r["flags"])])
        write_table(ws, ["車種", "單位", "車頭代號", "車頭名稱", "車頭備註", "空檔站所", "開始", "結束", "空檔分鐘",
                         "前班次代號", "前班次車廂", "前班次來自", "後班次代號", "後班次車廂", "後班次前往", "可跑站數", "可跑站所(來回分鐘)", "標記"],
                    data, [8, 7, 9, 11, 14, 9, 7, 7, 9, 10, 14, 9, 10, 14, 9, 8, 60, 22])
        for row in ws.iter_rows(min_row=2):
            dur = row[8].value
            if dur >= 180:
                row[8].fill = GRN
            elif dur >= 120:
                row[8].fill = YEL
            if row[17].value and "休息" in str(row[17].value):
                row[17].fill = RED
        ws = wb.create_sheet(f"站所別_{lab}")
        data = []
        for r in sorted(drows, key=lambda x: (x["station"], opt(x["start"]), x["code"])):
            data.append([r["station"], hhmm(r["start"]), hhmm(r["end"]), r["dur"], r["kind"], r["unit"], r["code"], r["name"],
                         (r["prev"] or {}).get("from"), (r["next"] or {}).get("to"), cand_text(r["cands"]), "；".join(r["flags"])])
        write_table(ws, ["站所", "開始", "結束", "空檔分鐘", "車種", "單位", "車頭代號", "車頭名稱", "來自", "下一站", "可跑站所(來回分鐘)", "標記"],
                    data, [9, 7, 7, 9, 8, 7, 9, 11, 9, 9, 60, 22])
        ws = wb.create_sheet(f"時間軸_{lab}")
        data = []
        for t in sorted([x for x in d["tractors"] if x["active"]], key=lambda x: (kind_rank(x["kind"]), unit_rank(x["unit"]), x["row"])):
            events = [("班次", l) for l in t["legs"]] + [("空檔", g) for g in t["gaps"]]
            events.sort(key=lambda e: e[1]["start"])
            for n, (kind, e) in enumerate(events, 1):
                if kind == "班次":
                    data.append([t["kind"], t["unit"], t["code"], t["name"], n, "班次", e["code"], e["trailer"], e["from"], hhmm(e["start"]),
                                 e["to"], hhmm(e["end"]), e["dur"], "新增/調整" if e["color"] == "yellow" else "原班次", e["note"]])
                else:
                    data.append([t["kind"], t["unit"], t["code"], t["name"], n, "空檔" if e["dur"] >= min_gap else "停留", None, None,
                                 e["station"], hhmm(e["start"]), e["station"], hhmm(e["end"]), e["dur"], None, None])
        write_table(ws, ["車種", "單位", "車頭代號", "車頭名稱", "序", "類型", "班次代號", "附掛車廂", "起站", "發車", "迄站", "到達", "分鐘", "狀態", "班次備註"],
                    data, [8, 7, 9, 11, 5, 6, 10, 16, 8, 7, 8, 7, 7, 10, 14])
        for row in ws.iter_rows(min_row=2):
            if row[5].value == "空檔":
                row[5].fill = GRN
                row[12].fill = GRN
            elif row[13].value == "新增/調整":
                row[13].fill = YEL
    if matched:
        write_needs_sheets(wb, matched, days, buffer)
    ws = wb.create_sheet("站間車程")
    data = [[a, b, v["n"], v["med"], v["min"], v["max"]] for (a, b), v in sorted(mat.items(), key=lambda kv: (kv[0][0], kv[1]["med"], kv[0][1]))]
    write_table(ws, ["起站", "迄站", "樣本數", "車程中位(分)", "最短", "最長"], data, [9, 9, 8, 12, 8, 8])
    wb.save(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--min", type=int, default=60)
    ap.add_argument("--buffer", type=int, default=30)
    ap.add_argument("--json", help="另存分析結果 JSON")
    ap.add_argument("--needs", help="加班需求清單 CSV/xlsx（ot_forecast.py 產出或自行編）")
    a = ap.parse_args()
    days = [load_day(f) for f in a.files]
    mat, rows = analyze(days, a.min, a.buffer)
    matched = None
    if a.needs:
        needs = read_needs(a.needs)
        matched = match_needs(days, mat, needs, a.min, a.buffer)
        for d in days:
            ms = [m for m in matched if m["day"] == d["label"]]
            byt = collections.Counter((m["need"]["tier"], bool(m["cands"])) for m in ms)
            print(f"{d['label']}: 需求 {len(ms)} 條、有候選車頭 {sum(1 for m in ms if m['cands'])} 條；"
                  + "；".join(f"{t} {byt[(t, True)]}/{byt[(t, True)] + byt[(t, False)]}" for t in ("幾乎每週", "很可能", "偶發")))
    write_excel(days, mat, rows, a.out, a.min, a.buffer, matched)
    for d in days:
        act = [t for t in d["tractors"] if t["active"]]
        dr = [r for r in rows if r["day"] == d["label"]]
        print(f"{d['label']}: 車頭 {len(act)} 台、空檔≥{a.min}分 {len(dr)} 筆、合計 {sum(r['dur'] for r in dr)} 分；有可跑站所 {sum(1 for r in dr if r['cands'])} 筆")
    print("站間車程配對", len(mat), "→", a.out)
    if a.json:
        json.dump({"rows": rows, "matrix": {f"{x}|{y}": v for (x, y), v in mat.items()}},
                  open(a.json, "w", encoding="utf8"), ensure_ascii=False, default=str)


if __name__ == "__main__":
    main()
