# -*- coding: utf-8 -*-
"""Parse 曳引車/全拖車 sheets from the embedded workbook into per-tractor timelines.
Colors: FFFF00 yellow = new, 00B050 green = deleted, none/white = original kept."""
import sys, os, json, re, datetime
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

YELLOW, GREEN = "FFFFFF00", "FF00B050"

def fill_color(cell):
    f = cell.fill
    if f is None or f.fill_type != "solid": return "white"
    fg = f.fgColor
    if fg.type == "rgb":
        if fg.rgb == YELLOW: return "yellow"
        if fg.rgb == GREEN: return "green"
        if fg.rgb in ("FFFFFFFF","00000000"): return "white"
        return "other:"+str(fg.rgb)
    if fg.type == "theme" and fg.theme == 0: return "white"
    return f"other:{fg.type}"

def norm(v):
    if v is None: return None
    if isinstance(v, str):
        v = v.strip()
        return v if v != "" else None
    return v

def parse_time(v):
    """return minutes-of-day or None"""
    v = norm(v)
    if v is None: return None
    if isinstance(v, (datetime.datetime, datetime.time)):
        return v.hour*60 + v.minute
    if isinstance(v, (int, float)):
        v = f"{int(v):04d}"
    s = str(v).replace("：",":").replace(" ","")
    m = re.fullmatch(r"(\d{1,2}):?(\d{2})", s)
    if not m:
        return ("BAD", s)
    h, mi = int(m.group(1)), int(m.group(2))
    if h > 24 or mi > 59: return ("BAD", s)
    return (h % 24)*60 + mi

def merged_map(ws):
    """col->(row, start_col) for merged ranges, keyed by (row,col)"""
    mm = {}
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, rng.max_row+1):
            for c in range(rng.min_col, rng.max_col+1):
                mm[(r,c)] = (rng.min_row, rng.min_col)
    return mm

def parse_sheet(ws, remark_col_letter):
    mm = merged_map(ws)
    remark_col = column_index_from_string(remark_col_letter)
    tractors = []
    cur = None
    r = 1
    maxr = ws.max_row
    while r <= maxr:
        if norm(ws.cell(r, 8).value) == "班次代號":
            code = norm(ws.cell(r, 6).value)
            name = norm(ws.cell(r, 7).value)
            if code is not None:
                cur = {"sheet": ws.title, "row": r, "unit": norm(ws.cell(r,3).value),
                       "idx": norm(ws.cell(r,4).value), "mileage": norm(ws.cell(r,5).value),
                       "code": str(code), "name": name,
                       "color": fill_color(ws.cell(r,6)),
                       "name_color": fill_color(ws.cell(r,7)),
                       "remark": norm(ws.cell(r, remark_col).value),
                       "remark_color": fill_color(ws.cell(r, remark_col)),
                       "blocks": []}
                tractors.append(cur)
            if cur is None:
                r += 5; continue
            block = {"row": r, "owner": norm(ws.cell(r,7).value) or norm(ws.cell(r+2,7).value),
                     "driver": norm(ws.cell(r+2,7).value), "points": [], "shifts": []}
            # shift codes on row r (merged), trailer names row r+1, notes r+4
            seen = set()
            for c in range(9, remark_col):
                key = mm.get((r,c), (r,c))
                if key in seen: continue
                seen.add(key)
                sc = norm(ws.cell(key[0], key[1]).value)
                tr = norm(ws.cell(r+1, key[1]).value)
                if sc is None and tr is None: continue
                block["shifts"].append({"col": get_column_letter(key[1]), "code": str(sc) if sc is not None else None,
                                        "trailer": tr, "color": fill_color(ws.cell(key[0], key[1])),
                                        "note": norm(ws.cell(r+4, key[1]).value)})
            for c in range(9, remark_col):
                tcell, scell = ws.cell(r+2, c), ws.cell(r+3, c)
                t, s = parse_time(tcell.value), norm(scell.value)
                if t is None and s is None: continue
                col = fill_color(tcell) if norm(tcell.value) is not None else fill_color(scell)
                key = mm.get((r,c), (r,c))
                sc = norm(ws.cell(key[0], key[1]).value)
                block["points"].append({"col": get_column_letter(c), "time": t, "station": s, "color": col,
                                        "shift": str(sc) if sc is not None else None,
                                        "tcolor": fill_color(tcell), "scolor": fill_color(scell)})
            cur["blocks"].append(block)
            r += 5
        else:
            r += 1
    return tractors

def main(xlsx, outjson):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    res = {}
    for sheet, rc in [("曳引車", None), ("全拖車", None)]:
        ws = wb[sheet]
        # find remark column on header row 3
        rcol = None
        for c in range(1, ws.max_column+1):
            if norm(ws.cell(3,c).value) == "備註": rcol = get_column_letter(c)
        res[sheet] = parse_sheet(ws, rcol)
        print(sheet, "remark col", rcol, "tractors", len(res[sheet]))
    json.dump(res, open(outjson,"w",encoding="utf8"), ensure_ascii=False, indent=1)

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
