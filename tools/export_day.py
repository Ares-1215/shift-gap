# -*- coding: utf-8 -*-
"""
把運務通知 .docx／班次表 .xlsx 解析成網站用的班次表 JSON（gzip），再由 push_to_site.ps1 推上網站。
結構與網站 index.html 的 rawTractors() 相同（col 用字母，網站會轉成數字）。

用法：python tools/export_day.py "運務通知_1106.docx" -o 班次表_1106.json.gz [--label 1106]
"""
import sys, os, json, gzip, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import analyze

RAW_T = ["row", "unit", "idx", "mileage", "code", "name", "color", "remark", "kind"]


def raw_tractors(tractors):
    out = []
    for t in tractors:
        o = {k: t.get(k) for k in RAW_T}
        o["blocks"] = [{
            "row": b["row"], "owner": b.get("owner"),
            "points": [{"col": p["col"], "time": p["time"] if isinstance(p["time"], int) else None,
                        "station": p["station"], "color": p["color"], "shift": p.get("shift")} for p in b["points"]],
            "shifts": [{"col": s["col"], "code": s.get("code"), "trailer": s.get("trailer"),
                        "color": s["color"], "note": s.get("note")} for s in b["shifts"]],
        } for b in t["blocks"]]
        out.append(o)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("-o", "--out", required=True, help="輸出 .json 或 .json.gz")
    ap.add_argument("--label", help="日期標籤 MMDD（預設從檔名抓）")
    a = ap.parse_args()
    day = analyze.load_day(a.file)
    obj = {"v": 1, "label": a.label or day["label"], "file": day["file"], "tractors": raw_tractors(day["tractors"])}
    data = json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8")
    if a.out.lower().endswith(".gz"):
        with gzip.open(a.out, "wb") as f:
            f.write(data)
    else:
        with open(a.out, "wb") as f:
            f.write(data)
    print(f"{day['file']}: 車頭 {len(obj['tractors'])} 台 → {a.out}（{os.path.getsize(a.out) // 1024} KB）")


if __name__ == "__main__":
    main()
