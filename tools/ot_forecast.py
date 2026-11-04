# -*- coding: utf-8 -*-
"""
加班需求推測：用 nls 205「發送封櫃明細」的歷史加班／追加櫃，
依「同星期、近 N 週」統計每條起迄站的發生機率、預估櫃數、發車時間窗，
產出「加班需求清單」CSV（可上傳到 shift-gap 網站的加班配車，或給 analyze.py --needs）。

用法：
  python ot_forecast.py 20261105 20261106 -o 加班預測_1105-1106.csv [--weeks 8] [--min 0.25]
"""
import sys, os, re, csv, argparse, datetime, statistics, collections
from pathlib import Path

NLS = Path(r"C:\Users\26516\notebookLM\nls-explorer")
sys.path.insert(0, str(NLS))
os.chdir(NLS)
import colmap                       # noqa: E402
from common import files, names_for  # noqa: E402

DOW = "一二三四五六日"


def load205(key):
    f = NLS / "data" / "205" / f"{key}.csv"
    txt = f.read_text(encoding="utf-8")
    return colmap.read_csv_named(txt, names_for("205")) if txt.strip() else []


def station_name(s):
    m = re.match(r"^\s*(\d{4})\s*(.*?)\s*$", s or "")
    return (m.group(2) or m.group(1)) if m else (s or "").strip()


def hhmm_to_min(t):
    m = re.search(r"(\d{1,2}):(\d{2})", t or "")
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def op(m):
    """營運時間：12:00 為一天起點（12:00→720、03:10→1630）"""
    return (m - 720) % 1440 + 720


def fmt(m):
    m = int(m) % 1440
    return f"{m // 60:02d}{m % 60:02d}"


def load_all():
    dayfiles = {k for k in files("205") if re.fullmatch(r"\d{8}", k)}
    seen, recs, alldays = set(), [], set()
    for key in files("205"):
        if key == "header":
            continue
        for r in load205(key):
            day = r.get("發送日", "")
            if not re.fullmatch(r"\d{8}", day):
                continue
            if not re.fullmatch(r"\d{8}", key) and day in dayfiles:
                continue
            alldays.add(day)
            cat = r.get("類別", "")
            if "加班" not in cat and "追加" not in cat:
                continue
            k = (day, r.get("班次代號", ""), r.get("櫃號", ""), r.get("起站", ""))
            if k in seen:
                continue
            seen.add(k)
            t = hhmm_to_min(r.get("實際離站時間"))
            src = "離站"
            if t is None:
                t = hhmm_to_min((r.get("封櫃時間") or "")[-8:])
                src = "封櫃"
            recs.append({"day": day, "cat": cat, "from": station_name(r.get("起站")), "to": station_name(r.get("迄站")),
                         "code": r.get("班次代號", ""), "name": r.get("班次名稱", ""), "box": r.get("櫃號", ""),
                         "load": r.get("發送積載率", ""), "t": None if t is None else op(t), "tsrc": src})
    return recs, sorted(alldays)


def pick_days(alldays, recs, target, weeks):
    """目標日同星期、最近 weeks 個營業日；加班櫃數低於同星期中位數 4 成的（假日）剔除再往前補"""
    wd = datetime.date(int(target[:4]), int(target[4:6]), int(target[6:])).weekday()
    cand = [d for d in alldays if d < target and datetime.date(int(d[:4]), int(d[4:6]), int(d[6:])).weekday() == wd]
    cnt = collections.Counter(r["day"] for r in recs)
    med = statistics.median([cnt[d] for d in cand]) if cand else 0
    used, skipped = [], []
    for d in reversed(cand):
        if cnt[d] < med * 0.4:
            skipped.append((d, cnt[d]))
            continue
        used.append(d)
        if len(used) >= weeks:
            break
    return wd, sorted(used), skipped, med


def forecast(recs, alldays, target, weeks, pmin):
    wd, used, skipped, med = pick_days(alldays, recs, target, weeks)
    n = len(used)
    rs = [r for r in recs if r["day"] in used]
    # 全期同星期（參考）
    allwd = [d for d in alldays if d < target and datetime.date(int(d[:4]), int(d[4:6]), int(d[6:])).weekday() == wd]
    rs_all = [r for r in recs if r["day"] in set(allwd)]
    byroute = collections.defaultdict(list)
    for r in rs:
        byroute[(r["from"], r["to"])].append(r)
    all_days_route = collections.defaultdict(set)
    for r in rs_all:
        all_days_route[(r["from"], r["to"])].add(r["day"])
    rows = []
    for (a, b), lst in byroute.items():
        days = {r["day"] for r in lst}
        p = len(days) / n
        if p < pmin:
            continue
        perday = collections.Counter(r["day"] for r in lst)
        boxes = statistics.median(perday.values())
        ts = sorted(r["t"] for r in lst if r["t"] is not None)
        if ts:
            q = statistics.quantiles(ts, n=4) if len(ts) >= 4 else [ts[0], ts[len(ts) // 2], ts[-1]]
            tmed, t1, t2 = statistics.median(ts), q[0], q[2]
        else:
            tmed = t1 = t2 = None
        tier = "幾乎每週" if p >= 0.75 else "很可能" if p >= 0.5 else "偶發"
        cats = collections.Counter(r["cat"] for r in lst)
        names = collections.Counter(r["name"].strip() for r in lst if r["name"].strip())
        rows.append({"日期": target[4:], "星期": DOW[wd], "起站": a, "迄站": b, "機率": round(p, 2), "等級": tier,
                     "出現天數": len(days), "樣本天數": n, "預估櫃數": boxes, "最多櫃數": max(perday.values()),
                     "時間中位": fmt(tmed) if tmed is not None else "", "時間窗起": fmt(t1) if t1 is not None else "", "時間窗迄": fmt(t2) if t2 is not None else "",
                     "時間樣本": len(ts), "類別": "/".join(f"{k}{v}" for k, v in cats.most_common()),
                     "常見班次": "、".join(k for k, _ in names.most_common(3)),
                     "全期機率": round(len(all_days_route[(a, b)]) / max(1, len(allwd)), 2), "全期天數": f"{len(all_days_route[(a, b)])}/{len(allwd)}"})
    rows.sort(key=lambda x: (-x["機率"], -x["預估櫃數"], x["起站"], x["迄站"]))
    return rows, used, skipped, med


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dates", nargs="+", help="目標日 YYYYMMDD")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--weeks", type=int, default=8)
    ap.add_argument("--min", type=float, default=0.25)
    ap.add_argument("--check", help="班次空檔分析的 result.json，用來檢查站名是否對得上")
    a = ap.parse_args()
    recs, alldays = load_all()
    print(f"205 資料 {alldays[0]}～{alldays[-1]} 共 {len(alldays)} 天，加班/追加櫃 {len(recs)} 筆（時間來源：離站 {sum(1 for r in recs if r['tsrc']=='離站')}、封櫃 {sum(1 for r in recs if r['tsrc']=='封櫃')}）")
    out = []
    for target in a.dates:
        rows, used, skipped, med = forecast(recs, alldays, target, a.weeks, a.min)
        print(f"== {target} 星期{DOW[datetime.date(int(target[:4]), int(target[4:6]), int(target[6:])).weekday()]}：採用 {used}；剔除疑似假日 {skipped}（同星期中位 {med:.0f} 櫃/日）")
        tiers = collections.Counter(r["等級"] for r in rows)
        print(f"   路線 {len(rows)} 條：{dict(tiers)}；預估櫃數合計 {sum(r['預估櫃數'] for r in rows):.0f}")
        out += rows
    cols = ["日期", "星期", "起站", "迄站", "機率", "等級", "出現天數", "樣本天數", "預估櫃數", "最多櫃數", "時間中位", "時間窗起", "時間窗迄", "時間樣本", "類別", "常見班次", "全期機率", "全期天數"]
    outp = Path(a.out) if os.path.isabs(a.out) else Path(os.environ.get("ORIG_CWD", ".")) / a.out
    with open(outp, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(out)
    print("→", outp)
    if a.check:
        import json
        R = json.load(open(a.check, encoding="utf8"))
        sched = {r["station"] for r in R["rows"]}
        for k in R.get("matrix", {}):
            sched.update(k.split("|"))
        names = collections.Counter()
        for r in out:
            names[r["起站"]] += r["預估櫃數"]
            names[r["迄站"]] += r["預估櫃數"]
        miss = [(k, v) for k, v in names.most_common() if k not in sched]
        print(f"站名檢查：預測用到 {len(names)} 站，班次表沒有的 {len(miss)} 站（依櫃數排）：", miss[:40])


if __name__ == "__main__":
    main()
