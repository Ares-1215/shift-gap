# -*- coding: utf-8 -*-
"""
每日自動：推測「明天、後天」的加班需求清單（ot_forecast.py），
若桌面「運務」資料夾已有該日的運務通知 .docx，順便跑 analyze.py --needs 產出配車 Excel。

由 Windows 工作排程 ShiftGap-Forecast 每天 13:40 呼叫（OpsHub-Daily 13:00 同步 205 之後）。
手動：python tools/daily_forecast.py [--days 2] [--date YYYYMMDD]
"""
import sys, os, re, glob, shutil, argparse, datetime, subprocess
from pathlib import Path

HERE = Path(__file__).resolve().parent
PY = sys.executable
DESK = Path(r"C:\Users\26516\Desktop\運務")
OUT = DESK / "加班預測"
LOG = HERE / "daily.log"


def log(msg):
    line = f"[{datetime.datetime.now():%Y-%m-%d %H:%M}] {msg}"
    print(line)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def run(args):
    r = subprocess.run([PY, *args], capture_output=True, text=True, encoding="utf-8", errors="replace",
                       env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    out = (r.stdout or "") + (r.stderr or "")
    for ln in out.strip().splitlines():
        log("  " + ln)
    return r.returncode == 0


PUSH = HERE / "push_to_site.ps1"


def push(path, kind, day=None, name=None):
    """把檔案推到網站（Edge Function shiftgap，同事開網站就看得到）；管理碼在 tools/site.local.json，沒有就略過"""
    if not (HERE / "site.local.json").exists():
        log("  （沒有 tools/site.local.json，略過推送網站）")
        return False
    args = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(PUSH), "-Path", str(path), "-Kind", kind]
    if day:
        args += ["-Day", day]
    if name:
        args += ["-Name", name]
    r = subprocess.run(args, capture_output=True, text=True, encoding="utf-8", errors="replace")
    for ln in ((r.stdout or "") + (r.stderr or "")).strip().splitlines():
        log("  " + ln)
    return r.returncode == 0


def find_docx(ymd):
    """該日的班次表：運務通知 .docx 或班次表 .xlsx（檔名含 MMDD 或 YYYYMMDD），同一天有多個取最新修改的"""
    pat = re.compile(rf"(?<!\d)(?:{ymd}|{ymd[4:]})(?!\d)")
    cands = [p for p in list(DESK.glob("*.docx")) + list(DESK.glob("*.xlsx"))
             if not p.name.startswith("~$") and not p.name.startswith("班次空檔分析") and pat.search(p.name)]
    return sorted(cands, key=lambda p: p.stat().st_mtime)[-1:]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=2, help="推測未來幾天（預設明天、後天）")
    ap.add_argument("--date", help="指定第一個目標日 YYYYMMDD（預設明天）")
    a = ap.parse_args()
    OUT.mkdir(parents=True, exist_ok=True)
    first = datetime.datetime.strptime(a.date, "%Y%m%d").date() if a.date else datetime.date.today() + datetime.timedelta(days=1)
    targets = [(first + datetime.timedelta(days=i)).strftime("%Y%m%d") for i in range(a.days)]
    tag = "-".join(t[4:] for t in targets)
    csv_path = OUT / f"加班預測_{tag}.csv"
    log(f"=== 推測 {targets} ===")
    ok = run([str(HERE / "ot_forecast.py"), *targets, "-o", str(csv_path)])
    if not ok or not csv_path.exists():
        log("ot_forecast 失敗")
        return 1
    shutil.copy(csv_path, OUT / "加班預測_最新.csv")
    log(f"清單 → {csv_path}")
    push(csv_path, "needs")   # 推到網站：開網站就有配車結果
    # 有通知就配車
    docs = []
    for t in targets:
        docs += [p for p in find_docx(t) if p not in docs]
    if docs:
        days_tag = "-".join(sorted({m.group(0) for p in docs for m in [re.search(r"(?<!\d)(1[0-2]|0?[1-9])(0[1-9]|[12]\d|3[01])(?!\d)", p.name)] if m}))
        xlsx = DESK / f"班次空檔分析_{days_tag or tag}.xlsx"
        log(f"找到通知 {[p.name for p in docs]} → 配車")
        run([str(HERE / "analyze.py"), *map(str, docs), "-o", str(xlsx), "--needs", str(csv_path)])
        log(f"配車 Excel → {xlsx}")
        for t in targets:   # 桌面有的班次表也推到網站：先解析成 JSON（約 100KB），同一天會覆蓋
            for p in find_docx(t):
                gz = OUT / f"班次表_{t[4:]}.json.gz"
                if run([str(HERE / "export_day.py"), str(p), "-o", str(gz), "--label", t[4:]]):
                    push(gz, "schedule", t[4:], name=p.name)
    else:
        log("桌面運務資料夾沒有這幾天的運務通知 .docx，只產清單（之後把通知放進去再跑一次，或上傳網站）")
    # 清掉 30 天前的舊清單
    for p in OUT.glob("加班預測_*.csv"):
        if p.name != "加班預測_最新.csv" and (datetime.date.today() - datetime.date.fromtimestamp(p.stat().st_mtime)).days > 30:
            p.unlink()
    log("=== 完成 ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
