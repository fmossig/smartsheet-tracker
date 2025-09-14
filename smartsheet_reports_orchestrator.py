from datetime import date, timedelta, datetime
import os, sys, subprocess, glob, re

def run(cmd, env=None):
    print("+", " ".join(cmd))
    res = subprocess.run(cmd, check=False, env=env)
    if res.returncode != 0:
        raise SystemExit(res.returncode)

def week_bounds(iso_week: str):
    y = int(iso_week.split("-W")[0]); w = int(iso_week.split("-W")[1])
    start = datetime.fromisocalendar(y, w, 1).date()  # Montag
    end   = start + timedelta(days=6)                 # Sonntag
    return start, end

def month_bounds(ym: str):
    y, m = [int(x) for x in ym.split("-")]
    start = date(y, m, 1)
    end   = date(y, 12, 31) if m == 12 else (date(y, m+1, 1) - timedelta(days=1))
    return start, end

def previous_month(today: date) -> str:
    y, m = today.year, today.month - 1
    if m == 0: y -= 1; m = 12
    return f"{y}-{m:02d}"

def iso_week_of_previous_week(today: date) -> str:
    last_monday = today - timedelta(days=today.weekday()+7)
    y, w, _ = last_monday.isocalendar()
    return f"{y}-W{w:02d}"

def resolve_tracker_script():
    for p in ["smartsheet_date_change_tracker.py",
              "smartsheet_date_change_tracker(1).py",
              "smartsheet_date_change_tracker (1).py"]:
        if os.path.exists(p): return p
    hits = sorted(glob.glob("smartsheet_date_change_tracker*.py"))
    if hits: return hits[0]
    raise FileNotFoundError("smartsheet_date_change_tracker*.py nicht gefunden.")

def resolve_status_report_script():
    for p in ["smartsheet_status_report.py",
              "smartsheet_status_report(1).py",
              "smartsheet_status_report (1).py"]:
        if os.path.exists(p): return p
    hits = sorted(glob.glob("smartsheet_status_report*.py"))
    if hits: return hits[0]
    raise FileNotFoundError("smartsheet_status_report*.py nicht gefunden.")

def trim_tracker_logs_to_last_n_months(folder: str, n: int):
    files = glob.glob(os.path.join(folder, "date_changes_log_*.csv"))
    def parse_my(p):
        m = re.search(r"date_changes_log_(\d{2})\.(\d{4})\.csv$", p)
        if not m: return None
        mm, yy = int(m.group(1)), int(m.group(2))
        return yy*100 + mm, p
    dated = [x for x in (parse_my(p) for p in files) if x]
    dated.sort(reverse=True)
    keep = set(p for _, p in dated[:n])
    for _, p in dated[n:]:
        os.remove(p)

def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("mode", choices=["nightly","weekly","monthly","bootstrap"])
    ap.add_argument("--tracker-logs", default="tracker_logs")
    ap.add_argument("--reports-dir", default="reports")
    ap.add_argument("--months", type=int, default=3)
    ap.add_argument("--python", default=sys.executable)
    args = ap.parse_args()

    tracker = resolve_tracker_script()
    status  = resolve_status_report_script()
    today   = date.today()

    if args.mode == "nightly":
        run([args.python, tracker, "--bootstrap", "--dedupe"])
        return

    if args.mode == "weekly":
        week = iso_week_of_previous_week(today)
        start, end = week_bounds(week)
        out_dir = os.path.join(args.reports_dir, "weekly", f"{start.year}", week)
        os.makedirs(out_dir, exist_ok=True)
        env = os.environ.copy()
        env["SMARTSHEET_TOKEN"] = env.get("SMARTSHEET_TOKEN") or env.get("SMARTSHEET_ACCESS_TOKEN","")
        env["REPORT_SINCE"]  = start.isoformat()
        env["REPORT_UNTIL"]  = end.isoformat()
        env["REPORT_LABEL"]  = week
        env["REPORT_OUTDIR"] = out_dir
        run([args.python, status], env=env)
        return

    if args.mode == "monthly":
        ym = previous_month(today)
        start, end = month_bounds(ym)
        out_dir = os.path.join(args.reports_dir, "monthly", f"{start.year}")
        os.makedirs(out_dir, exist_ok=True)
        env = os.environ.copy()
        env["SMARTSHEET_TOKEN"] = env.get("SMARTSHEET_TOKEN") or env.get("SMARTSHEET_ACCESS_TOKEN","")
        env["REPORT_SINCE"]  = start.isoformat()
        env["REPORT_UNTIL"]  = end.isoformat()
        env["REPORT_LABEL"]  = ym
        env["REPORT_OUTDIR"] = out_dir
        run([args.python, status], env=env)
        return

    if args.mode == "bootstrap":
        # 1) Logs aufbauen
        run([args.python, tracker, "--bootstrap"])
        # 2) Logs auf letzte N Monate kürzen
        trim_tracker_logs_to_last_n_months(args.tracker_logs, args.months)
        # 3) Für alle verbleibenden Monats-Logs PDFs bauen
        files = glob.glob(os.path.join(args.tracker_logs, "date_changes_log_*.csv"))
        months = set()
        for p in files:
            m = re.search(r"date_changes_log_(\d{2})\.(\d{4})\.csv$", os.path.basename(p))
            if m: months.add(f"{m.group(2)}-{m.group(1)}")
        for ym in sorted(months):
            start, end = month_bounds(ym)
            out_dir = os.path.join(args.reports_dir, "monthly", f"{start.year}")
            os.makedirs(out_dir, exist_ok=True)
            env = os.environ.copy()
            env["SMARTSHEET_TOKEN"] = env.get("SMARTSHEET_TOKEN") or env.get("SMARTSHEET_ACCESS_TOKEN","")
            env["REPORT_SINCE"]  = start.isoformat()
            env["REPORT_UNTIL"]  = end.isoformat()
            env["REPORT_LABEL"]  = ym
            env["REPORT_OUTDIR"] = out_dir
            run([args.python, status], env=env)
        return

if __name__ == "__main__":
    main()
