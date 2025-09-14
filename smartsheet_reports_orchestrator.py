
import os, sys, subprocess, argparse, glob, shutil
from datetime import datetime, date, timedelta

def iso_week_of_previous_week(today: date) -> str:
    # last completed ISO week (Mon-Sun)
    last_monday = today - timedelta(days=today.weekday()+7)
    y, w, _ = last_monday.isocalendar()
    return f"{y}-W{w:02d}"

def previous_month(today: date) -> str:
    y = today.year
    m = today.month - 1
    if m == 0:
        m = 12
        y -= 1
    return f"{y}-{m:02d}"

def run(cmd: list):
    print("+", " ".join(cmd))
    res = subprocess.run(cmd, check=False)
    if res.returncode != 0:
        raise SystemExit(res.returncode)

def trim_tracker_logs_to_last_n_months(folder: str, n: int):
    # filenames: date_changes_log_MM.YYYY.csv
    import re
    from datetime import datetime
    files = glob.glob(os.path.join(folder, "date_changes_log_*.csv"))
    def parse_my(p):
        m = re.search(r"date_changes_log_(\d{2})\.(\d{4})\.csv$", p)
        if not m:
            return None
        mm = int(m.group(1)); yy = int(m.group(2))
        return yy*100 + mm, p
    dated = [parse_my(p) for p in files]
    dated = [x for x in dated if x is not None]
    dated.sort(reverse=True)
    # keep n most recent months
    keep = set(p for _,p in dated[:n])
    deleted = []
    for _, p in dated[n:]:
        os.remove(p)
        deleted.append(p)
    return deleted

def main():
    ap = argparse.ArgumentParser(description="Orchestrator für Tracker + Reports")
    ap.add_argument("mode", choices=["nightly","weekly","monthly","bootstrap"], help="Was ausführen?")
    ap.add_argument("--python", default=sys.executable, help="Welches Python für Subprozesse")
    ap.add_argument("--tracker-script", default="smartsheet_date_change_tracker.py")
    ap.add_argument("--report-script", default="smartsheet_status_report.py")
    ap.add_argument("--tracker-logs", default="tracker_logs")
    ap.add_argument("--reports-dir", default="reports")
    ap.add_argument("--months", type=int, default=3, help="Für bootstrap: Anzahl Monate behalten")
    args = ap.parse_args()

    today = date.today()

    if args.mode == "nightly":
        # just run incremental tracker
        run([args.python, args.tracker_script])
        return

    if args.mode == "weekly":
        week = iso_week_of_previous_week(today)
        run([args.python, args.report_script, "--week", week, "--tracker-logs", args.tracker_logs, "--out-dir", args.reports_dir])
        return

    if args.mode == "monthly":
        month = previous_month(today)
        run([args.python, args.report_script, "--month", month, "--tracker-logs", args.tracker_logs, "--out-dir", args.reports_dir])
        return

    if args.mode == "bootstrap":
        # 1) run bootstrap to fill logs
        run([args.python, args.tracker_script, "--bootstrap"])
        # 2) trim tracker_logs to last N months
        deleted = trim_tracker_logs_to_last_n_months(args.tracker_logs, args.months)
        print(f"Removed {len(deleted)} old monthly tracker logs.")
        # 3) build a report for each kept month
        # Re-scan the folder and find remaining files
        import re
        files = glob.glob(os.path.join(args.tracker_logs, "date_changes_log_*.csv"))
        months = []
        for p in files:
            m = re.search(r"date_changes_log_(\d{2})\.(\d{4})\.csv$", os.path.basename(p))
            if m:
                months.append(f"{m.group(2)}-{m.group(1)}")
        for m in sorted(set(months)):
            run([args.python, args.report_script, "--month", m, "--tracker-logs", args.tracker_logs, "--out-dir", args.reports_dir])
        return

if __name__ == "__main__":
    main()
