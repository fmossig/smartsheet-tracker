import os, csv, glob, argparse
from datetime import datetime, date, timedelta
from collections import defaultdict, Counter
from typing import List, Dict, Tuple, Optional

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie

def parse_date_safe(s: str) -> Optional[date]:
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%m/%d/%Y", "%Y/%m/%d", "%d.%m.%y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except Exception:
            pass
    try:
        return datetime.fromisoformat((s.strip().split('T')[0])).date()
    except Exception:
        return None

def load_tracker_rows(tracker_folder: str) -> List[Dict[str, str]]:
    rows = []
    for path in glob.glob(os.path.join(tracker_folder, "date_changes_log_*.csv")):
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows.extend(reader)
    return rows

def within(d: date, start: date, end: date) -> bool:
    return (d is not None) and (start <= d <= end)

def aggregate(rows: List[Dict[str,str]], start: date, end: date):
    frows = []
    for r in rows:
        d = parse_date_safe(r.get("Datum","") or r.get("date",""))
        if within(d, start, end):
            rr = dict(r); rr["_parsed_date"] = d
            frows.append(rr)

    by_group = defaultdict(list)
    for r in frows:
        by_group[r.get("Produktgruppe","") or r.get("group","")].append(r)

    total_events = len(frows)
    phases = Counter()
    users = Counter()
    countries = Counter()
    for r in frows:
        p = str(r.get("Phase","") or r.get("phase","") or "").strip()
        phases[p] += 1
        u = (r.get("Mitarbeiter","") or r.get("user","") or "").strip()
        if u:
            users[u] += 1
        c = (r.get("Land/Marketplace","") or r.get("land","") or "").strip()
        if c:
            countries[c] += 1

    return {"window_rows": frows, "by_group": by_group, "total_events": total_events, "phases": phases, "users": users, "countries": countries}

def make_pie(data_pairs: List[Tuple[str,int]], w=100*mm, h=100*mm):
    d = Drawing(w, h)
    pie = Pie()
    pie.x = 10; pie.y = 10
    pie.width = w-20; pie.height = h-20
    values = [v for _,v in data_pairs] or [1]
    labels = [k for k,_ in data_pairs] or ["Keine Daten"]
    pie.data = values
    pie.labels = [f"{labels[i]} ({values[i]})" for i in range(len(values))]
    d.add(pie)
    return d

def make_bar(categories: List[str], series: List[int], w=160*mm, h=90*mm, title: str=""):
    d = Drawing(w, h)
    chart = VerticalBarChart()
    chart.x = 40; chart.y = 30
    chart.width = w-60; chart.height = h-60
    chart.data = [series or [0]]
    chart.categoryAxis.categoryNames = categories or ["Keine Daten"]
    chart.barSpacing = 2; chart.groupSpacing = 8
    chart.valueAxis.labels.boxAnchor = 'e'
    if title:
        d.add(String(0, h-10, title))
    d.add(chart)
    return d

def render_pdf(out_path: str, start: date, end: date, agg: Dict):
    styles = getSampleStyleSheet()
    # kollisionssichere Styles
    try:
        styles.add(ParagraphStyle(name="CustomH1", fontSize=18, leading=22, spaceAfter=8))
    except Exception:
        pass
    try:
        styles.add(ParagraphStyle(name="CustomH2", fontSize=13, leading=16, spaceBefore=10, spaceAfter=6))
    except Exception:
        pass
    try:
        styles.add(ParagraphStyle(name="CustomMuted", fontSize=9, textColor=colors.grey))
    except Exception:
        pass

    h1 = styles.get("CustomH1") or styles["Title"]
    h2 = styles.get("CustomH2") or styles["Heading2"]
    muted = styles.get("CustomMuted") or styles["Normal"]

    doc = SimpleDocTemplate(out_path, pagesize=A4, leftMargin=16*mm, rightMargin=16*mm, topMargin=18*mm, bottomMargin=18*mm)
    story = []

    period_label = f"{start.strftime('%d.%m.%Y')} – {end.strftime('%d.%m.%Y')}"
    story += [Paragraph(f"<b>Status Report</b> – {period_label}", h1), Spacer(1, 4*mm)]
    story.append(Paragraph(f"Gesamt protokollierte Änderungen: <b>{agg['total_events']}</b>", styles["Normal"]))
    story.append(Spacer(1, 2*mm))

    phase_pairs = sorted(agg["phases"].items(), key=lambda kv: (kv[0] or ""))
    story.append(Paragraph("Phasenverteilung", h2))
    story.append(make_pie(phase_pairs)); story.append(Spacer(1, 4*mm))

    top_users = agg["users"].most_common(15)
    story.append(Paragraph("Beiträge nach Mitarbeiter (Top 15)", h2))
    story.append(make_bar([u for u,_ in top_users], [v for _,v in top_users], title="Änderungen"))
    story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Aktivität nach Land/Marketplace (Top 20)", h2))
    rows = [["Land/Marketplace", "Änderungen"]]
    for c, n in agg["countries"].most_common(20):
        rows.append([c, str(n)])
    table = Table(rows, colWidths=[90*mm, 30*mm])
    table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.grey),
                               ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke),
                               ("ALIGN",(1,1),(-1,-1),"RIGHT")]))
    story.append(table); story.append(Spacer(1, 4*mm))

    story.append(Paragraph("Aktivität je Produktgruppe", h2))
    for g, rows_g in agg["by_group"].items():
        story.append(Paragraph(f"Produktgruppe: <b>{g or '–'}</b>", styles["Normal"]))
        story.append(Spacer(1, 1*mm))
        story.append(make_bar([g], [len(rows_g)], title="Änderungen insgesamt"))
        story.append(Spacer(1, 3*mm))

    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("Quelle: tracker_logs/*.csv | Struktur wie bisheriger PDF-Report", muted))

    doc.build(story)

def main():
    ap = argparse.ArgumentParser(description="Erzeuge wöchentliche/monatliche Status-PDFs aus tracker_logs.")
    ap.add_argument("--tracker-logs", default="tracker_logs", help="Ordner mit date_changes_log_*.csv")
    group = ap.add_mutually_exclusive_group(required=True)
    group.add_argument("--week", help="ISO-Wochenlabel z.B. 2025-W37 (ganze abgeschlossene Woche Mo–So)")
    group.add_argument("--month", help="Monatslabel z.B. 2025-08 (ganzer Kalendermonat)")
    group.add_argument("--from", dest="date_from", help="Startdatum YYYY-MM-DD")
    ap.add_argument("--to", dest="date_to", help="Enddatum YYYY-MM-DD (inklusive)")
    ap.add_argument("--out-dir", default="reports", help="Basis-Ausgabeordner")
    args = ap.parse_args()

    if args.week:
        y = int(args.week.split("-W")[0]); w = int(args.week.split("-W")[1])
        d = datetime.fromisocalendar(y, w, 1).date()
        start = d; end = d + timedelta(days=6)
        out_dir = os.path.join(args.out_dir, "weekly", f"{y}", f"{args.week}")
        out_name = f"report_weekly_{args.week}.pdf"
    elif args.month:
        y, m = [int(x) for x in args.month.split("-")]
        start = date(y, m, 1)
        end = date(y, 12, 31) if m == 12 else (date(y, m+1, 1) - timedelta(days=1))
        out_dir = os.path.join(args.out_dir, "monthly", f"{y}")
        out_name = f"report_monthly_{y}-{m:02d}.pdf"
    else:
        if not args.date_to:
            raise SystemExit("--to ist erforderlich wenn --from benutzt wird.")
        start = datetime.strptime(args.date_from, "%Y-%m-%d").date()
        end = datetime.strptime(args.date_to, "%Y-%m-%d").date()
        out_dir = os.path.join(args.out_dir, "custom")
        out_name = f"report_{start.isoformat()}_{end.isoformat()}.pdf"

    os.makedirs(out_dir, exist_ok=True)
    rows = load_tracker_rows(args.tracker_logs)
    agg = aggregate(rows, start, end)
    out_path = os.path.join(out_dir, out_name)
    render_pdf(out_path, start, end, agg)
    print(out_path)

if __name__ == "__main__":
    main()
