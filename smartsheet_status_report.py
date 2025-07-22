import csv, os
from datetime import datetime, timedelta, timezone
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.graphics.shapes import Drawing, String, Rect
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from collections import defaultdict

# Farb‑Mapping Produktgruppen
COLORS = {
    "NA": "#E63946",
    "NF": "#457B9D",
    "NH": "#2A9D8F",
    "NM": "#E9C46A",
    "NP": "#F4A261",
    "NT": "#9D4EDD",
    "NV": "#00B4D8"
}

# Farb‑Mapping Mitarbeiter
EMP_COLORS = {
    "DM": "#223459",
    "EK": "#6A5AAA",
    "HI": "#B45082",
    "SM": "#F9767F",
    "JHU": "#FFB142",
    "LK": "#FFDE70",
}

def read_snapshot_counts(date_str):
    """
    Summiert alle Phase-Events pro Produktgruppe.
    """
    snap_file = os.path.join("status", f"status_snapshot_{date_str}.csv")
    counts = {g: 0 for g in COLORS}
    with open(snap_file, encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader)  # header
        for row in reader:
            grp = row[0]
            if grp in counts:
                counts[grp] += 1
    # Reihenfolge wie COLORS
    return [counts[g] for g in COLORS]

def read_na_phase_employee(date_str):
    """
    Liefert dict: phase -> {emp: count} für Produktgruppe NA (nur definierte EMP_COLORS).
    """
    snap_file = os.path.join("status", f"status_snapshot_{date_str}.csv")
    counts = defaultdict(lambda: defaultdict(int))
    with open(snap_file, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["Produktgruppe"] != "NA":
                continue
            phase = int(row["Phase"])
            emp = (row["Mitarbeiter"] or "").strip()
            if emp in EMP_COLORS:
                counts[phase][emp] += 1
    return counts

def make_report():
    # --- Datum / Files ---
    now = datetime.now(timezone.utc)
    today = now.date()
    date_str = today.isoformat()
    cutoff = today - timedelta(days=30)
    pdf_file = os.path.join("status", f"status_report_{date_str}.pdf")

    # --- Report Setup ---
    doc = SimpleDocTemplate(
        pdf_file,
        pagesize=A4,
        leftMargin=20*mm, rightMargin=20*mm,
        topMargin=20*mm, bottomMargin=20*mm
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CoverTitle', parent=styles['Title'], fontName='Helvetica-Bold', fontSize=20, spaceAfter=12))
    styles.add(ParagraphStyle(name='CoverInfo',  parent=styles['Normal'], fontName='Helvetica',      fontSize=12, spaceAfter=6))
    styles.add(ParagraphStyle(name='ChartTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=14, spaceAfter=6))

    elems = []

    # --- Deckblatt ---
    elems.append(Paragraph("Amazon Content Management - Activity Report", styles['CoverTitle']))
    elems.append(Paragraph(f"Erstellungsdatum: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['CoverInfo']))
    elems.append(Paragraph(f"Abgrenzungsdatum: {cutoff.isoformat()}", styles['CoverInfo']))
    elems.append(PageBreak())

    # --- Chart 1: Produktgruppen ---
    groups = list(COLORS.keys())
    values = read_snapshot_counts(date_str)

    elems.append(Paragraph("Anzahl an eröffneten Phasen pro Produktgruppe (letzte 30 Tage)", styles['ChartTitle']))
    elems.append(Spacer(1, 6*mm))

    usable_width = A4[0] - 2*20*mm
    chart_height = 60*mm
    origin_x = 0
    origin_y = 15*mm

    total_gap = usable_width * 0.1
    gap = total_gap / (len(groups) + 1)
    bar_width = (usable_width - total_gap) / len(groups)

    d1 = Drawing(usable_width, chart_height + origin_y + 2*mm)
    max_val = max(values) if values else 1

    for idx, grp in enumerate(groups):
        val = values[idx]
        x = origin_x + gap * (idx + 1) + bar_width * idx
        h = (val / max_val) * chart_height

        d1.add(Rect(x, origin_y, bar_width, h,
                    fillColor=colors.HexColor(COLORS[grp]), strokeColor=None))
        d1.add(String(x + bar_width/2, origin_y + h + 4, str(val),
                      fontName='Helvetica', fontSize=9, textAnchor='middle'))
        d1.add(String(x + bar_width/2, origin_y - 10, grp,
                      fontName='Helvetica', fontSize=8, textAnchor='middle'))

    elems.append(d1)

    # --- Chart 2: Gestapeltes Balkendiagramm NA ---
    counts = read_na_phase_employee(date_str)
    phases_sorted = [1, 2, 3, 4, 5]
    emp_sorted = [e for e in EMP_COLORS if any(counts[p][e] for p in phases_sorted)]

    elems.append(PageBreak())
    elems.append(Paragraph("Mitarbeiterbasierte Phasenstatistik (NA, 30 Tage)", styles['ChartTitle']))
    elems.append(Spacer(1, 6*mm))

    usable_width2   = A4[0] - 2*15*mm
    left_axis_space = 18*mm
    row_h           = 8*mm
    gap_y           = 5*mm
    origin_x2       = 0
    origin_y2       = 10*mm

    total_h = len(phases_sorted) * (row_h + gap_y) + origin_y2 + 10*mm
    d2 = Drawing(usable_width2, total_h)

    max_total = max((sum(counts[p][e] for e in emp_sorted) for p in phases_sorted), default=1)

    for i, phase in enumerate(phases_sorted):
        y = origin_y2 + (len(phases_sorted) - 1 - i) * (row_h + gap_y)
        x = left_axis_space
        d2.add(String(0, y + row_h/2, f"Phase {phase}",
                      fontName='Helvetica', fontSize=9, textAnchor='start'))

        run_w = 0
        for emp in emp_sorted:
            val = counts[phase][emp]
            if val == 0:
                continue
            seg_w = (val / max_total) * (usable_width2 - left_axis_space - 5*mm)
            rect = Rect(x + run_w, y, seg_w, row_h,
                        fillColor=colors.HexColor(EMP_COLORS[emp]), strokeColor=None)
            d2.add(rect)

            if seg_w > 12:
                d2.add(String(x + run_w + seg_w/2, y + row_h/2, str(val),
                              fontName='Helvetica-Bold', fontSize=8,
                              textAnchor='middle', fillColor=colors.white))
            run_w += seg_w

    # Legende
    legend_x = usable_width2 - 40*mm
    legend_y = total_h - 10*mm
    box_h = 5*mm
    for j, emp in enumerate(emp_sorted):
        yy = legend_y - j*(box_h + 2*mm)
        d2.add(Rect(legend_x, yy, box_h, box_h,
                    fillColor=colors.HexColor(EMP_COLORS[emp]), strokeColor=None))
        d2.add(String(legend_x + box_h + 3*mm, yy + box_h/2,
                      emp, fontName='Helvetica', fontSize=8, textAnchor='start'))

    elems.append(d2)
    elems.append(Spacer(1, 6*mm))

    # --- Fußzeile ---
    elems.append(Paragraph(f"Report erstellt: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['Normal']))

    doc.build(elems)
    print(f"✅ PDF Report erstellt: {pdf_file}")

if __name__ == "__main__":
    make_report()
