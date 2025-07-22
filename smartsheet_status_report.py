import csv, os
from datetime import datetime, timedelta, timezone
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.graphics.shapes import Drawing, String, Rect
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm

# Farb‑Mapping
COLORS = {
    "NA": "#E63946",
    "NF": "#457B9D",
    "NH": "#2A9D8F",
    "NM": "#E9C46A",
    "NP": "#F4A261",
    "NT": "#9D4EDD",
    "NV": "#00B4D8"
}

def read_snapshot(date_str):
    """
    Liest die Snapshot-CSV und zählt die Anzahl der Phasen pro Produktgruppe.
    """
    snap_file = os.path.join("status", f"status_snapshot_{date_str}.csv")
    counts = {g: 0 for g in COLORS}
    with open(snap_file, encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader)
        for row in reader:
            grp = row[0]
            if grp in counts:
                counts[grp] += 1
    return [counts[g] for g in COLORS]

def make_report():
    """
    Erstellt einen PDF-Report mit Deckblatt und einem Bar-Chart.
    """
    # Datum
    now = datetime.now(timezone.utc)
    today = now.date()
    date_str = today.isoformat()
    cutoff = today - timedelta(days=30)
    cutoff_str = cutoff.isoformat()
    pdf_file = os.path.join("status", f"status_report_{date_str}.pdf")

    # Dokument
    doc = SimpleDocTemplate(
        pdf_file,
        pagesize=A4,
        leftMargin=20*mm,
        rightMargin=20*mm,
        topMargin=20*mm,
        bottomMargin=20*mm
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CoverTitle', parent=styles['Title'], fontName='Helvetica-Bold', fontSize=20, spaceAfter=12))
    styles.add(ParagraphStyle(name='CoverInfo', parent=styles['Normal'], fontName='Helvetica', fontSize=12, spaceAfter=6))
    styles.add(ParagraphStyle(name='ChartTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=14, spaceAfter=6))

    elems = []

    # Deckblatt
    elems.append(Paragraph("Amazon Content Management - Activity Report", styles['CoverTitle']))
    elems.append(Paragraph(f"Erstellungsdatum: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['CoverInfo']))
    elems.append(Paragraph(f"Abgrenzungsdatum: {cutoff_str}", styles['CoverInfo']))
    elems.append(PageBreak())

    # Daten
    groups = list(COLORS.keys())
    values = read_snapshot(date_str)

    # Chart-Titel
    elems.append(Paragraph("Anzahl an eröffneten Phasen pro Produktgruppe (letzte 30 Tage)", styles['ChartTitle']))
    elems.append(Spacer(1, 6*mm))

    # Bar-Chart manuell zeichnen
    usable_width = A4[0] - 2*20*mm  # nutzbare Breite
    chart_height = 60*mm
    origin_x = 0
    origin_y = 15*mm

    # Berechne Balkenbreite und Abstände
    total_gap = usable_width * 0.1
    gap = total_gap / (len(groups) + 1)
    bar_width = (usable_width - total_gap) / len(groups)

    drawing = Drawing(usable_width, chart_height + origin_y + 2*mm)
    for idx, grp in enumerate(groups):
        val = values[idx]
        x = origin_x + gap * (idx+1) + bar_width * idx
        height = (val / max(values or [1])) * chart_height

        # Balken zeichnen
        bar = Rect(x, origin_y, bar_width, height,
                   fillColor=colors.HexColor(COLORS[grp]), strokeColor=None)
        drawing.add(bar)

        # Wert-Label oben
        drawing.add(String(x + bar_width/2, origin_y + height + 4,
                           str(val), fontName='Helvetica', fontSize=9, textAnchor='middle'))

        # Gruppen-Label unten
        drawing.add(String(x + bar_width/2, origin_y - 10,
                           grp, fontName='Helvetica', fontSize=8, textAnchor='middle'))

    elems.append(drawing)
        # --- Gestapeltes Balkendiagramm NA (Phasen 1–5 nach Mitarbeiter) ---
    from collections import defaultdict

    # 1) Farben der Mitarbeitenden festlegen
    EMP_COLORS = {
        "DM": "#223459",
        "EK": "#6A5AAA",
        "HI": "#B45082",
        "SM": "#F9767F",
        "JHU": "#FFB142",
        "LK": "#FFDE70",
    }

    # 2) Snapshot einlesen und auf NA filtern
    snap_path = os.path.join("status", f"status_snapshot_{date_str}.csv")
    data_rows = []
    with open(snap_path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["Produktgruppe"] == "NA":
                phase = int(row["Phase"])
                emp = (row["Mitarbeiter"] or "").strip()
                if emp:  # nur wenn Mitarbeiter gesetzt
                    data_rows.append((phase, emp))

    # 3) Pivot: Phase -> Mitarbeiter -> Count
    counts = defaultdict(lambda: defaultdict(int))
    emp_set = set()
    for phase, emp in data_rows:
        counts[phase][emp] += 1
        emp_set.add(emp)

    phases_sorted = [1, 2, 3, 4, 5]
    # nur Mitarbeiter aufnehmen, die in EMP_COLORS stehen (Rest optional ignorieren)
    emp_sorted = [e for e in EMP_COLORS if e in emp_set]

    # 4) Neues Blatt/Seite + Titel
    elems.append(PageBreak())
    elems.append(Paragraph("Mitarbeiterbasierte Phasenstatistik (NA, 30 Tage)", styles['ChartTitle']))
    elems.append(Spacer(1, 6*mm))

    # 5) Zeichnungsfläche definieren
    usable_width2 = A4[0] - 2*15*mm
    row_h = 8*mm
    gap_y = 5*mm
    left_axis_space = 18*mm
    origin_x2 = 0
    origin_y2 = 10*mm
    total_h = len(phases_sorted) * (row_h + gap_y) + origin_y2 + 10*mm

    d2 = Drawing(usable_width2, total_h)

    # Max pro Phase für Skalierung
    max_total = max((sum(counts[p][e] for e in emp_sorted) for p in phases_sorted), default=1)

    # 6) Gestapelte Balken pro Phase zeichnen (horizontal)
    for idx_p, phase in enumerate(phases_sorted):
        y = origin_y2 + (len(phases_sorted) - 1 - idx_p) * (row_h + gap_y)
        x = left_axis_space
        # Phasenlabel links
        d2.add(String(0, y + row_h / 2, f"Phase {phase}", fontName='Helvetica', fontSize=9, textAnchor='start'))

        run_w = 0
        for emp in emp_sorted:
            val = counts[phase][emp]
            if val == 0:
                continue
            seg_w = (val / max_total) * (usable_width2 - left_axis_space - 5*mm)
            rect = Rect(x + run_w, y, seg_w, row_h,
                        fillColor=colors.HexColor(EMP_COLORS[emp]), strokeColor=None)
            d2.add(rect)

            # Beschriftung nur wenn genug Platz
            if seg_w > 12:
                d2.add(String(x + run_w + seg_w / 2, y + row_h / 2,
                              str(val), fontName='Helvetica-Bold', fontSize=8,
                              textAnchor='middle', fillColor=colors.white))
            run_w += seg_w

    # 7) Legende
    legend_x = usable_width2 - 40*mm
    legend_y = total_h - 10*mm
    box_h = 5*mm
    for i, emp in enumerate(emp_sorted):
        yy = legend_y - i * (box_h + 2*mm)
        d2.add(Rect(legend_x, yy, box_h, box_h,
                    fillColor=colors.HexColor(EMP_COLORS[emp]), strokeColor=None))
        d2.add(String(legend_x + box_h + 3*mm, yy + box_h / 2,
                      emp, fontName='Helvetica', fontSize=8, textAnchor='start'))

    elems.append(d2)
    elems.append(Spacer(1, 6*mm))

    # Fußzeile
    elems.append(Spacer(1, 6*mm))
    elems.append(Paragraph(f"Report erstellt: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['Normal']))

    # PDF bauen
    doc.build(elems)
    print(f"✅ PDF Report erstellt: {pdf_file}")

if __name__ == "__main__":
    make_report()
