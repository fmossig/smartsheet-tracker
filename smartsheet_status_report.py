import csv, os
from datetime import datetime, timedelta, timezone
from collections import defaultdict

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.graphics.shapes import Drawing, String, Rect
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm

# ---------- Farben ----------
GROUP_COLORS = {
    "NA": "#E63946",
    "NF": "#457B9D",
    "NH": "#2A9D8F",
    "NM": "#E9C46A",
    "NP": "#F4A261",
    "NT": "#9D4EDD",
    "NV": "#00B4D8"
}

EMP_COLORS = {
    "DM":  "#223459",
    "EK":  "#6A5AAA",
    "HI":  "#B45082",
    "SM":  "#F9767F",
    "JHU": "#FFB142",
    "LK":  "#FFDE70",
}

CODE_VERSION = "2025-07-22_17h30_banner"  # << zum Erkennen der aktuellen Version

# ---------- Datenhelpers ----------
def read_snapshot_counts(date_str):
    """Zählt alle Phase-Events pro Produktgruppe aus status_snapshot_<date>.csv."""
    snap_file = os.path.join("status", f"status_snapshot_{date_str}.csv")
    counts = {g: 0 for g in GROUP_COLORS}
    with open(snap_file, encoding="utf-8") as f:
        r = csv.reader(f)
        next(r)
        for row in r:
            grp = row[0]
            if grp in counts:
                counts[grp] += 1
    return [counts[g] for g in GROUP_COLORS]

def read_na_phase_employee(date_str):
    """dict: phase -> {emp: count} für NA, nur definierte EMP_COLORS."""
    snap_file = os.path.join("status", f"status_snapshot_{date_str}.csv")
    counts = defaultdict(lambda: defaultdict(int))
    with open(snap_file, encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            if row["Produktgruppe"] != "NA":
                continue
            phase = int(row["Phase"])
            emp = (row["Mitarbeiter"] or "").strip()
            if emp in EMP_COLORS:
                counts[phase][emp] += 1
    return counts

# ---------- Report ----------
def make_report():
    now = datetime.now(timezone.utc)
    today = now.date()
    date_str = today.isoformat()
    cutoff = today - timedelta(days=30)

    pdf_file = os.path.join("status", f"status_report_{date_str}.pdf")

    doc = SimpleDocTemplate(
        pdf_file,
        pagesize=A4,
        leftMargin=20*mm, rightMargin=20*mm,
        topMargin=20*mm, bottomMargin=20*mm
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CoverTitle', parent=styles['Title'],
                              fontName='Helvetica-Bold', fontSize=20, spaceAfter=12))
    styles.add(ParagraphStyle(name='CoverInfo', parent=styles['Normal'],
                              fontName='Helvetica', fontSize=12, spaceAfter=6))
    styles.add(ParagraphStyle(name='ChartTitle', parent=styles['Heading2'],
                              fontName='Helvetica-Bold', fontSize=14, spaceAfter=6))

    elems = []

    # ---- Deckblatt ----
    elems.append(Paragraph("Amazon Content Management - Activity Report", styles['CoverTitle']))
    elems.append(Paragraph(f"Erstellungsdatum: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['CoverInfo']))
    elems.append(Paragraph(f"Abgrenzungsdatum: {cutoff.isoformat()}", styles['CoverInfo']))
    elems.append(Paragraph(f"CODE_VERSION: {CODE_VERSION}", styles['CoverInfo']))
    elems.append(PageBreak())

    # ---- Chart 1: Produktgruppen ----
    groups = list(GROUP_COLORS.keys())
    values = read_snapshot_counts(date_str)

    elems.append(Paragraph("NA - Produktgruppen Daten (30 Tage)", styles['ChartTitle']))
    elems.append(Spacer(1, 6*mm))

    usable_width = A4[0] - 2*20*mm
    chart_height = 60*mm
    origin_y = 15*mm

    total_gap = usable_width * 0.1
    gap = total_gap / (len(groups) + 1)
    bar_width = (usable_width - total_gap) / len(groups)

    d1 = Drawing(usable_width, chart_height + origin_y + 2*mm)
    max_val = max(values) if values else 1

    for i, grp in enumerate(groups):
        val = values[i]
        x = gap * (i + 1) + bar_width * i
        h = (val / max_val) * chart_height

        d1.add(Rect(x, origin_y, bar_width, h,
                    fillColor=colors.HexColor(GROUP_COLORS[grp]), strokeColor=None))
        d1.add(String(x + bar_width/2, origin_y + h + 4, str(val),
                      fontName='Helvetica', fontSize=9, textAnchor='middle'))
        d1.add(String(x + bar_width/2, origin_y - 10, grp,
                      fontName='Helvetica', fontSize=8, textAnchor='middle'))

    elems.append(d1)

    # ---- Chart 2: Gestapelt NA ----
    counts = read_na_phase_employee(date_str)
    phases_sorted = [1, 2, 3, 4, 5]
    emp_sorted = [e for e in EMP_COLORS if any(counts[p][e] for p in phases_sorted)]

    # 1) Titel-Banner (volle Breite, NA-Farbe)
    banner_h = 12*mm
    banner_w = A4[0] - (doc.leftMargin + doc.rightMargin)
    banner = Drawing(banner_w, banner_h)
    banner.add(Rect(0, 0, banner_w, banner_h,
                    fillColor=colors.HexColor(GROUP_COLORS["NA"]),
                    strokeColor=None))
    banner.add(String(banner_w/2, banner_h/2,
                      "Mitarbeiterbasierte Phasenstatistik (NA, 30 Tage)",
                      fontName='Helvetica-Bold', fontSize=18,
                      textAnchor='middle', fillColor=colors.white))
    elems.append(PageBreak())
    elems.append(banner)
    elems.append(Spacer(1, 4*mm))

    # 2) Horizontale Legende
    legend_h  = 10*mm          # Gesamthöhe der Legendenzeile
    box_size  = 5*mm           # Größe der Quadrate
    font_sz   = 10             # Schriftgröße
    gap_item  = 14*mm          # horizontaler Abstand zwischen Einträgen

    legend_w  = banner_w
    leg = Drawing(legend_w, legend_h)

    y_center = legend_h / 2
    # Textbaseline sitzt etwas unter der Mitte -> optisch korrigieren:
    text_y = y_center - (font_sz * 0.35)  # Feinjustierung fürs Auge

    x_cursor = 0
    for emp in emp_sorted:
        # Box vertikal mittig
        leg.add(Rect(x_cursor,
                     y_center - box_size/2,
                     box_size, box_size,
                     fillColor=colors.HexColor(EMP_COLORS[emp]),
                     strokeColor=None))
        # Text daneben, ebenfalls optisch mittig
        leg.add(String(x_cursor + box_size + 2*mm,
                       text_y,
                       emp,
                       fontName='Helvetica',
                       fontSize=font_sz,
                       textAnchor='start'))
        x_cursor += box_size + 2*mm + gap_item

    elems.append(leg)
    elems.append(Spacer(1, 6*mm))


    # 3) Gestapeltes Chart (5% kleiner)
    shrink = 0.75
    chart_w   = (A4[0] - doc.leftMargin - doc.rightMargin) * 0.95
    left_ax   = 20*mm
    row_h     = 8*mm * shrink
    gap_y     = 4*mm * shrink
    origin_y2 = 10*mm

    # Höhe berechnen
    rows_drawn = sum(1 for p in phases_sorted if sum(counts[p][e] for e in emp_sorted) > 0)
    total_h = rows_drawn * (row_h + gap_y) + origin_y2 + 6*mm
    max_h   = A4[1] - doc.topMargin - doc.bottomMargin - 40*mm
    if total_h > max_h:
        scale = max_h / total_h
        row_h     *= scale
        gap_y     *= scale
        origin_y2 *= scale
        total_h    = max_h

    d2 = Drawing(chart_w, total_h)

    # Zeichnen
    y_index = 0
    for phase in phases_sorted:
        phase_total = sum(counts[phase][e] for e in emp_sorted)
        if phase_total == 0:
            continue

        y = origin_y2 + (rows_drawn - 1 - y_index) * (row_h + gap_y)
        y_index += 1
        x = left_ax

        d2.add(String(2*mm, y + row_h/2, f"Phase {phase}",
                      fontName='Helvetica', fontSize=8, textAnchor='start'))

        run_w   = 0
        avail_w = chart_w - left_ax - 5*mm
        for emp in emp_sorted:
            v = counts[phase][emp]
            if v == 0:
                continue
            seg_w = max((v / phase_total) * avail_w, 2)  # Mindestbreite
            rect = Rect(x + run_w, y, seg_w, row_h,
                        fillColor=colors.HexColor(EMP_COLORS[emp]), strokeColor=None)
            d2.add(rect)

            if seg_w > 14:
                d2.add(String(x + run_w + seg_w/2, y + row_h/2, str(v),
                              fontName='Helvetica-Bold', fontSize=7,
                              textAnchor='middle', fillColor=colors.white))
            run_w += seg_w

    elems.append(d2)
    elems.append(Spacer(1, 6*mm))

    # ---- Footer ----
    elems.append(Paragraph(f"Report erstellt: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['Normal']))

    doc.build(elems)
    print(f"✅ PDF Report erstellt: {pdf_file}")

if __name__ == "__main__":
    make_report()
