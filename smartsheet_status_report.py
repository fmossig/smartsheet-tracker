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
        next(reader)  # Header überspringen
        for row in reader:
            grp = row[0]
            if grp in counts:
                counts[grp] += 1
    return [counts[g] for g in COLORS]


def make_report():
    """
    Erstellt einen mehrseitigen PDF-Report mit Deckblatt und Bar-Chart.
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
    elems.append(Spacer(1, 4*mm))

    # Bar-Chart manuell zeichnen
    usable_width = (210 - 2*20) * mm  # A4 Breite minus Ränder
    chart_height = 80*mm
    origin_x = 0
    origin_y = 15*mm

    max_val = max(values) if values else 1
    num = len(groups)
    spacing = usable_width / (num * 1.5)
    bar_width = usable_width / (num * 1.2)

    drawing = Drawing(usable_width, chart_height + origin_y)
    for idx, grp in enumerate(groups):
        val = values[idx]
        x = origin_x + spacing/2 + idx * (bar_width + spacing)
        height = (val / max_val) * chart_height
        # Balken zeichnen
        bar = Rect(x, origin_y, bar_width, height,
                   fillColor=colors.HexColor(COLORS[grp]), strokeColor=None)
        drawing.add(bar)
        # Wert-Label über Balken
        drawing.add(String(x + bar_width/2, origin_y + height + 4,
                           str(val), fontName='Helvetica', fontSize=9, textAnchor='middle'))
        # Gruppen-Label unter Balken
        drawing.add(String(x + bar_width/2, origin_y - 10,
                           grp, fontName='Helvetica', fontSize=8, textAnchor='middle'))

    elems.append(drawing)

    # Fußzeile
    elems.append(Spacer(1, 6*mm))
    elems.append(Paragraph(f"Report erstellt: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['Normal']))

    # PDF bauen
    doc.build(elems)
    print(f"✅ PDF Report erstellt: {pdf_file}")


if __name__ == "__main__":
    make_report()
