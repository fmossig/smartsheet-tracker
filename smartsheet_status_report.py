import csv, os
from datetime import datetime, timedelta, timezone
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
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
    snap_file = f"status/status_snapshot_{date_str}.csv"
    with open(snap_file, encoding="utf-8") as f:
        data = list(csv.reader(f))[1:]
    counts = {g: 0 for g in COLORS}
    for grp, *_ in data:
        if grp in counts:
            counts[grp] += 1
    return [counts[g] for g in COLORS]


def make_report():
    # Setup
    now = datetime.now(timezone.utc)
    today = now.date()
    date_str = today.isoformat()
    cutoff = today - timedelta(days=30)
    cutoff_str = cutoff.isoformat()
    pdf_file = f"status/status_report_{date_str}.pdf"

    # PDF-Dokument
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

    # Daten laden
    groups = list(COLORS.keys())
    values = read_snapshot(date_str)

    # Chart Titel
    elems.append(Paragraph("Anzahl an eröffneten Phasen pro Produktgruppe (letzte 30 Tage)", styles['ChartTitle']))
    elems.append(Spacer(1, 4*mm))

    # Chart zeichnen
    drawing = Drawing(180*mm, 100*mm)
    chart = VerticalBarChart()
    chart.x = 15*mm
    chart.y = 15*mm
    chart.height = 80*mm
    chart.width = 150*mm
    chart.data = [values]
    chart.categoryAxis.categoryNames = groups
    chart.categoryAxis.labels.boxAnchor = 'n'

    # Achsen und Gitterlinien
    chart.valueAxis.valueMin = 0
    max_val = max(values) if values else 1
    chart.valueAxis.valueMax = max_val * 1.1
    chart.valueAxis.valueStep = max(1, int(max_val/10) or 1)
    chart.valueAxis.gridStrokeColor = colors.lightgrey

    # Balkenbreite und Abstand
    chart.barWidth = chart.width / (len(values) * 1.5)
    chart.groupSpacing = chart.barWidth / 2

    # Farben pro Balken und keine Umrandung
for idx, bar in enumerate(chart.bars):
    bar.fillColor = colors.HexColor(COLORS[groups[idx]])
    bar.strokeColor = None

# Labels über den Balken
    for idx, val in enumerate(values):
        x = chart.x + chart.groupSpacing + idx * (chart.barWidth + chart.groupSpacing) + chart.barWidth/2
        y = chart.y + (val / chart.valueAxis.valueMax) * chart.height + 6
        label = String(x, y, str(val), fontName='Helvetica', fontSize=9, textAnchor='middle')
        drawing.add(label)

    drawing.add(chart)
    elems.append(drawing)

    # Fußzeile
    elems.append(Spacer(1, 6*mm))
    elems.append(Paragraph(f"Report erstellt: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['Normal']))

    # PDF bauen
    doc.build(elems)
    print(f"✅ PDF Report erstellt: {pdf_file}")

if __name__ == "__main__":
    make_report()
