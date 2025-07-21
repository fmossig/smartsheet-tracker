import csv, os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# Lade Summery‑CSV
date_str = datetime.utcnow().strftime("%Y-%m-%d")
summ_file = f"status/status_summary_{date_str}.csv"
rows = list(csv.reader(open(summ_file, encoding="utf-8")))
# Header + Values
_, *headers = rows[0]
_, *phase_events = rows[1]

# Map Produktgruppen in fester Reihenfolge
groups = ["NA","NF","NH","NM","NP","NT","NV"]
# Hier würden wir counts aus der Phase‑Events Tabelle ziehen,
# aber Summary nur hat Gesamt‑Events – wir müssen erst den Snapshot

# Stattdessen: Chart1 aus Snapshot aggregieren:
snap_file = f"status/status_snapshot_{date_str}.csv"
data = list(csv.reader(open(snap_file, encoding="utf-8")))
# data[0] = header
# Gruppiere Zählung je Produktgruppe
counts = {g:0 for g in groups}
for row in data[1:]:
    grp = row[0]
    if grp in counts:
        counts[grp]+=1

values = [counts[g] for g in groups]

# Farben pro Gruppe
colmap = {
    "NA": "#E63946","NF":"#457B9D","NH":"#2A9D8F",
    "NM":"#E9C46A","NP":"#F4A261","NT":"#9D4EDD","NV":"#00B4D8"
}

# PDF aufsetzen
out = f"status/status_report_{date_str}.pdf"
doc = SimpleDocTemplate(out, pagesize=A4)
styles = getSampleStyleSheet()
elems = []

# Titel
elems.append(Paragraph("PRODUKTGRUPPEN ÜBERSICHT", styles['Title']))
elems.append(Spacer(1,12))

# Chart-Titel
elems.append(Paragraph("Anzahl an eröffneten Phasen pro Produktgruppe", styles['Heading2']))
elems.append(Spacer(1,12))

# BarChart
drawing = Drawing(500,200)
chart = VerticalBarChart()
chart.x = 50; chart.y = 20
chart.height = 150; chart.width = 400
chart.data = [values]
chart.categoryAxis.categoryNames = groups
chart.valueAxis.valueMin = 0
chart.valueAxis.valueMax = max(values)*1.1 if values else 1
chart.valueAxis.valueStep = max(1,int(max(values)/10)+1)

# Farben
chart.bars.fillColor = None  # disable default
# assign colors per bar
for i,grp in enumerate(groups):
    chart.bars[i].fillColor = colors.HexColor(colmap[grp])

drawing.add(chart)
elems.append(drawing)

doc.build(elems)
print(f"✅ PDF Report geschrieben: {out}")
