import os, csv, glob
from datetime import datetime, timedelta, timezone
from collections import defaultdict
import statistics

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.graphics.shapes import Drawing, String, Rect
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.utils import ImageReader

from reportlab.graphics.charts.piecharts import Pie

# Smartsheet
import smartsheet
from dotenv import load_dotenv

# --- PERIOD-OVERRIDES (Woche/Monat/Custom) ---
def _parse_ymd(s: str):
    return datetime.strptime(s, "%Y-%m-%d").date()

def get_period_override():
    """
    Liest optionale Umgebungsvariablen:
      REPORT_SINCE  (YYYY-MM-DD)
      REPORT_UNTIL  (YYYY-MM-DD)
      REPORT_OUTDIR (Pfad, default: "status")
      REPORT_LABEL  (z.B. "2025-W36" oder "2025-08")
    Gibt zurück: (since_date, until_date, out_dir, label) oder None
    """
    since = os.getenv("REPORT_SINCE")
    until = os.getenv("REPORT_UNTIL")
    if not (since and until):
        return None
    out_dir = os.getenv("REPORT_OUTDIR", "status")
    label   = os.getenv("REPORT_LABEL")  # optional
    return (_parse_ymd(since), _parse_ymd(until), out_dir, label)

# ---------- Pfade ----------
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
LOG_DIR    = os.path.join(BASE_DIR, "tracker_logs")  # Tracker-CSV Ordner

# ---------- Farben ----------
GROUP_COLORS = {
    "NA": "#E63946",
    "NF": "#457B9D",
    "NH": "#2A9D8F",
    "NM": "#E9C46A",
    "NP": "#F4A261",
    "NT": "#9D4EDD",
    "NV": "#00B4D8",
}

PHASE_COLORS = {
    1: "#1F77B4",
    2: "#FF7F0E",
    3: "#2CA02C",
    4: "#9467BD",
    5: "#D62728",
}

EMP_COLORS = {
    "DM":  "#223459",
    "EK":  "#6A5AAA",
    "HI":  "#B45082",
    "SM":  "#F9767F",
    "JHU": "#FFB142",
    "LK":  "#FFDE70",
}

CODE_VERSION = "2025-07-23_tracker"

# ---------- Smartsheet Spalten ----------
COL_ARTIKEL = "Artikel"
COL_LINK    = "Link"
COL_AMAZON  = "Amazon"
PHASE_DATE_COLS = ["Kontrolle", "BE am", "K am", "C am", "Reopen C2 am"]

# ---------- Sheet-IDs ----------
SHEET_IDS = {
    "NA": 6141179298008964,
    "NF": 615755411312516,
    "NH": 123340632051588,
    "NP": 3009924800925572,
    "NT": 2199739350077316,
    "NV": 8955413669040004,
    "NM": 4275419734822788,
}

# ---------- Footer ----------
def draw_footer(canvas, doc):
    amz_path = os.path.join(ASSETS_DIR, "amazon_logo.png")
    noc_path = os.path.join(ASSETS_DIR, "noctua_logo.png")

    logo_h   = 8*mm
    gap      = 4*mm
    padding  = 2*mm

    x0 = doc.leftMargin
    y0 = doc.bottomMargin - logo_h - padding

    def draw_if_exists(path, x, y, h):
        if not os.path.isfile(path):
            return 0
        img = ImageReader(path)
        w, h_img = img.getSize()
        scale = h / h_img
        w_pt  = w * scale
        canvas.drawImage(img, x, y, width=w_pt, height=h, mask='auto')
        return w_pt

    w_amz = draw_if_exists(amz_path, x0, y0, logo_h)
    draw_if_exists(noc_path, x0 + w_amz + gap, y0, logo_h)

# ---------- Parser ----------
def parse_date_fuzzy(s):
    if not s:
        return None
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        pass
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue
    return None

# ---------- Tracker-Log Loader & Aggregationen ----------
def load_log_rows(log_dir, since_date, until_date=None):
    rows = []
    for path in glob.glob(os.path.join(log_dir, "date_changes_log_*.csv")):
        with open(path, encoding="utf-8") as f:
            r = csv.DictReader(f)
            for row in r:
                d = parse_date_fuzzy(row.get("Datum", ""))
                if not d:
                    continue
                if d < since_date:
                    continue
                if until_date and d > until_date:
                    continue
                try:
                    phase_no = int(row.get("Phase", "0") or 0)
                except ValueError:
                    phase_no = 0
                rows.append({
                    "ts":    row.get("Änderung am", "").strip(),
                    "group": row.get("Produktgruppe", "").strip(),
                    "land":  row.get("Land/Marketplace", "").strip(),
                    "phase": phase_no,
                    "field": row.get("Feld", "").strip(),
                    "date":  d,
                    "user":  row.get("Mitarbeiter", "").strip(),
                })
    return rows

def aggregate_group_counts(log_rows):
    counts = defaultdict(int)
    for r in log_rows:
        counts[r["group"]] += 1
    return counts

def aggregate_phase_emp(log_rows, group_code):
    res = defaultdict(lambda: defaultdict(int))
    for r in log_rows:
        if r["group"] != group_code:
            continue
        if r["phase"] and r["user"]:
            res[r["phase"]][r["user"]] += 1
    return res

def aggregate_phase_dist(log_rows, group_code):
    dist = {p: 0 for p in range(1, 6)}
    for r in log_rows:
        if r["group"] == group_code and r["phase"] in dist:
            dist[r["phase"]] += 1
    return dist

# ---------- Pie ----------
def build_phase_pie(dist_dict, diam_mm):
    phases = [1,2,3,4,5]
    data   = [dist_dict.get(p, 0) for p in phases]

    d = Drawing(diam_mm*mm, diam_mm*mm)
    pie = Pie()
    pie.width  = pie.height = diam_mm*mm
    pie.x = pie.y = 0
    pie.data   = data
    pie.labels = []
    pie.sideLabels  = 0
    pie.simpleLabels = 1

    pie.slices.strokeColor = colors.white
    pie.slices.strokeWidth = 0.6

    for i, p in enumerate(phases):
        pie.slices[i].fillColor = colors.HexColor(PHASE_COLORS[p])

    d.add(pie)
    return d

# ---------- KPIs (Smartsheet) ----------
def calc_metrics_for_group(client, group_code, cutoff_date):
    sheet = client.Sheets.get_sheet(SHEET_IDS[group_code])
    col_map = {c.title: c.id for c in sheet.columns}

    artikel_count = 0
    mp_count = 0
    touched_rows = set()

    for row in sheet.rows:
        artikel_val = ""
        link_val = ""
        amazon_val = ""
        has_recent = False

        for cell in row.cells:
            if cell.column_id == col_map.get(COL_ARTIKEL):
                artikel_val = (cell.display_value or "").strip()
            elif cell.column_id == col_map.get(COL_LINK):
                link_val = (cell.display_value or "").strip()
            elif cell.column_id == col_map.get(COL_AMAZON):
                amazon_val = (cell.display_value or "").strip()

        for col_name in PHASE_DATE_COLS:
            cid = col_map.get(col_name)
            if not cid:
                continue
            cell = next((c for c in row.cells if c.column_id == cid), None)
            if cell and cell.value:
                try:
                    dt = datetime.fromisoformat(cell.value).date()
                    if dt >= cutoff_date:
                        has_recent = True
                except Exception:
                    pass

        if artikel_val and not link_val:
            artikel_count += 1
        if artikel_val and link_val and amazon_val:
            mp_count += 1
        if has_recent:
            touched_rows.add(row.id)

    bearbeitet = len(touched_rows)
    pct = (bearbeitet / mp_count * 100) if mp_count else 0.0
    return artikel_count, mp_count, bearbeitet, pct

# ---------- Länder-Statistik (Smartsheet) ----------
def country_age_stats_for_group(client, group_code, today_date):
    sheet = client.Sheets.get_sheet(SHEET_IDS[group_code])
    col_map = {c.title: c.id for c in sheet.columns}

    country_days = defaultdict(list)

    for row in sheet.rows:
        artikel_val = ""
        link_val = ""
        amazon_val = ""
        for cell in row.cells:
            if cell.column_id == col_map.get(COL_ARTIKEL):
                artikel_val = (cell.display_value or "").strip()
            elif cell.column_id == col_map.get(COL_LINK):
                link_val = (cell.display_value or "").strip()
            elif cell.column_id == col_map.get(COL_AMAZON):
                amazon_val = (cell.display_value or "").strip()

        if not (artikel_val and link_val and amazon_val):
            continue

        last_dt = None
        for col_name in PHASE_DATE_COLS:
            cid = col_map.get(col_name)
            if not cid:
                continue
            cell = next((c for c in row.cells if c.column_id == cid), None)
            if cell and cell.value:
                try:
                    dt = datetime.fromisoformat(cell.value).date()
                    if (last_dt is None) or (dt > last_dt):
                        last_dt = dt
                except Exception:
                    pass

        if last_dt:
            age_days = (today_date - last_dt).days
            country_days[amazon_val].append(age_days)

    stats = []
    for land, lst in country_days.items():
        if not lst:
            continue
        stats.append({
            "land": land,
            "avg": statistics.mean(lst),
            "median": statistics.median(lst),
            "max": max(lst),
            "count": len(lst)
        })
    return stats

# ---------- Header Chip ----------
def build_group_header(grp_code, grp_color_hex, period_text="Zeitraum: 30 Tage"):
    chip_font   = 'Helvetica-Bold'
    chip_fs     = 18
    chip_pad_x  = 4*mm
    chip_pad_y  = 2*mm

    line1_font  = 'Helvetica-Bold'
    line1_fs    = 14
    line2_font  = 'Helvetica'
    line2_fs    = 9

    gap_between = 6*mm

    CHIP_TEXT_UP = 1*mm
    TEXT_LEFT    = 2*mm

    chip_text_w = stringWidth(grp_code, chip_font, chip_fs)
    chip_w = chip_text_w + 2*chip_pad_x
    chip_h = chip_fs*1.2 + 2*chip_pad_y

    line1_w = stringWidth("Daten & Kennzahlen", line1_font, line1_fs)
    line2_w = stringWidth(period_text,          line2_font,  line2_fs)
    text_block_w = max(line1_w, line2_w)

    total_w = chip_w + gap_between + text_block_w
    total_h = chip_h

    d = Drawing(total_w, total_h)

    d.add(Rect(0, 0, chip_w, chip_h,
               fillColor=colors.HexColor(grp_color_hex),
               strokeColor=None))
    d.add(String(chip_pad_x,
                 chip_pad_y + chip_fs*0.1 + CHIP_TEXT_UP,
                 grp_code,
                 fontName=chip_font, fontSize=chip_fs,
                 fillColor=colors.white, textAnchor='start'))

    text_x   = chip_w + gap_between - TEXT_LEFT
    center_y = total_h/2.0
    line1_y  = center_y + line1_fs*0.35
    line2_y  = center_y - line2_fs*1.1

    d.add(String(text_x, line1_y,
                 "Daten & Kennzahlen",
                 fontName=line1_font, fontSize=line1_fs,
                 fillColor=colors.black, textAnchor='start'))
    d.add(String(text_x, line2_y,
                 period_text,
                 fontName=line2_font, fontSize=line2_fs,
                 fillColor=colors.black, textAnchor='start'))

    return d, total_h

# ---------- Länder-Rank-Tabellen ----------
def build_country_rank_tables(stats,
                              width_total,
                              gap_between_tables=12*mm,
                              shift_right=6*mm,
                              banner_shrink=0.8,
                              banner_height=9*mm):
    from reportlab.platypus import Spacer, Table, TableStyle
    longest = sorted(stats, key=lambda x: x["avg"], reverse=True)[:5]
    active  = sorted(stats, key=lambda x: x["avg"])[:5]

    def rows(lst):
        out = []
        for i, item in enumerate(lst, start=1):
            out.append([i, item["land"], f"~{int(round(item['avg']))}d"])
        return out

    left_rows  = rows(longest)
    right_rows = rows(active)

    def make_banner(txt, hexcolor, w, h=banner_height):
        d = Drawing(w, h)
        d.add(Rect(0, 0, w, h, fillColor=colors.HexColor(hexcolor), strokeColor=None))
        d.add(String(w/2, h/2 - 1.5*mm, txt,
                     fontName="Helvetica-Bold", fontSize=10,
                     textAnchor="middle", fillColor=colors.white))
        return d

    col_w = (width_total - gap_between_tables) / 2.0
    bn_w  = col_w * banner_shrink

    left_banner  = make_banner("Inaktivste Länder", "#8B0000", bn_w)
    right_banner = make_banner("Aktivste Länder",   "#2E8B57", bn_w)

    def center_draw(drawing, col_width):
        t = Table([[drawing]], colWidths=[col_width], hAlign="CENTER")
        t.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING",(0,0), (-1,-1), 0),
            ("TOPPADDING",  (0,0), (-1,-1), 0),
            ("BOTTOMPADDING",(0,0), (-1,-1), 0),
            ("BOX", (0,0), (-1,-1), 0, colors.white),
            ("INNERGRID", (0,0), (-1,-1), 0, colors.white),
        ]))
        return t

    left_banner_tbl  = center_draw(left_banner,  col_w)
    right_banner_tbl = center_draw(right_banner, col_w)

    header    = ["#", "Land", "Ø Alter"]
    base_fs   = 8
    second_fs = base_fs + 1
    first_fs  = base_fs + 2

    left_tbl  = Table([header] + left_rows,  colWidths=[8*mm, 28*mm, 22*mm], hAlign="CENTER")
    right_tbl = Table([header] + right_rows, colWidths=[8*mm, 28*mm, 22*mm], hAlign="CENTER")

    base_style = [
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), base_fs),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EEEEEE")),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("ALIGN", (0,1), (0,-1), "CENTER"),
        ("ALIGN", (2,1), (2,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("BOX", (0,0), (-1,-1), 0, colors.white),
        ("INNERGRID", (0,0), (-1,-1), 0, colors.white),
        ("LEFTPADDING",  (0,0), (-1,-1), 2),
        ("RIGHTPADDING", (0,0), (-1,-1), 2),
        ("TOPPADDING",   (0,0), (-1,-1), 2),
        ("BOTTOMPADDING",(0,0), (-1,-1), 2),
    ]
    left_tbl.setStyle(TableStyle(base_style))
    right_tbl.setStyle(TableStyle(base_style))

    def style_ranks(tbl):
        if len(tbl._cellvalues) > 1:
            tbl.setStyle(TableStyle([
                ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
                ("FONTSIZE", (0,1), (-1,1), first_fs),
            ]))
        if len(tbl._cellvalues) > 2:
            tbl.setStyle(TableStyle([
                ("FONTNAME", (0,2), (-1,2), "Helvetica-Bold"),
                ("FONTSIZE", (0,2), (-1,2), second_fs),
            ]))
        if len(tbl._cellvalues) > 3:
            tbl.setStyle(TableStyle([
                ("FONTNAME", (0,3), (-1,3), "Helvetica-Bold"),
                ("FONTSIZE", (0,3), (-1,3), base_fs),
            ]))
    style_ranks(left_tbl)
    style_ranks(right_tbl)

    block_spacer = Spacer(1, 2*mm)
    left_block  = [left_banner_tbl,  block_spacer, left_tbl]
    right_block = [right_banner_tbl, block_spacer, right_tbl]

    outer = Table([[left_block, "", right_block]],
                  colWidths=[col_w, gap_between_tables, col_w],
                  hAlign="CENTER",
                  style=TableStyle([
                      ("VALIGN", (0,0), (-1,-1), "TOP"),
                      ("LEFTPADDING", (0,0), (0,-1), shift_right),
                      ("LEFTPADDING", (2,0), (2,-1), shift_right),
                      ("TOPPADDING",  (0,0), (-1,-1), 0),
                      ("BOTTOMPADDING",(0,0), (-1,-1), 0),
                      ("BOX", (0,0), (-1,-1), 0, colors.white),
                      ("INNERGRID", (0,0), (-1,-1), 0, colors.white),
                  ]))

    return [outer]

# ---------- Report ----------
def make_report():
    load_dotenv()
    token = os.getenv("SMARTSHEET_TOKEN")
    client = smartsheet.Smartsheet(token)

    now = datetime.now(timezone.utc)
    today = now.date()

    # Zeitraum aus ENV übernehmen (falls gesetzt), sonst 30 Tage Standard
    override = get_period_override()
    if override:
        since_date, until_date, out_dir, label = override
        cutoff = since_date
        today  = until_date
        date_str = label or f"{since_date.isoformat()}_{until_date.isoformat()}"
        pdf_dir = out_dir
    else:
        cutoff = today - timedelta(days=30)
        date_str = today.isoformat()
        pdf_dir = "status"

    os.makedirs(pdf_dir, exist_ok=True)
    period_text = (
        f"Zeitraum: {cutoff:%d.%m.%Y} – {today:%d.%m.%Y}"
        if override else
        "Zeitraum: letzte 30 Tage"
    )



    # --- Tracker-Logs laden (30 Tage) ---
    log_rows = load_log_rows(LOG_DIR, since_date=cutoff, until_date=today)

    pdf_file = os.path.join(pdf_dir, f"status_report_{date_str}.pdf")



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
    elems.append(Paragraph(period_text, styles['CoverInfo']))
    elems.append(Paragraph(f"CODE_VERSION: {CODE_VERSION}", styles['CoverInfo']))
    elems.append(PageBreak())

    # ----- Seite: Produktgruppen Balken + Pies (aus Tracker) -----
    groups = list(GROUP_COLORS.keys())

    group_counts = aggregate_group_counts(log_rows)
    values = [group_counts.get(g, 0) for g in groups]

    elems.append(Paragraph("Produktgruppen Daten", styles['CoverTitle']))
    elems.append(Paragraph(period_text, styles['CoverInfo']))
    elems.append(Spacer(1, 6*mm))
    elems.append(Paragraph("Anzahl an eröffneten Phasen pro Produktgruppe", styles['ChartTitle']))
    elems.append(Spacer(1, 4*mm))

    usable_width = A4[0] - doc.leftMargin - doc.rightMargin
    chart_height = 55*mm
    origin_y = 25*mm
    total_gap = usable_width * 0.1
    gap = total_gap / (len(groups) + 1)
    bar_width = (usable_width - total_gap) / len(groups)

    d1 = Drawing(usable_width, chart_height + origin_y + 2*mm)
    max_val = max(values) if values else 1
    bar_x_centers = []

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

        bar_x_centers.append(x + bar_width/2)

    elems.append(d1)

    # --- Pies auf grauem Banner ---
    PIE_BANNER_COLOR = colors.HexColor("#F2F2F2")
    PIE_BANNER_H_MM  = 25
    pie_diam_mm = 20

    banner_h = PIE_BANNER_H_MM * mm
    pies_banner = Drawing(usable_width, banner_h)
    pies_banner.add(Rect(0, 0, usable_width, banner_h,
                         fillColor=PIE_BANNER_COLOR, strokeColor=None))

    pie_w = pie_diam_mm*mm
    pie_y = (banner_h - pie_w) / 2.0

    for idx, grp in enumerate(groups):
        dist = aggregate_phase_dist(log_rows, grp)
        pie_draw = build_phase_pie(dist, pie_diam_mm)
        pie_x = bar_x_centers[idx] - pie_w/2
        pie_draw.translate(pie_x, pie_y)
        pies_banner.add(pie_draw)

    elems.append(pies_banner)

    # Pie-Legende
    elems.append(Spacer(1, 3*mm))
    legend_w = usable_width
    legend_h = 12*mm
    leg = Drawing(legend_w, legend_h)
    box = 5*mm
    font_sz = 8
    gap_item = 18*mm
    items = [1,2,3,4,5]
    item_w = box + 2*mm + gap_item
    total_w = len(items)*item_w - gap_item
    x_cursor = (legend_w - total_w)/2
    y_center = legend_h/2
    for p in items:
        leg.add(Rect(x_cursor, y_center - box/2, box, box,
                     fillColor=colors.HexColor(PHASE_COLORS[p]), strokeColor=None))
        leg.add(String(x_cursor + box + 2*mm, y_center - 2,
                       f"Phase {p}", fontName='Helvetica', fontSize=font_sz,
                       textAnchor='start'))
        x_cursor += item_w
    elems.append(leg)

    # ---- pro Gruppe eigenes Board ----
    phases_sorted = [1, 2, 3, 4, 5]
    legend_items = list(EMP_COLORS.keys())

    banner_w_grp = usable_width
    legend_h_emp = 10*mm
    box_size = 5*mm
    font_sz_emp  = 10
    gap_item_emp = 14*mm

    shrink    = 0.75
    chart_w   = usable_width * 0.45
    left_ax   = 15*mm
    row_h     = 8*mm * shrink
    gap_y     = 4*mm * shrink
    origin_y2 = 10*mm

    for grp in groups:
        elems.append(PageBreak())

        hdr_draw, _ = build_group_header(grp, GROUP_COLORS[grp], period_text)
        elems.append(hdr_draw)
        elems.append(Spacer(1, 4*mm))

        # Grauer Banner für Emp-Legende
        leg_width  = banner_w_grp
        leg_height = 14*mm
        leg_draw = Drawing(leg_width, leg_height)
        leg_draw.add(Rect(0,0,leg_width,leg_height,
                          fillColor=colors.HexColor("#F2F2F2"), strokeColor=None))

        y_center = leg_height/2
        text_y   = y_center - (font_sz_emp * 0.35)
        item_w   = box_size + 2*mm + gap_item_emp
        total_w  = len(legend_items)*item_w - gap_item_emp
        x_cursor = (leg_width - total_w) / 2.0

        for emp in legend_items:
            leg_draw.add(Rect(x_cursor,
                              y_center - box_size/2,
                              box_size, box_size,
                              fillColor=colors.HexColor(EMP_COLORS[emp]),
                              strokeColor=None))
            leg_draw.add(String(x_cursor + box_size + 2*mm,
                                text_y,
                                emp,
                                fontName='Helvetica',
                                fontSize=font_sz_emp,
                                textAnchor='start'))
            x_cursor += item_w

        elems.append(leg_draw)
        elems.append(Spacer(1, 2*mm))

        # Gestapeltes Chart (Tracker-Daten)
        counts = aggregate_phase_emp(log_rows, grp)

        rows_drawn = len(phases_sorted)
        total_h = rows_drawn * (row_h + gap_y) + origin_y2 + 6*mm
        max_h   = A4[1] - doc.topMargin - doc.bottomMargin - 40*mm
        if total_h > max_h:
            scale = max_h / total_h
            row_h     *= scale
            gap_y     *= scale
            origin_y2 *= scale
            total_h    = max_h

        d2 = Drawing(chart_w, total_h)

        global_max = 0
        for p in phases_sorted:
            for emp in legend_items:
                global_max = max(global_max, counts[p][emp])

        avail_w     = chart_w - left_ax - 5*mm
        px_per_unit = avail_w / (global_max if global_max else 1)

        y_index = 0
        for phase in phases_sorted:
            y = origin_y2 + (rows_drawn - 1 - y_index) * (row_h + gap_y)
            y_index += 1
            x = left_ax

            d2.add(String(2*mm, y + row_h/2, f"Phase {phase}",
                          fontName='Helvetica', fontSize=8, textAnchor='start'))

            run_w = 0
            phase_total = sum(counts[phase][e] for e in legend_items)
            if phase_total == 0:
                d2.add(Rect(x, y, 1, row_h, fillColor=None, strokeColor=colors.grey))
                continue

            for emp in legend_items:
                v = counts[phase][emp]
                if v == 0:
                    continue
                seg_w = v * px_per_unit
                rect = Rect(x + run_w, y, seg_w, row_h,
                            fillColor=colors.HexColor(EMP_COLORS[emp]),
                            strokeColor=None)
                d2.add(rect)
                if seg_w > 14:
                    d2.add(String(x + run_w + seg_w/2, y + row_h/2, str(v),
                                  fontName='Helvetica-Bold', fontSize=9,
                                  textAnchor='middle', fillColor=colors.white))
                run_w += seg_w

        elems.append(d2)

        # KPIs (Smartsheet)
        artikel, mp, bearbeitet, pct = calc_metrics_for_group(client, grp, cutoff)
        elems.append(Spacer(1, 3*mm))

        red        = colors.HexColor(GROUP_COLORS[grp])
        box_h      = 18*mm
        spacing    = 10*mm
        font_val   = 14
        font_lab   = 7

        kpis = [
            ("Anzahl Artikel",       artikel),
            ("Marktplatzartikel",    mp),
            ("Bearbeitete Artikel",  bearbeitet),
            ("% bearbeitet",         f"{pct:.1f}%"),
        ]

        usable_full   = usable_width
        n             = len(kpis)
        total_spacing = spacing * (n - 1)
        box_w         = (usable_full * 0.88 - total_spacing) / n
        if box_w < 30*mm:
            box_w = 30*mm
        total_w = n * box_w + total_spacing
        start_x = (usable_full - total_w) / 2.0

        kpi_draw = Drawing(usable_full, box_h)
        for i, (label, value) in enumerate(kpis):
            x = start_x + i * (box_w + spacing)
            kpi_draw.add(Rect(x, 0, box_w, box_h,
                              fillColor=red, strokeColor=None))
            kpi_draw.add(String(x + box_w/2, box_h * 0.62,
                                str(value),
                                fontName='Helvetica-Bold', fontSize=font_val,
                                textAnchor='middle', fillColor=colors.white))
            kpi_draw.add(String(x + box_w/2, box_h * 0.28,
                                label,
                                fontName='Helvetica-Bold', fontSize=font_lab,
                                textAnchor='middle', fillColor=colors.white))

        elems.append(kpi_draw)
        elems.append(Spacer(1, 6*mm))

        # Länder-Rankings
        stats = country_age_stats_for_group(client, grp, today)
        if stats:
            elems.append(Spacer(1, 6*mm))
            elems.extend(build_country_rank_tables(stats, usable_full))
            elems.append(Spacer(1, 6*mm))

    # Footer
    elems.append(Paragraph(f"Report erstellt: {now.strftime('%Y-%m-%d %H:%M UTC')}", styles['Normal']))

    doc.build(elems, onFirstPage=draw_footer, onLaterPages=draw_footer)
    print(f"✅ PDF Report erstellt: {pdf_file}")

if __name__ == "__main__":
    make_report()
