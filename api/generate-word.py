"""Vercel Python serverless function — HCT-COHS KPI Word Report Generator.
Fetches live data from Smartsheet API and generates downloadable .docx files.
"""

import os, io, json, zipfile
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET

# ── Namespaces ──
W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NAVY  = '1C2340'
BRAND = '180E3F'
GREY  = '595959'
WHITE = 'FFFFFF'

# ── Regions (same as PPT generator) ──
REGIONS = {
    'AD Al Ain':       {'sheets': ['AAF','AAZ'], 'short': ['Falaj Hazza','Zakhir'],        'subtitle': 'Al Ain Falaj Hazza & Al Ain Zakhir'},
    'Abu Dhabi':       {'sheets': ['ADA','ADB'], 'short': ['Baniyas A','Baniyas B'],       'subtitle': 'Abu Dhabi Baniyas A & Abu Dhabi Baniyas B'},
    'AD Remote':       {'sheets': ['ADH','MZY'], 'short': ['Al Dhanna','Madinat Zayed'],   'subtitle': 'Al Dhanna Ruwais & Al Dhafra Madinat Zayed City'},
    'Dubai':           {'sheets': ['DMC','DBN'], 'short': ['Academic City','Al Nahda'],     'subtitle': 'Dubai Academic City & Dubai Al Nahda'},
    'Fujairah':        {'sheets': ['FJF','FJH'], 'short': ['Faseel','Hulaifat'],            'subtitle': 'Fujairah Faseel & Fujairah Hulaifat'},
    'Sharjah':         {'sheets': ['SJA','SJB'], 'short': ['Campus A','Campus B'],          'subtitle': 'Sharjah Campus A & Sharjah Campus B'},
    'Ras Al Khaimah':  {'sheets': ['RKA','RKB'], 'short': ['Campus A','Campus B'],          'subtitle': 'RAK Campus A & RAK Campus B'},
}

# ── KPI structure ──
PILLAR_KPIS = [
    {'pillar': 'Leadership, Accountability & Engagement', 'weight': 0.20, 'rows': [2,3,4,5,6]},
    {'pillar': 'Risk Management & Planning',              'weight': 0.20, 'rows': [7,8,9]},
    {'pillar': 'Training & Awareness',                    'weight': 0.10, 'rows': [10,11]},
    {'pillar': 'OCP & Emergency Preparedness',            'weight': 0.25, 'rows': [12,13,14,15]},
    {'pillar': 'Performance Evaluation & Improvement',    'weight': 0.25, 'rows': [16,17,18,19]},
]

KPI_WEIGHTS = {
    2: 0.30, 3: 0.10, 4: 0.10, 5: 0.25, 6: 0.25,
    7: 0.30, 8: 0.50, 9: 0.20,
    10: 0.50, 11: 0.50,
    12: 0.40, 13: 0.40, 14: 0.10, 15: 0.10,
    16: 0.30, 17: 0.30, 18: 0.20, 19: 0.20,
}

KPI_NAMES = {
    2: 'HS KPI Report Submission',
    3: 'Total Hours of Training',
    4: 'External Authority Compliance',
    5: 'HS Committee Meeting',
    6: 'Hazard Identification',
    7: 'Risk Assessment Closed',
    8: 'Risk Assessment Validated',
    9: 'Safe Working Procedure',
    10: 'Planned Training',
    11: 'Training Hours Delivered',
    12: 'Operational Control Procedures',
    13: 'Emergency Drills',
    14: 'Permit to Work',
    15: 'Onsite Safety Induction',
    16: 'Scheduled EHS Inspection',
    17: 'Findings Closed On Time',
    18: 'Investigation Completed on Time',
    19: 'Incident Notification on Time',
}

# ── Smartsheet sources (same as PPT) ──
SYNC_SOURCES = [
    {'key': 'v2_hs_kpi_report', 'reportId': '4811266391494532', 'campusCol': 'Campuses', 'monthCol': 'Primary', 'valueCol': 'Submitted', 'kpi_row': 2},
    {'key': 'v2_external_compliance', 'sheetId': '4198632256393092', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'Applicable Compliance', 'actualCol': 'Actual Compliance', 'kpi_row': 4},
    {'key': 'v2_hs_committee', 'sheetId': '435993944477572', 'campusCol': 'Committee', 'monthCol': 'Reporting Month', 'plannedCol': 'Meeting Planned', 'actualCol': 'Meeting Conducted', 'kpi_row': 5},
    {'key': 'v2_hazard_id', 'sheetId': '7323092115214212', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Controls Identified', 'actualCol': 'Implemented Controls', 'kpi_row': 7},
    {'key': 'v2_risk_closed', 'sheetId': '7323092115214212', 'campusCol': 'Campus', 'monthCol': 'Primary', 'plannedCol': 'Total Risk Assessments Registered', 'actualCol': 'Risk Assessment Closed', 'kpi_row': 8},
    {'key': 'v2_risk_validated', 'sheetId': '7323092115214212', 'campusCol': 'Campus', 'monthCol': 'Primary', 'plannedCol': 'Total Assessments Register', 'actualCol': 'RA Validated and Signed Off', 'kpi_row': 9},
    {'key': 'v2_planned_training', 'sheetId': '8549734774951812', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned (Yes/No)', 'actualCol': 'Planned (Yes/No)', 'kpi_row': 10, 'yesNoCount': True},
    {'key': 'v2_safe_working', 'sheetId': '1693592581001092', 'campusCol': 'Campus', 'monthCol': 'Primary', 'plannedCol': 'No. of SOPs Verified', 'actualCol': 'No. of SOPs Implemented', 'kpi_row': 12},
    {'key': 'v2_drills', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned Drill? (Yes/No)', 'actualCol': 'Are there any submission?', 'kpi_row': 13, 'yesNoCount': True},
    {'key': 'v2_permit_to_work', 'sheetId': '5899016251330436', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'No. of PTWs Issued', 'actualCol': 'Total Work Registered', 'kpi_row': 14},
    {'key': 'v2_onsite_induction', 'sheetId': '5899016251330436', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': "No. of New Contractors (Individuals)", 'actualCol': 'Contractors Inducted in the Reporting Month', 'kpi_row': 15},
    {'key': 'v2_ehs_inspection', 'sheetId': '4947401822392196', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of EHS Inspections Planned', 'actualCol': 'No. of EHS Inspections Completed', 'kpi_row': 16},
    {'key': 'v2_findings_on_time', 'sheetId': '4947401822392196', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of Findings in Reporting Month', 'actualCol': 'No. of Findings Due', 'kpi_row': 17},
    {'key': 'v2_investigation_on_time', 'reportId': '6831846506581892', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident Investigated', 'actualCol': 'Investigation Completed on Time', 'kpi_row': 18},
    {'key': 'notification', 'reportId': '1199821531598724', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident', 'actualCol': 'Notification Submitted on Time', 'kpi_row': 19},
]

WASTE_SOURCE = {'sheetId': '8150747345538948', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month'}
WASTE_TABLE_COLS = ['General Waste', 'Food Waste', 'Paper Waste', 'Aluminum',
                    'PET Bottle', 'Paper Cup/Carton', 'Single Use Plastic',
                    'Tissue', 'Scrap Metal', 'E-waste', 'Hazardous']
RECYCLABLE_COLS = ['Food Waste', 'Paper Waste', 'Aluminum', 'PET Bottle',
                   'Paper Cup/Carton', 'Single Use Plastic', 'Tissue',
                   'Scrap Metal', 'E-waste']

MONTH_NAMES = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']


# ── Smartsheet API ──

def _ss_fetch(endpoint, token):
    url = f'https://api.smartsheet.com/2.0/{endpoint}'
    req = Request(url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'})
    with urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())

def fetch_sheet_rows(sheet_id, token):
    data = _ss_fetch(f'sheets/{sheet_id}?pageSize=500', token)
    if not data.get('rows'): return []
    col_map = {c['id']: c['title'] for c in data.get('columns', [])}
    rows = []
    for row in data['rows']:
        rec = {}
        for cell in row.get('cells', []):
            title = col_map.get(cell.get('columnId'))
            if title:
                rec[title] = cell.get('displayValue') or cell.get('value') or ''
        if rec: rows.append(rec)
    return rows

def fetch_report_rows(report_id, token):
    data = _ss_fetch(f'reports/{report_id}?pageSize=500&level=1', token)
    if not data.get('rows'): return []
    col_map = {}
    for c in data.get('columns', []):
        if c.get('id'): col_map[c['id']] = c['title']
        if c.get('virtualId'): col_map[c['virtualId']] = c['title']
    rows = []
    for row in data['rows']:
        rec = {}
        for cell in row.get('cells', []):
            col_id = cell.get('virtualColumnId') or cell.get('columnId')
            title = col_map.get(col_id)
            if title:
                rec[title] = cell.get('displayValue') or cell.get('value') or ''
        if rec: rows.append(rec)
    return rows

def normalize_month(v):
    if not v: return None
    s = str(v).strip()
    if not s: return None
    for m in MONTH_NAMES:
        if m.lower() == s.lower(): return m
    abbr = s[:3].lower()
    abbr_map = {m[:3].lower(): m for m in MONTH_NAMES}
    if abbr in abbr_map: return abbr_map[abbr]
    try:
        from datetime import datetime as dt
        d = dt.strptime(s, '%Y-%m-%d')
        return MONTH_NAMES[d.month - 1]
    except: pass
    try:
        from datetime import datetime as dt
        d = dt.strptime(s, '%m/%d/%Y')
        return MONTH_NAMES[d.month - 1]
    except: pass
    return None

def safe_float(v, default=0.0):
    if v is None: return default
    try: return float(v)
    except: return default


# ── Fetch KPI data ──

def fetch_kpi_data(token, month_filter):
    data = {}
    for src in SYNC_SOURCES:
        try:
            if src.get('reportId'):
                rows = fetch_report_rows(src['reportId'], token)
            else:
                rows = fetch_sheet_rows(src['sheetId'], token)
        except Exception as e:
            print(f"  WARNING: Failed to fetch {src['key']}: {e}")
            continue

        kpi_row = src['kpi_row']
        campus_col = src['campusCol']
        month_col = src.get('monthCol')
        planned_col = src.get('plannedCol')
        actual_col = src.get('actualCol')
        value_col = src.get('valueCol')
        is_yes_no = src.get('yesNoCount', False)

        campus_agg = {}
        for row in rows:
            campus = str(row.get(campus_col, '')).strip()
            if not campus: continue
            if month_col and month_filter:
                row_month = normalize_month(row.get(month_col))
                if row_month != month_filter:
                    row_month = normalize_month(row.get('Reporting Month'))
                    if row_month != month_filter:
                        row_month = normalize_month(row.get('Date Reported'))
                        if row_month != month_filter:
                            row_month = normalize_month(row.get('Primary'))
                            if row_month != month_filter:
                                continue

            if campus not in campus_agg:
                campus_agg[campus] = {'planned': 0, 'actual': 0}

            if is_yes_no:
                p = 1 if str(row.get(planned_col, '')).strip().lower() == 'yes' else 0
                a = 1 if str(row.get(actual_col, '')).strip().lower() == 'yes' else 0
                campus_agg[campus]['planned'] += p
                campus_agg[campus]['actual'] += a
            elif planned_col and actual_col:
                campus_agg[campus]['planned'] += safe_float(row.get(planned_col))
                campus_agg[campus]['actual'] += safe_float(row.get(actual_col))
            elif value_col:
                v = safe_float(row.get(value_col))
                campus_agg[campus]['planned'] += v
                campus_agg[campus]['actual'] += v

        weight = KPI_WEIGHTS.get(kpi_row, 0.05)
        for campus, agg in campus_agg.items():
            if campus not in data:
                data[campus] = {}
            planned = agg['planned']
            achieved = agg['actual']
            calc = min(achieved / planned, 1.0) if planned > 0 else (1.0 if achieved > 0 else 0)
            data[campus][kpi_row] = {'planned': planned, 'achieved': achieved, 'calc': calc, 'weight': weight}

    return data

def fetch_waste_data(token, month_filter):
    try:
        rows = fetch_sheet_rows(WASTE_SOURCE['sheetId'], token)
    except:
        return {}
    waste = {}
    for row in rows:
        campus = str(row.get(WASTE_SOURCE['campusCol'], '')).strip()
        month = normalize_month(row.get(WASTE_SOURCE['monthCol']))
        if not campus or month != month_filter: continue
        entry = {}
        for col in WASTE_TABLE_COLS:
            entry[col] = safe_float(row.get(col))
        entry['Total Waste'] = safe_float(row.get('Total Waste'))
        if entry['Total Waste'] == 0:
            entry['Total Waste'] = sum(entry.get(c, 0) for c in WASTE_TABLE_COLS)
        waste[campus] = entry
    return waste


# ── KPI processing ──

def read_campus_data(kpi_data, sheet_name):
    campus = kpi_data.get(sheet_name, {})
    kpis = []
    pillar_scores = []
    for pillar in PILLAR_KPIS:
        p_kpis = []
        for row in pillar['rows']:
            d = campus.get(row, {'planned': 0, 'achieved': 0, 'calc': 1.0, 'weight': KPI_WEIGHTS.get(row, 0.05)})
            p_kpis.append(d)
            kpis.append(d)
        tw = sum(k['weight'] for k in p_kpis)
        score = sum(k['calc'] * k['weight'] for k in p_kpis) / tw if tw > 0 else 0
        pillar_scores.append(score)
    weights = [p['weight'] for p in PILLAR_KPIS]
    overall = sum(s * w for s, w in zip(pillar_scores, weights))
    return {'sheet': sheet_name, 'kpis': kpis, 'pillar_scores': pillar_scores, 'overall': overall}

def read_region_data(kpi_data, region_cfg):
    campuses = [read_campus_data(kpi_data, s) for s in region_cfg['sheets']]
    n = len(campuses)
    avg_p = [sum(c['pillar_scores'][i] for c in campuses)/n for i in range(5)]
    avg_o = sum(c['overall'] for c in campuses)/n
    return {'campuses': campuses, 'avg_pillar': avg_p, 'avg_overall': avg_o, 'short': region_cfg['short']}


# ── XML helper for OOXML Word ──

def w_el(tag, **attrs):
    e = ET.Element(f'{{{W}}}{tag}')
    for k, v in attrs.items():
        e.set(f'{{{W}}}{k}', v)
    return e

def make_run(text, bold=False, size=None, color=None, italic=False):
    r = w_el('r')
    rPr = w_el('rPr')
    if bold: rPr.append(w_el('b'))
    if italic: rPr.append(w_el('i'))
    if size:
        sz = w_el('sz'); sz.set(f'{{{W}}}val', str(size*2)); rPr.append(sz)
        sz2 = w_el('szCs'); sz2.set(f'{{{W}}}val', str(size*2)); rPr.append(sz2)
    if color:
        cl = w_el('color'); cl.set(f'{{{W}}}val', color); rPr.append(cl)
    r.append(rPr)
    t = w_el('t'); t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    return r

def make_para(text='', bold=False, size=None, color=None, center=False,
              style=None, space_before=0, space_after=0, italic=False):
    p = w_el('p')
    pPr = w_el('pPr')
    if style:
        ps = w_el('pStyle'); ps.set(f'{{{W}}}val', style); pPr.append(ps)
    if center:
        jc = w_el('jc'); jc.set(f'{{{W}}}val', 'center'); pPr.append(jc)
    if space_before or space_after:
        sp = w_el('spacing')
        if space_before: sp.set(f'{{{W}}}before', str(space_before))
        if space_after: sp.set(f'{{{W}}}after', str(space_after))
        pPr.append(sp)
    p.append(pPr)
    if text:
        p.append(make_run(text, bold=bold, size=size, color=color, italic=italic))
    return p

def make_page_break():
    p = w_el('p')
    r = w_el('r')
    br = w_el('br'); br.set(f'{{{W}}}type', 'page')
    r.append(br); p.append(r)
    return p

def make_cell(text, w_dxa, bold=False, size=10, color=None, fill=None, center=False, italic=False, colspan=None):
    tc = w_el('tc')
    tcPr = w_el('tcPr')
    tcW = w_el('tcW'); tcW.set(f'{{{W}}}w', str(w_dxa)); tcW.set(f'{{{W}}}type', 'dxa')
    tcPr.append(tcW)
    if colspan:
        gm = w_el('gridSpan'); gm.set(f'{{{W}}}val', str(colspan)); tcPr.append(gm)
    if fill:
        shd = w_el('shd')
        shd.set(f'{{{W}}}val', 'clear')
        shd.set(f'{{{W}}}color', 'auto')
        shd.set(f'{{{W}}}fill', fill)
        tcPr.append(shd)
    tc.append(tcPr)
    txt_color = color
    if txt_color is None and fill and fill not in ('auto', 'FFFFFF', ''):
        txt_color = WHITE
    tc.append(make_para(text, bold=bold, size=size, color=txt_color, center=center, italic=italic))
    return tc

def make_tbl(total_w=9360):
    tbl = w_el('tbl')
    tp = w_el('tblPr')
    ts = w_el('tblStyle'); ts.set(f'{{{W}}}val', 'TableGrid'); tp.append(ts)
    tw = w_el('tblW'); tw.set(f'{{{W}}}w', str(total_w)); tw.set(f'{{{W}}}type', 'dxa'); tp.append(tw)
    borders = w_el('tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = w_el(side)
        b.set(f'{{{W}}}val', 'single')
        b.set(f'{{{W}}}sz', '4')
        b.set(f'{{{W}}}space', '0')
        b.set(f'{{{W}}}color', 'auto')
        borders.append(b)
    tp.append(borders)
    tbl.append(tp)
    return tbl


# ── Report building ──

def make_cover_table(period, region_name, subtitle):
    tbl = make_tbl()

    def row2(label, value, bold_val=False, label_fill=BRAND, label_w=2880, val_w=6480):
        tr = w_el('tr')
        tr.append(make_cell(label, label_w, bold=True, size=10, fill=label_fill))
        tr.append(make_cell(value, val_w, bold=bold_val, size=10))
        return tr

    def row_header(text):
        tr = w_el('tr')
        tr.append(make_cell(text, 9360, bold=True, size=10, fill=NAVY, colspan=2))
        return tr

    tbl.append(row2('Report Title:', f'Corporate OHS Monthly SLA & KPI Report — {region_name}', bold_val=True))
    tbl.append(row2('Client Company:', 'Higher Colleges of Technology'))
    tbl.append(row2('Campuses:', subtitle))
    tbl.append(row2('Issued By:', 'Corporate OHS LLC OPC'))
    tbl.append(row2('Reporting Period:', period, bold_val=True))
    tbl.append(row_header('Document Production/Approval Record'))
    return tbl


def make_kpi_summary_table(region_data, region_cfg):
    tbl = make_tbl()
    W1, W2, W3, W4, W5 = 600, 2800, 1400, 1400, 1400

    # Header
    tr = w_el('tr')
    tr.append(make_cell('#', W1, bold=True, size=9, fill=BRAND, center=True))
    tr.append(make_cell('KPI', W2, bold=True, size=9, fill=BRAND, center=True))
    for name in region_cfg['short']:
        tr.append(make_cell(name, W3, bold=True, size=9, fill=BRAND, center=True))
    tr.append(make_cell('Average', W5, bold=True, size=9, fill=BRAND, center=True))
    tbl.append(tr)

    # KPI rows
    campuses = region_data['campuses']
    kpi_idx = 0
    for pillar in PILLAR_KPIS:
        # Pillar header
        tr = w_el('tr')
        tr.append(make_cell('', W1, size=9, fill='D9E2F3'))
        tr.append(make_cell(pillar['pillar'], W2, bold=True, size=9, fill='D9E2F3'))
        for _ in region_cfg['short']:
            tr.append(make_cell('', W3, size=9, fill='D9E2F3'))
        tr.append(make_cell('', W5, size=9, fill='D9E2F3'))
        tbl.append(tr)

        for row_num in pillar['rows']:
            tr = w_el('tr')
            tr.append(make_cell(str(row_num), W1, size=9, center=True))
            tr.append(make_cell(KPI_NAMES.get(row_num, f'KPI {row_num}'), W2, size=9))
            vals = []
            for c in campuses:
                kpi = c['kpis'][kpi_idx] if kpi_idx < len(c['kpis']) else {'calc': 0}
                pct = round(kpi['calc'] * 100)
                color = '00B050' if pct >= 90 else ('FFC000' if pct >= 70 else 'FF0000')
                tr.append(make_cell(f'{pct}%', W3, size=9, center=True, color=color))
                vals.append(kpi['calc'])
            avg = sum(vals) / len(vals) if vals else 0
            avg_pct = round(avg * 100)
            avg_color = '00B050' if avg_pct >= 90 else ('FFC000' if avg_pct >= 70 else 'FF0000')
            tr.append(make_cell(f'{avg_pct}%', W5, size=9, center=True, color=avg_color, bold=True))
            tbl.append(tr)
            kpi_idx += 1

    # Pillar scores
    tr = w_el('tr')
    tr.append(make_cell('', W1, size=9, fill=NAVY))
    tr.append(make_cell('Pillar Weighted Scores', W2, bold=True, size=9, fill=NAVY))
    for c in campuses:
        tr.append(make_cell(f'{round(c["overall"]*100)}%', W3, bold=True, size=9, fill=NAVY, center=True))
    tr.append(make_cell(f'{round(region_data["avg_overall"]*100)}%', W5, bold=True, size=9, fill=NAVY, center=True))
    tbl.append(tr)

    return tbl


def make_waste_table(waste_data, region_cfg):
    tbl = make_tbl()
    cols = ['Campus', 'Total Waste', 'General', 'Recyclable', 'Hazardous']
    widths = [2200, 1600, 1600, 1600, 2360]

    tr = w_el('tr')
    for lbl, w in zip(cols, widths):
        tr.append(make_cell(lbl, w, bold=True, size=9, fill=BRAND, center=True))
    tbl.append(tr)

    for sheet in region_cfg['sheets']:
        w = waste_data.get(sheet, {})
        total = w.get('Total Waste', 0)
        general = w.get('General Waste', 0)
        recyclable = sum(w.get(c, 0) for c in RECYCLABLE_COLS)
        hazardous = w.get('Hazardous', 0)
        tr = w_el('tr')
        tr.append(make_cell(sheet, 2200, size=9))
        tr.append(make_cell(f'{total:.1f}', 1600, size=9, center=True))
        tr.append(make_cell(f'{general:.1f}', 1600, size=9, center=True))
        tr.append(make_cell(f'{recyclable:.1f}', 1600, size=9, center=True))
        tr.append(make_cell(f'{hazardous:.1f}', 2360, size=9, center=True))
        tbl.append(tr)

    return tbl


def make_executive_summary_table(kpi_data, region_cfg):
    """Executive KPI Summary by Campus table."""
    tbl = make_tbl()
    headers = ['Campus', 'Drills Completion', 'EHS Inspection', 'Findings Closed',
               'Incident Notification', 'Risk Assessment', 'Total Incidents']
    kpi_rows_map = [13, 16, 17, 19, 8]  # KPI row numbers for percentage columns
    widths = [1300, 1300, 1300, 1300, 1400, 1300, 1360]

    # Header row
    tr = w_el('tr')
    for h, wd in zip(headers, widths):
        tr.append(make_cell(h, wd, bold=True, size=8, fill=BRAND, center=True))
    tbl.append(tr)

    # Data rows
    for sheet in region_cfg['sheets']:
        campus = kpi_data.get(sheet, {})
        tr = w_el('tr')
        tr.append(make_cell(sheet, 1300, size=9, bold=True))
        for kr in kpi_rows_map:
            d = campus.get(kr, {'calc': 0})
            pct = round(d.get('calc', 0) * 100)
            clr = '00B050' if pct >= 90 else ('FFC000' if pct >= 70 else 'FF0000')
            tr.append(make_cell(str(pct) + '%', widths[kpi_rows_map.index(kr) + 1], size=9, center=True, color=clr))
        # Total Incidents from notification source (planned = total incidents)
        incidents = int(campus.get(19, {}).get('planned', 0))
        tr.append(make_cell(str(incidents), 1360, size=9, center=True))
        tbl.append(tr)

    return tbl


# ── Build DOCX ──

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/><w:color w:val="180E3F"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/><w:color w:val="1C2340"/></w:rPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:tblPr><w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders></w:tblPr>
  </w:style>
</w:styles>"""

SETTINGS_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="720"/></w:settings>'

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""


def generate_report(region_name, month_name, year, token):
    region_cfg = REGIONS.get(region_name)
    if not region_cfg:
        raise ValueError(f'Unknown region: {region_name}')

    period = f'{month_name} {year}'
    print(f'Generating Word report for {region_name} — {period}')

    # Fetch data
    kpi_data = fetch_kpi_data(token, month_name)
    waste_data = fetch_waste_data(token, month_name)
    region_data = read_region_data(kpi_data, region_cfg)

    # Build document body
    body = w_el('body')

    # Cover page
    body.append(make_para('', space_after=400))
    body.append(make_para('Corporate OHS Monthly', bold=True, size=28, color=BRAND, center=True))
    body.append(make_para('KPI Performance Report', bold=True, size=28, color=BRAND, center=True))
    body.append(make_para('', space_after=200))
    body.append(make_para(region_name, bold=True, size=20, color=NAVY, center=True))
    body.append(make_para(region_cfg['subtitle'], bold=False, size=14, color=GREY, center=True))
    body.append(make_para('', space_after=200))
    body.append(make_para('Reporting Period', bold=True, size=14, color=GREY, center=True))
    body.append(make_para(period, bold=True, size=18, color=BRAND, center=True))
    body.append(make_para('', space_after=400))
    body.append(make_cover_table(period, region_name, region_cfg['subtitle']))
    body.append(make_page_break())

    # Section 1: KPI Summary
    body.append(make_para('1. KPI Performance Summary', style='Heading1'))
    body.append(make_para(f'The following table summarizes the KPI performance for {region_name} campuses during {period}.',
                          size=11, space_after=120))
    body.append(make_kpi_summary_table(region_data, region_cfg))
    body.append(make_para('', space_after=120))

    # Overall score
    overall_pct = round(region_data['avg_overall'] * 100)
    body.append(make_para(f'Overall Weighted Score: {overall_pct}%', bold=True, size=14,
                          color='00B050' if overall_pct >= 90 else ('FFC000' if overall_pct >= 70 else 'FF0000'),
                          center=True, space_before=120, space_after=120))
    body.append(make_page_break())

    # Section 2: Pillar breakdown
    body.append(make_para('2. Pillar Score Breakdown', style='Heading1'))
    for i, pillar in enumerate(PILLAR_KPIS):
        score = round(region_data['avg_pillar'][i] * 100)
        body.append(make_para(f'{pillar["pillar"]}  —  {score}%  (Weight: {int(pillar["weight"]*100)}%)',
                              style='Heading2'))
        # Campus detail for this pillar
        for ci, c in enumerate(region_data['campuses']):
            campus_score = round(c['pillar_scores'][i] * 100)
            body.append(make_para(f'  {region_cfg["short"][ci]}: {campus_score}%', size=11, space_after=40))
        body.append(make_para('', space_after=80))
    body.append(make_page_break())

    # Section 3: Waste Segregation
    body.append(make_para('3. Waste Segregation', style='Heading1'))
    body.append(make_para(f'Waste data for {period}:', size=11, space_after=120))
    body.append(make_waste_table(waste_data, region_cfg))
    body.append(make_para('', space_after=200))

    # Section 4: Executive KPI Summary by Campus
    body.append(make_para('4. Executive KPI Summary by Campus', style='Heading1'))
    body.append(make_para(f'Per-campus KPI performance summary for {region_name} during {period}.',
                          size=11, space_after=120))
    body.append(make_executive_summary_table(kpi_data, region_cfg))
    body.append(make_para('', space_after=120))
    body.append(make_page_break())

    # Section 5: Recommendations
    body.append(make_para('5. Recommendations & Action Items', style='Heading1'))
    body.append(make_para('(To be completed by the EHS team)', italic=True, size=11, color=GREY, space_after=200))

    # sectPr
    sectPr = w_el('sectPr')
    pgSz = w_el('pgSz'); pgSz.set(f'{{{W}}}w', '11906'); pgSz.set(f'{{{W}}}h', '16838')
    pgMar = w_el('pgMar')
    pgMar.set(f'{{{W}}}top', '1440'); pgMar.set(f'{{{W}}}right', '1440')
    pgMar.set(f'{{{W}}}bottom', '1440'); pgMar.set(f'{{{W}}}left', '1440')
    sectPr.append(pgSz); sectPr.append(pgMar)
    body.append(sectPr)

    doc_root = w_el('document')
    doc_root.append(body)

    # Register namespace
    ET.register_namespace('w', W)
    ET.register_namespace('r', R)

    # Build ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml', ET.tostring(doc_root, xml_declaration=True, encoding='UTF-8'))
        z.writestr('word/styles.xml', STYLES_XML)
        z.writestr('word/settings.xml', SETTINGS_XML)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('[Content_Types].xml', CONTENT_TYPES)

    buf.seek(0)
    return buf.getvalue()


# ── HTTP Handler ──

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        region = qs.get('region', [None])[0]
        month = qs.get('month', [None])[0]
        year = qs.get('year', ['2026'])[0]

        if not month:
            self.send_response(400)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'month required'}).encode())
            return

        if region != 'All' and region not in REGIONS:
            self.send_response(400)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': f'Invalid region. Available: {["All"] + list(REGIONS.keys())}'}).encode())
            return

        token = os.environ.get('SMARTSHEET_TOKEN', '')
        if not token:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'SMARTSHEET_TOKEN not set'}).encode())
            return

        try:
            if region == 'All':
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for rname in REGIONS:
                        docx_bytes = generate_report(rname, month, year, token)
                        fname = f'KPI_Report_{rname.replace(" ", "_")}_{month}_{year}.docx'
                        zf.writestr(fname, docx_bytes)
                zip_bytes = zip_buf.getvalue()
                filename = f'KPI_Report_All_Regions_{month}_{year}.zip'
                self.send_response(200)
                self.send_header('Content-Type', 'application/zip')
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.send_header('Content-Length', str(len(zip_bytes)))
                self.end_headers()
                self.wfile.write(zip_bytes)
            else:
                docx_bytes = generate_report(region, month, year, token)
                filename = f'KPI_Report_{region.replace(" ", "_")}_{month}_{year}.docx'
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.send_header('Content-Length', str(len(docx_bytes)))
                self.end_headers()
                self.wfile.write(docx_bytes)
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())
