"""Vercel Python serverless function — HCT-COHS KPI Word Report Generator.
Fetches live data from Smartsheet API and generates downloadable .docx files
matching the client's exact report template with 13 editable charts.
"""

import os, io, json, zipfile
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET

# ── Namespaces ──
W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

NAVY  = '1C2340'
BRAND = '180E3F'
TEAL  = '00249C'
GREY  = '595959'
WHITE = 'FFFFFF'

# ── All 14 campus codes (matches client template x-axis order) ──
ALL_CAMPUSES = ['ADA','ADB','AAF','AAZ','DMC','DBN','SJA','SJB','FJF','FJH','RKA','RKB','ADH','MZY']
ALL_CAMPUSES_WITH_HQ = ['ADA','ADB','AAF','AAZ','DMC','DBN','SJA','SJB','FJF','FJH','RKA','RKB','ADH','MZY','HQ']

# Chart display labels
CAMPUS_LABELS = ['ADA','ADB','AAF','AAZ','DMC','DBN','SJA','SJB','FJF','FJH','RAK A','RAK B','ADH','MZY']
CAMPUS_LABELS_HQ = CAMPUS_LABELS + ['HQ']

# Region groupings for committee charts (Charts 2,3)
REGION_GROUPS = {
    'Abu Dhabi Main': ['ADA','ADB'],
    'Al Ain':         ['AAF','AAZ'],
    'Dubai':          ['DMC','DBN'],
    'Sharjah':        ['SJA','SJB'],
    'Fujairah':       ['FJF','FJH'],
    'RAK':            ['RKA','RKB'],
    'Remote':         ['ADH','MZY'],
}
REGION_ORDER = ['Abu Dhabi Main','Al Ain','Dubai','Sharjah','Fujairah','RAK','Remote']

# ── Per-region config (for individual region reports) ──
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
    2: 'HS KPI Report Submission', 3: 'Total Hours of Training',
    4: 'External Authority Compliance', 5: 'HS Committee Meeting',
    6: 'Hazard Identification', 7: 'Risk Assessment Closed',
    8: 'Risk Assessment Validated', 9: 'Safe Working Procedure',
    10: 'Planned Training', 11: 'Training Hours Delivered',
    12: 'Operational Control Procedures', 13: 'Emergency Drills',
    14: 'Permit to Work', 15: 'Onsite Safety Induction',
    16: 'Scheduled EHS Inspection', 17: 'Findings Closed On Time',
    18: 'Investigation Completed on Time', 19: 'Incident Notification on Time',
}

# ── Smartsheet sources ──
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

TRAINING_SOURCE = {'sheetId': '8549734774951812', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'hoursCol': 'Total Training Hours Delivered'}

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

def fetch_training_hours(token, month_filter):
    try:
        rows = fetch_sheet_rows(TRAINING_SOURCE['sheetId'], token)
    except:
        return {}
    hours = {}
    for row in rows:
        campus = str(row.get(TRAINING_SOURCE['campusCol'], '')).strip()
        month = normalize_month(row.get(TRAINING_SOURCE['monthCol']))
        if not campus or month != month_filter: continue
        hours[campus] = hours.get(campus, 0) + safe_float(row.get(TRAINING_SOURCE['hoursCol']))
    return hours


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


# ── XML helpers ──

def _esc(text):
    return str(text).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

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


# ── OOXML Chart Generators ──

CHART_BLUE = '00249C'
CHART_ORANGE = 'ED7D31'
CHART_GREEN = '198754'
CHART_RED = 'DC3545'
CHART_YELLOW = 'FFC107'

def make_clustered_bar_xml(title, categories, series_list, show_data_labels=True, format_code='General'):
    """Build OOXML clustered column chart. series_list: [{name, values, color}]"""
    L = []
    L.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    L.append('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
             ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
             ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    L.append('<c:chart>')
    L.append('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>'
             f'<a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr>'
             f'<a:r><a:rPr lang="en-US" sz="1100" b="1"/><a:t>{_esc(title)}</a:t></a:r></a:p>'
             '</c:rich></c:tx><c:overlay val="0"/></c:title>')
    L.append('<c:autoTitleDeleted val="0"/>')
    L.append('<c:plotArea><c:layout/>')
    L.append('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>')
    for si, s in enumerate(series_list):
        clr = s.get('color', CHART_BLUE)
        L.append(f'<c:ser><c:idx val="{si}"/><c:order val="{si}"/>')
        L.append(f'<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/>'
                 f'<c:pt idx="0"><c:v>{_esc(s["name"])}</c:v></c:pt></c:strCache></c:strRef></c:tx>')
        L.append(f'<c:spPr><a:solidFill><a:srgbClr val="{clr}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>')
        if show_data_labels:
            L.append(f'<c:dLbls><c:numFmt formatCode="{_esc(format_code)}" sourceLinked="0"/>'
                     '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
                     '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
        L.append(f'<c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="{len(categories)}"/>')
        for ci, cat in enumerate(categories):
            L.append(f'<c:pt idx="{ci}"><c:v>{_esc(cat)}</c:v></c:pt>')
        L.append('</c:strCache></c:strRef></c:cat>')
        L.append(f'<c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>{_esc(format_code)}</c:formatCode>'
                 f'<c:ptCount val="{len(s["values"])}"/>')
        for vi, v in enumerate(s['values']):
            L.append(f'<c:pt idx="{vi}"><c:v>{v}</c:v></c:pt>')
        L.append('</c:numCache></c:numRef></c:val></c:ser>')
    L.append('<c:axId val="1"/><c:axId val="2"/></c:barChart>')
    L.append('<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
             '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>')
    L.append('<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
             '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>')
    L.append('</c:plotArea>')
    if len(series_list) > 1:
        L.append('<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>')
    L.append('<c:plotVisOnly val="1"/></c:chart></c:chartSpace>')
    return '\n'.join(L).encode('utf-8')


def make_pct_bar_xml(title, categories, values):
    """Single-series bar chart with per-bar color coding (green/yellow/red) for percentages."""
    L = []
    L.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    L.append('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
             ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
             ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    L.append('<c:chart>')
    L.append('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>'
             f'<a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr>'
             f'<a:r><a:rPr lang="en-US" sz="1100" b="1"/><a:t>{_esc(title)}</a:t></a:r></a:p>'
             '</c:rich></c:tx><c:overlay val="0"/></c:title>')
    L.append('<c:autoTitleDeleted val="0"/>')
    L.append('<c:plotArea><c:layout/>')
    L.append('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="1"/>')
    L.append('<c:ser><c:idx val="0"/><c:order val="0"/>')
    L.append(f'<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/>'
             f'<c:pt idx="0"><c:v>Percentage</c:v></c:pt></c:strCache></c:strRef></c:tx>')
    # Per-bar colors
    for vi, val in enumerate(values):
        clr = CHART_GREEN if val >= 90 else (CHART_YELLOW if val >= 70 else CHART_RED)
        L.append(f'<c:dPt><c:idx val="{vi}"/><c:spPr><a:solidFill><a:srgbClr val="{clr}"/>'
                 f'</a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:dPt>')
    L.append('<c:dLbls><c:numFmt formatCode="0&quot;%&quot;" sourceLinked="0"/>'
             '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
             '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
    L.append(f'<c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="{len(categories)}"/>')
    for ci, cat in enumerate(categories):
        L.append(f'<c:pt idx="{ci}"><c:v>{_esc(cat)}</c:v></c:pt>')
    L.append('</c:strCache></c:strRef></c:cat>')
    L.append(f'<c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>0</c:formatCode>'
             f'<c:ptCount val="{len(values)}"/>')
    for vi, val in enumerate(values):
        L.append(f'<c:pt idx="{vi}"><c:v>{val}</c:v></c:pt>')
    L.append('</c:numCache></c:numRef></c:val></c:ser>')
    L.append('<c:axId val="1"/><c:axId val="2"/></c:barChart>')
    L.append('<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
             '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>')
    L.append('<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:max val="100"/></c:scaling>'
             '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>')
    L.append('</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>')
    return '\n'.join(L).encode('utf-8')


def make_chart_para(rid, chart_idx):
    CX = 5486400
    CY = 3200400
    drawing_xml = (
        f'<w:drawing xmlns:w="{W}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}"'
        f' xmlns:c="{C_NS}" xmlns:r="{R_NS}">'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{CX}" cy="{CY}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{chart_idx + 100}" name="Chart {chart_idx}"/>'
        f'<a:graphic><a:graphicData uri="{C_NS}">'
        f'<c:chart r:id="{rid}"/>'
        f'</a:graphicData></a:graphic>'
        f'</wp:inline></w:drawing>'
    )
    p_el = w_el('p')
    pPr = w_el('pPr')
    jc = w_el('jc'); jc.set(f'{{{W}}}val', 'center'); pPr.append(jc)
    sp = w_el('spacing'); sp.set(f'{{{W}}}before', '120'); sp.set(f'{{{W}}}after', '80'); pPr.append(sp)
    p_el.append(pPr)
    r = w_el('r')
    drawing_el = ET.fromstring(drawing_xml)
    r.append(drawing_el)
    p_el.append(r)
    return p_el


# ── Build 13 Charts matching Client Template ──

def build_consolidated_charts(kpi_data, training_hours):
    """Build 13 charts matching client template exactly."""
    charts = []

    def campus_pct(kpi_row, campuses=ALL_CAMPUSES):
        vals = []
        for c in campuses:
            d = kpi_data.get(c, {}).get(kpi_row, {})
            vals.append(round(d.get('calc', 0) * 100))
        return vals

    def campus_values(kpi_row, field, campuses=ALL_CAMPUSES):
        vals = []
        for c in campuses:
            d = kpi_data.get(c, {}).get(kpi_row, {})
            vals.append(round(safe_float(d.get(field, 0))))
        return vals

    def region_agg(kpi_row, field):
        vals = []
        for rname in REGION_ORDER:
            total = 0
            for c in REGION_GROUPS[rname]:
                d = kpi_data.get(c, {}).get(kpi_row, {})
                total += safe_float(d.get(field, 0))
            vals.append(round(total))
        return vals

    # Chart 1: External Authority Compliance (%) - 14 campuses
    charts.append(('Figure 1: External Authority Compliance Requirements and Complied',
        make_pct_bar_xml('External Authority Compliance Rate', CAMPUS_LABELS,
            campus_pct(4))))

    # Chart 2: Committee Meetings - 2 series by region
    charts.append(('Figure 2: Committee Meetings Planned vs Conducted',
        make_clustered_bar_xml('Committee Meetings', REGION_ORDER, [
            {'name': 'Meeting Planned', 'values': region_agg(5, 'planned'), 'color': CHART_BLUE},
            {'name': 'Meeting Conducted', 'values': region_agg(5, 'achieved'), 'color': CHART_ORANGE},
        ])))

    # Chart 3: Committee Actions Closed - 2 series by region (using hazard/risk data)
    charts.append(('Figure 3: Committee Actions - Total vs Closed',
        make_clustered_bar_xml('Committee Actions Closed', REGION_ORDER, [
            {'name': 'Total', 'values': region_agg(6, 'planned'), 'color': CHART_BLUE},
            {'name': 'Closed', 'values': region_agg(6, 'achieved'), 'color': CHART_ORANGE},
        ])))

    # Chart 4: Risk Control Measures (%) - 14 campuses
    charts.append(('Figure 4: Risk Control Measures Implemented',
        make_pct_bar_xml('Risk Control Measures', CAMPUS_LABELS,
            campus_pct(7))))

    # Chart 5: Training Hours - 14 campuses (actual hours, not %)
    training_vals = [round(training_hours.get(c, 0)) for c in ALL_CAMPUSES]
    charts.append(('Figure 5: Actual Training Hours Delivered',
        make_clustered_bar_xml('Training Hours', CAMPUS_LABELS, [
            {'name': 'Actual Training Hours', 'values': training_vals, 'color': CHART_BLUE},
        ], format_code='#,##0')))

    # Chart 6: Compliance Activities (%) - 14 campuses (OCP)
    charts.append(('Figure 6: Operational Control Compliance Activities',
        make_pct_bar_xml('Compliance Activities', CAMPUS_LABELS,
            campus_pct(12))))

    # Chart 7: Emergency Drills (%) - 14 campuses
    charts.append(('Figure 7: Emergency Drills Completion',
        make_pct_bar_xml('Emergency Drills', CAMPUS_LABELS,
            campus_pct(13))))

    # Chart 8: PTW - 2 series, 14 campuses
    charts.append(('Figure 8: Permit to Work (PTW)',
        make_clustered_bar_xml('Permit to Work', CAMPUS_LABELS, [
            {'name': 'Total Works Registered', 'values': campus_values(14, 'planned'), 'color': CHART_BLUE},
            {'name': 'Number of PTWs Issued', 'values': campus_values(14, 'achieved'), 'color': CHART_ORANGE},
        ])))

    # Chart 9: Contractors - 2 series, 14 campuses
    charts.append(('Figure 9: Contractor Induction',
        make_clustered_bar_xml('Contractors', CAMPUS_LABELS, [
            {'name': 'Total Active Contractors', 'values': campus_values(15, 'planned'), 'color': CHART_BLUE},
            {'name': 'Inducted Contractors', 'values': campus_values(15, 'achieved'), 'color': CHART_ORANGE},
        ])))

    # Chart 10: Inspections - 2 series, 14 campuses
    charts.append(('Figure 10: EHS Inspections',
        make_clustered_bar_xml('EHS Inspections', CAMPUS_LABELS, [
            {'name': 'Inspection Planned', 'values': campus_values(16, 'planned'), 'color': CHART_BLUE},
            {'name': 'Inspection Completed', 'values': campus_values(16, 'achieved'), 'color': CHART_ORANGE},
        ])))

    # Chart 11: Findings Closed (%) - 14 campuses
    charts.append(('Figure 11: Findings Closed on Time',
        make_pct_bar_xml('Findings Closed', CAMPUS_LABELS,
            campus_pct(17))))

    # Chart 12: Incident Notifications (%) - 15 (with HQ)
    charts.append(('Figure 12: Incident Notification on Time',
        make_pct_bar_xml('Incident Notifications', CAMPUS_LABELS_HQ,
            campus_pct(19, ALL_CAMPUSES_WITH_HQ))))

    # Chart 13: Investigations (%) - 15 (with HQ)
    charts.append(('Figure 13: Investigation Completed on Time',
        make_pct_bar_xml('Investigations', CAMPUS_LABELS_HQ,
            campus_pct(18, ALL_CAMPUSES_WITH_HQ))))

    return charts


# ── Tables ──

def make_cover_table(period):
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
    tbl.append(row2('Report Title:', f'Corporate OHS Monthly SLA & KPI Report', bold_val=True))
    tbl.append(row2('Client Company:', 'Higher Colleges of Technology'))
    tbl.append(row2('Campuses:', 'All Campuses'))
    tbl.append(row2('Issued By:', 'Corporate OHS LLC OPC'))
    tbl.append(row2('Reporting Period:', period, bold_val=True))
    tbl.append(row_header('Document Production/Approval Record'))
    return tbl


def make_kpi_eval_table(kpi_data):
    """Section 1: EHS KPI Evaluation Table matching client template.
    Columns: #, KPI Description, Abu Dhabi (ADA+ADB avg), Al Ain, Dubai, Sharjah, Fujairah, RAK, Remote, Overall"""
    tbl = make_tbl()
    region_names = ['Abu Dhabi', 'Al Ain', 'Dubai', 'Sharjah', 'Fujairah', 'RAK', 'Remote']
    W0, WKPI = 400, 1800
    WREG = 900

    # Header
    tr = w_el('tr')
    tr.append(make_cell('#', W0, bold=True, size=8, fill=BRAND, center=True))
    tr.append(make_cell('KPI Description', WKPI, bold=True, size=8, fill=BRAND, center=True))
    for rn in region_names:
        tr.append(make_cell(rn, WREG, bold=True, size=7, fill=BRAND, center=True))
    tr.append(make_cell('Overall', WREG, bold=True, size=8, fill=BRAND, center=True))
    tbl.append(tr)

    kpi_idx = 0
    for pillar in PILLAR_KPIS:
        # Pillar header row
        tr = w_el('tr')
        tr.append(make_cell('', W0, size=7, fill='D9E2F3'))
        tr.append(make_cell(pillar['pillar'], WKPI, bold=True, size=7, fill='D9E2F3'))
        for _ in region_names:
            tr.append(make_cell('', WREG, size=7, fill='D9E2F3'))
        tr.append(make_cell('', WREG, size=7, fill='D9E2F3'))
        tbl.append(tr)

        for row_num in pillar['rows']:
            tr = w_el('tr')
            tr.append(make_cell(str(row_num), W0, size=7, center=True))
            tr.append(make_cell(KPI_NAMES.get(row_num, ''), WKPI, size=7))
            all_vals = []
            for ri, rn in enumerate(REGION_ORDER):
                campus_codes = REGION_GROUPS[rn]
                pcts = []
                for cc in campus_codes:
                    d = kpi_data.get(cc, {}).get(row_num, {})
                    pcts.append(d.get('calc', 0))
                avg = sum(pcts) / len(pcts) if pcts else 0
                pct = round(avg * 100)
                clr = '00B050' if pct >= 90 else ('FFC000' if pct >= 70 else 'FF0000')
                tr.append(make_cell(f'{pct}%', WREG, size=7, center=True, color=clr))
                all_vals.append(avg)
            overall = sum(all_vals) / len(all_vals) if all_vals else 0
            opct = round(overall * 100)
            oclr = '00B050' if opct >= 90 else ('FFC000' if opct >= 70 else 'FF0000')
            tr.append(make_cell(f'{opct}%', WREG, size=7, center=True, color=oclr, bold=True))
            tbl.append(tr)
            kpi_idx += 1

    return tbl


def make_waste_table_all(waste_data):
    """Section 5: Full waste segregation table - 14 campuses x 11 waste types matching client template."""
    tbl = make_tbl(total_w=14400)
    cols = ['Campus'] + WASTE_TABLE_COLS
    widths = [1200] + [1200] * len(WASTE_TABLE_COLS)

    tr = w_el('tr')
    for lbl, w in zip(cols, widths):
        short = lbl.replace('Single Use Plastic', 'SUP').replace('Paper Cup/Carton', 'Cup/Carton')
        tr.append(make_cell(short, w, bold=True, size=6, fill=BRAND, center=True))
    tbl.append(tr)

    for campus_code in ALL_CAMPUSES:
        w = waste_data.get(campus_code, {})
        tr = w_el('tr')
        tr.append(make_cell(campus_code, 1200, size=7, bold=True))
        for col in WASTE_TABLE_COLS:
            val = w.get(col, 0)
            tr.append(make_cell(f'{val:.1f}' if val else '-', 1200, size=7, center=True))
        tbl.append(tr)

    return tbl


# ── Document Structure ──

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
    <w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/><w:color w:val="2E75B6"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="2E75B6"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="40"/><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="2E75B6"/></w:rPr>
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


def generate_consolidated_report(month_name, year, token):
    """Generate the consolidated report matching client template exactly."""
    period = f'{month_name} {year}'
    print(f'Generating consolidated Word report — {period}')

    kpi_data = fetch_kpi_data(token, month_name)
    waste_data = fetch_waste_data(token, month_name)
    training_hours = fetch_training_hours(token, month_name)

    # Build 13 charts
    chart_items = build_consolidated_charts(kpi_data, training_hours)

    body = w_el('body')

    # ── Cover Page ──
    body.append(make_para('', space_after=600))
    body.append(make_para('HCT', bold=True, size=36, color=BRAND, center=True))
    body.append(make_para('Corporate OHS Monthly', bold=True, size=24, color=NAVY, center=True))
    body.append(make_para('COHS KPI Report', bold=True, size=24, color=NAVY, center=True))
    body.append(make_para('', space_after=200))
    body.append(make_para(period, bold=True, size=20, color=BRAND, center=True))
    body.append(make_para('', space_after=400))
    body.append(make_cover_table(period))
    body.append(make_page_break())

    # ── Table of Contents placeholder ──
    body.append(make_para('Table of Contents', bold=True, size=18, color=NAVY))
    body.append(make_para('(Update field to refresh page numbers)', italic=True, size=10, color=GREY))
    body.append(make_para('', space_after=200))
    body.append(make_page_break())

    # ── Section 1: EHS KPI Evaluation Table ──
    body.append(make_para('1. EHS KPI Evaluation', style='Heading1'))
    body.append(make_para(f'Monthly SLA & KPI performance evaluation for all campuses — {period}.', size=10, space_after=120))
    body.append(make_kpi_eval_table(kpi_data))
    body.append(make_page_break())

    # ── Section 2: Executive Summary ──
    body.append(make_para('2. Executive Summary', style='Heading1'))
    body.append(make_para('(Analysis to be added)', italic=True, size=10, color=GREY, space_after=200))
    body.append(make_page_break())

    # ── Section 3: Incidents ──
    body.append(make_para('3. Incidents', style='Heading1'))
    body.append(make_para('(Analysis to be added)', italic=True, size=10, color=GREY, space_after=200))
    body.append(make_page_break())

    # ── Section 4: Planned Actions and Initiatives ──
    body.append(make_para('4. Planned Actions and Initiatives', style='Heading1'))
    body.append(make_para('(Analysis to be added)', italic=True, size=10, color=GREY, space_after=200))
    body.append(make_page_break())

    # ── Section 5: Waste Segregation ──
    body.append(make_para('5. Waste Segregation', style='Heading1'))
    body.append(make_para(f'Waste segregation data for all campuses — {period}.', size=10, space_after=120))
    body.append(make_waste_table_all(waste_data))
    body.append(make_page_break())

    # ── Sections 6-10: Campus KPI Tracking with Charts ──
    # Prepare chart files
    chart_rels = []
    chart_overrides = []
    chart_files = {}

    # Section 6: Leadership, Accountability & Engagement (Charts 1-3)
    body.append(make_para('6. Campus KPI Tracking (Leadership, Accountability & Engagement)', style='Heading1'))
    body.append(make_para('6.1  % of H&S KPI Reports Submitted vs Planned', style='Heading2'))
    body.append(make_para('All campuses submitted their KPI reports on time using the approved KPI reporting template.', size=10, space_after=120))
    body.append(make_para('6.2  % of Audit Findings Closed', style='Heading2'))
    body.append(make_para('During the reporting month, no external environment, health, and safety audits were conducted by third-party entities.', size=10, space_after=120))
    body.append(make_para('6.3  % Authority Compliance Rate', style='Heading3'))

    # Chart 1
    ci = 0
    rid = f'rIdChart{ci+1}'
    chart_title, chart_xml = chart_items[ci]
    body.append(make_chart_para(rid, ci+1))
    body.append(make_para(chart_title, italic=True, size=9, color=GREY, center=True, space_after=80))
    body.append(make_para('(Analysis to be added)', italic=True, size=10, color=GREY, space_after=120))
    chart_files[f'word/charts/chart{ci+1}.xml'] = chart_xml
    chart_files[f'word/charts/_rels/chart{ci+1}.xml.rels'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    chart_rels.append(f'  <Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{ci+1}.xml"/>')
    chart_overrides.append(f'  <Override PartName="/word/charts/chart{ci+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')

    body.append(make_para('6.4  HS Committee', style='Heading3'))

    # Chart 2
    ci = 1
    rid = f'rIdChart{ci+1}'
    chart_title, chart_xml = chart_items[ci]
    body.append(make_chart_para(rid, ci+1))
    body.append(make_para(chart_title, italic=True, size=9, color=GREY, center=True, space_after=80))
    body.append(make_para('(Analysis to be added)', italic=True, size=10, color=GREY, space_after=120))
    chart_files[f'word/charts/chart{ci+1}.xml'] = chart_xml
    chart_files[f'word/charts/_rels/chart{ci+1}.xml.rels'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    chart_rels.append(f'  <Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{ci+1}.xml"/>')
    chart_overrides.append(f'  <Override PartName="/word/charts/chart{ci+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')

    # Chart 3
    ci = 2
    rid = f'rIdChart{ci+1}'
    chart_title, chart_xml = chart_items[ci]
    body.append(make_chart_para(rid, ci+1))
    body.append(make_para(chart_title, italic=True, size=9, color=GREY, center=True, space_after=80))
    body.append(make_para('(Analysis to be added)', italic=True, size=10, color=GREY, space_after=120))
    chart_files[f'word/charts/chart{ci+1}.xml'] = chart_xml
    chart_files[f'word/charts/_rels/chart{ci+1}.xml.rels'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    chart_rels.append(f'  <Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{ci+1}.xml"/>')
    chart_overrides.append(f'  <Override PartName="/word/charts/chart{ci+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
    body.append(make_page_break())

    # Section 7: Risk Management & Planning (Chart 4)
    body.append(make_para('7. Campus KPI Tracking (Risk Management & Planning)', style='Heading1'))
    ci = 3
    rid = f'rIdChart{ci+1}'
    chart_title, chart_xml = chart_items[ci]
    body.append(make_chart_para(rid, ci+1))
    body.append(make_para(chart_title, italic=True, size=9, color=GREY, center=True, space_after=80))
    body.append(make_para('(Analysis to be added)', italic=True, size=10, color=GREY, space_after=120))
    chart_files[f'word/charts/chart{ci+1}.xml'] = chart_xml
    chart_files[f'word/charts/_rels/chart{ci+1}.xml.rels'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    chart_rels.append(f'  <Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{ci+1}.xml"/>')
    chart_overrides.append(f'  <Override PartName="/word/charts/chart{ci+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
    body.append(make_page_break())

    # Section 8: Training & Awareness (Chart 5)
    body.append(make_para('8. Campus KPI Tracking (Training & Awareness)', style='Heading1'))
    ci = 4
    rid = f'rIdChart{ci+1}'
    chart_title, chart_xml = chart_items[ci]
    body.append(make_chart_para(rid, ci+1))
    body.append(make_para(chart_title, italic=True, size=9, color=GREY, center=True, space_after=80))
    body.append(make_para('(Analysis to be added)', italic=True, size=10, color=GREY, space_after=120))
    chart_files[f'word/charts/chart{ci+1}.xml'] = chart_xml
    chart_files[f'word/charts/_rels/chart{ci+1}.xml.rels'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    chart_rels.append(f'  <Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{ci+1}.xml"/>')
    chart_overrides.append(f'  <Override PartName="/word/charts/chart{ci+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
    body.append(make_page_break())

    # Section 9: Operational Control & Emergency Preparedness (Charts 6-9)
    body.append(make_para('9. Campus KPI Tracking (Operational Control & Emergency Preparedness)', style='Heading1'))
    for ci in range(5, 9):
        rid = f'rIdChart{ci+1}'
        chart_title, chart_xml = chart_items[ci]
        body.append(make_chart_para(rid, ci+1))
        body.append(make_para(chart_title, italic=True, size=9, color=GREY, center=True, space_after=80))
        body.append(make_para('(Analysis to be added)', italic=True, size=10, color=GREY, space_after=120))
        chart_files[f'word/charts/chart{ci+1}.xml'] = chart_xml
        chart_files[f'word/charts/_rels/chart{ci+1}.xml.rels'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        chart_rels.append(f'  <Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{ci+1}.xml"/>')
        chart_overrides.append(f'  <Override PartName="/word/charts/chart{ci+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
    body.append(make_page_break())

    # Section 10: Performance Evaluation & Continual Improvement (Charts 10-13)
    body.append(make_para('10. Campus KPI Tracking (Performance Evaluation & Continual Improvement)', style='Heading1'))
    for ci in range(9, 13):
        rid = f'rIdChart{ci+1}'
        chart_title, chart_xml = chart_items[ci]
        body.append(make_chart_para(rid, ci+1))
        body.append(make_para(chart_title, italic=True, size=9, color=GREY, center=True, space_after=80))
        body.append(make_para('(Analysis to be added)', italic=True, size=10, color=GREY, space_after=120))
        chart_files[f'word/charts/chart{ci+1}.xml'] = chart_xml
        chart_files[f'word/charts/_rels/chart{ci+1}.xml.rels'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        chart_rels.append(f'  <Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{ci+1}.xml"/>')
        chart_overrides.append(f'  <Override PartName="/word/charts/chart{ci+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
    body.append(make_page_break())

    # ── Appendix A: Action Plan Status ──
    body.append(make_para('Appendix A: Action Plan Status', style='Heading1'))
    body.append(make_para('(To be completed)', italic=True, size=10, color=GREY, space_after=200))

    # sectPr - A4
    sectPr = w_el('sectPr')
    pgSz = w_el('pgSz'); pgSz.set(f'{{{W}}}w', '11906'); pgSz.set(f'{{{W}}}h', '16838')
    pgMar = w_el('pgMar')
    pgMar.set(f'{{{W}}}top', '1440'); pgMar.set(f'{{{W}}}right', '1080')
    pgMar.set(f'{{{W}}}bottom', '1440'); pgMar.set(f'{{{W}}}left', '1080')
    sectPr.append(pgSz); sectPr.append(pgMar)
    body.append(sectPr)

    doc_root = w_el('document')
    doc_root.append(body)

    ET.register_namespace('w', W)
    ET.register_namespace('r', R)

    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n'
        + '\n'.join(chart_rels) + '\n'
        '</Relationships>'
    )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
        '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
        '  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
        + '\n'.join(chart_overrides) + '\n'
        '</Types>'
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml', ET.tostring(doc_root, xml_declaration=True, encoding='UTF-8'))
        z.writestr('word/styles.xml', STYLES_XML)
        z.writestr('word/settings.xml', SETTINGS_XML)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('[Content_Types].xml', content_types)
        for path, data in chart_files.items():
            z.writestr(path, data)

    buf.seek(0)
    return buf.getvalue()


def generate_region_report(region_name, month_name, year, token):
    """Generate per-region report (legacy)."""
    region_cfg = REGIONS.get(region_name)
    if not region_cfg:
        raise ValueError(f'Unknown region: {region_name}')

    period = f'{month_name} {year}'
    kpi_data = fetch_kpi_data(token, month_name)
    waste_data = fetch_waste_data(token, month_name)
    region_data = read_region_data(kpi_data, region_cfg)

    body = w_el('body')

    # Cover
    body.append(make_para('', space_after=400))
    body.append(make_para('Corporate OHS Monthly', bold=True, size=28, color=BRAND, center=True))
    body.append(make_para('KPI Performance Report', bold=True, size=28, color=BRAND, center=True))
    body.append(make_para('', space_after=200))
    body.append(make_para(region_name, bold=True, size=20, color=NAVY, center=True))
    body.append(make_para(region_cfg['subtitle'], size=14, color=GREY, center=True))
    body.append(make_para('', space_after=200))
    body.append(make_para(period, bold=True, size=18, color=BRAND, center=True))
    body.append(make_page_break())

    # KPI Summary
    body.append(make_para('1. KPI Performance Summary', style='Heading1'))
    body.append(make_para(f'KPI performance for {region_name} — {period}.', size=10, space_after=120))

    tbl = make_tbl()
    W1, W2, W3, W5 = 600, 2800, 1400, 1400
    tr = w_el('tr')
    tr.append(make_cell('#', W1, bold=True, size=9, fill=BRAND, center=True))
    tr.append(make_cell('KPI', W2, bold=True, size=9, fill=BRAND, center=True))
    for name in region_cfg['short']:
        tr.append(make_cell(name, W3, bold=True, size=9, fill=BRAND, center=True))
    tr.append(make_cell('Average', W5, bold=True, size=9, fill=BRAND, center=True))
    tbl.append(tr)

    campuses = region_data['campuses']
    kpi_idx = 0
    for pillar in PILLAR_KPIS:
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
            tr.append(make_cell(KPI_NAMES.get(row_num, ''), W2, size=9))
            vals = []
            for c in campuses:
                kpi = c['kpis'][kpi_idx] if kpi_idx < len(c['kpis']) else {'calc': 0}
                pct = round(kpi['calc'] * 100)
                clr = '00B050' if pct >= 90 else ('FFC000' if pct >= 70 else 'FF0000')
                tr.append(make_cell(f'{pct}%', W3, size=9, center=True, color=clr))
                vals.append(kpi['calc'])
            avg = sum(vals)/len(vals) if vals else 0
            avg_pct = round(avg * 100)
            avg_clr = '00B050' if avg_pct >= 90 else ('FFC000' if avg_pct >= 70 else 'FF0000')
            tr.append(make_cell(f'{avg_pct}%', W5, size=9, center=True, color=avg_clr, bold=True))
            tbl.append(tr)
            kpi_idx += 1
    body.append(tbl)

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
    ET.register_namespace('w', W)
    ET.register_namespace('r', R)

    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n'
        '</Relationships>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
        '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
        '  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
        '</Types>'
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('word/document.xml', ET.tostring(doc_root, xml_declaration=True, encoding='UTF-8'))
        z.writestr('word/styles.xml', STYLES_XML)
        z.writestr('word/settings.xml', SETTINGS_XML)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('[Content_Types].xml', content_types)

    buf.seek(0)
    return buf.getvalue()


# ── HTTP Handler ──

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        region = qs.get('region', ['All'])[0]
        month = qs.get('month', [None])[0]
        year = qs.get('year', ['2026'])[0]

        if not month:
            self.send_response(400)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'month required'}).encode())
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
                docx_bytes = generate_consolidated_report(month, year, token)
                filename = f'HCT_COHS_KPI_Report_{month}_{year}.docx'
            else:
                if region not in REGIONS:
                    self.send_response(400)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'error': f'Invalid region'}).encode())
                    return
                docx_bytes = generate_region_report(region, month, year, token)
                filename = f'KPI_Report_{region.replace(" ", "_")}_{month}_{year}.docx'

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(docx_bytes)))
            self.end_headers()
            self.wfile.write(docx_bytes)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())
