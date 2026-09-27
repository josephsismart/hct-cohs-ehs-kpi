"""Vercel Python serverless function --ÃÂÃÂ HCT-COHS KPI Word Report Generator.
Template-based: loads word_template.docx, updates charts with live Smartsheet data.
"""

import os, io, json, zipfile, re, copy, base64
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import requests as http_requests
from xml.etree import ElementTree as ET

# --ÃÂÃÂ--ÃÂÃÂ Namespaces --ÃÂÃÂ--ÃÂÃÂ
C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

ET.register_namespace('c', C_NS)
ET.register_namespace('a', A_NS)
ET.register_namespace('w', W_NS)

MONTH_NAMES = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']

# Campus code aliases (template chart labels -> data keys)
CAMPUS_ALIAS = {
    'RAK A': 'RKA', 'RAK B': 'RKB',
    'RAK A\xa0': 'RKA', 'RAK B\xa0': 'RKB',
    'RAK A ': 'RKA', 'RAK B ': 'RKB',
}

# --ÃÂÃÂ--ÃÂÃÂ Campus codes for all regions --ÃÂÃÂ--ÃÂÃÂ
ALL_CAMPUSES = ['ADA','ADB','AAF','AAZ','DMC','DBN','ADH','MZY','FJF','FJH','SJA','SJB','RKA','RKB']
REGION_CAMPUSES = {
    'Abu Dhabi Main': ['ADA','ADB'],
    'Al Ain': ['AAF','AAZ'],
    'Dubai': ['DMC','DBN'],
    'Sharjah': ['SJA','SJB'],
    'Fujairah': ['FJF','FJH'],
    'Ras Al Khaimah': ['RKA','RKB'],
    'Al Dhafra': ['ADH','MZY'],
}

COMMITTEE_MAP = {
    'Abu Dhabi Main': ['ADA','ADB'],
    'Abu Dhabi Main ': ['ADA','ADB'],
    'Al Ain': ['AAF','AAZ'],
    'Dubai': ['DMC','DBN'],
    'Dubai ': ['DMC','DBN'],
    'Sharjah': ['SJA','SJB'],
    'Sharjah ': ['SJA','SJB'],
    'Fujairah': ['FJF','FJH'],
    'Fujairah ': ['FJF','FJH'],
    'Ras Al Khaimah': ['RKA','RKB'],
    'Ras Al Khaimah ': ['RKA','RKB'],
    'RAK': ['RKA','RKB'],
    'Al Dhafra': ['ADH','MZY'],
}

# --ÃÂÃÂ--ÃÂÃÂ Smartsheet sources --ÃÂÃÂ--ÃÂÃÂ
SYNC_SOURCES = [
    {'key': 'v2_hs_kpi_report', 'reportId': '5852576405737348', 'campusCol': 'Committee', 'monthCol': 'Reporting Quarter', 'valueCol': 'KPI 1 - % of HS KPI Reports Submitted', 'kpi_row': 2, 'isolateFromCampusSet': True},
    {'key': 'v2_external_compliance', 'sheetId': '1325212455882628', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'Applicable Legal Compliance', 'actualCol': 'Legal Requirements Complied', 'kpi_row': 4},
    {'key': 'v2_hs_committee', 'sheetId': '5093607634587524', 'campusCol': 'Committee', 'monthCol': 'Reporting Month', 'plannedCol': 'Was a meeting held?', 'actualCol': 'Was a meeting held?', 'kpi_row': 5, 'yesNoCount': True, 'isolateFromCampusSet': True},
    {'key': 'v2_hazard_id', 'sheetId': '7524088825204612', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Controls Identified', 'actualCol': 'Implemented Controls', 'kpi_row': 7},
    {'key': 'v2_risk_closed', 'sheetId': '7524088825204612', 'campusCol': 'Campus', 'monthCol': 'Primary', 'plannedCol': 'Total Risk Assessments Registered', 'actualCol': 'Risk Assessment Closed', 'kpi_row': 8},
    {'key': 'v2_risk_validated', 'sheetId': '7524088825204612', 'campusCol': 'Campus', 'monthCol': 'Primary', 'plannedCol': 'Total Risk Assessments Registered', 'actualCol': 'Risk Assessment and Validation', 'kpi_row': 9},
    {'key': 'v2_planned_training', 'sheetId': '4456464805482372', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned (Yes/No)', 'actualCol': 'Planned (Yes/No)', 'kpi_row': 10, 'yesNoCount': True},
    {'key': 'v2_training_hours', 'sheetId': '4456464805482372', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'valueCol': 'Total Hours', 'kpi_row': 11},
    {'key': 'v2_safe_working', 'sheetId': '2717764266446724', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of activities checked', 'actualCol': 'No. of compliant activities', 'kpi_row': 12},
    {'key': 'v2_drills', 'sheetId': '7139786694283140', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'No. of Planned Drills', 'actualCol': 'No. of Planned Drills Conducted', 'kpi_row': 13},
    {'key': 'v2_permit_to_work', 'sheetId': '3519179394076548', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'No. of PTWs Issued', 'actualCol': 'Total Work Registered', 'kpi_row': 14},
    {'key': 'v2_onsite_induction', 'sheetId': '3519179394076548', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': "No. of New Contractors (Individuals)", 'actualCol': 'No. of Contractors Inducted', 'kpi_row': 15},
    {'key': 'v2_ehs_inspection', 'sheetId': '1510149721116548', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of EHS Inspections Planned', 'actualCol': 'No. of EHS Inspections Completed', 'kpi_row': 17},
    {'key': 'v2_findings_on_time', 'sheetId': '1510149721116548', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of Findings Due', 'actualCol': 'No. of Findings Closed', 'kpi_row': 16},
    {'key': 'v2_investigation_on_time', 'sheetId': '5977763159691140', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident Investigated', 'actualCol': 'Investigation Completed on Time', 'kpi_row': 19},
    {'key': 'notification', 'sheetId': '5977763159691140', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident', 'actualCol': 'Incident Notification Submitted on Time', 'kpi_row': 18},
]

WASTE_SOURCE = {'sheetId': '8150747345538948', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month'}

# Chart-to-KPI mapping: chart index -> KPI row and data type
# type: 'pct_campus' = percentage per campus, 'planned_actual_region' = two series by region,
#       'total_closed_pct_region' = three series by region, 'hours_campus' = raw hours
#       'planned_actual_campus' = two series by campus, 'pct_campus_hq' = percentage incl HQ
CHART_KPI_MAP = {
    1:  {'kpi_row': 2,  'type': 'pct_campus'},
    2:  {'kpi_row': 5,  'type': 'planned_actual_region'},
    3:  {'kpi_row': 6,  'type': 'total_closed_pct_region'},
    4:  {'kpi_row': 7,  'type': 'pct_campus'},
    5:  {'kpi_row': 11, 'type': 'hours_campus'},
    6:  {'kpi_row': 12, 'type': 'pct_campus'},
    7:  {'kpi_row': 13, 'type': 'pct_campus'},
    8:  {'kpi_row': 14, 'type': 'planned_actual_campus'},
    9:  {'kpi_row': 15, 'type': 'planned_actual_campus'},
    10: {'kpi_row': 16, 'type': 'planned_actual_campus'},
    11: {'kpi_row': 17, 'type': 'pct_campus'},
    12: {'kpi_row': 18, 'type': 'pct_campus'},
    13: {'kpi_row': 19, 'type': 'pct_campus'},
}

# --ÃÂÃÂ--ÃÂÃÂ Smartsheet API --ÃÂÃÂ--ÃÂÃÂ

def _ss_fetch(endpoint, token):
    url = f'https://api.smartsheet.com/2.0/{endpoint}'
    resp = http_requests.get(url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'}, timeout=30)
    resp.raise_for_status()
    return resp.json()

def fetch_sheet_rows(sheet_id, token):
    data = _ss_fetch(f'sheets/{sheet_id}?pageSize=10000', token)
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
    data = _ss_fetch(f'reports/{report_id}?pageSize=10000&level=1', token)
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
    # Handle "1. January" prefix format
    m = re.match(r'^\d+\.\s*(.+)$', s)
    if m: s = m.group(1).strip()
    for mn in MONTH_NAMES:
        if mn.lower() == s.lower(): return mn
    abbr = s[:3].lower()
    abbr_map = {mn[:3].lower(): mn for mn in MONTH_NAMES}
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
    except:
        try:
            s = str(v).strip().replace(',', '').replace('%', '')
            return float(s) if s else default
        except: return default

def yes_to_int(v):
    s = str(v).strip().lower()
    return 1 if s in ('yes', 'true', '1') else 0

# --ÃÂÃÂ--ÃÂÃÂ Fetch KPI data per campus --ÃÂÃÂ--ÃÂÃÂ

def fetch_all_kpi_data(token, month_filter):
    """Returns {campus_code: {kpi_row: {'planned': float, 'actual': float}}}"""
    data = {}
    committee_data = {}  # separate for committee (by region name)

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
        is_committee = src.get('isolateFromCampusSet', False)

        agg = {}
        for row in rows:
            campus = str(row.get(campus_col, '')).strip()
            if not campus: continue
            if campus == 'ADC': continue

            # Month filtering
            if month_col and month_filter:
                raw_mv = str(row.get(month_col, '')).strip().upper()
                if raw_mv in ('Q1', 'Q2', 'Q3', 'Q4'):
                    pass  # quarterly data passes through all month filters
                else:
                    row_month = normalize_month(row.get(month_col))
                    if row_month != month_filter:
                        row_month = normalize_month(row.get('Reporting Month'))
                        if row_month != month_filter:
                            row_month = normalize_month(row.get('Primary'))
                            if row_month != month_filter:
                                continue

            if campus not in agg:
                agg[campus] = {'planned': 0, 'actual': 0}

            if is_yes_no:
                agg[campus]['planned'] += yes_to_int(row.get(planned_col))
                agg[campus]['actual'] += yes_to_int(row.get(actual_col))
            elif planned_col and actual_col:
                agg[campus]['planned'] += safe_float(row.get(planned_col))
                agg[campus]['actual'] += safe_float(row.get(actual_col))
            elif value_col:
                v = safe_float(row.get(value_col))
                agg[campus]['planned'] += v
                agg[campus]['actual'] += v

        if is_committee:
            # Committee data is by region name, expand to campuses
            for region_name, vals in agg.items():
                campus_codes = COMMITTEE_MAP.get(region_name, COMMITTEE_MAP.get(region_name.strip(), []))
                if campus_codes:
                    for code in campus_codes:
                        if code not in data: data[code] = {}
                        data[code][kpi_row] = {'planned': vals['planned'], 'actual': vals['actual']}
                # Also store by region name for region-level charts
                committee_data[region_name.strip()] = vals
        else:
            for campus, vals in agg.items():
                if campus not in data: data[campus] = {}
                data[campus][kpi_row] = vals

    return data, committee_data

# --ÃÂÃÂ--ÃÂÃÂ Chart XML updater --ÃÂÃÂ--ÃÂÃÂ

def update_chart_xml(chart_xml_bytes, kpi_data, chart_info, committee_data=None):
    """Update chart XML with real data values."""
    # Fix: Register ALL namespaces from chart XML before parsing
    # to prevent ElementTree from rewriting prefixes (ns0, ns1, etc.)
    # which corrupts OOXML and causes "unreadable content" in Word
    raw_xml = chart_xml_bytes if isinstance(chart_xml_bytes, bytes) else chart_xml_bytes.encode('utf-8')
    xml_str = raw_xml.decode('utf-8')
    for prefix, uri in re.findall(r'xmlns:(\w+)=["\'](.*?)["\' ]', xml_str):
        try:
            ET.register_namespace(prefix, uri)
        except Exception:
            pass
    root = ET.fromstring(raw_xml)
    kpi_row = chart_info['kpi_row']
    chart_type = chart_info['type']

    # Find all series
    all_series = list(root.iter(f'{{{C_NS}}}ser'))

    if chart_type == 'pct_campus':
        # Single series: percentage per campus
        if all_series:
            ser = all_series[0]
            # Get categories from chart
            cats = []
            cat_el = ser.find(f'{{{C_NS}}}cat')
            if cat_el:
                for pt in cat_el.iter(f'{{{C_NS}}}pt'):
                    v = pt.find(f'{{{C_NS}}}v')
                    if v is not None: cats.append(v.text or '')

            # Update values
            val_el = ser.find(f'{{{C_NS}}}val')
            if val_el:
                cache = val_el.find(f'{{{C_NS}}}numRef')
                if cache is None: cache = val_el
                num_cache = cache.find(f'{{{C_NS}}}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                        idx = int(pt.get('idx', 0))
                        if idx < len(cats):
                            campus = cats[idx].strip()
                            campus = CAMPUS_ALIAS.get(campus, campus)
                            d = kpi_data.get(campus, {}).get(kpi_row, {})
                            planned = d.get('planned', 0)
                            actual = d.get('actual', 0)
                            pct = min(actual / planned * 100, 100) if planned > 0 else 0
                            v = pt.find(f'{{{C_NS}}}v')
                            if v is not None: v.text = str(round(pct))

    elif chart_type == 'hours_campus':
        # Single series: raw hours per campus
        if all_series:
            ser = all_series[0]
            cats = []
            cat_el = ser.find(f'{{{C_NS}}}cat')
            if cat_el:
                for pt in cat_el.iter(f'{{{C_NS}}}pt'):
                    v = pt.find(f'{{{C_NS}}}v')
                    if v is not None: cats.append(v.text or '')

            val_el = ser.find(f'{{{C_NS}}}val')
            if val_el:
                cache = val_el.find(f'{{{C_NS}}}numRef')
                if cache is None: cache = val_el
                num_cache = cache.find(f'{{{C_NS}}}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                        idx = int(pt.get('idx', 0))
                        if idx < len(cats):
                            campus = cats[idx].strip()
                            campus = CAMPUS_ALIAS.get(campus, campus)
                            d = kpi_data.get(campus, {}).get(kpi_row, {})
                            hours = d.get('actual', 0)
                            v = pt.find(f'{{{C_NS}}}v')
                            if v is not None: v.text = str(round(hours))

    elif chart_type == 'planned_actual_campus':
        # Two series: planned + actual per campus
        cats = []
        if all_series:
            cat_el = all_series[0].find(f'{{{C_NS}}}cat')
            if cat_el:
                for pt in cat_el.iter(f'{{{C_NS}}}pt'):
                    v = pt.find(f'{{{C_NS}}}v')
                    if v is not None: cats.append(v.text or '')

        for ser_idx, ser in enumerate(all_series):
            field = 'planned' if ser_idx == 0 else 'actual'
            val_el = ser.find(f'{{{C_NS}}}val')
            if val_el:
                cache = val_el.find(f'{{{C_NS}}}numRef')
                if cache is None: cache = val_el
                num_cache = cache.find(f'{{{C_NS}}}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                        idx = int(pt.get('idx', 0))
                        if idx < len(cats):
                            campus = cats[idx].strip()
                            campus = CAMPUS_ALIAS.get(campus, campus)
                            d = kpi_data.get(campus, {}).get(kpi_row, {})
                            val = d.get(field, 0)
                            v = pt.find(f'{{{C_NS}}}v')
                            if v is not None: v.text = str(round(val))

    elif chart_type == 'planned_actual_region':
        # Two series: planned + actual per region
        cats = []
        if all_series:
            cat_el = all_series[0].find(f'{{{C_NS}}}cat')
            if cat_el:
                for pt in cat_el.iter(f'{{{C_NS}}}pt'):
                    v = pt.find(f'{{{C_NS}}}v')
                    if v is not None: cats.append(v.text or '')

        for ser_idx, ser in enumerate(all_series):
            field = 'planned' if ser_idx == 0 else 'actual'
            val_el = ser.find(f'{{{C_NS}}}val')
            if val_el:
                cache = val_el.find(f'{{{C_NS}}}numRef')
                if cache is None: cache = val_el
                num_cache = cache.find(f'{{{C_NS}}}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                        idx = int(pt.get('idx', 0))
                        if idx < len(cats):
                            region_name = cats[idx].strip()
                            d = {}
                            if committee_data:
                                d = committee_data.get(region_name, {})
                            else:
                                # Aggregate from campuses
                                if not d:
                                    campus_codes = REGION_CAMPUSES.get(region_name, [])
                                    if campus_codes:
                                        tot_p = sum(kpi_data.get(c, {}).get(kpi_row, {}).get('planned', 0) for c in campus_codes)
                                        tot_a = sum(kpi_data.get(c, {}).get(kpi_row, {}).get('actual', 0) for c in campus_codes)
                                        d = {'planned': tot_p, 'actual': tot_a}
                            val = d.get(field, 0) if d else 0
                            v = pt.find(f'{{{C_NS}}}v')
                            if v is not None: v.text = str(round(val))

    elif chart_type == 'total_closed_pct_region':
        # Two series: total findings + closed findings, per region
        cats = []
        if all_series:
            cat_el = all_series[0].find(f'{{{C_NS}}}cat')
            if cat_el:
                for pt in cat_el.iter(f'{{{C_NS}}}pt'):
                    v = pt.find(f'{{{C_NS}}}v')
                    if v is not None: cats.append(v.text or '')

        for ser_idx, ser in enumerate(all_series):
            # series 0 = planned (total findings), series 1 = actual (closed)
            field = 'planned' if ser_idx == 0 else 'actual'
            val_el = ser.find(f'{{{C_NS}}}val')
            if val_el:
                cache = val_el.find(f'{{{C_NS}}}numRef')
                if cache is None: cache = val_el
                num_cache = cache.find(f'{{{C_NS}}}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                        idx = int(pt.get('idx', 0))
                        if idx < len(cats):
                            region_name = cats[idx].strip()
                            campus_codes = REGION_CAMPUSES.get(region_name, [])
                            if campus_codes:
                                tot = sum(kpi_data.get(c, {}).get(kpi_row, {}).get(field, 0) for c in campus_codes)
                            else:
                                tot = 0
                            v = pt.find(f'{{{C_NS}}}v')
                            if v is not None: v.text = str(round(tot))

    # Remove externalData references (causes "unreadable content" in Word)
    for ext_data in list(root.iter(f'{{{C_NS}}}externalData')):
        for parent in root.iter():
            if ext_data in list(parent):
                parent.remove(ext_data)
                break

    return ET.tostring(root, encoding='UTF-8', xml_declaration=True)

# --ÃÂÃÂ--ÃÂÃÂ Text replacement --ÃÂÃÂ--ÃÂÃÂ

def replace_text_in_doc(doc_xml_bytes, month_name, year):
    """Replace month/year placeholders in document.xml.
    Handles text split across XML runs (e.g. <w:t>March</w:t> in one run,
    <w:t xml:space="preserve"> 2026</w:t> in another).
    """
    text = doc_xml_bytes.decode('utf-8')
    year = str(year)

    def replace_wt_content(match):
        """Replace month names and years inside <w:t> tag content."""
        tag_open = match.group(1)
        content = match.group(2)
        tag_close = match.group(3)
        for mn in MONTH_NAMES:
            if mn != month_name:
                content = content.replace(mn, month_name)
        for y in range(2024, 2028):
            if str(y) != year:
                content = content.replace(str(y), year)
        return tag_open + content + tag_close

    # Match all <w:t> or <w:t ...> tags and replace month/year in their content
    text = re.sub(r'(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)', replace_wt_content, text)
    return text.encode('utf-8')

# --ÃÂÃÂ--ÃÂÃÂ Main generator --ÃÂÃÂ--ÃÂÃÂ

def generate_report(token, month, year, template_url):
    """Generate Word report from template."""

    month_name = normalize_month(month) or month

    # Download template
    print(f"  Downloading template from {template_url}")
    resp = http_requests.get(template_url, timeout=30)
    resp.raise_for_status()
    template_bytes = resp.content

    # Fetch all KPI data
    print(f"  Fetching KPI data for {month_name} {year}")
    kpi_data, committee_data = fetch_all_kpi_data(token, month_name)
    print(f"  Got data for {len(kpi_data)} campuses")
    # Debug: show per-kpi-row data counts
    kpi_rows_with_data = {}
    for campus, kpis in kpi_data.items():
        for kr, vals in kpis.items():
            if kr not in kpi_rows_with_data:
                kpi_rows_with_data[kr] = {'campuses': 0, 'total_planned': 0, 'total_actual': 0}
            kpi_rows_with_data[kr]['campuses'] += 1
            kpi_rows_with_data[kr]['total_planned'] += vals.get('planned', 0)
            kpi_rows_with_data[kr]['total_actual'] += vals.get('actual', 0)
    for kr in sorted(kpi_rows_with_data.keys()):
        d = kpi_rows_with_data[kr]
        print(f"  KPI row {kr}: {d['campuses']} campuses, planned={d['total_planned']}, actual={d['total_actual']}")

    # Unzip template
    template_io = io.BytesIO(template_bytes)
    output_io = io.BytesIO()

    with zipfile.ZipFile(template_io, 'r') as zin, \
         zipfile.ZipFile(output_io, 'w', zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            raw = zin.read(item.filename)

            # Update chart XML files
            m = re.match(r'word/charts/chart(\d+)\.xml$', item.filename)
            if m:
                chart_num = int(m.group(1))
                if chart_num in CHART_KPI_MAP:
                    print(f"  Updating chart{chart_num}.xml")
                    try:
                        raw = update_chart_xml(raw, kpi_data, CHART_KPI_MAP[chart_num], committee_data)
                    except Exception as e:
                        print(f"  WARNING: Failed to update chart{chart_num}: {e}")

            # Update document.xml (text replacements)
            if item.filename == 'word/document.xml':
                try:
                    raw = replace_text_in_doc(raw, month_name, year)
                except Exception as e:
                    print(f"  WARNING: Failed to update document.xml: {e}")

            # Strip external file references from chart .rels
            if item.filename.startswith('word/charts/_rels/') and item.filename.endswith('.rels'):
                try:
                    rels_text = raw.decode('utf-8')
                    rels_text = re.sub(r'<Relationship[^>]*Target="file:///[^"]*"[^/]*/>', '', rels_text)
                    raw = rels_text.encode('utf-8')
                except Exception as e:
                    print(f"  WARNING: Failed to clean rels {item.filename}: {e}")

            zout.writestr(item, raw)

    output_io.seek(0)
    return output_io.read()

# --ÃÂÃÂ--ÃÂÃÂ HTTP Handler for Vercel --ÃÂÃÂ--ÃÂÃÂ

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        qs = parse_qs(urlparse(self.path).query)
        token = (qs.get('token') or [os.environ.get('SMARTSHEET_TOKEN','')])[0]
        if not token:
            self.send_response(400)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"error": "No token"}).encode())
            return
        month = (qs.get('month') or ['August'])[0] or 'August'
        year = (qs.get('year') or ['2026'])[0]
        debug_mode = (qs.get('debug') or [''])[0] == '1'
        template_url = (qs.get('templateUrl') or ['https://raw.githubusercontent.com/josephsismart/hct-cohs-ehs-kpi/main/templates/word_template.docx'])[0]
        print(f"GET: Generating Word report for {month} {year}")

        if debug_mode:
            try:
                month_name = normalize_month(month) or month
                kpi_data, committee_data = fetch_all_kpi_data(token, month_name)
                debug_info = {
                    'month_filter': month_name,
                    'campuses_with_data': list(kpi_data.keys()),
                    'committee_data_keys': list(committee_data.keys()) if committee_data else [],
                    'kpi_rows': {},
                    'fetch_errors': []
                }
                for campus, kpis in kpi_data.items():
                    for kr, vals in kpis.items():
                        kr_str = str(kr)
                        if kr_str not in debug_info['kpi_rows']:
                            debug_info['kpi_rows'][kr_str] = {'campuses': {}, 'total_planned': 0, 'total_actual': 0}
                        debug_info['kpi_rows'][kr_str]['campuses'][campus] = vals
                        debug_info['kpi_rows'][kr_str]['total_planned'] += vals.get('planned', 0)
                        debug_info['kpi_rows'][kr_str]['total_actual'] += vals.get('actual', 0)
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps(debug_info, indent=2, default=str).encode())
                return
            except Exception as e:
                import traceback
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e), 'trace': traceback.format_exc()}).encode())
                return

        try:
            doc_bytes = generate_report(token, month, year, template_url)
            filename = f'KPI_Report_{month}_{year}.docx'
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(doc_bytes)))
            self.end_headers()
            self.wfile.write(doc_bytes)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length)) if length else {}

            token = body.get('token') or os.environ.get('SMARTSHEET_TOKEN')
            if not token:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"error": "No token provided"}).encode())
                return

            month = body.get('month', 'August')
            year = str(body.get('year', 2026))

            # Template URL from GitHub raw
            template_url = body.get('templateUrl',
                'https://raw.githubusercontent.com/josephsismart/hct-cohs-ehs-kpi/main/templates/word_template.docx')

            print(f"Generating Word report for {month} {year}")
            doc_bytes = generate_report(token, month, year, template_url)

            filename = f'KPI_Report_{month}_{year}.docx'
            b = base64.b64encode(doc_bytes).decode()

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                'filename': filename,
                'base64': b,
                'size': len(doc_bytes)
            }).encode())

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())
"""Vercel Python serverless function --ÂÂ HCT-COHS KPI Word Report Generator.
Template-based: loads word_template.docx, updates charts with live Smartsheet data.
"""

import os, io, json, zipfile, re, copy, base64
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import requests as http_requests
from xml.etree import ElementTree as ET

# --ÂÂ--ÂÂ Namespaces --ÂÂ--ÂÂ
C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

ET.register_namespace('c', C_NS)
ET.register_namespace('a', A_NS)
ET.register_namespace('w', W_NS)

MONTH_NAMES = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']

# Campus code aliases (template chart labels -> data keys)
CAMPUS_ALIAS = {
    'RAK A': 'RKA', 'RAK B': 'RKB',
    'RAK A\xa0': 'RKA', 'RAK B\xa0': 'RKB',
    'RAK A ': 'RKA', 'RAK B ': 'RKB',
}

# --ÂÂ--ÂÂ Campus codes for all regions --ÂÂ--ÂÂ
ALL_CAMPUSES = ['ADA','ADB','AAF','AAZ','DMC','DBN','ADH','MZY','FJF','FJH','SJA','SJB','RKA','RKB']
REGION_CAMPUSES = {
    'Abu Dhabi Main': ['ADA','ADB'],
    'Al Ain': ['AAF','AAZ'],
    'Dubai': ['DMC','DBN'],
    'Sharjah': ['SJA','SJB'],
    'Fujairah': ['FJF','FJH'],
    'Ras Al Khaimah': ['RKA','RKB'],
    'Al Dhafra': ['ADH','MZY'],
}

COMMITTEE_MAP = {
    'Abu Dhabi Main': ['ADA','ADB'],
    'Abu Dhabi Main ': ['ADA','ADB'],
    'Al Ain': ['AAF','AAZ'],
    'Dubai': ['DMC','DBN'],
    'Dubai ': ['DMC','DBN'],
    'Sharjah': ['SJA','SJB'],
    'Sharjah ': ['SJA','SJB'],
    'Fujairah': ['FJF','FJH'],
    'Fujairah ': ['FJF','FJH'],
    'Ras Al Khaimah': ['RKA','RKB'],
    'Ras Al Khaimah ': ['RKA','RKB'],
    'RAK': ['RKA','RKB'],
    'Al Dhafra': ['ADH','MZY'],
}

# --ÂÂ--ÂÂ Smartsheet sources --ÂÂ--ÂÂ
SYNC_SOURCES = [
    {'key': 'v2_hs_kpi_report', 'reportId': '5852576405737348', 'campusCol': 'Committee', 'monthCol': 'Reporting Quarter', 'valueCol': 'KPI 1 - % of HS KPI Reports Submitted', 'kpi_row': 2, 'isolateFromCampusSet': True},
    {'key': 'v2_external_compliance', 'sheetId': '1325212455882628', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'Applicable Legal Compliance', 'actualCol': 'Legal Requirements Complied', 'kpi_row': 4},
    {'key': 'v2_hs_committee', 'sheetId': '5093607634587524', 'campusCol': 'Committee', 'monthCol': 'Reporting Month', 'plannedCol': 'Was a meeting held?', 'actualCol': 'Was a meeting held?', 'kpi_row': 5, 'yesNoCount': True, 'isolateFromCampusSet': True},
    {'key': 'v2_hazard_id', 'sheetId': '7524088825204612', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Controls Identified', 'actualCol': 'Implemented Controls', 'kpi_row': 7},
    {'key': 'v2_risk_closed', 'sheetId': '7524088825204612', 'campusCol': 'Campus', 'monthCol': 'Primary', 'plannedCol': 'Total Risk Assessments Registered', 'actualCol': 'Risk Assessment Closed', 'kpi_row': 8},
    {'key': 'v2_risk_validated', 'sheetId': '7524088825204612', 'campusCol': 'Campus', 'monthCol': 'Primary', 'plannedCol': 'Total Risk Assessments Registered', 'actualCol': 'Risk Assessment and Validation', 'kpi_row': 9},
    {'key': 'v2_planned_training', 'sheetId': '4456464805482372', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned (Yes/No)', 'actualCol': 'Planned (Yes/No)', 'kpi_row': 10, 'yesNoCount': True},
    {'key': 'v2_training_hours', 'sheetId': '4456464805482372', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'valueCol': 'Total Hours', 'kpi_row': 11},
    {'key': 'v2_safe_working', 'sheetId': '2717764266446724', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of activities checked', 'actualCol': 'No. of compliant activities', 'kpi_row': 12},
    {'key': 'v2_drills', 'sheetId': '7139786694283140', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'No. of Planned Drills', 'actualCol': 'No. of Planned Drills Conducted', 'kpi_row': 13},
    {'key': 'v2_permit_to_work', 'sheetId': '3519179394076548', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'No. of PTWs Issued', 'actualCol': 'Total Work Registered', 'kpi_row': 14},
    {'key': 'v2_onsite_induction', 'sheetId': '3519179394076548', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': "No. of New Contractors (Individuals)", 'actualCol': 'No. of Contractors Inducted', 'kpi_row': 15},
    {'key': 'v2_ehs_inspection', 'sheetId': '1510149721116548', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of EHS Inspections Planned', 'actualCol': 'No. of EHS Inspections Completed', 'kpi_row': 17},
    {'key': 'v2_findings_on_time', 'sheetId': '1510149721116548', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of Findings Due', 'actualCol': 'No. of Findings Closed', 'kpi_row': 16},
    {'key': 'v2_investigation_on_time', 'sheetId': '5977763159691140', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident Investigated', 'actualCol': 'Investigation Completed on Time', 'kpi_row': 19},
    {'key': 'notification', 'sheetId': '5977763159691140', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident', 'actualCol': 'Incident Notification Submitted on Time', 'kpi_row': 18},
]

WASTE_SOURCE = {'sheetId': '8150747345538948', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month'}

# Chart-to-KPI mapping: chart index -> KPI row and data type
# type: 'pct_campus' = percentage per campus, 'planned_actual_region' = two series by region,
#       'total_closed_pct_region' = three series by region, 'hours_campus' = raw hours
#       'planned_actual_campus' = two series by campus, 'pct_campus_hq' = percentage incl HQ
CHART_KPI_MAP = {
    1:  {'kpi_row': 2,  'type': 'pct_campus'},
    2:  {'kpi_row': 5,  'type': 'planned_actual_region'},
    3:  {'kpi_row': 6,  'type': 'total_closed_pct_region'},
    4:  {'kpi_row': 7,  'type': 'pct_campus'},
    5:  {'kpi_row': 11, 'type': 'hours_campus'},
    6:  {'kpi_row': 12, 'type': 'pct_campus'},
    7:  {'kpi_row': 13, 'type': 'pct_campus'},
    8:  {'kpi_row': 14, 'type': 'planned_actual_campus'},
    9:  {'kpi_row': 15, 'type': 'planned_actual_campus'},
    10: {'kpi_row': 16, 'type': 'planned_actual_campus'},
    11: {'kpi_row': 17, 'type': 'pct_campus'},
    12: {'kpi_row': 18, 'type': 'pct_campus'},
    13: {'kpi_row': 19, 'type': 'pct_campus'},
}

# --ÂÂ--ÂÂ Smartsheet API --ÂÂ--ÂÂ

def _ss_fetch(endpoint, token):
    url = f'https://api.smartsheet.com/2.0/{endpoint}'
    resp = http_requests.get(url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'}, timeout=30)
    resp.raise_for_status()
    return resp.json()

def fetch_sheet_rows(sheet_id, token):
    data = _ss_fetch(f'sheets/{sheet_id}?pageSize=10000', token)
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
    data = _ss_fetch(f'reports/{report_id}?pageSize=10000&level=1', token)
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
    # Handle "1. January" prefix format
    m = re.match(r'^\d+\.\s*(.+)$', s)
    if m: s = m.group(1).strip()
    for mn in MONTH_NAMES:
        if mn.lower() == s.lower(): return mn
    abbr = s[:3].lower()
    abbr_map = {mn[:3].lower(): mn for mn in MONTH_NAMES}
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
    except:
        try:
            s = str(v).strip().replace(',', '').replace('%', '')
            return float(s) if s else default
        except: return default

def yes_to_int(v):
    s = str(v).strip().lower()
    return 1 if s in ('yes', 'true', '1') else 0

# --ÂÂ--ÂÂ Fetch KPI data per campus --ÂÂ--ÂÂ

def fetch_all_kpi_data(token, month_filter):
    """Returns {campus_code: {kpi_row: {'planned': float, 'actual': float}}}"""
    data = {}
    committee_data = {}  # separate for committee (by region name)

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
        is_committee = src.get('isolateFromCampusSet', False)

        agg = {}
        for row in rows:
            campus = str(row.get(campus_col, '')).strip()
            if not campus: continue
            if campus == 'ADC': continue

            # Month filtering
            if month_col and month_filter:
                raw_mv = str(row.get(month_col, '')).strip().upper()
                if raw_mv in ('Q1', 'Q2', 'Q3', 'Q4'):
                    pass  # quarterly data passes through all month filters
                else:
                    row_month = normalize_month(row.get(month_col))
                    if row_month != month_filter:
                        row_month = normalize_month(row.get('Reporting Month'))
                        if row_month != month_filter:
                            row_month = normalize_month(row.get('Primary'))
                            if row_month != month_filter:
                                continue

            if campus not in agg:
                agg[campus] = {'planned': 0, 'actual': 0}

            if is_yes_no:
                agg[campus]['planned'] += yes_to_int(row.get(planned_col))
                agg[campus]['actual'] += yes_to_int(row.get(actual_col))
            elif planned_col and actual_col:
                agg[campus]['planned'] += safe_float(row.get(planned_col))
                agg[campus]['actual'] += safe_float(row.get(actual_col))
            elif value_col:
                v = safe_float(row.get(value_col))
                agg[campus]['planned'] += v
                agg[campus]['actual'] += v

        if is_committee:
            # Committee data is by region name, expand to campuses
            for region_name, vals in agg.items():
                campus_codes = COMMITTEE_MAP.get(region_name, COMMITTEE_MAP.get(region_name.strip(), []))
                if campus_codes:
                    for code in campus_codes:
                        if code not in data: data[code] = {}
                        data[code][kpi_row] = {'planned': vals['planned'], 'actual': vals['actual']}
                # Also store by region name for region-level charts
                committee_data[region_name.strip()] = vals
        else:
            for campus, vals in agg.items():
                if campus not in data: data[campus] = {}
                data[campus][kpi_row] = vals

    return data, committee_data

# --ÂÂ--ÂÂ Chart XML updater --ÂÂ--ÂÂ

def update_chart_xml(chart_xml_bytes, kpi_data, chart_info, committee_data=None):
    """Update chart XML with real data values."""
    # Fix: Register ALL namespaces from chart XML before parsing
    # to prevent ElementTree from rewriting prefixes (ns0, ns1, etc.)
    # which corrupts OOXML and causes "unreadable content" in Word
    raw_xml = chart_xml_bytes if isinstance(chart_xml_bytes, bytes) else chart_xml_bytes.encode('utf-8')
    xml_str = raw_xml.decode('utf-8')
    for prefix, uri in re.findall(r'xmlns:(\w+)=["\'](.*?)["\' ]', xml_str):
        try:
            ET.register_namespace(prefix, uri)
        except Exception:
            pass
    root = ET.fromstring(raw_xml)
    kpi_row = chart_info['kpi_row']
    chart_type = chart_info['type']

    # Find all series
    all_series = list(root.iter(f'{{{C_NS}}}ser'))

    if chart_type == 'pct_campus':
        # Single series: percentage per campus
        if all_series:
            ser = all_series[0]
            # Get categories from chart
            cats = []
            cat_el = ser.find(f'{{{C_NS}}}cat')
            if cat_el:
                for pt in cat_el.iter(f'{{{C_NS}}}pt'):
                    v = pt.find(f'{{{C_NS}}}v')
                    if v is not None: cats.append(v.text or '')

            # Update values
            val_el = ser.find(f'{{{C_NS}}}val')
            if val_el:
                cache = val_el.find(f'{{{C_NS}}}numRef')
                if cache is None: cache = val_el
                num_cache = cache.find(f'{{{C_NS}}}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                        idx = int(pt.get('idx', 0))
                        if idx < len(cats):
                            campus = cats[idx].strip()
                            campus = CAMPUS_ALIAS.get(campus, campus)
                            d = kpi_data.get(campus, {}).get(kpi_row, {})
                            planned = d.get('planned', 0)
                            actual = d.get('actual', 0)
                            pct = min(actual / planned * 100, 100) if planned > 0 else 0
                            v = pt.find(f'{{{C_NS}}}v')
                            if v is not None: v.text = str(round(pct))

    elif chart_type == 'hours_campus':
        # Single series: raw hours per campus
        if all_series:
            ser = all_series[0]
            cats = []
            cat_el = ser.find(f'{{{C_NS}}}cat')
            if cat_el:
                for pt in cat_el.iter(f'{{{C_NS}}}pt'):
                    v = pt.find(f'{{{C_NS}}}v')
                    if v is not None: cats.append(v.text or '')

            val_el = ser.find(f'{{{C_NS}}}val')
            if val_el:
                cache = val_el.find(f'{{{C_NS}}}numRef')
                if cache is None: cache = val_el
                num_cache = cache.find(f'{{{C_NS}}}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                        idx = int(pt.get('idx', 0))
                        if idx < len(cats):
                            campus = cats[idx].strip()
                            campus = CAMPUS_ALIAS.get(campus, campus)
                            d = kpi_data.get(campus, {}).get(kpi_row, {})
                            hours = d.get('actual', 0)
                            v = pt.find(f'{{{C_NS}}}v')
                            if v is not None: v.text = str(round(hours))

    elif chart_type == 'planned_actual_campus':
        # Two series: planned + actual per campus
        cats = []
        if all_series:
            cat_el = all_series[0].find(f'{{{C_NS}}}cat')
            if cat_el:
                for pt in cat_el.iter(f'{{{C_NS}}}pt'):
                    v = pt.find(f'{{{C_NS}}}v')
                    if v is not None: cats.append(v.text or '')

        for ser_idx, ser in enumerate(all_series):
            field = 'planned' if ser_idx == 0 else 'actual'
            val_el = ser.find(f'{{{C_NS}}}val')
            if val_el:
                cache = val_el.find(f'{{{C_NS}}}numRef')
                if cache is None: cache = val_el
                num_cache = cache.find(f'{{{C_NS}}}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                        idx = int(pt.get('idx', 0))
                        if idx < len(cats):
                            campus = cats[idx].strip()
                            campus = CAMPUS_ALIAS.get(campus, campus)
                            d = kpi_data.get(campus, {}).get(kpi_row, {})
                            val = d.get(field, 0)
                            v = pt.find(f'{{{C_NS}}}v')
                            if v is not None: v.text = str(round(val))

    elif chart_type == 'planned_actual_region':
        # Two series: planned + actual per region
        cats = []
        if all_series:
            cat_el = all_series[0].find(f'{{{C_NS}}}cat')
            if cat_el:
                for pt in cat_el.iter(f'{{{C_NS}}}pt'):
                    v = pt.find(f'{{{C_NS}}}v')
                    if v is not None: cats.append(v.text or '')

        for ser_idx, ser in enumerate(all_series):
            field = 'planned' if ser_idx == 0 else 'actual'
            val_el = ser.find(f'{{{C_NS}}}val')
            if val_el:
                cache = val_el.find(f'{{{C_NS}}}numRef')
                if cache is None: cache = val_el
                num_cache = cache.find(f'{{{C_NS}}}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                        idx = int(pt.get('idx', 0))
                        if idx < len(cats):
                            region_name = cats[idx].strip()
                            d = {}
                            if committee_data:
                                d = committee_data.get(region_name, {})
                            else:
                                # Aggregate from campuses
                                if not d:
                                    campus_codes = REGION_CAMPUSES.get(region_name, [])
                                    if campus_codes:
                                        tot_p = sum(kpi_data.get(c, {}).get(kpi_row, {}).get('planned', 0) for c in campus_codes)
                                        tot_a = sum(kpi_data.get(c, {}).get(kpi_row, {}).get('actual', 0) for c in campus_codes)
                                        d = {'planned': tot_p, 'actual': tot_a}
                            val = d.get(field, 0) if d else 0
                            v = pt.find(f'{{{C_NS}}}v')
                            if v is not None: v.text = str(round(val))

    elif chart_type == 'total_closed_pct_region':
        # Two series: total findings + closed findings, per region
        cats = []
        if all_series:
            cat_el = all_series[0].find(f'{{{C_NS}}}cat')
            if cat_el:
                for pt in cat_el.iter(f'{{{C_NS}}}pt'):
                    v = pt.find(f'{{{C_NS}}}v')
                    if v is not None: cats.append(v.text or '')

        for ser_idx, ser in enumerate(all_series):
            # series 0 = planned (total findings), series 1 = actual (closed)
            field = 'planned' if ser_idx == 0 else 'actual'
            val_el = ser.find(f'{{{C_NS}}}val')
            if val_el:
                cache = val_el.find(f'{{{C_NS}}}numRef')
                if cache is None: cache = val_el
                num_cache = cache.find(f'{{{C_NS}}}numCache')
                if num_cache is not None:
                    for pt in num_cache.findall(f'{{{C_NS}}}pt'):
                        idx = int(pt.get('idx', 0))
                        if idx < len(cats):
                            region_name = cats[idx].strip()
                            campus_codes = REGION_CAMPUSES.get(region_name, [])
                            if campus_codes:
                                tot = sum(kpi_data.get(c, {}).get(kpi_row, {}).get(field, 0) for c in campus_codes)
                            else:
                                tot = 0
                            v = pt.find(f'{{{C_NS}}}v')
                            if v is not None: v.text = str(round(tot))

    # Remove externalData references (causes "unreadable content" in Word)
    for ext_data in list(root.iter(f'{{{C_NS}}}externalData')):
        for parent in root.iter():
            if ext_data in list(parent):
                parent.remove(ext_data)
                break

    return ET.tostring(root, encoding='UTF-8', xml_declaration=True)

# --ÂÂ--ÂÂ Text replacement --ÂÂ--ÂÂ

def replace_text_in_doc(doc_xml_bytes, month_name, year):
    """Replace month/year placeholders in document.xml"""
    text = doc_xml_bytes.decode('utf-8')
    # Replace common placeholders
    for mn in MONTH_NAMES:
        if mn != month_name:
            # Replace "August 2026" etc. with new month
            for y in range(2024, 2028):
                text = text.replace(f'{mn} {y}', f'{month_name} {year}')
    return text.encode('utf-8')

# --ÂÂ--ÂÂ Main generator --ÂÂ--ÂÂ

def generate_report(token, month, year, template_url):
    """Generate Word report from template."""

    month_name = normalize_month(month) or month

    # Download template
    print(f"  Downloading template from {template_url}")
    resp = http_requests.get(template_url, timeout=30)
    resp.raise_for_status()
    template_bytes = resp.content

    # Fetch all KPI data
    print(f"  Fetching KPI data for {month_name} {year}")
    kpi_data, committee_data = fetch_all_kpi_data(token, month_name)
    print(f"  Got data for {len(kpi_data)} campuses")
    # Debug: show per-kpi-row data counts
    kpi_rows_with_data = {}
    for campus, kpis in kpi_data.items():
        for kr, vals in kpis.items():
            if kr not in kpi_rows_with_data:
                kpi_rows_with_data[kr] = {'campuses': 0, 'total_planned': 0, 'total_actual': 0}
            kpi_rows_with_data[kr]['campuses'] += 1
            kpi_rows_with_data[kr]['total_planned'] += vals.get('planned', 0)
            kpi_rows_with_data[kr]['total_actual'] += vals.get('actual', 0)
    for kr in sorted(kpi_rows_with_data.keys()):
        d = kpi_rows_with_data[kr]
        print(f"  KPI row {kr}: {d['campuses']} campuses, planned={d['total_planned']}, actual={d['total_actual']}")

    # Unzip template
    template_io = io.BytesIO(template_bytes)
    output_io = io.BytesIO()

    with zipfile.ZipFile(template_io, 'r') as zin, \
         zipfile.ZipFile(output_io, 'w', zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            raw = zin.read(item.filename)

            # Update chart XML files
            m = re.match(r'word/charts/chart(\d+)\.xml$', item.filename)
            if m:
                chart_num = int(m.group(1))
                if chart_num in CHART_KPI_MAP:
                    print(f"  Updating chart{chart_num}.xml")
                    try:
                        raw = update_chart_xml(raw, kpi_data, CHART_KPI_MAP[chart_num], committee_data)
                    except Exception as e:
                        print(f"  WARNING: Failed to update chart{chart_num}: {e}")

            # Update document.xml (text replacements)
            if item.filename == 'word/document.xml':
                try:
                    raw = replace_text_in_doc(raw, month_name, year)
                except Exception as e:
                    print(f"  WARNING: Failed to update document.xml: {e}")

            # Strip external file references from chart .rels
            if item.filename.startswith('word/charts/_rels/') and item.filename.endswith('.rels'):
                try:
                    rels_text = raw.decode('utf-8')
                    rels_text = re.sub(r'<Relationship[^>]*Target="file:///[^"]*"[^/]*/>', '', rels_text)
                    raw = rels_text.encode('utf-8')
                except Exception as e:
                    print(f"  WARNING: Failed to clean rels {item.filename}: {e}")

            zout.writestr(item, raw)

    output_io.seek(0)
    return output_io.read()

# --ÂÂ--ÂÂ HTTP Handler for Vercel --ÂÂ--ÂÂ

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        qs = parse_qs(urlparse(self.path).query)
        token = (qs.get('token') or [os.environ.get('SMARTSHEET_TOKEN','')])[0]
        if not token:
            self.send_response(400)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"error": "No token"}).encode())
            return
        month = (qs.get('month') or ['August'])[0] or 'August'
        year = (qs.get('year') or ['2026'])[0]
        debug_mode = (qs.get('debug') or [''])[0] == '1'
        template_url = (qs.get('templateUrl') or ['https://raw.githubusercontent.com/josephsismart/hct-cohs-ehs-kpi/main/templates/word_template.docx'])[0]
        print(f"GET: Generating Word report for {month} {year}")

        if debug_mode:
            try:
                month_name = normalize_month(month) or month
                kpi_data, committee_data = fetch_all_kpi_data(token, month_name)
                debug_info = {
                    'month_filter': month_name,
                    'campuses_with_data': list(kpi_data.keys()),
                    'committee_data_keys': list(committee_data.keys()) if committee_data else [],
                    'kpi_rows': {},
                    'fetch_errors': []
                }
                for campus, kpis in kpi_data.items():
                    for kr, vals in kpis.items():
                        kr_str = str(kr)
                        if kr_str not in debug_info['kpi_rows']:
                            debug_info['kpi_rows'][kr_str] = {'campuses': {}, 'total_planned': 0, 'total_actual': 0}
                        debug_info['kpi_rows'][kr_str]['campuses'][campus] = vals
                        debug_info['kpi_rows'][kr_str]['total_planned'] += vals.get('planned', 0)
                        debug_info['kpi_rows'][kr_str]['total_actual'] += vals.get('actual', 0)
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps(debug_info, indent=2, default=str).encode())
                return
            except Exception as e:
                import traceback
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e), 'trace': traceback.format_exc()}).encode())
                return

        try:
            doc_bytes = generate_report(token, month, year, template_url)
            filename = f'KPI_Report_{month}_{year}.docx'
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(doc_bytes)))
            self.end_headers()
            self.wfile.write(doc_bytes)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length)) if length else {}

            token = body.get('token') or os.environ.get('SMARTSHEET_TOKEN')
            if not token:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"error": "No token provided"}).encode())
                return

            month = body.get('month', 'August')
            year = str(body.get('year', 2026))

            # Template URL from GitHub raw
            template_url = body.get('templateUrl',
                'https://raw.githubusercontent.com/josephsismart/hct-cohs-ehs-kpi/main/templates/word_template.docx')

            print(f"Generating Word report for {month} {year}")
            doc_bytes = generate_report(token, month, year, template_url)

            filename = f'KPI_Report_{month}_{year}.docx'
            b = base64.b64encode(doc_bytes).decode()

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                'filename': filename,
                'base64': b,
                'size': len(doc_bytes)
            }).encode())

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())
