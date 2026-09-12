"""Vercel Python serverless function - HCT-COHS KPI PPT Generator.
Generates quarterly (Q1+Q2) KPI reports using client's reference template.
Fetches live hdata from Smartsheet API.
"""

import os, re, io, json, zipfile, tempfile
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen
import ssl
import certifi
from datetime import datetime
from xml.etree import ElementTree as ET

# -- Namespaces --
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relahtionships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)
ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
ET.register_namespace('c14', 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart')
ET.register_namespace('c15', 'http://schemas.microsoft.com/office/drawing/2012/chart')
ET.register_namespace('c16r2', 'http://schemas.microsoft.com/office/drawing/2015/06/chart')

# -- Regions --
REGIONS = {
    'AD Al Ain': {'sheets': ['AAF','AAZ'], 'short': ['Falaj Hazza','Zakhir'], 'subtitle': 'Al Ain Falaj Hazza & Al Ain Zakhir'},
    'Abu Dhabi': {'sheets': ['ADA','ADB'], 'short': ['Baniyas A','Baniyas B'], 'subtitle': 'Abu Dhabi Baniyas A & Abu Dhabi Baniyas B'},
    'AD Remote': {'sheets': ['ADH','MZY'], 'short': ['Al Dhanna','Madinat Zayed'], 'subtitle': 'Al Dhanna Ruwais & Al Dhafra Madinat Zayed City'},
    'Dubai': {'sheets': ['DMC','DBN'], 'short': ['Academic City','Al Nahda'], 'subtitle': 'Dubai Academic City & Dubai Al Nahda'},
    'Fujairah': {'sheets': ['FJF','FJH'], 'short': ['Faseel','Hulaifat'], 'subtitle': 'Fujairah Faseel & Fujairah Hulaifat'},
    'Sharjah': {'sheets': ['SHJA','SHJB'], 'short': ['Campus A','Campus B'], 'subtitle': 'Sharjah Campus A & Sharjah Campus B'},
    'Ras Al Khaimah': {'sheets': ['RKA','RKB'], 'short': ['Campus A','Campus B'], 'subtitle': 'RAK Campus A & RAK Campus B'},
}

# -- KPI structure --
KPI_WEIGHTS = {
    2: 0.30, 3: 0.10, 4: 0.10, 5: 0.25, 6: 0.25,
    7: 0.30, 8: 0.50, 9: 0.20,
    10: 0.50, 11: 0.50,
    12: 0.40, 13: 0.40, 14: 0.10, 15: 0.10,
    16: 0.30, 17: 0.30, 18: 0.20, 19: 0.20,
}

PILLAR_CLASSIFICATION = [
    {'pillar': 'Leadership', 'weight': 0.20, 'groups': [
        {'class_weight': 0.5, 'kpis': [(2, 0.6), (3, 0.2), (4, 0.2)]},
        {'class_weight': 0.5, 'kpis': [(5, 0.5), (6, 0.5)]},
    ]},
    {'pillar': 'Risk Mgmt', 'weight': 0.20, 'groups': [
        {'class_weight': 0.3, 'kpis': [(7, 1.0)]},
        {'class_weight': 0.5, 'kpis': [(8, 1.0)]},
        {'class_weight': 0.2, 'kpis': [(9, 1.0)]},
    ]},
    {'pillar': 'Training', 'weight': 0.10, 'groups': [
        {'class_weight': 0.5, 'kpis': [(10, 1.0)]},
        {'class_weight': 0.5, 'kpis': [(11, 1.0)]},
    ]},
    {'pillar': 'OCP & Emerg', 'weight': 0.25, 'groups': [
        {'class_weight': 0.4, 'kpis': [(12, 1.0)]},
        {'class_weight': 0.4, 'kpis': [(13, 1.0)]},
        {'class_weight': 0.2, 'kpis': [(14, 0.5), (15, 0.5)]},
    ]},
    {'pillar': 'Perf Eval', 'weight': 0.25, 'groups': [
        {'class_weight': 0.6, 'kpis': [(16, 0.5), (17, 0.5)]},
        {'class_weight': 0.4, 'kpis': [(18, 0.5), (19, 0.5)]},
    ]},
]

# Chart file -> kpi_row mapping (kpi_row = KPI_display_number + 1)
# Q1 charts (slides 4-11)
Q1_CHART_MAP = {
    'chart2.xml': 4,    # KPI 3
    'chart3.xml': 5,    # KPI 4
    'chart4.xml': 6,    # KPI 5
    'chart5.xml': 7,    # KPI 6
    'chart6.xml': 10,   # KPI 9
    'chart7.xml': 12,   # KPI 11
    'chart8.xml': 13,   # KPI 12
    'chart9.xml': 14,   # KPI 13
    'chart10.xml': 15,  # KPI 14
    'chart11.xml': 16,  # KPI 15
    'chart12.xml': 17,  # KPI 16
    'chart13.xml': 18,  # KPI 17
    'chart14.xml': 19,  # KPI 18
}
# Q2 charts (slides 14-21)
Q2_CHART_MAP = {
    'chart16.xml': 4,   # KPI 3
    'chart17.xml': 5,   # KPI 4
    'chart18.xml': 6,   # KPI 5
    'chart19.xml': 7,   # KPI 6
    'chart20.xml': 10,  # KPI 9
    'chart21.xml': 12,  # KPI 11
    'chart22.xml': 13,  # KPI 12
    'chart23.xml': 14,  # KPI 13
    'chart24.xml': 15,  # KPI 14
    'chart25.xml': 16,  # KPI 15
    'chart26.xml': 17,  # KPI 16
    'chart27.xml': 18,  # KPI 17
    'chart28.xml': 19,  # KPI 18
}

PIE_CHARTS = {'chart1.xml', 'chart15.xml'}

# -- Smartsheet sources --
SYNC_SOURCES = [
    {'key': 'v2_hs_kpi_report', 'reportId': '5852576405737348', 'campusCol': 'Campuses', 'monthCol': 'Primary', 'valueCol': 'Submitted', 'kpi_row': 2},
    {'key': 'v2_external_compliance', 'sheetId': '1325212455882628', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'Applicable Legal Compliance', 'actualCol': 'Legal Requirements Complied', 'kpi_row': 4},
    {'key': 'v2_hs_committee', 'sheetId': '5093607634587524', 'campusCol': 'Committee', 'monthCol': 'Reporting Month', 'plannedCol': 'Was a meeting held?', 'actualCol': 'Was a meeting held?', 'kpi_row': 5},
    {'key': 'v2_hazard_id', 'sheetId': '7524088825204612', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Controls Identified', 'actualCol': 'Implemented Controls', 'kpi_row': 7},
    {'key': 'v2_risk_closed', 'sheetId': '7524088825204612', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'Total Risk Assessments Registered', 'actualCol': 'Risk Assessment Closed', 'kpi_row': 8},
    {'key': 'v2_risk_validated', 'sheetId': '7524088825204612', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'Total Risk Assessments Registered', 'actualCol': 'Risk Assessment and Validation', 'kpi_row': 9},
    {'key': 'v2_planned_training', 'sheetId': '4456464805482372', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned (Yes/No)', 'actualCol': 'Planned (Yes/No)', 'kpi_row': 10, 'yesNoCount': True},
    {'key': 'v2_safe_working', 'sheetId': '2717764266446724', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of activities checked', 'actualCol': 'No. of compliant activities', 'kpi_row': 12},
    {'key': 'v2_drills', 'sheetId': '7139786694283140', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'No. of Planned Drills', 'actualCol': 'No. of Planned Drills Conducted', 'kpi_row': 13},
    {'key': 'v2_permit_to_work', 'sheetId': '3519179394076548', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'No. of PTWs Issued', 'actualCol': 'Total Work Registered', 'kpi_row': 14},
    {'key': 'v2_onsite_induction', 'sheetId': '3519179394076548', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': "No. of New Contractors (Individuals)", 'actualCol': 'No. of Contractors Inducted', 'kpi_row': 15},
    {'key': 'v2_ehs_inspection', 'sheetId': '1510149721116548', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of EHS Inspections Planned', 'actualCol': 'No. of EHS Inspections Completed', 'kpi_row': 17},
    {'key': 'v2_findings_on_time', 'sheetId': '1510149721116548', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of Findings Due', 'actualCol': 'No. of Findings Closed', 'kpi_row': 16},
    {'key': 'v2_investigation_on_time', 'reportId': '5432865759121284', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident Investigated', 'actualCol': 'Investigation Completed on Time', 'kpi_row': 19},
    {'key': 'notification', 'reportId': '8527961731846020', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident', 'actualCol': 'Incident Notification Submitted on Time', 'kpi_row': 18},
]

COMMITTEE_MAP = {
    'Al Ain': ['AAF', 'AAZ'], 'Abu Dhabi': ['ADA', 'ADB'],
    'Dubai': ['DMC', 'DBN'], 'Fujairah': ['FJF', 'FJH'],
    'Sharjah': ['SHJA', 'SHJB'], 'Ras Al Khaimah': ['RKA', 'RKB'],
    'AD Remote': ['ADH', 'MZY'], 'Al Dhafra': ['ADH', 'MZY'],
    'Ruwais': ['ADH', 'MZY'],
}

MONTH_NAMES = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']
Q1_MONTHS = ['January', 'February', 'March']
Q2_MONTHS = ['April', 'May', 'June']

# KPI display names for NA boxes and text placeholders
KPI_NAMES = {
    2: 'Health and Safety Quarterly KPI Reports Submitted & Presented',
    3: 'External Audit Findings Closed',
    4: 'Legal Requirements Complied',
    5: 'H&S Committee Meetings Conducted',
    6: 'Action Items Closed',
    7: 'Hazard Identification Controls Implemented',
    8: 'Risk Assessments Closed',
    9: 'Risk Assessments and Validation',
    10: 'H&S Training Sessions Completed',
    11: 'H&S Awareness Campaigns Conducted',
    12: 'Safe Working Procedures Compliance',
    13: 'Emergency Drills Conducted',
    14: 'Permit to Work Compliance',
    15: 'Contractor Induction Compliance',
    16: 'Inspection Findings Closed on Time',
    17: 'EHS Inspections Completed',
    18: 'Incident Notification Submitted on Time',
    19: 'Incident Investigation Completed on Time',
}

# KPI chart titles matching client's KPI NAMES.xlsx
# Format: "KPI {display_num} - {name}" where display_num = kpi_row - 1
KPI_CHART_TITLES = {
    2: 'KPI 1 - % of Health and Safety KPI Reports Submitted vs Planned',
    3: 'KPI 2 - % of Audit Findings Closed',
    4: 'KPI 3 - % Authority Compliance Rate',
    5: 'KPI 4 - % of HS Committee Meetings Conducted as per TORs',
    6: 'KPI 5 - % of Committee Meetings & Management Review & Committee Meeting Actions Closed',
    7: 'KPI 6 - % of Implemented Control Measures',
    8: 'KPI 7 - % of Risk Assessments Closed',
    9: 'KPI 8 - % Risk Assessments and Validation',
    10: 'KPI 9 - % of Planned H&S Training Hours Delivered',
    11: 'KPI 10 - % of H&S Awareness Campaigns Conducted',
    12: 'KPI 11 - % of Compliance to Activities as per Set Health & Safety Procedures',
    13: 'KPI 12 - % of Emergency Drills Conducted on Schedule',
    14: 'KPI 13 - % Compliance to PTW (Permit-to-Work)',
    15: 'KPI 14 - % Compliance to Onsite Safety Induction',
    16: 'KPI 15 - % of Findings Closed On Time',
    17: 'KPI 16 - % of Inspections Completed',
    18: 'KPI 17 - % of Incident Notifications Reported On Time',
    19: 'KPI 18 - % of Investigations Completed On Time',
}

# -- Smartsheet API --

def _ss_fetch(endpoint, token):
    url = f'https://api.smartsheet.com/2.0/{endpoint}'
    req = Request(url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'})
    with urlopen(req, timeout=30, context=ssl.create_default_context()) as resp:
        return json.loads(resp.read())

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
    for m in MONTH_NAMES:
        if m.lower() == s.lower(): return m
    abbr = s[:3].lower()
    abbr_map = {m[:3].lower(): m for m in MONTH_NAMES}
    if abbr in abbr_map: return abbr_map[abbr]
    try:
        d = datetime.strptime(s, '%Y-%m-%d')
        return MONTH_NAMES[d.month - 1]
    except: pass
    try:
        d = datetime.strptime(s, '%m/%d/%Y')
        return MONTH_NAMES[d.month - 1]
    except: pass
    return None

def safe_float(v, default=0.0):
    if v is None: return default
    try: return float(v)
    except: return default

def pct_str(v):
    return f"{round(v * 100)}%"

# -- Fetch KPI data --

def fetch_kpi_data(token, month_list=None):
    """Fetch all KPI sources. month_list: list of month names to include, or None for all."""
    data = {}
    for src in SYNC_SOURCES:
        try:
            if src.get('reportId'):
                rows = fetch_report_rows(src['reportId'], token)
            else:
                rows = fetch_sheet_rows(src['sheetId'], token)
        except Exception as e:
            import traceback
            print(f"  ERROR: Failed to fetch {src['key']}: {e}")
            traceback.print_exc()
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
            if campus in ('ADC',) or str(row.get('Campus Code', '')).strip() in ('ADC',): continue

            if month_list and month_col:
                row_month = normalize_month(row.get(month_col))
                if not row_month:
                    row_month = normalize_month(row.get('Reporting Month'))
                if not row_month:
                    row_month = normalize_month(row.get('Date Reported'))
                if not row_month:
                    row_month = normalize_month(row.get('Primary'))
                if row_month not in month_list:
                    continue

            if campus not in campus_agg:
                campus_agg[campus] = {'planned': 0, 'actual': 0}

            if is_yes_no:
                pv = str(row.get(planned_col, '')).strip().lower()
                av = str(row.get(actual_col, '')).strip().lower()
                campus_agg[campus]['planned'] += 1 if pv in ('yes','true','1') else 0
                campus_agg[campus]['actual'] += 1 if av in ('yes','true','1') else 0
            elif planned_col and actual_col:
                campus_agg[campus]['planned'] += safe_float(row.get(planned_col))
                campus_agg[campus]['actual'] += safe_float(row.get(actual_col))
            elif value_col:
                v = safe_float(row.get(value_col))
                campus_agg[campus]['planned'] += v
                campus_agg[campus]['actual'] += v

        if src['key'] == 'v2_hs_committee':
            expanded = {}
            for cname, agg_val in campus_agg.items():
                if cname in COMMITTEE_MAP:
                    for code in COMMITTEE_MAP[cname]:
                        expanded[code] = dict(agg_val)
                else:
                    expanded[cname] = agg_val
            campus_agg = expanded

        weight = KPI_WEIGHTS.get(kpi_row, 0.05)
        for campus, agg in campus_agg.items():
            if campus not in data:
                data[campus] = {}
            planned = agg['planned']
            achieved = agg['actual']
            calc = min(achieved / planned, 1.0) if planned > 0 else 0.0
            data[campus][kpi_row] = {'planned': planned, 'achieved': achieved, 'calc': calc, 'weight': weight, 'na': planned == 0 and achieved == 0}
    return data

# -- Scoring --

def read_campus_data(kpi_data, sheet_name):
    """Two-level classification scoring matching Excel calc model."""
    campus = kpi_data.get(sheet_name, {})
    kpis = []
    pillar_scores = []
    for pillar in PILLAR_CLASSIFICATION:
        group_scores = []
        group_weights = []
        for grp in pillar['groups']:
            available = []
            for kpi_row, sub_weight in grp['kpis']:
                d = campus.get(kpi_row)
                if d is None:
                    kpis.append({'planned': 0, 'achieved': 0, 'calc': 0.0, 'weight': KPI_WEIGHTS.get(kpi_row, 0.05), 'na': True})
                    continue
                d.setdefault('weight', KPI_WEIGHTS.get(kpi_row, 0.05))
                kpis.append(d)
                if not d.get('na'):
                    available.append((d['calc'], sub_weight))
            if available:
                tw = sum(w for _, w in available)
                grp_score = sum(c * w for c, w in available) / tw if tw > 0 else 0
                group_scores.append(grp_score)
                group_weights.append(grp['class_weight'])
        tcw = sum(group_weights)
        p_score = sum(s * w for s, w in zip(group_scores, group_weights)) / tcw if tcw > 0 else 0
        pillar_scores.append(p_score)
    weights = [p['weight'] for p in PILLAR_CLASSIFICATION]
    overall = sum(s * w for s, w in zip(pillar_scores, weights))
    return {'sheet': sheet_name, 'kpis': kpis, 'pillar_scores': pillar_scores, 'overall': overall}

def read_region_data(kpi_data, region_cfg):
    campuses = [read_campus_data(kpi_data, s) for s in region_cfg['sheets']]
    n = len(campuses)
    avg_p = [sum(c['pillar_scores'][i] for c in campuses)/n for i in range(5)]
    avg_o = sum(c['overall'] for c in campuses)/n
    return {'campuses': campuses, 'avg_pillar': avg_p, 'avg_overall': avg_o, 'short': region_cfg['short']}

# -- Chart XML update --

def update_chart_title(xml_str, kpi_row):
    """Update chart title to match client's KPI naming convention."""
    title = KPI_CHART_TITLES.get(kpi_row)
    if not title:
        return xml_str
    safe_title = title.replace('&', '&amp;')
    title_match = re.search(r'<c:title>.*?</c:title>', xml_str, re.DOTALL)
    if not title_match:
        return xml_str
    title_block = title_match.group(0)
    # Clear all <a:t> content first, then set correct title in first run
    new_title_block = re.sub(r'(<a:t>)[^<]*(</a:t>)', r'\1\2', title_block)
    new_title_block = re.sub(r'(<a:t>)(</a:t>)', lambda m: m.group(1) + safe_title + m.group(2), new_title_block, count=1)
    return xml_str[:title_match.start()] + new_title_block + xml_str[title_match.end():]

def update_chart_xml(xml_str, c1_val, c2_val):
    """Update a 2-category clustered column chart: bar series = [c1, c2], line series = [avg, avg]."""
    avg_val = (c1_val + c2_val) / 2
    vals_bar = [c1_val, c2_val]
    vals_line = [avg_val, avg_val]

    def _update_num_vals(ser_xml, vals):
        def _replace_num(m):
            vx = m.group(0)
            vx = re.sub(r'<c:pt idx="\d+">\s*<c:v>[^<]*</c:v>\s*</c:pt>', '', vx)
            new_pts = ''.join(f'<c:pt idx="{i}"><c:v>{vals[i]}</c:v></c:pt>' for i in range(len(vals)))
            vx = vx.replace('</c:numCache>', new_pts + '</c:numCache>')
            return vx
        return re.sub(r'<c:val>.*?</c:val>', _replace_num, ser_xml, flags=re.DOTALL)

    # Update bar chart series (series 0 = bar values)
    bar_match = re.search(r'<c:barChart>.*?</c:barChart>', xml_str, re.DOTALL)
    if bar_match:
        bar_xml = bar_match.group(0)
        sers = list(re.finditer(r'<c:ser>.*?</c:ser>', bar_xml, re.DOTALL))
        for si, sm in enumerate(sers):
            old = sm.group(0)
            new = _update_num_vals(old, vals_bar if si == 0 else vals_line)
            bar_xml = bar_xml.replace(old, new, 1)
        xml_str = xml_str[:bar_match.start()] + bar_xml + xml_str[bar_match.end():]

    # Update line chart series (avg line)
    line_match = re.search(r'<c:lineChart>.*?</c:lineChart>', xml_str, re.DOTALL)
    if line_match:
        line_xml = line_match.group(0)
        sers = list(re.finditer(r'<c:ser>.*?</c:ser>', line_xml, re.DOTALL))
        for sm in sers:
            old = sm.group(0)
            new = _update_num_vals(old, vals_line)
            line_xml = line_xml.replace(old, new, 1)
        xml_str = xml_str[:line_match.start()] + line_xml + xml_str[line_match.end():]

    # Force Y-axis max to 1.0
    xml_str = _set_val_axis_max(xml_str, 1.0)
    return xml_str

def _set_val_axis_max(xml_str, max_val):
    valax_match = re.search(r'<c:valAx>(.*?)</c:valAx>', xml_str, re.DOTALL)
    if not valax_match: return xml_str
    valax = valax_match.group(0)
    scaling_match = re.search(r'<c:scaling>(.*?)</c:scaling>', valax, re.DOTALL)
    if scaling_match:
        inner = scaling_match.group(1)
        if '<c:max' in inner:
            inner = re.sub(r'<c:max val="[^"]*"/>', f'<c:max val="{max_val}"/>', inner)
        else:
            inner += f'<c:max val="{max_val}"/>'
        inner = re.sub(r'<c:min val="[^"]*"/>', '', inner)
        new_scaling = f'<c:scaling>{inner}</c:scaling>'
        new_valax = valax.replace(scaling_match.group(0), new_scaling)
    else:
        new_valax = valax.replace('<c:valAx>', f'<c:valAx><c:scaling><c:max val="{max_val}"/></c:scaling>')
    return xml_str.replace(valax, new_valax)

def remove_analysis_text(xml_str):
    """Remove Key Observations and analysis text shapes from slide content."""
    patterns_to_remove = ['Key Observations', 'Rectification', 'Key Observation',
                          'Corrective Actions']
    shapes = list(re.finditer(r'<p:sp\b[^>]*>.*?</p:sp>', xml_str, re.DOTALL))
    removals = []
    for m in shapes:
        shape_text = m.group(0)
        for pat in patterns_to_remove:
            if pat in shape_text:
                removals.append((m.start(), m.end()))
                break
    for start, end in sorted(removals, reverse=True):
        xml_str = xml_str[:start] + xml_str[end:]
    return xml_str

# -- Scoring summary slide update --

def update_scoring_slide(xml_str, region_data, short_names):
    """Update scoring summary slide: percentages, bar widths, campus names, totals."""
    campuses = region_data['campuses']
    c1 = campuses[0]
    c2 = campuses[1] if len(campuses) > 1 else campuses[0]

    # Replace campus name headers: DMC -> short_names[0], DBN -> short_names[1]
    xml_str = xml_str.replace('>DMC<', f'>{short_names[0]}<')
    xml_str = xml_str.replace('>DBN<', f'>{short_names[1]}<')

    # Replace total score labels
    xml_str = xml_str.replace('>DMC Total Score<', f'>{short_names[0]} Total Score<')
    xml_str = xml_str.replace('>DBN Total Score<', f'>{short_names[1]} Total Score<')

    # Build ordered list of percentage values to replace:
    # For each of 5 pillars: C1%, C2%, Avg%
    # Then totals: C1 total%, C2 total%, Avg total%
    new_pcts = []
    for pi in range(5):
        new_pcts.append(pct_str(c1['pillar_scores'][pi]))
        new_pcts.append(pct_str(c2['pillar_scores'][pi]))
        new_pcts.append(pct_str(region_data['avg_pillar'][pi]))
    new_pcts.append(pct_str(c1['overall']))
    new_pcts.append(pct_str(c2['overall']))
    new_pcts.append(pct_str(region_data['avg_overall']))

    # Find and replace all percentage text values
    pct_pattern = r'(<a:t>)(\d+%|N/A)(</a:t>)'
    matches = list(re.finditer(pct_pattern, xml_str))
    offset = 0
    pct_idx = 0
    for match in matches:
        if pct_idx >= len(new_pcts): break
        old_val = match.group(2)
        new_val = new_pcts[pct_idx]
        start = match.start(2) + offset
        end = match.end(2) + offset
        xml_str = xml_str[:start] + new_val + xml_str[end:]
        offset += len(new_val) - len(old_val)
        pct_idx += 1

    # Update progress bar widths
    # Pattern: pairs of shapes (bg bar + fill bar) for each pillar row per campus
    # Fill bar width = bg bar width * score percentage
    shapes = list(re.finditer(r'<p:sp>.*?</p:sp>', xml_str, re.DOTALL))

    # Identify fill bars by their position pattern:
    # For each pillar row (5 rows), there are 2 fill bars (C1, C2)
    # Fill bars are the ones with smaller cx values at the same position as bg bars
    bar_updates = []
    for pi in range(5):
        c1_score = c1['pillar_scores'][pi]
        c2_score = c2['pillar_scores'][pi]
        # Each pillar row has 7 shapes: C1_bg, C1_fill, C1_pct, C2_bg, C2_fill, C2_pct, Avg
        # Base index = pi * 7 + 1 (skip pillar label at index 0 for first row)
        # But pillar labels are at the end (indices 36-39), not interleaved
        # Actual order: shapes 0-35 are bar/pct data, 36-39 are labels, 40+ are headers/totals
        base = pi * 7
        c1_bg_idx = base + 1
        c1_fill_idx = base + 2
        c2_bg_idx = base + 4
        c2_fill_idx = base + 5

        if c1_fill_idx < len(shapes) and c1_bg_idx < len(shapes):
            bg_shape = shapes[c1_bg_idx].group(0)
            fill_shape = shapes[c1_fill_idx].group(0)
            bg_cx_m = re.search(r'<a:ext cx="(\d+)"', bg_shape)
            if bg_cx_m:
                bg_cx = int(bg_cx_m.group(1))
                new_cx = max(int(bg_cx * max(c1_score, 0.02)), 1) if c1_score > 0 else 1
                new_fill = re.sub(r'(<a:ext cx=")\d+(")', f'\\g<1>{new_cx}\\2', fill_shape, count=1)
                bar_updates.append((shapes[c1_fill_idx].start(), shapes[c1_fill_idx].end(), new_fill))

        if c2_fill_idx < len(shapes) and c2_bg_idx < len(shapes):
            bg_shape = shapes[c2_bg_idx].group(0)
            fill_shape = shapes[c2_fill_idx].group(0)
            bg_cx_m = re.search(r'<a:ext cx="(\d+)"', bg_shape)
            if bg_cx_m:
                bg_cx = int(bg_cx_m.group(1))
                new_cx = max(int(bg_cx * max(c2_score, 0.02)), 1) if c2_score > 0 else 1
                new_fill = re.sub(r'(<a:ext cx=")\d+(")', f'\\g<1>{new_cx}\\2', fill_shape, count=1)
                bar_updates.append((shapes[c2_fill_idx].start(), shapes[c2_fill_idx].end(), new_fill))

    # Apply bar width updates in reverse order
    for start, end, new_text in sorted(bar_updates, key=lambda x: x[0], reverse=True):
        xml_str = xml_str[:start] + new_text + xml_str[end:]

    return xml_str

# -- Main generation --

def generate_presentation(template_bytes, region_name, year, q1_data, q2_data):
    if region_name not in REGIONS:
        return None, f"Unknown region: {region_name}"

    region_cfg = REGIONS[region_name]
    short_names = region_cfg['short']
    campus_codes = region_cfg['sheets']

    q1_region = read_region_data(q1_data, region_cfg)
    q2_region = read_region_data(q2_data, region_cfg)

    # Read template ZIP
    buf_in = io.BytesIO(template_bytes)
    file_contents = {}
    with zipfile.ZipFile(buf_in, 'r') as zin:
        for item in zin.infolist():
            file_contents[item.filename] = zin.read(item.filename)

    # 1. Update cover slide (slide 1)
    slide1_path = 'ppt/slides/slide1.xml'
    if slide1_path in file_contents:
        xml = file_contents[slide1_path].decode('utf-8')
        date_str = datetime.now().strftime('%d-%m-%Y')
        # Replace date patterns (long format and DD-MM-YYYY)
        xml = re.sub(r'(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),\s+\w+\s+\d+,\s+\d{4}', date_str, xml)
        xml = re.sub(r'\d{2}-\d{2}-\d{4}', date_str, xml)
        # Replace subtitle with correct region
        subtitle_safe = region_cfg['subtitle'].replace('&', '&amp;')
        xml = re.sub(r'>Dubai Academic City[^<]*Al Nahda<', f'>{subtitle_safe}<', xml)
        xml = xml.replace('>DMC<', f'>{short_names[0]}<')
        xml = xml.replace('>DBN<', f'>{short_names[1]}<')
        file_contents[slide1_path] = xml.encode('utf-8')

    # 2. Update Q1 charts (slides 4-11)
    for chart_file, kpi_row in Q1_CHART_MAP.items():
        path = f"ppt/charts/{chart_file}"
        if path not in file_contents: continue
        xml_str = file_contents[path].decode('utf-8')
        c1_val = q1_data.get(campus_codes[0], {}).get(kpi_row, {}).get('calc', 0.0)
        c2_val = q1_data.get(campus_codes[1], {}).get(kpi_row, {}).get('calc', 0.0) if len(campus_codes) > 1 else 0.0
        new_xml = update_chart_xml(xml_str, c1_val, c2_val)
        new_xml = update_chart_title(new_xml, kpi_row)
        file_contents[path] = new_xml.encode('utf-8')

    # 3. Q2 charts REMOVED per client request (slides 12-21 deleted)
    # for chart_file, kpi_row in Q2_CHART_MAP.items():
        # path = f"ppt/charts/{chart_file}"
        # if path not in file_contents: continue
        # xml_str = file_contents[path].decode('utf-8')
        # c1_val = q2_data.get(campus_codes[0], {}).get(kpi_row, {}).get('calc', 0.0)
        # c2_val = q2_data.get(campus_codes[1], {}).get(kpi_row, {}).get('calc', 0.0) if len(campus_codes) > 1 else 0.0
        # new_xml = update_chart_xml(xml_str, c1_val, c2_val)
        # new_xml = update_chart_title(new_xml, kpi_row)
        # file_contents[path] = new_xml.encode('utf-8')

    # 4. Update Q1 scoring summary (slide 3)
    slide3_path = 'ppt/slides/slide3.xml'
    if slide3_path in file_contents:
        xml = file_contents[slide3_path].decode('utf-8')
        xml = update_scoring_slide(xml, q1_region, short_names)
        # Update quarter label in title
        xml = xml.replace('>Q1<', f'>Q1<')
        file_contents[slide3_path] = xml.encode('utf-8')

    # 5. Q2 scoring summary REMOVED per client request
    # slide13_path = 'ppt/slides/slide13.xml'
    # if slide13_path in file_contents:
    #     xml = file_contents[slide13_path].decode('utf-8')
    #     xml = update_scoring_slide(xml, q2_region, short_names)
    #     file_contents[slide13_path] = xml.encode('utf-8')

    # 6. Replace campus codes in ALL slides (DMC->short[0], DBN->short[1])
    for fname in list(file_contents.keys()):
        if fname.startswith('ppt/slides/slide') and fname.endswith('.xml'):
            xml = file_contents[fname].decode('utf-8')
            changed = False
            # Replace in chart category labels and text shapes
            if '>DMC<' in xml:
                xml = xml.replace('>DMC<', f'>{short_names[0]}<')
                changed = True
            if '>DBN<' in xml:
                xml = xml.replace('>DBN<', f'>{short_names[1]}<')
                changed = True
            if changed:
                file_contents[fname] = xml.encode('utf-8')

    # Also replace campus names in chart category caches
    for fname in list(file_contents.keys()):
        if fname.startswith('ppt/charts/chart') and fname.endswith('.xml') and fname.split('/')[-1] not in PIE_CHARTS:
            xml = file_contents[fname].decode('utf-8')
            changed = False
            if '>DMC<' in xml:
                xml = xml.replace('>DMC<', f'>{campus_codes[0]}<')
                changed = True
            if '>DBN<' in xml:
                xml = xml.replace('>DBN<', f'>{campus_codes[1]}<')
                changed = True
            if changed:
                file_contents[fname] = xml.encode('utf-8')

    # 7. Update NA box text with correct KPI names (keep "No scoring" text)
    for fname in list(file_contents.keys()):
        if not (fname.startswith('ppt/slides/slide') and fname.endswith('.xml')): continue
        xml = file_contents[fname].decode('utf-8')
        if 'No scoring' not in xml: continue
        # NA boxes contain "KPI X - Name\nNo scoring for this quarter"
        # Just ensure they display correctly - no data changes needed
        file_contents[fname] = xml.encode('utf-8')

    # 8. Force Y-axis max on all bar charts
    for fname in list(file_contents.keys()):
        if fname.startswith('ppt/charts/chart') and fname.endswith('.xml'):
            cname = fname.split('/')[-1]
            if cname in PIE_CHARTS: continue
            xml = file_contents[fname].decode('utf-8') if isinstance(file_contents[fname], bytes) else file_contents[fname]
            xml = _set_val_axis_max(xml, 1.0)
            file_contents[fname] = xml.encode('utf-8')

    # 9. Remove analysis text from content slides (per client: keep titles only)
    for slide_num in range(4, 12):
        slide_path = f'ppt/slides/slide{slide_num}.xml'
        if slide_path in file_contents:
            xml = file_contents[slide_path].decode('utf-8')
            xml = remove_analysis_text(xml)
            file_contents[slide_path] = xml.encode('utf-8')

    # 10. Remove Q2 duplicate slides (12-21) per client request
    pres_path = 'ppt/presentation.xml'
    rels_path_pres = 'ppt/_rels/presentation.xml.rels'
    if pres_path in file_contents and rels_path_pres in file_contents:
        pres_xml = file_contents[pres_path].decode('utf-8')
        rels_xml = file_contents[rels_path_pres].decode('utf-8')
        for slide_num in range(12, 22):
            # Find rId for this slide
            rid_match = re.search(rf'Id="(rId\d+)"[^>]*slides/slide{slide_num}\.xml', rels_xml)
            if rid_match:
                rid = rid_match.group(1)
                pres_xml = re.sub(rf'<p:sldId[^>]*{rid}[^/]*/>', '', pres_xml)
                rels_xml = re.sub(rf'<Relationship[^>]*{rid}[^>]*slides/slide{slide_num}\.xml[^/]*/>', '', rels_xml)
        file_contents[pres_path] = pres_xml.encode('utf-8')
        file_contents[rels_path_pres] = rels_xml.encode('utf-8')

    # Delete Q2 slide files and rels
    for slide_num in range(12, 22):
        for path in [f'ppt/slides/slide{slide_num}.xml', f'ppt/slides/_rels/slide{slide_num}.xml.rels']:
            if path in file_contents:
                del file_contents[path]
    # Delete Q2 chart files
    for chart_file in Q2_CHART_MAP:
        path = f'ppt/charts/{chart_file}'
        if path in file_contents:
            del file_contents[path]
    # Delete Q2 pie chart
    pie2_path = 'ppt/charts/chart15.xml'
    if pie2_path in file_contents:
        del file_contents[pie2_path]

        # Write output ZIP
    buf_out = io.BytesIO()
    with zipfile.ZipFile(buf_out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fname, data in file_contents.items():
            zout.writestr(fname, data)

    return buf_out.getvalue(), None

# -- HTTP Handler --

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)

        region = params.get('region', ['Abu Dhabi'])[0]
        year = params.get('year', [str(datetime.now().year)])[0]
        period = params.get('period', ['quarter'])[0]  # month, quarter, annual
        month = params.get('month', [None])[0]  # specific month name for monthly reports

        token = os.environ.get('SMARTSHEET_TOKEN')
        if not token:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'SMARTSHEET_TOKEN not set'}).encode())
            return

        if region.lower() == 'all':
            region = 'Abu Dhabi'

        if region not in REGIONS:
            self.send_response(400)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': f'Invalid region. Available: {list(REGIONS.keys())}'}).encode())
            return

        try:
            # Determine month lists based on period
            if period == 'annual':
                q1_months = ['January', 'February', 'March', 'April', 'May', 'June']
                q2_months = ['July', 'August', 'September', 'October', 'November', 'December']
            elif period == 'month' and month:
                q1_months = [month]
                q2_months = []
            else:
                q1_months = Q1_MONTHS
                q2_months = Q2_MONTHS

            # Fetch KPI data
            q1_data = fetch_kpi_data(token, q1_months)
            q2_data = fetch_kpi_data(token, q2_months) if q2_months else {}

            # Load template
            base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            template_path = os.path.join(base, 'templates', 'template.pptx')
            if not os.path.exists(template_path):
                template_path = os.path.join(os.getcwd(), 'templates', 'template.pptx')
            if not os.path.exists(template_path):
                template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates', 'template.pptx')
            with open(template_path, 'rb') as f:
                template_bytes = f.read()

            pptx_bytes, error = generate_presentation(template_bytes, region, year, q1_data, q2_data)

            if error:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': error}).encode())
                return

            safe_name = region.replace(' ', '_')
                    period_label = month.capitalize() if period == 'month' and month else ('Annual' if period == 'annual' else 'Q1_Q2')
                    filename = f"HCT_KPI_{safe_name}_{period_label}_{year}.pptx"

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(pptx_bytes)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(pptx_bytes)

        except Exception as e:
            import traceback
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e), 'trace': traceback.format_exc()}).encode())

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
