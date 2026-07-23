"""Vercel Python serverless — HCT Word Report Generator (template-based).
Unzips the client's template docx, replaces chart data with live Smartsheet
values, updates text placeholders, and rezips.
"""

import os, io, json, zipfile, re
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET

# ── Namespaces ──
C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

ALL_CAMPUSES = ['ADA','ADB','AAF','AAZ','DMC','DBN','SJA','SJB','FJF','FJH','RKA','RKB','ADH','MZY']
ALL_CAMPUSES_WITH_HQ = ALL_CAMPUSES + ['HQ']

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

KPI_WEIGHTS = {
    2: 0.30, 3: 0.10, 4: 0.10, 5: 0.25, 6: 0.25,
    7: 0.30, 8: 0.50, 9: 0.20,
    10: 0.50, 11: 0.50,
    12: 0.40, 13: 0.40, 14: 0.10, 15: 0.10,
    16: 0.30, 17: 0.30, 18: 0.20, 19: 0.20,
}

SYNC_SOURCES = [
    {'key': 'v2_kpi_report', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned Submission', 'actualCol': 'Actual Submission', 'kpi_row': 2},
    {'key': 'v2_training_hours', 'sheetId': '8549734774951812', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned Training Hours', 'actualCol': 'Total Training Hours Delivered', 'kpi_row': 3},
    {'key': 'v2_ext_authority', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Regulatory Requirements', 'actualCol': 'Complied Requirements', 'kpi_row': 4},
    {'key': 'v2_committee_meeting', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Committee Meeting Planned', 'actualCol': 'Committee Meeting Conducted', 'kpi_row': 5},
    {'key': 'v2_hazard_id', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Committee Actions', 'actualCol': 'Committee Actions Closed', 'kpi_row': 6},
    {'key': 'v2_risk_closed', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Risk Assessments', 'actualCol': 'Risk Assessments Closed', 'kpi_row': 7},
    {'key': 'v2_risk_validated', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Risk Assessments', 'actualCol': 'Risk Assessments Validated', 'kpi_row': 8},
    {'key': 'v2_swp', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total SWPs Due', 'actualCol': 'SWPs Completed', 'kpi_row': 9},
    {'key': 'v2_training_plan', 'sheetId': '8549734774951812', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned Training Hours', 'actualCol': 'Total Training Hours Delivered', 'kpi_row': 10},
    {'key': 'v2_awareness', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Awareness Campaigns Planned', 'actualCol': 'Awareness Campaigns Conducted', 'kpi_row': 11},
    {'key': 'v2_compliance_activity', 'sheetId': '5899016251330436', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Compliance Activities Sampled', 'actualCol': 'Compliance Activities Implemented', 'kpi_row': 12},
    {'key': 'v2_emergency_drill', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned Drill? (Yes/No)', 'actualCol': 'Are there any submission?', 'kpi_row': 13, 'yesNoCount': True},
    {'key': 'v2_permit_to_work', 'sheetId': '5899016251330436', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'No. of PTWs Issued', 'actualCol': 'Total Work Registered', 'kpi_row': 14},
    {'key': 'v2_onsite_induction', 'sheetId': '5899016251330436', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': "No. of New Contractors (Individuals)", 'actualCol': 'Contractors Inducted in the Reporting Month', 'kpi_row': 15},
    {'key': 'v2_ehs_inspection', 'sheetId': '4947401822392196', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of EHS Inspections Planned', 'actualCol': 'No. of EHS Inspections Completed', 'kpi_row': 16},
    {'key': 'v2_findings_on_time', 'sheetId': '4947401822392196', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of Findings in Reporting Month', 'actualCol': 'No. of Findings Due', 'kpi_row': 17},
    {'key': 'v2_investigation_on_time', 'reportId': '6831846506581892', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident Investigated', 'actualCol': 'Investigation Completed on Time', 'kpi_row': 18},
    {'key': 'notification', 'reportId': '1199821531598724', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident', 'actualCol': 'Notification Submitted on Time', 'kpi_row': 19},
]

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


# ── Data fetching ──

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
        is_yes_no = src.get('yesNoCount', False)

        campus_agg = {}
        for row in rows:
            campus = str(row.get(campus_col, '')).strip()
            if not campus: continue
            if month_col and month_filter:
                rm = normalize_month(row.get(month_col))
                if rm != month_filter:
                    rm = normalize_month(row.get('Reporting Month'))
                    if rm != month_filter:
                        rm = normalize_month(row.get('Date Reported'))
                        if rm != month_filter:
                            rm = normalize_month(row.get('Primary'))
                            if rm != month_filter:
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

        for campus, agg in campus_agg.items():
            if campus not in data:
                data[campus] = {}
            planned = agg['planned']
            achieved = agg['actual']
            calc = min(achieved / planned, 1.0) if planned > 0 else (1.0 if achieved > 0 else 0)
            data[campus][kpi_row] = {'planned': planned, 'achieved': achieved, 'calc': calc}

    return data

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


# ── Chart data replacement ──

def replace_chart_values(chart_xml_bytes, new_series_data):
    """Replace numCache values in a chart XML.
    new_series_data: list of dicts with 'values' (list of numbers).
    """
    tree = ET.ElementTree(ET.fromstring(chart_xml_bytes))
    root = tree.getroot()
    C = C_NS

    # Find all series
    sers = root.findall(f'.//{{{C}}}ser')

    for si, ser in enumerate(sers):
        if si >= len(new_series_data):
            break

        new_vals = new_series_data[si].get('values', [])

        # Replace numCache values
        num_cache = ser.find(f'.//{{{C}}}val/{{{C}}}numRef/{{{C}}}numCache')
        if num_cache is not None:
            # Update ptCount
            pt_count = num_cache.find(f'{{{C}}}ptCount')
            if pt_count is not None:
                pt_count.set('val', str(len(new_vals)))

            # Remove existing pts
            for pt in num_cache.findall(f'{{{C}}}pt'):
                num_cache.remove(pt)

            # Add new pts
            for idx, val in enumerate(new_vals):
                pt = ET.SubElement(num_cache, f'{{{C}}}pt')
                pt.set('idx', str(idx))
                v = ET.SubElement(pt, f'{{{C}}}v')
                v.text = str(val)

    buf = io.BytesIO()
    tree.write(buf, xml_declaration=True, encoding='UTF-8')
    return buf.getvalue()


def build_chart_data(kpi_data, training_hours):
    """Build replacement data for all 13 charts."""

    def campus_pct(kpi_row, campuses=ALL_CAMPUSES):
        return [round(kpi_data.get(c, {}).get(kpi_row, {}).get('calc', 0) * 100) for c in campuses]

    def campus_val(kpi_row, field, campuses=ALL_CAMPUSES):
        return [round(safe_float(kpi_data.get(c, {}).get(kpi_row, {}).get(field, 0))) for c in campuses]

    def region_agg(kpi_row, field):
        vals = []
        for rname in REGION_ORDER:
            total = sum(safe_float(kpi_data.get(c, {}).get(kpi_row, {}).get(field, 0)) for c in REGION_GROUPS[rname])
            vals.append(round(total))
        return vals

    charts = {}

    # Chart 1: External Authority Compliance (%) - 14 campuses, 1 series
    charts[1] = [{'values': campus_pct(4)}]

    # Chart 2: Committee Meetings - 7 regions, 2 series (planned, conducted)
    charts[2] = [
        {'values': region_agg(5, 'planned')},
        {'values': region_agg(5, 'achieved')},
    ]

    # Chart 3: Committee Actions Closed - 7 regions, 3 series (total, closed, %)
    total_actions = region_agg(6, 'planned')
    closed_actions = region_agg(6, 'achieved')
    pct_closed = [round(c/t*100) if t > 0 else 0 for t, c in zip(total_actions, closed_actions)]
    charts[3] = [
        {'values': total_actions},
        {'values': closed_actions},
        {'values': pct_closed},
    ]

    # Chart 4: Risk Control Measures (%) - 14 campuses, 1 series
    charts[4] = [{'values': campus_pct(7)}]

    # Chart 5: Training Hours - 14 campuses, 1 series (actual hours)
    charts[5] = [{'values': [round(training_hours.get(c, 0)) for c in ALL_CAMPUSES]}]

    # Chart 6: Compliance Activities (%) - 14 campuses, 1 series
    charts[6] = [{'values': campus_pct(12)}]

    # Chart 7: Emergency Drills (%) - varies (client has 7 cats)
    charts[7] = [{'values': campus_pct(13)}]

    # Chart 8: PTW - 14 campuses, 2 series (total work, PTWs issued)
    charts[8] = [
        {'values': campus_val(14, 'achieved')},
        {'values': campus_val(14, 'planned')},
    ]

    # Chart 9: Contractors - 14 campuses, 2 series
    charts[9] = [
        {'values': campus_val(15, 'planned')},
        {'values': campus_val(15, 'achieved')},
    ]

    # Chart 10: Inspections - 14 campuses, 2 series
    charts[10] = [
        {'values': campus_val(16, 'planned')},
        {'values': campus_val(16, 'achieved')},
    ]

    # Chart 11: Findings Closed (%) - 14 campuses, 1 series
    charts[11] = [{'values': campus_pct(17)}]

    # Chart 12: Incident Notifications (%) - 15 campuses with HQ, 1 series
    charts[12] = [{'values': campus_pct(19, ALL_CAMPUSES_WITH_HQ)}]

    # Chart 13: Investigations (%) - 15 campuses with HQ, 1 series
    charts[13] = [{'values': campus_pct(18, ALL_CAMPUSES_WITH_HQ)}]

    return charts


def update_document_text(doc_xml_bytes, month_name, year):
    """Replace month/date placeholders in document.xml."""
    text = doc_xml_bytes.decode('utf-8')

    # Replace "March" with actual month in common patterns
    text = text.replace('March 2026', f'{month_name} {year}')
    text = text.replace('March', month_name)

    # Replace "April 2026" in planned actions section
    month_idx = MONTH_NAMES.index(month_name) if month_name in MONTH_NAMES else 0
    next_month = MONTH_NAMES[(month_idx + 1) % 12]
    next_year = year if month_idx < 11 else str(int(year) + 1)
    text = text.replace('April 2026', f'{next_month} {next_year}')

    return text.encode('utf-8')


# ── Main generator ──

def generate_report(month_name, year, token):
    """Generate Word report by modifying template with live data."""
    print(f'Generating Word report — {month_name} {year}')

    # Load template
    template_path = os.path.join(os.path.dirname(__file__), '..', 'templates', 'word_template.docx')
    if not os.path.exists(template_path):
        template_path = os.path.join(os.path.dirname(__file__), 'word_template.docx')
    if not os.path.exists(template_path):
        # Try relative to current file
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates', 'word_template.docx')

    with open(template_path, 'rb') as f:
        template_bytes = f.read()

    # Fetch live data
    kpi_data = fetch_kpi_data(token, month_name)
    training_hours = fetch_training_hours(token, month_name)
    chart_data = build_chart_data(kpi_data, training_hours)

    # Modify template
    in_buf = io.BytesIO(template_bytes)
    out_buf = io.BytesIO()

    with zipfile.ZipFile(in_buf, 'r') as zin, zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)

            # Replace chart data
            m = re.match(r'word/charts/chart(\d+)\.xml$', item.filename)
            if m:
                chart_num = int(m.group(1))
                if chart_num in chart_data:
                    try:
                        data = replace_chart_values(data, chart_data[chart_num])
                    except Exception as e:
                        print(f'  WARNING: Failed to update chart {chart_num}: {e}')

            # Replace document text (month names)
            if item.filename == 'word/document.xml':
                data = update_document_text(data, month_name, year)

            zout.writestr(item, data)

    out_buf.seek(0)
    return out_buf.getvalue()


# ── HTTP Handler ──

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        month = qs.get('month', [''])[0]
        year = qs.get('year', ['2026'])[0]
        report_name = qs.get('name', [''])[0]

        # If no month specified, use current month
        if not month:
            from datetime import datetime
            month = MONTH_NAMES[datetime.now().month - 1]

        token = os.environ.get('SMARTSHEET_TOKEN', '')
        if not token:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'SMARTSHEET_TOKEN not set'}).encode())
            return

        try:
            docx_bytes = generate_report(month, year, token)
            filename = f'{report_name}.docx' if report_name else f'HCT_COHS_KPI_Report_{month}_{year}.docx'

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
