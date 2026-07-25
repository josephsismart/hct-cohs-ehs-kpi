"""Vercel Python serverless function â HCT-COHS KPI Word Report Generator.
Uses template-based approach: unzip template, replace chart data via regex, rezip.
"""

import os, io, json, re, zipfile
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen

# ââ Regions ââ
REGIONS = {
    'AD Al Ain':       {'sheets': ['AAF','AAZ'], 'short': ['Falaj Hazza','Zakhir'],        'subtitle': 'Al Ain Falaj Hazza & Al Ain Zakhir'},
    'Abu Dhabi':       {'sheets': ['ADA','ADB'], 'short': ['Baniyas A','Baniyas B'],       'subtitle': 'Abu Dhabi Baniyas A & Abu Dhabi Baniyas B'},
    'AD Remote':       {'sheets': ['ADH','MZY'], 'short': ['Al Dhanna','Madinat Zayed'],   'subtitle': 'Al Dhanna Ruwais & Al Dhafra Madinat Zayed City'},
    'Dubai':           {'sheets': ['DMC','DBN'], 'short': ['Academic City','Al Nahda'],     'subtitle': 'Dubai Academic City & Dubai Al Nahda'},
    'Fujairah':        {'sheets': ['FJF','FJH'], 'short': ['Faseel','Hulaifat'],            'subtitle': 'Fujairah Faseel & Fujairah Hulaifat'},
    'Sharjah':         {'sheets': ['SJA','SJB'], 'short': ['Campus A','Campus B'],          'subtitle': 'Sharjah Campus A & Sharjah Campus B'},
    'Ras Al Khaimah':  {'sheets': ['RKA','RKB'], 'short': ['Campus A','Campus B'],          'subtitle': 'RAK Campus A & RAK Campus B'},
}

# Campus order used in most charts (14 campuses)
CAMPUS_ORDER_14 = ['ADA','ADB','AAF','AAZ','DMC','DBN','SJA','SJB','FJF','FJH','RKA','RKB','ADH','MZY']
# Campus order for chart12/13 (15 campuses, includes HQ at index 1)
CAMPUS_ORDER_15 = ['ADA','HQ','ADB','AAF','AAZ','DMC','DBN','SJA','SJB','FJF','FJH','RKA','RKB','ADH','MZY']
# Region order for chart2/3
REGION_ORDER = ['Abu Dhabi', 'AD Al Ain', 'Dubai', 'Sharjah', 'Fujairah', 'Ras Al Khaimah', 'AD Remote']
# Region name mapping (chart labels â REGIONS keys)
REGION_LABEL_MAP = {
    'Abu Dhabi Main': 'Abu Dhabi', 'Abu Dhabi': 'Abu Dhabi',
    'Al Ain': 'AD Al Ain',
    'Dubai': 'Dubai',
    'Sharjah': 'Sharjah',
    'Fujairah': 'Fujairah',
    'Ras Al Khaimah': 'Ras Al Khaimah', 'RAK': 'Ras Al Khaimah',
    'Al Dhafra': 'AD Remote', 'Al Dhanna': 'AD Remote',
}

# ââ KPI weights ââ
KPI_WEIGHTS = {
    2: 0.30, 3: 0.10, 4: 0.10, 5: 0.25, 6: 0.25,
    7: 0.30, 8: 0.50, 9: 0.20,
    10: 0.50, 11: 0.50,
    12: 0.40, 13: 0.40, 14: 0.10, 15: 0.10,
    16: 0.30, 17: 0.30, 18: 0.20, 19: 0.20,
}

MONTH_NAMES = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']

# ââ Smartsheet sources ââ
SYNC_SOURCES = [
    {'key': 'v2_hs_kpi_report', 'reportId': '810227329879940', 'campusCol': 'Campus', 'monthCol': 'Reporting Month', 'valueCol': 'Submitted', 'kpi_row': 2},
    {'key': 'v2_external_compliance', 'sheetId': '4198632256393092', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'Applicable Compliance', 'actualCol': 'Actual Compliance', 'kpi_row': 4},
    {'key': 'v2_hs_committee', 'sheetId': '435993944477572', 'campusCol': 'Committee', 'monthCol': 'Reporting Month', 'plannedCol': 'Meeting Planned', 'actualCol': 'Meeting Conducted', 'kpi_row': 5},
    {'key': 'v2_hazard_id', 'sheetId': '7323092115214212', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Controls Identified', 'actualCol': 'Implemented Controls', 'kpi_row': 6},
    {'key': 'v2_risk_closed', 'sheetId': '7323092115214212', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'Total Risk Assessments Registered', 'actualCol': 'Risk Assessment Closed', 'kpi_row': 7},
    {'key': 'v2_risk_validated', 'sheetId': '7323092115214212', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'Total Assessments Register', 'actualCol': 'RA Validated and Signed Off', 'kpi_row': 8},
    {'key': 'v2_safe_working', 'sheetId': '1693592581001092', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of SOPs Verified', 'actualCol': 'No. of SOPs Implemented', 'kpi_row': 9},
    {'key': 'v2_planned_training', 'sheetId': '8549734774951812', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned (Yes/No)', 'actualCol': 'Are there any submission?', 'kpi_row': 10, 'yesNoCount': True},
    {'key': 'v2_drills', 'sheetId': '5053158949605252', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Planned Drill? (Yes/No)', 'actualCol': 'Are there any submission?', 'kpi_row': 13, 'yesNoCount': True},
    {'key': 'v2_permit_to_work', 'sheetId': '5899016251330436', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'No. of PTWs Issued', 'actualCol': 'Total Work Registered', 'kpi_row': 14},
    {'key': 'v2_onsite_induction', 'sheetId': '5899016251330436', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': "No. of New Contractors (Individuals)", 'actualCol': 'Contractors Inducted in the Reporting Month', 'kpi_row': 15},
    {'key': 'v2_ehs_inspection', 'sheetId': '4947401822392196', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of EHS Inspections Planned', 'actualCol': 'No. of EHS Inspections Completed', 'kpi_row': 16},
    {'key': 'v2_findings_on_time', 'sheetId': '4947401822392196', 'campusCol': 'Campus Code', 'monthCol': 'Primary', 'plannedCol': 'No. of Findings in Reporting Month', 'actualCol': 'No. of Findings Due', 'kpi_row': 17},
    {'key': 'v2_investigation_on_time', 'reportId': '6831846506581892', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident Investigated', 'actualCol': 'Investigation Completed on Time', 'kpi_row': 18},
    {'key': 'v2_notification', 'reportId': '1199821531598724', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'plannedCol': 'Total Incident', 'actualCol': 'Notification Submitted on Time', 'kpi_row': 19},
]

TRAINING_SOURCE = {'sheetId': '8549734774951812', 'campusCol': 'Campus Code', 'monthCol': 'Reporting Month', 'hoursCol': 'Total Hours'}


# ââ Smartsheet API ââ

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


# ââ Fetch KPI data ââ

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
    result = {}
    for row in rows:
        campus = str(row.get(TRAINING_SOURCE['campusCol'], '')).strip()
        month = normalize_month(row.get(TRAINING_SOURCE['monthCol']))
        if not campus: continue
        if month_filter and month != month_filter: continue
        hours = safe_float(row.get(TRAINING_SOURCE['hoursCol']))
        result[campus] = result.get(campus, 0) + hours
    return result


# ââ Chart data replacement (regex-based, preserves namespace prefixes) ââ

def replace_numcache_values(chart_xml, new_series_values):
    """Replace numCache values in chart XML. Namespace-agnostic (works with c:, ns0:, etc)."""
    # Detect the namespace prefix used (e.g. 'c:' or 'ns0:')
    ns_match = re.search(r'<(\w+):numCache>', chart_xml)
    if not ns_match:
        return chart_xml  # No numCache found
    ns = ns_match.group(1)  # e.g. 'c' or 'ns0'

    def replace_one(nc_text, new_vals):
        nc = nc_text
        nc = re.sub(rf'(<{ns}:ptCount val=")\d+(")', lambda m: m.group(1) + str(len(new_vals)) + m.group(2), nc)
        nc = re.sub(rf'<{ns}:pt\s+idx="\d+">\s*<{ns}:v>[^<]*</{ns}:v>\s*</{ns}:pt>', '', nc)
        pts = ''.join(f'<{ns}:pt idx="{i}"><{ns}:v>{v}</{ns}:v></{ns}:pt>' for i, v in enumerate(new_vals))
        nc = nc.replace(f'</{ns}:numCache>', pts + f'</{ns}:numCache>')
        return nc

    pattern = re.compile(rf'<{ns}:numCache>.*?</{ns}:numCache>', re.DOTALL)
    matches = list(pattern.finditer(chart_xml))

    result = chart_xml
    for i in range(min(len(matches), len(new_series_values)) - 1, -1, -1):
        m = matches[i]
        result = result[:m.start()] + replace_one(m.group(), new_series_values[i]) + result[m.end():]

    return result


# ââ Build chart data from KPI data ââ

def get_campus_val(kpi_data, campus, kpi_row, field='calc'):
    return kpi_data.get(campus, {}).get(kpi_row, {}).get(field, 0)

def region_agg(kpi_data, kpi_row, region_name, field):
    cfg = REGIONS.get(region_name)
    if not cfg: return 0
    vals = [get_campus_val(kpi_data, s, kpi_row, field) for s in cfg['sheets']]
    return sum(vals)

def build_chart_data(kpi_data, training_hours):
    """Build replacement data for all 13 charts."""
    charts = {}

    # Helper: percentage per campus (14 campuses)
    def pct_14(kpi_row):
        return [[get_campus_val(kpi_data, c, kpi_row, 'calc') for c in CAMPUS_ORDER_14]]

    # Helper: percentage per campus (15 campuses, HQ=0 at idx 1)
    def pct_15(kpi_row):
        vals = []
        for c in CAMPUS_ORDER_15:
            if c == 'HQ':
                vals.append(0)
            else:
                vals.append(get_campus_val(kpi_data, c, kpi_row, 'calc'))
        return [vals]

    # Helper: planned + actual per campus (14), 2 series
    def planned_actual_14(kpi_row):
        planned = [get_campus_val(kpi_data, c, kpi_row, 'planned') for c in CAMPUS_ORDER_14]
        actual = [get_campus_val(kpi_data, c, kpi_row, 'achieved') for c in CAMPUS_ORDER_14]
        return [planned, actual]

    # chart1: KPI Report % (14 campuses, 1 series)
    charts[1] = pct_14(2)

    # chart2: Committee meeting (7 regions, 2 series: planned, conducted)
    planned = []
    conducted = []
    for rname in REGION_ORDER:
        p = 0
        a = 0
        # Committee data uses region names as campus key
        for alias, rkey in REGION_LABEL_MAP.items():
            if rkey == rname:
                p += get_campus_val(kpi_data, alias, 5, 'planned')
                a += get_campus_val(kpi_data, alias, 5, 'achieved')
        # Also check direct region name
        p += get_campus_val(kpi_data, rname, 5, 'planned')
        a += get_campus_val(kpi_data, rname, 5, 'achieved')
        planned.append(round(p))
        conducted.append(round(a))
    charts[2] = [planned, conducted]

    # chart3: Incident investigation by region (7 regions, 3 series: total, closed, % close-out)
    totals = []
    closed = []
    pct_close = []
    for rname in REGION_ORDER:
        t = region_agg(kpi_data, 18, rname, 'planned')
        c = region_agg(kpi_data, 18, rname, 'achieved')
        totals.append(round(t))
        closed.append(round(c))
        pct_close.append(round(c / t, 4) if t > 0 else 0)
    charts[3] = [totals, closed, pct_close]

    # chart4: External compliance % (14 campuses)
    charts[4] = pct_14(4)

    # chart5: Training hours (14 campuses, 1 series â actual hours)
    charts[5] = [[training_hours.get(c, 0) for c in CAMPUS_ORDER_14]]

    # chart6: Hazard ID % (14 campuses)
    charts[6] = pct_14(6)

    # chart7: SWP % (7 campuses: ADA, ADB, FJF, FJH, RKA, RKB, ADH)
    chart7_campuses = ['ADA','ADB','FJF','FJH','RKA','RKB','ADH']
    charts[7] = [[get_campus_val(kpi_data, c, 9, 'calc') for c in chart7_campuses]]

    # chart8: PTW (14 campuses, 2 series)
    charts[8] = planned_actual_14(14)

    # chart9: Onsite induction (14 campuses, 2 series)
    charts[9] = planned_actual_14(15)

    # chart10: EHS inspection (14 campuses, 2 series)
    charts[10] = planned_actual_14(16)

    # chart11: Findings % (14 campuses)
    charts[11] = pct_14(17)

    # chart12: Investigation % (15 campuses incl HQ)
    charts[12] = pct_15(18)

    # chart13: Notification % (15 campuses incl HQ)
    charts[13] = pct_15(19)

    return charts


# ââ Generate report using template ââ

def generate_report(month_name, year, token):
    if not month_name:
        month_name = None
    period_label = f'{month_name} {year}' if month_name else f'Overall {year}'
    print(f'Generating Word report (overall) for {period_label}')

    # Fetch data
    kpi_data = fetch_kpi_data(token, month_name)
    training_hours = fetch_training_hours(token, month_name)
    print(f'  KPI data: {len(kpi_data)} campuses')
    print(f'  Training hours: {len(training_hours)} campuses')

    # Build chart replacement data
    chart_data = build_chart_data(kpi_data, training_hours)

    # Load template
    tmpl_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates', 'word_template.docx')
    with open(tmpl_path, 'rb') as f:
        tmpl_bytes = f.read()

    # Unzip, modify charts, rezip
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(tmpl_bytes), 'r') as zin:
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                # Replace chart data
                chart_match = re.match(r'word/charts/chart(\d+)\.xml', item.filename)
                if chart_match:
                    chart_num = int(chart_match.group(1))
                    if chart_num in chart_data:
                        xml_str = data.decode('utf-8')
                        xml_str = replace_numcache_values(xml_str, chart_data[chart_num])
                        data = xml_str.encode('utf-8')
                        print(f'  Updated chart{chart_num} with {len(chart_data[chart_num])} series')

                zout.writestr(item, data)

    buf.seek(0)
    return buf.getvalue()


# ââ HTTP Handler ââ

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        month = qs.get('month', [None])[0]
        month = month if month else None
        year = qs.get('year', ['2026'])[0]
        report_name = qs.get('reportName', ['KPI_Report'])[0]


        token = os.environ.get('SMARTSHEET_TOKEN', '')
        if not token:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'SMARTSHEET_TOKEN not set'}).encode())
            return

        try:
            docx_bytes = generate_report(month, year, token)
            filename = f'{report_name.replace(" ", "_")}_{month}_{year}.docx' if month else f'{report_name.replace(" ", "_")}_Overall_{year}.docx'
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
