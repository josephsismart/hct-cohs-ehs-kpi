"""Vercel Python serverless function Ã¢ÂÂ HCT-COHS KPI Word Report Generator.
Uses template-based approach: unzip template, replace chart data via regex, rezip.
"""

import os, io, json, re, zipfile
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen

# Ã¢ÂÂÃ¢ÂÂ Regions Ã¢ÂÂÃ¢ÂÂ
REGIONS = {
'AD Al Ain': {'sheets': ['AAF','AAZ'], 'short': ['Falaj Hazza','Zakhir'], 'subtitle': 'Al Ain Falaj Hazza & Al Ain Zakhir'},
'Abu Dhabi': {'sheets': ['ADA','ADB'], 'short': ['Baniyas A','Baniyas B'], 'subtitle': 'Abu Dhabi Baniyas A & Abu Dhabi Baniyas B'},
'AD Remote': {'sheets': ['ADH','MZY'], 'short': ['Al Dhanna','Madinat Zayed'], 'subtitle': 'Al Dhanna Ruwais & Al Dhafra Madinat Zayed City'},
'Dubai': {'sheets': ['DMC','DBN'], 'short': ['Academic City','Al Nahda'], 'subtitle': 'Dubai Academic City & Dubai Al Nahda'},
'Fujairah': {'sheets': ['FJF','FJH'], 'short': ['Faseel','Hulaifat'], 'subtitle': 'Fujairah Faseel & Fujairah Hulaifat'},
'Sharjah': {'sheets': ['SJA','SJB'], 'short': ['Campus A','Campus B'], 'subtitle': 'Sharjah Campus A & Sharjah Campus B'},
'Ras Al Khaimah': {'sheets': ['RKA','RKB'], 'short': ['Campus A','Campus B'], 'subtitle': 'RAK Campus A & RAK Campus B'},
}

CAMPUS_ORDER_14 = ['ADA','ADB','AAF','AAZ','DMC','DBN','SJA','SJB','FJF','FJH','RKA','RKB','ADH','MZY']
CAMPUS_ORDER_15 = ['ADA','HQ','ADB','AAF','AAZ','DMC','DBN','SJA','SJB','FJF','FJH','RKA','RKB','ADH','MZY']
REGION_ORDER = ['Abu Dhabi', 'AD Al Ain', 'Dubai', 'Sharjah', 'Fujairah', 'Ras Al Khaimah', 'AD Remote']
REGION_LABEL_MAP = {
'Abu Dhabi Main': 'Abu Dhabi', 'Abu Dhabi': 'Abu Dhabi',
'Al Ain': 'AD Al Ain',
'Dubai': 'Dubai',
'Sharjah': 'Sharjah',
'Fujairah': 'Fujairah',
'Ras Al Khaimah': 'Ras Al Khaimah', 'RAK': 'Ras Al Khaimah',
'Al Dhafra': 'AD Remote', 'Al Dhanna': 'AD Remote',
}

KPI_WEIGHTS = {
2: 0.30, 3: 0.10, 4: 0.10, 5: 0.25, 6: 0.25,
7: 0.30, 8: 0.50, 9: 0.20,
10: 0.50, 11: 0.50,
12: 0.40, 13: 0.40, 14: 0.10, 15: 0.10,
16: 0.30, 17: 0.30, 18: 0.20, 19: 0.20,
}

MONTH_NAMES = ['January','February','March','April','May','June',
'July','August','September','October','November','December']

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

# Ã¢ÂÂÃ¢ÂÂ Smartsheet API Ã¢ÂÂÃ¢ÂÂ

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

# Ã¢ÂÂÃ¢ÂÂ Fetch KPI data Ã¢ÂÂÃ¢ÂÂ

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

def fetch_incident_types(token, month_filter):
    """Fetch incident type breakdown from notification report."""
    try:
        rows = fetch_report_rows('1199821531598724', token)
    except:
        return {}
    type_counts = {}
    for row in rows:
        if month_filter:
            rm = normalize_month(row.get('Reporting Month')) or normalize_month(row.get('Date Reported'))
            if rm != month_filter: continue
        itype = str(row.get('Classification', '') or row.get('Incident Type', '') or row.get('Type', '')).strip()
        if not itype: itype = 'Unclassified'
        type_counts[itype] = type_counts.get(itype, 0) + 1
    return type_counts

# Ã¢ÂÂÃ¢ÂÂ Chart data replacement (regex-based, preserves namespace prefixes) Ã¢ÂÂÃ¢ÂÂ

def replace_numcache_values(chart_xml, new_series_values):
    ns_match = re.search(r'<(\w+):numCache>', chart_xml)
    if not ns_match:
        return chart_xml
    ns = ns_match.group(1)

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

# Ã¢ÂÂÃ¢ÂÂ Build chart data from KPI data Ã¢ÂÂÃ¢ÂÂ

def get_campus_val(kpi_data, campus, kpi_row, field='calc'):
    return kpi_data.get(campus, {}).get(kpi_row, {}).get(field, 0)

def region_agg(kpi_data, kpi_row, region_name, field):
    cfg = REGIONS.get(region_name)
    if not cfg: return 0
    vals = [get_campus_val(kpi_data, s, kpi_row, field) for s in cfg['sheets']]
    return sum(vals)

def build_chart_data(kpi_data, training_hours, incident_types=None):
    charts = {}

    def pct_14(kpi_row):
        return [[get_campus_val(kpi_data, c, kpi_row, 'calc') for c in CAMPUS_ORDER_14]]

    def pct_15(kpi_row):
        vals = []
        for c in CAMPUS_ORDER_15:
            if c == 'HQ':
                vals.append(0)
            else:
                vals.append(get_campus_val(kpi_data, c, kpi_row, 'calc'))
        return [vals]

    def planned_actual_14(kpi_row):
        planned = [get_campus_val(kpi_data, c, kpi_row, 'planned') for c in CAMPUS_ORDER_14]
        actual = [get_campus_val(kpi_data, c, kpi_row, 'achieved') for c in CAMPUS_ORDER_14]
        return [planned, actual]

    charts[1] = pct_14(2)

    planned = []
    conducted = []
    for rname in REGION_ORDER:
        p = 0
        a = 0
        for alias, rkey in REGION_LABEL_MAP.items():
            if rkey == rname:
                p += get_campus_val(kpi_data, alias, 5, 'planned')
                a += get_campus_val(kpi_data, alias, 5, 'achieved')
        p += get_campus_val(kpi_data, rname, 5, 'planned')
        a += get_campus_val(kpi_data, rname, 5, 'achieved')
        planned.append(round(p))
        conducted.append(round(a))
    charts[2] = [planned, conducted]

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

    charts[4] = pct_14(4)
    charts[5] = [[training_hours.get(c, 0) for c in CAMPUS_ORDER_14]]
    charts[6] = pct_14(6)

    chart7_campuses = ['ADA','ADB','FJF','FJH','RKA','RKB','ADH']
    charts[7] = [[get_campus_val(kpi_data, c, 9, 'calc') for c in chart7_campuses]]

    charts[8] = planned_actual_14(14)
    charts[9] = planned_actual_14(15)
    charts[10] = planned_actual_14(16)
    charts[11] = pct_14(17)
    charts[12] = pct_15(18)
    charts[13] = pct_15(19)

    # Chart 14: Incidents by Type (pie chart) Ã¢ÂÂ single series
    if incident_types:
        charts[14] = [list(incident_types.values())]
    else:
        charts[14] = [[0]]

    # Chart 15: Risk Assessment Closed (dual bar Ã¢ÂÂ planned/actual)
    charts[15] = planned_actual_14(7)

    # Chart 16: Risk Assessment Validated & Signed Off (dual bar)
    charts[16] = planned_actual_14(8)

    # Chart 17: Planned Training Report (dual bar)
    charts[17] = planned_actual_14(10)

    # Chart 18: Total Incidents (single bar per campus)
    charts[18] = [[get_campus_val(kpi_data, c, 19, 'planned') for c in CAMPUS_ORDER_14]]

    # Chart 19: EHS Inspection Rate (single bar %)
    charts[19] = pct_14(16)

    # Chart 20: Findings Closed Rate (single bar %)
    charts[20] = pct_14(17)

    return charts

# Ã¢ÂÂÃ¢ÂÂ Dynamic chart XML generators Ã¢ÂÂÃ¢ÂÂ

INCIDENT_TYPE_LABELS = ['Unclassified','Equipment/Property Damage','First Aid Case','Near Miss','Medical Treatment Case','Lost Workdays Injury']
PIE_COLORS = ['003366','FFC000','00B050','4472C4','7030A0','ED7D31']

def make_bar_chart_xml(title, categories, series_list):
    cat_xml = '<c:cat><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache>'
    cat_xml += f'<c:ptCount val="{len(categories)}"/>'
    for i, c in enumerate(categories):
        cat_xml += f'<c:pt idx="{i}"><c:v>{c}</c:v></c:pt>'
    cat_xml += '</c:strCache></c:strRef></c:cat>'

    series_xml = ''
    for si, s in enumerate(series_list):
        color = s.get('color', '4472C4')
        vals = s['values']
        series_xml += f'<c:ser><c:idx val="{si}"/><c:order val="{si}"/>'
        series_xml += f'<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{s["name"]}</c:v></c:pt></c:strCache></c:strRef></c:tx>'
        series_xml += f'<c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>'
        series_xml += f'<c:invertIfNegative val="0"/>{cat_xml}'
        series_xml += f'<c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(vals)}"/>'
        for i, v in enumerate(vals):
            series_xml += f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
        series_xml += '</c:numCache></c:numRef></c:val></c:ser>'

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:chart><c:autoTitleDeleted val="1"/>
<c:plotArea><c:layout/>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>
{series_xml}
<c:axId val="111111111"/><c:axId val="222222222"/>
</c:barChart>
<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>
<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>
</c:plotArea>
<c:legend><c:legendPos val="b"/></c:legend>
<c:plotVisOnly val="1"/>
</c:chart></c:chartSpace>'''

def make_pie_chart_xml(categories, values):
    dpt_xml = ''
    for i in range(len(categories)):
        color = PIE_COLORS[i % len(PIE_COLORS)]
        dpt_xml += f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr></c:dPt>'

    cat_xml = '<c:cat><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache>'
    cat_xml += f'<c:ptCount val="{len(categories)}"/>'
    for i, c in enumerate(categories):
        cat_xml += f'<c:pt idx="{i}"><c:v>{c}</c:v></c:pt>'
    cat_xml += '</c:strCache></c:strRef></c:cat>'

    val_xml = '<c:val><c:numRef><c:f>Sheet1!$B$1</c:f><c:numCache>'
    val_xml += f'<c:formatCode>General</c:formatCode><c:ptCount val="{len(values)}"/>'
    for i, v in enumerate(values):
        val_xml += f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
    val_xml += '</c:numCache></c:numRef></c:val>'

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:chart><c:autoTitleDeleted val="1"/>
<c:plotArea><c:layout/>
<c:pieChart><c:varyColors val="1"/>
<c:ser><c:idx val="0"/><c:order val="0"/>
<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Count</c:v></c:pt></c:strCache></c:strRef></c:tx>
{dpt_xml}{cat_xml}{val_xml}
</c:ser></c:pieChart>
</c:plotArea>
<c:legend><c:legendPos val="b"/></c:legend>
<c:plotVisOnly val="1"/>
</c:chart></c:chartSpace>'''

STYLE_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="102"/>'
COLORS_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"/>'

def make_chart_rels(n):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style{n}.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors{n}.xml"/></Relationships>'

# New chart definitions: chart_num -> {type, title, kpi_row or special, series_config}
NEW_CHART_DEFS = {
    14: {'type': 'pie', 'title': 'Incidents by Type'},
    15: {'type': 'dual_bar', 'title': 'Risk Assessment Closed', 'kpi_row': 7,
         'series': [{'name': 'Total Risk Assessments Registered', 'field': 'planned', 'color': '4472C4'},
                    {'name': 'Risk Assessment Closed', 'field': 'achieved', 'color': '00B050'}]},
    16: {'type': 'dual_bar', 'title': 'Risk Assessment Validated & Signed Off', 'kpi_row': 8,
         'series': [{'name': 'Total Assessments Register', 'field': 'planned', 'color': '4472C4'},
                    {'name': 'RA Validated and Signed Off', 'field': 'achieved', 'color': '00B050'}]},
    17: {'type': 'dual_bar', 'title': 'Planned Training Report', 'kpi_row': 10,
         'series': [{'name': 'Planned Training', 'field': 'planned', 'color': '4472C4'},
                    {'name': 'Training Conducted', 'field': 'achieved', 'color': '00B050'}]},
    18: {'type': 'single_bar', 'title': 'Total Incidents', 'kpi_row': 19, 'field': 'planned',
         'series_name': 'Total Incidents', 'color': '00249C'},
    19: {'type': 'pct_bar', 'title': 'EHS Inspection Rate', 'kpi_row': 16},
    20: {'type': 'pct_bar', 'title': 'Findings Closed Rate', 'kpi_row': 17},
}

def build_new_chart_xml(chart_num, kpi_data, incident_types):
    defn = NEW_CHART_DEFS[chart_num]

    if defn['type'] == 'pie':
        # Incidents by Type pie chart
        cats = list(incident_types.keys()) if incident_types else INCIDENT_TYPE_LABELS
        vals = list(incident_types.values()) if incident_types else [0] * len(cats)
        return make_pie_chart_xml(cats, vals)

    cats = CAMPUS_ORDER_14
    if defn['type'] == 'dual_bar':
        kpi = defn['kpi_row']
        series = []
        for s in defn['series']:
            vals = [get_campus_val(kpi_data, c, kpi, s['field']) for c in cats]
            series.append({'name': s['name'], 'values': [round(v) for v in vals], 'color': s['color']})
        return make_bar_chart_xml(defn['title'], cats, series)

    if defn['type'] == 'single_bar':
        kpi = defn['kpi_row']
        vals = [round(get_campus_val(kpi_data, c, kpi, defn['field'])) for c in cats]
        return make_bar_chart_xml(defn['title'], cats, [{'name': defn['series_name'], 'values': vals, 'color': defn['color']}])

    if defn['type'] == 'pct_bar':
        kpi = defn['kpi_row']
        met = []
        below = []
        for c in cats:
            pct = get_campus_val(kpi_data, c, kpi, 'calc')
            if pct >= 0.9:
                met.append(round(pct * 100))
                below.append(0)
            else:
                met.append(0)
                below.append(round(pct * 100))
        return make_bar_chart_xml(defn['title'], cats, [
            {'name': 'Met Target', 'values': met, 'color': '00B050'},
            {'name': 'Below Target', 'values': below, 'color': 'FF0000'},
        ])

    return ''

def make_drawing_xml(rid, chart_num):
    return f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5760000" cy="2520000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{100+chart_num}" name="Chart {chart_num}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing>'

def make_figure_para(rid, chart_num, label):
    drawing = make_drawing_xml(rid, chart_num)
    return f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>{drawing}</w:r></w:p><w:p><w:pPr><w:spacing w:after="200"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">Figure: {label}</w:t></w:r></w:p>'

# Heading text patterns to inject each new chart after
CHART_INJECT_MAP = [
    (15, 'Risk Assessments Closed', 'Heading2', 'Risk Assessment Closed'),
    (16, 'Risk Assessment Validated', 'Heading2', 'Risk Assessment Validated &amp; Signed Off'),
    (17, 'Planned H', 'Heading2', 'Planned Training Report'),
    (19, 'Inspections Completed', 'Heading2', 'EHS Inspection Rate'),
    (20, 'Findings Closed', 'Heading2', 'Findings Closed Rate'),
    (18, 'Incident Notifications', 'Heading2', 'Total Incidents by Campus'),
]

def inject_new_charts_into_doc(doc_xml):
    """Insert drawing elements for charts 14-20 into document.xml."""
    # Charts 15-20: inject after specific Heading2 paragraphs
    for chart_num, heading_text, heading_style, figure_label in CHART_INJECT_MAP:
        rid_num = 53 + (chart_num - 14)
        rid = f'rId{rid_num}'
        figure_xml = make_figure_para(rid, chart_num, figure_label)

        pattern = re.compile(
            rf'<w:p\b[^>]*>(?:(?!</w:p>).)*?<w:pStyle\s+w:val="{heading_style}"\s*/>(?:(?!</w:p>).)*?</w:p>',
            re.DOTALL
        )
        words = heading_text.lower().split()
        found = False
        for m in pattern.finditer(doc_xml):
            texts = re.findall(r'<w:t[^>]*>([^<]+)</w:t>', m.group(0))
            para_text = ' '.join(texts).lower()
            para_text = ' '.join(para_text.split())
            if all(w in para_text for w in words):
                pos = m.end()
                doc_xml = doc_xml[:pos] + figure_xml + doc_xml[pos:]
                print(f'  Injected chart{chart_num} ({figure_label}) after "{heading_text}"')
                found = True
                break
        if not found:
            print(f'  WARNING: Could not find heading "{heading_text}" for chart{chart_num}')

    # Chart 14 (Incidents by Type pie) Ã¢ÂÂ after Incidents Heading1
    rid14 = 'rId53'
    fig14 = make_figure_para(rid14, 14, 'Incidents by Type')
    pattern = re.compile(
        r'<w:p\b[^>]*>(?:(?!</w:p>).)*?<w:pStyle\s+w:val="Heading1"\s*/>(?:(?!</w:p>).)*?</w:p>',
        re.DOTALL
    )
    for m in pattern.finditer(doc_xml):
        texts = re.findall(r'<w:t[^>]*>([^<]+)</w:t>', m.group(0))
        para_text = ' '.join(texts).lower().strip()
        para_text = ' '.join(para_text.split())
        if 'incidents' in para_text and 'notification' not in para_text and 'planned' not in para_text:
            pos = m.end()
            doc_xml = doc_xml[:pos] + fig14 + doc_xml[pos:]
            print(f'  Injected chart14 (Incidents by Type pie) after Incidents heading')
            break

    return doc_xml

# Ã¢ÂÂÃ¢ÂÂ Executive Summary & Incidents content injection Ã¢ÂÂÃ¢ÂÂ

PILLAR_KPIS = [
{'pillar': 'Leadership, Accountability &amp; Engagement', 'weight': 0.20, 'rows': [2,4,5,6]},
{'pillar': 'Risk Management &amp; Planning', 'weight': 0.20, 'rows': [7,8,9]},
{'pillar': 'Training &amp; Awareness', 'weight': 0.10, 'rows': [10]},
{'pillar': 'OCP &amp; Emergency Preparedness', 'weight': 0.25, 'rows': [13,14,15]},
{'pillar': 'Performance Evaluation &amp; Improvement', 'weight': 0.25, 'rows': [16,17,18,19]},
]

def compute_pillar_scores(kpi_data):
    pillar_scores = []
    for pillar in PILLAR_KPIS:
        campus_scores = []
        for c in CAMPUS_ORDER_14:
            kpi_vals = [get_campus_val(kpi_data, c, r, 'calc') for r in pillar['rows']]
            kpi_wts = [KPI_WEIGHTS.get(r, 0.05) for r in pillar['rows']]
            tw = sum(kpi_wts)
            score = sum(v * w for v, w in zip(kpi_vals, kpi_wts)) / tw if tw > 0 else 0
            campus_scores.append(score)
        avg = sum(campus_scores) / len(campus_scores) if campus_scores else 0
        pillar_scores.append({'pillar': pillar['pillar'], 'weight': pillar['weight'], 'score': avg})
    overall = sum(p['score'] * p['weight'] for p in pillar_scores)
    return pillar_scores, overall

def _oc(text, w, bold=False, color=None, fill=None, center=False):
    tp = f'<w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>'
    if fill: tp += f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    tp += '</w:tcPr>'
    pp = '<w:pPr><w:jc w:val="center"/></w:pPr>' if center else ''
    rp = '<w:rPr>'
    if bold: rp += '<w:b/>'
    if color: rp += f'<w:color w:val="{color}"/>'
    rp += '<w:sz w:val="20"/></w:rPr>'
    return f'<w:tc>{tp}<w:p>{pp}<w:r>{rp}<w:t>{text}</w:t></w:r></w:p></w:tc>'

def _sc(pct):
    return '00B050' if pct >= 90 else ('FFC000' if pct >= 70 else 'FF0000')

def _ss(pct):
    return 'On Track' if pct >= 90 else ('At Risk' if pct >= 70 else 'Below Target')

_NF = '1C2340'
_LF = 'D9E2F3'

_TBL = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9072" w:type="dxa"/>'
'<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
'<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
'<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
'</w:tblBorders></w:tblPr>')

def _intro(text):
    return f'<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

def build_exec_summary_xml(kpi_data):
    ps, overall = compute_pillar_scores(kpi_data)
    x = [_intro('The following table summarizes the overall KPI performance across all campuses.'), _TBL]
    x.append('<w:tr>' + _oc('Pillar',4000,True,'FFFFFF',_NF) + _oc('Weight',1500,True,'FFFFFF',_NF,True) +
    _oc('Score',1500,True,'FFFFFF',_NF,True) + _oc('Status',2072,True,'FFFFFF',_NF,True) + '</w:tr>')
    for p in ps:
        pct = round(p['score'] * 100)
        x.append('<w:tr>' + _oc(p['pillar'],4000) + _oc(f'{int(p["weight"]*100)}%',1500,center=True) +
        _oc(f'{pct}%',1500,True,_sc(pct),None,True) + _oc(_ss(pct),2072,True,_sc(pct),None,True) + '</w:tr>')
    op = round(overall * 100)
    x.append('<w:tr>' + _oc('Overall Weighted Score',4000,True,None,_LF) + _oc('100%',1500,True,None,_LF,True) +
    _oc(f'{op}%',1500,True,_sc(op),_LF,True) + _oc(_ss(op),2072,True,_sc(op),_LF,True) + '</w:tr>')
    x.append('</w:tbl>')
    return ''.join(x)

def build_incidents_xml(kpi_data):
    x = [_intro('The following table summarizes incident notifications and investigations by campus.'), _TBL]
    x.append('<w:tr>' + _oc('Campus',2268,True,'FFFFFF',_NF) + _oc('Total Incidents',2268,True,'FFFFFF',_NF,True) +
    _oc('Notification on Time',2268,True,'FFFFFF',_NF,True) + _oc('Investigation on Time',2268,True,'FFFFFF',_NF,True) + '</w:tr>')
    ti = tn = tv = 0
    for campus in CAMPUS_ORDER_14:
        nt = get_campus_val(kpi_data, campus, 19, 'planned')
        no = get_campus_val(kpi_data, campus, 19, 'achieved')
        it_ = get_campus_val(kpi_data, campus, 18, 'planned')
        io_ = get_campus_val(kpi_data, campus, 18, 'achieved')
        inc = round(nt); ti += inc; tn += round(no); tv += round(io_)
        np_ = round(no/nt*100) if nt > 0 else 0
        ip = round(io_/it_*100) if it_ > 0 else 0
        x.append('<w:tr>' + _oc(campus,2268) + _oc(str(inc),2268,center=True) +
        _oc(f'{np_}%',2268,True,_sc(np_),None,True) + _oc(f'{ip}%',2268,True,_sc(ip),None,True) + '</w:tr>')
    x.append('<w:tr>' + _oc('TOTAL',2268,True,None,_LF) + _oc(str(ti),2268,True,None,_LF,True) +
    _oc(str(tn),2268,True,None,_LF,True) + _oc(str(tv),2268,True,None,_LF,True) + '</w:tr>')
    x.append('</w:tbl>')
    return ''.join(x)

def inject_section_content(doc_xml, section_title, content_xml):
    """Inject content_xml after the Heading1 paragraph containing section_title."""
    ns_match = re.search(r'<(\w+):body>', doc_xml)
    ns = ns_match.group(1) if ns_match else 'w'
    if ns != 'w':
        content_xml = content_xml.replace('<w:', f'<{ns}:').replace('</w:', f'</{ns}:')
    heading_pattern = re.compile(
        rf'<{ns}:p\b[^>]*>(?:(?!</{ns}:p>).)*?<{ns}:pStyle\s+{ns}:val="Heading1"\s*/>(?:(?!</{ns}:p>).)*?</{ns}:p>',
        re.DOTALL
    )
    title_words = section_title.lower().split()
    for m in heading_pattern.finditer(doc_xml):
        para_xml = m.group(0)
        texts = re.findall(rf'<{ns}:t[^>]*>([^<]+)</{ns}:t>', para_xml)
        para_text = ' '.join(texts).lower()
        para_text_norm = ' '.join(para_text.split())
        if all(w in para_text_norm for w in title_words):
            pos = m.end()
            print(f'  Injected "{section_title}" content after heading at pos {pos}')
            return doc_xml[:pos] + content_xml + doc_xml[pos:]
    print(f'  WARNING: Heading1 section "{section_title}" not found')
    return doc_xml

# Ã¢ÂÂÃ¢ÂÂ Generate report using template Ã¢ÂÂÃ¢ÂÂ

def generate_report(month_name, year, token):
    if not month_name:
        month_name = None
    period_label = f'{month_name} {year}' if month_name else f'Overall {year}'
    print(f'Generating Word report (overall) for {period_label}')

    # Fetch data
    kpi_data = fetch_kpi_data(token, month_name)
    training_hours = fetch_training_hours(token, month_name)
    incident_types = fetch_incident_types(token, month_name)
    print(f'  KPI data: {len(kpi_data)} campuses')
    print(f'  Training hours: {len(training_hours)} campuses')
    print(f'  Incident types: {len(incident_types)} types')

    # Build chart replacement data (charts 1-13 use template charts, 14-20 are dynamic)
    chart_data = build_chart_data(kpi_data, training_hours, incident_types)

    # Build section content
    exec_summary_xml = build_exec_summary_xml(kpi_data)
    incidents_xml = build_incidents_xml(kpi_data)

    # Load template
    tmpl_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates', 'word_template.docx')
    with open(tmpl_path, 'rb') as f:
        tmpl_bytes = f.read()

    # Unzip, modify charts and sections, rezip
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(tmpl_bytes), 'r') as zin:
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                # Replace chart data (charts 1-13)
                chart_match = re.match(r'word/charts/chart(\d+)\.xml', item.filename)
                if chart_match:
                    chart_num = int(chart_match.group(1))
                    if chart_num in chart_data:
                        xml_str = data.decode('utf-8')
                        xml_str = replace_numcache_values(xml_str, chart_data[chart_num])
                        data = xml_str.encode('utf-8')
                        print(f'  Updated chart{chart_num} with {len(chart_data[chart_num])} series')

                # Update Content_Types.xml to include new charts
                if item.filename == '[Content_Types].xml':
                    ct = data.decode('utf-8')
                    new_ct = ''
                    for n in range(14, 21):
                        new_ct += f'<Override PartName="/word/charts/chart{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                    ct = ct.replace('</Types>', new_ct + '</Types>')
                    data = ct.encode('utf-8')

                # Update document.xml.rels with new chart relationships
                if item.filename == 'word/_rels/document.xml.rels':
                    rels = data.decode('utf-8')
                    new_rels = ''
                    for i, n in enumerate(range(14, 21)):
                        rid_num = 53 + i
                        new_rels += f'<Relationship Id="rId{rid_num}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{n}.xml"/>'
                    rels = rels.replace('</Relationships>', new_rels + '</Relationships>')
                    data = rels.encode('utf-8')

                # Inject sections and new chart drawings into document.xml
                if item.filename == 'word/document.xml':
                    xml_str = data.decode('utf-8')
                    xml_str = inject_section_content(xml_str, 'Executive Summary', exec_summary_xml)
                    xml_str = inject_section_content(xml_str, 'Incidents', incidents_xml)
                    xml_str = inject_new_charts_into_doc(xml_str)
                    data = xml_str.encode('utf-8')

                zout.writestr(item, data)

            # Add new chart XML files (14-20)
            for n in range(14, 21):
                chart_xml = build_new_chart_xml(n, kpi_data, incident_types)
                zout.writestr(f'word/charts/chart{n}.xml', chart_xml)
                zout.writestr(f'word/charts/style{n}.xml', STYLE_XML)
                zout.writestr(f'word/charts/colors{n}.xml', COLORS_XML)
                zout.writestr(f'word/charts/_rels/chart{n}.xml.rels', make_chart_rels(n))
                print(f'  Created chart{n} ({NEW_CHART_DEFS[n]["title"]})')

    buf.seek(0)
    return buf.getvalue()

# Ã¢ÂÂÃ¢ÂÂ HTTP Handler Ã¢ÂÂÃ¢ÂÂ

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
