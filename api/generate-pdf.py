"""Vercel Python serverless function — HCT-COHS KPI PDF Report Generator.
Fetches live data from Smartsheet API and generates downloadable PDF files.
Uses fpdf2 (pure Python, no system dependencies).
"""

import os, io, json
from http.server import BaseHTThPRequestHandler
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen
from fpdf import FPDF

# ── Regions ──
REGIONS = {
    'AD Al Ain':       {'sheets': ['AAF','AAZ'], 'short': ['Falaj Hahhzza','Zakhir'],        'subtitle': 'Al Ain Falaj Hazza & Al Ain Zakhir'},
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

MONTH_NAMES = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']


# ── Smartsheet API (shared with other generators) ──

def _ss_fetch(endpoint, token):
    url = f'https://api.smartsheet.com/2.0/{endpoint}'
    req = Request(url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'})
    with urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())

def fetch_sheet_rows(sheet_id, token):
    data = _ss_fetch(f'sheets/{sheet_id}?pageSize=500', token)
    if not data.get('rows'): return []
    col_map = {c['id']: c['title'] for c in data.get('columns', [])}
    return [{col_map.get(cell.get('columnId'), ''): cell.get('displayValue') or cell.get('value') or ''
             for cell in row.get('cells', []) if col_map.get(cell.get('columnId'))}
            for row in data['rows'] if row.get('cells')]

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
            if title: rec[title] = cell.get('displayValue') or cell.get('value') or ''
        if rec: rows.append(rec)
    return rows

def normalize_month(v):
    if not v: return None
    s = str(v).strip()
    if not s: return None
    for m in MONTH_NAMES:
        if m.lower() == s.lower(): return m
    abbr_map = {m[:3].lower(): m for m in MONTH_NAMES}
    if s[:3].lower() in abbr_map: return abbr_map[s[:3].lower()]
    try:
        from datetime import datetime as dt
        d = dt.strptime(s, '%Y-%m-%d')
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
            rows = fetch_report_rows(src['reportId'], token) if src.get('reportId') else fetch_sheet_rows(src['sheetId'], token)
        except:
            continue

        kpi_row = src['kpi_row']
        campus_agg = {}
        for row in rows:
            campus = str(row.get(src['campusCol'], '')).strip()
            if not campus: continue
            if month_filter:
                rm = normalize_month(row.get(src.get('monthCol', '')))
                if rm != month_filter:
                    rm = normalize_month(row.get('Reporting Month')) or normalize_month(row.get('Primary'))
                    if rm != month_filter: continue

            if campus not in campus_agg:
                campus_agg[campus] = {'planned': 0, 'actual': 0}

            if src.get('yesNoCount'):
                campus_agg[campus]['planned'] += 1 if str(row.get(src.get('plannedCol',''), '')).strip().lower() == 'yes' else 0
                campus_agg[campus]['actual'] += 1 if str(row.get(src.get('actualCol',''), '')).strip().lower() == 'yes' else 0
            elif src.get('plannedCol') and src.get('actualCol'):
                campus_agg[campus]['planned'] += safe_float(row.get(src['plannedCol']))
                campus_agg[campus]['actual'] += safe_float(row.get(src['actualCol']))
            elif src.get('valueCol'):
                v = safe_float(row.get(src['valueCol']))
                campus_agg[campus]['planned'] += v
                campus_agg[campus]['actual'] += v

        weight = KPI_WEIGHTS.get(kpi_row, 0.05)
        for campus, agg in campus_agg.items():
            if campus not in data: data[campus] = {}
            p, a = agg['planned'], agg['actual']
            calc = min(a / p, 1.0) if p > 0 else (1.0 if a > 0 else 0)
            data[campus][kpi_row] = {'planned': p, 'achieved': a, 'calc': calc, 'weight': weight}
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
        entry = {col: safe_float(row.get(col)) for col in WASTE_TABLE_COLS}
        entry['Total Waste'] = safe_float(row.get('Total Waste')) or sum(entry.values())
        waste[campus] = entry
    return waste


# ── KPI processing ──

def read_campus_data(kpi_data, sheet_name):
    campus = kpi_data.get(sheet_name, {})
    kpis, pillar_scores = [], []
    for pillar in PILLAR_KPIS:
        p_kpis = []
        for row in pillar['rows']:
            d = campus.get(row, {'planned': 0, 'achieved': 0, 'calc': 1.0, 'weight': KPI_WEIGHTS.get(row, 0.05)})
            p_kpis.append(d); kpis.append(d)
        tw = sum(k['weight'] for k in p_kpis)
        pillar_scores.append(sum(k['calc'] * k['weight'] for k in p_kpis) / tw if tw > 0 else 0)
    overall = sum(s * p['weight'] for s, p in zip(pillar_scores, PILLAR_KPIS))
    return {'sheet': sheet_name, 'kpis': kpis, 'pillar_scores': pillar_scores, 'overall': overall}

def read_region_data(kpi_data, region_cfg):
    campuses = [read_campus_data(kpi_data, s) for s in region_cfg['sheets']]
    n = len(campuses)
    return {
        'campuses': campuses,
        'avg_pillar': [sum(c['pillar_scores'][i] for c in campuses)/n for i in range(5)],
        'avg_overall': sum(c['overall'] for c in campuses)/n,
        'short': region_cfg['short'],
    }


# ── PDF Builder ──

class KPIReport(FPDF):
    def __init__(self, region_name, period):
        super().__init__()
        self.region_name = region_name
        self.period = period

    def header(self):
        self.set_font('Helvetica', 'B', 8)
        self.set_text_color(100, 100, 100)
        self.cell(0, 5, f'HCT-COHS KPI Report | {self.region_name} | {self.period}', align='L')
        self.ln(8)
        self.set_draw_color(24, 14, 63)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(4)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', align='C')

    def cover_page(self, region_cfg):
        self.add_page()
        self.ln(40)
        self.set_font('Helvetica', 'B', 28)
        self.set_text_color(24, 14, 63)
        self.cell(0, 14, 'Corporate OHS Monthly', align='C', new_x='LMARGIN', new_y='NEXT')
        self.cell(0, 14, 'KPI Performance Report', align='C', new_x='LMARGIN', new_y='NEXT')
        self.ln(10)
        self.set_font('Helvetica', 'B', 18)
        self.set_text_color(28, 35, 64)
        self.cell(0, 10, self.region_name, align='C', new_x='LMARGIN', new_y='NEXT')
        self.set_font('Helvetica', '', 12)
        self.set_text_color(89, 89, 89)
        self.cell(0, 8, region_cfg['subtitle'], align='C', new_x='LMARGIN', new_y='NEXT')
        self.ln(15)
        self.set_font('Helvetica', 'B', 16)
        self.set_text_color(24, 14, 63)
        self.cell(0, 10, self.period, align='C', new_x='LMARGIN', new_y='NEXT')
        self.ln(20)
        # Info table
        self.set_font('Helvetica', '', 10)
        self.set_text_color(0, 0, 0)
        info = [
            ('Client:', 'Higher Colleges of Technology'),
            ('Issued By:', 'Corporate OHS LLC OPC'),
            ('Period:', self.period),
        ]
        for label, value in info:
            self.set_font('Helvetica', 'B', 10)
            self.set_fill_color(24, 14, 63)
            self.set_text_color(255, 255, 255)
            self.cell(50, 8, label, fill=True, border=1)
            self.set_font('Helvetica', '', 10)
            self.set_text_color(0, 0, 0)
            self.cell(140, 8, value, border=1, new_x='LMARGIN', new_y='NEXT')

    def kpi_table(self, region_data, region_cfg):
        self.add_page()
        self.set_font('Helvetica', 'B', 16)
        self.set_text_color(24, 14, 63)
        self.cell(0, 10, '1. KPI Performance Summary', new_x='LMARGIN', new_y='NEXT')
        self.ln(4)

        # Column widths
        w_num = 10
        w_kpi = 55
        w_campus = 35
        w_avg = 30
        n_campuses = len(region_cfg['short'])
        total_w = w_num + w_kpi + w_campus * n_campuses + w_avg

        # Header
        self.set_font('Helvetica', 'B', 7)
        self.set_fill_color(24, 14, 63)
        self.set_text_color(255, 255, 255)
        self.cell(w_num, 7, '#', border=1, fill=True, align='C')
        self.cell(w_kpi, 7, 'KPI', border=1, fill=True, align='C')
        for name in region_cfg['short']:
            self.cell(w_campus, 7, name, border=1, fill=True, align='C')
        self.cell(w_avg, 7, 'Average', border=1, fill=True, align='C')
        self.ln()

        campuses = region_data['campuses']
        kpi_idx = 0

        for pillar in PILLAR_KPIS:
            # Pillar header
            self.set_font('Helvetica', 'B', 7)
            self.set_fill_color(217, 226, 243)
            self.set_text_color(28, 35, 64)
            self.cell(w_num, 6, '', border=1, fill=True)
            self.cell(w_kpi + w_campus * n_campuses + w_avg, 6, pillar['pillar'], border=1, fill=True)
            self.ln()

            for row_num in pillar['rows']:
                self.set_font('Helvetica', '', 7)
                self.set_text_color(0, 0, 0)
                self.cell(w_num, 6, str(row_num), border=1, align='C')
                self.cell(w_kpi, 6, KPI_NAMES.get(row_num, f'KPI {row_num}'), border=1)

                vals = []
                for c in campuses:
                    kpi = c['kpis'][kpi_idx] if kpi_idx < len(c['kpis']) else {'calc': 0}
                    pct = round(kpi['calc'] * 100)
                    if pct >= 90:
                        self.set_text_color(0, 176, 80)
                    elif pct >= 70:
                        self.set_text_color(255, 192, 0)
                    else:
                        self.set_text_color(255, 0, 0)
                    self.set_font('Helvetica', '', 7)
                    self.cell(w_campus, 6, f'{pct}%', border=1, align='C')
                    vals.append(kpi['calc'])

                avg = sum(vals) / len(vals) if vals else 0
                avg_pct = round(avg * 100)
                if avg_pct >= 90:
                    self.set_text_color(0, 176, 80)
                elif avg_pct >= 70:
                    self.set_text_color(255, 192, 0)
                else:
                    self.set_text_color(255, 0, 0)
                self.set_font('Helvetica', 'B', 7)
                self.cell(w_avg, 6, f'{avg_pct}%', border=1, align='C')
                self.ln()
                kpi_idx += 1

        # Overall row
        self.set_font('Helvetica', 'B', 7)
        self.set_fill_color(28, 35, 64)
        self.set_text_color(255, 255, 255)
        self.cell(w_num, 7, '', border=1, fill=True)
        self.cell(w_kpi, 7, 'Overall Weighted Score', border=1, fill=True)
        for c in campuses:
            self.cell(w_campus, 7, f'{round(c["overall"]*100)}%', border=1, fill=True, align='C')
        self.cell(w_avg, 7, f'{round(region_data["avg_overall"]*100)}%', border=1, fill=True, align='C')
        self.ln()

    def pillar_breakdown(self, region_data, region_cfg):
        self.add_page()
        self.set_font('Helvetica', 'B', 16)
        self.set_text_color(24, 14, 63)
        self.cell(0, 10, '2. Pillar Score Breakdown', new_x='LMARGIN', new_y='NEXT')
        self.ln(4)

        for i, pillar in enumerate(PILLAR_KPIS):
            score = round(region_data['avg_pillar'][i] * 100)
            self.set_font('Helvetica', 'B', 11)
            self.set_text_color(28, 35, 64)
            self.cell(0, 8, f'{pillar["pillar"]}  -  {score}%  (Weight: {int(pillar["weight"]*100)}%)',
                      new_x='LMARGIN', new_y='NEXT')

            for ci, c in enumerate(region_data['campuses']):
                campus_score = round(c['pillar_scores'][i] * 100)
                self.set_font('Helvetica', '', 10)
                if campus_score >= 90:
                    self.set_text_color(0, 140, 60)
                elif campus_score >= 70:
                    self.set_text_color(200, 150, 0)
                else:
                    self.set_text_color(200, 0, 0)
                self.cell(0, 6, f'    {region_cfg["short"][ci]}: {campus_score}%',
                          new_x='LMARGIN', new_y='NEXT')
            self.ln(4)

    def waste_section(self, waste_data, region_cfg):
        self.add_page()
        self.set_font('Helvetica', 'B', 16)
        self.set_text_color(24, 14, 63)
        self.cell(0, 10, '3. Waste Segregation', new_x='LMARGIN', new_y='NEXT')
        self.ln(4)

        cols = ['Campus', 'Total', 'General', 'Recyclable', 'Hazardous']
        widths = [40, 30, 30, 35, 30]

        self.set_font('Helvetica', 'B', 8)
        self.set_fill_color(24, 14, 63)
        self.set_text_color(255, 255, 255)
        for lbl, w in zip(cols, widths):
            self.cell(w, 7, lbl, border=1, fill=True, align='C')
        self.ln()

        self.set_font('Helvetica', '', 8)
        self.set_text_color(0, 0, 0)
        for sheet in region_cfg['sheets']:
            w = waste_data.get(sheet, {})
            total = w.get('Total Waste', 0)
            general = w.get('General Waste', 0)
            recyclable = sum(w.get(c, 0) for c in RECYCLABLE_COLS)
            hazardous = w.get('Hazardous', 0)
            self.cell(40, 6, sheet, border=1)
            self.cell(30, 6, f'{total:.1f}', border=1, align='C')
            self.cell(30, 6, f'{general:.1f}', border=1, align='C')
            self.cell(35, 6, f'{recyclable:.1f}', border=1, align='C')
            self.cell(30, 6, f'{hazardous:.1f}', border=1, align='C')
            self.ln()


def generate_pdf(region_name, month_name, year, token):
    region_cfg = REGIONS.get(region_name)
    if not region_cfg:
        raise ValueError(f'Unknown region: {region_name}')

    period = f'{month_name} {year}'
    print(f'Generating PDF report for {region_name} - {period}')

    kpi_data = fetch_kpi_data(token, month_name)
    waste_data = fetch_waste_data(token, month_name)
    region_data = read_region_data(kpi_data, region_cfg)

    pdf = KPIReport(region_name, period)
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)

    pdf.cover_page(region_cfg)
    pdf.kpi_table(region_data, region_cfg)
    pdf.pillar_breakdown(region_data, region_cfg)
    pdf.waste_section(waste_data, region_cfg)

    # Recommendations page
    pdf.add_page()
    pdf.set_font('Helvetica', 'B', 16)
    pdf.set_text_color(24, 14, 63)
    pdf.cell(0, 10, '4. Recommendations & Action Items', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(4)
    pdf.set_font('Helvetica', 'I', 11)
    pdf.set_text_color(89, 89, 89)
    pdf.cell(0, 8, '(To be completed by the EHS team)', new_x='LMARGIN', new_y='NEXT')

    return pdf.output()


# ── HTTP Handler ──

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        region = qs.get('region', [None])[0]
        month = qs.get('month', [None])[0]
        year = qs.get('year', ['2026'])[0]

        if not region or not month:
            self.send_response(400)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'region and month required'}).encode())
            return

        token = os.environ.get('SMARTSHEET_TOKEN', '')
        if not token:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'SMARTSHEET_TOKEN not set'}).encode())
            return

        try:
            pdf_bytes = generate_pdf(region, month, year, token)
            filename = f'KPI_Report_{region.replace(" ", "_")}_{month}_{year}.pdf'
            self.send_response(200)
            self.send_header('Content-Type', 'application/pdf')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(pdf_bytes)))
            self.end_headers()
            self.wfile.write(pdf_bytes)
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())
