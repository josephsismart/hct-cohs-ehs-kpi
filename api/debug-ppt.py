"""Debug endpoint - dump Smartsheet column names for problem sheets."""
import os, json
from http.server import BaseHTTPRequestHandler
from urllib.request import Request, urlopen

def _ss_fetch(endpoint, token):
    url = f'https://api.smartsheet.com/2.0/{endpoint}'
    req = Request(url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'})
    with urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())

SHEETS_TO_CHECK = {
    'v2_onsite_induction (KPI14)': '3519179394076548',
    'v2_findings_on_time (KPI15)': '1510149721116548',
    'v2_hs_committee (KPI4)': '5093607634587524',
    'v2_planned_training (KPI9)': '4456464805482372',
    'v2_hazard_id (KPI6)': '7524088825204612',
    'v2_drills (KPI12)': '7139786694283140',
}

REPORTS_TO_CHECK = {
    'v2_hs_kpi_report (KPI1)': '5852576405737348',
    'v2_investigation (KPI18)': '5432865759121284',
    'notification (KPI17)': '8527961731846020',
}

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        token = os.environ.get('SMARTSHEET_TOKEN')
        result = {}

        for name, sid in SHEETS_TO_CHECK.items():
            try:
                data = _ss_fetch(f'sheets/{sid}?pageSize=1', token)
                cols = [c['title'] for c in data.get('columns', [])]
                result[name] = {'columns': cols}
            except Exception as e:
                result[name] = {'error': str(e)}

        for name, rid in REPORTS_TO_CHECK.items():
            try:
                data = _ss_fetch(f'reports/{rid}?pageSize=1&level=1', token)
                cols = list(set(c['title'] for c in data.get('columns', [])))
                result[name] = {'columns': sorted(cols)}
            except Exception as e:
                result[name] = {'error': str(e)}

        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(result, indent=2).encode())
