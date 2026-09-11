"""Debug endpoint - test Smartsheet API connectivity from Python runtime."""
import os, json
from http.server import BaseHTTPRequestHandler
from urllib.request import Request, urlopen

def _ss_fetch(endpoint, token):
    url = f'https://api.smartsheet.com/2.0/{endpoint}'
    req = Request(url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/json'})
    with urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        token = os.environ.get('SMARTSHEET_TOKEN')
        result = {'token_exists': bool(token), 'token_len': len(token) if token else 0}
        test_sheet = '2717764266446724'
        try:
            data = _ss_fetch(f'sheets/{test_sheet}?pageSize=5', token)
            col_map = {c['id']: c['title'] for c in data.get('columns', [])}
            cols = list(col_map.values())
            row_count = len(data.get('rows', []))
            first_row = {}
            if data.get('rows'):
                for cell in data['rows'][0].get('cells', []):
                    title = col_map.get(cell.get('columnId'))
                    if title:
                        first_row[title] = cell.get('displayValue') or cell.get('value') or ''
            result['sheet_test'] = {'status': 'ok', 'total_rows': row_count, 'columns': cols[:10], 'first_row': first_row}
        except Exception as e:
            result['sheet_test'] = {'status': 'error', 'error': str(e)}
        test_report = '8527961731846020'
        try:
            data = _ss_fetch(f'reports/{test_report}?pageSize=5&level=1', token)
            col_map = {}
            for c in data.get('columns', []):
                if c.get('id'): col_map[c['id']] = c['title']
                if c.get('virtualId'): col_map[c['virtualId']] = c['title']
            cols = list(set(col_map.values()))
            row_count = len(data.get('rows', []))
            result['report_test'] = {'status': 'ok', 'total_rows': row_count, 'columns': cols[:10]}
        except Exception as e:
            result['report_test'] = {'status': 'error', 'error': str(e)}
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(result, indent=2).encode())
