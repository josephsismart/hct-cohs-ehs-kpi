from http.server import BaseHTTPRequestHandler
import json, os

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        diag = {}
        diag["env_ok"] = bool(os.environ.get("SMARTSHEET_TOKEN"))
        diag["tkn_len"] = len(os.environ.get("SMARTSHEET_TOKEN",""))
        try:
            import requests as _rq
            diag["req_ok"] = True
            diag["req_ver"] = _rq.__version__
        except Exception as ex:
            diag["req_ok"] = False
            diag["req_err"] = str(ex)
        if diag.get("req_ok") and diag.get("env_ok"):
            try:
                tk = os.environ["SMARTSHEET_TOKEN"]
                r = _rq.get("https://api.smartsheet.com/2.0/sheets/1325212455882628?pageSize=5", headers={"Authorization": "Bearer " + tk, "Accept": "application/json"}, timeout=15)
                diag["ss_code"] = r.status_code
                if r.ok:
                    dd = r.json()
                    diag["ss_rows"] = len(dd.get("rows",[]))
                    diag["ss_name"] = dd.get("name","")
            except Exception as ex:
                diag["ss_err"] = str(ex)
        self.send_response(200)
        self.send_header("Content-Type","application/json")
        self.send_header("Access-Control-Allow-Origin","*")
        self.end_headers()
        self.wfile.write(json.dumps(diag, indent=2, default=str).encode())
