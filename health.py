import argparse
import csv
import json
import mimetypes
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse


ROOT = Path(__file__).resolve().parent
WEB_ROOT = ROOT / "web"
DATA_FILE = ROOT / "table_data.csv"
ROWS = 30
COLUMNS = 6


def normalise_rows(rows):
    clean = []

    for row in rows[:ROWS]:
        values = [str(value) if value is not None else "" for value in row[:COLUMNS]]
        values.extend([""] * (COLUMNS - len(values)))
        clean.append(values)

    while len(clean) < ROWS:
        clean.append([""] * COLUMNS)

    return clean


def read_table():
    if not DATA_FILE.exists():
        return normalise_rows([])

    with DATA_FILE.open("r", newline="", encoding="utf-8-sig") as file:
        return normalise_rows(list(csv.reader(file)))


def write_table(rows):
    clean = normalise_rows(rows)
    temp_file = DATA_FILE.with_suffix(".csv.tmp")

    with temp_file.open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerows(clean)

    temp_file.replace(DATA_FILE)
    return clean


def csv_text(rows):
    import io

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerows(normalise_rows(rows))
    return output.getvalue()


class HealthHandler(BaseHTTPRequestHandler):
    server_version = "HealthApp/2.0"

    def do_GET(self):
        parsed = urlparse(self.path)

        if parsed.path == "/api/table":
            self.send_json({"rows": read_table(), "source": "csv"})
            return

        if parsed.path == "/api/export.csv":
            body = csv_text(read_table()).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition", 'attachment; filename="table_data.csv"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        self.serve_static(parsed.path)

    def do_POST(self):
        parsed = urlparse(self.path)

        if parsed.path != "/api/table":
            self.send_error(404)
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            rows = payload.get("rows", [])
            saved = write_table(rows)
        except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
            self.send_json({"error": str(exc)}, status=400)
            return

        self.send_json({"rows": saved, "source": "csv"})

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def send_json(self, payload, status=200):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def serve_static(self, request_path):
        path = unquote(request_path).lstrip("/")
        if path == "":
            path = "index.html"

        target = (WEB_ROOT / path).resolve()

        if WEB_ROOT.resolve() not in target.parents and target != WEB_ROOT.resolve():
            self.send_error(403)
            return

        if target.is_dir():
            target = target / "index.html"

        if not target.exists():
            self.send_error(404)
            return

        mime_type, _ = mimetypes.guess_type(str(target))
        body = target.read_bytes()

        self.send_response(200)
        self.send_header("Content-Type", mime_type or "application/octet-stream")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        print("%s - %s" % (self.address_string(), format % args))


def main():
    parser = argparse.ArgumentParser(description="Run the Health web app.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=8765, type=int)
    args = parser.parse_args()

    address = (args.host, args.port)
    server = ThreadingHTTPServer(address, HealthHandler)
    print(f"Health app running at http://{args.host}:{args.port}")
    print("Use --host 0.0.0.0 to make it reachable from another device on your network.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping Health app.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
