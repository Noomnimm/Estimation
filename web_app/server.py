from __future__ import annotations

import cgi
import json
import mimetypes
import os
import traceback
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from material_logic import MaterialWorkbook


ROOT = Path(__file__).resolve().parent
UPLOADS = ROOT / "uploads"
OUTPUTS = ROOT / "outputs"
STATIC = ROOT / "static"

WORKBOOK = MaterialWorkbook()


class AppHandler(SimpleHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_file(STATIC / "index.html", "text/html; charset=utf-8")
            return
        if parsed.path == "/api/heads":
            query = parse_qs(parsed.query)
            self.handle_json(lambda: {"heads": WORKBOOK.get_heads(query.get("size", [""])[0])})
            return
        if parsed.path.startswith("/static/"):
            target = STATIC / parsed.path.removeprefix("/static/")
            self.send_file(target)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        routes = {
            "/api/load-base": self.load_base,
            "/api/load-set": self.load_set,
            "/api/calculate": self.calculate,
            "/api/expand-set": self.expand_set,
            "/api/export": self.export_summary,
        }
        route = routes.get(parsed.path)
        if route is None:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        route()

    def load_base(self) -> None:
        self.handle_json(lambda: WORKBOOK.load_base(save_upload(self, "base")))

    def load_set(self) -> None:
        self.handle_json(lambda: WORKBOOK.load_set(save_upload(self, "set")))

    def calculate(self) -> None:
        payload = self.read_json()
        self.handle_json(lambda: WORKBOOK.calculate(payload.get("pages", [])))

    def expand_set(self) -> None:
        self.handle_json(WORKBOOK.expand_set)

    def export_summary(self) -> None:
        try:
            data = WORKBOOK.export_summary()
            OUTPUTS.mkdir(parents=True, exist_ok=True)
            path = OUTPUTS / "material_summary_web.xlsx"
            path.write_bytes(data)
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="material_summary_web.xlsx"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def handle_json(self, action) -> None:
        try:
            self.send_json(action())
        except Exception as exc:
            traceback.print_exc()
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length)
        if not raw:
            return {}
        return json.loads(raw.decode("utf-8"))

    def send_json(self, data: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path: Path, content_type: str | None = None) -> None:
        if not path.exists() or not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        data = path.read_bytes()
        guessed = content_type or mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", guessed)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def save_upload(handler: AppHandler, prefix: str) -> Path:
    form = cgi.FieldStorage(
        fp=handler.rfile,
        headers=handler.headers,
        environ={
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": handler.headers.get("Content-Type", ""),
            "CONTENT_LENGTH": handler.headers.get("Content-Length", "0"),
        },
    )
    field = form["file"] if "file" in form else None
    if field is None or not field.filename:
        raise ValueError("ไม่พบไฟล์ที่อัปโหลด")
    filename = Path(field.filename).name
    if not filename.lower().endswith(".xlsx"):
        raise ValueError("รองรับเฉพาะไฟล์ .xlsx")
    UPLOADS.mkdir(parents=True, exist_ok=True)
    path = UPLOADS / f"{prefix}_{filename}"
    data = field.file.read()
    path.write_bytes(data)
    return path


def main() -> None:
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer((host, port), AppHandler)
    print(f"Material Calculator Web is running at http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
