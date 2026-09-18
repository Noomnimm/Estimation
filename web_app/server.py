from __future__ import annotations

import cgi
import base64
import hashlib
import hmac
import json
import mimetypes
import os
import traceback
import time
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pandas as pd

from material_logic import MaterialWorkbook, SIZE_COL, HEAD_COL, MATERIAL_COL, CODE_COL, QTY_COL
from cloud_store import GoogleSheetProjectStore


ROOT = Path(__file__).resolve().parent
UPLOADS = ROOT / "uploads"
OUTPUTS = ROOT / "outputs"
STATIC = ROOT / "static"
DEFAULT_BASE = ROOT.parent / "Newdata.xlsx"
DEFAULT_SET = ROOT.parent / "New folder" / "Allset.xlsx"

WORKBOOK = MaterialWorkbook()
CLOUD_STORE = GoogleSheetProjectStore()
ADMIN_USERNAME = os.environ.get("BASE_ADMIN_USERNAME", "").strip()
ADMIN_PASSWORD = os.environ.get("BASE_ADMIN_PASSWORD", "")
ADMIN_SECRET = os.environ.get("BASE_ADMIN_SESSION_SECRET", "").strip() or ADMIN_PASSWORD
if DEFAULT_BASE.exists():
    WORKBOOK.load_base(DEFAULT_BASE)
if DEFAULT_SET.exists():
    WORKBOOK.load_set(DEFAULT_SET)


def reload_approved_base() -> None:
    if not DEFAULT_BASE.exists():
        return
    WORKBOOK.load_base(DEFAULT_BASE)
    try:
        approved = CLOUD_STORE.list_approved_base_rows()
    except Exception:
        traceback.print_exc()
        return
    if not approved:
        return
    grouped: dict[str, list[dict]] = {}
    for row in approved:
        grouped.setdefault(str(row.get("request_id", "")), []).append(row)
    for request_rows in grouped.values():
        first = request_rows[0]
        size, head = str(first.get("size", "")).strip(), str(first.get("head", "")).strip()
        if first.get("action") == "replace":
            WORKBOOK.base_df = WORKBOOK.base_df[
                ~((WORKBOOK.base_df[SIZE_COL].astype(str).str.strip() == size) & (WORKBOOK.base_df[HEAD_COL].astype(str).str.strip() == head))
            ]
        additions = [{
            SIZE_COL: size, HEAD_COL: head, MATERIAL_COL: str(row.get("material", "")).strip(),
            CODE_COL: str(row.get("code", "")).strip(), QTY_COL: float(row.get("quantity", 0)),
        } for row in request_rows]
        WORKBOOK.base_df = pd.concat([WORKBOOK.base_df, pd.DataFrame(additions)], ignore_index=True)


def create_admin_token() -> str:
    expires = int(time.time()) + 8 * 60 * 60
    payload = f"{ADMIN_USERNAME}:{expires}"
    signature = hmac.new(ADMIN_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
    return base64.urlsafe_b64encode(f"{payload}:{signature}".encode()).decode()


def verify_admin_token(token: str) -> str:
    if not ADMIN_USERNAME or not ADMIN_PASSWORD or not ADMIN_SECRET:
        raise PermissionError("ยังไม่ได้ตั้งค่าบัญชี Admin บน Render")
    try:
        decoded = base64.urlsafe_b64decode(token.encode()).decode()
        username, expires_text, signature = decoded.rsplit(":", 2)
        payload = f"{username}:{expires_text}"
        expected = hmac.new(ADMIN_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
        if username != ADMIN_USERNAME or int(expires_text) < int(time.time()) or not hmac.compare_digest(signature, expected):
            raise ValueError
    except Exception as exc:
        raise PermissionError("เซสชัน Admin ไม่ถูกต้องหรือหมดอายุ") from exc
    return username


reload_approved_base()


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
        if parsed.path == "/api/status":
            self.handle_json(WORKBOOK.get_status)
            return
        if parsed.path == "/api/cloud-config":
            self.handle_json(CLOUD_STORE.public_config)
            return
        if parsed.path == "/api/cloud-projects":
            self.cloud_projects()
            return
        if parsed.path == "/api/base-admin/config":
            self.send_json({"configured": bool(ADMIN_USERNAME and ADMIN_PASSWORD and CLOUD_STORE.service_configured)})
            return
        if parsed.path == "/api/base-requests/admin":
            self.list_base_requests()
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
            "/api/export-pages": self.export_pages,
            "/api/export-page-hardware": self.export_page_hardware,
            "/api/export-page-insulators": self.export_page_insulators,
            "/api/export-page-crossarms": self.export_page_crossarms,
            "/api/cloud-projects": self.save_cloud_project,
            "/api/cloud-projects/delete": self.delete_cloud_project,
            "/api/base-requests": self.submit_base_request,
            "/api/base-admin/login": self.admin_login,
            "/api/base-requests/review": self.review_base_request,
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

    def export_pages(self) -> None:
        try:
            payload = self.read_json()
            data = WORKBOOK.export_page_summary(payload.get("pages", []))
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="page_summary.xlsx"')
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def export_page_hardware(self) -> None:
        try:
            data = WORKBOOK.export_page_hardware(self.read_json().get("pages", []))
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="page_insulators_hardware.xlsx"')
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def export_page_insulators(self) -> None:
        try:
            data = WORKBOOK.export_page_insulators(self.read_json().get("pages", []))
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="page_insulators.xlsx"')
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def export_page_crossarms(self) -> None:
        try:
            data = WORKBOOK.export_page_crossarms(self.read_json().get("pages", []))
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="page_crossarms.xlsx"')
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def submit_base_request(self) -> None:
        try:
            request = CLOUD_STORE.submit_base_request(self.read_json())
            self.send_json({"request": request}, HTTPStatus.CREATED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def admin_login(self) -> None:
        try:
            payload = self.read_json()
            username, password = str(payload.get("username", "")), str(payload.get("password", ""))
            if not ADMIN_USERNAME or not ADMIN_PASSWORD:
                raise PermissionError("ยังไม่ได้ตั้งค่าบัญชี Admin บน Render")
            if not hmac.compare_digest(username, ADMIN_USERNAME) or not hmac.compare_digest(password, ADMIN_PASSWORD):
                raise PermissionError("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")
            self.send_json({"token": create_admin_token(), "username": ADMIN_USERNAME})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)

    def list_base_requests(self) -> None:
        try:
            verify_admin_token(self.admin_token())
            self.send_json({"requests": CLOUD_STORE.list_base_requests()})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def review_base_request(self) -> None:
        try:
            admin = verify_admin_token(self.admin_token())
            payload = self.read_json()
            request = CLOUD_STORE.review_base_request(
                str(payload.get("requestId", "")), bool(payload.get("approve")), admin,
                str(payload.get("note", "")),
            )
            if payload.get("approve"):
                reload_approved_base()
            self.send_json({"request": request})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def admin_token(self) -> str:
        header = self.headers.get("Authorization", "")
        return header.removeprefix("Bearer ").strip() if header.startswith("Bearer ") else ""

    def cloud_projects(self) -> None:
        try:
            user = CLOUD_STORE.verify_user(self.bearer_token())
            self.send_json({"projects": CLOUD_STORE.list_projects(), "user": user})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            traceback.print_exc()
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def save_cloud_project(self) -> None:
        try:
            user = CLOUD_STORE.verify_user(self.bearer_token())
            project = CLOUD_STORE.save_project(self.read_json().get("project", {}), user)
            self.send_json({"project": project, "user": user})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            traceback.print_exc()
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def delete_cloud_project(self) -> None:
        try:
            user = CLOUD_STORE.verify_user(self.bearer_token())
            CLOUD_STORE.delete_project(str(self.read_json().get("projectId", "")).strip(), user)
            self.send_json({"deleted": True})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            traceback.print_exc()
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def bearer_token(self) -> str:
        authorization = self.headers.get("Authorization", "")
        prefix = "Bearer "
        return authorization[len(prefix):].strip() if authorization.startswith(prefix) else ""

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
        self.send_header("Cache-Control", "no-store")
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
        self.send_header("Cache-Control", "no-store")
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
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer((host, port), AppHandler)
    print(f"Material Calculator Web is running at http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
