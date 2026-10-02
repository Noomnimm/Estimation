from __future__ import annotations

import cgi
import base64
import hashlib
import hmac
import html
import json
import mimetypes
import os
import threading
import traceback
import time
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pandas as pd

from material_logic import MaterialWorkbook, SIZE_COL, HEAD_COL, MATERIAL_COL, CODE_COL, QTY_COL, DEPARTMENT_COL, DEFAULT_DEPARTMENT, INSULATOR_UPRIGHT_COL, INSULATOR_HORIZONTAL_COL
from cloud_store import GoogleSheetProjectStore


ROOT = Path(__file__).resolve().parent
UPLOADS = ROOT / "uploads"
OUTPUTS = ROOT / "outputs"
STATIC = ROOT / "static"
DEFAULT_BASE = ROOT.parent / "Newdata.xlsx"
DEFAULT_SET = ROOT.parent / "New folder (2)" / "Allset.xlsx"
DEFAULT_TRANSFORMER_BASE = ROOT.parent / "แผนกหม้อแปลง.xlsx"
DEFAULT_TRANSMISSION = ROOT.parent / "สายส่ง 115kV.xlsx"

WORKBOOK = MaterialWorkbook()
CLOUD_STORE = GoogleSheetProjectStore()
HEAD_IMAGES: dict[tuple[str, str, str], dict[str, str]] = {}
HEAD_EDITORS: dict[tuple[str, str, str], str] = {}
ADMIN_USERNAME = os.environ.get("BASE_ADMIN_USERNAME", "").strip()
ADMIN_PASSWORD = os.environ.get("BASE_ADMIN_PASSWORD", "")
ADMIN_SECRET = os.environ.get("BASE_ADMIN_SESSION_SECRET", "").strip() or ADMIN_PASSWORD
if DEFAULT_BASE.exists():
    WORKBOOK.load_base(DEFAULT_BASE)
if DEFAULT_TRANSFORMER_BASE.exists():
    WORKBOOK.load_department_base(DEFAULT_TRANSFORMER_BASE, "แผนกหม้อแปลง", "หม้อแปลง")
if DEFAULT_TRANSMISSION.exists():
    WORKBOOK.load_keycode_catalog(DEFAULT_TRANSMISSION, "แผนกสายส่ง")
if DEFAULT_SET.exists():
    WORKBOOK.load_set(DEFAULT_SET)


def reload_approved_base() -> None:
    HEAD_IMAGES.clear()
    HEAD_EDITORS.clear()
    if not DEFAULT_BASE.exists():
        return
    WORKBOOK.load_base(DEFAULT_BASE)
    if DEFAULT_TRANSFORMER_BASE.exists():
        WORKBOOK.load_department_base(DEFAULT_TRANSFORMER_BASE, "แผนกหม้อแปลง", "หม้อแปลง")
    if DEFAULT_TRANSMISSION.exists():
        WORKBOOK.load_keycode_catalog(DEFAULT_TRANSMISSION, "แผนกสายส่ง")
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
        department = str(first.get("department", "")).strip() or DEFAULT_DEPARTMENT
        action = str(first.get("action", "add")).strip()
        if action in {"replace", "rename", "delete"}:
            remove_size = str(first.get("source_size", "")).strip() or size
            remove_head = str(first.get("source_head", "")).strip() or head
            WORKBOOK.base_df = WORKBOOK.base_df[
                ~((WORKBOOK.base_df[SIZE_COL].astype(str).str.strip() == remove_size)
                  & (WORKBOOK.base_df[HEAD_COL].astype(str).str.strip() == remove_head)
                  & (WORKBOOK.base_df[DEPARTMENT_COL].astype(str).str.strip() == department))
            ]
            HEAD_IMAGES.pop((department, remove_size, remove_head), None)
            HEAD_EDITORS.pop((department, remove_size, remove_head), None)
        additions = [{
            SIZE_COL: size, HEAD_COL: head, MATERIAL_COL: str(row.get("material", "")).strip(),
            CODE_COL: str(row.get("code", "")).strip(), QTY_COL: float(row.get("quantity", 0)),
            DEPARTMENT_COL: department,
            INSULATOR_UPRIGHT_COL: pd.NA if row.get("insulator_upright", "") == "" else row.get("insulator_upright"),
            INSULATOR_HORIZONTAL_COL: pd.NA if row.get("insulator_horizontal", "") == "" else row.get("insulator_horizontal"),
        } for row in request_rows if action != "delete"]
        if additions:
            WORKBOOK.base_df = pd.concat([WORKBOOK.base_df, pd.DataFrame(additions)], ignore_index=True)
            HEAD_EDITORS[(department, size, head)] = str(first.get("last_modified_by", "")).strip() or "Admin"
        image_file_id = str(first.get("image_file_id", "")).strip()
        if image_file_id:
            HEAD_IMAGES[(department, size, head)] = {
                "id": image_file_id,
                "name": str(first.get("image_name", "")).strip(),
                "mimeType": str(first.get("image_mime_type", "")).strip(),
            }


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


class AppHandler(SimpleHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_file(STATIC / "index.html", "text/html; charset=utf-8")
            return
        if parsed.path == "/api/heads":
            query = parse_qs(parsed.query)
            size = query.get("size", [""])[0]
            department = query.get("department", [DEFAULT_DEPARTMENT])[0]
            def heads_response():
                heads = WORKBOOK.get_heads(size, department)
                return {
                    "heads": heads,
                    "insulatorRates": {head: WORKBOOK.get_insulator_rate(size, head, department) for head in heads},
                    "headImages": {head: HEAD_IMAGES[(department, str(size).strip(), head)] for head in heads if (department, str(size).strip(), head) in HEAD_IMAGES},
                    "lastEditors": {head: HEAD_EDITORS.get((department, str(size).strip(), head), "Admin") for head in heads},
                }
            self.handle_json(heads_response)
            return
        if parsed.path == "/api/sizes":
            query = parse_qs(parsed.query)
            department = query.get("department", [DEFAULT_DEPARTMENT])[0]
            self.handle_json(lambda: {"sizes": WORKBOOK.get_sizes(department)})
            return
        if parsed.path == "/api/base-entry":
            query = parse_qs(parsed.query)
            self.base_entry(query.get("size", [""])[0], query.get("head", [""])[0], query.get("department", [DEFAULT_DEPARTMENT])[0])
            return
        if parsed.path == "/api/head-image":
            query = parse_qs(parsed.query)
            self.head_image(query.get("id", [""])[0])
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
            self.send_json({
                "configured": bool(ADMIN_USERNAME and ADMIN_PASSWORD and CLOUD_STORE.service_configured),
                "driveOAuthReady": CLOUD_STORE.drive_oauth_ready,
                "driveConnected": CLOUD_STORE.drive_oauth_configured,
            })
            return
        if parsed.path == "/api/google-drive/oauth/callback":
            self.google_drive_oauth_callback(parsed)
            return
        if parsed.path == "/api/base-requests/admin":
            self.list_base_requests()
            return
        if parsed.path == "/api/base-requests/pending":
            self.handle_json(CLOUD_STORE.public_base_request_overview)
            return
        if parsed.path == "/api/base-admin/cloud-trash":
            self.list_admin_cloud_trash()
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
            "/api/set-components": self.set_components,
            "/api/export": self.export_summary,
            "/api/export-pages": self.export_pages,
            "/api/export-page-hardware": self.export_page_hardware,
            "/api/export-page-insulators": self.export_page_insulators,
            "/api/export-page-crossarms": self.export_page_crossarms,
            "/api/cloud-projects": self.save_cloud_project,
            "/api/cloud-projects/delete": self.delete_cloud_project,
            "/api/base-admin/cloud-trash/restore": self.restore_admin_cloud_project,
            "/api/base-admin/cloud-trash/purge": self.purge_admin_cloud_project,
            "/api/cloud-folders": self.create_cloud_folder,
            "/api/cloud-folders/delete": self.delete_cloud_folder,
            "/api/cloud-projects/move": self.move_cloud_project,
            "/api/base-requests": self.submit_base_request,
            "/api/base-admin/login": self.admin_login,
            "/api/base-requests/review": self.review_base_request,
            "/api/base-requests/clear-approved": self.clear_approved_requests,
            "/api/base-images/upload": self.upload_base_image,
            "/api/google-drive/oauth/start": self.google_drive_oauth_start,
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

    def base_entry(self, size: str, head: str, department: str = DEFAULT_DEPARTMENT) -> None:
        try:
            if WORKBOOK.base_df is None:
                raise ValueError("ยังไม่ได้โหลด BaseData")
            size, head = str(size).strip(), str(head).strip()
            matches = WORKBOOK.base_df[
                (WORKBOOK.base_df[SIZE_COL].astype(str).str.strip() == size)
                & (WORKBOOK.base_df[HEAD_COL].astype(str).str.strip() == head)
                & (WORKBOOK.base_df[DEPARTMENT_COL].astype(str).str.strip() == department)
            ]
            rows = [{
                "material": str(row[MATERIAL_COL]).strip(),
                "code": str(row[CODE_COL]).strip(),
                "quantity": float(row[QTY_COL]),
            } for _, row in matches.iterrows()]
            upright, horizontal = WORKBOOK.get_insulator_rate(size, head, department)
            self.send_json({
                "size": size, "head": head, "rows": rows,
                "insulatorUpright": upright, "insulatorHorizontal": horizontal,
                "image": HEAD_IMAGES.get((department, size, head), {}),
                "lastModifiedBy": HEAD_EDITORS.get((department, size, head), "Admin"),
            })
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def calculate(self) -> None:
        payload = self.read_json()
        self.handle_json(lambda: WORKBOOK.calculate(payload.get("pages", []), payload.get("structurePages", [])))

    def expand_set(self) -> None:
        self.handle_json(WORKBOOK.expand_set)

    def set_components(self) -> None:
        payload = self.read_json()
        self.handle_json(lambda: WORKBOOK.get_set_components(payload.get("code", ""), payload.get("quantity", 1)))

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
            data = WORKBOOK.export_page_hardware_combined(self.read_json().get("pages", []))
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

    def upload_base_image(self) -> None:
        try:
            payload = self.read_json()
            data_url = str(payload.get("data", ""))
            if "," not in data_url:
                raise ValueError("ข้อมูลรูปไม่ถูกต้อง")
            header, encoded = data_url.split(",", 1)
            mime_type = header.removeprefix("data:").split(";", 1)[0].strip().lower()
            try:
                content = base64.b64decode(encoded, validate=True)
            except Exception as exc:
                raise ValueError("ข้อมูลรูปไม่ถูกต้อง") from exc
            image = CLOUD_STORE.upload_head_image(content, str(payload.get("name", "head-image")), mime_type)
            self.send_json({"image": image}, HTTPStatus.CREATED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def head_image(self, file_id: str) -> None:
        try:
            content, mime_type, name = CLOUD_STORE.download_head_image(str(file_id).strip())
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", mime_type)
            # BaseHTTPRequestHandler encodes headers as Latin-1. Keep the inline
            # header ASCII-only so Thai (or other Unicode) Drive filenames do
            # not break the response after it has already started.
            self.send_header("Content-Disposition", "inline")
            self.send_header("Cache-Control", "private, max-age=3600")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.FORBIDDEN)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.NOT_FOUND)

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

    @staticmethod
    def google_drive_redirect_uri() -> str:
        base_url = os.environ.get("RENDER_EXTERNAL_URL", "http://localhost:8000").strip().rstrip("/")
        return f"{base_url}/api/google-drive/oauth/callback"

    def google_drive_oauth_start(self) -> None:
        try:
            verify_admin_token(self.admin_token())
            if not CLOUD_STORE.drive_oauth_ready:
                raise ValueError("กรุณาตั้งค่า GOOGLE_CLIENT_SECRET บน Render ก่อน")
            from google_auth_oauthlib.flow import Flow
            flow = Flow.from_client_config({
                "web": {
                    "client_id": CLOUD_STORE.client_id,
                    "client_secret": CLOUD_STORE.oauth_client_secret,
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                }
            }, scopes=["https://www.googleapis.com/auth/drive"], state=self.admin_token(),
                autogenerate_code_verifier=False)
            flow.redirect_uri = self.google_drive_redirect_uri()
            authorization_url, _ = flow.authorization_url(
                access_type="offline", prompt="consent", include_granted_scopes="true",
            )
            self.send_json({"authorizationUrl": authorization_url})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def google_drive_oauth_callback(self, parsed) -> None:
        try:
            query = parse_qs(parsed.query)
            state = query.get("state", [""])[0]
            verify_admin_token(state)
            if query.get("error"):
                raise ValueError(f"Google ปฏิเสธการเชื่อมต่อ: {query['error'][0]}")
            from google_auth_oauthlib.flow import Flow
            flow = Flow.from_client_config({
                "web": {
                    "client_id": CLOUD_STORE.client_id,
                    "client_secret": CLOUD_STORE.oauth_client_secret,
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                }
            }, scopes=["https://www.googleapis.com/auth/drive"], state=state,
                autogenerate_code_verifier=False)
            flow.redirect_uri = self.google_drive_redirect_uri()
            flow.fetch_token(code=query.get("code", [""])[0])
            refresh_token = flow.credentials.refresh_token
            if not refresh_token:
                raise ValueError("Google ไม่ได้ส่ง Refresh Token กรุณาถอนสิทธิ์แอปแล้วเชื่อมต่อใหม่")
            safe_token = html.escape(refresh_token)
            body = f"""<!doctype html><html lang=\"th\"><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width\"><title>เชื่อม Google Drive</title><style>body{{font-family:sans-serif;background:#f5eff9;padding:32px;color:#25113d}}main{{max-width:760px;margin:auto;background:white;padding:28px;border-radius:18px}}textarea{{width:100%;min-height:130px;box-sizing:border-box}}code{{font-weight:bold}}button{{padding:10px 18px;background:#5b2584;color:white;border:0;border-radius:9px}}</style><main><h1>เชื่อมต่อสำเร็จ</h1><p>คัดลอกค่าด้านล่างไปใส่ใน Render Environment โดยใช้ Key <code>GOOGLE_DRIVE_REFRESH_TOKEN</code></p><textarea id=\"token\" readonly>{safe_token}</textarea><p><button onclick=\"navigator.clipboard.writeText(document.getElementById('token').value)\">คัดลอก Token</button></p><p>จากนั้น Save และ Deploy ใหม่ แล้วปิดหน้านี้ได้เลย</p></main></html>"""
            encoded = body.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(encoded)))
            self.end_headers()
            self.wfile.write(encoded)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def list_base_requests(self) -> None:
        try:
            verify_admin_token(self.admin_token())
            self.send_json({"requests": CLOUD_STORE.list_base_requests()})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def list_admin_cloud_trash(self) -> None:
        try:
            verify_admin_token(self.admin_token())
            self.send_json({"projects": CLOUD_STORE.list_all_deleted_projects()})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def restore_admin_cloud_project(self) -> None:
        try:
            verify_admin_token(self.admin_token())
            payload = self.read_json()
            project = CLOUD_STORE.restore_project_as_admin(
                str(payload.get("projectId", "")).strip(), str(payload.get("ownerEmail", "")).strip()
            )
            self.send_json({"project": project})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def purge_admin_cloud_project(self) -> None:
        try:
            verify_admin_token(self.admin_token())
            payload = self.read_json()
            CLOUD_STORE.permanently_delete_project_as_admin(
                str(payload.get("projectId", "")).strip(), str(payload.get("ownerEmail", "")).strip()
            )
            self.send_json({"deleted": True})
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
                str(payload.get("note", "")), payload.get("edits"),
            )
            if payload.get("approve"):
                reload_approved_base()
            self.send_json({"request": request})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def clear_approved_requests(self) -> None:
        try:
            verify_admin_token(self.admin_token())
            self.send_json({"cleared": CLOUD_STORE.clear_approved_requests()})
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
            self.send_json({"projects": CLOUD_STORE.list_projects(user), "folders": CLOUD_STORE.list_project_folders(user), "user": user})
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

    def restore_cloud_project(self) -> None:
        try:
            user = CLOUD_STORE.verify_user(self.bearer_token())
            project = CLOUD_STORE.restore_project(str(self.read_json().get("projectId", "")).strip(), user)
            self.send_json({"project": project})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def purge_cloud_project(self) -> None:
        try:
            user = CLOUD_STORE.verify_user(self.bearer_token())
            CLOUD_STORE.permanently_delete_project(str(self.read_json().get("projectId", "")).strip(), user)
            self.send_json({"deleted": True})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def create_cloud_folder(self) -> None:
        try:
            user = CLOUD_STORE.verify_user(self.bearer_token())
            folder = CLOUD_STORE.create_project_folder(self.read_json().get("name", ""), user)
            self.send_json({"folder": folder})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def move_cloud_project(self) -> None:
        try:
            user = CLOUD_STORE.verify_user(self.bearer_token())
            payload = self.read_json()
            project = CLOUD_STORE.move_project(str(payload.get("projectId", "")).strip(), str(payload.get("folderId", "")).strip(), user)
            self.send_json({"project": project})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)

    def delete_cloud_folder(self) -> None:
        try:
            user = CLOUD_STORE.verify_user(self.bearer_token())
            CLOUD_STORE.delete_project_folder(str(self.read_json().get("folderId", "")).strip(), user)
            self.send_json({"deleted": True})
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.UNAUTHORIZED)
        except Exception as exc:
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
    print(f"Material Calculator Web is running at http://{host}:{port}", flush=True)
    threading.Thread(target=reload_approved_base, name="approved-base-loader", daemon=True).start()
    server.serve_forever()


if __name__ == "__main__":
    main()
