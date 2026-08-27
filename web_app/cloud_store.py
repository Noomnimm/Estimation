from __future__ import annotations

import json
import os
import threading
from datetime import datetime, timezone
from typing import Any


SHEET_NAME = "Projects"
HEADERS = [
    "project_id",
    "project_name",
    "plan_number",
    "pages_json",
    "results_json",
    "result_meta",
    "created_at",
    "updated_at",
    "created_by",
    "updated_by",
    "deleted_at",
]


class GoogleSheetProjectStore:
    def __init__(self) -> None:
        self.spreadsheet_id = os.environ.get("GOOGLE_SHEET_ID", "").strip()
        self.client_id = os.environ.get("GOOGLE_CLIENT_ID", "").strip()
        self.allowed_domain = os.environ.get("GOOGLE_ALLOWED_DOMAIN", "").strip().lower().lstrip("@")
        self.allowed_emails = {
            email.strip().lower()
            for email in os.environ.get("GOOGLE_ALLOWED_EMAILS", "").split(",")
            if email.strip()
        }
        self._credentials_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        self._service = None
        self._lock = threading.RLock()

    @property
    def configured(self) -> bool:
        return bool(self.spreadsheet_id and self.client_id and self._credentials_json)

    def public_config(self) -> dict[str, Any]:
        return {
            "configured": self.configured,
            "clientId": self.client_id if self.configured else "",
            "allowedDomain": self.allowed_domain,
        }

    def verify_user(self, credential: str) -> dict[str, str]:
        if not self.configured:
            raise ValueError("ยังไม่ได้ตั้งค่าการเชื่อมต่อ Google Cloud")
        if not credential:
            raise PermissionError("กรุณาเข้าสู่ระบบด้วย Google")

        from google.auth.transport import requests as google_requests
        from google.oauth2 import id_token

        try:
            info = id_token.verify_oauth2_token(credential, google_requests.Request(), self.client_id)
        except Exception as exc:
            raise PermissionError("เซสชัน Google หมดอายุ กรุณาเข้าสู่ระบบใหม่") from exc

        email = str(info.get("email", "")).strip().lower()
        if not email or not info.get("email_verified"):
            raise PermissionError("บัญชี Google นี้ยังไม่ได้ยืนยันอีเมล")
        if self.allowed_emails and email not in self.allowed_emails:
            raise PermissionError("อีเมลนี้ไม่ได้รับอนุญาตให้ใช้งาน")
        if self.allowed_domain and not email.endswith(f"@{self.allowed_domain}"):
            raise PermissionError(f"อนุญาตเฉพาะบัญชี @{self.allowed_domain}")
        if not self.allowed_emails and not self.allowed_domain:
            raise PermissionError("ยังไม่ได้กำหนดรายชื่อหรือโดเมนผู้ใช้งาน")
        return {"email": email, "name": str(info.get("name", email)).strip()}

    def list_projects(self) -> list[dict[str, Any]]:
        with self._lock:
            rows = self._read_rows()
        projects = [self._row_to_project(row) for row in rows if not row.get("deleted_at")]
        return sorted(projects, key=lambda project: project.get("updatedAt", ""), reverse=True)

    def save_project(self, project: dict[str, Any], user: dict[str, str]) -> dict[str, Any]:
        project_id = str(project.get("id", "")).strip()
        name = str(project.get("name", "")).strip()
        if not project_id or not name:
            raise ValueError("ข้อมูล Project ID หรือชื่องานไม่ครบ")

        now = datetime.now(timezone.utc).isoformat()
        with self._lock:
            rows = self._read_rows()
            existing = next((row for row in rows if row.get("project_id") == project_id), None)
            record = {
                "project_id": project_id,
                "project_name": name,
                "plan_number": str(project.get("planNumber", "")).strip(),
                "pages_json": json.dumps(project.get("pages", []), ensure_ascii=False, separators=(",", ":")),
                "results_json": json.dumps(project.get("results", []), ensure_ascii=False, separators=(",", ":")),
                "result_meta": str(project.get("resultMeta", "")),
                "created_at": (existing or {}).get("created_at") or str(project.get("createdAt", "")) or now,
                "updated_at": now,
                "created_by": (existing or {}).get("created_by") or user["email"],
                "updated_by": user["email"],
                "deleted_at": "",
            }
            self._write_record(record, existing.get("_row_number") if existing else None)
        return self._row_to_project(record)

    def delete_project(self, project_id: str, user: dict[str, str]) -> None:
        with self._lock:
            rows = self._read_rows()
            existing = next((row for row in rows if row.get("project_id") == project_id and not row.get("deleted_at")), None)
            if not existing:
                raise ValueError("ไม่พบงานที่ต้องการลบ")
            existing["deleted_at"] = datetime.now(timezone.utc).isoformat()
            existing["updated_by"] = user["email"]
            self._write_record(existing, existing["_row_number"])

    def _get_service(self):
        if self._service is None:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build

            info = json.loads(self._credentials_json)
            credentials = service_account.Credentials.from_service_account_info(
                info,
                scopes=["https://www.googleapis.com/auth/spreadsheets"],
            )
            self._service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
        return self._service

    def _ensure_sheet(self) -> None:
        service = self._get_service()
        metadata = service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
        titles = {sheet["properties"]["title"] for sheet in metadata.get("sheets", [])}
        if SHEET_NAME not in titles:
            service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": [{"addSheet": {"properties": {"title": SHEET_NAME}}}]},
            ).execute()

        values = service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{SHEET_NAME}'!A1:K1",
        ).execute().get("values", [])
        if not values:
            service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{SHEET_NAME}'!A1:K1",
                valueInputOption="RAW",
                body={"values": [HEADERS]},
            ).execute()
        elif values[0] != HEADERS:
            raise ValueError(f"หัวตารางในแท็บ {SHEET_NAME} ไม่ตรงกับรูปแบบของระบบ")

    def _read_rows(self) -> list[dict[str, Any]]:
        self._ensure_sheet()
        values = self._get_service().spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{SHEET_NAME}'!A2:K",
        ).execute().get("values", [])
        rows = []
        for row_number, values_row in enumerate(values, start=2):
            padded = values_row + [""] * (len(HEADERS) - len(values_row))
            row = dict(zip(HEADERS, padded[:len(HEADERS)]))
            if row.get("project_id"):
                row["_row_number"] = row_number
                rows.append(row)
        return rows

    def _write_record(self, record: dict[str, Any], row_number: int | None) -> None:
        values = [[record.get(header, "") for header in HEADERS]]
        service = self._get_service().spreadsheets().values()
        if row_number:
            service.update(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{SHEET_NAME}'!A{row_number}:K{row_number}",
                valueInputOption="RAW",
                body={"values": values},
            ).execute()
        else:
            service.append(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{SHEET_NAME}'!A:K",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": values},
            ).execute()

    @staticmethod
    def _row_to_project(row: dict[str, Any]) -> dict[str, Any]:
        def parse_json(value: str, fallback):
            try:
                return json.loads(value) if value else fallback
            except json.JSONDecodeError:
                return fallback

        return {
            "id": row.get("project_id", ""),
            "name": row.get("project_name", ""),
            "planNumber": row.get("plan_number", ""),
            "pages": parse_json(row.get("pages_json", ""), []),
            "results": parse_json(row.get("results_json", ""), []),
            "resultMeta": row.get("result_meta", ""),
            "createdAt": row.get("created_at", ""),
            "updatedAt": row.get("updated_at", ""),
            "createdBy": row.get("created_by", ""),
            "updatedBy": row.get("updated_by", ""),
        }
