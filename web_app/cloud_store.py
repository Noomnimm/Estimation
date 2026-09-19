from __future__ import annotations

import json
import os
import threading
import uuid
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

REQUEST_SHEET = "BaseDataRequests"
REQUEST_HEADERS = [
    "request_id", "submitted_at", "submitter_name", "employee_id", "department",
    "action", "size", "head", "rows_json", "note", "status", "reviewed_at",
    "reviewed_by", "review_note",
]
APPROVED_SHEET = "ApprovedBaseData"
APPROVED_HEADERS = [
    "size", "head", "material", "code", "quantity", "action", "request_id",
    "approved_at", "approved_by",
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

    @property
    def service_configured(self) -> bool:
        return bool(self.spreadsheet_id and self._credentials_json)

    def submit_base_request(self, payload: dict[str, Any]) -> dict[str, Any]:
        required = {
            "submitter_name": "ชื่อผู้เสนอ", "employee_id": "รหัสพนักงาน",
            "department": "สังกัด", "size": "ขนาดเสา", "head": "รหัสหัวเสา",
        }
        values = {key: str(payload.get(key, "")).strip() for key in required}
        missing = [label for key, label in required.items() if not values[key]]
        if missing:
            raise ValueError(f"กรุณากรอก {', '.join(missing)}")
        rows = payload.get("rows", [])
        if not isinstance(rows, list) or not rows:
            raise ValueError("กรุณาเพิ่มรายการวัสดุอย่างน้อย 1 รายการ")
        if len(rows) > 100:
            raise ValueError("หนึ่งคำขอเพิ่มรายการวัสดุได้ไม่เกิน 100 รายการ")
        clean_rows = []
        for row in rows:
            material = str(row.get("material", "")).strip()
            code = str(row.get("code", "")).strip()
            try:
                quantity = float(row.get("quantity", 0))
            except (TypeError, ValueError) as exc:
                raise ValueError("จำนวนวัสดุต้องเป็นตัวเลข") from exc
            if not material or not code or quantity == 0:
                raise ValueError("ทุกรายการต้องมีชื่อวัสดุ รหัสพัสดุ/SET และจำนวนที่ไม่เป็นศูนย์")
            clean_rows.append({"material": material, "code": code, "quantity": quantity})
        action = str(payload.get("action", "add")).strip().lower()
        if action not in {"add", "replace"}:
            action = "add"
        original_rows = payload.get("original_rows", []) if action == "replace" else []
        if not isinstance(original_rows, list):
            original_rows = []
        clean_original_rows = []
        for row in original_rows:
            try:
                quantity = float(row.get("quantity", 0))
            except (TypeError, ValueError):
                quantity = 0
            clean_original_rows.append({
                "material": str(row.get("material", "")).strip(),
                "code": str(row.get("code", "")).strip(),
                "quantity": quantity,
            })
        record = {
            "request_id": uuid.uuid4().hex,
            "submitted_at": datetime.now(timezone.utc).isoformat(),
            **values,
            "action": action,
            "rows_json": json.dumps(
                {"new": clean_rows, "original": clean_original_rows},
                ensure_ascii=False, separators=(",", ":"),
            ),
            "note": str(payload.get("note", "")).strip(),
            "status": "pending", "reviewed_at": "", "reviewed_by": "", "review_note": "",
        }
        with self._lock:
            self._ensure_named_sheet(REQUEST_SHEET, REQUEST_HEADERS)
            self._append_named_record(REQUEST_SHEET, REQUEST_HEADERS, record)
        return self._request_to_public(record)

    def list_base_requests(self) -> list[dict[str, Any]]:
        with self._lock:
            rows = self._read_named_rows(REQUEST_SHEET, REQUEST_HEADERS, "request_id")
        return [self._request_to_public(row) for row in reversed(rows)]

    def review_base_request(self, request_id: str, approve: bool, reviewer: str, note: str = "") -> dict[str, Any]:
        with self._lock:
            rows = self._read_named_rows(REQUEST_SHEET, REQUEST_HEADERS, "request_id")
            record = next((row for row in rows if row["request_id"] == request_id), None)
            if not record:
                raise ValueError("ไม่พบคำขอที่ต้องการตรวจ")
            if record.get("status") != "pending":
                raise ValueError("คำขอนี้ถูกตรวจแล้ว")
            now = datetime.now(timezone.utc).isoformat()
            record.update({
                "status": "approved" if approve else "rejected", "reviewed_at": now,
                "reviewed_by": reviewer, "review_note": str(note).strip(),
            })
            self._update_named_record(REQUEST_SHEET, REQUEST_HEADERS, record, record["_row_number"])
            if approve:
                self._ensure_named_sheet(APPROVED_SHEET, APPROVED_HEADERS)
                stored_rows = json.loads(record["rows_json"])
                new_rows = stored_rows.get("new", []) if isinstance(stored_rows, dict) else stored_rows
                for item in new_rows:
                    self._append_named_record(APPROVED_SHEET, APPROVED_HEADERS, {
                        "size": record["size"], "head": record["head"],
                        "material": item["material"], "code": item["code"], "quantity": item["quantity"],
                        "action": record["action"], "request_id": request_id,
                        "approved_at": now, "approved_by": reviewer,
                    })
        return self._request_to_public(record)

    def list_approved_base_rows(self) -> list[dict[str, Any]]:
        if not self.service_configured:
            return []
        with self._lock:
            return self._read_named_rows(APPROVED_SHEET, APPROVED_HEADERS, "request_id")

    def _ensure_named_sheet(self, sheet_name: str, headers: list[str]) -> None:
        if not self.service_configured:
            raise ValueError("ยังไม่ได้ตั้งค่า Google Sheet สำหรับเก็บคำขอ BaseData")
        service = self._get_service()
        metadata = service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
        titles = {sheet["properties"]["title"] for sheet in metadata.get("sheets", [])}
        if sheet_name not in titles:
            service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]},
            ).execute()
        current = service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id, range=f"'{sheet_name}'!A1:{self._column_letter(len(headers))}1",
        ).execute().get("values", [])
        if not current:
            service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id, range=f"'{sheet_name}'!A1",
                valueInputOption="RAW", body={"values": [headers]},
            ).execute()
        elif current[0] != headers:
            raise ValueError(f"หัวตารางในแท็บ {sheet_name} ไม่ตรงกับรูปแบบของระบบ")

    def _read_named_rows(self, sheet_name: str, headers: list[str], key: str) -> list[dict[str, Any]]:
        self._ensure_named_sheet(sheet_name, headers)
        values = self._get_service().spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id, range=f"'{sheet_name}'!A2:{self._column_letter(len(headers))}",
        ).execute().get("values", [])
        result = []
        for row_number, value_row in enumerate(values, start=2):
            padded = value_row + [""] * (len(headers) - len(value_row))
            row = dict(zip(headers, padded[:len(headers)]))
            if row.get(key):
                row["_row_number"] = row_number
                result.append(row)
        return result

    def _append_named_record(self, sheet_name: str, headers: list[str], record: dict[str, Any]) -> None:
        self._get_service().spreadsheets().values().append(
            spreadsheetId=self.spreadsheet_id, range=f"'{sheet_name}'!A:{self._column_letter(len(headers))}",
            valueInputOption="RAW", insertDataOption="INSERT_ROWS",
            body={"values": [[record.get(header, "") for header in headers]]},
        ).execute()

    def _update_named_record(self, sheet_name: str, headers: list[str], record: dict[str, Any], row_number: int) -> None:
        self._get_service().spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{sheet_name}'!A{row_number}:{self._column_letter(len(headers))}{row_number}",
            valueInputOption="RAW", body={"values": [[record.get(header, "") for header in headers]]},
        ).execute()

    @staticmethod
    def _column_letter(number: int) -> str:
        result = ""
        while number:
            number, remainder = divmod(number - 1, 26)
            result = chr(65 + remainder) + result
        return result

    @staticmethod
    def _request_to_public(record: dict[str, Any]) -> dict[str, Any]:
        try:
            stored_rows = json.loads(record.get("rows_json", "[]"))
        except json.JSONDecodeError:
            stored_rows = []
        if isinstance(stored_rows, dict):
            rows = stored_rows.get("new", [])
            original_rows = stored_rows.get("original", [])
        else:
            rows = stored_rows
            original_rows = []
        return {
            "id": record.get("request_id", ""), "submittedAt": record.get("submitted_at", ""),
            "submitterName": record.get("submitter_name", ""), "employeeId": record.get("employee_id", ""),
            "department": record.get("department", ""), "action": record.get("action", "add"),
            "size": record.get("size", ""), "head": record.get("head", ""), "rows": rows,
            "originalRows": original_rows,
            "note": record.get("note", ""), "status": record.get("status", "pending"),
            "reviewedAt": record.get("reviewed_at", ""), "reviewedBy": record.get("reviewed_by", ""),
            "reviewNote": record.get("review_note", ""),
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
