from __future__ import annotations

import json
import os
import re
import threading
import uuid
from io import BytesIO
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
REQUEST_JOURNAL_SHEET = "BaseDataRequestJournal"
REQUEST_JOURNAL_HEADERS = ["event_id", "request_id", "event", "record_json", "created_at", "actor"]
APPROVED_SHEET = "ApprovedBaseData"
APPROVED_HEADERS = [
    "size", "head", "material", "code", "quantity", "action", "request_id",
    "approved_at", "approved_by", "department", "insulator_upright", "insulator_horizontal",
    "source_size", "source_head",
    "image_file_id", "image_name", "image_mime_type",
    "last_modified_by",
]
FOLDER_SHEET = "ProjectFolders"
FOLDER_HEADERS = ["folder_id", "folder_name", "owner_email", "created_at", "updated_at", "deleted_at"]


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
        self.drive_folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "").strip()
        self.oauth_client_secret = os.environ.get("GOOGLE_CLIENT_SECRET", "").strip()
        self.drive_refresh_token = os.environ.get("GOOGLE_DRIVE_REFRESH_TOKEN", "").strip()
        self._service = None
        self._drive_service = None
        self._lock = threading.RLock()
        self._pending_request_cache: dict[str, dict[str, Any]] = {}
        self._pending_record_cache: dict[str, dict[str, Any]] = {}
        self._base_request_snapshot: list[dict[str, Any]] = []
        self._checked_named_sheets: set[tuple[str, tuple[str, ...]]] = set()

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

    @property
    def drive_configured(self) -> bool:
        if self.drive_oauth_ready:
            return bool(self.drive_folder_id and self.drive_oauth_configured)
        return bool(self.drive_folder_id and self._credentials_json)

    @property
    def drive_oauth_ready(self) -> bool:
        return bool(self.client_id and self.oauth_client_secret)

    @property
    def drive_oauth_configured(self) -> bool:
        return bool(self.drive_oauth_ready and self.drive_refresh_token)

    def submit_base_request(self, payload: dict[str, Any]) -> dict[str, Any]:
        required = {
            "submitter_name": "ชื่อผู้เสนอ", "employee_id": "รหัสพนักงาน",
            "department": "สังกัด", "size": "ขนาดเสา", "head": "รหัสหัวเสา",
        }
        values = {key: str(payload.get(key, "")).strip() for key in required}
        missing = [label for key, label in required.items() if not values[key]]
        if missing:
            raise ValueError(f"กรุณากรอก {', '.join(missing)}")
        action = str(payload.get("action", "add")).strip().lower()
        if action not in {"add", "replace", "rename", "copy", "delete"}:
            action = "add"
        rows = payload.get("rows", [])
        if not isinstance(rows, list) or (not rows and action != "delete"):
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
        original_rows = payload.get("original_rows", []) if action != "add" else []
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
        target_department = str(payload.get("target_department", "แผนกแรงสูง")).strip() or "แผนกแรงสูง"
        insulator_upright = self._nonnegative_number(payload.get("insulator_upright"), "ลูกถ้วยตั้ง")
        insulator_horizontal = self._nonnegative_number(payload.get("insulator_horizontal"), "ลูกถ้วยนอน")
        image_file_id = str(payload.get("image_file_id", "")).strip()
        image_name = str(payload.get("image_name", "")).strip()
        image_mime_type = str(payload.get("image_mime_type", "")).strip()
        source_size = str(payload.get("source_size", "")).strip() or (values["size"] if action in {"replace", "delete"} else "")
        source_head = str(payload.get("source_head", "")).strip() or (values["head"] if action in {"replace", "delete"} else "")
        if action in {"rename", "copy", "delete"} and (not source_size or not source_head):
            raise ValueError("กรุณาเลือกขนาดเสาและหัวเสาต้นฉบับ")
        record = {
            "request_id": uuid.uuid4().hex,
            "submitted_at": datetime.now(timezone.utc).isoformat(),
            **values,
            "action": action,
            "rows_json": json.dumps(
                {"new": clean_rows, "original": clean_original_rows, "department": target_department,
                 "insulator_upright": insulator_upright, "insulator_horizontal": insulator_horizontal,
                 "source_size": source_size, "source_head": source_head,
                 "image_file_id": image_file_id, "image_name": image_name, "image_mime_type": image_mime_type},
                ensure_ascii=False, separators=(",", ":"),
            ),
            "note": str(payload.get("note", "")).strip(),
            "status": "pending", "reviewed_at": "", "reviewed_by": "", "review_note": "",
        }
        with self._lock:
            self._ensure_named_sheet(REQUEST_SHEET, REQUEST_HEADERS)
            row_number = self._append_named_record(REQUEST_SHEET, REQUEST_HEADERS, record)
            if row_number:
                record["_row_number"] = row_number
            try:
                self._append_request_journal(record, "submitted", values["submitter_name"])
            except Exception:
                # The primary request row is already durable; do not make the user submit twice.
                pass
            public_record = self._request_to_public(record)
            self._pending_request_cache[public_record["id"]] = public_record
            self._pending_record_cache[public_record["id"]] = dict(record)
            self._base_request_snapshot = [public_record, *[item for item in self._base_request_snapshot if item["id"] != public_record["id"]]]
        return public_record

    def list_base_requests(self) -> list[dict[str, Any]]:
        with self._lock:
            try:
                rows = self._read_named_rows(REQUEST_SHEET, REQUEST_HEADERS, "request_id")
            except Exception:
                journal_records = self._read_request_journal_records()
                if journal_records:
                    rows = []
                elif self._base_request_snapshot or self._pending_request_cache:
                    cached = list(self._base_request_snapshot)
                    pending = list(self._pending_request_cache.values())
                    pending_ids = {request["id"] for request in pending}
                    return [*pending, *[request for request in cached if request["id"] not in pending_ids]]
                else:
                    raise
            else:
                journal_records = self._read_request_journal_records()
            requests = [self._request_to_public(row) for row in reversed(rows)]
            stored_by_id = {request["id"]: request for request in requests}
            for journal_record in journal_records:
                request_id = journal_record.get("request_id")
                if request_id and request_id not in stored_by_id:
                    public_record = self._request_to_public(journal_record)
                    requests.insert(0, public_record)
                    stored_by_id[request_id] = public_record
            for request_id, request in list(self._pending_request_cache.items()):
                stored = stored_by_id.get(request_id)
                if stored and stored.get("status") != "pending":
                    self._pending_request_cache.pop(request_id, None)
                elif not stored:
                    requests.insert(0, request)
            self._base_request_snapshot = list(requests)
            return requests

    def list_pending_base_requests(self) -> list[dict[str, Any]]:
        """Return pending request details that are safe to show without Admin access."""
        return self.public_base_request_overview()["requests"]

    def public_base_request_overview(self) -> dict[str, Any]:
        """Return safe pending details plus IDs that have completed review."""
        public_fields = {
            "id", "submittedAt", "action", "targetDepartment", "size", "head", "rows",
            "originalRows", "sourceSize", "sourceHead", "insulatorUpright", "insulatorHorizontal",
        }
        all_requests = self.list_base_requests()
        pending = [
            {key: value for key, value in request.items() if key in public_fields}
            for request in all_requests
            if request.get("status") == "pending"
        ]
        return {
            "requests": pending,
            "resolvedRequestIds": [request["id"] for request in all_requests if request.get("status") != "pending"],
        }

    def clear_approved_requests(self) -> int:
        """Clear approved request history without touching published BaseData rows."""
        with self._lock:
            rows = self._read_named_rows(REQUEST_SHEET, REQUEST_HEADERS, "request_id")
            approved_rows = [row for row in rows if row.get("status") == "approved"]
            row_numbers = [int(row["_row_number"]) for row in approved_rows]
            for row in approved_rows:
                archived = dict(row)
                archived["status"] = "archived"
                self._append_request_journal(archived, "archived", "admin")
            if row_numbers:
                self._clear_named_rows(REQUEST_SHEET, len(REQUEST_HEADERS), row_numbers)
        return len(row_numbers)

    def review_base_request(
        self, request_id: str, approve: bool, reviewer: str, note: str = "",
        edits: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        with self._lock:
            rows = self._read_named_rows(REQUEST_SHEET, REQUEST_HEADERS, "request_id")
            record = next((row for row in rows if row["request_id"] == request_id), None)
            if not record:
                cached_record = self._pending_record_cache.get(request_id)
                record = dict(cached_record) if cached_record and cached_record.get("_row_number") else None
            if not record:
                record = next((item for item in self._read_request_journal_records() if item.get("request_id") == request_id and item.get("status") == "pending"), None)
            if not record:
                raise ValueError("ไม่พบคำขอที่ต้องการตรวจ")
            if record.get("status") != "pending":
                raise ValueError("คำขอนี้ถูกตรวจแล้ว")
            if edits is not None:
                if not approve:
                    raise ValueError("แก้ไขข้อมูลได้เฉพาะตอนอนุมัติ")
                if not isinstance(edits, dict):
                    raise ValueError("ข้อมูลที่แก้ไขไม่ถูกต้อง")
                stored_rows = json.loads(record.get("rows_json", "{}"))
                if not isinstance(stored_rows, dict):
                    stored_rows = {"new": stored_rows if isinstance(stored_rows, list) else []}
                edited_size = str(edits.get("size", "")).strip()
                edited_head = str(edits.get("head", "")).strip()
                if not edited_size or not edited_head:
                    raise ValueError("กรุณากรอกขนาดเสาและรหัสหัวเสา")
                edited_rows = edits.get("rows", [])
                if not isinstance(edited_rows, list) or not edited_rows:
                    raise ValueError("กรุณาเพิ่มรายการวัสดุอย่างน้อย 1 รายการ")
                if len(edited_rows) > 100:
                    raise ValueError("หนึ่งคำขอเพิ่มรายการวัสดุได้ไม่เกิน 100 รายการ")
                clean_rows = []
                for item in edited_rows:
                    material = str(item.get("material", "")).strip()
                    code = str(item.get("code", "")).strip()
                    try:
                        quantity = float(item.get("quantity", 0))
                    except (TypeError, ValueError) as exc:
                        raise ValueError("จำนวนวัสดุต้องเป็นตัวเลข") from exc
                    if not material or not code or quantity == 0:
                        raise ValueError("ทุกรายการต้องมีชื่อวัสดุ รหัสพัสดุ/SET และจำนวนที่ไม่เป็นศูนย์")
                    clean_rows.append({"material": material, "code": code, "quantity": quantity})
                stored_rows["submitted"] = {
                    "size": record.get("size", ""), "head": record.get("head", ""),
                    "new": stored_rows.get("new", []),
                    "insulator_upright": stored_rows.get("insulator_upright", ""),
                    "insulator_horizontal": stored_rows.get("insulator_horizontal", ""),
                }
                stored_rows["new"] = clean_rows
                stored_rows["insulator_upright"] = self._nonnegative_number(edits.get("insulator_upright"), "ลูกถ้วยตั้ง")
                stored_rows["insulator_horizontal"] = self._nonnegative_number(edits.get("insulator_horizontal"), "ลูกถ้วยนอน")
                stored_rows["admin_edited"] = True
                record["size"] = edited_size
                record["head"] = edited_head
                record["rows_json"] = json.dumps(stored_rows, ensure_ascii=False, separators=(",", ":"))
            now = datetime.now(timezone.utc).isoformat()
            record.update({
                "status": "approved" if approve else "rejected", "reviewed_at": now,
                "reviewed_by": reviewer, "review_note": str(note).strip(),
            })
            if record.get("_row_number"):
                self._update_named_record(REQUEST_SHEET, REQUEST_HEADERS, record, record["_row_number"])
            if approve:
                self._ensure_named_sheet(APPROVED_SHEET, APPROVED_HEADERS)
                stored_rows = json.loads(record["rows_json"])
                new_rows = stored_rows.get("new", []) if isinstance(stored_rows, dict) else stored_rows
                if record["action"] == "delete":
                    new_rows = [{"material": "", "code": "", "quantity": 0}]
                for item in new_rows:
                    self._append_named_record(APPROVED_SHEET, APPROVED_HEADERS, {
                        "size": record["size"], "head": record["head"],
                        "material": item["material"], "code": item["code"], "quantity": item["quantity"],
                        "action": record["action"], "request_id": request_id,
                        "approved_at": now, "approved_by": reviewer,
                        "department": stored_rows.get("department", "แผนกแรงสูง") if isinstance(stored_rows, dict) else "แผนกแรงสูง",
                        "insulator_upright": stored_rows.get("insulator_upright", "") if isinstance(stored_rows, dict) else "",
                        "insulator_horizontal": stored_rows.get("insulator_horizontal", "") if isinstance(stored_rows, dict) else "",
                        "source_size": stored_rows.get("source_size", "") if isinstance(stored_rows, dict) else "",
                        "source_head": stored_rows.get("source_head", "") if isinstance(stored_rows, dict) else "",
                        "image_file_id": stored_rows.get("image_file_id", "") if isinstance(stored_rows, dict) else "",
                        "image_name": stored_rows.get("image_name", "") if isinstance(stored_rows, dict) else "",
                        "image_mime_type": stored_rows.get("image_mime_type", "") if isinstance(stored_rows, dict) else "",
                        "last_modified_by": record.get("submitter_name", "") or "Admin",
                    })
            self._pending_request_cache.pop(request_id, None)
            self._pending_record_cache.pop(request_id, None)
            public_record = self._request_to_public(record)
            self._base_request_snapshot = [public_record, *[item for item in self._base_request_snapshot if item["id"] != request_id]]
            self._append_request_journal(record, "approved" if approve else "rejected", reviewer)
        return public_record

    def _append_request_journal(self, record: dict[str, Any], event: str, actor: str) -> None:
        self._ensure_named_sheet(REQUEST_JOURNAL_SHEET, REQUEST_JOURNAL_HEADERS)
        clean_record = {key: value for key, value in record.items() if not key.startswith("_")}
        self._append_named_record(REQUEST_JOURNAL_SHEET, REQUEST_JOURNAL_HEADERS, {
            "event_id": uuid.uuid4().hex,
            "request_id": record.get("request_id", ""),
            "event": event,
            "record_json": json.dumps(clean_record, ensure_ascii=False, separators=(",", ":")),
            "created_at": datetime.now(timezone.utc).isoformat(),
            "actor": str(actor).strip(),
        })

    def _read_request_journal_records(self) -> list[dict[str, Any]]:
        try:
            events = self._read_named_rows(REQUEST_JOURNAL_SHEET, REQUEST_JOURNAL_HEADERS, "event_id")
        except Exception:
            return []
        latest: dict[str, dict[str, Any]] = {}
        for event in events:
            try:
                record = json.loads(event.get("record_json", "{}"))
            except (TypeError, json.JSONDecodeError):
                continue
            request_id = str(event.get("request_id", "")).strip()
            if request_id and isinstance(record, dict):
                latest[request_id] = record
        return [record for record in latest.values() if record.get("status") != "archived"]

    def list_approved_base_rows(self) -> list[dict[str, Any]]:
        if not self.service_configured:
            return []
        with self._lock:
            return self._read_named_rows(APPROVED_SHEET, APPROVED_HEADERS, "request_id")

    def upload_head_image(self, content: bytes, name: str, mime_type: str) -> dict[str, str]:
        if not self.drive_configured:
            if self.drive_oauth_ready and not self.drive_refresh_token:
                raise ValueError("ยังไม่ได้เชื่อมต่อ Google Drive OAuth กรุณาเข้าเมนูรายการอนุมัติ (Admin) แล้วกดเชื่อมต่อ Google Drive")
            raise ValueError("ยังไม่ได้ตั้งค่า GOOGLE_DRIVE_FOLDER_ID บน Render")
        if len(content) > 5 * 1024 * 1024:
            raise ValueError("รูปต้องมีขนาดไม่เกิน 5 MB")
        allowed = {"image/jpeg": (b"\xff\xd8\xff",), "image/png": (b"\x89PNG\r\n\x1a\n",), "image/webp": (b"RIFF",)}
        if mime_type not in allowed or not any(content.startswith(signature) for signature in allowed[mime_type]):
            raise ValueError("รองรับเฉพาะรูป JPG, PNG หรือ WEBP")
        if mime_type == "image/webp" and content[8:12] != b"WEBP":
            raise ValueError("ไฟล์ WEBP ไม่ถูกต้อง")
        from googleapiclient.http import MediaIoBaseUpload
        safe_name = "".join(character for character in str(name) if character.isalnum() or character in "._- ").strip() or "head-image"
        uploaded = self._get_drive_service().files().create(
            body={"name": f"{uuid.uuid4().hex[:10]}-{safe_name}", "parents": [self.drive_folder_id]},
            media_body=MediaIoBaseUpload(BytesIO(content), mimetype=mime_type, resumable=False),
            fields="id,name,mimeType",
            supportsAllDrives=True,
        ).execute()
        return {"id": uploaded["id"], "name": uploaded.get("name", safe_name), "mimeType": uploaded.get("mimeType", mime_type)}

    def download_head_image(self, file_id: str) -> tuple[bytes, str, str]:
        if not self.drive_configured or not file_id:
            if self.drive_oauth_ready and not self.drive_refresh_token:
                raise ValueError("ยังไม่ได้เชื่อมต่อ Google Drive OAuth")
            raise ValueError("ไม่พบรูปประกอบ")
        service = self._get_drive_service()
        metadata = service.files().get(fileId=file_id, fields="id,name,mimeType,parents", supportsAllDrives=True).execute()
        if self.drive_folder_id not in metadata.get("parents", []):
            raise PermissionError("รูปนี้ไม่ได้อยู่ในโฟลเดอร์รูปหัวเสา")
        content = service.files().get_media(fileId=file_id, supportsAllDrives=True).execute()
        return content, metadata.get("mimeType", "application/octet-stream"), metadata.get("name", "head-image")

    @staticmethod
    def _nonnegative_number(value: Any, label: str) -> float:
        try:
            number = float(value if value not in (None, "") else 0)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{label}ต้องเป็นตัวเลข") from exc
        if number < 0:
            raise ValueError(f"{label}ต้องไม่ติดลบ")
        return number

    def _ensure_named_sheet(self, sheet_name: str, headers: list[str]) -> None:
        cache_key = (sheet_name, tuple(headers))
        if cache_key in self._checked_named_sheets:
            return
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
            if headers[:len(current[0])] == current[0]:
                service.spreadsheets().values().update(
                    spreadsheetId=self.spreadsheet_id, range=f"'{sheet_name}'!A1",
                    valueInputOption="RAW", body={"values": [headers]},
                ).execute()
            else:
                raise ValueError(f"หัวตารางในแท็บ {sheet_name} ไม่ตรงกับรูปแบบของระบบ")
        self._checked_named_sheets.add(cache_key)

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

    def _append_named_record(self, sheet_name: str, headers: list[str], record: dict[str, Any]) -> int | None:
        response = self._get_service().spreadsheets().values().append(
            spreadsheetId=self.spreadsheet_id, range=f"'{sheet_name}'!A:{self._column_letter(len(headers))}",
            valueInputOption="RAW", insertDataOption="INSERT_ROWS",
            body={"values": [[record.get(header, "") for header in headers]]},
        ).execute()
        updated_range = str(response.get("updates", {}).get("updatedRange", ""))
        match = re.search(r"![A-Z]+(\d+)(?::|$)", updated_range)
        return int(match.group(1)) if match else None

    def _update_named_record(self, sheet_name: str, headers: list[str], record: dict[str, Any], row_number: int) -> None:
        self._get_service().spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f"'{sheet_name}'!A{row_number}:{self._column_letter(len(headers))}{row_number}",
            valueInputOption="RAW", body={"values": [[record.get(header, "") for header in headers]]},
        ).execute()

    def _clear_named_rows(self, sheet_name: str, column_count: int, row_numbers: list[int]) -> None:
        last_column = self._column_letter(column_count)
        self._get_service().spreadsheets().values().batchClear(
            spreadsheetId=self.spreadsheet_id,
            body={"ranges": [f"'{sheet_name}'!A{row}:{last_column}{row}" for row in row_numbers]},
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
            department = stored_rows.get("department", "แผนกแรงสูง")
            insulator_upright = stored_rows.get("insulator_upright", "")
            insulator_horizontal = stored_rows.get("insulator_horizontal", "")
            source_size = stored_rows.get("source_size", "")
            source_head = stored_rows.get("source_head", "")
            image_file_id = stored_rows.get("image_file_id", "")
            image_name = stored_rows.get("image_name", "")
            image_mime_type = stored_rows.get("image_mime_type", "")
            submitted = stored_rows.get("submitted")
            admin_edited = bool(stored_rows.get("admin_edited"))
        else:
            rows = stored_rows
            original_rows = []
            department = "แผนกแรงสูง"
            insulator_upright = ""
            insulator_horizontal = ""
            source_size = ""
            source_head = ""
            image_file_id = ""
            image_name = ""
            image_mime_type = ""
            submitted = None
            admin_edited = False
        return {
            "id": record.get("request_id", ""), "submittedAt": record.get("submitted_at", ""),
            "submitterName": record.get("submitter_name", ""), "employeeId": record.get("employee_id", ""),
            "department": record.get("department", ""), "action": record.get("action", "add"),
            "targetDepartment": department,
            "insulatorUpright": insulator_upright, "insulatorHorizontal": insulator_horizontal,
            "size": record.get("size", ""), "head": record.get("head", ""), "rows": rows,
            "originalRows": original_rows,
            "sourceSize": source_size, "sourceHead": source_head,
            "imageFileId": image_file_id, "imageName": image_name, "imageMimeType": image_mime_type,
            "note": record.get("note", ""), "status": record.get("status", "pending"),
            "reviewedAt": record.get("reviewed_at", ""), "reviewedBy": record.get("reviewed_by", ""),
            "reviewNote": record.get("review_note", ""),
            "adminEdited": admin_edited, "submitted": submitted,
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

    @staticmethod
    def _project_owner(row: dict[str, Any]) -> str:
        return str(row.get("created_by") or row.get("updated_by") or "").strip().lower()

    def list_projects(self, user: dict[str, str]) -> list[dict[str, Any]]:
        owner_email = str(user.get("email", "")).strip().lower()
        with self._lock:
            rows = self._read_rows()
        projects = [
            self._row_to_project(row) for row in rows
            if not row.get("deleted_at") and self._project_owner(row) == owner_email
        ]
        return sorted(projects, key=lambda project: project.get("updatedAt", ""), reverse=True)

    def list_deleted_projects(self, user: dict[str, str]) -> list[dict[str, Any]]:
        owner_email = str(user.get("email", "")).strip().lower()
        with self._lock:
            rows = self._read_rows()
        projects = [self._row_to_project(row) for row in rows
                    if row.get("deleted_at") and self._project_owner(row) == owner_email]
        return sorted(projects, key=lambda project: project.get("deletedAt", ""), reverse=True)

    def list_all_deleted_projects(self) -> list[dict[str, Any]]:
        with self._lock:
            rows = self._read_rows()
        return sorted([self._row_to_project(row) for row in rows if row.get("deleted_at")],
                      key=lambda project: project.get("deletedAt", ""), reverse=True)

    def restore_project_as_admin(self, project_id: str, owner_email: str) -> dict[str, Any]:
        owner_email = str(owner_email).strip().lower()
        with self._lock:
            rows = self._read_rows()
            existing = next((row for row in rows if row.get("project_id") == project_id
                             and row.get("deleted_at") and self._project_owner(row) == owner_email), None)
            if not existing:
                raise ValueError("ไม่พบงานในถังขยะ")
            existing["deleted_at"] = ""
            existing["updated_at"] = datetime.now(timezone.utc).isoformat()
            self._write_record(existing, existing["_row_number"])
        return self._row_to_project(existing)

    def permanently_delete_project_as_admin(self, project_id: str, owner_email: str) -> None:
        owner_email = str(owner_email).strip().lower()
        with self._lock:
            rows = self._read_rows()
            existing = next((row for row in rows if row.get("project_id") == project_id
                             and row.get("deleted_at") and self._project_owner(row) == owner_email), None)
            if not existing:
                raise ValueError("ไม่พบงานในถังขยะ")
            self._write_record({}, existing["_row_number"])

    def list_project_folders(self, user: dict[str, str]) -> list[dict[str, Any]]:
        owner_email = str(user.get("email", "")).strip().lower()
        with self._lock:
            rows = self._read_named_rows(FOLDER_SHEET, FOLDER_HEADERS, "folder_id")
        return sorted([
            {"id": row.get("folder_id", ""), "name": row.get("folder_name", ""),
             "createdAt": row.get("created_at", ""), "updatedAt": row.get("updated_at", "")}
            for row in rows
            if not row.get("deleted_at") and str(row.get("owner_email", "")).strip().lower() == owner_email
        ], key=lambda folder: folder["name"].casefold())

    def create_project_folder(self, name: str, user: dict[str, str]) -> dict[str, Any]:
        name = " ".join(str(name).strip().split())
        if not name:
            raise ValueError("กรุณาระบุชื่อโฟลเดอร์")
        if len(name) > 80:
            raise ValueError("ชื่อโฟลเดอร์ยาวเกิน 80 ตัวอักษร")
        owner_email = str(user.get("email", "")).strip().lower()
        now = datetime.now(timezone.utc).isoformat()
        with self._lock:
            rows = self._read_named_rows(FOLDER_SHEET, FOLDER_HEADERS, "folder_id")
            if any(not row.get("deleted_at") and str(row.get("owner_email", "")).strip().lower() == owner_email
                   and str(row.get("folder_name", "")).strip().casefold() == name.casefold() for row in rows):
                raise ValueError("มีโฟลเดอร์ชื่อนี้แล้ว")
            record = {"folder_id": uuid.uuid4().hex, "folder_name": name, "owner_email": owner_email,
                      "created_at": now, "updated_at": now, "deleted_at": ""}
            self._append_named_record(FOLDER_SHEET, FOLDER_HEADERS, record)
        return {"id": record["folder_id"], "name": name, "createdAt": now, "updatedAt": now}

    def move_project(self, project_id: str, folder_id: str, user: dict[str, str]) -> dict[str, Any]:
        owner_email = str(user.get("email", "")).strip().lower()
        folder_id = str(folder_id or "").strip()
        with self._lock:
            if folder_id:
                folders = self._read_named_rows(FOLDER_SHEET, FOLDER_HEADERS, "folder_id")
                valid = any(row.get("folder_id") == folder_id and not row.get("deleted_at")
                            and str(row.get("owner_email", "")).strip().lower() == owner_email for row in folders)
                if not valid:
                    raise ValueError("ไม่พบโฟลเดอร์ปลายทาง")
            rows = self._read_rows()
            existing = next((row for row in rows if row.get("project_id") == project_id
                             and not row.get("deleted_at") and self._project_owner(row) == owner_email), None)
            if not existing:
                raise ValueError("ไม่พบงานที่ต้องการย้าย")
            pages_data = json.loads(existing.get("pages_json") or "{}")
            if not isinstance(pages_data, dict):
                pages_data = {"version": 2, "activeDepartment": "แผนกแรงสูง", "departments": {}, "pages": pages_data}
            pages_data["folderId"] = folder_id
            existing["pages_json"] = json.dumps(pages_data, ensure_ascii=False, separators=(",", ":"))
            existing["updated_at"] = datetime.now(timezone.utc).isoformat()
            existing["updated_by"] = owner_email
            self._write_record(existing, existing["_row_number"])
        return self._row_to_project(existing)

    def delete_project_folder(self, folder_id: str, user: dict[str, str]) -> None:
        owner_email = str(user.get("email", "")).strip().lower()
        with self._lock:
            if any(project.get("folderId") == folder_id for project in self.list_projects(user)):
                raise ValueError("กรุณาย้ายงานออกจากโฟลเดอร์ก่อนลบ")
            rows = self._read_named_rows(FOLDER_SHEET, FOLDER_HEADERS, "folder_id")
            existing = next((row for row in rows if row.get("folder_id") == folder_id
                             and not row.get("deleted_at")
                             and str(row.get("owner_email", "")).strip().lower() == owner_email), None)
            if not existing:
                raise ValueError("ไม่พบโฟลเดอร์ที่ต้องการลบ")
            existing["deleted_at"] = datetime.now(timezone.utc).isoformat()
            existing["updated_at"] = existing["deleted_at"]
            self._update_named_record(FOLDER_SHEET, FOLDER_HEADERS, existing, existing["_row_number"])

    def save_project(self, project: dict[str, Any], user: dict[str, str]) -> dict[str, Any]:
        project_id = str(project.get("id", "")).strip()
        name = str(project.get("name", "")).strip()
        if not project_id or not name:
            raise ValueError("ข้อมูล Project ID หรือชื่องานไม่ครบ")

        now = datetime.now(timezone.utc).isoformat()
        owner_email = str(user.get("email", "")).strip().lower()
        with self._lock:
            rows = self._read_rows()
            existing = next((
                row for row in rows
                if row.get("project_id") == project_id and self._project_owner(row) == owner_email
            ), None)
            record = {
                "project_id": project_id,
                "project_name": name,
                "plan_number": str(project.get("planNumber", "")).strip(),
                "pages_json": json.dumps({
                    "version": 2,
                    "folderId": project.get("folderId", (self._row_to_project(existing).get("folderId", "") if existing else "")),
                    "activeDepartment": project.get("department", "แผนกแรงสูง"),
                    "departments": project.get("departments", {}),
                    "pages": project.get("pages", []),
                    "structurePages": project.get("structurePages", []),
                }, ensure_ascii=False, separators=(",", ":")),
                "results_json": json.dumps(project.get("results", []), ensure_ascii=False, separators=(",", ":")),
                "result_meta": str(project.get("resultMeta", "")),
                "created_at": (existing or {}).get("created_at") or str(project.get("createdAt", "")) or now,
                "updated_at": now,
                "created_by": (existing or {}).get("created_by") or owner_email,
                "updated_by": owner_email,
                "deleted_at": "",
            }
            self._write_record(record, existing.get("_row_number") if existing else None)
        return self._row_to_project(record)

    def delete_project(self, project_id: str, user: dict[str, str]) -> None:
        owner_email = str(user.get("email", "")).strip().lower()
        with self._lock:
            rows = self._read_rows()
            existing = next((
                row for row in rows
                if row.get("project_id") == project_id
                and not row.get("deleted_at")
                and self._project_owner(row) == owner_email
            ), None)
            if not existing:
                raise ValueError("ไม่พบงานที่ต้องการลบ")
            existing["deleted_at"] = datetime.now(timezone.utc).isoformat()
            existing["updated_by"] = owner_email
            self._write_record(existing, existing["_row_number"])

    def restore_project(self, project_id: str, user: dict[str, str]) -> dict[str, Any]:
        owner_email = str(user.get("email", "")).strip().lower()
        with self._lock:
            rows = self._read_rows()
            existing = next((row for row in rows if row.get("project_id") == project_id
                             and row.get("deleted_at") and self._project_owner(row) == owner_email), None)
            if not existing:
                raise ValueError("ไม่พบงานในถังขยะ")
            existing["deleted_at"] = ""
            existing["updated_at"] = datetime.now(timezone.utc).isoformat()
            existing["updated_by"] = owner_email
            self._write_record(existing, existing["_row_number"])
        return self._row_to_project(existing)

    def permanently_delete_project(self, project_id: str, user: dict[str, str]) -> None:
        owner_email = str(user.get("email", "")).strip().lower()
        with self._lock:
            rows = self._read_rows()
            existing = next((row for row in rows if row.get("project_id") == project_id
                             and row.get("deleted_at") and self._project_owner(row) == owner_email), None)
            if not existing:
                raise ValueError("ไม่พบงานในถังขยะ")
            self._write_record({}, existing["_row_number"])

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

    def _get_drive_service(self):
        if self._drive_service is None:
            from googleapiclient.discovery import build
            if self.drive_oauth_configured:
                from google.oauth2.credentials import Credentials
                credentials = Credentials(
                    token=None,
                    refresh_token=self.drive_refresh_token,
                    token_uri="https://oauth2.googleapis.com/token",
                    client_id=self.client_id,
                    client_secret=self.oauth_client_secret,
                    scopes=["https://www.googleapis.com/auth/drive"],
                )
            else:
                from google.oauth2 import service_account
                info = json.loads(self._credentials_json)
                credentials = service_account.Credentials.from_service_account_info(
                    info, scopes=["https://www.googleapis.com/auth/drive"],
                )
            self._drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
        return self._drive_service

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

        pages_data = parse_json(row.get("pages_json", ""), [])
        if isinstance(pages_data, dict):
            pages = pages_data.get("pages", [])
            departments = pages_data.get("departments", {})
            department = pages_data.get("activeDepartment", "แผนกแรงสูง")
            folder_id = pages_data.get("folderId", "")
            structure_pages = pages_data.get("structurePages", [])
        else:
            pages, departments, department, folder_id, structure_pages = pages_data, {}, "", "", []
        return {
            "id": row.get("project_id", ""),
            "name": row.get("project_name", ""),
            "planNumber": row.get("plan_number", ""),
            "pages": pages,
            "structurePages": structure_pages,
            "departments": departments,
            "department": department,
            "results": parse_json(row.get("results_json", ""), []),
            "resultMeta": row.get("result_meta", ""),
            "createdAt": row.get("created_at", ""),
            "updatedAt": row.get("updated_at", ""),
            "createdBy": row.get("created_by", ""),
            "updatedBy": row.get("updated_by", ""),
            "folderId": folder_id,
            "deletedAt": row.get("deleted_at", ""),
        }
