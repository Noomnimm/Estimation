import json
import unittest
from unittest.mock import patch

from cloud_store import GoogleSheetProjectStore


class CloudProjectStorageTests(unittest.TestCase):
    def _configured_store(self):
        store = GoogleSheetProjectStore()
        store.projects_drive_folder_id = "projects-folder"
        store.client_id = "client-id"
        store.oauth_client_secret = "client-secret"
        store.drive_refresh_token = "refresh-token"
        return store

    def test_large_project_uses_drive_pointer_instead_of_oversized_sheet_cell(self):
        store = self._configured_store()
        user = {"email": "worker@example.com"}
        project = {
            "id": "project-1",
            "name": "งานทดสอบ",
            "pages": [[{"material": "x" * 60000, "code": "1001"}]],
            "departments": {"แรงสูง": {"pages": [[{"material": "x" * 60000}]]}},
            "results": [{"material": "x" * 3000}],
        }
        saved_rows = []
        uploaded_payloads = []

        def upload(payload_json, project_id, owner_email):
            uploaded_payloads.append(json.loads(payload_json))
            return "drive-file-1"

        def write(record, row_number):
            saved_rows.append(dict(record, _row_number=row_number or 2))

        with patch.object(store, "_read_rows", return_value=[]), \
             patch.object(store, "_write_record", side_effect=write), \
             patch.object(store, "_upload_project_payload", side_effect=upload):
            store._delete_project_payload_file = lambda *args: None
            summary = store.save_project(project, user)

        self.assertEqual(summary["id"], "project-1")
        self.assertEqual(len(uploaded_payloads), 1)
        pointer = json.loads(saved_rows[0]["pages_json"])
        self.assertEqual(pointer["storage"], "drive-json-v1")
        self.assertEqual(pointer["fileId"], "drive-file-1")
        self.assertLess(len(saved_rows[0]["pages_json"]), 1000)
        self.assertEqual(saved_rows[0]["results_json"], "")

    def test_legacy_project_rows_still_load_without_drive_storage(self):
        store = GoogleSheetProjectStore()
        legacy = {
            "project_id": "old-project",
            "project_name": "งานเก่า",
            "plan_number": "P-1",
            "pages_json": json.dumps({
                "version": 2,
                "folderId": "folder-1",
                "activeDepartment": "แผนกแรงสูง",
                "departments": {"แผนกแรงสูง": {"pages": [[{"code": "1001"}]]}},
                "pages": [[{"code": "1001"}]],
                "structurePages": [],
            }),
            "results_json": "[]",
            "result_meta": "ผลเดิม",
            "created_by": "worker@example.com",
        }

        project = store._row_to_project(legacy)

        self.assertEqual(project["pages"][0][0]["code"], "1001")
        self.assertEqual(project["folderId"], "folder-1")
        self.assertEqual(project["resultMeta"], "ผลเดิม")

    def test_drive_backed_project_loads_on_demand(self):
        store = self._configured_store()
        row = {
            "project_id": "project-2",
            "project_name": "งาน Drive",
            "pages_json": json.dumps({
                "version": 3, "storage": "drive-json-v1", "fileId": "file-2", "folderId": "folder-2",
            }),
            "created_by": "worker@example.com",
        }
        payload = {
            "pages": [[{"code": "2002"}]],
            "structurePages": [],
            "departments": {},
            "activeDepartment": "แผนกแรงสูง",
            "results": [],
            "resultMeta": "",
        }

        with patch.object(store, "_read_rows", return_value=[row]), \
             patch.object(store, "_download_project_payload", return_value=payload) as download:
            summaries = store.list_projects({"email": "worker@example.com"})
            project = store.get_project("project-2", {"email": "worker@example.com"})

        self.assertNotIn("pages", summaries[0])
        self.assertEqual(project["pages"][0][0]["code"], "2002")
        download.assert_called_once_with("file-2", "project-2", "worker@example.com")

    def test_large_project_without_drive_config_fails_before_writing_sheet(self):
        store = GoogleSheetProjectStore()
        project = {
            "id": "project-1",
            "name": "งานทดสอบ",
            "pages": [[{"material": "x" * 50000}]],
        }
        write = unittest.mock.Mock()

        with patch.object(store, "_read_rows", return_value=[]), \
             patch.object(store, "_write_record", write):
            with self.assertRaisesRegex(ValueError, "ตั้งค่าโฟลเดอร์ Drive"):
                store.save_project(project, {"email": "worker@example.com"})

        write.assert_not_called()


if __name__ == "__main__":
    unittest.main()
