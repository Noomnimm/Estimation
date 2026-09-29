import json
import unittest

from web_app.cloud_store import GoogleSheetProjectStore


class MemoryBaseRequestStore(GoogleSheetProjectStore):
    def __init__(self):
        super().__init__()
        self.rows = []
        self.approved = []

    def _ensure_named_sheet(self, sheet_name, headers):
        return None

    def _append_named_record(self, sheet_name, headers, record):
        saved = dict(record)
        if sheet_name == "BaseDataRequests":
            saved["_row_number"] = len(self.rows) + 2
            self.rows.append(saved)
        else:
            self.approved.append(saved)

    def _read_named_rows(self, sheet_name, headers, key):
        return [dict(row) for row in self.rows]

    def _update_named_record(self, sheet_name, headers, record, row_number):
        self.rows[row_number - 2] = dict(record)

    def _clear_named_rows(self, sheet_name, column_count, row_numbers):
        targets = set(row_numbers)
        self.rows = [row for index, row in enumerate(self.rows, start=2) if index not in targets]
        for index, row in enumerate(self.rows, start=2):
            row["_row_number"] = index


class BaseRequestTests(unittest.TestCase):
    def test_cloud_project_reads_multi_department_payload(self):
        payload = {
            "version": 2, "activeDepartment": "แผนกแรงสูง TAC",
            "departments": {"แผนกแรงสูง": {"pages": [[{"head": "BA"}]]}, "แผนกแรงสูง TAC": {"pages": [[{"head": "TAC"}]]}},
            "pages": [[{"head": "TAC"}]],
        }
        project = GoogleSheetProjectStore._row_to_project({
            "project_id": "p1", "project_name": "งานเดียวหลายแผนก", "pages_json": json.dumps(payload),
        })
        self.assertEqual(project["department"], "แผนกแรงสูง TAC")
        self.assertEqual(len(project["departments"]), 2)
        self.assertEqual(project["pages"][0][0]["head"], "TAC")

    def test_request_stays_pending_until_admin_approval(self):
        store = MemoryBaseRequestStore()
        request = store.submit_base_request({
            "submitter_name": "ผู้ทดสอบ", "employee_id": "123456", "department": "กวว.",
            "target_department": "แผนกแรงสูง TAC",
            "insulator_upright": 5,
            "insulator_horizontal": 10,
            "action": "replace", "size": "14.3", "head": "TEST HEAD",
            "original_rows": [{"material": "OLD SET", "code": "Set00001", "quantity": 1}],
            "rows": [
                {"material": "SET TEST", "code": "Set99999", "quantity": 1},
                {"material": "BOLT TEST", "code": "1010110001", "quantity": 3},
            ],
        })
        self.assertEqual(request["status"], "pending")
        self.assertEqual(request["targetDepartment"], "แผนกแรงสูง TAC")
        self.assertEqual(request["insulatorUpright"], 5.0)
        self.assertEqual(request["insulatorHorizontal"], 10.0)
        self.assertEqual(request["originalRows"], [{"material": "OLD SET", "code": "Set00001", "quantity": 1.0}])
        self.assertEqual(store.approved, [])

        reviewed = store.review_base_request(request["id"], True, "admin")
        self.assertEqual(reviewed["status"], "approved")
        self.assertEqual(len(store.approved), 2)
        self.assertTrue(all(row["request_id"] == request["id"] for row in store.approved))
        self.assertTrue(all(row["action"] == "replace" for row in store.approved))
        self.assertTrue(all(row["department"] == "แผนกแรงสูง TAC" for row in store.approved))
        self.assertTrue(all(row["insulator_upright"] == 5.0 for row in store.approved))
        self.assertTrue(all(row["insulator_horizontal"] == 10.0 for row in store.approved))

    def test_reject_does_not_publish_rows(self):
        store = MemoryBaseRequestStore()
        request = store.submit_base_request({
            "submitter_name": "ผู้ทดสอบ", "employee_id": "123456", "department": "กวว.",
            "size": "99", "head": "REJECTED",
            "rows": [{"material": "TEST", "code": "Set1", "quantity": 1}],
        })
        store.review_base_request(request["id"], False, "admin", "ข้อมูลไม่ครบ")
        self.assertEqual(store.approved, [])
        self.assertEqual(store.rows[0]["status"], "rejected")

    def test_copy_request_keeps_source_and_new_target(self):
        store = MemoryBaseRequestStore()
        request = store.submit_base_request({
            "submitter_name": "ผู้ทดสอบ", "employee_id": "123456", "department": "กวว.",
            "target_department": "แผนกแรงสูง", "action": "copy",
            "source_size": "12.2", "source_head": "DDE",
            "size": "12.2", "head": "DDE รุ่นแก้ไข",
            "insulator_upright": 6, "insulator_horizontal": 24,
            "original_rows": [{"material": "SET เดิม", "code": "Set1", "quantity": 1}],
            "rows": [{"material": "SET ใหม่", "code": "Set2", "quantity": 1}],
        })
        self.assertEqual(request["action"], "copy")
        self.assertEqual(request["sourceSize"], "12.2")
        self.assertEqual(request["sourceHead"], "DDE")
        store.review_base_request(request["id"], True, "admin")
        self.assertEqual(store.approved[0]["action"], "copy")
        self.assertEqual(store.approved[0]["source_size"], "12.2")
        self.assertEqual(store.approved[0]["source_head"], "DDE")

    def test_rename_request_requires_source(self):
        store = MemoryBaseRequestStore()
        with self.assertRaisesRegex(ValueError, "หัวเสาต้นฉบับ"):
            store.submit_base_request({
                "submitter_name": "ผู้ทดสอบ", "employee_id": "123456", "department": "กวว.",
                "action": "rename", "size": "12.2", "head": "ชื่อใหม่",
                "rows": [{"material": "SET", "code": "Set1", "quantity": 1}],
            })

    def test_clear_approved_history_keeps_published_rows_and_rejections(self):
        store = MemoryBaseRequestStore()
        base_payload = {
            "submitter_name": "ผู้ทดสอบ", "employee_id": "123456", "department": "กวว.",
            "size": "99", "rows": [{"material": "TEST", "code": "Set1", "quantity": 1}],
        }
        approved = store.submit_base_request({**base_payload, "head": "APPROVED"})
        rejected = store.submit_base_request({**base_payload, "head": "REJECTED"})
        store.review_base_request(approved["id"], True, "admin")
        store.review_base_request(rejected["id"], False, "admin")
        self.assertEqual(store.clear_approved_requests(), 1)
        self.assertEqual(len(store.approved), 1)
        remaining = store.list_base_requests()
        self.assertEqual([item["head"] for item in remaining], ["REJECTED"])


if __name__ == "__main__":
    unittest.main()
