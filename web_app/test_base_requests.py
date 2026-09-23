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


class BaseRequestTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
