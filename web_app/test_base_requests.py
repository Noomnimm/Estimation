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


class MemoryProjectStore(GoogleSheetProjectStore):
    def __init__(self):
        super().__init__()
        self.project_rows = []
        self.folder_rows = []

    def _read_rows(self):
        return [dict(row) for row in self.project_rows]

    def _write_record(self, record, row_number):
        saved = dict(record)
        if row_number:
            saved["_row_number"] = row_number
            self.project_rows[row_number - 2] = saved
        else:
            saved["_row_number"] = len(self.project_rows) + 2
            self.project_rows.append(saved)

    def _read_named_rows(self, sheet_name, headers, key):
        return [dict(row) for row in self.folder_rows]

    def _append_named_record(self, sheet_name, headers, record):
        saved = dict(record)
        saved["_row_number"] = len(self.folder_rows) + 2
        self.folder_rows.append(saved)

    def _update_named_record(self, sheet_name, headers, record, row_number):
        saved = dict(record)
        saved["_row_number"] = row_number
        self.folder_rows[row_number - 2] = saved


class BaseRequestTests(unittest.TestCase):
    def test_cloud_projects_are_isolated_by_owner_email(self):
        store = MemoryProjectStore()
        alice = {"email": "alice@example.com", "name": "Alice"}
        bob = {"email": "bob@example.com", "name": "Bob"}
        payload = {"id": "same-project-id", "name": "งานของฉัน", "pages": [], "departments": {}}
        store.save_project(payload, alice)
        store.save_project({**payload, "name": "งานของ Bob"}, bob)

        self.assertEqual([project["name"] for project in store.list_projects(alice)], ["งานของฉัน"])
        self.assertEqual([project["name"] for project in store.list_projects(bob)], ["งานของ Bob"])

        store.delete_project("same-project-id", bob)
        self.assertEqual(len(store.list_projects(alice)), 1)
        self.assertEqual(store.list_projects(bob), [])

    def test_cloud_deleted_project_can_be_restored_or_purged(self):
        store = MemoryProjectStore()
        user = {"email": "alice@example.com", "name": "Alice"}
        payload = {"id": "p-trash", "name": "งานทดสอบถังขยะ", "pages": [], "departments": {}}
        store.save_project(payload, user)
        store.delete_project("p-trash", user)
        self.assertEqual(store.list_projects(user), [])
        self.assertEqual([project["id"] for project in store.list_deleted_projects(user)], ["p-trash"])

        restored = store.restore_project("p-trash", user)
        self.assertEqual(restored["id"], "p-trash")
        self.assertEqual(len(store.list_projects(user)), 1)

        store.delete_project("p-trash", user)
        store.permanently_delete_project("p-trash", user)
        self.assertEqual(store.list_deleted_projects(user), [])

    def test_admin_trash_lists_all_owners_and_targets_owner(self):
        store = MemoryProjectStore()
        alice = {"email": "alice@example.com", "name": "Alice"}
        bob = {"email": "bob@example.com", "name": "Bob"}
        payload = {"id": "same-id", "name": "งานในถังขยะ", "pages": [], "departments": {}}
        store.save_project(payload, alice)
        store.save_project(payload, bob)
        store.delete_project("same-id", alice)
        store.delete_project("same-id", bob)
        self.assertEqual({project["createdBy"] for project in store.list_all_deleted_projects()}, {"alice@example.com", "bob@example.com"})

        store.restore_project_as_admin("same-id", "alice@example.com")
        self.assertEqual([project["createdBy"] for project in store.list_all_deleted_projects()], ["bob@example.com"])
        store.permanently_delete_project_as_admin("same-id", "bob@example.com")
        self.assertEqual(store.list_all_deleted_projects(), [])

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

    def test_cloud_project_folders_are_owned_and_projects_can_move(self):
        store = MemoryProjectStore()
        alice = {"email": "alice@example.com", "name": "Alice"}
        bob = {"email": "bob@example.com", "name": "Bob"}
        alice_folder = store.create_project_folder("งานปรับปรุง", alice)
        store.create_project_folder("งานของ Bob", bob)
        store.save_project({"id": "p1", "name": "โครงการหนึ่ง", "pages": [], "departments": {}}, alice)

        moved = store.move_project("p1", alice_folder["id"], alice)
        self.assertEqual(moved["folderId"], alice_folder["id"])
        self.assertEqual([folder["name"] for folder in store.list_project_folders(alice)], ["งานปรับปรุง"])
        with self.assertRaises(ValueError):
            store.delete_project_folder(alice_folder["id"], alice)

        store.move_project("p1", "", alice)
        store.delete_project_folder(alice_folder["id"], alice)
        self.assertEqual(store.list_project_folders(alice), [])

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

    def test_public_pending_requests_hide_submitter_identity(self):
        store = MemoryBaseRequestStore()
        pending = store.submit_base_request({
            "submitter_name": "ผู้ทดสอบ", "employee_id": "123456", "department": "กวว.",
            "size": "12", "head": "PENDING",
            "rows": [{"material": "TEST", "code": "Set1", "quantity": 1}],
        })
        rejected = store.submit_base_request({
            "submitter_name": "อีกคน", "employee_id": "654321", "department": "กวว.",
            "size": "12", "head": "REJECTED",
            "rows": [{"material": "TEST", "code": "Set2", "quantity": 1}],
        })
        store.review_base_request(rejected["id"], False, "admin")

        requests = store.list_pending_base_requests()

        self.assertEqual([request["id"] for request in requests], [pending["id"]])
        self.assertNotIn("submitterName", requests[0])
        self.assertNotIn("employeeId", requests[0])
        self.assertEqual(requests[0]["rows"][0]["code"], "Set1")

    def test_just_submitted_request_survives_temporarily_stale_sheet_read(self):
        store = MemoryBaseRequestStore()
        request = store.submit_base_request({
            "submitter_name": "ผู้ทดสอบ", "employee_id": "123456", "department": "กวว.",
            "size": "12", "head": "WAITING FOR SHEET",
            "rows": [{"material": "TEST", "code": "Set1", "quantity": 1}],
        })
        saved_rows = store.rows
        store.rows = []

        requests = store.list_base_requests()

        self.assertEqual([item["id"] for item in requests], [request["id"]])
        store.rows = saved_rows

    def test_admin_can_edit_request_while_approving_and_original_is_audited(self):
        store = MemoryBaseRequestStore()
        request = store.submit_base_request({
            "submitter_name": "ผู้ทดสอบ", "employee_id": "123456", "department": "กวว.",
            "size": "12", "head": "หัวเดิม", "insulator_upright": 1, "insulator_horizontal": 2,
            "rows": [{"material": "รายการเดิม", "code": "Set1", "quantity": 1}],
        })

        reviewed = store.review_base_request(request["id"], True, "admin", edits={
            "size": "12.2", "head": "หัวที่ Admin แก้", "insulator_upright": 3,
            "insulator_horizontal": 4,
            "rows": [{"material": "รายการที่แก้", "code": "Set2", "quantity": 5}],
        })

        self.assertTrue(reviewed["adminEdited"])
        self.assertEqual(reviewed["submitted"]["head"], "หัวเดิม")
        self.assertEqual(reviewed["submitted"]["new"][0]["code"], "Set1")
        self.assertEqual(reviewed["head"], "หัวที่ Admin แก้")
        self.assertEqual(store.approved[0]["code"], "Set2")
        self.assertEqual(store.approved[0]["quantity"], 5.0)
        self.assertEqual(store.approved[0]["insulator_upright"], 3.0)

    def test_invalid_admin_edit_keeps_request_pending(self):
        store = MemoryBaseRequestStore()
        request = store.submit_base_request({
            "submitter_name": "ผู้ทดสอบ", "employee_id": "123456", "department": "กวว.",
            "size": "12", "head": "TEST",
            "rows": [{"material": "รายการเดิม", "code": "Set1", "quantity": 1}],
        })
        with self.assertRaisesRegex(ValueError, "อย่างน้อย 1 รายการ"):
            store.review_base_request(request["id"], True, "admin", edits={
                "size": "12", "head": "TEST", "rows": [],
            })
        self.assertEqual(store.rows[0]["status"], "pending")
        self.assertEqual(store.approved, [])

    def test_copy_request_keeps_source_and_new_target(self):
        store = MemoryBaseRequestStore()
        request = store.submit_base_request({
            "submitter_name": "ผู้ทดสอบ", "employee_id": "123456", "department": "กวว.",
            "target_department": "แผนกแรงสูง", "action": "copy",
            "source_size": "12.2", "source_head": "DDE",
            "size": "12.2", "head": "DDE รุ่นแก้ไข",
            "insulator_upright": 6, "insulator_horizontal": 24,
            "image_file_id": "drive-image-1", "image_name": "dde.jpg", "image_mime_type": "image/jpeg",
            "original_rows": [{"material": "SET เดิม", "code": "Set1", "quantity": 1}],
            "rows": [{"material": "SET ใหม่", "code": "Set2", "quantity": 1}],
        })
        self.assertEqual(request["action"], "copy")
        self.assertEqual(request["sourceSize"], "12.2")
        self.assertEqual(request["sourceHead"], "DDE")
        self.assertEqual(request["imageFileId"], "drive-image-1")
        store.review_base_request(request["id"], True, "admin")
        self.assertEqual(store.approved[0]["action"], "copy")
        self.assertEqual(store.approved[0]["source_size"], "12.2")
        self.assertEqual(store.approved[0]["source_head"], "DDE")
        self.assertEqual(store.approved[0]["image_file_id"], "drive-image-1")

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
