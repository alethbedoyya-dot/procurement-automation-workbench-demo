import json
import os
import tempfile
import unittest
from unittest import mock

import web_query


class SavedLoginStateTests(unittest.TestCase):
    def test_saved_state_is_bound_to_the_current_windows_user(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "auth_state.json")
            meta_path = os.path.join(temp_dir, "auth_state.meta.json")
            with open(state_path, "w", encoding="utf-8") as file_obj:
                json.dump({"cookies": []}, file_obj)
            with open(meta_path, "w", encoding="utf-8") as file_obj:
                json.dump({"computer": "pc01", "user": "mary"}, file_obj)

            with mock.patch.object(web_query, "AUTH_STATE_FILE", state_path), \
                    mock.patch.object(web_query, "AUTH_META_FILE", meta_path), \
                    mock.patch.dict(
                        os.environ,
                        {"COMPUTERNAME": "PC01", "USERNAME": "Mary"},
                        clear=False,
                    ):
                self.assertTrue(web_query._saved_auth_is_local())
                os.environ["USERNAME"] = "OtherUser"
                self.assertFalse(web_query._saved_auth_is_local())


class RecoveryAndFieldParsingTests(unittest.TestCase):
    def test_manifest_wbs_check_only_returns_confirmed_missing_results(self):
        manifest = {
            "results": [
                {"po": "9000000001", "wbs": "WBS-001"},
                {"po": "9000000002", "wbs": ""},
                {"po": "9000000003", "wbs": None},
            ]
        }

        self.assertEqual(
            web_query._find_manifest_pos_with_missing_wbs(
                manifest, ["9000000001", "9000000002", "9000000003", "9000000004"]
            ),
            ["9000000002", "9000000003"],
        )

    def test_project_name_parser_preserves_symbols_and_stops_at_next_field(self):
        value, source = web_query._extract_project_name_from_body_text(
            "项目名称：NO.2026G01地块/A&B区#1【一期】 项目经理 演示用户"
        )

        self.assertEqual(value, "NO.2026G01地块/A&B区#1【一期】")
        self.assertEqual(source, "同行完整文本")


if __name__ == "__main__":
    unittest.main()
