import runpy
import sys
import tempfile
import types
import unittest
import json
from pathlib import Path
from unittest.mock import patch

import openpyxl

from workbench_config import ConfigurationError, load_web_urls


PROJECT_DIR = Path(__file__).resolve().parents[1]
GUI_SCRIPT = PROJECT_DIR / "procurement_workbench.py"


class PublicDemoConfigurationTests(unittest.TestCase):
    def setUp(self):
        ttk_stub = types.ModuleType("ttkbootstrap")
        with patch.dict(sys.modules, {"ttkbootstrap": ttk_stub}):
            self.module = runpy.run_path(str(GUI_SCRIPT))

    def test_categories_use_obvious_synthetic_identifiers(self):
        categories = self.module["CATEGORIES"]

        self.assertIn("Demo Category A", categories)
        self.assertGreaterEqual(len(categories), 3)
        for category_name, config in categories.items():
            self.assertTrue(category_name.startswith("Demo Category"))
            self.assertTrue(config["label"].startswith("Demo Category"))
            self.assertTrue(config["data_sheet"].startswith("DEMO_"))
            for material_id in config["filter_materials"]:
                self.assertTrue(str(material_id).startswith("DEMO-MAT-"))

    def test_demo_material_can_be_filtered_without_network_access(self):
        category = "Demo Category B"
        config = self.module["CATEGORIES"][category]

        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "demo_purchase_records.xlsx"
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.title = "Sheet1"
            worksheet.append(["物料", "订单净值", "短文本", "净价", "采购订单数量"])
            worksheet.append([config["filter_materials"][0], 240, "DEMO-MATERIAL", 160, 2])
            worksheet.append(["DEMO-MAT-UNRELATED", 300, "UNRELATED", 120, 1])
            workbook.save(workbook_path)
            workbook.close()

            function_globals = self.module["_enhance_and_filter"].__globals__
            with patch.dict(function_globals, {"EXCEL_FILE": str(workbook_path)}):
                _, row_count, _ = self.module["_enhance_and_filter"](
                    config, lambda _message: None
                )

            self.assertEqual(row_count, 1)
            workbook = openpyxl.load_workbook(workbook_path, data_only=True)
            self.assertIn(config["data_sheet"], workbook.sheetnames)
            self.assertEqual(workbook[config["data_sheet"]].cell(2, 1).value, config["filter_materials"][0])
            workbook.close()

    def test_configuration_template_is_parseable_and_uses_placeholder_urls(self):
        template_path = PROJECT_DIR / "config.example.json"
        template = json.loads(template_path.read_text(encoding="utf-8"))

        self.assertEqual(set(template["web"]), {"home_url", "search_url"})
        for value in template["web"].values():
            self.assertIn(".example", value)

    def test_missing_runtime_urls_fail_before_any_browser_workflow(self):
        with self.assertRaises(ConfigurationError):
            load_web_urls(config_path=PROJECT_DIR / "missing-config.json", environ={})


if __name__ == "__main__":
    unittest.main()
