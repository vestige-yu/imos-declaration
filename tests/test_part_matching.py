import importlib.util
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CONFIG = ROOT / "最新逻辑_请优先看" / "报关单配置关系表.xlsx"


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


MODULES = (
    load_module("web_app", ROOT / "app.py"),
    load_module("desktop_app", ROOT / "local_desktop_app" / "app.py"),
)


class PartMatchingTests(unittest.TestCase):
    CASES = (
        ("3073400200-00", "307340020000", True),
        ("001234", "1234", True),
        ("ABC001-01", "ABC001-02", True),
        ("CABC001", "ABC001", False),
        ("ＡＢＣ　００１", "ABC001", True),
        ("123456.0", "123456", True),
        ("ABO001", "AB0001", False),
        ("OLD123", "NEW456", False),
    )

    def test_confirmed_rules(self):
        for module in MODULES:
            for left, right, expected in self.CASES:
                with self.subTest(module=module.__name__, left=left, right=right):
                    actual = bool(
                        set(module.part_match_keys(left))
                        & set(module.part_match_keys(right))
                    )
                    self.assertEqual(expected, actual)

    def test_exact_version_weight_has_priority(self):
        for module in MODULES:
            mapping = {}
            for key in module.part_match_keys("ABC001-01"):
                mapping[key] = 10
            for key in module.part_match_keys("ABC001-02"):
                mapping[key] = 30 if key == "ABC001" else 20
            with self.subTest(module=module.__name__):
                self.assertEqual(10, module.part_lookup(mapping, "ABC001-01"))
                self.assertEqual(20, module.part_lookup(mapping, "ABC001-02"))

    def test_alphanumeric_version_suffix_uses_base_key(self):
        for module in MODULES:
            with self.subTest(module=module.__name__):
                self.assertIn(
                    "3031100034",
                    module.part_match_keys("3031100034-03A1"),
                )
                self.assertNotIn("PCTN", module.part_match_keys("PCTN-018"))
                self.assertNotIn("EE", module.part_match_keys("EE-0000000391"))

    def test_notes_are_not_part_numbers(self):
        for module in MODULES:
            with self.subTest(module=module.__name__):
                self.assertFalse(module.is_plausible_part_no("未找到图，待拍实物"))
                self.assertFalse(module.is_plausible_part_no("2024未出货"))
                self.assertTrue(module.is_plausible_part_no("EE-0000000391"))

    def test_real_sample_regressions(self):
        samples = (
            {
                "id": "SP26003466",
                "invoice": ROOT / "样例与测试资料/debug_01/SP26003466-CASCO Schoeller sea-20260710-Invoice.xls",
                "packing": ROOT / "样例与测试资料/debug_01/SP26003466-CASCO Schoeller sea-20260710-Packing.xls",
                "expected": (5415.0, 61271.58, 571.21, 571.2, 695.38, 5, 7),
            },
            {
                "id": "SP26000096",
                "invoice": ROOT / "样例与测试资料/test/SP26000096 Invoice -CASCO Imos Italia -sea 20260116-1.xls",
                "packing": ROOT / "样例与测试资料/test/SP26000096 Packing list -CASCO Imos Italia-sea 20260116-1.xls",
                "expected": (112546.0, 246950.57, 3428.89, 3428.89, 3935.85, 27, 15),
            },
            {
                "id": "SP26000001",
                "invoice": ROOT / "样例与测试资料/源文件_样例/SP26000001 Invoice -CASCO Imos Italia -sea 20260104.xls",
                "packing": ROOT / "样例与测试资料/源文件_样例/SP26000001 Packing list -CASCO Imos Italia-sea 20260104.xls",
                "expected": (148097.0, 413972.72, 4644.71, 4644.71, 5337.33, 36, 12),
            },
        )
        if not all(sample["invoice"].is_file() and sample["packing"].is_file() for sample in samples):
            self.skipTest("本地真实报关样例未包含在当前检出目录中")

        for module in MODULES:
            rules = module.load_rules(CONFIG)
            for sample in samples:
                with self.subTest(module=module.__name__, sample=sample["id"]):
                    invoice = module.parse_invoice(sample["invoice"])
                    packing = module.parse_packing(
                        sample["packing"],
                        invoice.get("contractNo"),
                    )
                    preview = module.merge_preview(invoice, packing, rules)
                    actual = (
                        preview["totals"]["quantity"],
                        preview["totals"]["amount"],
                        preview["totals"]["netWeight"],
                        preview["netWeight"],
                        preview["grossWeight"],
                        preview["packageCount"],
                        len(preview["commodityLines"]),
                    )
                    self.assertEqual(sample["expected"], actual)
                    self.assertEqual([], preview["warnings"])


if __name__ == "__main__":
    unittest.main()
