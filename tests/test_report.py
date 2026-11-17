from argparse import Namespace
from datetime import date
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from unittest.mock import MagicMock

from openpyxl import load_workbook

from cs_returns_report.cli import API, cached_fetch, load_annotations, run, write_workbook, zopo_rates, append_annotation_templates, review_documents
from cs_returns_report.rules import DEFAULT_SCOPE, THREE_YEAR_SCOPE, build_row, retired_at_proposal, select_expected


def insolvency(proposal="2023-01-01"):
    return {"self": {"id": 12023}, "reference_number": {"case_number": 1, "year": 2023}, "proposal_date": proposal, "state": "ODDLUŽENÍ", "debtors": [], "discharge": {"href": "/discharge/1"}}


def discharge(default=30, short=50, approved=True):
    return {"discharge_allowed_date": "2023-02-01", "approval": {"discharge_approved_date": "2023-03-01"} if approved else None, "satisfaction_estimates": [{"estimation_scope": DEFAULT_SCOPE, "rate": default}, {"estimation_scope": THREE_YEAR_SCOPE, "rate": short}]}


class RulesTests(unittest.TestCase):
    def test_regime_boundaries_and_threshold(self):
        d = discharge(short=60)
        for proposal, expected in [("2021-05-31", 30), ("2021-06-01", 60), ("2024-12-31", 60), ("2025-01-01", None)]:
            with self.subTest(proposal=proposal):
                self.assertEqual(select_expected(insolvency(proposal), d, [])[0], expected)
        d["approval"]["approved_expected_return_rate"] = 0
        self.assertEqual(select_expected(insolvency("2025-01-01"), d, [])[0], 0)

    def test_retirement_upper_limit_and_missing_birth_date(self):
        born = [{"birth_date": "1959-01-02", "gender": "F"}]
        self.assertFalse(retired_at_proposal(born, date(2023, 1, 1)))
        self.assertTrue(retired_at_proposal(born, date(2023, 1, 2)))
        self.assertFalse(retired_at_proposal([{}], date(2023, 1, 2)))
        self.assertEqual(select_expected(insolvency(), discharge(), born), (30, []))
        self.assertEqual(select_expected(insolvency("2023-01-02"), discharge(), born), (50, []))
        self.assertEqual(select_expected(insolvency(), discharge(short=None), [{"birth_date": "1940-01-01"}]), (30, []))

    def test_joint_debtor_and_zero(self):
        self.assertEqual(select_expected(insolvency(), discharge(short=0), [{}, {"birth_date": "1940-01-01"}]), (0, []))

    def test_annotation_gate_and_overrides(self):
        ins = insolvency("2025-01-01")
        row, reasons = build_row(ins, discharge(), [], None, {})
        self.assertTrue(reasons)
        self.assertIsNone(row["PREDP_USP"])
        row, reasons = build_row(ins, discharge(), [], None, {"expected_return_rate": 0})
        self.assertEqual((row["PREDP_USP"], reasons), (0, []))
        row, reasons = build_row(ins, discharge(), [], None, {"note": "Not stated in approval"})
        self.assertEqual(reasons, [])
        self.assertEqual(row["NOTE"], "Not stated in approval")
        self.assertEqual(build_row(ins, discharge(approved=False), [], None, {})[1], [])
        self.assertTrue(build_row(ins, {}, [], None, {})[1])

    def test_conflicting_estimates_require_review(self):
        d = discharge()
        d["satisfaction_estimates"].append({"estimation_scope": DEFAULT_SCOPE, "rate": 40})
        self.assertTrue(select_expected(insolvency(), d, [])[1])


class PipelineTests(unittest.TestCase):
    def test_creditor_page_progress(self):
        with patch.dict(os.environ, {"REKENBER_API_TOKEN": "test"}):
            api = API()
        bar = MagicMock()
        with patch("cs_returns_report.cli.tqdm") as progress, patch.object(api, "get", side_effect=[
            {"results": [1], "pagination": {"total_pages": 2, "next": "/next"}},
            {"results": [2], "pagination": {"total_pages": 2}},
        ]) as get:
            progress.return_value.__enter__.return_value = bar
            self.assertEqual(api.pages("/start", progress_label="Creditor test", progress_position=1), [1, 2])
            self.assertEqual(bar.total, 2)
            self.assertEqual(bar.update.call_count, 2)
            self.assertEqual(get.call_args_list[0].args, ("/start?",))
            bar.set_postfix_str.assert_any_call("requesting page 1; 0 rows loaded")
            bar.set_postfix_str.assert_any_call("2 rows loaded")

    def test_review_document_types_and_cutoff(self):
        def event(kind, published):
            return {"publication_date": published, "event_types": [{"id": kind}], "document": {"id": kind}}
        zpo = event(1036, "2026-03-01")
        fallback = event(1037, "2026-04-01")
        earlier_fallback = {**event(1037, "2026-02-01"), "document": {"id": 987}}
        older_case = insolvency("2024-12-31")
        self.assertTrue(review_documents(older_case, [fallback, zpo])[0].endswith("=1036"))
        self.assertTrue(review_documents(older_case, [fallback, earlier_fallback])[0].endswith("=1037"))
        for kind in [174, 175, 1065]:
            with self.subTest(kind=kind):
                source, zpo_url = review_documents(insolvency("2025-01-01"), [zpo, event(kind, "2025-02-01")])
                self.assertTrue(source.endswith(f"={kind}"))
                self.assertTrue(zpo_url.endswith("=1036"))
        events = [zpo, event(174, "2025-01-01"), event(175, "2025-02-01"), event(1065, "2025-01-15")]
        self.assertTrue(review_documents(insolvency("2025-01-01"), events)[0].endswith("=175"))
        self.assertEqual(review_documents(insolvency("2024-12-31"), events), (review_documents(insolvency("2025-01-01"), events)[1],) * 2)
        self.assertIsNone(review_documents(insolvency("2025-01-01"), [zpo])[0])

    def test_append_only_templates_preserve_existing_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "annotations.jsonl"
            original = '{"insolvency_id":"1","expected_return_rate":42,"note":"reviewed"}'
            path.write_text(original)
            missing = [{"insolvency_id": "1"}, {"insolvency_id": "2", "reason": "missing"}]
            append_annotation_templates(path, missing, load_annotations(path))
            content = path.read_text()
            self.assertTrue(content.startswith(original + "\n"))
            self.assertIsNone(load_annotations(path)["2"]["expected_return_rate"])
            append_annotation_templates(path, missing, load_annotations(path))
            self.assertEqual(path.read_text(), content)

    def test_pagination_and_cycles(self):
        with patch.dict(os.environ, {"REKENBER_API_TOKEN": "test"}):
            api = API()
        with patch.object(api, "get", side_effect=[{"results": [1], "pagination": {"next": "/next"}}, {"results": [2]}]):
            self.assertEqual(api.pages("/start"), [1, 2])
        with patch.object(api, "get", return_value={"results": [], "pagination": {"next": "/start?"}}):
            with self.assertRaises(ValueError):
                api.pages("/start")

    def test_cache_empty_and_checkpoint_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.assertEqual(cached_fetch(root, "test", ["a"], lambda k: []), {"a": []})
            self.assertEqual(cached_fetch(root, "test", ["a"], lambda k: self.fail("cache missed")), {"a": []})
            with self.assertRaises(RuntimeError):
                cached_fetch(root, "test", ["a", "b"], lambda k: (_ for _ in ()).throw(RuntimeError("failed")))
            self.assertEqual(json.loads((root / "test_cache.json").read_text()), {"a": []})

    def test_latest_zopo_no_older_fallback(self):
        events = {"1": [{"publication_date": "2026-01-01", "event_types": [{"id": 39}], "document": {"id": 1}}, {"publication_date": "2026-02-01", "event_types": [{"id": 39}], "document": {"id": 2}}]}
        with tempfile.TemporaryDirectory() as directory, patch.dict(os.environ, {"REKENBER_ANNOTATION_DB_URL": "test"}), patch("cs_returns_report.cli.fetch_extracts", return_value={"2": {"obligations_summary": {"current_return_rate": 0}}}) as fetch:
            self.assertEqual(zopo_rates(Path(directory), events)[0], {"1": 0})
            fetch.assert_called_once_with("test", ["2"])
            zopo_rates(Path(directory), events)
            self.assertEqual(fetch.call_count, 1)

    def test_zopo_missing_extract_and_database_failure(self):
        events = {"1": [{"publication_date": "2026-01-01", "event_types": [{"id": 39}], "document": {"id": 7}}]}
        with tempfile.TemporaryDirectory() as directory, patch.dict(os.environ, {"REKENBER_ANNOTATION_DB_URL": "test"}):
            root = Path(directory)
            with patch("cs_returns_report.cli.fetch_extracts", side_effect=RuntimeError("connection failed")):
                with self.assertRaises(RuntimeError):
                    zopo_rates(root, events)
            self.assertFalse((root / "zopo_extracts_cache.json").exists())
            with patch("cs_returns_report.cli.fetch_extracts", return_value={"7": None}):
                self.assertEqual(zopo_rates(root, events)[0], {"1": None})
            with patch("cs_returns_report.cli.fetch_extracts", side_effect=AssertionError("must use cache")):
                self.assertEqual(zopo_rates(root, events)[0], {"1": None})

    def test_annotation_roundtrip_and_duplicate_rejection(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "annotations.jsonl"
            item = {"insolvency_id": "12023", "expected_return_rate": 0, "note": "=literal"}
            append_annotation_templates(path, [item], {})
            self.assertEqual(load_annotations(path)["12023"]["note"], "=literal")
            path.write_text(path.read_text() * 2)
            with self.assertRaises(ValueError):
                load_annotations(path)

    def test_full_annotation_loop_deduplication_and_snapshot(self):
        with tempfile.TemporaryDirectory() as directory, patch.dict(os.environ, {"REKENBER_API_TOKEN": "test", "REKENBER_ANNOTATIONS_PATH": "data/annotations.jsonl"}):
            previous = Path.cwd()
            os.chdir(directory)
            try:
                ins = insolvency("2025-01-01")
                args = Namespace(report_name="2026-11", max_workers=1, refresh=False, skip_zopo=False)
                def pages(path, **params):
                    return [ins] if "insolvencies" in path else []
                with patch.object(API, "pages", side_effect=pages), patch.object(API, "get", side_effect=lambda url: discharge() if "discharge" in url else ins):
                    with self.assertRaises(SystemExit):
                        run(args)
                self.assertFalse(Path("workspace/2026-11/final_report.xlsx").exists())
                annotations = load_annotations(Path("data/annotations.jsonl"))
                annotations["12023"]["expected_return_rate"] = 42
                Path("data/annotations.jsonl").write_text("".join(json.dumps(item) + "\n" for item in annotations.values()))
                with patch.object(API, "get", side_effect=AssertionError("must use cache")), patch.object(API, "pages", side_effect=AssertionError("must use cache")):
                    run(args)
                workbook = load_workbook("workspace/2026-11/final_report.xlsx")
                sheet = workbook["Report"]
                self.assertEqual(sheet.max_row, 2)
                self.assertEqual(sheet["A2"].value, "INS 1/2023")
                self.assertEqual(sheet["D2"].value, 42)
                self.assertEqual(sheet["B2"].value.date(), date(2023, 2, 1))
                workbook.close()
                # A refreshed snapshot applies current state and removes the old output.
                args.refresh = True
                with patch.object(API, "pages", return_value=[]):
                    run(args)
                workbook = load_workbook("workspace/2026-11/final_report.xlsx")
                self.assertEqual(workbook["Report"].max_row, 1)
                workbook.close()
                self.assertTrue(Path("workspace/2026-11/previous_final_report.xlsx").exists())
            finally:
                os.chdir(previous)


if __name__ == "__main__":
    unittest.main()
