"""Tests for the Phase 0 discovery engine.

Run with:  python -m unittest discover -s tests -v

Uses unittest rather than pytest deliberately: the whole engine must run on a
corporate laptop with nothing installed beyond PyYAML, and a test suite that
needs a pip install is a test suite that does not get run.
"""

from __future__ import annotations

import json
import shutil
import tempfile
import unittest
import zipfile
from dataclasses import replace
from pathlib import Path
from typing import Dict, List

from sc.configuration import Config, ConfigError, load_config
from sc.discovery import classify
from sc.discovery.csv_probe import probe_csv
from sc.discovery.datamashup import (
    MashupError,
    extract_sources,
    parse_queries,
    probe_mashup,
    read_section_m,
    split_declarations,
)
from sc.discovery.ooxml import NotAnOoxmlPackage, column_index, open_workbook, resolve_part
from sc.discovery.report import render_discovery_md, write_manifest
from sc.discovery.scan import run_discovery
from sc.discovery.vba import probe_vba
from sc.discovery.workbook_probe import probe_workbook
from sc.discovery.xlsb import probe_xlsb
from tests.build_estate import build
from tests.make_fixtures import build_mashup_blob, build_vba_project, write_xlsb, write_xlsx


class TestOoxmlPlumbing(unittest.TestCase):
    def test_relationship_target_resolution(self) -> None:
        self.assertEqual(resolve_part("xl/workbook.xml", "worksheets/sheet1.xml"), "xl/worksheets/sheet1.xml")
        self.assertEqual(resolve_part("xl/workbook.xml", "/xl/theme/theme1.xml"), "xl/theme/theme1.xml")
        self.assertEqual(resolve_part("xl/worksheets/sheet1.xml", "../tables/table1.xml"), "xl/tables/table1.xml")

    def test_column_index(self) -> None:
        self.assertEqual([column_index(c) for c in ("A", "Z", "AA", "BX")], [0, 25, 26, 75])
        self.assertEqual(column_index(""), -1)

    def test_non_package_raises_clearly(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            fake: Path = Path(tmp) / "legacy.xls"
            fake.write_bytes(b"\xd0\xcf\x11\xe0legacy binary")
            with self.assertRaises(NotAnOoxmlPackage):
                open_workbook(str(fake))


class TestMParser(unittest.TestCase):
    SECTION: str = (
        'section Section1;\n\n'
        'shared #"Open PO" = let\n'
        '    S = Csv.Document(File.Contents("C:\\Temp\\po.csv"),[Delimiter=";"]),\n'
        '    // a semicolon in a comment ; must not split\n'
        '    N = "a string with ; and an escaped quote "" inside"\n'
        'in\n'
        '    N;\n\n'
        '[Description="second"]\n'
        'shared Second = let X = #"Open PO" in X;\n'
    )

    def test_semicolons_inside_strings_and_comments_do_not_split(self) -> None:
        # section header + 2 queries
        self.assertEqual(len(split_declarations(self.SECTION)), 3)

    def test_query_names_and_metadata(self) -> None:
        queries = parse_queries(self.SECTION)
        self.assertEqual([q.name for q in queries], ["Open PO", "Second"])
        self.assertIn("second", queries[1].metadata)

    def test_query_reference_is_a_dependency(self) -> None:
        second = parse_queries(self.SECTION)[1]
        self.assertIn("Open PO", [s["location"] for s in second.sources if s["kind"] == "query_ref"])

    def test_source_kinds_are_typed(self) -> None:
        body = (
            'let A = File.Contents("\\\\HOST\\share\\b.xlsx"), '
            'B = SharePoint.Files("https://predatorgroup.sharepoint.com/sites/SC"), '
            'C = Sql.Database("SRV01","NetSuite") in A'
        )
        kinds: Dict[str, str] = {s["location"]: s["kind"] for s in extract_sources(body)}
        self.assertEqual(kinds["\\\\HOST\\share\\b.xlsx"], "file")
        self.assertEqual(kinds["https://predatorgroup.sharepoint.com/sites/SC"], "sharepoint")
        self.assertEqual(kinds["SRV01"], "sql")

    def test_each_location_reported_once(self) -> None:
        body = 'let A = Excel.Workbook(File.Contents("\\\\HOST\\share\\b.xlsx")) in A'
        locations: List[str] = [s["location"] for s in extract_sources(body)]
        self.assertEqual(len(locations), len(set(locations)))

    def test_safe_filenames(self) -> None:
        query = parse_queries(self.SECTION)[0]
        self.assertEqual(query.safe_filename(), "Open_PO.m")


class TestMashupBlob(unittest.TestCase):
    def test_round_trip_through_a_real_nested_zip(self) -> None:
        section: str = 'section Section1;\n\nshared Q = let X = 1 in X;\n'
        with tempfile.TemporaryDirectory() as tmp:
            path: Path = Path(tmp) / "wb.xlsx"
            write_xlsx(path, [("Sheet1", [["A", "B"], [1, 2]], "visible")], mashup_section=section)
            with open_workbook(str(path)) as zf:
                probe = probe_mashup(zf)
        self.assertTrue(probe.found)
        self.assertIsNone(probe.error)
        self.assertEqual(probe.query_names, ["Q"])

    def test_absent_power_query_is_not_an_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path: Path = Path(tmp) / "plain.xlsx"
            write_xlsx(path, [("Sheet1", [["A"], [1]], "visible")])
            with open_workbook(str(path)) as zf:
                probe = probe_mashup(zf)
        self.assertFalse(probe.found)
        self.assertIsNone(probe.error)

    def test_corrupt_blob_reports_rather_than_crashes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path: Path = Path(tmp) / "bad.xlsx"
            write_xlsx(path, [("Sheet1", [["A"], [1]], "visible")], mashup_section="x")
            # Replace the payload with base64 that decodes to something that is not a zip.
            source = zipfile.ZipFile(path)
            entries = {n: source.read(n) for n in source.namelist()}
            source.close()
            entries["customXml/item1.xml"] = (
                b'<?xml version="1.0"?><DataMashup '
                b'xmlns="http://schemas.microsoft.com/DataMashup">bm90YXppcA==</DataMashup>'
            )
            with zipfile.ZipFile(path, "w") as rebuilt:
                for name, payload in entries.items():
                    rebuilt.writestr(name, payload)
            with open_workbook(str(path)) as zf:
                probe = probe_mashup(zf)
        self.assertTrue(probe.found)
        self.assertIsNotNone(probe.error)
        self.assertEqual(probe.queries, [])


class TestWorkbookProbe(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.mkdtemp()
        self.path = Path(self.tmp) / "probe.xlsx"
        rows = [
            ["Quarterly Report", None, None],
            [None, None, None],
            ["SKU", "On Hand", "Location"],
            ["PRED-1", 10, "JAX"],
            ["PRED-2", 20, "JAX"],
        ]
        write_xlsx(
            self.path,
            [("Data", rows, "visible"), ("Hidden_Calc", [["a", "b"], [1, 2]], "hidden")],
            tables={"Data": ("tbl_Data", "A3:C5", ["SKU", "On Hand", "Location"])},
            formula_cells={"Data": [(4, 1, "SUM(B4:B5)")]},
            defined_names=[("MyRange", "Data!$A$3:$C$5")],
        )

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_header_row_detected_below_a_title_block(self) -> None:
        sheet = probe_workbook(str(self.path)).sheets[0]
        self.assertEqual(sheet.header_row, 3)
        self.assertEqual(sheet.headers, ["SKU", "On Hand", "Location"])
        self.assertEqual(sheet.data_row_estimate, 2)

    def test_hidden_sheets_are_reported(self) -> None:
        probe = probe_workbook(str(self.path))
        self.assertEqual(probe.visible_sheets, ["Data"])
        self.assertEqual(probe.hidden_sheets, ["Hidden_Calc"])

    def test_tables_defined_names_and_formulas(self) -> None:
        probe = probe_workbook(str(self.path))
        self.assertEqual(probe.sheets[0].tables[0]["name"], "tbl_Data")
        self.assertEqual(probe.sheets[0].formula_count, 1)
        self.assertEqual([n["name"] for n in probe.defined_names], ["MyRange"])

    def test_inline_strings_resolve(self) -> None:
        path: Path = Path(self.tmp) / "inline.xlsx"
        write_xlsx(
            path,
            [("Inline", [["SKU", "Qty"], ["PRED-9", 5]], "visible")],
            inline_string_sheets=["Inline"],
        )
        self.assertEqual(probe_workbook(str(path)).sheets[0].headers, ["SKU", "Qty"])

    def test_external_link_target_is_captured(self) -> None:
        path: Path = Path(self.tmp) / "linked.xlsx"
        write_xlsx(path, [("S", [["A"], [1]], "visible")],
                   external_link_target=r"file:///\\HOST\share\other.xlsx")
        self.assertEqual(probe_workbook(str(path)).external_links,
                         [r"file:///\\HOST\share\other.xlsx"])


class TestXlsb(unittest.TestCase):
    def test_sheet_names_hidden_state_and_headers(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path: Path = Path(tmp) / "book.xlsb"
            write_xlsb(path, [
                ("Overstock", [["SKU", "Warehouse", "On Hand"], ["A", "JAX", "10"]], "visible"),
                ("Notes", [["note"], ["x"]], "hidden"),
            ])
            sheets = probe_xlsb(str(path))
        self.assertEqual([s.name for s in sheets], ["Overstock", "Notes"])
        self.assertEqual([s.state for s in sheets], ["visible", "hidden"])
        self.assertEqual(sheets[0].headers, ["SKU", "Warehouse", "On Hand"])
        self.assertEqual(sheets[0].row_count, 2)
        self.assertIsNone(sheets[0].error)


class TestVba(unittest.TestCase):
    def test_module_names_without_execution(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path: Path = Path(tmp) / "macro.xlsm"
            write_xlsx(path, [("S", [["A"], [1]], "visible")],
                       vba=build_vba_project("SCUpdate", ["mod_A", "mod_B"], ["cls_C"], protected=True))
            with open_workbook(str(path)) as zf:
                probe = probe_vba(zf)
        self.assertTrue(probe.present)
        self.assertTrue(probe.protected)
        self.assertEqual(probe.project_name, "SCUpdate")
        for expected in ("mod_A", "mod_B", "cls_C"):
            self.assertIn(expected, probe.module_names)

    def test_unprotected_project_not_flagged(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path: Path = Path(tmp) / "open.xlsm"
            write_xlsx(path, [("S", [["A"], [1]], "visible")],
                       vba=build_vba_project("Open", ["mod_A"], [], protected=False))
            with open_workbook(str(path)) as zf:
                probe = probe_vba(zf)
        self.assertFalse(probe.protected)


class TestCsvProbe(unittest.TestCase):
    def test_bom_delimiter_headers_and_ragged_rows(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path: Path = Path(tmp) / "export.csv"
            path.write_bytes(
                "\ufeffPO #;Line;Item;Qty\r\nPO1;1;A;5\r\nPO2;1;B\r\nPO3;1;C;7\r\n".encode("utf-8")
            )
            probe = probe_csv(path, 65536)
        self.assertEqual(probe.encoding, "utf-8-sig")
        self.assertEqual(probe.delimiter, ";")
        self.assertEqual(probe.headers, ["PO #", "Line", "Item", "Qty"])
        self.assertEqual(probe.row_count, 3)
        self.assertEqual(probe.ragged_rows, 1)

    def test_empty_file_reports_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path: Path = Path(tmp) / "empty.csv"
            path.write_bytes(b"")
            self.assertIsNotNone(probe_csv(path, 4096).error)


class TestClassify(unittest.TestCase):
    def test_pol_is_not_a_purchase_order(self) -> None:
        keys = classify.detect_key_fields(["Container #", "SKU", "POL", "POD", "ETA"])
        self.assertNotIn("po_number", keys)
        self.assertEqual(keys["port_of_loading"], "POL")
        self.assertEqual(keys["port_of_discharge"], "POD")

    def test_purchase_order_headers_detected(self) -> None:
        keys = classify.detect_key_fields(["PO #", "Line", "Item", "Committed Date"])
        self.assertEqual(keys["po_number"], "PO #")
        self.assertEqual(keys["sku"], "Item")

    def test_file_url_wrapped_unc_still_scores_risk(self) -> None:
        assessment = classify.assess_dependency_risk(
            locations=[r"file:///\\PGFS01\SC\book.xlsx"], has_vba=False, vba_protected=False,
            hidden_sheet_count=0, row_count_capped=False, header_confidence=1.0, probe_errors=[],
        )
        self.assertGreater(assessment.score, 0)
        self.assertTrue(any("UNC" in f for f in assessment.findings))

    def test_downloads_path_outranks_sharepoint(self) -> None:
        downloads = classify.assess_dependency_risk(
            locations=[r"C:\Users\X\Downloads\a.csv"], has_vba=False, vba_protected=False,
            hidden_sheet_count=0, row_count_capped=False, header_confidence=1.0, probe_errors=[])
        sharepoint = classify.assess_dependency_risk(
            locations=["https://x.sharepoint.com/sites/y"], has_vba=False, vba_protected=False,
            hidden_sheet_count=0, row_count_capped=False, header_confidence=1.0, probe_errors=[])
        self.assertGreater(downloads.score, sharepoint.score)

    def test_derived_copy_versus_true_source(self) -> None:
        derived = classify.classify_role(
            query_external_sources=2, query_count=3, external_link_count=0, connection_count=1,
            formula_count=0, total_rows=100, has_vba=False, is_csv=False,
            heavy_formula_count=500, heavy_formulas_per_row=1.0)
        source = classify.classify_role(
            query_external_sources=0, query_count=0, external_link_count=0, connection_count=0,
            formula_count=3, total_rows=100, has_vba=False, is_csv=False,
            heavy_formula_count=500, heavy_formulas_per_row=1.0)
        self.assertEqual(derived.role, classify.ROLE_DERIVED_COPY)
        self.assertEqual(source.role, classify.ROLE_TRUE_SOURCE)

    def test_csv_is_always_a_true_source(self) -> None:
        verdict = classify.classify_role(
            query_external_sources=0, query_count=0, external_link_count=0, connection_count=1,
            formula_count=0, total_rows=10, has_vba=False, is_csv=True,
            heavy_formula_count=500, heavy_formulas_per_row=1.0)
        self.assertEqual(verdict.role, classify.ROLE_TRUE_SOURCE)

    def test_domain_from_collapsed_filename(self) -> None:
        domains = {"open_po": ["items on purchase order", "purchase order"], "supplier": ["vendor"]}
        domain, _hits = classify.classify_domain(
            "NetSuite Exports/ItemsOnPurchaseOrder.csv", ["PO #", "Vendor"], domains)
        self.assertEqual(domain, "open_po")

    def test_same_file_sheets_are_not_duplicate_truth(self) -> None:
        signature = classify.header_signature(["SKU", "Qty", "ETA", "POL", "POD"])
        pairs = classify.find_header_overlaps(
            [("A.xlsx", "S1", "in_transit", 10, signature, {}),
             ("A.xlsx", "S2", "in_transit", 20, signature, {})], 0.7, 4)
        self.assertEqual(pairs, [])

    def test_cross_file_duplicate_truth_detected_with_row_delta(self) -> None:
        left = classify.header_signature(["SKU", "Qty", "ETA", "POL", "POD"])
        right = classify.header_signature(["SKU", "Qty", "ETA", "POL", "POD", "Days Late"])
        pairs = classify.find_header_overlaps(
            [("A.xlsx", "S1", "in_transit", 141, left, {}),
             ("B.xlsx", "S1", "in_transit", 118, right, {})], 0.7, 4)
        self.assertEqual(len(pairs), 1)
        self.assertEqual(pairs[0].row_delta, 23)
        self.assertTrue(pairs[0].is_row_count_conflict)


class TestConfiguration(unittest.TestCase):
    def test_seeds_and_thresholds_load(self) -> None:
        config: Config = load_config()
        self.assertGreater(len(config.discovery.seeds), 10)
        self.assertEqual(config.thresholds.overstock_months_supply, 8.0)
        self.assertEqual(config.thresholds.abc_basis, "revenue")
        self.assertEqual(config.thresholds.worst_flag_hierarchy[0], "CAPITAL TRAP")

    def test_missing_config_raises(self) -> None:
        with self.assertRaises(ConfigError):
            load_config(Path("/nonexistent/config.yaml"))

    def test_unknown_path_key_raises(self) -> None:
        with self.assertRaises(ConfigError):
            load_config().path("not_a_real_key")


class TestEndToEndSweep(unittest.TestCase):
    """The whole Phase 0 run against a simulated estate."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.tmp = tempfile.mkdtemp()
        cls.estate = build(Path(cls.tmp) / "estate")
        base: Config = load_config()
        cls.config = replace(
            base,
            discovery=replace(base.discovery, roots=[str(cls.estate)], extra_roots=[],
                              auto_detect_onedrive=False),
            repo_root=Path(cls.tmp) / "out",
        )
        cls.result = run_discovery(cls.config, write_queries=True)

    @classmethod
    def tearDownClass(cls) -> None:
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def _by_name(self, needle: str):
        return next(s for s in self.result.sources if needle in s.relative_path)

    def test_skip_rules_exclude_archive_backup_and_lock_files(self) -> None:
        paths: List[str] = [s.relative_path for s in self.result.sources]
        self.assertEqual(len(paths), 7)
        for excluded in ("Archive", "Backup", "~$"):
            self.assertFalse([p for p in paths if excluded in p], f"{excluded} was not skipped")

    def test_no_probe_failures(self) -> None:
        failures = [(s.relative_path, s.errors) for s in self.result.sources
                    if s.probe_status == "failed"]
        self.assertEqual(failures, [])

    def test_power_query_stack_extracted_and_written(self) -> None:
        record = self._by_name("In Transit.xlsx")
        self.assertEqual(record.power_query["query_count"], 3)
        self.assertEqual(record.refresh_mechanism.startswith("power_query"), True)
        written: List[str] = [Path(p).name for p in self.result.query_files_written]
        self.assertIn("_Section1.m", written)
        self.assertIn("Items_on_Purchase_Order.m", written)

    def test_fragile_download_path_is_critical(self) -> None:
        record = self._by_name("In Transit.xlsx")
        self.assertEqual(record.risk_band, "critical")
        self.assertTrue(any("Downloads" in f for f in record.risk_findings))

    def test_seed_domain_overrides_keyword_guess(self) -> None:
        self.assertEqual(self._by_name("ItemsOnPurchaseOrder.csv").domain, "open_po")

    def test_key_fields_come_from_the_primary_sheet_only(self) -> None:
        record = self._by_name("Supply_Chain_Update_Meeting_Workbook.xlsm")
        self.assertEqual(record.primary_sheet, "Inventory Health")
        # "value" on the veryHidden _lookup sheet must not leak into the grain.
        self.assertNotIn("currency", record.key_fields)

    def test_header_row_four_detected_under_title_block(self) -> None:
        record = self._by_name("Supply_Chain_Update_Meeting_Workbook.xlsm")
        sheet = next(s for s in record.sheets if s.name == "Inventory Health")
        self.assertEqual(sheet.header_row, 4)
        self.assertIn("Worst Flag", sheet.headers)

    def test_xlsb_is_not_a_blind_spot(self) -> None:
        record = self._by_name("Overstock_Static_Inventory.xlsb")
        self.assertEqual(record.primary_sheet, "Overstock")
        self.assertIn("Months Supply", record.sheets[0].headers)

    def test_protected_vba_is_reported(self) -> None:
        record = self._by_name("Supply_Chain_Update_Meeting_Workbook.xlsm")
        self.assertTrue(record.vba["protected"])
        self.assertIn("mod_Refresh", record.vba["modules"])

    def test_duplicate_truth_pair_found_across_files(self) -> None:
        self.assertEqual(len(self.result.overlaps), 1)
        self.assertTrue(self.result.overlaps[0]["is_row_count_conflict"])

    def test_missing_seeds_are_named_not_assumed(self) -> None:
        missing = [s["id"] for s in self.result.seed_status if not s["found"]]
        self.assertIn("revo_production", missing)
        self.assertIn("master_reference", missing)   # only present under Backup/, which is skipped

    def test_manifest_is_valid_json_and_traces_every_source(self) -> None:
        out: Path = Path(self.tmp) / "manifest_out"
        payload = json.loads(write_manifest(self.result, out).read_text(encoding="utf-8"))
        self.assertEqual(payload["manifest_version"], 1)
        self.assertEqual(len(payload["sources"]), 7)
        for source in payload["sources"]:
            for required in ("path", "modified_at", "sample_hash", "domain", "role", "probe_status"):
                self.assertIn(required, source)

    def test_discovery_md_leads_with_the_conclusion(self) -> None:
        text: str = render_discovery_md(self.result)
        self.assertLess(text.index("## Executive summary"), text.index("## Sources, ranked worst-first"))
        self.assertIn("Duplicate truth", text)
        self.assertIn("what breaks if a path moves", text)

    def test_unreachable_root_is_reported_not_silently_empty(self) -> None:
        config = replace(
            self.config,
            discovery=replace(self.config.discovery, roots=["/definitely/not/here"]),
        )
        result = run_discovery(config, write_queries=False)
        self.assertEqual(result.sources, [])
        self.assertFalse(result.roots[0]["reachable"])
        self.assertIn("does not exist", result.roots[0]["note"])

    def test_rerun_is_idempotent(self) -> None:
        again = run_discovery(self.config, write_queries=True)
        self.assertEqual(len(again.sources), len(self.result.sources))
        self.assertEqual(again.query_diffs, [])   # nothing changed, so no query diffs


class TestQueryDiffDetection(unittest.TestCase):
    def test_changed_m_source_is_flagged_on_the_next_run(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            estate: Path = Path(tmp) / "estate"
            estate.mkdir()
            path: Path = estate / "Pipeline.xlsx"
            write_xlsx(path, [("S", [["SKU", "Qty"], ["A", 1]], "visible")],
                       mashup_section='section Section1;\n\nshared Q = let X = 1 in X;\n')

            base: Config = load_config()
            config = replace(
                base,
                discovery=replace(base.discovery, roots=[str(estate)], extra_roots=[],
                                  auto_detect_onedrive=False),
                repo_root=Path(tmp) / "out",
            )
            first = run_discovery(config, write_queries=True)
            self.assertEqual(first.query_diffs, [])

            # Edit the query, as someone would in the Power Query editor.
            write_xlsx(path, [("S", [["SKU", "Qty"], ["A", 1]], "visible")],
                       mashup_section='section Section1;\n\nshared Q = let X = 1, Y = 2 in Y;\n')
            second = run_discovery(config, write_queries=True)
            self.assertEqual(len(second.query_diffs), 1)
            self.assertEqual(second.query_diffs[0]["query"], "Q")


if __name__ == "__main__":
    unittest.main(verbosity=2)
