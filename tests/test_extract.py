"""Tests for the canonical extraction layer.

Every case here corresponds to something the real estate got wrong, found by
`sc.cli reconcile`. These are the regression tests for the consolidation.
"""

from __future__ import annotations

import shutil
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from typing import List

try:
    import pandas as pd

    HAS_PANDAS: bool = True
except ImportError:
    HAS_PANDAS = False

from sc.configuration import Config, load_config


@unittest.skipUnless(HAS_PANDAS, "pandas not installed")
class TestCanonicalSku(unittest.TestCase):
    """The rule that three queries applied three ways."""

    def setUp(self) -> None:
        self.rules = load_config().canonical

    def _sku(self, values: List[object]) -> List[object]:
        from sc.extract.normalize import canonical_sku

        return canonical_sku(pd.Series(values), self.rules).tolist()

    def test_matrix_item_takes_the_child_and_trims(self) -> None:
        # Missing Text.Trim in 2 of 3 queries is what broke the joins.
        self.assertEqual(self._sku(["PREDATOR : PRED-BK-01 "]), ["PRED-BK-01"])

    def test_case_variants_collapse_to_one_key(self) -> None:
        result = self._sku(["PREDATOR : PRED-BK-01", "predator : pred-bk-01"])
        self.assertEqual(result[0], result[1])

    def test_plain_item_without_delimiter_is_kept(self) -> None:
        self.assertEqual(self._sku(["  revo-12.4 "]), ["REVO-12.4"])

    def test_last_delimiter_wins(self) -> None:
        self.assertEqual(self._sku(["A : B : C-9"]), ["C-9"])

    def test_empty_becomes_missing_not_an_empty_string(self) -> None:
        result = self._sku(["   ", "PARENT : ", None])
        self.assertTrue(all(pd.isna(value) for value in result))

    def test_internal_whitespace_collapses(self) -> None:
        self.assertEqual(self._sku(["PRED   BK    01"]), ["PRED BK 01"])


@unittest.skipUnless(HAS_PANDAS, "pandas not installed")
class TestLocationDimension(unittest.TestCase):
    """Three queries each invented a location grouping; there is now one."""

    def setUp(self) -> None:
        self.rules = load_config().canonical

    def test_all_three_groupings_come_from_one_join(self) -> None:
        from sc.extract.normalize import attach_location_attributes

        frame = pd.DataFrame({"location_id": ["HIE", "US B2B", "Jax Tradeshows"]})
        out = attach_location_attributes(frame, "location_id", self.rules)
        self.assertEqual(out.loc[0, "regional_id"], "China")
        self.assertEqual(out.loc[0, "revo_region"], "AP")
        self.assertFalse(bool(out.loc[0, "jax"]))
        self.assertTrue(bool(out.loc[1, "jax"]))
        self.assertEqual(out.loc[2, "regional_id"], "Tradeshow")

    def test_unmapped_location_yields_nulls_and_is_reported(self) -> None:
        from sc.extract.normalize import attach_location_attributes, unknown_locations

        frame = pd.DataFrame({"location_id": ["Narnia"]})
        out = attach_location_attributes(frame, "location_id", self.rules)
        self.assertTrue(pd.isna(out.loc[0, "regional_id"]))
        self.assertEqual(unknown_locations(frame, "location_id", self.rules), ["Narnia"])

    def test_join_cannot_duplicate_rows(self) -> None:
        """validate="many_to_one" guards against a dimension with duplicate keys."""
        from sc.extract.normalize import attach_location_attributes

        frame = pd.DataFrame({"location_id": ["HIE"] * 5})
        self.assertEqual(len(attach_location_attributes(frame, "location_id", self.rules)), 5)


@unittest.skipUnless(HAS_PANDAS, "pandas not installed")
class TestNumericCoercion(unittest.TestCase):
    def test_thousands_and_accounting_negatives(self) -> None:
        from sc.extract.normalize import coerce_numeric

        result = coerce_numeric(pd.Series(["1,204", "(35)", "7", ""]))
        self.assertEqual(result.tolist()[:3], [1204.0, -35.0, 7.0])

    def test_unparseable_becomes_missing_never_zero(self) -> None:
        from sc.extract.normalize import coerce_numeric

        result = coerce_numeric(pd.Series(["x"]))
        self.assertTrue(pd.isna(result.iloc[0]))


@unittest.skipUnless(HAS_PANDAS, "pandas not installed")
class TestReaderContract(unittest.TestCase):
    """Headers are validated by name; quoting is always honoured."""

    def setUp(self) -> None:
        from sc.extract.schema_store import SourceColumn, SourceSchema

        self.rules = load_config().canonical
        self.schema = SourceSchema(
            slug="test_export", key="k", label="l", kind="csv",
            columns=[
                SourceColumn("Item", "string", "type text"),
                SourceColumn("Product Category", "string", "type text"),
                SourceColumn("On Hand", "Int64", "Int64.Type"),
            ],
        )

    def _read(self, csv_text: str, strict: bool = True):
        from sc.extract.fetch import Fetched
        from sc.extract.readers import read_export

        return read_export(Fetched("test_export", csv_text.encode("utf-8-sig"), "memory", "local"),
                           self.schema, self.rules, strict_columns=strict)

    def test_quoted_comma_does_not_shift_columns(self) -> None:
        """The exact failure QuoteStyle.None causes in the live queries."""
        frame, report = self._read(
            'Item,Product Category,On Hand\nPRED-1,"Cues, Playing",1204\n'
        )
        self.assertEqual(report.rows, 1)
        self.assertEqual(frame.loc[0, "Product Category"], "Cues, Playing")
        self.assertEqual(frame.loc[0, "On Hand"], 1204)

    def test_missing_declared_column_fails_loudly_and_names_it(self) -> None:
        from sc.extract.readers import ReadError

        with self.assertRaises(ReadError) as caught:
            self._read("Item,On Hand\nPRED-1,5\n")
        self.assertIn("Product Category", str(caught.exception))

    def test_unexpected_column_is_reported_not_fatal(self) -> None:
        _frame, report = self._read(
            "Item,Product Category,On Hand,Surprise\nPRED-1,Cues,5,x\n"
        )
        self.assertEqual(report.unexpected_columns, ["Surprise"])
        self.assertTrue(report.ok)

    def test_coercion_failures_are_counted(self) -> None:
        _frame, report = self._read("Item,Product Category,On Hand\nPRED-1,Cues,abc\n")
        self.assertEqual(report.coercion_failures.get("On Hand"), 1)


@unittest.skipUnless(HAS_PANDAS, "pandas not installed")
class TestEntityBuildersEndToEnd(unittest.TestCase):
    """The full chain, against exports containing the estate's real hazards."""

    @classmethod
    def setUpClass(cls) -> None:
        from tests.make_export_fixtures import build

        cls.tmp = tempfile.mkdtemp()
        base: Config = load_config()
        drop: Path = Path(cls.tmp) / "drop"
        build(base.repo_root / base.extract.schema_file, drop)
        cls.config = replace(
            base,
            extract=replace(base.extract, drop_dir="drop", fetch_mode="local"),
            repo_root=Path(cls.tmp),
        )

    @classmethod
    def tearDownClass(cls) -> None:
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def _build(self, slug: str):
        from sc.extract.entities import BUILDERS
        from sc.extract.fetch import fetch
        from sc.extract.readers import read_export
        from sc.extract.schema_store import load_schemas

        schema = next(
            s for s in load_schemas(load_config().repo_root / load_config().extract.schema_file)
            if s.slug == slug
        )
        fetched = fetch(schema, self.config.extract, self.config.repo_root)
        raw, _report = read_export(fetched, schema, self.config.canonical)
        return BUILDERS[slug](raw, self.config)

    def test_inventory_separates_on_hand_from_available(self) -> None:
        """The blocking conflict: both were emitted as "Available"."""
        build = self._build("netsuite_2600947")
        self.assertIn("qty_on_hand", build.frame.columns)
        self.assertIn("qty_available", build.frame.columns)
        row = build.frame[(build.frame["sku"] == "PRED-BK-01")
                          & (build.frame["location_id"] == "US B2B")].iloc[0]
        self.assertEqual(int(row["qty_on_hand"]), 1204)
        self.assertEqual(int(row["qty_available"]), 1100)

    def test_inventory_grain_is_sku_by_location(self) -> None:
        build = self._build("netsuite_2600947")
        self.assertEqual(build.grain, ["sku", "location_id"])
        self.assertFalse(build.frame.duplicated(subset=build.grain).any())

    def test_null_key_blocks_the_publish(self) -> None:
        build = self._build("netsuite_2600947")
        self.assertTrue(build.blocked)
        gate = next(r for r in build.results if r.gate == "no_null_keys")
        self.assertFalse(gate.passed)
        self.assertTrue(gate.blocks_publish)

    def test_unmapped_location_blocks_the_publish(self) -> None:
        build = self._build("netsuite_2600947")
        gate = next(r for r in build.results if r.gate == "known_locations")
        self.assertFalse(gate.passed)
        self.assertIn("Narnia", gate.detail)

    def test_negative_on_hand_warns_but_does_not_block(self) -> None:
        build = self._build("netsuite_2600947")
        gate = next(r for r in build.results if r.gate == "non_negative")
        self.assertFalse(gate.passed)
        self.assertFalse(gate.blocks_publish)

    def test_open_po_recomputes_remaining_and_records_the_conflict(self) -> None:
        build = self._build("netsuite_2600946")
        line = build.frame[(build.frame["po_number"] == "PO1001")
                           & (build.frame["po_line"] == 2)].iloc[0]
        self.assertEqual(int(line["qty_remaining"]), 200)      # computed, not the export's 150
        self.assertFalse(build.conflicts.empty)
        conflict = build.conflicts.iloc[0]
        self.assertEqual(int(conflict["computed_value"]), 200)
        self.assertEqual(int(conflict["source_value"]), 150)
        self.assertEqual(int(conflict["delta"]), 50)

    def test_open_po_remaining_is_floored_at_zero(self) -> None:
        build = self._build("netsuite_2600946")
        self.assertTrue((build.frame["qty_remaining"].fillna(0) >= 0).all())

    def test_allocation_unpivots_every_channel(self) -> None:
        build = self._build("netsuite_2620991")
        self.assertEqual(build.grain, ["sku", "region_channel"])
        self.assertEqual(len(build.frame), 2 * len(self.config.canonical.channels))
        self.assertEqual(
            set(build.frame["region_channel"].dropna().unique()),
            set(self.config.canonical.channels.values()),
        )

    def test_allocation_maps_asia_to_apac(self) -> None:
        build = self._build("netsuite_2620991")
        self.assertIn("apac", set(build.frame["region_channel"]))
        self.assertNotIn("ASIA", set(build.frame["region_channel"]))


@unittest.skipUnless(HAS_PANDAS, "pandas not installed")
class TestWarehouse(unittest.TestCase):
    def test_lineage_is_stamped_and_vintage_stays_unknown(self) -> None:
        from datetime import datetime, timezone

        from sc.extract.warehouse import attach_lineage

        frame = pd.DataFrame({"sku": ["A"]})
        out = attach_lineage(frame, source_slug="netsuite_1", source_origin="/x.csv",
                             extracted_at=datetime(2026, 9, 2, tzinfo=timezone.utc))
        self.assertEqual(out.loc[0, "source_slug"], "netsuite_1")
        # data_as_of stays null: the export does not state its own vintage.
        self.assertTrue(pd.isna(out.loc[0, "data_as_of"]))

    def test_round_trip_through_parquet_and_sqlite(self) -> None:
        import sqlite3

        from sc.extract.warehouse import write_sqlite, write_table

        with tempfile.TemporaryDirectory() as tmp:
            frame = pd.DataFrame({"sku": ["A", "B"], "qty_on_hand": [1, 2],
                                  "extracted_at": pd.to_datetime(["2026-09-02", "2026-09-02"])})
            report = write_table(frame, "inventory_onhand", Path(tmp))
            self.assertEqual(report.rows, 2)
            self.assertEqual(len(pd.read_parquet(report.parquet_path)), 2)

            db = write_sqlite({"inventory_onhand": frame}, Path(tmp))
            with sqlite3.connect(db) as connection:
                total = connection.execute("select sum(qty_on_hand) from inventory_onhand").fetchone()
            self.assertEqual(total[0], 3)

    def test_writes_are_atomic_leaving_no_staging_files(self) -> None:
        from sc.extract.warehouse import write_table

        with tempfile.TemporaryDirectory() as tmp:
            write_table(pd.DataFrame({"a": [1]}), "t", Path(tmp))
            self.assertEqual([p.name for p in Path(tmp).glob(".*tmp")], [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
