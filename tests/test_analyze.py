"""Tests for query reconciliation, schema recovery and formula mapping.

The cases are the real defects these tools found in the estate.
"""

from __future__ import annotations

import unittest
from typing import List

from sc.analyze.export_schema import build_export_specs, pandas_dtype_for
from sc.analyze.formula_map import ColumnFormulas, normalize_formula
from sc.analyze.m_ast import (
    call_arguments,
    parse_let_steps,
    profile_query,
    split_top_level,
    strip_one_brace_layer,
)
from sc.analyze.reconcile import SEVERITY_BLOCKING, reconcile

INVENTORY_QUERY_MEETING: str = '''shared #"Current Inventory" = let
    Source = Csv.Document(Web.Contents("https://x/media.nl?id=2600947&h=<REDACTED>"),[Delimiter=",", Columns=24, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Name", type text}, {"On Hand", Int64.Type}, {"Available", Int64.Type}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Name", "Inventory Location"}, {{"Ttl Available", each List.Sum([On Hand]), type nullable number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Grouped Rows",{{"Ttl Available", "InvQty"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([InvQty] <> null)),
    #"Added Custom1" = Table.AddColumn(#"Filtered Rows", "SKU", each if Text.Contains([Item], ":") then Text.Trim(Text.AfterDelimiter([Item], ":")) else Text.Trim([Item]))
in
    #"Added Custom1";'''

INVENTORY_QUERY_HIE: str = '''shared #"Regional Inv Loc" = let
    Source = Csv.Document(Web.Contents("https://x/media.nl?id=2600947&h=<REDACTED>"),[Delimiter=",", Columns=24, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Name", type text}, {"On Hand", Int64.Type}, {"Available", Int64.Type}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Name", "Inventory Location"}, {{"Available", each List.Sum([Available]), type nullable number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "SKU", each if Text.Contains([Name],":") then Text.AfterDelimiter([Name],":") else [Name])
in
    #"Added Custom";'''

LOCATION_QUERY_HIE: str = '''shared Location_Inventory = let
    Source = Csv.Document(Web.Contents("https://x/media.nl?id=2600947&h=<REDACTED>"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Name", type text}, {"On Hand", Int64.Type}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Inventory Location", "Name"}, {{"Available", each List.Sum([On Hand]), type nullable number}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows", each ([Custom] = "JAX"))
in
    #"Filtered Rows";'''


class TestMTokenizer(unittest.TestCase):
    def test_split_respects_brackets_and_strings(self) -> None:
        parts = split_top_level('a = f(1, 2), b = "x, y", c = {1, 2}')
        self.assertEqual(len(parts), 3)

    def test_strip_one_brace_layer_only(self) -> None:
        """lstrip('{') removes every brace and flattens a nested list too far."""
        self.assertEqual(strip_one_brace_layer('{{"a", 1}}'), '{"a", 1}')

    def test_call_arguments_of_nested_call(self) -> None:
        args = call_arguments('Table.Group(prev, {"A","B"}, {{"n", each List.Sum([q])}})', "Table.Group")
        self.assertEqual(len(args or []), 3)

    def test_let_steps_parsed(self) -> None:
        steps = parse_let_steps(INVENTORY_QUERY_MEETING)
        self.assertIn("Grouped Rows", [step.name for step in steps])


class TestQueryProfile(unittest.TestCase):
    def test_measure_lineage_follows_renames(self) -> None:
        """"InvQty" must be traceable back to List.Sum([On Hand])."""
        profile = profile_query("Meeting.xlsm", "Current Inventory", INVENTORY_QUERY_MEETING)
        self.assertEqual(profile.measures, {"InvQty": "List.Sum([On Hand])"})

    def test_grain_is_the_group_keys(self) -> None:
        profile = profile_query("Meeting.xlsm", "Current Inventory", INVENTORY_QUERY_MEETING)
        self.assertEqual(set(profile.group_keys), {"Name", "Inventory Location"})

    def test_parse_options_unquoted(self) -> None:
        profile = profile_query("Meeting.xlsm", "Current Inventory", INVENTORY_QUERY_MEETING)
        self.assertEqual(profile.source.option("Delimiter"), ",")
        self.assertEqual(profile.source.option("QuoteStyle"), "QuoteStyle.None")

    def test_declared_schema_recovered_without_fetching(self) -> None:
        profile = profile_query("Meeting.xlsm", "Current Inventory", INVENTORY_QUERY_MEETING)
        self.assertEqual(profile.declared_schema["On Hand"], "Int64.Type")
        self.assertEqual(profile.declared_schema["Name"], "type text")

    def test_key_derivation_found_under_its_final_name(self) -> None:
        profile = profile_query("Meeting.xlsm", "Current Inventory", INVENTORY_QUERY_MEETING)
        self.assertIn("Text.Trim", profile.key_derivations["SKU"])


class TestReconciliation(unittest.TestCase):
    def _profiles(self):
        left = profile_query("Meeting.xlsm", "Current Inventory", INVENTORY_QUERY_MEETING)
        right = profile_query("HIE.xlsm", "Regional Inv Loc", INVENTORY_QUERY_HIE)
        third = profile_query("HIE.xlsm", "Location_Inventory", LOCATION_QUERY_HIE)
        for profile in (left, right, third):
            profile.upstream_key = "netsuite:media:2600947"
            profile.upstream_label = "NetSuite export 2600947"
        return [left, right, third]

    def test_same_name_different_measure_is_blocking(self) -> None:
        """The real defect: "Available" meant On Hand in one query, Available in another."""
        _groups, findings = reconcile(self._profiles())
        conflicts = [f for f in findings if f.kind == "measure_definition_conflict"]
        self.assertTrue(conflicts)
        self.assertEqual(conflicts[0].severity, SEVERITY_BLOCKING)
        self.assertEqual(
            {conflicts[0].left_value, conflicts[0].right_value},
            {"List.Sum([On Hand])", "List.Sum([Available])"},
        )

    def test_missing_trim_on_a_join_key_is_blocking(self) -> None:
        _groups, findings = reconcile(self._profiles())
        keys = [f for f in findings if f.kind == "key_normalization_conflict"]
        self.assertTrue(keys)
        self.assertEqual(keys[0].severity, SEVERITY_BLOCKING)
        self.assertIn("Text.Trim", keys[0].consequence)

    def test_quotestyle_none_versus_csv_is_blocking(self) -> None:
        _groups, findings = reconcile(self._profiles())
        parse = [f for f in findings if f.kind == "parse_option_conflict"
                 and "QuoteStyle" in f.subject]
        self.assertTrue(parse)
        self.assertTrue(all(f.severity == SEVERITY_BLOCKING for f in parse))

    def test_identical_queries_produce_no_findings(self) -> None:
        left = profile_query("A.xlsm", "Q", INVENTORY_QUERY_MEETING)
        right = profile_query("B.xlsm", "Q", INVENTORY_QUERY_MEETING)
        for profile in (left, right):
            profile.upstream_key = "k"
            profile.upstream_label = "same export"
        _groups, findings = reconcile([left, right])
        self.assertEqual(findings, [])

    def test_a_single_consumer_is_not_reconciled(self) -> None:
        only = profile_query("A.xlsm", "Q", INVENTORY_QUERY_MEETING)
        only.upstream_key, only.upstream_label = "k", "one export"
        groups, findings = reconcile([only])
        self.assertEqual(findings, [])
        self.assertFalse(groups[0].is_forked)


class TestExportSchemaRecovery(unittest.TestCase):
    def test_type_mapping(self) -> None:
        self.assertEqual(pandas_dtype_for("Int64.Type"), "Int64")
        self.assertEqual(pandas_dtype_for("type nullable number"), "Float64")
        self.assertEqual(pandas_dtype_for("type date"), "datetime64[ns]")

    def test_unknown_type_falls_back_to_string(self) -> None:
        """A wrong numeric coercion changes values; a string fails loudly later."""
        self.assertEqual(pandas_dtype_for("type whatever"), "string")

    def test_declarations_merge_and_conflicts_are_kept(self) -> None:
        left = profile_query("A.xlsm", "Q1", INVENTORY_QUERY_MEETING)
        right = profile_query("B.xlsm", "Q2", INVENTORY_QUERY_HIE)
        for profile in (left, right):
            profile.upstream_key = "netsuite:media:2600947"
            profile.upstream_label = "export 2600947"
        specs = build_export_specs([left, right])
        self.assertEqual(len(specs), 1)
        self.assertIn("On Hand", specs[0].column_names)
        self.assertEqual(specs[0].slug, "netsuite_2600947")


class TestFormulaMapping(unittest.TestCase):
    def test_row_numbers_collapse_to_a_pattern(self) -> None:
        self.assertEqual(
            normalize_formula('IF(B5>8,"OVERSTOCK","MONITOR")'),
            normalize_formula('IF(B6>8,"OVERSTOCK","MONITOR")'),
        )

    def test_absolute_references_normalize_too(self) -> None:
        self.assertEqual(normalize_formula("VLOOKUP($A$5,X!$A$1:$B$9,2,FALSE)"),
                         normalize_formula("VLOOKUP($A$6,X!$A$1:$B$9,2,FALSE)"))

    def _column(self, patterns: dict, broken: int = 0) -> ColumnFormulas:
        column = ColumnFormulas(column_index=0, column_letter="A")
        for pattern, count in patterns.items():
            column.patterns[pattern] = count
        column.broken_reference_cells = broken
        return column

    def test_uniform_formula_column_is_not_suspicious(self) -> None:
        self.assertFalse(self._column({"SUM(A{r}:B{r})": 500}).is_suspicious)

    def test_mostly_uniform_column_with_exceptions_is_suspicious(self) -> None:
        column = self._column({"SUM(A{r}:B{r})": 480, "SUM(A{r}:C{r})": 20})
        self.assertTrue(column.is_suspicious)

    def test_varied_text_literals_are_not_a_defect(self) -> None:
        """"No completion data (non-BCP)" is a literal despite its parentheses."""
        column = self._column({
            '"Flag: below 85% complete"': 34,
            '"No completion data (non-BCP)"': 28,
            '"Requested item"': 17,
            '"At or above 85% complete"': 9,
        })
        self.assertTrue(column.is_constant_column)
        self.assertFalse(column.is_suspicious)

    def test_broken_references_are_counted_separately(self) -> None:
        column = self._column({"'Open Order Report'!#REF!": 2909,
                               "'Open Order Report'!A{r}": 6453}, broken=2909)
        self.assertTrue(column.has_broken_references)
        self.assertEqual(column.broken_reference_cells, 2909)

    def test_consistency_label_never_rounds_up_to_a_clean_hundred(self) -> None:
        column = self._column({"A{r}": 4752, "B{r}": 10})
        self.assertEqual(column.consistency_label, "99.8%")


if __name__ == "__main__":
    unittest.main(verbosity=2)
