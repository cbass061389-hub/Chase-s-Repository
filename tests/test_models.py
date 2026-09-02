"""Tests for the proposed canonical model.

Skipped when pydantic is absent: the Phase 0 discovery engine must run with
nothing but PyYAML installed, so its test suite cannot hard-depend on pydantic.
"""

from __future__ import annotations

import unittest
from datetime import date, datetime, timezone
from decimal import Decimal

try:
    import pydantic  # noqa: F401

    HAS_PYDANTIC: bool = True
except ImportError:
    HAS_PYDANTIC = False

LINEAGE = dict(
    source_file="In Transit.xlsx",
    source_sheet="InTransit",
    source_id="f2ee6bb9320c",
    extracted_at=datetime(2026, 9, 2, 12, 0, tzinfo=timezone.utc),
    data_as_of=date(2026, 9, 1),
)


@unittest.skipUnless(HAS_PYDANTIC, "pydantic not installed")
class TestCanonicalRules(unittest.TestCase):
    def test_rule1_sku_is_normalized(self) -> None:
        from sc.models import OpenPo

        po = OpenPo(po_number="PO1", po_line=1, sku="  pred-bk-01 ", uom="ea",
                    qty_ordered=Decimal(10), **LINEAGE)
        self.assertEqual(po.sku, "PRED-BK-01")

    def test_rule1_empty_sku_rejected(self) -> None:
        from sc.models import OpenPo

        with self.assertRaises(Exception):
            OpenPo(po_number="PO1", po_line=1, sku="   ", uom="ea",
                   qty_ordered=Decimal(10), **LINEAGE)

    def test_rule2_lineage_is_not_optional(self) -> None:
        from sc.models import OpenPo

        with self.assertRaises(Exception):
            OpenPo(po_number="PO1", po_line=1, sku="A", uom="ea", qty_ordered=Decimal(1))

    def test_rule3_quantity_requires_uom(self) -> None:
        from sc.models import InventoryOnHand, Ownership

        with self.assertRaises(Exception):
            InventoryOnHand(sku="A", location_id="JAX", ownership=Ownership.PREDATOR_PAID,
                            snapshot_date=date(2026, 9, 1), qty_on_hand=Decimal(5), uom="", **LINEAGE)

    def test_rule3_value_requires_currency(self) -> None:
        from sc.models import InventoryOnHand, Ownership

        with self.assertRaises(Exception):
            InventoryOnHand(sku="A", location_id="JAX", ownership=Ownership.PREDATOR_PAID,
                            snapshot_date=date(2026, 9, 1), qty_on_hand=Decimal(5), uom="EA",
                            extended_value=Decimal(100), **LINEAGE)

    def test_rule3_non_usd_requires_fx_rate(self) -> None:
        from sc.models import MoneyDeclared

        with self.assertRaises(Exception):
            MoneyDeclared(currency="EUR")
        self.assertIsNotNone(MoneyDeclared(currency="USD"))

    def test_rule4_exception_holds_both_sides(self) -> None:
        from sc.models import Exception_, ExceptionSeverity, ExceptionType

        conflict = Exception_(
            exception_id="e1", exception_type=ExceptionType.SOURCE_CONFLICT,
            severity=ExceptionSeverity.BLOCKING, entity="in_transit_line",
            entity_key="MSCU1|PRED-BK-01", message="row count disagreement",
            left_value="141", right_value="118", delta=Decimal(23),
            delta_dollars=Decimal(41000), detected_at=datetime.now(timezone.utc),
        )
        self.assertTrue(conflict.blocks_publish)
        self.assertNotIn("winner", conflict.model_dump())

    def test_derived_quantities_are_computed_not_read(self) -> None:
        from sc.models import OpenPo

        po = OpenPo(po_number="PO1", po_line=1, sku="A", uom="EA",
                    qty_ordered=Decimal(500), qty_received=Decimal(600), **LINEAGE)
        self.assertEqual(po.qty_remaining, Decimal(0))   # floored, never negative

    def test_in_transit_requires_identity_and_an_eta(self) -> None:
        from sc.models import InTransitShipment

        with self.assertRaises(Exception):
            InTransitShipment(shipment_id="S1", container_number="MSCU1", **LINEAGE)
        with self.assertRaises(Exception):
            InTransitShipment(shipment_id="S2", eta=date(2026, 10, 1), **LINEAGE)
        self.assertIsNotNone(
            InTransitShipment(shipment_id="S3", container_number="MSCU1",
                              eta=date(2026, 10, 1), **LINEAGE)
        )

    def test_unknown_source_column_is_an_error(self) -> None:
        from sc.models import Item

        with self.assertRaises(Exception):
            Item(sku="A", base_uom="EA", surprise_new_column="?", **LINEAGE)

    def test_worst_flag_hierarchy_matches_config(self) -> None:
        from sc.configuration import load_config
        from sc.models import WorstFlag

        configured = load_config().thresholds.worst_flag_hierarchy
        self.assertEqual([flag.value for flag in WorstFlag], configured)


if __name__ == "__main__":
    unittest.main(verbosity=2)
