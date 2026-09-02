"""Canonical entity builders — one per upstream export.

Each builder turns a validated raw frame into a canonical entity at a declared
grain, applying the agreed rules from ``sc/extract/normalize.py``. The point is
that the transformation exists once. Where the estate had three queries deriving
a SKU three ways, there is one ``canonical_sku``.

Two conventions that matter:

* **Derived quantities are computed, never read.** ``open_po`` recomputes
  ``qty_remaining`` even though the export supplies its own column, then
  compares the two and records a conflict row where they disagree. Sources
  differ on whether their "remaining" nets cancellations; the difference is
  data, not something to pick a side on silently.
* **Lineage on every row.** ``source_slug``, ``source_origin``, ``extracted_at``
  and ``data_as_of`` are attached by the pipeline, so any number traces back.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Sequence, Tuple

import pandas as pd

from ..configuration import CanonicalRules, Config
from . import gates
from .gates import GateResult
from .normalize import attach_location_attributes, canonical_sku


@dataclass
class EntityBuild:
    """A built entity plus the gates that ran on it."""

    entity: str
    grain: List[str]
    frame: pd.DataFrame
    results: List[GateResult] = field(default_factory=list)
    conflicts: pd.DataFrame = field(default_factory=pd.DataFrame)
    notes: List[str] = field(default_factory=list)

    @property
    def blocked(self) -> bool:
        return any(result.blocks_publish for result in self.results)

    @property
    def rows(self) -> int:
        return int(len(self.frame))


def build_inventory_onhand(raw: pd.DataFrame, config: Config) -> EntityBuild:
    """``inventory_onhand`` from the NetSuite inventory export (id 2600947).

    Grain: ``sku x location_id``. Both ``qty_on_hand`` and ``qty_available`` are
    emitted under unambiguous names — the estate labelled ``List.Sum([On Hand])``
    and ``List.Sum([Available])`` identically as "Available", which is why two
    workbooks disagreed.
    """
    rules: CanonicalRules = config.canonical
    frame: pd.DataFrame = raw.copy()
    frame["sku"] = canonical_sku(frame["Name"], rules)
    frame["location_id"] = frame["Inventory Location"].astype("string").str.strip()

    grouped: pd.DataFrame = (
        frame.groupby(["sku", "location_id"], dropna=False, observed=True)
        .agg(
            qty_on_hand=("On Hand", "sum"),
            qty_available=("Available", "sum"),
            qty_on_sales_order=("On Sales Order", "sum"),
            qty_on_purchase_order=("On Purchase Order", "sum"),
            safety_stock_level=("Safety Stock Level", "max"),
            reorder_point=("Location Reorder Point", "max"),
            lead_time_days=("Lead Time", "max"),
            average_cost=("Average Cost", "mean"),
            product_category=("Product Category", "first"),
            region_of_origin=("Region of Origin", "first"),
            source_rows=("Name", "size"),
        )
        .reset_index()
    )
    grouped = attach_location_attributes(grouped, "location_id", rules)

    # Ownership is derived from the location's supply side. Whether the four-way
    # split is right is decision 4 in SCHEMA.md, so the basis is recorded rather
    # than being baked in silently.
    grouped["ownership"] = grouped["supply_side"].map(
        {"domestic": "predator_paid", "asia": "supplier_held_unpaid"}
    ).astype("string")
    grouped["ownership_basis"] = "derived from canonical.locations.supply_side"

    results: List[GateResult] = [
        gates.gate_min_rows(grouped, "inventory_onhand", int(config.extract.gates.get("min_rows", 1))),
        gates.gate_no_null_keys(grouped, "inventory_onhand", "sku",
                                config.extract.gates.get("max_null_key_ratio", 0.0)),
        gates.gate_unique_grain(grouped, "inventory_onhand", ["sku", "location_id"]),
        gates.gate_known_locations(grouped, "inventory_onhand", "location_id", "regional_id",
                                   config.extract.gates.get("max_unknown_location_ratio", 0.02)),
        gates.gate_non_negative(grouped, "inventory_onhand", "qty_on_hand"),
    ]

    build: EntityBuild = EntityBuild(
        entity="inventory_onhand", grain=["sku", "location_id"], frame=grouped, results=results
    )
    build.notes.append(
        "qty_on_hand and qty_available are separate columns on purpose: the estate emitted "
        "both under the name 'Available'."
    )
    return build


def build_open_po(raw: pd.DataFrame, config: Config) -> EntityBuild:
    """``open_po`` from the Items-on-Purchase-Order export (id 2600946).

    Grain: ``po_number x po_line``. ``qty_remaining`` is recomputed and checked
    against the export's own column.
    """
    rules: CanonicalRules = config.canonical
    frame: pd.DataFrame = raw.copy()
    frame["sku"] = canonical_sku(frame["Item"], rules)
    frame["po_number"] = frame["Document Number"].astype("string").str.strip()
    frame["po_line"] = frame["Line ID"]
    frame["supplier_name"] = frame["Name"].astype("string").str.strip()
    frame["location_id"] = frame["Inventory Location"].astype("string").str.strip()
    frame["qty_ordered"] = frame["Quantity"]
    frame["qty_received"] = frame["Quantity Fulfilled/Received"].fillna(0)

    # Computed, not read. The export's own "Quantity Remaining" is compared below.
    frame["qty_remaining"] = (frame["qty_ordered"] - frame["qty_received"]).clip(lower=0)

    frame["committed_date"] = frame["Expected Receipt Date"]
    frame["original_pi_date"] = frame["Original PI Date"]
    frame["order_date"] = frame["Date"]
    frame["status"] = frame["Status"].astype("string").str.strip()
    frame["closed"] = frame["Closed"].astype("string").str.strip().str.lower().isin(
        ["yes", "true", "t", "y"]
    )
    frame["amount"] = frame["Amount"]

    conflicts: pd.DataFrame = pd.DataFrame()
    if "Quantity Remaining" in frame.columns:
        source_remaining: "pd.Series[float]" = frame["Quantity Remaining"].fillna(0)
        mismatch: "pd.Series[bool]" = source_remaining != frame["qty_remaining"].fillna(0)
        if bool(mismatch.any()):
            conflicts = pd.DataFrame({
                "entity": "open_po",
                "entity_key": frame.loc[mismatch, "po_number"].astype("string")
                              + ":" + frame.loc[mismatch, "po_line"].astype("string"),
                "exception_type": "source_conflict",
                "severity": "warning",
                "subject": "qty_remaining",
                "computed_value": frame.loc[mismatch, "qty_remaining"],
                "source_value": source_remaining[mismatch],
                "delta": frame.loc[mismatch, "qty_remaining"] - source_remaining[mismatch],
                "message": "ordered minus received disagrees with the export's own "
                           "Quantity Remaining — the export may net cancellations differently",
            })

    columns: List[str] = [
        "po_number", "po_line", "sku", "supplier_name", "location_id",
        "qty_ordered", "qty_received", "qty_remaining", "committed_date",
        "original_pi_date", "order_date", "status", "closed", "amount",
    ]
    tidy: pd.DataFrame = frame[columns].copy()
    tidy = attach_location_attributes(tidy, "location_id", rules)

    results: List[GateResult] = [
        gates.gate_min_rows(tidy, "open_po", int(config.extract.gates.get("min_rows", 1))),
        gates.gate_no_null_keys(tidy, "open_po", "sku",
                                config.extract.gates.get("max_null_key_ratio", 0.0)),
        gates.gate_unique_grain(tidy, "open_po", ["po_number", "po_line"]),
        gates.gate_non_negative(tidy, "open_po", "qty_ordered"),
    ]
    build: EntityBuild = EntityBuild(
        entity="open_po", grain=["po_number", "po_line"], frame=tidy,
        results=results, conflicts=conflicts,
    )
    build.notes.append(
        "qty_remaining is computed as ordered minus received, floored at zero; the export's "
        "own column is compared and any disagreement is written to conflicts."
    )
    return build


def build_allocation(raw: pd.DataFrame, config: Config) -> EntityBuild:
    """``allocation`` from the new-product allocation export (id 2620991).

    The export is wide — one column per channel — so it is unpivoted to
    ``sku x region_channel``. A wide table cannot answer "what is allocated to
    EMEA across all items" without a formula per column.
    """
    rules: CanonicalRules = config.canonical
    frame: pd.DataFrame = raw.copy()
    frame["sku"] = canonical_sku(frame["Name"], rules)

    channel_columns: List[str] = [c for c in rules.channels if c in frame.columns]
    missing: List[str] = [c for c in rules.channels if c not in frame.columns]

    long: pd.DataFrame = frame.melt(
        id_vars=["sku", "Description", "Base Price", "Targeted Launch Date"],
        value_vars=channel_columns,
        var_name="channel_column",
        value_name="qty_allocated",
    )
    long["region_channel"] = long["channel_column"].map(rules.channels).astype("string")
    long["description"] = long["Description"].astype("string")
    long["base_price"] = long["Base Price"]
    long["targeted_launch_date"] = long["Targeted Launch Date"]
    long["np_launch_control"] = True          # this export is new-product only
    long["qty_allocated"] = long["qty_allocated"].fillna(0)

    tidy: pd.DataFrame = long[[
        "sku", "region_channel", "qty_allocated", "np_launch_control",
        "targeted_launch_date", "base_price", "description",
    ]].copy()

    results: List[GateResult] = [
        gates.gate_min_rows(tidy, "allocation", int(config.extract.gates.get("min_rows", 1))),
        gates.gate_no_null_keys(tidy, "allocation", "sku",
                                config.extract.gates.get("max_null_key_ratio", 0.0)),
        gates.gate_unique_grain(tidy, "allocation", ["sku", "region_channel"]),
        gates.gate_non_negative(tidy, "allocation", "qty_allocated"),
    ]
    build: EntityBuild = EntityBuild(
        entity="allocation", grain=["sku", "region_channel"], frame=tidy, results=results
    )
    build.notes.append(
        f"unpivoted {len(channel_columns)} channel column(s) to rows"
        + (f"; configured channels absent from the export: {missing}" if missing else "")
    )
    return build


#: Upstream slug -> builder. A slug absent here is landed raw and not modelled;
#: `python -m sc.cli extract` says so rather than pretending coverage.
BUILDERS: Dict[str, Callable[[pd.DataFrame, Config], EntityBuild]] = {
    "netsuite_2600947": build_inventory_onhand,
    "netsuite_2600946": build_open_po,
    "netsuite_2620991": build_allocation,
}

#: Why the remaining exports are not modelled yet. Each names the blocking decision.
UNMODELLED: Dict[str, str] = {
    "netsuite_2600949": "demand(open_so) — grain depends on decision 2: three workbooks "
                        "transform this export differently and none is confirmed correct",
    "netsuite_2600950": "demand(actual_shipped) — the only query reading it declares 3 of its "
                        "columns, so the recovered schema is partial",
    "netsuite_2606291": "demand(forecast) — 50 columns of period buckets needing an unpivot; "
                        "blocked on decision 8, the forecast horizon",
    "netsuite_2613629": "item cost — no declared grain yet; one query, no confirmed key",
}
