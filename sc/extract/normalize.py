"""The canonical transforms. One rule per concept, defined once.

Every function here replaced a per-query variant that discovery found. The
reconciliation report names the variants; this module is the resolution. If a
rule needs to change, it changes here and every consumer follows — which is the
entire point of the exercise.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

import pandas as pd

from ..configuration import CanonicalRules


def canonical_sku(values: "pd.Series[Any]", rules: CanonicalRules) -> "pd.Series[str]":
    """Derive the canonical SKU from a NetSuite item name, vectorized.

    NetSuite writes matrix items as ``PARENT : CHILD``. Three queries derived a
    SKU from that and only one trimmed the result, so the same item yielded two
    key strings and rows that should have joined did not.

    Rules come from ``canonical.sku`` in config.yaml, so the behaviour is
    auditable without reading this code.
    """
    delimiter: str = str(rules.sku.get("take_after_delimiter") or "")
    text: "pd.Series[str]" = values.astype("string")

    if delimiter:
        # Take the segment after the LAST delimiter; a parent name may contain one.
        tail: "pd.Series[str]" = text.str.rsplit(delimiter, n=1).str[-1]
        text = tail.where(text.str.contains(delimiter, regex=False, na=False), text)

    if rules.sku.get("trim", True):
        text = text.str.strip()
    if rules.sku.get("collapse_internal_whitespace", True):
        text = text.str.replace(r"\s+", " ", regex=True)
    if rules.sku.get("uppercase", False):
        text = text.str.upper()

    # An empty string is not a key. It becomes missing, and the gates reject it.
    return text.replace("", pd.NA)


def attach_location_attributes(
    frame: pd.DataFrame,
    location_column: str,
    rules: CanonicalRules,
    attributes: Optional[List[str]] = None,
) -> pd.DataFrame:
    """Join the location dimension onto *frame*.

    Three queries each built their own location grouping over the same set of
    locations — a regional id, a JAX flag, a China/HIE grouping. They were never
    in conflict, just duplicated, so they are columns of one dimension here.

    An unmapped location yields nulls rather than being dropped or guessed at;
    ``gates.unknown_locations`` decides whether that blocks a publish.
    """
    wanted: List[str] = attributes or ["regional_id", "revo_region", "jax", "supply_side"]
    dimension: pd.DataFrame = location_dimension(rules)[["location", *wanted]]

    keys: "pd.Series[str]" = frame[location_column].astype("string").str.strip()
    merged: pd.DataFrame = frame.assign(**{"__location_key": keys}).merge(
        dimension.rename(columns={"location": "__location_key"}),
        on="__location_key",
        how="left",
        validate="many_to_one",
    )
    return merged.drop(columns="__location_key")


def location_dimension(rules: CanonicalRules) -> pd.DataFrame:
    """The location dimension as a frame, built from config."""
    records: List[Dict[str, Any]] = [
        {"location": name, **attributes} for name, attributes in rules.locations.items()
    ]
    frame: pd.DataFrame = pd.DataFrame.from_records(records)
    for column in ("regional_id", "revo_region", "supply_side"):
        if column in frame.columns:
            frame[column] = frame[column].astype("string")
    if "jax" in frame.columns:
        frame["jax"] = frame["jax"].astype("boolean")
    return frame


def unknown_locations(frame: pd.DataFrame, location_column: str, rules: CanonicalRules) -> List[str]:
    """Locations present in the data but absent from the configured dimension."""
    seen = set(frame[location_column].astype("string").str.strip().dropna().unique())
    return sorted(seen - set(rules.known_locations))


def coerce_numeric(values: "pd.Series[Any]", thousands: str = ",") -> "pd.Series[Any]":
    """Parse a numeric column that may carry thousands separators.

    Values that do not parse become missing, never zero — a failed parse and a
    genuine zero must not look the same downstream.
    """
    if pd.api.types.is_numeric_dtype(values):
        return values
    cleaned: "pd.Series[str]" = values.astype("string").str.strip()
    if thousands:
        cleaned = cleaned.str.replace(thousands, "", regex=False)
    cleaned = cleaned.str.replace(r"^\((.*)\)$", r"-\1", regex=True)   # (1,234) -> -1234
    return pd.to_numeric(cleaned, errors="coerce")
