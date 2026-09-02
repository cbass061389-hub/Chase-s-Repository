"""Canonical entities — the executable form of SCHEMA.md.

**Proposed, pending sign-off.** Nothing in Phase 2 imports this until the model
is agreed; it exists now so the proposal is reviewable as code rather than only
as prose, and so the four rules are enforced by the type system instead of by
convention.

The rules, as enforced here:

1. ``sku`` is normalized on every model that carries one, so the join key cannot
   drift on whitespace or case.
2. ``LineageMixin`` is mandatory on every fact model — there is no way to
   construct a fact row without saying where it came from.
3. A quantity requires its UOM and money requires its currency, enforced by
   validators rather than documented and hoped for.
4. Conflicts are represented by :class:`Exception_`, which holds *both* values
   and *both* lineages. There is no field for "the right answer".
"""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

# --------------------------------------------------------------------------
# Enumerations — closed sets, so a typo in a source becomes a loud failure
# --------------------------------------------------------------------------

class ItemStatus(str, Enum):
    ACTIVE = "active"
    PHASE_OUT = "phase_out"
    DISCONTINUED = "discontinued"
    NEW_PRODUCT = "new_product"


class AbcClass(str, Enum):
    A = "A"
    B = "B"
    C = "C"


class Ownership(str, Enum):
    """Who owns and who has paid for the stock. A dimension, not a flag."""

    PREDATOR_PAID = "predator_paid"
    SUPPLIER_HELD_UNPAID = "supplier_held_unpaid"
    CONSIGNMENT = "consignment"
    IN_BOND = "in_bond"


class TransportMode(str, Enum):
    OCEAN = "ocean"
    AIR = "air"
    TRUCK = "truck"
    COURIER = "courier"


class RegionChannel(str, Enum):
    AMERICAS = "americas"
    EMEA = "emea"
    APAC = "apac"
    B2C = "b2c"
    TRADESHOW = "tradeshow"
    SPONSORSHIP = "sponsorship"
    MARKETING = "marketing"
    HQ_RESERVE = "hq_reserve"


class DsrDecision(str, Enum):
    APPROVED = "approved"
    DENIED = "denied"
    OVERRIDDEN = "overridden"
    PENDING = "pending"


class DemandType(str, Enum):
    FORECAST = "forecast"
    OPEN_SO = "open_so"
    ACTUAL_SHIPPED = "actual_shipped"
    BUDGET = "budget"


class InspectionType(str, Enum):
    PRE_SHIPMENT = "pre_shipment"
    INCOMING = "incoming"
    IN_PROCESS = "in_process"
    FINAL = "final"


class Severity(str, Enum):
    CRITICAL = "critical"
    MAJOR = "major"
    MINOR = "minor"


class WorstFlag(str, Enum):
    """Evaluated top-down; first match wins. Order is the hierarchy."""

    CAPITAL_TRAP = "CAPITAL TRAP"
    MOSTLY_DEAD = "MOSTLY DEAD"
    AT_RISK = "AT RISK"
    OVERSTOCK = "OVERSTOCK"
    LOW_CONFIDENCE = "LOW CONFIDENCE"
    OVER_FORECASTED = "OVER-FORECASTED"
    MONITOR = "MONITOR"


class ExceptionType(str, Enum):
    DUPLICATE_KEY = "duplicate_key"
    ORPHAN_SKU = "orphan_sku"
    NEGATIVE_QTY = "negative_qty"
    STALE_SNAPSHOT = "stale_snapshot"
    MISSING_CONTAINER_OR_ETA = "missing_container_or_eta"
    ALLOCATION_OVER_SUPPLY = "allocation_over_supply"
    UOM_MISMATCH = "uom_mismatch"
    CURRENCY_UNDECLARED = "currency_undeclared"
    SOURCE_CONFLICT = "source_conflict"


class ExceptionSeverity(str, Enum):
    BLOCKING = "blocking"
    WARNING = "warning"
    INFO = "info"


# --------------------------------------------------------------------------
# Base
# --------------------------------------------------------------------------

class ScModel(BaseModel):
    """Strict base: unknown fields are an error, not a shrug.

    A source growing a column must be a visible failure, because an unnoticed
    new column is usually a changed definition of an existing one.
    """

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True, use_enum_values=False)


class LineageMixin(ScModel):
    """Rule 2 — every fact row states where it came from and how old it is."""

    source_file: str
    source_sheet: str
    source_row: Optional[int] = None
    source_id: str = Field(description="stable 12-char id from discovery/manifest.json")
    extracted_at: datetime
    data_as_of: Optional[date] = Field(
        default=None,
        description="vintage of the DATA, not the file mtime; null means the source did not state it",
    )


def _normalize_sku(value: str) -> str:
    return str(value).strip().upper()


class SkuKeyed(LineageMixin):
    """Rule 1 — one canonical SKU spelling, normalized at construction."""

    sku: str

    @field_validator("sku")
    @classmethod
    def _sku_normalized(cls, value: str) -> str:
        normalized: str = _normalize_sku(value)
        if not normalized:
            raise ValueError("sku is empty after normalization; the row cannot be joined")
        return normalized


class QuantityDeclared(ScModel):
    """Rule 3 — a quantity is meaningless without its unit."""

    uom: str = Field(min_length=1, description="required; there is no default unit")

    @field_validator("uom")
    @classmethod
    def _uom_present(cls, value: str) -> str:
        if not str(value).strip():
            raise ValueError("uom must be declared, never inferred")
        return str(value).strip().upper()


class MoneyDeclared(ScModel):
    """Rule 3 — money without a currency is a number, not an amount."""

    currency: str = Field(min_length=3, max_length=3, description="ISO 4217")
    fx_rate_to_usd: Optional[Decimal] = None
    fx_as_of: Optional[date] = None

    @model_validator(mode="after")
    def _non_usd_needs_a_rate(self) -> "MoneyDeclared":
        if self.currency.upper() != "USD" and self.fx_rate_to_usd is None:
            raise ValueError(
                f"currency {self.currency} requires fx_rate_to_usd; a USD roll-up cannot be "
                "produced from mixed currencies without one"
            )
        return self


# --------------------------------------------------------------------------
# Entities
# --------------------------------------------------------------------------

class Item(LineageMixin):
    """Grain: sku."""

    sku: str
    description: Optional[str] = None
    family: Optional[str] = None
    product_line: Optional[str] = None
    brand: Optional[str] = None
    status: ItemStatus = ItemStatus.ACTIVE
    abc_class: Optional[AbcClass] = Field(
        default=None, description="by trailing-12-month revenue, never units"
    )
    standard_cost: Optional[Decimal] = None
    cost_currency: Optional[str] = None
    base_uom: str = Field(min_length=1)

    @field_validator("sku")
    @classmethod
    def _normalize(cls, value: str) -> str:
        return _normalize_sku(value)

    @model_validator(mode="after")
    def _cost_needs_currency(self) -> "Item":
        if self.standard_cost is not None and not self.cost_currency:
            raise ValueError(f"{self.sku}: standard_cost present with no cost_currency")
        return self


class ItemCrosswalk(LineageMixin):
    """Grain: supplier_id + supplier_item_code. Date-effective on purpose."""

    supplier_id: str
    supplier_item_code: str
    sku: str
    valid_from: Optional[date] = None
    valid_to: Optional[date] = None

    @field_validator("sku")
    @classmethod
    def _normalize(cls, value: str) -> str:
        return _normalize_sku(value)


class InventoryOnHand(SkuKeyed, QuantityDeclared):
    """Grain: sku + location_id + ownership + snapshot_date."""

    location_id: str
    ownership: Ownership
    snapshot_date: date
    qty_on_hand: Decimal
    qty_available: Optional[Decimal] = None
    qty_committed: Optional[Decimal] = None
    extended_value: Optional[Decimal] = None
    currency: Optional[str] = None

    @model_validator(mode="after")
    def _value_needs_currency(self) -> "InventoryOnHand":
        if self.extended_value is not None and not self.currency:
            raise ValueError(f"{self.sku}@{self.location_id}: extended_value present with no currency")
        return self


class OpenPo(SkuKeyed, QuantityDeclared):
    """Grain: po_number + po_line."""

    po_number: str
    po_line: int
    supplier_id: Optional[str] = None
    qty_ordered: Decimal
    qty_received: Decimal = Decimal(0)
    committed_date: Optional[date] = None
    revised_date: Optional[date] = None
    unit_cost: Optional[Decimal] = None
    cost_currency: Optional[str] = None
    status: Optional[str] = None

    @property
    def qty_remaining(self) -> Decimal:
        """Derived here, never read from a source column.

        Sources disagree on whether their own "remaining" nets cancellations,
        which is exactly why this is computed.
        """
        return max(self.qty_ordered - self.qty_received, Decimal(0))

    def days_late(self, as_of: date) -> Optional[int]:
        """Days past the committed date for an unreceived line. None when N/A."""
        if self.committed_date is None or self.qty_remaining <= 0:
            return None
        return max((as_of - self.committed_date).days, 0)


class InTransitShipment(LineageMixin):
    """Grain: shipment_id. Header-level — ETA lives here, exactly once."""

    shipment_id: str
    container_number: Optional[str] = None
    invoice_number: Optional[str] = None
    bl_number: Optional[str] = None
    supplier_id: Optional[str] = None
    pol: Optional[str] = None
    pod: Optional[str] = None
    etd: Optional[date] = None
    eta: Optional[date] = None
    revised_eta: Optional[date] = None
    mode: Optional[TransportMode] = None
    vessel: Optional[str] = None
    carrier: Optional[str] = None
    status: Optional[str] = None

    @model_validator(mode="after")
    def _must_be_identifiable_and_predictable(self) -> "InTransitShipment":
        if not (self.container_number or self.invoice_number):
            raise ValueError(
                f"shipment {self.shipment_id}: no container number and no invoice number — "
                "the shipment cannot be identified"
            )
        if not (self.eta or self.revised_eta):
            raise ValueError(
                f"shipment {self.shipment_id}: no ETA — the row cannot answer "
                "'what is going to happen?', which is the point of the table"
            )
        return self


class InTransitLine(SkuKeyed, QuantityDeclared):
    """Grain: shipment_id + sku + channel. Channel splits live here, not on the header."""

    shipment_id: str
    channel: RegionChannel
    qty: Decimal
    allocation_ref: Optional[str] = None
    unit_cost: Optional[Decimal] = None
    cost_currency: Optional[str] = None


class Allocation(SkuKeyed, QuantityDeclared):
    """Grain: sku + region_channel + effective_date."""

    region_channel: RegionChannel
    effective_date: date
    qty_allocated: Decimal
    qty_consumed: Decimal = Decimal(0)
    np_launch_control: bool = False
    hq_lock_qty: Optional[Decimal] = None
    lock_expires: Optional[date] = None
    approver: Optional[str] = None

    @property
    def qty_remaining(self) -> Decimal:
        return max(self.qty_allocated - self.qty_consumed, Decimal(0))


class DsrRequest(LineageMixin):
    """Grain: request_id."""

    request_id: str
    requester: str
    region: RegionChannel
    requested_at: datetime
    decision: DsrDecision = DsrDecision.PENDING
    decision_reason: Optional[str] = None
    decided_by: Optional[str] = None
    decided_at: Optional[datetime] = None

    @model_validator(mode="after")
    def _decisions_are_attributable(self) -> "DsrRequest":
        if self.decision is not DsrDecision.PENDING and self.decided_at is None:
            raise ValueError(
                f"DSR {self.request_id}: decision '{self.decision.value}' with no decided_at; "
                "an unattributable decision cannot be reviewed"
            )
        return self


class DsrLine(SkuKeyed, QuantityDeclared):
    """Grain: request_id + sku."""

    request_id: str
    qty_requested: Decimal
    qty_approved: Optional[Decimal] = None
    inventory_snapshot: Dict[str, Any] = Field(
        default_factory=dict,
        description="availability position at the moment of the decision; without it an "
                    "override cannot be judged after the fact",
    )


class Production(SkuKeyed, QuantityDeclared):
    """Grain: production_order_id + sku + stage."""

    production_order_id: str
    stage: str
    supplier_id: Optional[str] = None
    qty_planned: Decimal
    qty_completed: Decimal = Decimal(0)
    qty_b_grade: Decimal = Decimal(0)
    qty_rework: Decimal = Decimal(0)
    qty_scrap: Decimal = Decimal(0)
    scheduled_completion: Optional[date] = None
    actual_completion: Optional[date] = None
    status: Optional[str] = None

    @property
    def schedule_variance_days(self) -> Optional[int]:
        if self.scheduled_completion is None or self.actual_completion is None:
            return None
        return (self.actual_completion - self.scheduled_completion).days


class Demand(SkuKeyed, QuantityDeclared):
    """Grain: sku + period_start + demand_type."""

    demand_type: DemandType
    period_start: date
    period_end: date
    qty: Decimal
    forecast_version: Optional[str] = None
    confidence_score: Optional[float] = Field(default=None, ge=0.0, le=1.0)
    confidence_basis: Optional[str] = Field(
        default=None, description="e.g. '3-month lookback variance'"
    )

    @model_validator(mode="after")
    def _period_ordered(self) -> "Demand":
        if self.period_end < self.period_start:
            raise ValueError(f"{self.sku}: period_end {self.period_end} precedes {self.period_start}")
        return self


class Quality(SkuKeyed, QuantityDeclared):
    """Grain: inspection_id + sku + defect_code."""

    inspection_id: str
    inspection_date: date
    inspection_type: InspectionType
    supplier_id: Optional[str] = None
    qty_inspected: Decimal
    qty_rejected: Decimal = Decimal(0)
    defect_code: Optional[str] = None
    defect_description: Optional[str] = None
    severity: Optional[Severity] = None
    car_number: Optional[str] = None
    car_status: Optional[str] = None
    disposition: Optional[str] = None

    @property
    def reject_rate(self) -> Optional[Decimal]:
        if self.qty_inspected <= 0:
            return None
        return self.qty_rejected / self.qty_inspected


class Supplier(LineageMixin):
    """Grain: supplier_id. P50/P90 lead time, not an average."""

    supplier_id: str
    name: str
    country: Optional[str] = None
    nominal_lead_time_days: Optional[int] = None
    actual_lead_time_p50: Optional[int] = None
    actual_lead_time_p90: Optional[int] = Field(
        default=None, description="the tail is what breaks a launch date"
    )
    open_spend: Optional[Decimal] = None
    spend_currency: Optional[str] = None
    status: Optional[str] = None


class Tariff(LineageMixin):
    """Grain: trade_lane + hts_category + effective_date. Never overwritten."""

    trade_lane: str
    hts_category: str
    effective_date: date
    rate_pct: Decimal
    additional_duty_pct: Decimal = Decimal(0)
    basis: str = Field(default="customs_value", description="fob | cif | customs_value")
    expires_date: Optional[date] = None
    authority: Optional[str] = None
    notes: Optional[str] = None


class Exception_(ScModel):
    """Rule 4 — a conflict, holding both sides. There is no 'winner' field.

    Named with a trailing underscore to avoid shadowing the builtin; the table
    is ``exceptions``.
    """

    exception_id: str
    exception_type: ExceptionType
    severity: ExceptionSeverity
    entity: str
    entity_key: str
    message: str
    left_value: Optional[str] = None
    right_value: Optional[str] = None
    delta: Optional[Decimal] = None
    delta_dollars: Optional[Decimal] = Field(
        default=None, description="ranks the queue; counts do not"
    )
    left_lineage: Optional[Dict[str, Any]] = None
    right_lineage: Optional[Dict[str, Any]] = None
    detected_at: datetime

    @property
    def blocks_publish(self) -> bool:
        return self.severity is ExceptionSeverity.BLOCKING


class SourceVintage(ScModel):
    """Grain: source_id + run_id. Populated from discovery/manifest.json."""

    source_id: str
    run_id: str
    path: str
    file_modified_at: datetime
    data_as_of: Optional[date] = None
    row_count: int
    refresh_status: str
    probe_status: str
    risk_band: str


#: Every fact entity, in the order Phase 2 loads them (dependencies first).
LOAD_ORDER: List[type[ScModel]] = [
    Item,
    ItemCrosswalk,
    Supplier,
    InventoryOnHand,
    OpenPo,
    InTransitShipment,
    InTransitLine,
    Allocation,
    DsrRequest,
    DsrLine,
    Production,
    Demand,
    Quality,
    Tariff,
]
