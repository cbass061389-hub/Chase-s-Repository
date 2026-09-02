"""Turn raw probe output into the judgements DISCOVERY.md has to make.

Four questions per source, all of them answered from evidence in the file
rather than from its name:

1. Which domain does it belong to?
2. Is it a **true source** or a **downstream copy**? (formula density, query
   sources and external links decide this, not the filename)
3. What is its grain and key?
4. How badly does it break when something moves? (dependency risk)

Plus one cross-file question: **do two workbooks claim the same data?**
Answered by comparing normalized header signatures, so it catches a duplicate
even when the two files are named nothing alike.
"""

from __future__ import annotations

import fnmatch
import re
from dataclasses import dataclass, field
from itertools import combinations
from pathlib import PurePath
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

# --------------------------------------------------------------------------
# Key-field detection — supply chain vocabulary, not generic column guessing
# --------------------------------------------------------------------------
KEY_PATTERNS: Dict[str, Tuple[str, ...]] = {
    "sku": ("sku", "item", "item code", "item name", "itemid", "part number", "part no", "model", "upc"),
    "po_number": ("po #", "po#", "po number", "po no", "purchase order", "document number"),
    "po_line": ("line", "line #", "line no", "seq"),
    "container": ("container", "container #", "cntr", "ctnr", "equipment"),
    "port_of_loading": ("pol", "port of loading", "load port", "origin port"),
    "port_of_discharge": ("pod", "port of discharge", "discharge port", "destination port"),
    "invoice": ("invoice", "invoice #", "inv #", "inv no"),
    "location": ("location", "warehouse", "whse", "site", "bin", "subsidiary"),
    "quantity": ("qty", "quantity", "units", "pcs", "on hand", "onhand", "available"),
    "date": ("date", "eta", "etd", "committed", "ship date", "due date", "receipt date", "as of"),
    "region_channel": ("region", "channel", "market", "americas", "emea", "b2c", "customer group"),
    "supplier": ("vendor", "supplier", "manufacturer", "factory"),
    "work_order": ("wo", "wo #", "work order", "production order"),
    "currency": ("currency", "curr", "fx", "usd", "cost", "price", "amount", "value"),
    "uom": ("uom", "unit of measure", "units of measure", "each", "per"),
}

# Locations that are fragile by construction. Ranked worst-first.
FRAGILE_LOCATION_RULES: Tuple[Tuple[str, str, int], ...] = (
    (r"^[A-Za-z]:\\Users\\", "hardcoded per-user local path — breaks on any other machine", 40),
    (r"^\\\\", "UNC network share — breaks off-VPN and on any file-server change", 25),
    (r"(?i)onedrive", "OneDrive path — breaks if the tenant folder is renamed or resynced", 20),
    (r"(?i)sharepoint\.com", "SharePoint URL — survives moves better, but needs auth at refresh", 8),
    (r"(?i)^https?://", "web endpoint — external availability dependency", 8),
    (r"(?i)\\Downloads\\|\\Desktop\\", "reads from Downloads/Desktop — a manual save step is baked in", 45),
)

def normalize_location(location: str) -> str:
    """Strip URL wrappers so a path is matched by what it actually points at.

    Excel writes external links as ``file:///\\\\server\\share\\book.xlsx``; left
    wrapped, the UNC rule never fires and a network dependency scores zero risk.
    """
    text: str = str(location).strip()
    for prefix in ("file:///", "file://", "file:"):
        if text.lower().startswith(prefix):
            text = text[len(prefix):]
            break
    return text.replace("%20", " ")


PATH_MATCH_WEIGHT: int = 3

ROLE_TRUE_SOURCE: str = "true_source"
ROLE_DERIVED_COPY: str = "derived_copy"
ROLE_CALCULATED_OUTPUT: str = "calculated_output"
ROLE_HYBRID: str = "hybrid"
ROLE_UNKNOWN: str = "unknown"


def normalize_header(value: str) -> str:
    """Collapse a header to a comparable token: lowercase, alphanumeric only."""
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def header_signature(headers: Iterable[str]) -> Set[str]:
    """Normalized header set used for cross-file overlap comparison."""
    return {token for token in (normalize_header(h) for h in headers) if token}


def detect_key_fields(headers: Sequence[str]) -> Dict[str, str]:
    """Map each detected key role to the actual header that satisfied it.

    Exact normalized matches beat substring matches, so a column literally named
    "Item" wins over "Item Description" for the ``sku`` role.
    """
    candidates: List[Tuple[str, List[str], str]] = [
        (normalize_header(h), _tokenize(h), h) for h in headers if str(h).strip()
    ]
    found: Dict[str, str] = {}
    for role, patterns in KEY_PATTERNS.items():
        exact_targets: List[str] = [normalize_header(p) for p in patterns]
        exact: Optional[str] = next(
            (original for norm, _tokens, original in candidates if norm in exact_targets), None
        )
        if exact is not None:
            found[role] = exact
            continue
        token_targets: List[List[str]] = [_tokenize(p) for p in patterns]
        partial: Optional[str] = next(
            (
                original
                for _norm, tokens, original in candidates
                if any(target and _contains_tokens(tokens, target) for target in token_targets)
            ),
            None,
        )
        if partial is not None:
            found[role] = partial
    return found


def _tokenize(value: str) -> List[str]:
    """Split a header into lowercase alphanumeric words."""
    return [token for token in re.split(r"[^a-z0-9]+", str(value).lower()) if token]


def _contains_tokens(haystack: Sequence[str], needle: Sequence[str]) -> bool:
    """True when *needle* appears as a contiguous run of whole words in *haystack*.

    Whole-word matching is the point: substring matching made the in-transit
    column "POL" (port of loading) satisfy the "po #" pattern and register as a
    purchase-order key.
    """
    span: int = len(needle)
    if span == 0 or span > len(haystack):
        return False
    return any(list(haystack[i : i + span]) == list(needle) for i in range(len(haystack) - span + 1))


def infer_grain(key_fields: Dict[str, str]) -> str:
    """Human-readable grain statement built from the detected keys."""
    ordered: List[str] = []
    for role in ("po_number", "po_line", "container", "invoice", "work_order",
                 "sku", "location", "region_channel", "date"):
        # ports describe a shipment, they are not part of its grain

        if role in key_fields:
            ordered.append(role)
    return " x ".join(ordered) if ordered else "unresolved — header detection found no key columns"


def classify_domain(
    path: str,
    haystacks: Sequence[str],
    domains: Dict[str, List[str]],
) -> Tuple[str, List[str]]:
    """Tag a source with a domain and report which keywords matched.

    Scored rather than first-match-wins on the path alone, because filenames in
    this estate are unreliable ("Table PO Status" is in-transit data).
    """
    path_text: str = str(path).lower()
    header_text: str = " | ".join(str(h).lower() for h in haystacks)
    # "ItemsOnPurchaseOrder.csv" must match the keyword "items on purchase order",
    # so every corpus is also compared with separators removed.
    path_flat: str = _flatten(path_text)
    header_flat: str = _flatten(header_text)

    scores: Dict[str, Tuple[int, List[str]]] = {}
    for domain, keywords in domains.items():
        weight: int = 0
        hits: List[str] = []
        for keyword in keywords:
            flat: str = _flatten(keyword)
            in_path: bool = keyword in path_text or (len(flat) > 3 and flat in path_flat)
            in_headers: bool = keyword in header_text or (len(flat) > 3 and flat in header_flat)
            if in_path:
                weight += PATH_MATCH_WEIGHT      # the folder/filename is the stronger signal
                hits.append(f"{keyword} (path)")
            elif in_headers:
                weight += 1
                hits.append(keyword)
        if hits:
            scores[domain] = (weight, hits)

    if not scores:
        return "unclassified", []
    order: List[str] = list(domains)
    best: str = max(scores, key=lambda d: (scores[d][0], -order.index(d)))
    return best, scores[best][1]


def _flatten(value: str) -> str:
    """Lowercase with every non-alphanumeric character removed."""
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def match_seeds(filename: str, seeds: Sequence["SeedLike"]) -> List[str]:
    """Seed ids whose glob patterns match this filename (case-insensitive)."""
    lowered: str = filename.lower()
    return [
        seed.id
        for seed in seeds
        if any(fnmatch.fnmatch(lowered, pattern.lower()) for pattern in seed.patterns)
    ]


class SeedLike:
    """Structural protocol for a seed entry (see configuration.SeedSpec)."""

    id: str
    patterns: List[str]


@dataclass
class RoleVerdict:
    role: str
    reasons: List[str] = field(default_factory=list)
    confidence: str = "medium"


def classify_role(
    *,
    query_external_sources: int,
    query_count: int,
    external_link_count: int,
    connection_count: int,
    formula_count: int,
    total_rows: int,
    has_vba: bool,
    is_csv: bool,
    heavy_formula_count: float,
    heavy_formulas_per_row: float,
) -> RoleVerdict:
    """Decide whether a file originates data or restates someone else's.

    This is the judgement that decides what Phase 2 reads. Reading a downstream
    copy instead of its source is how a pipeline ends up reporting a stale
    number that nobody can trace.
    """
    reasons: List[str] = []

    if is_csv:
        return RoleVerdict(ROLE_TRUE_SOURCE, ["flat export — originates outside Excel (ERP extract)"], "high")

    pulls_data: bool = query_external_sources > 0 or external_link_count > 0 or connection_count > 0
    formulas_per_row: float = (formula_count / total_rows) if total_rows else 0.0
    heavy_formulas: bool = (
        formula_count > heavy_formula_count and formulas_per_row >= heavy_formulas_per_row
    )

    if pulls_data:
        reasons.append(
            f"pulls from {query_external_sources} external query source(s), "
            f"{external_link_count} workbook link(s), {connection_count} connection(s)"
        )
    if heavy_formulas:
        reasons.append(f"{formula_count:,} formulas ({formulas_per_row:.1f}/row) — it computes, it does not hold")
    if has_vba:
        reasons.append("carries VBA — behaviour may not be visible in the grid")

    if pulls_data and heavy_formulas:
        return RoleVerdict(ROLE_HYBRID, reasons + ["imports and then recomputes — split these in Phase 2"], "high")
    if pulls_data:
        return RoleVerdict(ROLE_DERIVED_COPY, reasons, "high")
    if heavy_formulas:
        return RoleVerdict(ROLE_CALCULATED_OUTPUT, reasons, "medium")
    if total_rows > 0:
        return RoleVerdict(
            ROLE_TRUE_SOURCE,
            reasons + [f"{total_rows:,} rows, no external pull, low formula density — hand-maintained or pasted"],
            "medium",
        )
    return RoleVerdict(ROLE_UNKNOWN, reasons + ["no rows read — probe found nothing to classify"], "low")


def refresh_mechanism(
    *,
    query_count: int,
    connection_count: int,
    external_link_count: int,
    has_vba: bool,
    is_csv: bool,
) -> str:
    """How this source gets new data today."""
    if is_csv:
        return "netsuite_export (manual or scheduled File Cabinet extract)"
    parts: List[str] = []
    if query_count:
        parts.append(f"power_query ({query_count} queries)")
    if connection_count:
        parts.append(f"workbook_connections ({connection_count})")
    if external_link_count:
        parts.append(f"formula_links ({external_link_count} workbook link(s))")
    if has_vba:
        parts.append("vba_macro")
    return " + ".join(parts) if parts else "manual (typed or pasted by a person)"


@dataclass
class RiskAssessment:
    score: int
    band: str
    findings: List[str] = field(default_factory=list)


def assess_dependency_risk(
    *,
    locations: Sequence[str],
    has_vba: bool,
    vba_protected: bool,
    hidden_sheet_count: int,
    row_count_capped: bool,
    header_confidence: float,
    probe_errors: Sequence[str],
) -> RiskAssessment:
    """Score what breaks this source. Higher is worse.

    The score exists to rank the migration list, not to be precise. What matters
    is that a workbook reading from someone's Downloads folder ranks above one
    reading from SharePoint.
    """
    score: int = 0
    findings: List[str] = []

    for location in locations:
        for pattern, message, weight in FRAGILE_LOCATION_RULES:
            if re.search(pattern, normalize_location(location)):
                score += weight
                findings.append(f"{message} -> {location}")
                break

    if has_vba:
        score += 10
        findings.append("VBA present — logic outside the query stack and outside version control")
    if vba_protected:
        score += 10
        findings.append("VBA project is password-protected — logic cannot be reviewed or ported")
    if hidden_sheet_count:
        score += 3 * hidden_sheet_count
        findings.append(f"{hidden_sheet_count} hidden sheet(s) — staging logic a reader must not miss")
    if row_count_capped:
        score += 15
        findings.append("row count hit the configured cap — file is larger than the probe scanned")
    if 0.0 < header_confidence < 0.55:
        score += 12
        findings.append(
            f"header detection confidence {header_confidence:.2f} — layout is not a clean table, "
            "a reader will need an explicit header row"
        )
    if probe_errors:
        score += 20
        findings.append(f"{len(probe_errors)} structural probe error(s) — see errors list")

    band: str = "critical" if score >= 70 else "high" if score >= 40 else "medium" if score >= 15 else "low"
    return RiskAssessment(score=score, band=band, findings=findings)


# --------------------------------------------------------------------------
# Duplicate-truth detection
# --------------------------------------------------------------------------

@dataclass
class OverlapPair:
    """Two sheets in different files that appear to hold the same data."""

    left_file: str
    left_sheet: str
    left_rows: int
    right_file: str
    right_sheet: str
    right_rows: int
    jaccard: float
    shared_columns: List[str]
    row_delta: int
    domain: str

    @property
    def is_row_count_conflict(self) -> bool:
        """Same columns, different row counts — one of them is stale or filtered."""
        return self.row_delta != 0


def find_header_overlaps(
    entries: Sequence[Tuple[str, str, str, int, Set[str], Dict[str, str]]],
    threshold: float,
    min_columns: int,
) -> List[OverlapPair]:
    """Compare every cross-file sheet pair by header signature.

    *entries* is ``(file_path, sheet_name, domain, row_count, signature, header_map)``.
    Pairs within the same file are skipped — a workbook restating itself across
    tabs is normal; two workbooks doing it is a reconciliation problem.
    """
    candidates = [e for e in entries if len(e[4]) >= min_columns]
    pairs: List[OverlapPair] = []

    for left, right in combinations(candidates, 2):
        if left[0] == right[0]:
            continue
        left_sig: Set[str] = left[4]
        right_sig: Set[str] = right[4]
        intersection: Set[str] = left_sig & right_sig
        union: Set[str] = left_sig | right_sig
        if not union:
            continue
        jaccard: float = len(intersection) / len(union)
        if jaccard < threshold:
            continue
        shared: List[str] = sorted(
            {left[5].get(token) or right[5].get(token) or token for token in intersection}
        )
        pairs.append(
            OverlapPair(
                left_file=left[0], left_sheet=left[1], left_rows=left[3],
                right_file=right[0], right_sheet=right[1], right_rows=right[3],
                jaccard=round(jaccard, 3),
                shared_columns=shared,
                row_delta=left[3] - right[3],
                domain=left[2] if left[2] == right[2] else f"{left[2]}|{right[2]}",
            )
        )

    pairs.sort(key=lambda p: (-abs(p.row_delta), -p.jaccard))
    return pairs
