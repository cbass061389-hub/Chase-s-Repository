"""Reconcile queries that read the same upstream.

A forked upstream — one export, several workbooks — is only a problem because
the transformations differ. This module says *how* they differ, in the terms
that decide whether the numbers can agree:

======================================  ==========  ===================================
Finding                                 Severity    Why it matters
======================================  ==========  ===================================
measure_definition_conflict             blocking    same column name, different source
key_normalization_conflict              blocking    join keys built differently
parse_option_conflict                   blocking    one reader can corrupt rows
filter_divergence                       warning     different row scope
grain_difference                        warning     different aggregation level
measure_name_conflict                   info        same number, different label
======================================  ==========  ===================================

Everything is derived from the M source, so it holds without running a refresh
and without Excel.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from itertools import combinations
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

from .m_ast import QueryProfile

SEVERITY_BLOCKING: str = "blocking"
SEVERITY_WARNING: str = "warning"
SEVERITY_INFO: str = "info"

SEVERITY_ORDER: Dict[str, int] = {SEVERITY_BLOCKING: 0, SEVERITY_WARNING: 1, SEVERITY_INFO: 2}

#: Parse options where a mismatch changes what actually gets read.
CRITICAL_PARSE_OPTIONS: Tuple[str, ...] = ("QuoteStyle", "Columns", "Delimiter", "Encoding")

#: Normalization calls whose presence or absence changes a key's value.
NORMALIZERS: Tuple[str, ...] = ("Text.Trim", "Text.Upper", "Text.Lower", "Text.Clean", "Text.Proper")


@dataclass
class Finding:
    """One concrete disagreement between two queries on the same upstream."""

    kind: str
    severity: str
    upstream_label: str
    left: str
    right: str
    subject: str
    left_value: str
    right_value: str
    consequence: str
    recommendation: str = ""

    @property
    def sort_key(self) -> Tuple[int, str, str]:
        return (SEVERITY_ORDER.get(self.severity, 9), self.upstream_label, self.kind)


@dataclass
class ForkGroup:
    """Every query reading one canonical upstream."""

    upstream_key: str
    upstream_label: str
    profiles: List[QueryProfile] = field(default_factory=list)

    @property
    def workbooks(self) -> Set[str]:
        return {profile.workbook for profile in self.profiles}

    @property
    def is_forked(self) -> bool:
        return len(self.workbooks) > 1


def _label(profile: QueryProfile) -> str:
    return f"{profile.workbook}!{profile.query}"


def _normalizers_used(expression: str) -> Set[str]:
    return {name for name in NORMALIZERS if name in expression}


def _strip_expression(expression: str) -> str:
    """Collapse whitespace so two identical expressions compare equal."""
    return re.sub(r"\s+", " ", expression).strip()


def compare_pair(group: ForkGroup, left: QueryProfile, right: QueryProfile) -> List[Finding]:
    """Every disagreement between two queries reading the same upstream."""
    findings: List[Finding] = []
    left_name: str = _label(left)
    right_name: str = _label(right)

    # ---- 1. Parse options: a different reader can produce different rows ----
    for option in CRITICAL_PARSE_OPTIONS:
        left_value: str = left.source.option(option)
        right_value: str = right.source.option(option)
        if left_value == right_value:
            continue
        if not left_value and not right_value:
            continue
        dangerous: bool = "QuoteStyle.None" in (left_value, right_value)
        findings.append(Finding(
            kind="parse_option_conflict",
            severity=SEVERITY_BLOCKING if dangerous else SEVERITY_WARNING,
            upstream_label=group.upstream_label,
            left=left_name, right=right_name,
            subject=f"Csv.Document option {option}",
            left_value=left_value or "(not set)",
            right_value=right_value or "(not set)",
            consequence=(
                "QuoteStyle.None ignores CSV quoting, so any field containing a comma "
                "splits across columns and silently shifts every value after it on that row. "
                "The two queries are reading the same bytes into different tables."
                if dangerous else
                "The same file is being parsed with different options, so the two queries "
                "can disagree on rows or columns before any transformation runs."
            ),
            recommendation=(
                "Read the export once, with QuoteStyle.Csv, and reuse it."
                if dangerous else
                "Standardise the parse options on one extraction."
            ),
        ))

    # ---- 2. Measures: the same output name must mean the same thing ----
    left_measures: Dict[str, str] = left.measures
    right_measures: Dict[str, str] = right.measures
    for name in sorted(set(left_measures) & set(right_measures)):
        if left_measures[name] == right_measures[name]:
            continue
        findings.append(Finding(
            kind="measure_definition_conflict",
            severity=SEVERITY_BLOCKING,
            upstream_label=group.upstream_label,
            left=left_name, right=right_name,
            subject=f'output column "{name}"',
            left_value=left_measures[name],
            right_value=right_measures[name],
            consequence=(
                f'Both queries emit a column called "{name}" from the same export, but they '
                "aggregate different source columns. Any comparison, join or roll-up across "
                "the two is comparing different measures under one name — and nothing in "
                "either workbook shows that."
            ),
            recommendation=(
                "Define this measure once in the canonical layer and give each variant a "
                "distinct, honest name."
            ),
        ))

    # Same underlying measure emitted under different names.
    reverse_left: Dict[str, List[str]] = {}
    for name, signature in left_measures.items():
        reverse_left.setdefault(signature, []).append(name)
    for signature, right_names in (
        (sig, [n for n, s in right_measures.items() if s == sig]) for sig in set(right_measures.values())
    ):
        left_names: List[str] = reverse_left.get(signature, [])
        if not left_names or not right_names or set(left_names) == set(right_names):
            continue
        findings.append(Finding(
            kind="measure_name_conflict",
            severity=SEVERITY_INFO,
            upstream_label=group.upstream_label,
            left=left_name, right=right_name,
            subject=signature,
            left_value=", ".join(sorted(left_names)),
            right_value=", ".join(sorted(right_names)),
            consequence="The same measure is labelled differently in each workbook, so a reader "
                        "cannot tell they are the same number.",
            recommendation="Adopt one column name in the canonical layer.",
        ))

    # ---- 3. Key derivation: a join key built two ways will not join ----
    left_keys: Dict[str, str] = left.key_derivations
    right_keys: Dict[str, str] = right.key_derivations
    for name in sorted(set(left_keys) & set(right_keys)):
        if _strip_expression(left_keys[name]) == _strip_expression(right_keys[name]):
            continue
        left_norms: Set[str] = _normalizers_used(left_keys[name])
        right_norms: Set[str] = _normalizers_used(right_keys[name])
        difference: Set[str] = left_norms ^ right_norms
        findings.append(Finding(
            kind="key_normalization_conflict",
            severity=SEVERITY_BLOCKING if difference else SEVERITY_WARNING,
            upstream_label=group.upstream_label,
            left=left_name, right=right_name,
            subject=f'key column "{name}"',
            left_value=_strip_expression(left_keys[name])[:200],
            right_value=_strip_expression(right_keys[name])[:200],
            consequence=(
                f"Only one side applies {', '.join(sorted(difference))}. The same NetSuite value "
                "therefore produces two different key strings, so rows that should join do not, "
                "and the same item can appear twice in a combined view."
                if difference else
                "The join key is derived differently on each side, so the two can disagree on "
                "which rows correspond."
            ),
            recommendation="Normalise the key once, in the canonical layer, and never per query.",
        ))

    # Keys derived on one side but not the other, under the same name.
    for name in sorted(set(left_keys) ^ set(right_keys)):
        holder, other = (left, right) if name in left_keys else (right, left)
        if name in other.output_columns:
            continue
        findings.append(Finding(
            kind="key_derivation_missing",
            severity=SEVERITY_WARNING,
            upstream_label=group.upstream_label,
            left=left_name, right=right_name,
            subject=f'key column "{name}"',
            left_value=_strip_expression(left_keys.get(name, "(absent)"))[:160],
            right_value=_strip_expression(right_keys.get(name, "(absent)"))[:160],
            consequence=f"Only {_label(holder)} derives a key named \"{name}\"; the other side "
                        "carries the raw value, so the two cannot be joined without a translation step.",
            recommendation="Derive the key once, upstream of both.",
        ))

    # ---- 4. Grain ----
    if set(left.group_keys) != set(right.group_keys) and (left.group_keys or right.group_keys):
        findings.append(Finding(
            kind="grain_difference",
            severity=SEVERITY_WARNING,
            upstream_label=group.upstream_label,
            left=left_name, right=right_name,
            subject="group-by keys",
            left_value=" x ".join(left.group_keys) or "(no grouping)",
            right_value=" x ".join(right.group_keys) or "(no grouping)",
            consequence="The two aggregate the export to different levels, so their row counts "
                        "and totals are not comparable.",
            recommendation="If both grains are genuinely needed, derive both from one extraction "
                           "rather than reading the export twice.",
        ))

    # ---- 5. Filters ----
    left_filters: Set[str] = {_strip_expression(f) for f in left.filters}
    right_filters: Set[str] = {_strip_expression(f) for f in right.filters}
    if left_filters != right_filters:
        findings.append(Finding(
            kind="filter_divergence",
            severity=SEVERITY_WARNING,
            upstream_label=group.upstream_label,
            left=left_name, right=right_name,
            subject="row filters",
            left_value="; ".join(sorted(left_filters)) or "(no filter)",
            right_value="; ".join(sorted(right_filters)) or "(no filter)",
            consequence="Different row scope from the same export. This is the usual reason two "
                        "reports show different totals for what looks like the same thing.",
            recommendation="Decide whether the narrower scope is correct, and apply it once.",
        ))

    return findings


def reconcile(profiles: Sequence[QueryProfile]) -> Tuple[List[ForkGroup], List[Finding]]:
    """Group profiles by upstream and compare every pair within each forked group."""
    groups: Dict[str, ForkGroup] = {}
    for profile in profiles:
        if not profile.upstream_key:
            continue
        group: ForkGroup = groups.setdefault(
            profile.upstream_key,
            ForkGroup(upstream_key=profile.upstream_key, upstream_label=profile.upstream_label),
        )
        group.profiles.append(profile)

    findings: List[Finding] = []
    for group in groups.values():
        if len(group.profiles) < 2:
            continue
        for left, right in combinations(sorted(group.profiles, key=_label), 2):
            findings.extend(compare_pair(group, left, right))

    findings.sort(key=lambda finding: finding.sort_key)
    ordered: List[ForkGroup] = sorted(
        groups.values(), key=lambda g: (-len(g.profiles), g.upstream_label)
    )
    return ordered, findings
