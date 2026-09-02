"""Render the reconciliation report."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Sequence

from .m_ast import QueryProfile
from .reconcile import (
    SEVERITY_BLOCKING,
    SEVERITY_INFO,
    SEVERITY_WARNING,
    Finding,
    ForkGroup,
)

KIND_TITLES: Dict[str, str] = {
    "measure_definition_conflict": "Same column name, different measure",
    "key_normalization_conflict": "Join key built differently",
    "key_derivation_missing": "Join key derived on one side only",
    "parse_option_conflict": "Same file, different parser",
    "filter_divergence": "Different row scope",
    "grain_difference": "Different grain",
    "measure_name_conflict": "Same measure, different name",
}


def _table(rows: Sequence[Sequence[str]], headers: Sequence[str]) -> List[str]:
    lines: List[str] = ["| " + " | ".join(headers) + " |",
                        "|" + "|".join("---" for _ in headers) + "|"]
    lines.extend("| " + " | ".join(str(c).replace("|", "\\|") for c in row) + " |" for row in rows)
    return lines


def render(groups: Sequence[ForkGroup], findings: Sequence[Finding], generated_at: str) -> str:
    lines: List[str] = []
    add = lines.append

    blocking: List[Finding] = [f for f in findings if f.severity == SEVERITY_BLOCKING]
    warning: List[Finding] = [f for f in findings if f.severity == SEVERITY_WARNING]
    info: List[Finding] = [f for f in findings if f.severity == SEVERITY_INFO]
    forked: List[ForkGroup] = [g for g in groups if g.is_forked]

    add("# Query reconciliation — where the forked logic diverges")
    add("")
    add(f"Generated {generated_at} from the committed M source in `queries/`. "
        "No refresh, no Excel.")
    add("")

    add("## Executive summary")
    add("")
    add(f"- **{len(forked)} upstream export(s) are read by more than one workbook.**")
    add(f"- **{len(blocking)} blocking conflict(s)** — the numbers cannot agree until these are resolved.")
    add(f"- {len(warning)} divergence(s) that change scope or grain, and {len(info)} naming inconsistency(ies).")
    add("")

    if blocking:
        add("### The blocking conflicts")
        add("")
        for index, finding in enumerate(blocking, start=1):
            add(f"**{index}. {KIND_TITLES.get(finding.kind, finding.kind)} — {finding.subject}**")
            add("")
            add(f"- Upstream: {finding.upstream_label}")
            add(f"- `{finding.left}` -> `{finding.left_value}`")
            add(f"- `{finding.right}` -> `{finding.right_value}`")
            add(f"- {finding.consequence}")
            if finding.recommendation:
                add(f"- **Fix:** {finding.recommendation}")
            add("")
    else:
        add("No blocking conflicts found.")
        add("")

    add("## All findings")
    add("")
    if findings:
        rows: List[List[str]] = [
            [
                f.severity,
                KIND_TITLES.get(f.kind, f.kind),
                f.upstream_label[:44],
                f.subject[:40],
                f"`{f.left}`",
                f.left_value[:70],
                f"`{f.right}`",
                f.right_value[:70],
            ]
            for f in findings
        ]
        lines.extend(_table(rows, ["Severity", "Finding", "Upstream", "Subject",
                                   "Query A", "A value", "Query B", "B value"]))
    else:
        add("None.")
    add("")

    add("## Per-upstream detail")
    add("")
    for group in groups:
        marker: str = " — **FORKED**" if group.is_forked else ""
        add(f"### {group.upstream_label}{marker}")
        add("")
        rows = []
        for profile in sorted(group.profiles, key=lambda p: (p.workbook, p.query)):
            measures: str = "; ".join(f"{name} = {sig}" for name, sig in sorted(profile.measures.items()))
            rows.append([
                f"`{profile.workbook}`",
                profile.query,
                str(profile.step_count),
                profile.source.option("QuoteStyle") or "—",
                " x ".join(profile.group_keys) or "—",
                measures or "—",
                "; ".join(profile.filters)[:60] or "—",
            ])
        lines.extend(_table(rows, ["Workbook", "Query", "Steps", "QuoteStyle", "Grain",
                                   "Measures (what they really aggregate)", "Filters"]))
        add("")

    add("## What to do with this")
    add("")
    add("Each forked export should be read **once** into the canonical layer, then reused. "
        "The reconciliation above is the specification for that single extraction: for every "
        "conflict, one side is right, and the answer belongs in `SCHEMA.md` rather than in "
        "seven copies of a query.")
    add("")
    add("Order of work:")
    add("")
    add("1. Resolve the blocking conflicts — they are cases where two workbooks are already "
        "reporting different numbers under the same label.")
    add("2. Decide, per forked export, whether differing grain and scope are intentional. "
        "Where they are, both variants derive from one extraction.")
    add("3. Settle the naming, once.")
    add("")
    return "\n".join(lines)


def write_reports(
    groups: Sequence[ForkGroup],
    findings: Sequence[Finding],
    profiles: Sequence[QueryProfile],
    out_dir: Path,
    generated_at: str,
) -> List[Path]:
    """Write RECONCILIATION.md plus a machine-readable companion."""
    out_dir.mkdir(parents=True, exist_ok=True)
    markdown: Path = out_dir / "RECONCILIATION.md"
    markdown.write_text(render(groups, findings, generated_at), encoding="utf-8")

    payload: Dict[str, Any] = {
        "generated_at": generated_at,
        "counts": {
            "upstreams": len(groups),
            "forked_upstreams": sum(1 for g in groups if g.is_forked),
            "findings": len(findings),
            "blocking": sum(1 for f in findings if f.severity == SEVERITY_BLOCKING),
            "warning": sum(1 for f in findings if f.severity == SEVERITY_WARNING),
            "info": sum(1 for f in findings if f.severity == SEVERITY_INFO),
        },
        "findings": [
            {
                "kind": f.kind, "severity": f.severity, "upstream": f.upstream_label,
                "subject": f.subject, "left": f.left, "left_value": f.left_value,
                "right": f.right, "right_value": f.right_value,
                "consequence": f.consequence, "recommendation": f.recommendation,
            }
            for f in findings
        ],
        "query_profiles": [
            {
                "workbook": p.workbook, "query": p.query,
                "upstream_key": p.upstream_key, "upstream_label": p.upstream_label,
                "reader": p.source.reader, "fetcher": p.source.fetcher,
                "parse_options": p.source.options,
                "grain": p.group_keys,
                "measures": p.measures,
                "filters": p.filters,
                "key_derivations": p.key_derivations,
                "output_columns": p.output_columns,
                "step_count": p.step_count,
            }
            for p in sorted(profiles, key=lambda p: (p.workbook, p.query))
        ],
    }
    json_path: Path = out_dir / "reconciliation.json"
    json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return [markdown, json_path]
