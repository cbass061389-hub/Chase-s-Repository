"""Phase 0 outputs: ``discovery/manifest.json`` and ``discovery/DISCOVERY.md``.

manifest.json is the machine contract every later phase reads.
DISCOVERY.md is written for a decision, not for completeness: conclusion first,
then the ranked table, then what has to happen next.
"""

from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path
from typing import Any, Dict, List, Sequence

from .scan import STATUS_FAILED, DiscoveryResult, SourceRecord

ROLE_LABELS: Dict[str, str] = {
    "true_source": "SOURCE",
    "derived_copy": "COPY",
    "calculated_output": "OUTPUT",
    "hybrid": "HYBRID",
    "unknown": "?",
}
MANIFEST_VERSION: int = 1


def _plural(count: int, singular: str, plural: str) -> str:
    """Agreement helper — this report is read by executives, not by a parser."""
    return singular if count == 1 else plural


def write_manifest(result: DiscoveryResult, discovery_dir: Path) -> Path:
    """Serialize the full result. Nothing is summarized away here."""
    discovery_dir.mkdir(parents=True, exist_ok=True)
    target: Path = discovery_dir / "manifest.json"
    payload: Dict[str, Any] = {
        "manifest_version": MANIFEST_VERSION,
        "run": {
            "started_at": result.started_at,
            "finished_at": result.finished_at,
            "duration_seconds": result.duration_seconds,
            "config_path": result.config_path,
        },
        "counts": result.counts,
        "roots": result.roots,
        "seed_status": result.seed_status,
        "sources": [asdict(source) for source in result.sources],
        "duplicate_truth_candidates": result.overlaps,
        "query_files_written": result.query_files_written,
        "query_diffs": result.query_diffs,
        "skipped": result.skipped,
    }
    target.write_text(json.dumps(payload, indent=2, sort_keys=False), encoding="utf-8")
    return target


def _table(rows: Sequence[Sequence[str]], headers: Sequence[str]) -> List[str]:
    """Markdown table with no column-width padding (GitHub renders it fine)."""
    lines: List[str] = ["| " + " | ".join(headers) + " |",
                        "|" + "|".join("---" for _ in headers) + "|"]
    lines.extend("| " + " | ".join(str(cell).replace("|", "\\|") for cell in row) + " |" for row in rows)
    return lines


def _priority(source: SourceRecord) -> tuple:
    """Ranking for the source table: broken first, then risky, then large."""
    return (
        0 if source.probe_status == STATUS_FAILED else 1,
        -source.risk_score,
        -source.total_rows,
    )


def render_discovery_md(result: DiscoveryResult) -> str:
    """The human-readable Phase 0 report."""
    counts: Dict[str, int] = result.counts
    lines: List[str] = []
    add = lines.append

    unreachable: List[Dict[str, Any]] = [r for r in result.roots if not r["reachable"]]
    missing_seeds: List[Dict[str, Any]] = [s for s in result.seed_status if not s["found"]]
    conflicts: List[Dict[str, Any]] = [o for o in result.overlaps if o["is_row_count_conflict"]]
    failed: List[SourceRecord] = [s for s in result.sources if s.probe_status == STATUS_FAILED]
    critical: List[SourceRecord] = [s for s in result.sources if s.risk_band in ("critical", "high")]

    add("# Phase 0 — Discovery")
    add("")
    add(f"Generated {result.finished_at} in {result.duration_seconds}s from `{result.config_path}`.")
    add("Read-only sweep. No source workbook was opened in Excel, written to, or had a macro executed.")
    add("")

    # ---------------- Executive summary ----------------
    add("## Executive summary")
    add("")
    if unreachable:
        add(f"**Sweep was incomplete.** {len(unreachable)} configured root(s) could not be read, "
            "so this manifest is not a full picture of the estate:")
        for root in unreachable:
            add(f"- `{root['root']}` — {root['note']}")
        add("")

    reachable: int = sum(1 for r in result.roots if r["reachable"])
    add(f"- **{counts['sources']} {_plural(counts['sources'], 'source', 'sources')}** found "
        f"across {reachable} reachable {_plural(reachable, 'root', 'roots')}.")
    add(f"- **{counts['power_query_workbooks']} "
        f"{_plural(counts['power_query_workbooks'], 'workbook carries', 'workbooks carry')} Power Query** "
        f"({len(result.query_files_written)} M files now version-controlled under `queries/`).")
    add(f"- **{counts['vba_workbooks']} "
        f"{_plural(counts['vba_workbooks'], 'workbook carries', 'workbooks carry')} VBA** — "
        "logic that lives outside the query stack.")
    add(f"- **{len(critical)} {_plural(len(critical), 'source ranks', 'sources rank')} high or critical** "
        "on dependency risk.")
    add(f"- **{counts['duplicate_truth_pairs']} duplicate-truth "
        f"{_plural(counts['duplicate_truth_pairs'], 'candidate', 'candidates')}** "
        f"({len(conflicts)} of them {_plural(len(conflicts), 'disagrees', 'disagree')} on row count).")
    add(f"- **{counts['seeds_missing']} of {len(result.seed_status)} expected workbooks were not found.**")
    if counts["failed_probes"]:
        add(f"- **{counts['failed_probes']} "
            f"{_plural(counts['failed_probes'], 'file', 'files')} could not be probed at all.**")
    add("")

    add("### Key risks")
    add("")
    risk_written: bool = False
    if conflicts:
        worst = conflicts[0]
        add(f"1. **Two sources claim the same data with different numbers.** Worst case: "
            f"`{worst['left_file']}!{worst['left_sheet']}` ({worst['left_rows']:,} rows) vs "
            f"`{worst['right_file']}!{worst['right_sheet']}` ({worst['right_rows']:,} rows) — "
            f"{worst['jaccard']:.0%} header overlap, {abs(worst['row_delta']):,} row difference. "
            "Until one is named the source, every downstream number is arguable.")
        risk_written = True
    if critical:
        worst_risk: SourceRecord = max(critical, key=lambda s: s.risk_score)
        add(f"{'2' if risk_written else '1'}. **Fragile paths.** `{worst_risk.relative_path}` scores "
            f"{worst_risk.risk_score} ({worst_risk.risk_band}): "
            + "; ".join(worst_risk.risk_findings[:2]))
        risk_written = True
    if missing_seeds:
        add(f"{'3' if risk_written else '1'}. **{len(missing_seeds)} expected "
            f"{_plural(len(missing_seeds), 'workbook', 'workbooks')} not located** — "
            "either renamed, in a skipped folder, or outside the configured roots. "
            "Named below; each needs a path before Phase 2 can claim coverage.")
        risk_written = True
    if not risk_written:
        add("No blocking risks found in this sweep.")
    add("")

    add("### Recommended action")
    add("")
    add("1. Resolve the duplicate-truth pairs below — for each, tell me which file is the source.")
    add("2. Confirm or correct the domain and grain assignments in the source table.")
    add("3. Point `discovery.roots` / `extra_roots` in `sc/config.yaml` at anything the sweep missed, then re-run.")
    add("4. Sign off on the canonical model in `SCHEMA.md` — Phase 2 does not start before that.")
    add("")

    # ---------------- Ranked source table ----------------
    add("## Sources, ranked worst-first")
    add("")
    add("`ROLE` is inferred from evidence in the file — formula density, query sources and external "
        "links — not from its name. A `COPY` must not be read by Phase 2 when its source is also present.")
    add("")
    rows: List[List[str]] = []
    for source in sorted(result.sources, key=_priority):
        sheet_summary: str = (
            f"{len(source.sheets)} ({sum(1 for s in source.sheets if s.state != 'visible')} hidden)"
            if source.sheets else "—"
        )
        rows.append([
            f"`{source.relative_path}`",
            source.domain,
            ROLE_LABELS.get(source.role, source.role),
            source.grain[:60],
            source.refresh_mechanism[:44],
            f"{source.risk_score} {source.risk_band}",
            f"{source.total_rows:,}",
            sheet_summary,
            source.probe_status,
        ])
    lines.extend(_table(rows, ["File", "Domain", "Role", "Grain (inferred)", "Refresh",
                              "Risk", "Rows", "Sheets", "Probe"]))
    add("")

    # ---------------- Conflicts ----------------
    add("## Duplicate truth — two sources, same columns")
    add("")
    if not result.overlaps:
        add("None detected above the configured header-overlap threshold.")
    else:
        add("Matched on normalized header signature, so this catches duplicates whose filenames "
            "look nothing alike. A row-count difference means at least one is stale or filtered.")
        add("")
        overlap_rows: List[List[str]] = [
            [
                f"`{o['left_file']}`!{o['left_sheet']}",
                f"{o['left_rows']:,}",
                f"`{o['right_file']}`!{o['right_sheet']}",
                f"{o['right_rows']:,}",
                f"{o['jaccard']:.0%}",
                f"{o['row_delta']:+,}" if o["row_delta"] else "0",
                o["domain"],
                ", ".join(o["shared_columns"][:6]),
            ]
            for o in result.overlaps
        ]
        lines.extend(_table(overlap_rows, ["Source A", "Rows A", "Source B", "Rows B",
                                           "Header overlap", "Row delta", "Domain", "Shared columns"]))
    add("")

    # ---------------- Seed confirmation ----------------
    add("## Expected workbooks — confirmed vs missing")
    add("")
    seed_rows: List[List[str]] = [
        [
            "found" if seed["found"] else "**MISSING**",
            seed["label"],
            seed["domain"],
            ", ".join(f"`{m['path']}`" for m in seed["matches"]) if seed["matches"]
            else ", ".join(f"`{p}`" for p in seed["patterns"]),
        ]
        for seed in sorted(result.seed_status, key=lambda s: (s["found"], s["label"]))
    ]
    lines.extend(_table(seed_rows, ["Status", "Expected asset", "Domain", "Found at / searched for"]))
    add("")

    # ---------------- External dependency map ----------------
    add("## External dependency map — what breaks if a path moves")
    add("")
    dependency_rows: List[List[str]] = []
    for source in result.sources:
        for location in source.external_locations:
            dependency_rows.append([
                f"`{source.relative_path}`",
                location["kind"],
                location.get("via", ""),
                f"`{location['location'][:96]}`",
            ])
    if dependency_rows:
        dependency_rows.sort(key=lambda r: (r[1], r[0]))
        lines.extend(_table(dependency_rows, ["Dependent file", "Kind", "Via", "Location"]))
    else:
        add("No external locations detected. Every source is self-contained or manually maintained.")
    add("")

    # ---------------- Power Query inventory ----------------
    add("## Power Query inventory")
    add("")
    pq_sources = [s for s in result.sources if s.power_query.get("query_count")]
    if not pq_sources:
        add("No Power Query found in the swept estate.")
    else:
        for source in sorted(pq_sources, key=lambda s: -int(s.power_query["query_count"])):
            add(f"### `{source.relative_path}` — {source.power_query['query_count']} queries")
            add("")
            query_rows: List[List[str]] = [
                [
                    q["name"],
                    str(q["lines"]),
                    ", ".join(sorted({s["kind"] for s in q["sources"] if s["kind"] != "query_ref"})) or "—",
                    ", ".join(s["location"] for s in q["sources"] if s["kind"] == "query_ref") or "—",
                ]
                for q in source.power_query["queries"]
            ]
            lines.extend(_table(query_rows, ["Query", "Lines", "External source kinds", "Depends on queries"]))
            add("")

    # ---------------- VBA ----------------
    vba_sources = [s for s in result.sources if s.vba.get("present")]
    add("## VBA present")
    add("")
    if not vba_sources:
        add("No VBA found. Macros were never executed by this sweep in any case.")
    else:
        add("Module names only — nothing was executed. A password-protected project cannot be "
            "reviewed or ported, which makes it a hard blocker rather than a risk.")
        add("")
        vba_rows: List[List[str]] = [
            [
                f"`{s.relative_path}`",
                s.vba.get("project_name") or "—",
                "**yes**" if s.vba.get("protected") else "no",
                str(s.vba.get("module_count", 0)),
                ", ".join(s.vba.get("modules", [])[:10]),
            ]
            for s in vba_sources
        ]
        lines.extend(_table(vba_rows, ["File", "Project", "Protected", "Components", "Names"]))
    add("")

    # ---------------- Failures ----------------
    add("## Probe failures and warnings")
    add("")
    problem_sources = [s for s in result.sources if s.errors]
    if not problem_sources and not result.skipped:
        add("Clean sweep — no probe errors.")
    else:
        for source in sorted(problem_sources, key=_priority):
            add(f"- `{source.relative_path}` ({source.probe_status})")
            for error in source.errors:
                add(f"  - {error}")
        for entry in result.skipped:
            add(f"- `{entry['path']}` — {entry['reason']}")
    add("")

    if result.query_diffs:
        add("## Power Query changes since the last discovery run")
        add("")
        add("A query edit that breaks the pipeline is visible here instead of being debugged blind.")
        add("")
        diff_rows: List[List[str]] = [
            [f"`{d['workbook']}`", d["query"], d["previous_lines"], d["current_lines"], d["change"]]
            for d in result.query_diffs
        ]
        lines.extend(_table(diff_rows, ["Workbook", "Query", "Lines before", "Lines now", "Change"]))
        add("")

    add("## Next steps")
    add("")
    add("| # | Action | Owner |")
    add("|---|---|---|")
    add("| 1 | Name the source of record for each duplicate-truth pair | Chase |")
    add("| 2 | Correct any wrong domain/grain call in the source table | Chase |")
    add("| 3 | Supply paths for the missing expected workbooks | Chase |")
    add("| 4 | Sign off `SCHEMA.md` | Chase |")
    add("| 5 | Build Phase 2 readers against the confirmed sources only | Claude |")
    add("")
    add("Phase 2 is blocked on items 1-4 by design. Extraction code written before the model is "
        "agreed is code that gets rewritten.")
    add("")
    return "\n".join(lines)


def write_discovery_md(result: DiscoveryResult, discovery_dir: Path) -> Path:
    discovery_dir.mkdir(parents=True, exist_ok=True)
    target: Path = discovery_dir / "DISCOVERY.md"
    target.write_text(render_discovery_md(result), encoding="utf-8")
    return target
