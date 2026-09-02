"""Command entry point.

    python -m sc.cli discover                      # Phase 0 sweep
    python -m sc.cli discover --root "D:/Prototypes"
    python -m sc.cli run --refresh all --build reference,html

Phase 0 is implemented. ``run`` is deliberately gated: the extraction, refresh,
reference-workbook and HTML phases are not built until the canonical model in
SCHEMA.md is signed off, per the build spec's rules of engagement. It reports
that instead of silently doing nothing or pretending to work.
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional, Sequence

from .configuration import Config, ConfigError, load_config

EXIT_OK: int = 0
EXIT_ERROR: int = 1
EXIT_BLOCKED: int = 2

LOG_FORMAT: str = "%(asctime)s %(levelname)-7s %(name)-22s %(message)s"
logger: logging.Logger = logging.getLogger("sc.cli")


def _configure_logging(verbose: bool, log_file: Optional[Path]) -> None:
    handlers: List[logging.Handler] = [logging.StreamHandler(sys.stderr)]
    if log_file is not None:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format=LOG_FORMAT,
        handlers=handlers,
        force=True,
    )


def _apply_root_overrides(config: Config, roots: Sequence[str]) -> Config:
    """Replace the configured roots for this run only. Nothing on disk changes."""
    if not roots:
        return config
    from dataclasses import replace

    return replace(config, discovery=replace(config.discovery, roots=list(roots), extra_roots=[]))


def command_discover(args: argparse.Namespace) -> int:
    """Phase 0: read-only sweep, manifest, and DISCOVERY.md."""
    from .discovery.report import write_discovery_md, write_manifest
    from .discovery.scan import run_discovery

    config: Config = _apply_root_overrides(load_config(args.config), args.root or [])
    run_stamp: str = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    run_dir: Path = config.path("runs_dir") / run_stamp
    _configure_logging(args.verbose, run_dir / "discover.log")

    logger.info("phase 0 discovery starting; config=%s", config.config_path)
    for root in config.discovery.resolved_roots():
        logger.info("root: %s (exists=%s)", root, root.is_dir())

    result = run_discovery(config, write_queries=not args.no_queries)

    discovery_dir: Path = config.path("discovery_dir")
    manifest_path: Path = write_manifest(result, discovery_dir)
    logger.info("wrote %s", manifest_path)

    md_path: Optional[Path] = None
    if not args.json_only:
        md_path = write_discovery_md(result, discovery_dir)
        logger.info("wrote %s", md_path)

    run_dir.mkdir(parents=True, exist_ok=True)
    (run_dir / "manifest.json").write_text(
        json.dumps(
            {
                "command": "discover",
                "started_at": result.started_at,
                "finished_at": result.finished_at,
                "duration_seconds": result.duration_seconds,
                "counts": result.counts,
                "roots": result.roots,
                "per_source_rows": {s.relative_path: s.total_rows for s in result.sources},
                "outputs": [str(manifest_path)] + ([str(md_path)] if md_path else []),
            },
            indent=2,
        ),
        encoding="utf-8",
    )

    counts = result.counts
    logger.info(
        "discovery complete: %d sources, %d PQ workbooks, %d VBA workbooks, "
        "%d duplicate-truth pairs, %d missing seeds, %d failed probes (%.2fs)",
        counts["sources"], counts["power_query_workbooks"], counts["vba_workbooks"],
        counts["duplicate_truth_pairs"], counts["seeds_missing"], counts["failed_probes"],
        result.duration_seconds,
    )

    unreachable: List[str] = [r["root"] for r in result.roots if not r["reachable"]]
    if unreachable:
        logger.error("unreachable roots — this manifest is NOT a full picture: %s", unreachable)
        return EXIT_BLOCKED
    if not result.sources:
        logger.error("no sources found under any root; check discovery.roots in %s", config.config_path)
        return EXIT_BLOCKED
    if counts["failed_probes"]:
        logger.warning("%d file(s) could not be probed; see DISCOVERY.md", counts["failed_probes"])
        return EXIT_BLOCKED
    return EXIT_OK


def command_reconcile(args: argparse.Namespace) -> int:
    """Compare the queries that read the same upstream export.

    Runs entirely from committed state: the manifest supplies the lineage, the
    ``queries/*/_Section1.m`` files supply the M source. No refresh, no Excel,
    no access to the original workbooks.
    """
    from .analyze.m_ast import profile_query
    from .analyze.reconcile import SEVERITY_BLOCKING, reconcile
    from .analyze.reconcile_report import write_reports
    from .discovery.datamashup import parse_queries

    config: Config = load_config(args.config)
    _configure_logging(args.verbose, None)

    discovery_dir: Path = config.path("discovery_dir")
    manifest_path: Path = discovery_dir / "manifest.json"
    if not manifest_path.is_file():
        logger.error("no manifest at %s — run `python -m sc.cli discover` first", manifest_path)
        return EXIT_BLOCKED

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    edges = manifest.get("lineage", {}).get("upstream_edges", [])
    if not edges:
        logger.error("manifest has no lineage edges; re-run discover with this version")
        return EXIT_BLOCKED

    # (workbook, query) -> canonical upstream
    upstream_by_query: dict[tuple[str, str], tuple[str, str]] = {
        (edge["consumer"], edge["query"]): (edge["upstream_key"], edge["upstream_label"])
        for edge in edges
    }

    queries_dir: Path = config.path("queries_dir")
    profiles = []
    for source in manifest["sources"]:
        if not source.get("power_query", {}).get("query_count"):
            continue
        stem: str = Path(source["filename"]).stem
        section: Path = queries_dir / f"{stem}__{source['id']}" / "_Section1.m"
        if not section.is_file():
            logger.warning("no extracted M for %s (expected %s)", source["relative_path"], section)
            continue
        for query in parse_queries(section.read_text(encoding="utf-8")):
            profile = profile_query(source["relative_path"], query.name, query.source)
            key, label = upstream_by_query.get((source["relative_path"], query.name), ("", ""))
            profile.upstream_key, profile.upstream_label = key, label
            profiles.append(profile)

    if not profiles:
        logger.error("no queries could be profiled; check %s", queries_dir)
        return EXIT_BLOCKED

    groups, findings = reconcile(profiles)
    generated_at: str = datetime.now(timezone.utc).isoformat(timespec="seconds")
    written = write_reports(groups, findings, profiles, discovery_dir, generated_at)
    for path in written:
        logger.info("wrote %s", path)

    blocking = [f for f in findings if f.severity == SEVERITY_BLOCKING]
    logger.info(
        "reconciled %d queries across %d upstreams (%d forked): %d findings, %d blocking",
        len(profiles), len(groups), sum(1 for g in groups if g.is_forked),
        len(findings), len(blocking),
    )
    for finding in blocking:
        logger.warning("BLOCKING %s | %s | %s: %s vs %s",
                       finding.kind, finding.upstream_label, finding.subject,
                       finding.left_value, finding.right_value)
    return EXIT_BLOCKED if blocking else EXIT_OK



def _load_query_profiles(config: Config) -> tuple[list, dict, dict]:
    """Profile every committed query. Shared by `reconcile`, `schemas` and `extract`.

    Runs from committed state only: the manifest supplies lineage, the extracted
    _Section1.m files supply the M source.
    """
    from .analyze.m_ast import profile_query
    from .discovery.datamashup import parse_queries

    manifest_path: Path = config.path("discovery_dir") / "manifest.json"
    if not manifest_path.is_file():
        raise FileNotFoundError(
            f"no manifest at {manifest_path} — run `python -m sc.cli discover` first"
        )
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    edges = manifest.get("lineage", {}).get("upstream_edges", [])

    upstream_by_query: dict[tuple[str, str], tuple[str, str, str]] = {
        (edge["consumer"], edge["query"]):
            (edge["upstream_key"], edge["upstream_label"], edge["upstream_kind"])
        for edge in edges
    }
    kind_by_key: dict[str, str] = {
        edge["upstream_key"]: edge["upstream_kind"] for edge in edges
    }

    queries_dir: Path = config.path("queries_dir")
    profiles: list = []
    for source in manifest["sources"]:
        if not source.get("power_query", {}).get("query_count"):
            continue
        stem: str = Path(source["filename"]).stem
        section: Path = queries_dir / f"{stem}__{source['id']}" / "_Section1.m"
        if not section.is_file():
            logger.warning("no extracted M for %s (expected %s)", source["relative_path"], section)
            continue
        for query in parse_queries(section.read_text(encoding="utf-8")):
            profile = profile_query(source["relative_path"], query.name, query.source)
            key, label, _kind = upstream_by_query.get(
                (source["relative_path"], query.name), ("", "", "")
            )
            profile.upstream_key, profile.upstream_label = key, label
            profiles.append(profile)
    return profiles, kind_by_key, manifest


def command_schemas(args: argparse.Namespace) -> int:
    """Recover each export's schema from the committed M and write it to disk."""
    from .analyze.export_schema import build_export_specs
    from .extract.schema_store import SourceColumn, SourceSchema, dump_schemas, preserved_annotations

    config: Config = load_config(args.config)
    _configure_logging(args.verbose, None)

    profiles, kind_by_key, _manifest = _load_query_profiles(config)
    specs = build_export_specs(profiles, kind_by_key)
    if not specs:
        logger.error("no declared column types found in the committed M source")
        return EXIT_BLOCKED

    target: Path = config.repo_root / config.extract.schema_file
    keep = preserved_annotations(target)

    schemas = [
        SourceSchema(
            slug=spec.slug, key=spec.key, label=spec.label, kind=spec.kind,
            partial=spec.partial, declared_by=spec.declared_by,
            entity=keep.get(spec.slug, {}).get("entity", ""),
            notes=keep.get(spec.slug, {}).get("notes", ""),
            columns=[
                SourceColumn(name=c.name, dtype=c.pandas_dtype, m_type=c.m_type,
                             type_conflict=c.conflicting_types)
                for c in spec.columns
            ],
        )
        for spec in specs
    ]
    dump_schemas(schemas, target)
    logger.info("wrote %s (%d sources)", target, len(schemas))

    for spec in specs:
        flags: list[str] = []
        if spec.partial:
            flags.append("PARTIAL")
        if spec.conflicts:
            flags.append(f"{len(spec.conflicts)} type conflict(s)")
        logger.info("  %-24s %3d columns  %s%s", spec.slug, len(spec.columns),
                    spec.label[:52], f"  [{', '.join(flags)}]" if flags else "")
    conflicted = [c.name for spec in specs for c in spec.conflicts]
    if conflicted:
        logger.warning("columns whose declared type differs between queries: %s", conflicted)
    return EXIT_OK



def command_extract(args: argparse.Namespace) -> int:
    """Phase 2: fetch, read, validate and land the canonical layer.

    Each export is read once, transformed once, and written to the warehouse.
    A blocking gate failure stops the publish and exits non-zero with the
    exception list in the log.
    """
    import pandas as pd

    from .extract import gates as gate_module
    from .extract.entities import BUILDERS, UNMODELLED, EntityBuild
    from .extract.fetch import FetchError, fetch
    from .extract.readers import ReadError, read_export
    from .extract.schema_store import load_schemas
    from .extract.warehouse import attach_lineage, write_sqlite, write_table

    config: Config = load_config(args.config)
    run_stamp: str = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    run_dir: Path = config.path("runs_dir") / run_stamp
    _configure_logging(args.verbose, run_dir / "extract.log")

    schema_path: Path = config.repo_root / config.extract.schema_file
    try:
        schemas = load_schemas(schema_path)
    except (FileNotFoundError, ValueError) as exc:
        logger.error("%s", exc)
        return EXIT_BLOCKED

    wanted: Optional[set[str]] = set(args.source) if args.source else None
    warehouse_dir: Path = config.path("warehouse_dir")
    extracted_at: datetime = datetime.now(timezone.utc)

    tables: dict[str, "pd.DataFrame"] = {}
    all_gate_results: list = []
    all_conflicts: list = []
    summary: list[dict[str, object]] = []
    fetch_failures: list[str] = []

    for schema in schemas:
        if wanted is not None and schema.slug not in wanted:
            continue
        if schema.slug not in BUILDERS:
            reason: str = UNMODELLED.get(schema.slug, "no canonical entity mapped yet")
            logger.info("skip %-22s not modelled: %s", schema.slug, reason)
            summary.append({"source": schema.slug, "status": "not_modelled", "reason": reason})
            continue

        try:
            fetched = fetch(schema, config.extract, config.repo_root, args.fetch_mode)
        except FetchError as exc:
            logger.error("%s", exc)
            fetch_failures.append(schema.slug)
            summary.append({"source": schema.slug, "status": "fetch_failed", "reason": str(exc)})
            continue

        try:
            raw, read_report = read_export(fetched, schema, config.canonical)
        except ReadError as exc:
            logger.error("%s", exc)
            summary.append({"source": schema.slug, "status": "read_failed", "reason": str(exc)})
            continue

        logger.info("read %-22s %7d rows x %2d cols from %s (%s)",
                    schema.slug, read_report.rows, read_report.columns,
                    fetched.origin, read_report.encoding)
        if read_report.unexpected_columns:
            logger.warning("  %s: export carries %d column(s) not in the schema: %s",
                           schema.slug, len(read_report.unexpected_columns),
                           read_report.unexpected_columns[:6])
        for column, failures in read_report.coercion_failures.items():
            logger.warning("  %s: %d value(s) in '%s' did not parse and are missing, not zero",
                           schema.slug, failures, column)

        build: EntityBuild = BUILDERS[schema.slug](raw, config)
        stamped = attach_lineage(
            build.frame, source_slug=schema.slug, source_origin=fetched.origin,
            extracted_at=extracted_at,
        )
        tables[build.entity] = stamped
        all_gate_results.extend(build.results)
        if not build.conflicts.empty:
            all_conflicts.append(build.conflicts)

        blocked_gates = [r for r in build.results if r.blocks_publish]
        logger.info("built %-22s -> %-18s %7d rows  grain=%s  %s",
                    schema.slug, build.entity, build.rows, " x ".join(build.grain),
                    "BLOCKED" if blocked_gates else "gates ok")
        for note in build.notes:
            logger.info("  note: %s", note)
        for result in build.results:
            level = logger.error if result.blocks_publish else (
                logger.warning if not result.passed else logger.debug
            )
            level("  gate %-18s %-8s %s", result.gate,
                  "PASS" if result.passed else result.severity.upper(), result.detail)

        summary.append({
            "source": schema.slug, "status": "built", "entity": build.entity,
            "rows": build.rows, "grain": build.grain,
            "blocked": bool(blocked_gates),
        })

    if not tables:
        logger.error("nothing was built. Drop the export CSVs in %s (named <slug>.csv) or "
                     "run with --fetch-mode url and the per-source tokens set.",
                     config.repo_root / config.extract.drop_dir)
        return EXIT_BLOCKED

    exceptions = pd.concat(
        [gate_module.to_exception_rows(all_gate_results, "extract")]
        + ([pd.concat(all_conflicts, ignore_index=True)] if all_conflicts else []),
        ignore_index=True,
    )
    tables["exceptions"] = exceptions

    blocking = [r for r in all_gate_results if r.blocks_publish]
    if blocking and not args.force:
        logger.error("%d blocking gate failure(s); nothing written. Re-run with --force to "
                     "land the data anyway with the exceptions table.", len(blocking))
        for result in blocking:
            logger.error("  %s / %s: %s", result.entity, result.gate, result.detail)
        return EXIT_BLOCKED

    written = [write_table(frame, name, warehouse_dir) for name, frame in tables.items()]
    sqlite_path = write_sqlite(tables, warehouse_dir)
    for report in written:
        logger.info("wrote %-18s %7d rows x %2d cols -> %s",
                    report.table, report.rows, report.columns, report.parquet_path)
    logger.info("wrote sqlite mirror -> %s", sqlite_path)

    run_dir.mkdir(parents=True, exist_ok=True)
    (run_dir / "manifest.json").write_text(json.dumps({
        "command": "extract",
        "extracted_at": extracted_at.isoformat(timespec="seconds"),
        "fetch_mode": args.fetch_mode or config.extract.fetch_mode,
        "sources": summary,
        "tables": [{"table": r.table, "rows": r.rows, "columns": r.columns} for r in written],
        "blocking_gate_failures": [
            {"entity": r.entity, "gate": r.gate, "detail": r.detail} for r in blocking
        ],
        "fetch_failures": fetch_failures,
    }, indent=2), encoding="utf-8")

    if blocking:
        logger.warning("landed with %d blocking failure(s) because --force was given; see the "
                       "exceptions table", len(blocking))
        return EXIT_BLOCKED
    return EXIT_OK



def command_formulas(args: argparse.Namespace) -> int:
    """Map the calculation engines: what each formula column computes."""
    from .analyze.formula_map import map_workbook_formulas
    from .analyze.formula_report import write_reports
    from .discovery.ooxml import NotAnOoxmlPackage

    config: Config = load_config(args.config)
    _configure_logging(args.verbose, None)

    targets: list[Path] = []
    for root in args.root or []:
        base: Path = Path(root).expanduser()
        if base.is_file():
            targets.append(base)
            continue
        targets.extend(sorted(p for p in base.rglob("*.xls*") if not p.name.startswith("~$")))
    if not targets:
        manifest_path: Path = config.path("discovery_dir") / "manifest.json"
        if not manifest_path.is_file():
            logger.error("pass --root PATH, or run `python -m sc.cli discover` first")
            return EXIT_BLOCKED
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        targets = [Path(s["path"]) for s in manifest["sources"]
                   if s["kind"].startswith("workbook") and Path(s["path"]).is_file()]

    if not targets:
        logger.error("no readable workbooks found; the manifest's paths may be from another machine")
        return EXIT_BLOCKED

    mapped: list = []
    for target in targets:
        try:
            sheets = map_workbook_formulas(str(target), min_formulas=args.min_formulas,
                                           sheet_names=args.sheet or None)
        except NotAnOoxmlPackage as exc:
            logger.warning("skip %s: %s", target.name, exc)
            continue
        logger.info("%-46s %2d engine sheet(s), %8d formulas",
                    target.name, len(sheets), sum(s.formula_count for s in sheets))
        mapped.extend(sheets)

    if not mapped:
        logger.error("no sheet reached the %d-formula threshold; lower it with --min-formulas",
                     args.min_formulas)
        return EXIT_BLOCKED

    mapped.sort(key=lambda sheet: -sheet.formula_count)
    generated_at: str = datetime.now(timezone.utc).isoformat(timespec="seconds")
    for path in write_reports(mapped, config.path("discovery_dir"), generated_at, args.min_formulas):
        logger.info("wrote %s", path)

    suspicious = [(s, c) for s in mapped for c in s.suspicious_columns]
    logger.info("%d sheets, %d formula columns, %d column(s) with mixed formulas",
                len(mapped), sum(len(s.columns) for s in mapped), len(suspicious))
    for sheet, column in sorted(suspicious, key=lambda pair: pair[1].consistency)[:8]:
        logger.warning("MIXED %s!%s (%s): %d cells, %d patterns, %s consistent",
                       sheet.sheet, column.column_letter, column.header or "no header",
                       column.total, column.distinct_patterns, column.consistency_label)
    return EXIT_OK


def command_run(args: argparse.Namespace) -> int:
    """Phase 2-6 pipeline. Gated until the canonical model is signed off."""
    load_config(args.config)          # fail fast on a broken config even when gated
    logger.error(
        "`run` is not implemented yet, on purpose.\n"
        "  Requested: refresh=%s build=%s publish=%s\n"
        "  Phase 0 (discover) is complete and runnable.\n"
        "  Phases 2-6 are blocked on two sign-offs from the build spec:\n"
        "    1. the discovery manifest — which files are sources vs downstream copies\n"
        "    2. SCHEMA.md — the canonical entities, keys, grain and units\n"
        "  Writing readers before those are agreed produces code that gets rewritten.\n"
        "  Run `python -m sc.cli discover` first, then review discovery/DISCOVERY.md.",
        args.refresh, args.build, args.publish,
    )
    return EXIT_BLOCKED


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m sc.cli",
        description="Predator Group supply chain reference & update engine.",
    )
    parser.add_argument("--config", type=Path, default=None,
                        help="path to config.yaml (default: sc/config.yaml)")
    parser.add_argument("-v", "--verbose", action="store_true", help="debug logging")
    subparsers = parser.add_subparsers(dest="command", required=True)

    discover = subparsers.add_parser("discover", help="Phase 0 read-only sweep of the workbook estate")
    discover.add_argument("--root", action="append", default=[],
                          help="override discovery roots (repeatable); config roots are ignored")
    discover.add_argument("--no-queries", action="store_true",
                          help="skip writing queries/*.m")
    discover.add_argument("--json-only", action="store_true",
                          help="write manifest.json but not DISCOVERY.md")
    discover.set_defaults(func=command_discover)

    schemas_parser = subparsers.add_parser(
        "schemas",
        help="recover each export's schema from the committed M source",
    )
    schemas_parser.set_defaults(func=command_schemas)

    reconcile_parser = subparsers.add_parser(
        "reconcile",
        help="compare queries that read the same upstream export (needs a manifest)",
    )
    reconcile_parser.set_defaults(func=command_reconcile)

    formulas_parser = subparsers.add_parser(
        "formulas", help="map the calculation engines (what each formula column computes)"
    )
    formulas_parser.add_argument("--root", action="append", default=[],
                                 help="workbook or folder to map (repeatable); "
                                      "defaults to the manifest's workbooks")
    formulas_parser.add_argument("--sheet", action="append", default=[],
                                 help="limit to these sheet names (repeatable)")
    formulas_parser.add_argument("--min-formulas", type=int, default=1000,
                                 help="skip sheets below this formula count (default 1000)")
    formulas_parser.set_defaults(func=command_formulas)

    extract_parser = subparsers.add_parser(
        "extract", help="Phase 2: fetch, read, validate and land the canonical layer"
    )
    extract_parser.add_argument("--source", action="append", default=[],
                                help="limit to these source slugs (repeatable)")
    extract_parser.add_argument("--fetch-mode", choices=["local", "url"], default=None,
                                help="override extract.fetch_mode from config")
    extract_parser.add_argument("--force", action="store_true",
                                help="land the data even when a blocking gate fails")
    extract_parser.set_defaults(func=command_extract)

    run = subparsers.add_parser("run", help="full pipeline (gated — see Phase 0 sign-off)")
    run.add_argument("--refresh", default="none")
    run.add_argument("--build", default="")
    run.add_argument("--publish", default="")
    run.add_argument("--dry-run", action="store_true")
    run.set_defaults(func=command_run)
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args: argparse.Namespace = build_parser().parse_args(argv)
    if not logging.getLogger().handlers:
        _configure_logging(getattr(args, "verbose", False), None)
    try:
        return int(args.func(args))
    except ConfigError as exc:
        logger.error("configuration error: %s", exc)
        return EXIT_ERROR
    except KeyboardInterrupt:
        logger.error("interrupted")
        return EXIT_ERROR


if __name__ == "__main__":
    raise SystemExit(main())
