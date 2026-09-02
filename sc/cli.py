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
