"""Phase 0 orchestrator: sweep -> probe -> classify -> manifest.

Strictly read-only against every source. Workbooks are opened as zip archives
in read mode, macros are never executed, and nothing under a configured root is
written to. The only writes are into the repo's own ``discovery/`` and
``queries/`` directories.
"""

from __future__ import annotations

import fnmatch
import hashlib
import os
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

from ..configuration import Config
from . import classify
from .csv_probe import CsvProbe, probe_csv
from .datamashup import MashupProbe, probe_mashup
from .ooxml import NotAnOoxmlPackage, open_workbook
from .vba import VbaProbe, probe_vba
from .workbook_probe import WorkbookProbe, probe_workbook
from .xlsb import BiffError, XlsbSheet, probe_xlsb

KIND_WORKBOOK_XML: str = "workbook_ooxml"
KIND_WORKBOOK_XLSB: str = "workbook_xlsb"
KIND_CSV: str = "csv"
KIND_COMPANION: str = "companion_asset"

STATUS_OK: str = "ok"
STATUS_PARTIAL: str = "partial"
STATUS_FAILED: str = "failed"


@dataclass
class SheetRecord:
    name: str
    state: str
    rows: int
    data_rows: int
    columns: int
    formulas: int
    header_row: Optional[int]
    header_confidence: float
    headers: List[str] = field(default_factory=list)
    tables: List[Dict[str, str]] = field(default_factory=list)
    rows_capped: bool = False
    key_fields: Dict[str, str] = field(default_factory=dict)
    error: Optional[str] = None


@dataclass
class SourceRecord:
    """One discovered source, fully described. This is the manifest row."""

    id: str
    path: str
    relative_path: str
    filename: str
    extension: str
    kind: str
    root: str
    size_bytes: int
    modified_at: str
    sample_hash: str
    domain: str = "unclassified"
    domain_keywords: List[str] = field(default_factory=list)
    seed_ids: List[str] = field(default_factory=list)
    role: str = classify.ROLE_UNKNOWN
    role_confidence: str = "low"
    role_reasons: List[str] = field(default_factory=list)
    refresh_mechanism: str = ""
    grain: str = ""
    primary_sheet: str = ""
    key_fields: Dict[str, str] = field(default_factory=dict)
    sheets: List[SheetRecord] = field(default_factory=list)
    power_query: Dict[str, Any] = field(default_factory=dict)
    connections: List[Dict[str, str]] = field(default_factory=list)
    external_links: List[str] = field(default_factory=list)
    external_locations: List[Dict[str, str]] = field(default_factory=list)
    vba: Dict[str, Any] = field(default_factory=dict)
    risk_score: int = 0
    risk_band: str = "low"
    risk_findings: List[str] = field(default_factory=list)
    probe_status: str = STATUS_OK
    errors: List[str] = field(default_factory=list)

    @property
    def total_rows(self) -> int:
        return sum(sheet.rows for sheet in self.sheets)


@dataclass
class DiscoveryResult:
    """Everything Phase 0 produced, ready to serialize."""

    started_at: str
    finished_at: str
    duration_seconds: float
    config_path: str
    roots: List[Dict[str, Any]] = field(default_factory=list)
    sources: List[SourceRecord] = field(default_factory=list)
    overlaps: List[Dict[str, Any]] = field(default_factory=list)
    seed_status: List[Dict[str, Any]] = field(default_factory=list)
    query_files_written: List[str] = field(default_factory=list)
    query_diffs: List[Dict[str, str]] = field(default_factory=list)
    skipped: List[Dict[str, str]] = field(default_factory=list)

    @property
    def counts(self) -> Dict[str, int]:
        by_kind: Dict[str, int] = {}
        for source in self.sources:
            by_kind[source.kind] = by_kind.get(source.kind, 0) + 1
        return {
            "sources": len(self.sources),
            "failed_probes": sum(1 for s in self.sources if s.probe_status == STATUS_FAILED),
            "partial_probes": sum(1 for s in self.sources if s.probe_status == STATUS_PARTIAL),
            "power_query_workbooks": sum(1 for s in self.sources if s.power_query.get("query_count")),
            "vba_workbooks": sum(1 for s in self.sources if s.vba.get("present")),
            "duplicate_truth_pairs": len(self.overlaps),
            "seeds_missing": sum(1 for s in self.seed_status if not s["found"]),
            "skipped": len(self.skipped),
            **{f"kind_{k}": v for k, v in sorted(by_kind.items())},
        }


# --------------------------------------------------------------------------
# Filesystem sweep
# --------------------------------------------------------------------------

def _should_skip_dir(name: str, skip_dirs: Sequence[str]) -> bool:
    lowered: str = name.lower()
    return any(lowered == skip.lower() for skip in skip_dirs)


def _should_skip_file(name: str, skip_globs: Sequence[str]) -> bool:
    lowered: str = name.lower()
    return any(fnmatch.fnmatch(lowered, pattern.lower()) for pattern in skip_globs)


def sweep(config: Config) -> Tuple[List[Tuple[Path, Path]], List[Dict[str, str]], List[Dict[str, Any]]]:
    """Walk the configured roots.

    Returns ``(candidates, skipped, root_status)`` where each candidate is
    ``(root, absolute_path)``. Unreachable roots are reported, not silently
    dropped — an empty manifest caused by a renamed OneDrive folder must look
    like a failure, not like a clean estate.
    """
    disc = config.discovery
    wanted: Set[str] = set(disc.extensions) | set(disc.companion_extensions)
    candidates: List[Tuple[Path, Path]] = []
    skipped: List[Dict[str, str]] = []
    root_status: List[Dict[str, Any]] = []

    for root in disc.resolved_roots():
        if not root.is_dir():
            root_status.append({"root": str(root), "reachable": False, "files": 0,
                                "note": "path does not exist or is not readable from this machine"})
            continue

        found: int = 0
        for current_dir, subdirs, filenames in os.walk(root, onerror=lambda e: skipped.append(
            {"path": getattr(e, "filename", str(root)), "reason": f"walk error: {e}"}
        )):
            subdirs[:] = [d for d in subdirs if not _should_skip_dir(d, disc.skip_dirs)]
            for filename in filenames:
                if _should_skip_file(filename, disc.skip_file_globs):
                    continue
                extension: str = Path(filename).suffix.lower()
                if extension not in wanted:
                    continue
                candidates.append((root, Path(current_dir) / filename))
                found += 1
        root_status.append({"root": str(root), "reachable": True, "files": found, "note": ""})

    return candidates, skipped, root_status


def _sample_hash(path: Path, sample_bytes: int) -> str:
    """Hash of head+tail bytes plus size — cheap duplicate detection on huge files."""
    digest = hashlib.sha256()
    try:
        size: int = path.stat().st_size
        digest.update(str(size).encode())
        with path.open("rb") as handle:
            digest.update(handle.read(sample_bytes))
            if size > sample_bytes * 2:
                handle.seek(-sample_bytes, os.SEEK_END)
                digest.update(handle.read(sample_bytes))
    except OSError as exc:
        return f"unhashable:{exc.errno}"
    return digest.hexdigest()[:16]


def _stable_id(relative_path: str) -> str:
    return hashlib.sha1(relative_path.lower().encode("utf-8")).hexdigest()[:12]


# --------------------------------------------------------------------------
# Per-file probing
# --------------------------------------------------------------------------

def _base_record(root: Path, path: Path, kind: str, sample_bytes: int) -> SourceRecord:
    stat = path.stat()
    try:
        relative: str = str(path.relative_to(root))
    except ValueError:
        relative = str(path)
    return SourceRecord(
        id=_stable_id(relative),
        path=str(path),
        relative_path=relative,
        filename=path.name,
        extension=path.suffix.lower(),
        kind=kind,
        root=str(root),
        size_bytes=stat.st_size,
        modified_at=datetime.fromtimestamp(stat.st_mtime, tz=timezone.utc).isoformat(timespec="seconds"),
        sample_hash=_sample_hash(path, sample_bytes),
    )


def _kind_for(path: Path, config: Config) -> str:
    extension: str = path.suffix.lower()
    if extension == ".csv":
        return KIND_CSV
    if extension == ".xlsb":
        return KIND_WORKBOOK_XLSB
    if extension in config.discovery.extensions:
        return KIND_WORKBOOK_XML
    return KIND_COMPANION


def _probe_ooxml(record: SourceRecord, config: Config) -> Tuple[WorkbookProbe, MashupProbe, VbaProbe]:
    """Structure + Power Query + VBA in a single package open."""
    probe: WorkbookProbe = probe_workbook(record.path, row_cap=config.discovery.row_count_cap)
    with open_workbook(record.path) as zf:
        mashup: MashupProbe = probe_mashup(zf)
        vba: VbaProbe = probe_vba(zf)
    return probe, mashup, vba


def probe_source(record: SourceRecord, config: Config) -> Tuple[SourceRecord, Optional[MashupProbe]]:
    """Fill a record from the file on disk. Never raises; failures land in ``errors``."""
    disc = config.discovery
    mashup: Optional[MashupProbe] = None
    headers_pool: List[str] = []
    haystacks: List[str] = []

    if record.size_bytes > disc.max_file_bytes:
        record.errors.append(
            f"{record.size_bytes / 1048576:.0f} MB exceeds max_file_mb={disc.max_file_mb}; "
            "structure probed, rows not counted"
        )

    try:
        if record.kind == KIND_CSV:
            csv_result: CsvProbe = probe_csv(record.path, disc.csv_sniff_bytes, disc.row_count_cap)
            record.sheets = [
                SheetRecord(
                    name="(flat file)", state="visible", rows=csv_result.row_count,
                    data_rows=csv_result.row_count, columns=csv_result.column_count,
                    formulas=0, header_row=1, header_confidence=1.0 if csv_result.headers else 0.0,
                    headers=csv_result.headers, rows_capped=csv_result.row_count_capped,
                    error=csv_result.error,
                )
            ]
            record.power_query = {"found": False, "query_count": 0}
            record.vba = {"present": False}
            headers_pool = list(csv_result.headers)
            haystacks = list(csv_result.headers)
            if csv_result.error:
                record.errors.append(csv_result.error)
            record.connections = [
                {"name": "flat_file", "type": "csv",
                 "connection_string": f"encoding={csv_result.encoding}; delimiter={csv_result.delimiter!r}",
                 "command": "", "description": "", "refresh_on_load": "0",
                 "background_refresh": "0", "source_file": record.path}
            ]
            if csv_result.ragged_rows:
                record.errors.append(
                    f"{csv_result.ragged_rows} row(s) have a column count that differs from the header"
                )

        elif record.kind == KIND_WORKBOOK_XLSB:
            sheets: List[XlsbSheet] = probe_xlsb(record.path, disc.row_count_cap)
            for sheet in sheets:
                record.sheets.append(
                    SheetRecord(
                        name=sheet.name, state=sheet.state, rows=sheet.row_count,
                        data_rows=max(sheet.row_count - (sheet.header_row or 0), 0),
                        columns=sheet.last_column + 1, formulas=sheet.formula_count,
                        header_row=sheet.header_row, header_confidence=1.0 if sheet.headers else 0.0,
                        headers=sheet.headers, error=sheet.error,
                    )
                )
                headers_pool.extend(sheet.headers)
                haystacks.append(sheet.name)
                if sheet.error:
                    record.errors.append(sheet.error)
            # .xlsb still stores Power Query and VBA as package parts.
            with open_workbook(record.path) as zf:
                mashup = probe_mashup(zf)
                vba_probe: VbaProbe = probe_vba(zf)
            record.vba = {
                "present": vba_probe.present, "project_name": vba_probe.project_name,
                "protected": vba_probe.protected, "modules": vba_probe.module_names,
                "module_count": vba_probe.count,
            }

        elif record.kind == KIND_WORKBOOK_XML:
            wb_probe: WorkbookProbe
            wb_probe, mashup, vba_probe = _probe_ooxml(record, config)
            for sheet in wb_probe.sheets:
                record.sheets.append(
                    SheetRecord(
                        name=sheet.name, state=sheet.state, rows=sheet.row_count,
                        data_rows=sheet.data_row_estimate, columns=sheet.max_column_index + 1,
                        formulas=sheet.formula_count, header_row=sheet.header_row,
                        header_confidence=sheet.header_confidence, headers=sheet.headers,
                        tables=sheet.tables, rows_capped=sheet.row_count_capped, error=sheet.error,
                    )
                )
                headers_pool.extend(sheet.headers)
                haystacks.append(sheet.name)
                haystacks.extend(table["name"] for table in sheet.tables)
            record.connections = wb_probe.connections
            record.external_links = wb_probe.external_links
            record.errors.extend(wb_probe.errors)
            record.vba = {
                "present": vba_probe.present, "project_name": vba_probe.project_name,
                "protected": vba_probe.protected, "modules": vba_probe.module_names,
                "module_count": vba_probe.count,
            }
            if wb_probe.shared_strings_truncated:
                record.errors.append(
                    "shared string table exceeded the probe cap; some headers may read blank"
                )
            haystacks.extend(name["name"] for name in wb_probe.defined_names)

        else:  # companion assets: .py/.ps1/.js/.m/.json/.pbix
            record.power_query = {"found": False, "query_count": 0}
            record.vba = {"present": False}
            record.refresh_mechanism = "n/a (companion asset — code or model, not a dataset)"

    except (NotAnOoxmlPackage, BiffError) as exc:
        record.errors.append(str(exc))
        record.probe_status = STATUS_FAILED
    except (OSError, PermissionError) as exc:
        record.errors.append(f"unreadable on this machine: {exc}")
        record.probe_status = STATUS_FAILED
    except Exception as exc:  # a probe bug must not abort the whole sweep
        record.errors.append(f"unhandled probe error ({type(exc).__name__}): {exc}")
        record.probe_status = STATUS_FAILED

    if mashup is not None:
        record.power_query = {
            "found": mashup.found,
            "part": mashup.part,
            "query_count": len(mashup.queries),
            "error": mashup.error,
            "queries": [
                {"name": q.name, "lines": q.line_count, "shared": q.is_shared,
                 "metadata": q.metadata, "sources": q.sources}
                for q in mashup.queries
            ],
        }
        haystacks.extend(mashup.query_names)
        if mashup.error:
            record.errors.append(f"power query: {mashup.error}")

    # Every external location this file depends on, from any mechanism.
    locations: Dict[str, Dict[str, str]] = {}
    for query in record.power_query.get("queries", []):
        for source in query["sources"]:
            if source["kind"] == "query_ref":
                continue
            locations.setdefault(source["location"], {**source, "via": f"PQ:{query['name']}"})
    for connection in record.connections:
        for value in (connection.get("source_file", ""), connection.get("connection_string", "")):
            if value and ("\\" in value or value.lower().startswith("http")):
                locations.setdefault(value, {"kind": "connection", "via": connection.get("name", ""),
                                             "location": value})
    for link in record.external_links:
        locations.setdefault(link, {"kind": "workbook_link", "via": "formula link", "location": link})
    record.external_locations = list(locations.values())

    # ---- classification ----
    record.domain, record.domain_keywords = classify.classify_domain(
        record.relative_path, haystacks, disc.domains
    )
    record.seed_ids = classify.match_seeds(record.filename, disc.seeds)

    # A seed match is explicit configuration, so it outranks keyword inference.
    seed_domains: List[str] = [
        seed.domain for seed in disc.seeds if seed.id in record.seed_ids
    ]
    if seed_domains and seed_domains[0] != record.domain:
        record.domain_keywords.append(
            f"domain set from seed match '{record.seed_ids[0]}' "
            f"(keyword inference said '{record.domain}')"
        )
        record.domain = seed_domains[0]

    # Keys are detected per sheet. Pooling headers across every tab lets a hidden
    # lookup sheet contribute a spurious key to the whole workbook's grain.
    for sheet in record.sheets:
        sheet.key_fields = classify.detect_key_fields(sheet.headers)
    primary: Optional[SheetRecord] = max(
        (s for s in record.sheets if s.state == "visible" and s.key_fields),
        key=lambda s: (len(s.key_fields), s.data_rows),
        default=None,
    )
    if primary is None:
        primary = max(record.sheets, key=lambda s: s.data_rows, default=None)
    if primary is not None:
        record.primary_sheet = primary.name
        record.key_fields = primary.key_fields
    record.grain = classify.infer_grain(record.key_fields)

    verdict = classify.classify_role(
        query_external_sources=sum(
            1 for location in record.external_locations if location["kind"] != "query_ref"
        ),
        query_count=int(record.power_query.get("query_count") or 0),
        external_link_count=len(record.external_links),
        connection_count=len([c for c in record.connections if c.get("type") != "csv"]),
        formula_count=sum(sheet.formulas for sheet in record.sheets),
        total_rows=record.total_rows,
        has_vba=bool(record.vba.get("present")),
        is_csv=record.kind == KIND_CSV,
        heavy_formula_count=disc.role_heuristics["heavy_formula_count"],
        heavy_formulas_per_row=disc.role_heuristics["heavy_formulas_per_row"],
    )
    record.role, record.role_reasons, record.role_confidence = (
        verdict.role, verdict.reasons, verdict.confidence
    )
    if record.kind != KIND_COMPANION:
        record.refresh_mechanism = classify.refresh_mechanism(
            query_count=int(record.power_query.get("query_count") or 0),
            connection_count=len([c for c in record.connections if c.get("type") != "csv"]),
            external_link_count=len(record.external_links),
            has_vba=bool(record.vba.get("present")),
            is_csv=record.kind == KIND_CSV,
        )

    best_confidence: float = max(
        (sheet.header_confidence for sheet in record.sheets), default=0.0
    )
    risk = classify.assess_dependency_risk(
        locations=[location["location"] for location in record.external_locations],
        has_vba=bool(record.vba.get("present")),
        vba_protected=bool(record.vba.get("protected")),
        hidden_sheet_count=sum(1 for sheet in record.sheets if sheet.state != "visible"),
        row_count_capped=any(sheet.rows_capped for sheet in record.sheets),
        header_confidence=best_confidence,
        probe_errors=record.errors,
    )
    record.risk_score, record.risk_band, record.risk_findings = risk.score, risk.band, risk.findings

    if record.probe_status != STATUS_FAILED:
        record.probe_status = STATUS_PARTIAL if record.errors else STATUS_OK
    return record, mashup


# --------------------------------------------------------------------------
# Query version control
# --------------------------------------------------------------------------

def write_query_files(
    record: SourceRecord, mashup: Optional[MashupProbe], queries_dir: Path
) -> Tuple[List[str], List[Dict[str, str]]]:
    """Write each M query to ``queries/<workbook>/<query>.m`` and diff against last run.

    A changed query is reported so a pipeline break traces to a query edit
    instead of being debugged from scratch.
    """
    if mashup is None or not mashup.queries:
        return [], []

    stem: str = Path(record.filename).stem
    target_dir: Path = queries_dir / f"{stem}__{record.id}"
    target_dir.mkdir(parents=True, exist_ok=True)

    written: List[str] = []
    diffs: List[Dict[str, str]] = []
    for query in mashup.queries:
        target: Path = target_dir / query.safe_filename()
        new_text: str = query.source.rstrip() + "\n"
        if target.exists():
            previous: str = target.read_text(encoding="utf-8")
            if previous != new_text:
                diffs.append({
                    "workbook": record.relative_path,
                    "query": query.name,
                    "file": str(target),
                    "change": "M source changed since the last discovery run",
                    "previous_lines": str(previous.count("\n")),
                    "current_lines": str(new_text.count("\n")),
                })
        target.write_text(new_text, encoding="utf-8")
        written.append(str(target))

    (target_dir / "_Section1.m").write_text(mashup.section_text, encoding="utf-8")
    written.append(str(target_dir / "_Section1.m"))
    return written, diffs


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------

def run_discovery(config: Config, write_queries: bool = True) -> DiscoveryResult:
    """Execute the full Phase 0 sweep."""
    start: float = time.time()
    started_at: str = datetime.now(timezone.utc).isoformat(timespec="seconds")

    candidates, skipped, root_status = sweep(config)
    queries_dir: Path = config.path("queries_dir")

    sources: List[SourceRecord] = []
    query_files: List[str] = []
    query_diffs: List[Dict[str, str]] = []

    for root, path in candidates:
        record: SourceRecord = _base_record(root, path, _kind_for(path, config), config.discovery.hash_sample_bytes)
        record, mashup = probe_source(record, config)
        sources.append(record)
        if write_queries:
            written, diffs = write_query_files(record, mashup, queries_dir)
            query_files.extend(written)
            query_diffs.extend(diffs)

    # ---- duplicate-truth detection across files ----
    entries: List[Tuple[str, str, str, int, Set[str], Dict[str, str]]] = []
    for record in sources:
        for sheet in record.sheets:
            if not sheet.headers:
                continue
            signature: Set[str] = classify.header_signature(sheet.headers)
            header_map: Dict[str, str] = {classify.normalize_header(h): h for h in sheet.headers}
            entries.append((record.relative_path, sheet.name, record.domain, sheet.rows, signature, header_map))

    overlaps = classify.find_header_overlaps(
        entries,
        config.discovery.header_overlap_threshold,
        config.discovery.header_overlap_min_columns,
    )

    # ---- seed confirmation ----
    seed_status: List[Dict[str, Any]] = []
    for seed in config.discovery.seeds:
        matches: List[SourceRecord] = [r for r in sources if seed.id in r.seed_ids]
        seed_status.append({
            "id": seed.id,
            "label": seed.label,
            "domain": seed.domain,
            "patterns": seed.patterns,
            "found": bool(matches),
            "matches": [{"path": m.relative_path, "modified_at": m.modified_at,
                         "rows": m.total_rows, "role": m.role} for m in matches],
        })

    finished: float = time.time()
    return DiscoveryResult(
        started_at=started_at,
        finished_at=datetime.now(timezone.utc).isoformat(timespec="seconds"),
        duration_seconds=round(finished - start, 2),
        config_path=str(config.config_path),
        roots=root_status,
        sources=sources,
        overlaps=[asdict(pair) | {"is_row_count_conflict": pair.is_row_count_conflict} for pair in overlaps],
        seed_status=seed_status,
        query_files_written=query_files,
        query_diffs=query_diffs,
        skipped=skipped,
    )
