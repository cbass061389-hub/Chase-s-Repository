"""Cross-workbook lineage: who reads whom, and where the same upstream forks.

The single most useful thing Phase 0 can produce for a consolidation is not a
list of files — it is the dependency graph between them. Three failure modes
only become visible at graph level:

**Circular dependency.** Workbook A refreshes from B while B refreshes from A.
Refresh order is then undefined: whichever runs second reads the other's stale
output, and the numbers move depending on the order somebody happened to click.

**Forked upstream.** One export feeding several workbooks, each applying its own
transformation. This is the *root cause* of two workbooks disagreeing — they are
not disagreeing about the data, they are disagreeing about the logic. Fixing the
copies without consolidating the logic fixes nothing.

**Version skew.** One workbook reading two different files that are clearly
versions of the same thing ("... HIE .xlsm" and "... HIE - Updated.xlsm"). Some
of its data is then sourced from a superseded file.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import PurePath, PureWindowsPath
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

#: Trailing words that mark a filename as a variant of the same asset.
VERSION_WORDS: Tuple[str, ...] = (
    "updated", "update", "new", "final", "latest", "current", "copy", "draft",
    "rev", "revised", "old", "backup", "bak", "v2", "v3", "temp", "working",
)


@dataclass(frozen=True)
class Upstream:
    """A canonical identity for something a query reads from."""

    key: str
    kind: str
    label: str

    @property
    def is_workbook_file(self) -> bool:
        return self.kind == "file"


@dataclass
class Edge:
    """One workbook reading one upstream, via one named query."""

    consumer_id: str
    consumer_path: str
    query: str
    upstream: Upstream
    location: str
    resolved_source_id: Optional[str] = None   # set when the upstream is a swept file


@dataclass
class LineageFindings:
    edges: List[Edge] = field(default_factory=list)
    cycles: List[List[str]] = field(default_factory=list)
    forked_upstreams: List[Dict[str, object]] = field(default_factory=list)
    version_skew: List[Dict[str, object]] = field(default_factory=list)
    unresolved_files: List[Dict[str, str]] = field(default_factory=list)


def canonical_upstream(kind: str, location: str) -> Upstream:
    """Reduce a raw location to a stable identity.

    Two queries hitting the same NetSuite File Cabinet export must land on the
    same node even though their URLs differ in ordering and token, so the media
    id is the identity and the token is irrelevant to it.
    """
    text: str = str(location).strip()

    media: Optional[re.Match[str]] = re.search(r"media\.nl\?[^\"]*?\bid=(\d+)", text, re.IGNORECASE)
    if media is not None:
        account: str = ""
        host = re.search(r"https?://(\d+)\.app\.netsuite\.com", text, re.IGNORECASE)
        if host is not None:
            account = host.group(1)
        return Upstream(
            key=f"netsuite:media:{media.group(1)}",
            kind="netsuite_file_cabinet",
            label=f"NetSuite File Cabinet export id={media.group(1)}"
                  + (f" (account {account})" if account else ""),
        )

    lowered: str = text.lower()
    if "sharepoint.com" in lowered:
        path: str = re.sub(r"[?#].*$", "", text)
        return Upstream(key=f"sharepoint:{path.lower()}", kind="sharepoint", label=path)

    if lowered.startswith(("http://", "https://")):
        return Upstream(key=f"web:{re.sub(r'[?#].*$', '', lowered)}", kind="web",
                        label=re.sub(r"[?#].*$", "", text))

    if "\\" in text or "/" in text:
        name: str = PureWindowsPath(text.replace("/", "\\")).name or text
        return Upstream(key=f"file:{name.lower()}", kind="file", label=name)

    return Upstream(key=f"{kind}:{lowered}", kind=kind, label=text)


def _normalize_filename(filename: str) -> str:
    """Filename reduced to alphanumerics, so separators cannot break a match."""
    return re.sub(r"[^a-z0-9]+", "", str(filename).lower())


def _version_stem(filename: str) -> str:
    """Normalize a filename to the asset it is a version of.

    "Shipment Request HIE .xlsm" and "Shipment Request HIE - Updated.xlsm"
    both reduce to "shipmentrequesthie".
    """
    stem: str = PurePath(filename).stem.lower()
    tokens: List[str] = [t for t in re.split(r"[^a-z0-9]+", stem) if t]
    while tokens and tokens[-1] in VERSION_WORDS:
        tokens.pop()
    # A trailing bare number is a version marker too ("... 2", "... 2026").
    while len(tokens) > 1 and tokens[-1].isdigit():
        tokens.pop()
    return "".join(tokens)


def build_lineage(
    consumers: Sequence[Tuple[str, str, str, Sequence[Tuple[str, str, str]]]],
) -> LineageFindings:
    """Build the graph and run every cross-file check.

    *consumers* is ``(source_id, path, filename, [(query, kind, location)])``.
    """
    findings: LineageFindings = LineageFindings()
    # Indexed by exact and normalized filename. A query written as
    # "Supply Chain Update Meeting Workbook.xlsm" must resolve to a swept file
    # named "Supply_Chain_Update_Meeting_Workbook.xlsm" — otherwise a real
    # circular dependency goes undetected because of a space versus underscore.
    by_filename: Dict[str, Tuple[str, str]] = {}
    for source_id, path, filename, _dependencies in consumers:
        by_filename.setdefault(filename.lower(), (source_id, path))
        by_filename.setdefault(_normalize_filename(filename), (source_id, path))

    for source_id, path, _filename, dependencies in consumers:
        for query, kind, location in dependencies:
            upstream: Upstream = canonical_upstream(kind, location)
            edge: Edge = Edge(
                consumer_id=source_id, consumer_path=path, query=query,
                upstream=upstream, location=location,
            )
            if upstream.is_workbook_file:
                match: Optional[Tuple[str, str]] = by_filename.get(
                    upstream.label.lower()
                ) or by_filename.get(_normalize_filename(upstream.label))
                if match is not None:
                    edge.resolved_source_id = match[0]
                else:
                    findings.unresolved_files.append({
                        "consumer": path, "query": query, "location": location,
                        "note": "referenced workbook was not in the sweep — add its folder to "
                                "discovery.roots or confirm the path is dead",
                    })
            findings.edges.append(edge)

    findings.cycles = _find_cycles(findings.edges)
    findings.forked_upstreams = _find_forked_upstreams(findings.edges)
    findings.version_skew = _find_version_skew(findings.edges)
    return findings


def _find_cycles(edges: Sequence[Edge]) -> List[List[str]]:
    """Every simple cycle in the workbook-to-workbook graph.

    The graph here is a handful of nodes, so an exhaustive DFS is the right
    trade: correct and obvious, with no cleverness to get wrong.
    """
    graph: Dict[str, Set[str]] = {}
    label: Dict[str, str] = {}
    for edge in edges:
        if edge.resolved_source_id is None:
            continue
        graph.setdefault(edge.consumer_id, set()).add(edge.resolved_source_id)
        label[edge.consumer_id] = edge.consumer_path

    cycles: List[List[str]] = []
    seen: Set[Tuple[str, ...]] = set()

    def walk(node: str, stack: List[str]) -> None:
        for neighbour in sorted(graph.get(node, ())):
            if neighbour in stack:
                cycle: List[str] = stack[stack.index(neighbour):] + [neighbour]
                fingerprint: Tuple[str, ...] = tuple(sorted(set(cycle)))
                if fingerprint not in seen:
                    seen.add(fingerprint)
                    cycles.append([label.get(n, n) for n in cycle])
                continue
            if len(stack) < 12:                       # guards against pathological graphs
                walk(neighbour, stack + [neighbour])

    for start in sorted(graph):
        walk(start, [start])
    return cycles


def _find_forked_upstreams(edges: Sequence[Edge]) -> List[Dict[str, object]]:
    """One upstream read by more than one workbook — the duplicate-truth root cause."""
    grouped: Dict[str, List[Edge]] = {}
    for edge in edges:
        grouped.setdefault(edge.upstream.key, []).append(edge)

    forked: List[Dict[str, object]] = []
    for key, group in grouped.items():
        consumers: Set[str] = {edge.consumer_path for edge in group}
        if len(consumers) < 2:
            continue
        forked.append({
            "upstream_key": key,
            "upstream_kind": group[0].upstream.kind,
            "upstream_label": group[0].upstream.label,
            "consumer_count": len(consumers),
            "consumers": [
                {"workbook": edge.consumer_path, "query": edge.query} for edge in group
            ],
        })
    forked.sort(key=lambda item: -int(item["consumer_count"]))
    return forked


def _find_version_skew(edges: Sequence[Edge]) -> List[Dict[str, object]]:
    """One workbook reading two files that are versions of the same asset."""
    per_consumer: Dict[str, Dict[str, Dict[str, List[str]]]] = {}
    for edge in edges:
        if not edge.upstream.is_workbook_file:
            continue
        stem: str = _version_stem(edge.upstream.label)
        if not stem:
            continue
        bucket = per_consumer.setdefault(edge.consumer_path, {}).setdefault(stem, {})
        bucket.setdefault(edge.upstream.label, []).append(edge.query)

    skew: List[Dict[str, object]] = []
    for consumer, stems in per_consumer.items():
        for stem, variants in stems.items():
            if len(variants) < 2:
                continue
            skew.append({
                "workbook": consumer,
                "asset": stem,
                "variants": [
                    {"file": name, "queries": queries} for name, queries in sorted(variants.items())
                ],
            })
    return skew
