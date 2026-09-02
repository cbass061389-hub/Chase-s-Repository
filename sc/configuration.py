"""Typed loader for ``sc/config.yaml``.

Every path, threshold and source list used by the engine arrives through this
module. Code never contains a literal path. Loading is strict: a missing or
malformed key raises immediately with the offending key name, because a
silently defaulted config produces a plausible-looking wrong answer.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List

import yaml

CONFIG_FILENAME: str = "config.yaml"


class ConfigError(RuntimeError):
    """Raised when the configuration file is missing or structurally invalid."""


def _require(mapping: Dict[str, Any], key: str, expected: type, where: str) -> Any:
    if key not in mapping:
        raise ConfigError(f"missing required key '{key}' in {where}")
    value: Any = mapping[key]
    if not isinstance(value, expected):
        raise ConfigError(
            f"key '{key}' in {where} must be {expected.__name__}, got {type(value).__name__}"
        )
    return value


def expand(raw: str) -> Path:
    """Expand ``${ENV_VAR}`` and ``~`` in a configured path string."""
    return Path(os.path.expandvars(os.path.expanduser(str(raw))))


@dataclass(frozen=True)
class SeedSpec:
    """A workbook the build spec expects to exist. Presence is verified, never assumed."""

    id: str
    domain: str
    label: str
    patterns: List[str]


@dataclass(frozen=True)
class DiscoveryConfig:
    roots: List[str]
    auto_detect_onedrive: bool
    extra_roots: List[str]
    extensions: List[str]
    skip_dirs: List[str]
    skip_file_globs: List[str]
    max_file_mb: int
    csv_sniff_bytes: int
    hash_sample_bytes: int
    row_count_cap: int
    header_overlap_threshold: float
    header_overlap_min_columns: int
    domains: Dict[str, List[str]]
    seeds: List[SeedSpec]
    companion_extensions: List[str]
    role_heuristics: Dict[str, float]
    redact_patterns: List[Dict[str, str]]
    bloat_min_rows: int
    bloat_value_ratio: float

    @property
    def max_file_bytes(self) -> int:
        return int(self.max_file_mb) * 1024 * 1024

    def resolved_roots(self) -> List[Path]:
        """Configured roots plus auto-detected OneDrive folders, deduplicated.

        Non-existent roots are returned as well; the sweep reports them as
        unreachable rather than skipping them quietly, so a renamed tenant
        folder surfaces as a finding instead of an empty manifest.
        """
        candidates: List[Path] = [expand(r) for r in list(self.roots) + list(self.extra_roots)]
        if self.auto_detect_onedrive:
            profile: str = os.environ.get("USERPROFILE") or os.environ.get("HOME") or ""
            if profile:
                candidates.extend(sorted(Path(profile).glob("OneDrive*")))
        seen: Dict[str, Path] = {}
        for candidate in candidates:
            seen.setdefault(str(candidate).rstrip("\\/").lower(), candidate)
        return list(seen.values())


@dataclass(frozen=True)
class Thresholds:
    overstock_months_supply: float
    dead_stock_forecast_units: float
    forecast_confidence_lookback_months: int
    snapshot_stale_days: int
    abc_basis: str
    abc_cutoffs: Dict[str, float]
    worst_flag_hierarchy: List[str]


@dataclass(frozen=True)
class CanonicalRules:
    """The single agreed rule per concept. See `canonical:` in config.yaml."""

    sku: Dict[str, Any]
    csv: Dict[str, Any]
    locations: Dict[str, Dict[str, Any]]
    measures: Dict[str, Dict[str, Any]]
    channels: Dict[str, str]

    def location_attribute(self, location: str, attribute: str) -> Any:
        """One attribute of one location, or None when the location is unmapped."""
        entry: Dict[str, Any] = self.locations.get(str(location).strip(), {})
        return entry.get(attribute)

    @property
    def known_locations(self) -> List[str]:
        return sorted(self.locations)


@dataclass(frozen=True)
class ExtractConfig:
    schema_file: str
    drop_dir: str
    fetch_mode: str
    token_env_prefix: str
    request_timeout_seconds: int
    gates: Dict[str, float]


@dataclass(frozen=True)
class Config:
    project: Dict[str, Any]
    discovery: DiscoveryConfig
    thresholds: Thresholds
    canonical: CanonicalRules
    extract: ExtractConfig
    repo_root: Path
    config_path: Path
    _paths: Dict[str, str] = field(default_factory=dict)

    def path(self, name: str) -> Path:
        """Resolve a named output path from ``paths:``, relative to the repo root."""
        if name not in self._paths:
            raise ConfigError(f"unknown path key '{name}'; known keys: {sorted(self._paths)}")
        configured: Path = expand(self._paths[name])
        return configured if configured.is_absolute() else self.repo_root / configured


def load_config(config_path: Path | str | None = None) -> Config:
    """Load and validate the engine configuration.

    The repo root is taken as the parent of the directory holding config.yaml,
    so output paths resolve identically from any working directory.
    """
    resolved: Path = (
        Path(config_path).expanduser().resolve()
        if config_path is not None
        else (Path(__file__).resolve().parent / CONFIG_FILENAME)
    )
    if not resolved.is_file():
        raise ConfigError(f"config file not found: {resolved}")

    with resolved.open("r", encoding="utf-8") as handle:
        raw: Any = yaml.safe_load(handle)
    if not isinstance(raw, dict):
        raise ConfigError(f"{resolved} did not parse to a mapping")

    disc_raw: Dict[str, Any] = _require(raw, "discovery", dict, "config root")
    seeds: List[SeedSpec] = []
    for index, entry in enumerate(_require(disc_raw, "seeds", list, "discovery")):
        where: str = f"discovery.seeds[{index}]"
        if not isinstance(entry, dict):
            raise ConfigError(f"{where} must be a mapping")
        seeds.append(
            SeedSpec(
                id=_require(entry, "id", str, where),
                domain=_require(entry, "domain", str, where),
                label=_require(entry, "label", str, where),
                patterns=list(_require(entry, "patterns", list, where)),
            )
        )

    discovery: DiscoveryConfig = DiscoveryConfig(
        roots=list(_require(disc_raw, "roots", list, "discovery")),
        auto_detect_onedrive=bool(disc_raw.get("auto_detect_onedrive", False)),
        extra_roots=list(disc_raw.get("extra_roots") or []),
        extensions=[str(e).lower() for e in _require(disc_raw, "extensions", list, "discovery")],
        skip_dirs=list(_require(disc_raw, "skip_dirs", list, "discovery")),
        skip_file_globs=list(_require(disc_raw, "skip_file_globs", list, "discovery")),
        max_file_mb=int(disc_raw.get("max_file_mb", 400)),
        csv_sniff_bytes=int(disc_raw.get("csv_sniff_bytes", 65536)),
        hash_sample_bytes=int(disc_raw.get("hash_sample_bytes", 1048576)),
        row_count_cap=int(disc_raw.get("row_count_cap", 2_000_000)),
        header_overlap_threshold=float(disc_raw.get("header_overlap_threshold", 0.70)),
        header_overlap_min_columns=int(disc_raw.get("header_overlap_min_columns", 4)),
        domains={str(k): [str(v).lower() for v in vals]
                 for k, vals in _require(disc_raw, "domains", dict, "discovery").items()},
        seeds=seeds,
        companion_extensions=[str(e).lower() for e in disc_raw.get("companion_extensions") or []],
        role_heuristics={
            str(k): float(v)
            for k, v in (_require(disc_raw, "role_heuristics", dict, "discovery")).items()
        },
        redact_patterns=[
            {"name": str(entry["name"]), "pattern": str(entry["pattern"])}
            for entry in _require(disc_raw, "redact_patterns", list, "discovery")
        ],
        bloat_min_rows=int(disc_raw.get("bloat_min_rows", 5000)),
        bloat_value_ratio=float(disc_raw.get("bloat_value_ratio", 0.10)),
    )

    thr_raw: Dict[str, Any] = _require(raw, "thresholds", dict, "config root")
    thresholds: Thresholds = Thresholds(
        overstock_months_supply=float(_require(thr_raw, "overstock_months_supply", float, "thresholds")),
        dead_stock_forecast_units=float(thr_raw.get("dead_stock_forecast_units", 0)),
        forecast_confidence_lookback_months=int(thr_raw.get("forecast_confidence_lookback_months", 3)),
        snapshot_stale_days=int(thr_raw.get("snapshot_stale_days", 7)),
        abc_basis=str(thr_raw.get("abc_basis", "revenue")),
        abc_cutoffs={str(k): float(v) for k, v in (thr_raw.get("abc_cutoffs") or {}).items()},
        worst_flag_hierarchy=[str(f) for f in _require(thr_raw, "worst_flag_hierarchy", list, "thresholds")],
    )

    canon_raw: Dict[str, Any] = _require(raw, "canonical", dict, "config root")
    canonical: CanonicalRules = CanonicalRules(
        sku=dict(_require(canon_raw, "sku", dict, "canonical")),
        csv=dict(_require(canon_raw, "csv", dict, "canonical")),
        locations={
            str(name): dict(attributes or {})
            for name, attributes in _require(canon_raw, "locations", dict, "canonical").items()
        },
        measures={
            str(name): dict(spec or {})
            for name, spec in _require(canon_raw, "measures", dict, "canonical").items()
        },
        channels={
            str(column): str(channel)
            for column, channel in _require(canon_raw, "channels", dict, "canonical").items()
        },
    )

    ext_raw: Dict[str, Any] = _require(raw, "extract", dict, "config root")
    extract: ExtractConfig = ExtractConfig(
        schema_file=str(_require(ext_raw, "schema_file", str, "extract")),
        drop_dir=str(_require(ext_raw, "drop_dir", str, "extract")),
        fetch_mode=str(ext_raw.get("fetch_mode", "local")),
        token_env_prefix=str(ext_raw.get("token_env_prefix", "SC_TOKEN_")),
        request_timeout_seconds=int(ext_raw.get("request_timeout_seconds", 120)),
        gates={str(k): float(v) for k, v in (ext_raw.get("gates") or {}).items()},
    )

    return Config(
        project=dict(raw.get("project") or {}),
        discovery=discovery,
        thresholds=thresholds,
        canonical=canonical,
        extract=extract,
        repo_root=resolved.parent.parent,
        config_path=resolved,
        _paths={str(k): str(v) for k, v in _require(raw, "paths", dict, "config root").items()},
    )
