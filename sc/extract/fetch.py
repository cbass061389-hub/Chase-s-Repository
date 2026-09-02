"""Get the bytes of an export, without mixing that up with parsing.

Fetching and parsing are separate on purpose: the readers are then testable
against a local file with no credentials, and the credential handling lives in
exactly one place.

Two modes:

``local``
    Read ``data/drop/<slug>.csv``. This is the default and the only mode used in
    tests. It is also what a scheduled NetSuite saved-search delivery writes to.

``url``
    Fetch the File Cabinet URL. The access token is **not** stored in config or
    in git — it is read from the environment, per source, as
    ``SC_TOKEN_<SLUG>``. A missing token is an error, never a silent skip.
"""

from __future__ import annotations

import os
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from ..configuration import ExtractConfig
from .schema_store import SourceSchema

MODE_LOCAL: str = "local"
MODE_URL: str = "url"


class FetchError(RuntimeError):
    """Raised when an export cannot be retrieved. Never swallowed."""


@dataclass
class Fetched:
    slug: str
    payload: bytes
    origin: str
    mode: str

    @property
    def size_bytes(self) -> int:
        return len(self.payload)


def token_env_var(prefix: str, slug: str) -> str:
    """Environment variable holding this source's access token."""
    return f"{prefix}{slug.upper()}"


def fetch(
    schema: SourceSchema,
    config: ExtractConfig,
    repo_root: Path,
    mode: Optional[str] = None,
) -> Fetched:
    """Retrieve one export's bytes."""
    resolved_mode: str = (mode or config.fetch_mode or MODE_LOCAL).lower()

    if resolved_mode == MODE_LOCAL:
        drop: Path = repo_root / config.drop_dir
        candidates = [drop / f"{schema.slug}.csv", drop / f"{schema.slug}.CSV"]
        for candidate in candidates:
            if candidate.is_file():
                return Fetched(schema.slug, candidate.read_bytes(), str(candidate), MODE_LOCAL)
        raise FetchError(
            f"{schema.slug}: no local export at {candidates[0]}. Either drop the CSV there, "
            f"or run with --fetch-mode url and set {token_env_var(config.token_env_prefix, schema.slug)}."
        )

    if resolved_mode == MODE_URL:
        if not schema.label.startswith("http") and "netsuite" not in schema.key:
            raise FetchError(f"{schema.slug}: not a fetchable URL source ({schema.kind})")
        variable: str = token_env_var(config.token_env_prefix, schema.slug)
        token: str = os.environ.get(variable, "")
        if not token:
            raise FetchError(
                f"{schema.slug}: {variable} is not set. The File Cabinet access token is "
                "deliberately not stored in config or git; export it in the environment."
            )
        url: str = _rebuild_url(schema, token)
        try:
            with urllib.request.urlopen(url, timeout=config.request_timeout_seconds) as response:
                payload: bytes = response.read()
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as exc:
            raise FetchError(f"{schema.slug}: fetch failed ({exc})") from exc
        if not payload:
            raise FetchError(f"{schema.slug}: fetch returned no content")
        return Fetched(schema.slug, payload, _redacted(url), MODE_URL)

    raise FetchError(f"unknown fetch mode '{resolved_mode}'; expected '{MODE_LOCAL}' or '{MODE_URL}'")


def _rebuild_url(schema: SourceSchema, token: str) -> str:
    """Reassemble the export URL with the token supplied from the environment.

    The committed schema holds the endpoint with the credential stripped, so the
    token has to be reattached at fetch time.
    """
    base: str = schema.label
    if "media.nl" in base and "&h=" not in base:
        return f"{base}&h={token}"
    if "<REDACTED>" in base:
        return base.replace("<REDACTED>", token)
    return base


def _redacted(url: str) -> str:
    import re

    return re.sub(r"([?&]h=)[^&]+", r"\1<REDACTED>", url)
