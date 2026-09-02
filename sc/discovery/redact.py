"""Strip secrets before anything reaches disk.

Power Query text routinely embeds credentials. A NetSuite File Cabinet media
link carries its access token in the `h=` parameter, and that token is enough to
pull the export without logging in. Since `queries/` is committed so query
changes show up in a diff, redaction is not optional — committing the extracted
M source verbatim would publish live access tokens into git history.

Patterns come from `config.yaml`. Every redaction is counted and reported, so
the report can say what was stripped rather than silently altering the record.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Sequence, Tuple

REDACTION_MARKER: str = "<REDACTED>"


@dataclass
class Redactor:
    """Compiled redaction rules plus a running tally of what they matched."""

    rules: List[Tuple[str, re.Pattern[str]]] = field(default_factory=list)
    counts: Dict[str, int] = field(default_factory=dict)

    @classmethod
    def from_config(cls, patterns: Sequence[Dict[str, str]]) -> "Redactor":
        compiled: List[Tuple[str, re.Pattern[str]]] = []
        for entry in patterns:
            name: str = entry["name"]
            try:
                compiled.append((name, re.compile(entry["pattern"])))
            except re.error as exc:
                raise ValueError(f"redact pattern '{name}' is not a valid regex: {exc}") from exc
        return cls(rules=compiled)

    def scrub(self, text: str) -> str:
        """Return *text* with every configured secret replaced.

        Each pattern keeps its leading capture group (the parameter name or URL
        prefix) so the reader can still see *what kind* of credential was there
        and which endpoint it belonged to.
        """
        if not text:
            return text
        scrubbed: str = text
        for name, pattern in self.rules:
            scrubbed, hits = pattern.subn(
                lambda match: (match.group(1) if match.lastindex else "") + REDACTION_MARKER,
                scrubbed,
            )
            if hits:
                self.counts[name] = self.counts.get(name, 0) + hits
        return scrubbed

    @property
    def total(self) -> int:
        return sum(self.counts.values())

    def summary(self) -> List[Dict[str, str]]:
        return [{"pattern": name, "occurrences": str(count)} for name, count in sorted(self.counts.items())]
