"""A shallow structural reader for Power Query M.

Not a full M parser, and it does not try to be. It answers the specific
questions that decide whether two queries reading the same export agree:

* what parse options the source was read with,
* what the query grouped by, and **which column each output measure actually
  aggregates**,
* what it filtered,
* how it derived its key columns,
* what it renamed, removed and finally emitted.

The measure lineage is the part that matters. Two queries can both emit a column
called "Available" while one sums ``[On Hand]`` and the other sums ``[Available]``
— the names agree and the numbers cannot.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

OPEN_BRACKETS: str = "([{"
CLOSE_BRACKETS: str = ")]}"


def split_top_level(text: str, separator: str = ",") -> List[str]:
    """Split *text* on *separator* at bracket depth zero, ignoring strings/comments."""
    parts: List[str] = []
    buffer: List[str] = []
    depth: int = 0
    index: int = 0
    length: int = len(text)
    in_string: bool = False
    in_line_comment: bool = False
    in_block_comment: bool = False

    while index < length:
        char: str = text[index]
        nxt: str = text[index + 1] if index + 1 < length else ""

        if in_line_comment:
            buffer.append(char)
            if char == "\n":
                in_line_comment = False
            index += 1
            continue
        if in_block_comment:
            buffer.append(char)
            if char == "*" and nxt == "/":
                buffer.append(nxt)
                index += 2
                in_block_comment = False
                continue
            index += 1
            continue
        if in_string:
            buffer.append(char)
            if char == '"':
                if nxt == '"':
                    buffer.append(nxt)
                    index += 2
                    continue
                in_string = False
            index += 1
            continue

        if char == "/" and nxt == "/":
            in_line_comment = True
            buffer.append(char)
            index += 1
            continue
        if char == "/" and nxt == "*":
            in_block_comment = True
            buffer.append(char)
            index += 1
            continue
        if char == '"':
            in_string = True
            buffer.append(char)
            index += 1
            continue
        if char in OPEN_BRACKETS:
            depth += 1
        elif char in CLOSE_BRACKETS:
            depth = max(depth - 1, 0)
        elif char == separator and depth == 0:
            parts.append("".join(buffer))
            buffer = []
            index += 1
            continue

        buffer.append(char)
        index += 1

    tail: str = "".join(buffer)
    if tail.strip():
        parts.append(tail)
    return [part.strip() for part in parts if part.strip()]


def strip_one_brace_layer(text: str) -> str:
    """Remove exactly one enclosing ``{...}``.

    ``str.lstrip("{")`` strips *every* leading brace, which flattens a nested M
    list literal one level too far and splits an aggregation record into its
    own fields.
    """
    trimmed: str = text.strip()
    if trimmed.startswith("{") and trimmed.endswith("}"):
        return trimmed[1:-1].strip()
    return trimmed


def unquote(text: str) -> str:
    """Strip surrounding double quotes from an M literal, if present."""
    trimmed: str = text.strip()
    if len(trimmed) >= 2 and trimmed[0] == '"' and trimmed[-1] == '"':
        return trimmed[1:-1].replace('""', '"')
    return trimmed


def quoted_strings(text: str) -> List[str]:
    """Every double-quoted literal in *text*, unescaped."""
    return [m.group(1).replace('""', '"') for m in re.finditer(r'"((?:[^"]|"")*)"', text)]


def call_arguments(expression: str, function: str) -> Optional[List[str]]:
    """Top-level arguments of the first ``function(...)`` call in *expression*."""
    match: Optional[re.Match[str]] = re.search(rf"\b{re.escape(function)}\s*\(", expression)
    if match is None:
        return None
    start: int = match.end()
    depth: int = 1
    index: int = start
    in_string: bool = False
    while index < len(expression) and depth > 0:
        char: str = expression[index]
        if in_string:
            if char == '"':
                if index + 1 < len(expression) and expression[index + 1] == '"':
                    index += 2
                    continue
                in_string = False
        elif char == '"':
            in_string = True
        elif char in OPEN_BRACKETS:
            depth += 1
        elif char in CLOSE_BRACKETS:
            depth -= 1
        index += 1
    return split_top_level(expression[start : index - 1])


@dataclass
class MStep:
    """One ``let`` step: a name and its expression."""

    name: str
    expression: str

    @property
    def leading_function(self) -> str:
        match = re.match(r"\s*([A-Za-z][A-Za-z0-9_]*(?:\.[A-Za-z][A-Za-z0-9_]*)*)\s*\(", self.expression)
        return match.group(1) if match else ""


@dataclass
class SourceRead:
    """How the upstream bytes were parsed."""

    reader: str = ""
    fetcher: str = ""
    location: str = ""
    options: Dict[str, str] = field(default_factory=dict)

    def option(self, name: str) -> str:
        return self.options.get(name, "")


@dataclass
class Aggregation:
    output: str
    function: str
    column: str

    @property
    def signature(self) -> str:
        return f"{self.function}([{self.column}])"


@dataclass
class GroupBy:
    keys: List[str] = field(default_factory=list)
    aggregations: List[Aggregation] = field(default_factory=list)


@dataclass
class QueryProfile:
    """The semantic fingerprint of one query."""

    workbook: str
    query: str
    upstream_key: str = ""
    upstream_label: str = ""
    source: SourceRead = field(default_factory=SourceRead)
    groups: List[GroupBy] = field(default_factory=list)
    filters: List[str] = field(default_factory=list)
    derived: List[Tuple[str, str]] = field(default_factory=list)
    renames: List[Tuple[str, str]] = field(default_factory=list)
    removed: List[str] = field(default_factory=list)
    output_columns: List[str] = field(default_factory=list)
    column_types: List[Tuple[str, str]] = field(default_factory=list)
    step_count: int = 0

    @property
    def declared_schema(self) -> Dict[str, str]:
        """Column -> M type, as the query declares the upstream.

        Table.TransformColumnTypes enumerates every column of the export, so the
        schema of a NetSuite export can be recovered from any query that reads
        it — no credentials and no fetch required.
        """
        return dict(self.column_types)

    @property
    def measures(self) -> Dict[str, str]:
        """Final column name -> the aggregation that actually produced it.

        Renames are replayed in order, so a measure emitted as "Available" is
        traced back to the column it summed.
        """
        current: Dict[str, str] = {}
        for group in self.groups:
            for aggregation in group.aggregations:
                current[aggregation.output] = aggregation.signature
        for old, new in self.renames:
            if old in current:
                current[new] = current.pop(old)
        return current

    @property
    def group_keys(self) -> List[str]:
        """Grain, expressed as the last grouping applied."""
        return list(self.groups[-1].keys) if self.groups else []

    @property
    def resolved_derived(self) -> Dict[str, str]:
        """Derived columns under their FINAL names, renames replayed.

        Power Query names an added column "Custom.1" and renames it later; the
        derivation is only findable under the name it ends up with.
        """
        current: Dict[str, str] = dict(self.derived)
        for old, new in self.renames:
            if old in current:
                current[new] = current.pop(old)
        return current

    @property
    def key_derivations(self) -> Dict[str, str]:
        """Derived columns that look like a join key, with their expressions."""
        wanted: Tuple[str, ...] = ("sku", "item", "name", "id", "key", "canon")
        return {
            name: expression
            for name, expression in self.resolved_derived.items()
            if any(token in name.lower() for token in wanted)
        }


def parse_let_steps(query_source: str) -> List[MStep]:
    """Split a query's ``let ... in ...`` body into steps."""
    body: str = query_source
    match = re.search(r"\blet\b", body)
    if match is None:
        return []
    body = body[match.end():]
    # Trim the trailing "in <expr>;" at depth zero.
        # Handled by splitting first, then discarding the final `in` clause.
    steps: List[MStep] = []
    for chunk in split_top_level(body):
        assignment = re.match(r'\s*(?:#"((?:[^"]|"")*)"|([A-Za-z_][A-Za-z0-9_.]*))\s*=\s*(.*)$',
                              chunk, re.DOTALL)
        if assignment is None:
            continue
        name: str = (assignment.group(1) or assignment.group(2) or "").replace('""', '"')
        expression: str = assignment.group(3).strip()
        expression = re.sub(r"\bin\b\s+[^\n]*;?\s*$", "", expression, flags=re.DOTALL).strip()
        steps.append(MStep(name=name, expression=expression))
    return steps


def _parse_source(step: MStep, profile: QueryProfile) -> None:
    """Extract reader, fetcher, location and parse options from a source step."""
    for reader in ("Csv.Document", "Excel.Workbook", "Json.Document", "Xml.Tables"):
        if reader in step.expression:
            profile.source.reader = reader
            break
    for fetcher in ("Web.Contents", "File.Contents", "SharePoint.Files", "Folder.Files"):
        if fetcher in step.expression:
            profile.source.fetcher = fetcher
            break
    literals: List[str] = quoted_strings(step.expression)
    if literals:
        profile.source.location = literals[0]

    # Options arrive as an M record: [Delimiter=",", Columns=24, QuoteStyle=QuoteStyle.None]
    record = re.search(r"\[([^\[\]]*(?:Delimiter|QuoteStyle|Columns|Encoding)[^\[\]]*)\]",
                       step.expression)
    if record is not None:
        for entry in split_top_level(record.group(1)):
            if "=" not in entry:
                continue
            key, _sep, value = entry.partition("=")
            profile.source.options[key.strip()] = unquote(value)


def _parse_group(step: MStep, profile: QueryProfile) -> None:
    arguments: Optional[List[str]] = call_arguments(step.expression, "Table.Group")
    if arguments is None or len(arguments) < 2:
        return
    group: GroupBy = GroupBy(keys=quoted_strings(arguments[1]))
    if len(arguments) >= 3:
        for entry in split_top_level(strip_one_brace_layer(arguments[2])):
            names: List[str] = quoted_strings(entry)
            if not names:
                continue
            aggregation = re.search(r"each\s+(List\.[A-Za-z]+|Table\.[A-Za-z]+)\s*\(\s*\[([^\]]+)\]",
                                    entry)
            group.aggregations.append(
                Aggregation(
                    output=names[0],
                    function=aggregation.group(1) if aggregation else "unparsed",
                    column=aggregation.group(2).strip() if aggregation else "?",
                )
            )
    profile.groups.append(group)


def profile_query(workbook: str, query: str, source_text: str) -> QueryProfile:
    """Build the semantic fingerprint of one query from its M source."""
    profile: QueryProfile = QueryProfile(workbook=workbook, query=query)
    steps: List[MStep] = parse_let_steps(source_text)
    profile.step_count = len(steps)

    for step in steps:
        function: str = step.leading_function
        expression: str = step.expression

        if function in ("Csv.Document", "Excel.Workbook", "Json.Document", "Xml.Tables") or (
            not profile.source.reader and ("Web.Contents" in expression or "File.Contents" in expression)
        ):
            _parse_source(step, profile)

        if function == "Table.Group":
            _parse_group(step, profile)

        elif function == "Table.SelectRows":
            arguments = call_arguments(expression, "Table.SelectRows")
            if arguments and len(arguments) >= 2:
                predicate: str = re.sub(r"\s+", " ", arguments[1]).strip()
                predicate = re.sub(r"^each\s+", "", predicate)
                if predicate and predicate != "true":
                    profile.filters.append(predicate)

        elif function == "Table.AddColumn":
            arguments = call_arguments(expression, "Table.AddColumn")
            if arguments and len(arguments) >= 3:
                names = quoted_strings(arguments[1])
                if names:
                    body: str = re.sub(r"\s+", " ", arguments[2]).strip()
                    profile.derived.append((names[0], re.sub(r"^each\s+", "", body)))

        elif function == "Table.RenameColumns":
            arguments = call_arguments(expression, "Table.RenameColumns")
            if arguments and len(arguments) >= 2:
                for pair in split_top_level(strip_one_brace_layer(arguments[1])):
                    names = quoted_strings(pair)
                    if len(names) >= 2:
                        profile.renames.append((names[0], names[1]))

        elif function == "Table.RemoveColumns":
            arguments = call_arguments(expression, "Table.RemoveColumns")
            if arguments and len(arguments) >= 2:
                profile.removed.extend(quoted_strings(arguments[1]))

        elif function == "Table.TransformColumnTypes":
            arguments = call_arguments(expression, "Table.TransformColumnTypes")
            if arguments and len(arguments) >= 2:
                for pair in split_top_level(strip_one_brace_layer(arguments[1])):
                    fields: List[str] = split_top_level(strip_one_brace_layer(pair))
                    if len(fields) < 2:
                        continue
                    names = quoted_strings(fields[0])
                    if names:
                        profile.column_types.append((names[0], fields[1].strip()))

        elif function in ("Table.ReorderColumns", "Table.SelectColumns"):
            arguments = call_arguments(expression, function)
            if arguments and len(arguments) >= 2:
                profile.output_columns = quoted_strings(arguments[1])

    return profile
