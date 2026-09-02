"""Power Query extraction: DataMashup blob -> Section1.m -> individual queries.

Excel stores the entire Power Query stack as a base64 ``DataMashup`` element in
a customXml part. Inside that base64 is a *nested zip* whose ``Formulas/Section1.m``
holds every query as M source. Pulling it out turns an opaque binary blob into
version-controlled text, which is the only way a query change that breaks the
pipeline becomes visible in a diff.

Format of the decoded blob (little-endian):
    int32  version
    int32  package_parts_length
    bytes  package_parts        <- zip: Config/Package.xml, Formulas/Section1.m
    int32  permissions_length
    ...    (permissions, metadata, bindings — not needed here)

The header is trusted first, then a signature scan is used as a fallback,
because Excel has shipped more than one layout of the trailing sections.
"""

from __future__ import annotations

import base64
import io
import re
import struct
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

from .ooxml import local_name

SECTION_PART: str = "Formulas/Section1.m"
ZIP_LOCAL_HEADER: bytes = b"PK\x03\x04"
ZIP_EOCD: bytes = b"PK\x05\x06"


class MashupError(RuntimeError):
    """Raised when a DataMashup blob is present but cannot be decoded."""


@dataclass
class MQuery:
    """One ``shared`` declaration from Section1.m."""

    name: str
    source: str
    is_shared: bool = True
    metadata: str = ""
    sources: List[Dict[str, str]] = field(default_factory=list)

    @property
    def line_count(self) -> int:
        return self.source.count("\n") + 1

    def safe_filename(self) -> str:
        """Filesystem-safe name for writing to ``queries/*.m``."""
        cleaned: str = re.sub(r"[^A-Za-z0-9._-]+", "_", self.name).strip("_")
        return f"{cleaned or 'unnamed'}.m"


@dataclass
class MashupProbe:
    """Result of looking for Power Query inside one workbook."""

    found: bool = False
    part: Optional[str] = None
    section_text: str = ""
    queries: List[MQuery] = field(default_factory=list)
    package_parts: List[str] = field(default_factory=list)
    all_parts: List[str] = field(default_factory=list)
    error: Optional[str] = None

    @property
    def query_names(self) -> List[str]:
        return [q.name for q in self.queries]


# --------------------------------------------------------------------------
# Blob location and decoding
# --------------------------------------------------------------------------

def _decode_head(head: bytes) -> str:
    """Decode a part's leading bytes, honouring a UTF-16 BOM.

    Excel writes the DataMashup part as UTF-16LE with a BOM. An ASCII-only
    substring check therefore never matches it, and the search falls through to
    ``itemProps*.xml`` — which mentions DataMashup only in a schema reference.
    That combination reports "no Power Query" on a workbook that is full of it.
    """
    if head[:2] in (b"\xff\xfe", b"\xfe\xff"):
        encoding: str = "utf-16-le" if head[:2] == b"\xff\xfe" else "utf-16-be"
        return head[2:].decode(encoding, errors="replace")
    return head.decode("utf-8", errors="replace")


def find_mashup_parts(zf: zipfile.ZipFile) -> List[str]:
    """Every package part that really holds a DataMashup element.

    Excel has used ``customXml/item1.xml`` and ``xl/customXml/item1.xml``
    depending on version, so candidates are located by content, not by path.
    Two things this must get right:

    * ``itemProps*.xml`` is excluded. It carries a ``ds:schemaRef`` naming the
      DataMashup namespace without holding any mashup, and "customXml/item"
      is a substring of "customXml/itemProps".
    * The element name is matched after decoding, so a UTF-16 part is found.
    """
    candidates: List[str] = []
    for name in zf.namelist():
        lowered: str = name.lower()
        if not lowered.endswith(".xml") or "customxml/item" not in lowered:
            continue
        if "itemprops" in lowered or "/_rels/" in lowered:
            continue
        try:
            head: str = _decode_head(zf.read(name)[:8192])
        except (KeyError, zipfile.BadZipFile):
            continue
        if "<DataMashup" in head or ":DataMashup" in head:
            candidates.append(name)
    return candidates


def find_data_model_parts(zf: zipfile.ZipFile) -> List[str]:
    """Power Pivot / Excel Data Model parts.

    The data model is stored under the internal "Gemini" codename. A workbook
    carrying one has a second modelling layer beside the grid and the query
    stack, with its own DAX measures and relationships — worth knowing about,
    because it is another place a definition can live.
    """
    found: List[str] = []
    for name in zf.namelist():
        lowered: str = name.lower()
        if lowered.endswith(".xml") and "customxml/item" in lowered and "itemprops" not in lowered:
            try:
                head: str = _decode_head(zf.read(name)[:2048])
            except (KeyError, zipfile.BadZipFile):
                continue
            if "<Gemini" in head or ":Gemini" in head:
                found.append(name)
        elif "model/item" in lowered or lowered.endswith("xl/model/model.data"):
            found.append(name)
    return found


def _decode_base64_element(raw: bytes) -> bytes:
    """Pull the base64 payload out of a DataMashup XML part."""
    try:
        root: ET.Element = ET.fromstring(raw)
    except ET.ParseError as exc:
        raise MashupError(f"DataMashup part is not well-formed XML: {exc}") from exc
    if local_name(root.tag) != "DataMashup":
        # Some writers nest it one level down.
        for node in root.iter():
            if local_name(node.tag) == "DataMashup":
                root = node
                break
        else:
            raise MashupError("no DataMashup element found in part")
    payload: str = (root.text or "").strip()
    if not payload:
        raise MashupError("DataMashup element is empty")
    try:
        return base64.b64decode(payload, validate=False)
    except (ValueError, TypeError) as exc:
        raise MashupError(f"DataMashup base64 did not decode: {exc}") from exc


def _slice_inner_zip(blob: bytes) -> bytes:
    """Extract the nested zip package from a decoded DataMashup blob."""
    if len(blob) >= 8:
        try:
            _version, length = struct.unpack_from("<II", blob, 0)
            candidate: bytes = blob[8 : 8 + length]
            if candidate.startswith(ZIP_LOCAL_HEADER) and length > 0:
                return candidate
        except struct.error:
            pass

    # Fallback: locate the archive by signature and trim to the end-of-central-directory
    # record, since zipfile cannot tolerate arbitrary trailing bytes.
    start: int = blob.find(ZIP_LOCAL_HEADER)
    if start < 0:
        raise MashupError("decoded DataMashup contains no zip package")
    eocd: int = blob.rfind(ZIP_EOCD)
    if eocd < start:
        raise MashupError("decoded DataMashup zip has no end-of-central-directory record")
    comment_length: int = struct.unpack_from("<H", blob, eocd + 20)[0]
    return blob[start : eocd + 22 + comment_length]


def read_section_m(zf: zipfile.ZipFile, part: str) -> Tuple[str, List[str]]:
    """Return ``(Section1.m text, nested package part names)`` for one mashup part."""
    inner: bytes = _slice_inner_zip(_decode_base64_element(zf.read(part)))
    try:
        with zipfile.ZipFile(io.BytesIO(inner)) as package:
            names: List[str] = package.namelist()
            section: Optional[str] = next(
                (n for n in names if n.replace("\\", "/").endswith(SECTION_PART)), None
            )
            if section is None:
                raise MashupError(f"nested package has no {SECTION_PART}; parts={names}")
            return package.read(section).decode("utf-8-sig", errors="replace"), names
    except zipfile.BadZipFile as exc:
        raise MashupError(f"nested DataMashup package is not a readable zip: {exc}") from exc


# --------------------------------------------------------------------------
# M source parsing
# --------------------------------------------------------------------------

_DECL_PATTERN: re.Pattern[str] = re.compile(
    r"""^\s*
        (?P<meta>\[(?:[^\[\]]|\[[^\]]*\])*\]\s*)?   # optional [Description="..."] block
        (?P<shared>shared\s+)?
        (?:\#"(?P<quoted>(?:[^"]|"")*)"|(?P<plain>[A-Za-z_][A-Za-z0-9_.]*))
        \s*=\s*
    """,
    re.VERBOSE | re.DOTALL,
)


def split_declarations(section_text: str) -> List[str]:
    """Split a Section1.m document into top-level declarations.

    Splits on semicolons at bracket depth zero, while ignoring semicolons inside
    string literals and comments. A naive ``split(';')`` mangles any query
    containing a delimiter or a URL, which is most of them.
    """
    chunks: List[str] = []
    buffer: List[str] = []
    depth: int = 0
    index: int = 0
    length: int = len(section_text)
    in_string: bool = False
    in_line_comment: bool = False
    in_block_comment: bool = False

    while index < length:
        char: str = section_text[index]
        nxt: str = section_text[index + 1] if index + 1 < length else ""

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
                if nxt == '"':          # "" is an escaped quote inside M strings
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
        if char in "([{":
            depth += 1
        elif char in ")]}":
            depth = max(depth - 1, 0)
        elif char == ";" and depth == 0:
            chunks.append("".join(buffer))
            buffer = []
            index += 1
            continue

        buffer.append(char)
        index += 1

    tail: str = "".join(buffer).strip()
    if tail:
        chunks.append(tail)
    return [c for c in (chunk.strip() for chunk in chunks) if c]


def parse_queries(section_text: str) -> List[MQuery]:
    """Parse Section1.m into named queries, skipping the ``section`` header."""
    queries: List[MQuery] = []
    for chunk in split_declarations(section_text):
        if re.match(r"^\s*section\b", chunk):
            continue
        match: Optional[re.Match[str]] = _DECL_PATTERN.match(chunk)
        if match is None:
            continue
        quoted: Optional[str] = match.group("quoted")
        name: str = (quoted.replace('""', '"') if quoted is not None else match.group("plain") or "").strip()
        if not name:
            continue
        body: str = chunk[match.end():].strip()
        query: MQuery = MQuery(
            name=name,
            source=f"{chunk};",
            is_shared=match.group("shared") is not None,
            metadata=(match.group("meta") or "").strip(),
        )
        query.sources = extract_sources(body)
        queries.append(query)
    return queries


# --------------------------------------------------------------------------
# External source extraction — "what breaks if a path moves"
# --------------------------------------------------------------------------

_SOURCE_FUNCTIONS: Dict[str, str] = {
    "File.Contents": "file",
    "Folder.Files": "folder",
    "Folder.Contents": "folder",
    "Excel.Workbook": "workbook",
    "Csv.Document": "csv",
    "Json.Document": "json",
    "Web.Contents": "web",
    "SharePoint.Files": "sharepoint",
    "SharePoint.Contents": "sharepoint",
    "SharePoint.Tables": "sharepoint",
    "Sql.Database": "sql",
    "Sql.Databases": "sql",
    "Odbc.DataSource": "odbc",
    "Odbc.Query": "odbc",
    "OData.Feed": "odata",
    "AzureStorage.Blobs": "azure_blob",
    "AzureStorage.DataLake": "azure_datalake",
    "Xml.Tables": "xml",
    "Xml.Document": "xml",
}

_FUNCTION_CALL: re.Pattern[str] = re.compile(
    r"\b(" + "|".join(re.escape(fn) for fn in _SOURCE_FUNCTIONS) + r")\s*\(\s*\"((?:[^\"]|\"\")*)\""
)

# Bare literals that are unmistakably a location, caught even when wrapped in a
# helper function or built by concatenation.
_PATH_LITERAL: re.Pattern[str] = re.compile(
    r"\"((?:[A-Za-z]:\\|\\\\)[^\"]{3,}|https?://[^\"]{5,})\""
)


def extract_sources(m_body: str) -> List[Dict[str, str]]:
    """Every external location a query reads from, deduplicated and typed."""
    found: Dict[Tuple[str, str], Dict[str, str]] = {}

    for match in _FUNCTION_CALL.finditer(m_body):
        function: str = match.group(1)
        value: str = match.group(2).replace('""', '"')
        found.setdefault(
            (_SOURCE_FUNCTIONS[function], value),
            {"kind": _SOURCE_FUNCTIONS[function], "via": function, "location": value},
        )

    already: set[str] = {location for _kind, location in found}
    for match in _PATH_LITERAL.finditer(m_body):
        value = match.group(1).replace('""', '"')
        if value in already:
            continue                # already typed by its source function above
        kind: str = "sharepoint" if "sharepoint.com" in value.lower() else (
            "web" if value.lower().startswith("http") else
            "unc" if value.startswith("\\\\") else "file"
        )
        found.setdefault((kind, value), {"kind": kind, "via": "literal", "location": value})

    # A query referencing another query is a dependency too, and the usual cause
    # of a refresh-order bug. Resolved against the section's real query names in
    # resolve_query_refs() — every `let` step is also written as #"Name", so
    # matching the syntax alone reports "Promoted Headers" as a dependency.
    return list(found.values())


def _identifier_references(m_body: str) -> List[str]:
    """Every ``#"..."`` identifier in a query body, in order of appearance."""
    return [
        match.group(1).replace('""', '"')
        for match in re.finditer(r'#"((?:[^"]|"")*)"', m_body)
    ]


def resolve_query_refs(queries: Sequence["MQuery"]) -> None:
    """Attach cross-query dependencies, matched exactly against known query names.

    Mutates each query's ``sources`` in place. A reference is a dependency only
    when it names another query in the same section; anything else is one of the
    query's own ``let`` steps.
    """
    known: Dict[str, str] = {q.name.strip().lower(): q.name for q in queries}
    for query in queries:
        own: str = query.name.strip().lower()
        for referenced in _identifier_references(query.source):
            canonical: Optional[str] = known.get(referenced.strip().lower())
            if canonical is None or referenced.strip().lower() == own:
                continue
            if any(s["kind"] == "query_ref" and s["location"] == canonical for s in query.sources):
                continue
            query.sources.append(
                {"kind": "query_ref", "via": "reference", "location": canonical}
            )


def probe_mashup(zf: zipfile.ZipFile) -> MashupProbe:
    """Locate and decode the Power Query stack in an open workbook package."""
    parts: List[str] = find_mashup_parts(zf)
    if not parts:
        return MashupProbe(found=False)

    probe: MashupProbe = MashupProbe(found=True, part=parts[0])
    probe.all_parts = parts
    errors: List[str] = []
    for part in parts:
        try:
            section_text, package_parts = read_section_m(zf, part)
        except MashupError as exc:
            errors.append(f"{part}: {exc}")
            continue
        probe.part = part
        probe.section_text, probe.package_parts = section_text, package_parts
        probe.queries = parse_queries(section_text)
        resolve_query_refs(probe.queries)
        break
    if not probe.queries and errors:
        probe.error = "; ".join(errors)
    return probe
