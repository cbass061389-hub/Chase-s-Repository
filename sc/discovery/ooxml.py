"""Shared OOXML/zip plumbing for the discovery probes.

Everything here is stdlib-only and read-only. Workbooks are opened as zip
archives and their XML parts streamed, never loaded through Excel or openpyxl:
the estate contains multi-hundred-megabyte sheets that defeat DOM parsers, and
the tool has to run on a locked-down laptop with no installs.
"""

from __future__ import annotations

import posixpath
import zipfile
from typing import Dict, Iterator, List, Optional, Tuple
from xml.etree import ElementTree as ET

NS_MAIN: str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL: str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL: str = "http://schemas.openxmlformats.org/package/2006/relationships"

# Relationship type suffixes we care about, matched on the tail of the Type URI.
REL_WORKSHEET: str = "/worksheet"
REL_TABLE: str = "/table"
REL_EXTERNAL_LINK: str = "/externalLink"
REL_EXTERNAL_LINK_PATH: str = "/externalLinkPath"
REL_CONNECTIONS: str = "/connections"
REL_VBA: str = "/vbaProject"
REL_CUSTOM_XML: str = "/customXml"


def qn(namespace: str, tag: str) -> str:
    """Qualified XML name."""
    return f"{{{namespace}}}{tag}"


def local_name(tag: str) -> str:
    """Strip the namespace from an element tag."""
    return tag.rsplit("}", 1)[-1]


def resolve_part(base_part: str, target: str) -> str:
    """Resolve a relationship Target against the part that declared it.

    Handles absolute targets ("/xl/worksheets/sheet1.xml"), relative targets
    ("worksheets/sheet1.xml") and parent traversal ("../sharedStrings.xml").
    Returns a zip-entry path with no leading slash.
    """
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir: str = posixpath.dirname(base_part)
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def rels_part_for(part: str) -> str:
    """Path of the ``_rels`` part describing *part*."""
    base_dir: str = posixpath.dirname(part)
    name: str = posixpath.basename(part)
    return posixpath.join(base_dir, "_rels", f"{name}.rels").lstrip("/")


class Relationship:
    """One entry from a ``_rels`` part."""

    __slots__ = ("rel_id", "type_uri", "target", "target_mode", "resolved")

    def __init__(self, rel_id: str, type_uri: str, target: str, target_mode: str, resolved: str) -> None:
        self.rel_id: str = rel_id
        self.type_uri: str = type_uri
        self.target: str = target
        self.target_mode: str = target_mode
        self.resolved: str = resolved

    @property
    def is_external(self) -> bool:
        return self.target_mode.lower() == "external"

    def type_is(self, suffix: str) -> bool:
        return self.type_uri.endswith(suffix)


def read_rels(zf: zipfile.ZipFile, part: str) -> Dict[str, Relationship]:
    """Parse the relationships declared for *part*. Missing rels yields ``{}``."""
    rels_path: str = rels_part_for(part)
    if rels_path not in zf.namelist():
        return {}
    out: Dict[str, Relationship] = {}
    with zf.open(rels_path) as handle:
        tree: ET.Element = ET.parse(handle).getroot()
    for node in tree.findall(qn(NS_PKG_REL, "Relationship")):
        rel_id: str = node.get("Id", "")
        target: str = node.get("Target", "")
        mode: str = node.get("TargetMode", "Internal")
        resolved: str = "" if mode.lower() == "external" else resolve_part(part, target)
        out[rel_id] = Relationship(rel_id, node.get("Type", ""), target, mode, resolved)
    return out


def iter_shared_strings(zf: zipfile.ZipFile, cap: int) -> Tuple[List[str], bool]:
    """Stream ``xl/sharedStrings.xml`` into a list, capped at *cap* entries.

    Returns ``(strings, truncated)``. Truncation is reported rather than hidden,
    because a header resolved from a truncated table would be silently wrong.
    """
    part: str = "xl/sharedStrings.xml"
    if part not in zf.namelist():
        return [], False
    strings: List[str] = []
    truncated: bool = False
    si_tag: str = qn(NS_MAIN, "si")
    t_tag: str = qn(NS_MAIN, "t")
    with zf.open(part) as handle:
        for event, element in ET.iterparse(handle, events=("end",)):
            if element.tag != si_tag:
                continue
            # An <si> is either a single <t> or a run of <r><t> fragments.
            strings.append("".join(node.text or "" for node in element.iter(t_tag)))
            element.clear()
            if len(strings) >= cap:
                truncated = True
                break
    return strings, truncated


def cell_column_letters(reference: Optional[str]) -> str:
    """Column portion of an A1 reference; ``''`` when absent."""
    if not reference:
        return ""
    return "".join(character for character in reference if character.isalpha())


def column_index(letters: str) -> int:
    """Zero-based column index for column letters. ``-1`` when unparseable."""
    if not letters:
        return -1
    index: int = 0
    for character in letters.upper():
        if not ("A" <= character <= "Z"):
            return -1
        index = index * 26 + (ord(character) - 64)
    return index - 1


def open_workbook(path: str) -> zipfile.ZipFile:
    """Open an OOXML package read-only, raising a clear error for legacy/corrupt files."""
    try:
        return zipfile.ZipFile(path, "r")
    except zipfile.BadZipFile as exc:  # .xls, encrypted, or truncated download
        raise NotAnOoxmlPackage(f"{path}: not a readable OOXML package ({exc})") from exc


class NotAnOoxmlPackage(RuntimeError):
    """Raised for files that are not zip-based Office packages (legacy .xls, encrypted, corrupt)."""


def iter_parts(zf: zipfile.ZipFile) -> Iterator[Tuple[str, int]]:
    """Yield ``(part_name, uncompressed_size)`` for every entry in the package."""
    for info in zf.infolist():
        yield info.filename, info.file_size
