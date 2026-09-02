"""VBA module inventory from ``xl/vbaProject.bin`` — names only, never executed.

The project is an OLE compound file whose module directory is RLE-compressed,
but the ``PROJECT`` stream is stored as plain text and lists every component:

    Module=mod_Refresh
    Class=cls_Allocation
    Document=Sheet1/&H00000000
    BaseClass=frmDSR

Scanning for those lines gives the module inventory without an OLE parser and
without any dependency, which matters because this has to run with nothing but
the standard library installed. It is explicitly a heuristic: the result is
reported as "declared components", and a workbook whose macros matter gets read
properly in a later phase.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List

VBA_PART: str = "xl/vbaProject.bin"

_COMPONENT: re.Pattern[bytes] = re.compile(
    rb"(?m)^(Module|Class|BaseClass|Document|Package)=([^\r\n/&]{1,64})"
)
_PROJECT_NAME: re.Pattern[bytes] = re.compile(rb"(?m)^Name=\"?([A-Za-z_][A-Za-z0-9_]{0,63})")


@dataclass
class VbaProbe:
    """What the VBA project declares about itself."""

    present: bool = False
    project_name: str = ""
    components: Dict[str, List[str]] = field(default_factory=dict)
    protected: bool = False

    @property
    def module_names(self) -> List[str]:
        """Every declared component name, code modules first."""
        ordered: List[str] = []
        for kind in ("Module", "Class", "BaseClass", "Document", "Package"):
            ordered.extend(self.components.get(kind, []))
        return ordered

    @property
    def count(self) -> int:
        return len(self.module_names)

    @property
    def code_modules(self) -> List[str]:
        """Standard and class modules — the components that hold real logic.

        Excel declares one ``Document=`` component per worksheet, so a 76-tab
        workbook lists 70-odd "Sheet12"-style entries that say nothing. Those are
        separated out so the modules worth reading stay visible.
        """
        named: List[str] = []
        for kind in ("Module", "Class", "BaseClass"):
            named.extend(self.components.get(kind, []))
        return named

    @property
    def document_modules(self) -> List[str]:
        return list(self.components.get("Document", []))


def probe_vba(zf: zipfile.ZipFile) -> VbaProbe:
    """Inventory the VBA project in an open workbook package. Reads only."""
    if VBA_PART not in zf.namelist():
        return VbaProbe(present=False)

    raw: bytes = zf.read(VBA_PART)
    probe: VbaProbe = VbaProbe(present=True)

    # A locked project stores CMG/DPB/GC protection keys in the PROJECT stream.
    probe.protected = b"DPB=" in raw and b"CMG=" in raw

    name_match = _PROJECT_NAME.search(raw)
    if name_match is not None:
        probe.project_name = name_match.group(1).decode("latin-1", errors="replace")

    seen: Dict[str, List[str]] = {}
    for match in _COMPONENT.finditer(raw):
        kind: str = match.group(1).decode("ascii")
        component: str = match.group(2).decode("latin-1", errors="replace").strip().strip('"')
        if not component:
            continue
        bucket: List[str] = seen.setdefault(kind, [])
        if component not in bucket:
            bucket.append(component)
    probe.components = seen
    return probe
