"""Preserve worksheet DrawingML that openpyxl cannot round-trip.

Excel frequently uses ordinary DrawingML shapes as macro buttons.  openpyxl
supports charts and images in drawings, but does not load arbitrary shapes;
an unrelated workbook save can therefore remove the shape, its worksheet
relationship, and its content-type declaration.  This module snapshots and
restores that OOXML state without interpreting or rewriting the drawing.
"""

from __future__ import annotations

import posixpath
from pathlib import PurePosixPath
from typing import Any
from zipfile import ZipFile

from openpyxl.xml.functions import SubElement, fromstring, tostring

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
DRAWING_REL = f"{DOCUMENT_REL_NS}/drawing"


def relationship_part_name(part: str) -> str:
    path = PurePosixPath(part)
    return str(path.parent / "_rels" / f"{path.name}.rels")


def _resolve_package_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(
        str(PurePosixPath(source_part).parent / target)
    )


def _collect_related_package_parts(
    archive: ZipFile,
    roots: set[str],
) -> set[str]:
    """Return internal OOXML parts reachable from the supplied roots."""

    names = set(archive.namelist())
    found: set[str] = set()
    pending = list(roots)
    while pending:
        part = pending.pop()
        if part in found or part not in names:
            continue
        found.add(part)
        relationships_part = relationship_part_name(part)
        if relationships_part not in names:
            continue
        found.add(relationships_part)
        relationships = fromstring(archive.read(relationships_part))
        for relationship in relationships:
            if relationship.get("TargetMode") == "External":
                continue
            target = relationship.get("Target")
            if not target:
                continue
            resolved = _resolve_package_target(part, target)
            if resolved not in found:
                pending.append(resolved)
    return found


def _discover_content_types(
    archive: ZipFile,
    preserved_parts: set[str],
) -> dict[str, dict[str, str]]:
    root = fromstring(archive.read("[Content_Types].xml"))
    extensions = {
        PurePosixPath(part).suffix.lstrip(".")
        for part in preserved_parts
        if PurePosixPath(part).suffix
    }
    defaults = {
        item.get("Extension"): item.get("ContentType")
        for item in root.findall(f"{{{CONTENT_TYPES_NS}}}Default")
        if item.get("Extension") in extensions and item.get("ContentType")
    }
    overrides = {
        item.get("PartName"): item.get("ContentType")
        for item in root.findall(f"{{{CONTENT_TYPES_NS}}}Override")
        if (item.get("PartName") or "").lstrip("/") in preserved_parts
        and item.get("ContentType")
    }
    return {"defaults": defaults, "overrides": overrides}


def discover_drawing_state(
    archive: ZipFile,
    sheet_parts: dict[str, str],
) -> dict[str, Any]:
    """Capture drawings, dependencies, and per-sheet relationships."""

    names = set(archive.namelist())
    relationships_by_sheet: dict[str, dict[str, str]] = {}
    drawing_roots: set[str] = set()
    rel_id_key = f"{{{DOCUMENT_REL_NS}}}id"

    for sheet_name, sheet_part in sheet_parts.items():
        worksheet = fromstring(archive.read(sheet_part))
        drawing = worksheet.find(f"{{{MAIN_NS}}}drawing")
        if drawing is None:
            continue
        rel_id = drawing.get(rel_id_key)
        relationships_part = relationship_part_name(sheet_part)
        if not rel_id or relationships_part not in names:
            continue

        relationships = fromstring(archive.read(relationships_part))
        relationship = next(
            (item for item in relationships if item.get("Id") == rel_id),
            None,
        )
        if relationship is None or relationship.get("Type") != DRAWING_REL:
            continue
        target = relationship.get("Target")
        if not target:
            continue

        snapshot = {
            "Id": rel_id,
            "Type": DRAWING_REL,
            "Target": target,
        }
        if relationship.get("TargetMode"):
            snapshot["TargetMode"] = relationship.get("TargetMode")
        relationships_by_sheet[sheet_name] = snapshot
        drawing_roots.add(_resolve_package_target(sheet_part, target))

    drawing_parts = _collect_related_package_parts(archive, drawing_roots)
    return {
        "package_parts": {
            name: archive.read(name) for name in drawing_parts
        },
        "content_types": _discover_content_types(archive, drawing_parts),
        "worksheet_drawings": relationships_by_sheet,
    }


def patch_preserved_content_types(
    xml: bytes,
    preserved_content_types: dict[str, dict[str, str]],
) -> bytes:
    root = fromstring(xml)
    defaults = {
        item.get("Extension"): item
        for item in root.findall(f"{{{CONTENT_TYPES_NS}}}Default")
    }
    overrides = {
        item.get("PartName"): item
        for item in root.findall(f"{{{CONTENT_TYPES_NS}}}Override")
    }

    for extension, content_type in preserved_content_types.get(
        "defaults", {}
    ).items():
        item = defaults.get(extension)
        if item is None:
            item = SubElement(
                root,
                f"{{{CONTENT_TYPES_NS}}}Default",
                {"Extension": extension},
            )
            defaults[extension] = item
        item.set("ContentType", content_type)

    for part_name, content_type in preserved_content_types.get(
        "overrides", {}
    ).items():
        item = overrides.get(part_name)
        if item is None:
            item = SubElement(
                root,
                f"{{{CONTENT_TYPES_NS}}}Override",
                {"PartName": part_name},
            )
            overrides[part_name] = item
        item.set("ContentType", content_type)
    return tostring(root)


def _next_relationship_id(root: Any) -> str:
    used = {item.get("Id") for item in root}
    number = 1
    while f"rId{number}" in used:
        number += 1
    return f"rId{number}"


def restore_drawing_relationship(
    xml: bytes | None,
    snapshot: dict[str, str],
) -> tuple[bytes, str]:
    if xml is None:
        root = fromstring(
            f'<Relationships xmlns="{PACKAGE_REL_NS}"/>'.encode()
        )
    else:
        root = fromstring(xml)

    relationship = next(
        (
            item
            for item in root
            if item.get("Type") == DRAWING_REL
            and item.get("Target") == snapshot["Target"]
        ),
        None,
    )
    if relationship is None:
        relationship = next(
            (item for item in root if item.get("Type") == DRAWING_REL),
            None,
        )

    if relationship is None:
        desired_id = snapshot["Id"]
        used_ids = {item.get("Id") for item in root}
        relationship_id = (
            desired_id if desired_id not in used_ids else _next_relationship_id(root)
        )
        relationship = SubElement(
            root,
            f"{{{PACKAGE_REL_NS}}}Relationship",
            {"Id": relationship_id},
        )
    else:
        relationship_id = relationship.get("Id") or _next_relationship_id(root)
        relationship.set("Id", relationship_id)

    relationship.set("Type", snapshot["Type"])
    relationship.set("Target", snapshot["Target"])
    if snapshot.get("TargetMode"):
        relationship.set("TargetMode", snapshot["TargetMode"])
    else:
        relationship.attrib.pop("TargetMode", None)
    return tostring(root), relationship_id


def restore_worksheet_drawing(root: Any, relationship_id: str) -> None:
    rel_id_key = f"{{{DOCUMENT_REL_NS}}}id"
    drawings = root.findall(f"{{{MAIN_NS}}}drawing")
    if drawings:
        drawing = drawings[0]
        drawing.set(rel_id_key, relationship_id)
        for extra in drawings[1:]:
            root.remove(extra)
        return

    drawing = fromstring(
        (
            f'<drawing xmlns="{MAIN_NS}" '
            f'xmlns:r="{DOCUMENT_REL_NS}" r:id="{relationship_id}"/>'
        ).encode()
    )
    later_elements = {
        "legacyDrawing",
        "legacyDrawingHF",
        "picture",
        "oleObjects",
        "controls",
        "webPublishItems",
        "tableParts",
        "extLst",
    }
    for index, child in enumerate(root):
        if child.tag.rsplit("}", 1)[-1] in later_elements:
            root.insert(index, drawing)
            break
    else:
        root.append(drawing)


__all__ = [
    "discover_drawing_state",
    "patch_preserved_content_types",
    "relationship_part_name",
    "restore_drawing_relationship",
    "restore_worksheet_drawing",
]
