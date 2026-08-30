"""Preserve modern Excel threaded comments across openpyxl saves."""

from __future__ import annotations

import posixpath
from pathlib import PurePosixPath
from typing import Any
from zipfile import ZipFile

from openpyxl.xml.functions import SubElement, fromstring, tostring

from openpyxl365.ooxml_drawings import relationship_part_name

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
COMMENTS_REL = f"{DOCUMENT_REL_NS}/comments"
VML_DRAWING_REL = f"{DOCUMENT_REL_NS}/vmlDrawing"
THREADED_COMMENT_REL = (
    "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
)
PERSON_REL = (
    "http://schemas.microsoft.com/office/2017/10/relationships/person"
)


def _resolve_package_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(
        str(PurePosixPath(source_part).parent / target)
    )


def _next_relationship_id(root: Any) -> str:
    used = {item.get("Id") for item in root}
    number = 1
    while f"rId{number}" in used:
        number += 1
    return f"rId{number}"


def _collect_related_package_parts(
    archive: ZipFile,
    roots: set[str],
) -> set[str]:
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
            if target:
                pending.append(_resolve_package_target(part, target))
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


def discover_threaded_comment_state(
    archive: ZipFile,
    sheet_parts: dict[str, str],
) -> dict[str, Any]:
    """Capture threaded comments and their legacy compatibility payloads."""

    names = set(archive.namelist())
    worksheets: dict[str, dict[str, Any]] = {}
    package_roots: set[str] = set()
    rel_id_key = f"{{{DOCUMENT_REL_NS}}}id"

    for sheet_name, sheet_part in sheet_parts.items():
        relationships_part = relationship_part_name(sheet_part)
        if relationships_part not in names:
            continue
        relationships = fromstring(archive.read(relationships_part))
        if not any(
            item.get("Type") == THREADED_COMMENT_REL
            for item in relationships
        ):
            continue

        snapshots: list[dict[str, str]] = []
        for relationship in relationships:
            if relationship.get("Type") not in {
                COMMENTS_REL,
                VML_DRAWING_REL,
                THREADED_COMMENT_REL,
            }:
                continue
            target = relationship.get("Target")
            if not target:
                continue
            snapshots.append(dict(relationship.attrib))
            if relationship.get("TargetMode") != "External":
                package_roots.add(
                    _resolve_package_target(sheet_part, target)
                )

        worksheet = fromstring(archive.read(sheet_part))
        legacy_drawing = worksheet.find(f"{{{MAIN_NS}}}legacyDrawing")
        worksheets[sheet_name] = {
            "relationships": snapshots,
            "legacy_drawing_relationship_id": (
                legacy_drawing.get(rel_id_key)
                if legacy_drawing is not None
                else None
            ),
        }

    workbook_relationships = fromstring(
        archive.read("xl/_rels/workbook.xml.rels")
    )
    preserved_workbook_relationships: list[dict[str, str]] = []
    for relationship in workbook_relationships:
        if relationship.get("Type") != PERSON_REL:
            continue
        target = relationship.get("Target")
        if not target:
            continue
        preserved_workbook_relationships.append(dict(relationship.attrib))
        if relationship.get("TargetMode") != "External":
            package_roots.add(
                _resolve_package_target("xl/workbook.xml", target)
            )

    package_parts = _collect_related_package_parts(archive, package_roots)
    return {
        "worksheets": worksheets,
        "workbook_relationships": preserved_workbook_relationships,
        "package_parts": {
            name: archive.read(name) for name in package_parts
        },
        "content_types": _discover_content_types(archive, package_parts),
    }


def restore_threaded_comment_relationships(
    xml: bytes | None,
    *,
    sheet_part: str,
    snapshot: dict[str, Any],
    original_package_parts: set[str],
) -> tuple[bytes, dict[str, str], set[str]]:
    """Restore sheet comment relationships and identify generated orphan parts."""

    if xml is None:
        root = fromstring(
            f'<Relationships xmlns="{PACKAGE_REL_NS}"/>'.encode()
        )
    else:
        root = fromstring(xml)

    snapshot_vml_targets = {
        _resolve_package_target(sheet_part, item["Target"])
        for item in snapshot.get("relationships", [])
        if item.get("Type") == VML_DRAWING_REL and item.get("Target")
    }
    removed_package_parts: set[str] = set()
    for relationship in list(root):
        relationship_type = relationship.get("Type")
        target = relationship.get("Target")
        resolved_target = (
            _resolve_package_target(sheet_part, target) if target else None
        )
        generated_comment_vml = (
            relationship_type == VML_DRAWING_REL
            and resolved_target is not None
            and (
                resolved_target in snapshot_vml_targets
                or "commentsDrawing" in PurePosixPath(resolved_target).name
            )
        )
        if relationship_type in {COMMENTS_REL, THREADED_COMMENT_REL} or (
            generated_comment_vml
        ):
            if (
                resolved_target
                and resolved_target not in original_package_parts
            ):
                removed_package_parts.add(resolved_target)
                generated_rels = relationship_part_name(resolved_target)
                removed_package_parts.add(generated_rels)
            root.remove(relationship)

    relationship_ids: dict[str, str] = {}
    for relationship_snapshot in snapshot.get("relationships", []):
        desired_id = relationship_snapshot.get("Id")
        used_ids = {item.get("Id") for item in root}
        relationship_id = (
            desired_id
            if desired_id and desired_id not in used_ids
            else _next_relationship_id(root)
        )
        relationship = SubElement(
            root,
            f"{{{PACKAGE_REL_NS}}}Relationship",
            {"Id": relationship_id},
        )
        for name in ("Type", "Target", "TargetMode"):
            value = relationship_snapshot.get(name)
            if value:
                relationship.set(name, value)
        if desired_id:
            relationship_ids[desired_id] = relationship_id
    return tostring(root), relationship_ids, removed_package_parts


def restore_threaded_comment_worksheet(
    xml: bytes,
    *,
    snapshot: dict[str, Any],
    relationship_ids: dict[str, str],
) -> bytes:
    """Restore the legacyDrawing pointer used by comment compatibility data."""

    root = fromstring(xml)
    rel_id_key = f"{{{DOCUMENT_REL_NS}}}id"
    legacy_drawings = root.findall(f"{{{MAIN_NS}}}legacyDrawing")
    source_id = snapshot.get("legacy_drawing_relationship_id")
    restored_id = relationship_ids.get(source_id) if source_id else None

    if restored_id:
        if legacy_drawings:
            legacy_drawing = legacy_drawings[0]
            for extra in legacy_drawings[1:]:
                root.remove(extra)
        else:
            legacy_drawing = fromstring(
                (
                    f'<legacyDrawing xmlns="{MAIN_NS}" '
                    f'xmlns:r="{DOCUMENT_REL_NS}" r:id="{restored_id}"/>'
                ).encode()
            )
            extension_list = root.find(f"{{{MAIN_NS}}}extLst")
            if extension_list is None:
                root.append(legacy_drawing)
            else:
                root.insert(list(root).index(extension_list), legacy_drawing)
        legacy_drawing.set(rel_id_key, restored_id)
    else:
        for legacy_drawing in legacy_drawings:
            root.remove(legacy_drawing)
    return tostring(root)


__all__ = [
    "COMMENTS_REL",
    "PERSON_REL",
    "THREADED_COMMENT_REL",
    "VML_DRAWING_REL",
    "discover_threaded_comment_state",
    "restore_threaded_comment_relationships",
    "restore_threaded_comment_worksheet",
]
