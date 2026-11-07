"""Edit existing Excel 365 workbooks with openpyxl dynamic-array support.

openpyxl can read and write array formula text but currently drops the OOXML
cell metadata that tells Excel a formula is a dynamic (spilling) array.  This
module keeps openpyxl as the workbook editor, preserves pre-existing metadata
(including rich-value metadata and its Rich Data package parts), x14 data
validations, and threaded comments.  It also restores dynamic-array anchors
and delegates unsupported DrawingML preservation to
:mod:`openpyxl365.ooxml_drawings`.

The module does not calculate formulae.  Open the saved file in Microsoft
Excel and recalculate it before treating cached values as current.
"""

from __future__ import annotations

import errno
import os
import re
import shutil
import sys
import tempfile
from io import BytesIO
from pathlib import Path, PurePosixPath
from typing import Any
from urllib.parse import urlsplit
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.reader.strings import read_string_table
from openpyxl.cell.rich_text import CellRichText
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.packaging.relationship import Relationship
from openpyxl.workbook.properties import CalcProperties
from openpyxl.workbook.external_link.external import (
    ExternalLink, ExternalBook, ExternalSheetNames,
)
from openpyxl.xml.functions import SubElement, fromstring, tostring

from openpyxl365.ooxml_drawings import (
    discover_drawing_state,
    patch_preserved_content_types,
    relationship_part_name,
    restore_drawing_relationship,
    restore_worksheet_drawing,
)
from openpyxl365.ooxml_data_validations import (
    discover_extended_data_validations,
    restore_extended_data_validations,
)
from openpyxl365.ooxml_threaded_comments import (
    discover_threaded_comment_state,
    restore_threaded_comment_relationships,
    restore_threaded_comment_worksheet,
)


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SHEET_METADATA_REL = f"{DOCUMENT_REL_NS}/sheetMetadata"
SHEET_METADATA_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml."
    "sheetMetadata+xml"
)
RICH_DATA_PART_PREFIX = "xl/richData/"
PYTHON_REL = "http://schemas.microsoft.com/office/2023/09/relationships/Python"

_DYNAMIC_FUNCTIONS = {
    "ANCHORARRAY",
    "BYCOL",
    "BYROW",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DROP",
    "EXPAND",
    "FILTER",
    "HSTACK",
    "LAMBDA",
    "LET",
    "MAKEARRAY",
    "MAP",
    "RANDARRAY",
    "REDUCE",
    "SCAN",
    "SEQUENCE",
    "SINGLE",
    "SORT",
    "SORTBY",
    "SWITCH",
    "TAKE",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "UNIQUE",
    "VSTACK",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
}

_FUTURE_FUNCTIONS = _DYNAMIC_FUNCTIONS | {
    "ARRAYTOTEXT",
    "TEXTAFTER",
    "TEXTBEFORE",
    "VALUETOTEXT",
    "XMATCH",
}

_XLWS_FUNCTIONS = {"FILTER", "SORT"}

_SPILL_REFERENCE = re.compile(
    r"(?P<ref>"
    r"(?:(?:'(?:[^']|'')*'|[A-Za-z_][A-Za-z0-9_.]*|"
    r"\[[^\]]+\][A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"\$?[A-Za-z]{1,3}\$?\d+"
    r")#"
)

_METADATA_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
 <metadataTypes count="1">
  <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1"
   pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1"
   clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>
 </metadataTypes>
 <futureMetadata name="XLDAPR" count="1">
  <bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">
   <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>
  </ext></extLst></bk>
 </futureMetadata>
 <cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>
</metadata>
"""


class UnsupportedWorkbookMetadataError(ValueError):
    """Raised when existing metadata cannot represent a new dynamic array."""


def _replace_spill_operators(formula: str) -> str:
    """Convert formula-bar ``A1#`` syntax without changing quoted strings."""

    chunks: list[str] = []
    start = 0
    index = 0
    length = len(formula)

    while index < length:
        if formula[index] != '"':
            index += 1
            continue

        chunks.append(
            _SPILL_REFERENCE.sub(
                lambda match: f"ANCHORARRAY({match.group('ref')})",
                formula[start:index],
            )
        )
        quote_start = index
        index += 1
        while index < length:
            if formula[index] != '"':
                index += 1
                continue
            if index + 1 < length and formula[index + 1] == '"':
                index += 2
                continue
            index += 1
            break
        chunks.append(formula[quote_start:index])
        start = index

    chunks.append(
        _SPILL_REFERENCE.sub(
            lambda match: f"ANCHORARRAY({match.group('ref')})",
            formula[start:],
        )
    )
    return "".join(chunks)


def _base_function_name(name: str) -> str:
    upper = name.upper()
    for prefix in ("_XLFN._XLWS.", "_XLFN.", "_XLWS."):
        if upper.startswith(prefix):
            return upper[len(prefix) :]
    return upper


def normalize_excel365_formula(formula: str) -> tuple[str, set[str]]:
    """Return storage-safe formula text and the recognised function names.

    The returned formula includes the leading equals sign.  Modern functions
    receive their required ``_xlfn.`` prefix and spill references such as
    ``D1#`` become ``_xlfn.ANCHORARRAY(D1)``.
    """

    if not isinstance(formula, str) or not formula.startswith("="):
        raise ValueError("Excel formula must be a string beginning with '='")

    prepared = _replace_spill_operators(formula)
    tokenizer = Tokenizer(prepared)
    functions: set[str] = set()
    pieces = ["="]

    for token in tokenizer.items:
        value = token.value
        if token.type == Token.FUNC and token.subtype == Token.OPEN:
            name = value[:-1]
            base = _base_function_name(name)
            functions.add(base)
            if base in _FUTURE_FUNCTIONS and not name.upper().startswith("_XLFN."):
                if base in _XLWS_FUNCTIONS:
                    value = f"_xlfn._xlws.{name}("
                else:
                    value = f"_xlfn.{name}("
        pieces.append(value)

    return "".join(pieces), functions


CALCULATION_ATTRIBUTES = (
    "calcId", "calcMode", "fullCalcOnLoad", "refMode", "iterate", "iterateCount",
    "iterateDelta", "fullPrecision", "calcCompleted", "calcOnSave",
    "concurrentCalc", "concurrentManualCount", "forceFullCalc",
)


def _sync_calculation_properties(path: Path, workbook: Workbook) -> None:
    """Report only calculation attributes actually present in the source file."""

    try:
        with ZipFile(path) as archive:
            root = fromstring(archive.read("xl/workbook.xml"))
    except KeyError:
        return
    element = root.find(f"{{{MAIN_NS}}}calcPr")
    present = set(element.attrib) if element is not None else set()
    calculation = workbook.calculation
    if calculation is None:
        return
    for name in CALCULATION_ATTRIBUTES:
        if name not in present and hasattr(calculation, name):
            setattr(calculation, name, None)


def _normalise_package_target(target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return str(PurePosixPath("xl") / target)


def _sheet_part_map(archive: ZipFile) -> dict[str, str]:
    workbook = fromstring(archive.read("xl/workbook.xml"))
    rels = fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.get("Id"): _normalise_package_target(rel.get("Target"))
        for rel in rels
        if rel.get("Type", "").endswith("/worksheet")
    }
    result: dict[str, str] = {}
    rel_id_key = f"{{{DOCUMENT_REL_NS}}}id"
    for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
        rel_id = sheet.get(rel_id_key)
        if rel_id in targets:
            result[sheet.get("name")] = targets[rel_id]
    return result


def _source_metadata_types(archive: ZipFile) -> set[str]:
    if "xl/metadata.xml" not in archive.namelist():
        return set()
    root = fromstring(archive.read("xl/metadata.xml"))
    return {
        item.get("name")
        for item in root.findall(f".//{{{MAIN_NS}}}metadataType")
        if item.get("name")
    }


def _dynamic_metadata_indices(metadata_xml: bytes | None) -> set[str]:
    """Return the 1-based cellMetadata records that point to XLDAPR."""

    if metadata_xml is None:
        return set()

    root = fromstring(metadata_xml)
    type_names = [
        item.get("name")
        for item in root.findall(
            f"./{{{MAIN_NS}}}metadataTypes/{{{MAIN_NS}}}metadataType"
        )
    ]
    dynamic_type_numbers = {
        str(index)
        for index, name in enumerate(type_names, start=1)
        if name == "XLDAPR"
    }
    cell_metadata = root.find(f"{{{MAIN_NS}}}cellMetadata")
    if cell_metadata is None:
        return set()

    result: set[str] = set()
    for index, block in enumerate(cell_metadata, start=1):
        if any(
            record.get("t") in dynamic_type_numbers
            for record in block.findall(f"{{{MAIN_NS}}}rc")
        ):
            result.add(str(index))
    return result


def _discover_content_types_for_parts(
    archive: ZipFile,
    preserved_parts: set[str],
) -> dict[str, dict[str, str]]:
    """Capture content-type declarations required by preserved OOXML parts."""

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


def _merge_content_type_snapshots(
    *snapshots: dict[str, dict[str, str]],
) -> dict[str, dict[str, str]]:
    merged = {"defaults": {}, "overrides": {}}
    for snapshot in snapshots:
        merged["defaults"].update(snapshot.get("defaults", {}))
        merged["overrides"].update(snapshot.get("overrides", {}))
    return merged


def _discover_package_state(path: Path) -> dict[str, Any]:
    """Capture OOXML state that openpyxl does not round-trip itself."""

    with ZipFile(path) as archive:
        # openpyxl parses external links but rewrites their cached sheet data.
        # Some real-world XLA links (notably iFinD) are then rejected by Excel.
        # Formula editing does not require touching these parts, so keep them
        # byte-for-byte from the source package.
        preserved_package_parts = {
            name: archive.read(name)
            for name in archive.namelist()
            if name.startswith("xl/externalLinks/")
        }
        sheet_parts = _sheet_part_map(archive)
        drawing_state = discover_drawing_state(archive, sheet_parts)
        extended_data_validations = discover_extended_data_validations(
            archive, sheet_parts
        )
        threaded_comment_state = discover_threaded_comment_state(
            archive, sheet_parts
        )
        preserved_package_parts.update(drawing_state["package_parts"])
        preserved_package_parts.update(
            threaded_comment_state["package_parts"]
        )
        shared_strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_strings = read_string_table(
                BytesIO(archive.read("xl/sharedStrings.xml"))
            )
        rich_data_parts = {
            name
            for name in archive.namelist()
            if name.startswith(RICH_DATA_PART_PREFIX)
        }
        preserved_package_parts.update(
            {name: archive.read(name) for name in rich_data_parts}
        )
        workbook_relationships = fromstring(
            archive.read("xl/_rels/workbook.xml.rels")
        )
        # PY() references code/environment definitions outside the worksheet.
        # Preserve that payload and its relationship alongside Rich Data.
        python_parts = set()
        for relationship in workbook_relationships:
            if (
                relationship.get("Type") == PYTHON_REL
                and relationship.get("TargetMode") != "External"
            ):
                target = _normalise_package_target(relationship.get("Target", ""))
                if target in archive.namelist():
                    python_parts.add(target)
        preserved_package_parts.update(
            {name: archive.read(name) for name in python_parts}
        )
        preserved_workbook_relationships: list[dict[str, str]] = []
        for relationship in workbook_relationships:
            target = relationship.get("Target")
            if (
                not target
                or relationship.get("TargetMode") == "External"
                or not (
                    _normalise_package_target(target).startswith(RICH_DATA_PART_PREFIX)
                    or _normalise_package_target(target) in python_parts
                )
            ):
                continue
            preserved_workbook_relationships.append(dict(relationship.attrib))
        preserved_workbook_relationships.extend(
            threaded_comment_state["workbook_relationships"]
        )
        rich_data_content_types = _discover_content_types_for_parts(
            archive, rich_data_parts | python_parts
        )
        metadata_xml = (
            archive.read("xl/metadata.xml")
            if "xl/metadata.xml" in archive.namelist()
            else None
        )
        metadata_types = _source_metadata_types(archive)
        dynamic_indices = _dynamic_metadata_indices(metadata_xml)
        dynamic_cells: dict[str, set[str]] = {}
        cell_attributes: dict[str, dict[str, dict[str, str]]] = {}
        formula_texts: dict[str, dict[str, str]] = {}
        formula_attributes: dict[str, dict[str, dict[str, Any]]] = {}

        for sheet_name, part in sheet_parts.items():
            root = fromstring(archive.read(part))
            sheet_dynamic: set[str] = set()
            sheet_cell_attributes: dict[str, dict[str, str]] = {}
            sheet_formula_texts: dict[str, str] = {}
            sheet_formula_attributes: dict[str, dict[str, Any]] = {}
            for cell in root.findall(f".//{{{MAIN_NS}}}c"):
                coordinate = cell.get("r")
                if not coordinate:
                    continue
                preserved = {
                    name: cell.get(name)
                    for name in ("cm", "vm")
                    if cell.get(name) is not None
                }
                if preserved:
                    sheet_cell_attributes[coordinate] = preserved

                formula = cell.find(f"{{{MAIN_NS}}}f")
                if formula is None:
                    continue
                sheet_formula_texts[coordinate] = f"={formula.text or ''}"
                cached_value = cell.find(f"{{{MAIN_NS}}}v")
                sheet_formula_attributes[coordinate] = {
                    "text": formula.text or "",
                    "attributes": dict(formula.attrib),
                    "cell_type": cell.get("t"),
                    "has_cached_value": cached_value is not None,
                    "cached_value": (
                        cached_value.text if cached_value is not None else None
                    ),
                }
                snapshot = sheet_formula_attributes[coordinate]
                # openpyxl writes inline strings and drops sharedStrings.xml.
                # A restored formula cache must contain text, not an orphaned
                # index into the source workbook's string table.
                if cell.get("t") == "s" and cached_value is not None:
                    snapshot["cached_value"] = shared_strings[int(cached_value.text)]
                    snapshot["cell_type"] = "str"
                if cell.get("cm") in dynamic_indices:
                    sheet_dynamic.add(coordinate)

            if sheet_dynamic:
                dynamic_cells[sheet_name] = sheet_dynamic
            if sheet_cell_attributes:
                cell_attributes[sheet_name] = sheet_cell_attributes
            if sheet_formula_texts:
                formula_texts[sheet_name] = sheet_formula_texts
            if sheet_formula_attributes:
                formula_attributes[sheet_name] = sheet_formula_attributes

        return {
            "preserved_package_parts": preserved_package_parts,
            "preserved_content_types": _merge_content_type_snapshots(
                drawing_state["content_types"],
                rich_data_content_types,
                threaded_comment_state["content_types"],
            ),
            "preserved_workbook_relationships": (
                preserved_workbook_relationships
            ),
            "worksheet_drawings": drawing_state["worksheet_drawings"],
            "extended_data_validations": extended_data_validations,
            "threaded_comment_state": threaded_comment_state,
            "metadata_xml": metadata_xml,
            "metadata_types": metadata_types,
            "dynamic_indices": dynamic_indices,
            "dynamic_cells": dynamic_cells,
            "cell_attributes": cell_attributes,
            "formula_texts": formula_texts,
            "formula_attributes": formula_attributes,
        }


def _next_relationship_id(root: Any) -> str:
    used = {item.get("Id") for item in root}
    number = 1
    while f"rId{number}" in used:
        number += 1
    return f"rId{number}"


def _patch_relationships(
    xml: bytes,
    *,
    include_metadata: bool,
    preserved_workbook_relationships: list[dict[str, str]],
) -> bytes:
    root = fromstring(xml)
    if include_metadata:
        metadata_relationship = next(
            (
                relationship
                for relationship in root
                if relationship.get("Type") == SHEET_METADATA_REL
            ),
            None,
        )
        if metadata_relationship is None:
            metadata_relationship = SubElement(
                root,
                f"{{{PACKAGE_REL_NS}}}Relationship",
                {"Id": _next_relationship_id(root)},
            )
        metadata_relationship.set("Type", SHEET_METADATA_REL)
        metadata_relationship.set("Target", "metadata.xml")
        metadata_relationship.attrib.pop("TargetMode", None)

    for snapshot in preserved_workbook_relationships:
        relationship_type = snapshot.get("Type")
        target = snapshot.get("Target")
        if not relationship_type or not target:
            continue
        resolved_target = _normalise_package_target(target)
        relationship = next(
            (
                item
                for item in root
                if item.get("Type") == relationship_type
                and item.get("Target")
                and _normalise_package_target(item.get("Target"))
                == resolved_target
            ),
            None,
        )
        if relationship is None:
            desired_id = snapshot.get("Id")
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
        relationship.set("Type", relationship_type)
        relationship.set("Target", target)
        if snapshot.get("TargetMode"):
            relationship.set("TargetMode", snapshot["TargetMode"])
        else:
            relationship.attrib.pop("TargetMode", None)
    return tostring(root)


def _patch_content_types(
    xml: bytes,
    *,
    include_metadata: bool,
    preserved_content_types: dict[str, dict[str, str]],
    removed_package_parts: set[str],
) -> bytes:
    xml = patch_preserved_content_types(xml, preserved_content_types)
    root = fromstring(xml)
    removed_part_names = {f"/{name.lstrip('/')}" for name in removed_package_parts}
    for item in list(root.findall(f"{{{CONTENT_TYPES_NS}}}Override")):
        if item.get("PartName") in removed_part_names:
            root.remove(item)
    if not include_metadata:
        return tostring(root)

    overrides = {
        item.get("PartName"): item
        for item in root.findall(f"{{{CONTENT_TYPES_NS}}}Override")
    }
    item = overrides.get("/xl/metadata.xml")
    if item is None:
        item = SubElement(
            root,
            f"{{{CONTENT_TYPES_NS}}}Override",
            {"PartName": "/xl/metadata.xml"},
        )
    item.set("ContentType", SHEET_METADATA_CONTENT_TYPE)
    return tostring(root)


def _patch_worksheet(
    xml: bytes,
    dynamic_coordinates: set[str],
    dynamic_metadata_index: str,
    preserved_cell_attributes: dict[str, dict[str, str]],
    preserved_formula_attributes: dict[str, dict[str, Any]],
    invalidated_formula_caches: set[str],
    drawing_relationship_id: str | None,
) -> bytes:
    root = fromstring(xml)
    # Rich-text runs containing spaces require xml:space. openpyxl omits it
    # on some runs, and Excel repairs those strings by stripping the spaces.
    for text_node in root.findall(f".//{{{MAIN_NS}}}is//{{{MAIN_NS}}}t"):
        text = text_node.text
        if text and text != text.strip():
            text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    cells = {
        cell.get("r"): cell
        for cell in root.findall(f".//{{{MAIN_NS}}}c")
        if cell.get("r")
    }
    for coordinate, attributes in preserved_cell_attributes.items():
        cell = cells.get(coordinate)
        if cell is None:
            continue
        for name, value in attributes.items():
            formula = cell.find(f"{{{MAIN_NS}}}f")
            snapshot = preserved_formula_attributes.get(coordinate)
            if snapshot is not None:
                no_array = formula is None or (
                    formula.get("t") != "array" and coordinate not in dynamic_coordinates
                )
                changed_formula = formula is None or (formula.text or "") != snapshot.get(
                    "loaded_text", snapshot["text"]
                )
                if (name == "cm" and no_array) or (name == "vm" and changed_formula):
                    # Neither the array marker nor its old rich/error value
                    # metadata belongs on a replacement scalar/formula.
                    cell.attrib.pop(name, None)
                    continue
            cell.set(name, value)

    # A shared formula is one OOXML group: one master carries the formula and
    # range, while followers carry only the shared index. If any member was
    # edited, restoring the remaining members independently can leave an
    # orphaned group (followers without a master). In that case keep
    # openpyxl's expanded ordinary formulas for the whole group instead.
    shared_groups: dict[str, list[tuple[str, dict[str, Any]]]] = {}
    for coordinate, snapshot in preserved_formula_attributes.items():
        attributes = snapshot.get("attributes", {})
        if attributes.get("t") == "shared" and attributes.get("si") is not None:
            shared_groups.setdefault(attributes["si"], []).append(
                (coordinate, snapshot)
            )
    restorable_shared_groups = {
        shared_index
        for shared_index, members in shared_groups.items()
        if all(
            (cell := cells.get(coordinate)) is not None
            and (formula := cell.find(f"{{{MAIN_NS}}}f")) is not None
            and (formula.text or "")
            == snapshot.get("loaded_text", snapshot["text"])
            for coordinate, snapshot in members
        )
    }

    # openpyxl retains array formula text/ref but drops optional attributes
    # such as aca/ca.  Restore them only when the formula itself is unchanged.
    for coordinate, snapshot in preserved_formula_attributes.items():
        cell = cells.get(coordinate)
        if cell is None:
            continue
        formula = cell.find(f"{{{MAIN_NS}}}f")
        expected_text = snapshot.get("loaded_text", snapshot["text"])
        if formula is None or (formula.text or "") != expected_text:
            continue
        source_attributes = snapshot["attributes"]
        shared_index = (
            source_attributes.get("si")
            if source_attributes.get("t") == "shared"
            else None
        )
        if shared_index is None or shared_index in restorable_shared_groups:
            formula.attrib.clear()
            formula.attrib.update(source_attributes)
            formula.text = snapshot["text"] or None
        if coordinate in invalidated_formula_caches:
            continue

        cell_type = snapshot.get("cell_type")
        if cell_type is None:
            cell.attrib.pop("t", None)
        else:
            cell.set("t", cell_type)

        cached_value = cell.find(f"{{{MAIN_NS}}}v")
        if snapshot.get("has_cached_value"):
            if cached_value is None:
                cached_value = SubElement(cell, f"{{{MAIN_NS}}}v")
            cached_value.text = snapshot.get("cached_value")
        elif cached_value is not None:
            cell.remove(cached_value)

    for coordinate in dynamic_coordinates:
        cell = cells.get(coordinate)
        if cell is None:
            raise ValueError(f"Dynamic-array anchor is missing from worksheet: {coordinate}")
        formula = cell.find(f"{{{MAIN_NS}}}f")
        if formula is None:
            raise ValueError(f"Dynamic-array anchor has no formula: {coordinate}")
        cell.set("cm", dynamic_metadata_index)
        formula.set("t", "array")
        formula.set("ref", coordinate)
    shared = [
        formula
        for cell in cells.values()
        if (formula := cell.find(f"{{{MAIN_NS}}}f")) is not None
        and formula.get("t") == "shared"
    ]
    shared_indices = {formula.get("si") for formula in shared}
    shared_masters = {
        formula.get("si")
        for formula in shared
        if formula.get("ref") and (formula.text or "").strip()
    }
    if orphaned := sorted(shared_indices - shared_masters):
        raise RuntimeError(f"Worksheet contains orphaned shared-formula groups: {orphaned}")
    if drawing_relationship_id is not None:
        restore_worksheet_drawing(root, drawing_relationship_id)
    return tostring(root)


def _rewrite_package(
    source: Path,
    destination: Path,
    dynamic_cells: dict[str, set[str]],
    *,
    metadata_xml: bytes | None,
    dynamic_metadata_index: str,
    preserved_cell_attributes: dict[str, dict[str, dict[str, str]]],
    preserved_formula_attributes: dict[str, dict[str, dict[str, Any]]],
    preserved_package_parts: dict[str, bytes],
    preserved_content_types: dict[str, dict[str, str]],
    preserved_workbook_relationships: list[dict[str, str]],
    worksheet_drawings: dict[str, dict[str, str]],
    extended_data_validations: dict[str, list[bytes]],
    threaded_comment_state: dict[str, Any],
    invalidated_formula_caches: dict[str, set[str]],
) -> None:
    with ZipFile(source, "r") as archive:
        sheet_parts = _sheet_part_map(archive)
        replacements: dict[str, bytes] = dict(preserved_package_parts)
        if metadata_xml is not None or preserved_workbook_relationships:
            replacements["xl/_rels/workbook.xml.rels"] = _patch_relationships(
                archive.read("xl/_rels/workbook.xml.rels"),
                include_metadata=metadata_xml is not None,
                preserved_workbook_relationships=(
                    preserved_workbook_relationships
                ),
            )
        if metadata_xml is not None:
            replacements.update(
                {
                    "xl/metadata.xml": metadata_xml,
                }
            )
        archive_names = set(archive.namelist())
        threaded_worksheets = threaded_comment_state.get("worksheets", {})
        threaded_relationship_ids: dict[str, dict[str, str]] = {}
        removed_package_parts: set[str] = set()
        original_threaded_parts = set(
            threaded_comment_state.get("package_parts", {})
        )
        for sheet_name, snapshot in threaded_worksheets.items():
            if sheet_name not in sheet_parts:
                raise ValueError(f"Worksheet missing after save: {sheet_name}")
            sheet_part = sheet_parts[sheet_name]
            relationships_part = relationship_part_name(sheet_part)
            relationships_xml = (
                archive.read(relationships_part)
                if relationships_part in archive_names
                else None
            )
            patched, relationship_ids, removed = (
                restore_threaded_comment_relationships(
                    relationships_xml,
                    sheet_part=sheet_part,
                    snapshot=snapshot,
                    original_package_parts=original_threaded_parts,
                )
            )
            replacements[relationships_part] = patched
            threaded_relationship_ids[sheet_name] = relationship_ids
            removed_package_parts.update(removed)

        drawing_relationship_ids: dict[str, str] = {}
        for sheet_name, snapshot in worksheet_drawings.items():
            if sheet_name not in sheet_parts:
                raise ValueError(f"Worksheet missing after save: {sheet_name}")
            relationships_part = relationship_part_name(sheet_parts[sheet_name])
            relationships_xml = (
                replacements.get(relationships_part)
                or archive.read(relationships_part)
                if relationships_part in archive_names
                else replacements.get(relationships_part)
            )
            patched, relationship_id = restore_drawing_relationship(
                relationships_xml, snapshot
            )
            replacements[relationships_part] = patched
            drawing_relationship_ids[sheet_name] = relationship_id

        if (
            metadata_xml is not None
            or any(preserved_content_types.values())
            or removed_package_parts
        ):
            replacements["[Content_Types].xml"] = _patch_content_types(
                archive.read("[Content_Types].xml"),
                include_metadata=metadata_xml is not None,
                preserved_content_types=preserved_content_types,
                removed_package_parts=removed_package_parts,
            )

        sheets_to_patch = (
            set(dynamic_cells)
            | set(preserved_cell_attributes)
            | set(preserved_formula_attributes)
            | set(invalidated_formula_caches)
            | set(drawing_relationship_ids)
            | set(extended_data_validations)
            | set(threaded_worksheets)
        )
        for sheet_name in sheets_to_patch:
            if sheet_name not in sheet_parts:
                raise ValueError(f"Worksheet missing after save: {sheet_name}")
            part = sheet_parts[sheet_name]
            patched_worksheet = _patch_worksheet(
                archive.read(part),
                dynamic_cells.get(sheet_name, set()),
                dynamic_metadata_index,
                preserved_cell_attributes.get(sheet_name, {}),
                preserved_formula_attributes.get(sheet_name, {}),
                invalidated_formula_caches.get(sheet_name, set()),
                drawing_relationship_ids.get(sheet_name),
            )
            patched_worksheet = restore_extended_data_validations(
                patched_worksheet,
                extended_data_validations.get(sheet_name, []),
            )
            if sheet_name in threaded_worksheets:
                patched_worksheet = restore_threaded_comment_worksheet(
                    patched_worksheet,
                    snapshot=threaded_worksheets[sheet_name],
                    relationship_ids=threaded_relationship_ids.get(
                        sheet_name, {}
                    ),
                )
            replacements[part] = patched_worksheet

        with ZipFile(destination, "w", ZIP_DEFLATED, allowZip64=True) as output:
            written: set[str] = set()
            for info in archive.infolist():
                if info.filename in removed_package_parts:
                    continue
                data = replacements.get(info.filename, archive.read(info.filename))
                output.writestr(info, data)
                written.add(info.filename)
            for name, data in replacements.items():
                if name not in written:
                    output.writestr(name, data)


class Excel365Workbook:
    """An openpyxl workbook with safe Excel 365 dynamic-array saves."""

    def __init__(
        self,
        workbook: Workbook,
        *,
        source_path: Path | None = None,
        package_state: dict[str, Any] | None = None,
    ) -> None:
        self.workbook = workbook
        self.source_path = source_path
        package_state = package_state or {}
        self._source_metadata_xml: bytes | None = package_state.get("metadata_xml")
        self._preserved_package_parts: dict[str, bytes] = dict(
            package_state.get("preserved_package_parts", {})
        )
        self._preserved_content_types: dict[str, dict[str, str]] = {
            name: dict(values)
            for name, values in package_state.get(
                "preserved_content_types", {}
            ).items()
        }
        self._preserved_workbook_relationships: list[dict[str, str]] = [
            dict(values)
            for values in package_state.get(
                "preserved_workbook_relationships", []
            )
        ]
        self._worksheet_drawings: dict[str, dict[str, str]] = {
            name: dict(values)
            for name, values in package_state.get("worksheet_drawings", {}).items()
        }
        self._extended_data_validations: dict[str, list[bytes]] = {
            name: list(values)
            for name, values in package_state.get(
                "extended_data_validations", {}
            ).items()
        }
        self._threaded_comment_state: dict[str, Any] = dict(
            package_state.get("threaded_comment_state", {})
        )
        self._source_metadata_types: set[str] = set(
            package_state.get("metadata_types", set())
        )
        self._source_dynamic_indices: set[str] = set(
            package_state.get("dynamic_indices", set())
        )
        self._preserved_cell_attributes = package_state.get("cell_attributes", {})
        self._preserved_formula_attributes = package_state.get(
            "formula_attributes", {}
        )
        # This map intentionally contains only formulas explicitly added in
        # the current editing session. Existing Excel spill anchors are kept
        # by the preserved cell/formula metadata below; treating them as new
        # arrays would erase their calculated spill children on an unrelated
        # edit.
        self._dynamic_by_worksheet_id: dict[int, set[str]] = {}
        self._invalidated_formula_caches_by_worksheet_id: dict[int, set[str]] = {}
        self._source_formula_texts_by_worksheet_id: dict[int, dict[str, str]] = {}
        for sheet_name, formula_texts in package_state.get(
            "formula_texts", {}
        ).items():
            if sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
                loaded_formula_texts: dict[str, str] = {}
                for coordinate in formula_texts:
                    snapshot = self._preserved_formula_attributes.get(
                        sheet_name, {}
                    ).get(coordinate, {})
                    if (
                        not snapshot.get("text")
                        and snapshot.get("attributes", {}).get("t") in {None, "normal"}
                    ):
                        # Excel emits empty <f ca="1"/> on some spill children.
                        # These are cached outputs, not an executable "=" formula.
                        cached = snapshot.get("cached_value")
                        kind = snapshot.get("cell_type")
                        if cached is not None and kind not in {"str", "e", "s"}:
                            cached = bool(int(cached)) if kind == "b" else float(cached)
                        worksheet[coordinate].value = cached
                        continue
                    value = worksheet[coordinate].value
                    if isinstance(value, ArrayFormula):
                        formula = value.text
                    elif isinstance(value, str) and value.startswith("="):
                        formula = value
                    else:
                        continue
                    # Shared-formula followers have an empty <f/> in the raw
                    # OOXML, while openpyxl expands them to their translated
                    # formula text.  Snapshot the loaded representation so an
                    # unrelated edit does not make those followers look new.
                    loaded_formula_texts[coordinate] = normalize_excel365_formula(
                        formula
                    )[0]
                    snapshot = self._preserved_formula_attributes.get(
                        sheet_name, {}
                    ).get(coordinate)
                    if snapshot is not None:
                        loaded_text = loaded_formula_texts[coordinate]
                        snapshot["loaded_text"] = (
                            loaded_text[1:]
                            if loaded_text.startswith("=")
                            else loaded_text
                        )
                self._source_formula_texts_by_worksheet_id[id(worksheet)] = (
                    loaded_formula_texts
                )
        self._source_array_outputs = {}
        for sheet_name, snapshots in self._preserved_formula_attributes.items():
            if sheet_name not in workbook.sheetnames:
                continue
            ws = workbook[sheet_name]
            for coordinate, snapshot in snapshots.items():
                attributes = snapshot["attributes"]
                if attributes.get("t") != "array" or not attributes.get("ref"):
                    continue
                c1, r1, c2, r2 = range_boundaries(attributes["ref"])
                if c1 == c2 and r1 == r2:
                    continue
                if (c2 - c1 + 1) * (r2 - r1 + 1) <= len(ws._cells):
                    cells = (ws._cells.get((row, col))
                             for row in range(r1, r2 + 1)
                             for col in range(c1, c2 + 1))
                else:
                    cells = (cell for cell in ws._cells.values()
                             if r1 <= cell.row <= r2 and c1 <= cell.column <= c2)
                children = {cell.coordinate: cell.value for cell in cells
                            if cell is not None and cell.coordinate != coordinate}
                self._source_array_outputs[id(ws), coordinate] = children

    @classmethod
    def open(cls, filename: str | os.PathLike[str], **kwargs: Any) -> "Excel365Workbook":
        path = Path(filename).expanduser().resolve()
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Excel365Workbook supports .xlsx and .xlsm files")
        if kwargs.get("data_only"):
            raise ValueError("data_only=True would discard formulas needed for editing")
        if kwargs.get("read_only"):
            raise ValueError("read_only=True workbooks cannot be saved")

        kwargs.setdefault("data_only", False)
        kwargs.setdefault("read_only", False)
        kwargs.setdefault("keep_links", True)
        kwargs.setdefault("rich_text", True)
        if path.suffix.lower() == ".xlsm":
            kwargs.setdefault("keep_vba", True)

        package_state = _discover_package_state(path)
        workbook = load_workbook(path, **kwargs)
        _sync_calculation_properties(path, workbook)
        return cls(
            workbook,
            source_path=path,
            package_state=package_state,
        )

    @classmethod
    def create(cls) -> "Excel365Workbook":
        return cls(Workbook())

    @property
    def wb(self) -> Workbook:
        """Short alias for callers already accustomed to an openpyxl workbook."""

        return self.workbook

    @staticmethod
    def _validate_cloud_workbook_url(url: str) -> str:
        """Validate an explicit workbook URL; never infer an account or path."""
        parsed = urlsplit(url)
        if (parsed.scheme != "https" or not parsed.hostname or parsed.username
                or parsed.password or parsed.query or parsed.fragment
                or "\\" in url or any(ord(c) < 32 for c in url)):
            raise ValueError("Use a full HTTPS workbook URL, not a sharing URL or local path")
        if not parsed.path.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".xls")):
            raise ValueError("Cloud URL must identify a workbook file")
        if parsed.hostname.lower() == "d.docs.live.net":
            if not re.match(r"^/[0-9a-fA-F]{16}/.+", parsed.path):
                raise ValueError("Personal OneDrive URL requires a verified 16-digit CID")
            if re.search(r"/(?:Mac/Home|Library/CloudStorage)/", parsed.path, re.I):
                raise ValueError("OneDrive cloud URL contains a local Mac path")
        return url

    def external_link_targets(self) -> list[dict[str, Any]]:
        """Return formula link indices and all saved path alternatives, read-only."""
        result = []
        for index, link in enumerate(self.workbook._external_links, 1):
            part = f"xl/externalLinks/_rels/externalLink{index}.xml.rels"
            raw = self._preserved_package_parts.get(part)
            targets = ([r.get("Target") for r in fromstring(raw)
                        if r.get("Type") == f"{DOCUMENT_REL_NS}/externalLinkPath"]
                       if raw else [link.file_link.Target])
            result.append({"index": index, "targets": targets})
        return result

    def set_external_link_target(self, index: int, source_url: str) -> None:
        """Relocate the same source workbook to HTTPS, preserving formulas/caches.

        All externalLinkPath alternatives are replaced, including Excel's local
        fallback paths. This does not refresh data or verify cloud access. Do not
        use retained caches as evidence that a different source contains them.
        """
        source_url = self._validate_cloud_workbook_url(source_url)
        if isinstance(index, bool) or not isinstance(index, int) or not (
            1 <= index <= len(self.workbook._external_links)
        ):
            raise ValueError("External link index must identify an existing workbook link")
        link = self.workbook._external_links[index - 1]
        if link.externalBook is None or link.file_link is None:
            raise ValueError("Only external workbook links can be relocated")
        part = f"xl/externalLinks/_rels/externalLink{index}.xml.rels"
        raw = self._preserved_package_parts.get(part)
        if raw:
            root = fromstring(raw)
            for relationship in root:
                if relationship.get("Type") == f"{DOCUMENT_REL_NS}/externalLinkPath":
                    relationship.set("Target", source_url)
                    relationship.set("TargetMode", "External")
            self._preserved_package_parts[part] = tostring(root)
        link.file_link.Target = source_url
        link.file_link.TargetMode = "External"

    def set_external_reference(
        self, sheet: str | Any, coordinate: str, *, source_url: str,
        source_sheet: str, source_cell: str,
    ) -> int:
        """Write a single-cell cloud reference and its real OOXML link records.

        Returns the 1-based link index. Reuses a source URL when possible, without
        guessing OneDrive identities. New formulas have no calculated cache.
        """
        source_url = self._validate_cloud_workbook_url(source_url)
        if not source_sheet or len(source_sheet) > 31 or re.search(r"[\[\]:*?/\\]", source_sheet):
            raise ValueError("Invalid source worksheet name")
        if not re.fullmatch(r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*", source_cell):
            raise ValueError("source_cell must be one A1-style cell reference")
        c1, r1, _, _ = range_boundaries(source_cell)
        if c1 > 16384 or r1 > 1048576:
            raise ValueError("source_cell exceeds Excel worksheet bounds")
        worksheet = self.workbook[sheet] if isinstance(sheet, str) else sheet
        if isinstance(worksheet[coordinate].value, ArrayFormula):
            raise ValueError("Cannot replace an array anchor with a scalar reference")
        index = next((item["index"] for item in self.external_link_targets()
                      if source_url in item["targets"]), None)
        if index is None:
            link = ExternalLink(externalBook=ExternalBook(
                id="rId1", sheetNames=ExternalSheetNames(sheetName=[source_sheet]),
            ))
            link.file_link = Relationship(
                Id="rId1", Type=f"{DOCUMENT_REL_NS}/externalLinkPath",
                Target=source_url, TargetMode="External",
            )
            self.workbook._external_links.append(link)
            index = len(self.workbook._external_links)
        else:
            # Remove a prior local alternative when reusing a verified cloud link.
            self.set_external_link_target(index, source_url)
            link = self.workbook._external_links[index - 1]
            if link.externalBook.sheetNames is None:
                link.externalBook.sheetNames = ExternalSheetNames()
            if source_sheet not in link.externalBook.sheetNames.sheetName:
                link.externalBook.sheetNames.sheetName.append(source_sheet)
                part = f"xl/externalLinks/externalLink{index}.xml"
                if part in self._preserved_package_parts:
                    root = fromstring(self._preserved_package_parts[part])
                    book = root.find(f"{{{MAIN_NS}}}externalBook")
                    names = book.find(f"{{{MAIN_NS}}}sheetNames")
                    if names is None:
                        names = SubElement(book, f"{{{MAIN_NS}}}sheetNames")
                        book.remove(names)
                        book.insert(0, names)
                    SubElement(names, f"{{{MAIN_NS}}}sheetName", {"val": source_sheet})
                    self._preserved_package_parts[part] = tostring(root)
        escaped_sheet = source_sheet.replace("'", "''")
        worksheet[coordinate] = f"='[{index}]{escaped_sheet}'!{source_cell}"
        self.invalidate_formula_cache(worksheet, coordinate)
        return index

    def set_dynamic_formula(
        self,
        sheet: str | Any,
        coordinate: str,
        formula: str,
    ) -> None:
        worksheet = self.workbook[sheet] if isinstance(sheet, str) else sheet
        existing = worksheet[coordinate].value
        if isinstance(existing, ArrayFormula) and existing.ref:
            min_col, min_row, max_col, max_row = range_boundaries(existing.ref)
            for row in range(min_row, max_row + 1):
                for column in range(min_col, max_col + 1):
                    if (row, column) != (
                        worksheet[coordinate].row,
                        worksheet[coordinate].column,
                    ):
                        worksheet.cell(row=row, column=column).value = None
        normalised, _ = normalize_excel365_formula(formula)
        worksheet[coordinate] = ArrayFormula(ref=coordinate, text=normalised)
        self._dynamic_by_worksheet_id.setdefault(id(worksheet), set()).add(coordinate)
        self.invalidate_formula_cache(worksheet, coordinate)

    def invalidate_formula_cache(
        self,
        sheet: str | Any,
        coordinate: str,
    ) -> None:
        """Prevent a changed formula from inheriting its source cached value."""

        worksheet = self.workbook[sheet] if isinstance(sheet, str) else sheet
        self._invalidated_formula_caches_by_worksheet_id.setdefault(
            id(worksheet), set()
        ).add(coordinate)

    def _prepare_formulas(self) -> dict[str, set[str]]:
        dynamic: dict[str, set[str]] = {}

        for worksheet in self.workbook.worksheets:
            explicit = self._dynamic_by_worksheet_id.get(id(worksheet), set())
            source_formula_texts = self._source_formula_texts_by_worksheet_id.get(
                id(worksheet), {}
            )
            # Clear only unedited former spill outputs when the source anchor
            # changes or becomes a scalar. Explicit edits in that area survive.
            for (ws_id, coordinate), children in self._source_array_outputs.items():
                if ws_id != id(worksheet):
                    continue
                value = worksheet[coordinate].value
                formula = value.text if isinstance(value, ArrayFormula) else value
                current = (
                    normalize_excel365_formula(formula)[0]
                    if isinstance(formula, str) and formula.startswith("=")
                    else None
                )
                if current == source_formula_texts.get(coordinate):
                    continue
                for address, cached in children.items():
                    if worksheet[address].value == cached:
                        worksheet[address].value = None
            anchors: set[str] = set()
            spill_outputs_to_clear: set[tuple[int, int]] = set()
            for cell in list(worksheet._cells.values()):
                value = cell.value
                if isinstance(value, ArrayFormula):
                    normalised, functions = normalize_excel365_formula(value.text)
                    unchanged_source_formula = (
                        source_formula_texts.get(cell.coordinate) == normalised
                    )
                    is_dynamic = (
                        cell.coordinate in explicit
                        or (
                            not unchanged_source_formula
                            and bool(functions & _DYNAMIC_FUNCTIONS)
                        )
                    )
                    cell.value = ArrayFormula(ref=value.ref, text=normalised)
                    if is_dynamic:
                        if value.ref:
                            min_col, min_row, max_col, max_row = range_boundaries(
                                value.ref
                            )
                            for row in range(min_row, max_row + 1):
                                for column in range(min_col, max_col + 1):
                                    if (row, column) != (cell.row, cell.column):
                                        spill_outputs_to_clear.add((row, column))
                        cell.value.ref = cell.coordinate
                        anchors.add(cell.coordinate)
                elif isinstance(value, str) and value.startswith("="):
                    normalised, functions = normalize_excel365_formula(value)
                    unchanged_source_formula = (
                        source_formula_texts.get(cell.coordinate) == normalised
                    )
                    if cell.coordinate in explicit or (
                        not unchanged_source_formula
                        and functions & _DYNAMIC_FUNCTIONS
                    ):
                        cell.value = ArrayFormula(
                            ref=cell.coordinate,
                            text=normalised,
                        )
                        anchors.add(cell.coordinate)
                    elif normalised != value:
                        cell.value = normalised

            # Excel writes calculated spill children as ordinary cached cells.
            # They must be blank again before Excel can recreate the spill.
            for row, column in spill_outputs_to_clear:
                worksheet.cell(row=row, column=column).value = None

            if anchors:
                dynamic[worksheet.title] = anchors
                self._dynamic_by_worksheet_id[id(worksheet)] = set(anchors)

        return dynamic

    def save(
        self,
        filename: str | os.PathLike[str],
        *,
        recalculate_on_open: bool = False,
    ) -> Path:
        """Atomically save with metadata restored and no calculation chain.

        Formula evaluation is left to Excel. Set ``recalculate_on_open=True``
        to explicitly request full calculation when the workbook is opened.
        """

        destination = Path(filename).expanduser().resolve()
        if destination.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Output must use the .xlsx or .xlsm extension")
        destination.parent.mkdir(parents=True, exist_ok=True)

        # Include rich-text-only sheets in package patching, even without
        # formulas or metadata, so new workbooks receive the same fix.
        for worksheet in self.workbook.worksheets:
            if any(
                isinstance(cell.value, CellRichText)
                for cell in worksheet._cells.values()
            ):
                self._preserved_cell_attributes.setdefault(worksheet.title, {})
        dynamic_cells = self._prepare_formulas()
        invalidated_formula_caches = {
            worksheet.title: set(
                self._invalidated_formula_caches_by_worksheet_id.get(
                    id(worksheet), set()
                )
            )
            for worksheet in self.workbook.worksheets
            if self._invalidated_formula_caches_by_worksheet_id.get(id(worksheet))
        }
        metadata_xml = self._source_metadata_xml
        dynamic_metadata_indices = set(self._source_dynamic_indices)
        if dynamic_cells and metadata_xml is None:
            metadata_xml = _METADATA_XML
            dynamic_metadata_indices = {"1"}
        if dynamic_cells and not dynamic_metadata_indices:
            metadata_names = ", ".join(sorted(self._source_metadata_types)) or "none"
            raise UnsupportedWorkbookMetadataError(
                "Workbook metadata cannot accept new dynamic arrays because it "
                "has no XLDAPR cellMetadata record. Existing metadata types: "
                f"{metadata_names}."
            )
        dynamic_metadata_index = min(
            dynamic_metadata_indices or {"1"}, key=int
        )
        # Leave Excel defaults implicit: automatic mode, no forced full calc.
        calculation = self.workbook.calculation
        if calculation is None:
            calculation = self.workbook.calculation = CalcProperties()
        calculation.calcMode = None
        calculation.fullCalcOnLoad = True if recalculate_on_open else None
        calculation.forceFullCalc = None

        # Stage outside the destination directory: OneDrive registers any
        # short-lived sibling temp file and then retries it forever once it
        # is renamed away (ghost "itemNotFound" downloads).
        staging_dir = _staging_dir()
        stage_fd, stage_name = tempfile.mkstemp(
            dir=staging_dir,
            prefix=f"{destination.stem}.openpyxl-",
            suffix=destination.suffix,
        )
        output_fd, output_name = tempfile.mkstemp(
            dir=staging_dir,
            prefix=f"{destination.stem}.excel365-",
            suffix=destination.suffix,
        )
        os.close(stage_fd)
        os.close(output_fd)
        stage_path = Path(stage_name)
        output_path = Path(output_name)

        try:
            self.workbook.save(stage_path)
            if (
                metadata_xml is not None
                or self._preserved_package_parts
                or self._preserved_cell_attributes
                or self._preserved_formula_attributes
                or self._extended_data_validations
                or self._threaded_comment_state.get("worksheets")
            ):
                _rewrite_package(
                    stage_path,
                    output_path,
                    dynamic_cells,
                    metadata_xml=metadata_xml,
                    dynamic_metadata_index=dynamic_metadata_index,
                    preserved_cell_attributes=self._preserved_cell_attributes,
                    preserved_formula_attributes=self._preserved_formula_attributes,
                    preserved_package_parts=self._preserved_package_parts,
                    preserved_content_types=self._preserved_content_types,
                    preserved_workbook_relationships=(
                        self._preserved_workbook_relationships
                    ),
                    worksheet_drawings=self._worksheet_drawings,
                    extended_data_validations=(
                        self._extended_data_validations
                    ),
                    threaded_comment_state=self._threaded_comment_state,
                    invalidated_formula_caches=invalidated_formula_caches,
                )
            else:
                os.replace(stage_path, output_path)
            _publish(output_path, destination)
        finally:
            stage_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)

        self.source_path = destination
        return destination


def _staging_dir() -> Path:
    """Local (non-synced) directory for save-time temp files.

    Override with ``EXCEL365_STAGING_DIR``; defaults to ``~/Library/Caches``
    on macOS and the system temp directory elsewhere.
    """

    override = os.environ.get("EXCEL365_STAGING_DIR")
    if override:
        base = Path(override).expanduser()
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Caches" / "excel365"
    else:
        base = Path(tempfile.gettempdir()) / "excel365"
    base.mkdir(parents=True, exist_ok=True)
    return base


def _publish(source: Path, destination: Path) -> None:
    """Move the finished workbook into place without leaving a sibling temp file.

    Same-volume moves are a single atomic rename.  Across volumes, copy to a
    ``.tmp`` sibling and rename that.
    """

    try:
        os.replace(source, destination)
        return
    except OSError as error:
        if error.errno != errno.EXDEV:
            raise
    sibling_fd, sibling_name = tempfile.mkstemp(
        dir=destination.parent, prefix=f"~{destination.stem}.", suffix=".tmp"
    )
    os.close(sibling_fd)
    sibling = Path(sibling_name)
    try:
        shutil.copyfile(source, sibling)
        os.replace(sibling, destination)
    finally:
        sibling.unlink(missing_ok=True)


def inspect_excel365_dynamic_arrays(
    filename: str | os.PathLike[str],
) -> dict[str, Any]:
    """Inspect the OOXML pieces required for Excel 365 spill behaviour."""

    path = Path(filename).expanduser().resolve()
    formulas: list[dict[str, str | None]] = []
    errors: list[str] = []

    with ZipFile(path) as archive:
        names = set(archive.namelist())
        has_metadata = "xl/metadata.xml" in names
        metadata_xml = archive.read("xl/metadata.xml") if has_metadata else None
        dynamic_metadata_indices = _dynamic_metadata_indices(metadata_xml)
        metadata_types = sorted(_source_metadata_types(archive))

        rels = fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        has_relationship = any(
            relationship.get("Type") == SHEET_METADATA_REL
            for relationship in rels
        )

        content_types = fromstring(archive.read("[Content_Types].xml"))
        has_content_type = any(
            item.get("PartName") == "/xl/metadata.xml"
            and item.get("ContentType") == SHEET_METADATA_CONTENT_TYPE
            for item in content_types.findall(f"{{{CONTENT_TYPES_NS}}}Override")
        )

        for sheet_name, part in _sheet_part_map(archive).items():
            root = fromstring(archive.read(part))
            for cell in root.findall(f".//{{{MAIN_NS}}}c"):
                formula = cell.find(f"{{{MAIN_NS}}}f")
                if formula is None or not formula.text:
                    continue
                if cell.get("cm") not in dynamic_metadata_indices:
                    continue
                record = {
                    "sheet": sheet_name,
                    "cell": cell.get("r"),
                    "formula": formula.text,
                    "cm": cell.get("cm"),
                    "type": formula.get("t"),
                    "ref": formula.get("ref"),
                }
                formulas.append(record)
                if record["type"] != "array":
                    errors.append(f"{sheet_name}!{record['cell']} is not t=array")
                reference_anchor = (record["ref"] or "").split(":", 1)[0]
                if reference_anchor != record["cell"]:
                    errors.append(f"{sheet_name}!{record['cell']} has wrong ref")
                if "#" in (record["formula"] or ""):
                    errors.append(f"{sheet_name}!{record['cell']} stores raw # syntax")

    if formulas and not has_metadata:
        errors.append("xl/metadata.xml is missing")
    if formulas and not has_relationship:
        errors.append("sheetMetadata relationship is missing")
    if formulas and not has_content_type:
        errors.append("sheetMetadata content type is missing")

    return {
        "path": str(path),
        "valid": not errors,
        "has_metadata": has_metadata,
        "has_relationship": has_relationship,
        "has_content_type": has_content_type,
        "metadata_types": metadata_types,
        "dynamic_metadata_indices": sorted(dynamic_metadata_indices, key=int),
        "formulas": formulas,
        "errors": errors,
    }


__all__ = [
    "Excel365Workbook",
    "UnsupportedWorkbookMetadataError",
    "inspect_excel365_dynamic_arrays",
    "normalize_excel365_formula",
]
