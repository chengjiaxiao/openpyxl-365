"""Preserve Excel 365 x14 data-validation extensions across openpyxl saves."""

from __future__ import annotations

from zipfile import ZipFile

from openpyxl.xml.functions import SubElement, fromstring, tostring

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
DATA_VALIDATION_EXTENSION_URI = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"


def discover_extended_data_validations(
    archive: ZipFile,
    sheet_parts: dict[str, str],
) -> dict[str, list[bytes]]:
    """Capture worksheet extension blocks containing x14 data validations."""

    result: dict[str, list[bytes]] = {}
    for sheet_name, part in sheet_parts.items():
        root = fromstring(archive.read(part))
        extension_list = root.find(f"{{{MAIN_NS}}}extLst")
        if extension_list is None:
            continue
        snapshots = [
            tostring(extension)
            for extension in extension_list
            if extension.find(f".//{{{X14_NS}}}dataValidations") is not None
        ]
        if snapshots:
            result[sheet_name] = snapshots
    return result


def restore_extended_data_validations(
    xml: bytes,
    snapshots: list[bytes],
) -> bytes:
    """Restore captured x14 validation extensions into a saved worksheet."""

    if not snapshots:
        return xml

    root = fromstring(xml)
    extension_list = root.find(f"{{{MAIN_NS}}}extLst")
    if extension_list is None:
        extension_list = SubElement(root, f"{{{MAIN_NS}}}extLst")
    else:
        for extension in list(extension_list):
            if extension.find(f".//{{{X14_NS}}}dataValidations") is not None:
                extension_list.remove(extension)

    for snapshot in snapshots:
        extension_list.append(fromstring(snapshot))
    return tostring(root)


def count_extended_data_validations(
    filename: str,
) -> dict[str, int]:
    """Return x14 data-validation counts per worksheet for verification."""

    from openpyxl365.workbook import _sheet_part_map

    counts: dict[str, int] = {}
    with ZipFile(filename) as archive:
        for sheet_name, part in _sheet_part_map(archive).items():
            root = fromstring(archive.read(part))
            count = len(
                root.findall(f".//{{{X14_NS}}}dataValidation")
            )
            if count:
                counts[sheet_name] = count
    return counts


__all__ = [
    "DATA_VALIDATION_EXTENSION_URI",
    "count_extended_data_validations",
    "discover_extended_data_validations",
    "restore_extended_data_validations",
]
