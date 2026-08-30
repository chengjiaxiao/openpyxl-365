from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.xml.functions import SubElement, fromstring, tostring
from openpyxl365.workbook import (
    MAIN_NS,
    Excel365Workbook,
    _sheet_part_map,
    inspect_excel365_dynamic_arrays,
    normalize_excel365_formula,
)


class Excel365FormulaTests(unittest.TestCase):
    def test_formula_normalisation(self) -> None:
        formula, functions = normalize_excel365_formula(
            '="UNIQUE("&VSTACK(UNIQUE(A1:A3),FILTER(B1:B3,B1:B3>0))'
        )
        self.assertIn('"UNIQUE("', formula)
        self.assertIn("_xlfn.VSTACK(", formula)
        self.assertIn("_xlfn.UNIQUE(", formula)
        self.assertIn("_xlfn._xlws.FILTER(", formula)
        self.assertEqual({"FILTER", "UNIQUE", "VSTACK"}, functions)

    def test_spill_operator_normalisation(self) -> None:
        formula, functions = normalize_excel365_formula(
            "=XLOOKUP('My Sheet'!D1#,A1:A5,B1:B5)"
        )
        self.assertEqual(
            "=_xlfn.XLOOKUP("
            "_xlfn.ANCHORARRAY('My Sheet'!D1),A1:A5,B1:B5)",
            formula,
        )
        self.assertEqual({"ANCHORARRAY", "XLOOKUP"}, functions)


class Excel365RoundTripTests(unittest.TestCase):
    def _rewrite_part(
        self,
        source: Path,
        destination: Path,
        part_name: str,
        replacement: bytes,
    ) -> None:
        with ZipFile(source, "r") as archive, ZipFile(
            destination, "w", ZIP_DEFLATED
        ) as output:
            for info in archive.infolist():
                data = replacement if info.filename == part_name else archive.read(
                    info.filename
                )
                output.writestr(info, data)

    def _build_workbook(self, path: Path) -> None:
        document = Excel365Workbook.create()
        sheet = document.wb.active
        sheet.title = "DynamicArrays"
        rows = [
            ("apple", 10),
            ("banana", 20),
            ("apple", 30),
            ("pear", 40),
            ("banana", 50),
        ]
        for row, values in enumerate(rows, 1):
            sheet.cell(row, 1, values[0])
            sheet.cell(row, 2, values[1])

        # Direct openpyxl-style assignments are upgraded automatically.
        sheet["D1"] = "=UNIQUE(A1:A5)"
        sheet["F1"] = "=VSTACK(B1:B3,B1:B3)"
        sheet["H1"] = '=XLOOKUP("pear",A1:A5,B1:B5)'
        sheet["J1"] = "=XLOOKUP(D1#,A1:A5,B1:B5)"
        document.save(path)

    def test_create_and_modify_existing_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            initial = Path(temp_dir) / "initial.xlsx"
            modified = Path(temp_dir) / "modified.xlsx"
            self._build_workbook(initial)

            first_report = inspect_excel365_dynamic_arrays(initial)
            self.assertTrue(first_report["valid"], first_report["errors"])
            self.assertEqual(4, len(first_report["formulas"]))

            document = Excel365Workbook.open(initial)
            document.wb["DynamicArrays"]["C1"] = "modified with openpyxl"
            document.save(modified)

            second_report = inspect_excel365_dynamic_arrays(modified)
            self.assertTrue(second_report["valid"], second_report["errors"])
            self.assertEqual(4, len(second_report["formulas"]))

            workbook = load_workbook(modified, data_only=False)
            sheet = workbook["DynamicArrays"]
            self.assertEqual("modified with openpyxl", sheet["C1"].value)
            self.assertIsInstance(sheet["D1"].value, ArrayFormula)
            self.assertEqual(
                "=_xlfn.UNIQUE(A1:A5)",
                sheet["D1"].value.text,
            )
            self.assertEqual(
                "=_xlfn.XLOOKUP("
                "_xlfn.ANCHORARRAY(D1),A1:A5,B1:B5)",
                sheet["J1"].value.text,
            )

    def test_excel_spill_cache_cells_are_cleared_before_resave(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "resaved.xlsx"
            document = Excel365Workbook.create()
            sheet = document.wb.active
            document.set_dynamic_formula(sheet, "D2", "=UNIQUE(A2:A6)")

            # This is how an Excel-saved dynamic array looks to openpyxl:
            # the anchor ref covers the spill and child results look static.
            sheet["D2"].value.ref = "D2:D4"
            sheet["D3"] = "banana"
            sheet["D4"] = "pear"
            document.save(output)

            workbook = load_workbook(output, data_only=False)
            result = workbook.active
            self.assertIsInstance(result["D2"].value, ArrayFormula)
            self.assertEqual("D2", result["D2"].value.ref)
            self.assertIsNone(result["D3"].value)
            self.assertIsNone(result["D4"].value)

    def test_existing_excel_spill_ref_and_cached_children_are_preserved(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            initial = Path(temp_dir) / "initial.xlsx"
            excel_saved = Path(temp_dir) / "excel-saved.xlsx"
            output = Path(temp_dir) / "roundtrip.xlsx"

            document = Excel365Workbook.create()
            sheet = document.wb.active
            sheet.title = "DynamicArrays"
            for row, value in enumerate(("apple", "banana", "pear"), 1):
                sheet.cell(row, 1, value)
            document.set_dynamic_formula(sheet, "D1", "=UNIQUE(A1:A3)")
            # Excel stores calculated spill children as ordinary cached cells.
            sheet["D2"] = "banana"
            sheet["D3"] = "pear"
            document.save(initial)

            with ZipFile(initial) as archive:
                part = _sheet_part_map(archive)["DynamicArrays"]
                root = fromstring(archive.read(part))
            anchor = root.find(
                f'.//{{{MAIN_NS}}}c[@r="D1"]/{{{MAIN_NS}}}f'
            )
            self.assertIsNotNone(anchor)
            anchor.set("ref", "D1:D3")
            self._rewrite_part(initial, excel_saved, part, tostring(root))

            reopened = Excel365Workbook.open(excel_saved)
            reopened.wb["DynamicArrays"]["C1"] = "unrelated edit"
            reopened.save(output)

            workbook = load_workbook(output, data_only=False)
            result = workbook["DynamicArrays"]
            self.assertIsInstance(result["D1"].value, ArrayFormula)
            self.assertEqual("D1:D3", result["D1"].value.ref)
            self.assertEqual("banana", result["D2"].value)
            self.assertEqual("pear", result["D3"].value)

    def test_replacing_existing_dynamic_formula_clears_old_spill_children(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            initial = Path(temp_dir) / "initial.xlsx"
            excel_saved = Path(temp_dir) / "excel-saved.xlsx"
            output = Path(temp_dir) / "replaced.xlsx"

            document = Excel365Workbook.create()
            sheet = document.wb.active
            sheet.title = "DynamicArrays"
            for row, value in enumerate(("apple", "banana", "pear"), 1):
                sheet.cell(row, 1, value)
            document.set_dynamic_formula(sheet, "D1", "=UNIQUE(A1:A3)")
            sheet["D2"] = "banana"
            sheet["D3"] = "pear"
            document.save(initial)

            with ZipFile(initial) as archive:
                part = _sheet_part_map(archive)["DynamicArrays"]
                root = fromstring(archive.read(part))
            anchor = root.find(
                f'.//{{{MAIN_NS}}}c[@r="D1"]/{{{MAIN_NS}}}f'
            )
            self.assertIsNotNone(anchor)
            anchor.set("ref", "D1:D3")
            self._rewrite_part(initial, excel_saved, part, tostring(root))

            reopened = Excel365Workbook.open(excel_saved)
            reopened.set_dynamic_formula(
                "DynamicArrays",
                "D1",
                "=SORT(UNIQUE(A1:A3))",
            )
            reopened.save(output)

            result = load_workbook(output, data_only=False)["DynamicArrays"]
            self.assertIsInstance(result["D1"].value, ArrayFormula)
            self.assertEqual("D1", result["D1"].value.ref)
            self.assertIsNone(result["D2"].value)
            self.assertIsNone(result["D3"].value)

    def test_shared_formula_follower_is_not_misclassified_as_new_array(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            initial = Path(temp_dir) / "initial.xlsx"
            shared = Path(temp_dir) / "shared.xlsx"
            output = Path(temp_dir) / "roundtrip.xlsx"

            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Shared"
            sheet["A1"] = "apple"
            sheet["A2"] = "pear"
            sheet["B1"] = "=_xlfn.XLOOKUP(A1,$A$1:$A$2,$A$1:$A$2)"
            sheet["B2"] = "=_xlfn.XLOOKUP(A2,$A$1:$A$2,$A$1:$A$2)"
            workbook.save(initial)

            with ZipFile(initial) as archive:
                part = _sheet_part_map(archive)["Shared"]
                root = fromstring(archive.read(part))
            leader = root.find(
                f'.//{{{MAIN_NS}}}c[@r="B1"]/{{{MAIN_NS}}}f'
            )
            follower = root.find(
                f'.//{{{MAIN_NS}}}c[@r="B2"]/{{{MAIN_NS}}}f'
            )
            self.assertIsNotNone(leader)
            self.assertIsNotNone(follower)
            leader.attrib.update({"t": "shared", "si": "0", "ref": "B1:B2"})
            follower.attrib.update({"t": "shared", "si": "0"})
            follower.text = None
            self._rewrite_part(initial, shared, part, tostring(root))

            reopened = Excel365Workbook.open(shared)
            reopened.wb["Shared"]["C1"] = "unrelated edit"
            reopened.save(output)

            result = load_workbook(output, data_only=False)["Shared"]
            self.assertNotIsInstance(result["B1"].value, ArrayFormula)
            self.assertNotIsInstance(result["B2"].value, ArrayFormula)
            self.assertEqual(
                "=_xlfn.XLOOKUP(A2,$A$1:$A$2,$A$1:$A$2)",
                result["B2"].value,
            )

    def test_rich_data_parts_relationships_and_content_types_are_preserved(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            initial = Path(temp_dir) / "initial.xlsx"
            with_rich_data = Path(temp_dir) / "with-rich-data.xlsx"
            output = Path(temp_dir) / "roundtrip.xlsx"

            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "RichData"
            sheet["A1"] = 15
            workbook.save(initial)

            metadata_xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
 <metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes>
 <futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>
 <valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>
</metadata>"""
            rich_parts = {
                "xl/richData/rdrichvalue.xml": b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><rv s="0"><v>13</v><v>3</v></rv></rvData>""",
                "xl/richData/rdrichvaluestructure.xml": b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><s t="_error"><k n="errorType" t="i"/><k n="subType" t="i"/></s></rvStructures>""",
                "xl/richData/rdRichValueTypes.xml": b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"><global/></rvTypesInfo>""",
            }
            relationship_types = {
                "xl/richData/rdrichvalue.xml": "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue",
                "xl/richData/rdrichvaluestructure.xml": "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure",
                "xl/richData/rdRichValueTypes.xml": "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes",
            }
            content_types_by_part = {
                "/xl/metadata.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
                "/xl/richData/rdrichvalue.xml": "application/vnd.ms-excel.rdrichvalue+xml",
                "/xl/richData/rdrichvaluestructure.xml": "application/vnd.ms-excel.rdrichvaluestructure+xml",
                "/xl/richData/rdRichValueTypes.xml": "application/vnd.ms-excel.rdrichvaluetypes+xml",
            }

            with ZipFile(initial) as archive:
                sheet_part = _sheet_part_map(archive)["RichData"]
                worksheet = fromstring(archive.read(sheet_part))
                worksheet.find(
                    f'.//{{{MAIN_NS}}}c[@r="A1"]'
                ).set("vm", "1")

                relationships = fromstring(
                    archive.read("xl/_rels/workbook.xml.rels")
                )
                SubElement(
                    relationships,
                    "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                    {
                        "Id": "rId20",
                        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
                        "Target": "metadata.xml",
                    },
                )
                for index, (part, relationship_type) in enumerate(
                    relationship_types.items(), start=21
                ):
                    SubElement(
                        relationships,
                        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                        {
                            "Id": f"rId{index}",
                            "Type": relationship_type,
                            "Target": part.removeprefix("xl/"),
                        },
                    )

                content_types = fromstring(archive.read("[Content_Types].xml"))
                for part_name, content_type in content_types_by_part.items():
                    SubElement(
                        content_types,
                        "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
                        {"PartName": part_name, "ContentType": content_type},
                    )

                replacements = {
                    sheet_part: tostring(worksheet),
                    "xl/_rels/workbook.xml.rels": tostring(relationships),
                    "[Content_Types].xml": tostring(content_types),
                }
                with ZipFile(with_rich_data, "w", ZIP_DEFLATED) as patched:
                    for info in archive.infolist():
                        patched.writestr(
                            info,
                            replacements.get(info.filename, archive.read(info.filename)),
                        )
                    patched.writestr("xl/metadata.xml", metadata_xml)
                    for part, data in rich_parts.items():
                        patched.writestr(part, data)

            document = Excel365Workbook.open(with_rich_data)
            document.wb["RichData"]["B1"] = "edited"
            document.save(output)

            with ZipFile(output) as archive:
                self.assertEqual(metadata_xml, archive.read("xl/metadata.xml"))
                for part, data in rich_parts.items():
                    self.assertEqual(data, archive.read(part))

                relationships = fromstring(
                    archive.read("xl/_rels/workbook.xml.rels")
                )
                saved_relationships = {
                    item.get("Type"): item.get("Target")
                    for item in relationships
                }
                for part, relationship_type in relationship_types.items():
                    self.assertEqual(
                        part.removeprefix("xl/"),
                        saved_relationships[relationship_type],
                    )

                content_types = fromstring(archive.read("[Content_Types].xml"))
                saved_content_types = {
                    item.get("PartName"): item.get("ContentType")
                    for item in content_types
                }
                for part_name, content_type in content_types_by_part.items():
                    self.assertEqual(
                        content_type,
                        saved_content_types[part_name],
                    )

                worksheet = fromstring(
                    archive.read(_sheet_part_map(archive)["RichData"])
                )
                self.assertEqual(
                    "1",
                    worksheet.find(
                        f'.//{{{MAIN_NS}}}c[@r="A1"]'
                    ).get("vm"),
                )

    def test_macro_assigned_drawing_shape_is_preserved(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            initial = Path(temp_dir) / "initial.xlsx"
            with_shape = Path(temp_dir) / "with-shape.xlsx"
            output = Path(temp_dir) / "roundtrip.xlsx"

            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "CHECK"
            sheet["A1"] = "keep me"
            workbook.save(initial)

            drawing_xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <xdr:twoCellAnchor>
  <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:sp macro="[0]!SetFormulaA10" textlink="">
   <xdr:nvSpPr><xdr:cNvPr id="2" name="Macro button"/><xdr:cNvSpPr/></xdr:nvSpPr>
   <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
   <xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Run macro</a:t></a:r></a:p></xdr:txBody>
  </xdr:sp>
  <xdr:clientData/>
 </xdr:twoCellAnchor>
</xdr:wsDr>"""

            with ZipFile(initial) as archive:
                sheet_part = _sheet_part_map(archive)["CHECK"]
                worksheet = fromstring(archive.read(sheet_part))
                worksheet.append(
                    fromstring(
                        b'<drawing xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                        b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>'
                    )
                )
                content_types = fromstring(archive.read("[Content_Types].xml"))
                SubElement(
                    content_types,
                    "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
                    {
                        "PartName": "/xl/drawings/drawing1.xml",
                        "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml",
                    },
                )
                replacements = {
                    sheet_part: tostring(worksheet),
                    "[Content_Types].xml": tostring(content_types),
                }
                with ZipFile(with_shape, "w", ZIP_DEFLATED) as patched:
                    for info in archive.infolist():
                        patched.writestr(
                            info,
                            replacements.get(info.filename, archive.read(info.filename)),
                        )
                    patched.writestr("xl/drawings/drawing1.xml", drawing_xml)
                    patched.writestr(
                        "xl/worksheets/_rels/sheet1.xml.rels",
                        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
                        b'</Relationships>',
                    )

            document = Excel365Workbook.open(with_shape)
            document.wb["CHECK"]["B1"] = "edited"
            document.save(output)

            with ZipFile(output) as archive:
                self.assertEqual(
                    drawing_xml,
                    archive.read("xl/drawings/drawing1.xml"),
                )
                sheet_part = _sheet_part_map(archive)["CHECK"]
                worksheet = fromstring(archive.read(sheet_part))
                drawing = worksheet.find(f"{{{MAIN_NS}}}drawing")
                self.assertIsNotNone(drawing)
                relationship_id = drawing.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
                )
                rels_part = "xl/worksheets/_rels/sheet1.xml.rels"
                relationships = fromstring(archive.read(rels_part))
                relationship = next(
                    item
                    for item in relationships
                    if item.get("Id") == relationship_id
                )
                self.assertTrue(relationship.get("Type", "").endswith("/drawing"))
                self.assertEqual("../drawings/drawing1.xml", relationship.get("Target"))
                shape = fromstring(archive.read("xl/drawings/drawing1.xml")).find(
                    ".//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}sp"
                )
                self.assertEqual("[0]!SetFormulaA10", shape.get("macro"))

                content_types = fromstring(archive.read("[Content_Types].xml"))
                drawing_override = next(
                    item
                    for item in content_types
                    if item.get("PartName") == "/xl/drawings/drawing1.xml"
                )
                self.assertEqual(
                    "application/vnd.openxmlformats-officedocument.drawing+xml",
                    drawing_override.get("ContentType"),
                )


if __name__ == "__main__":
    unittest.main()
