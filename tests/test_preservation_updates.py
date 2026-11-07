"""Synthetic regression fixtures: no customer workbooks or cloud accounts."""
import errno
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.xml.functions import SubElement, fromstring, tostring

from openpyxl365 import Excel365Workbook
from openpyxl365.workbook import MAIN_NS, PYTHON_REL, _publish

REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
SHEET = "xl/worksheets/sheet1.xml"


class PreservationUpdatesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "source.xlsx"
        self.output = self.root / "output.xlsx"

    def rewrite(self, replacements):
        with ZipFile(self.source) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        parts.update(replacements)
        with ZipFile(self.source, "w") as archive:
            for name, data in parts.items():
                archive.writestr(name, data)

    def xml(self, path, part=SHEET):
        with ZipFile(path) as archive:
            return fromstring(archive.read(part))

    def cell(self, root, coordinate):
        return root.find(f'.//{{{MAIN_NS}}}c[@r="{coordinate}"]')

    def test_calculation_flags_match_file_and_save_defaults(self):
        wb = Workbook()
        wb.active["A1"] = "=1+1"
        wb.save(self.source)
        doc = Excel365Workbook.open(self.source)
        self.assertTrue(doc.wb.calculation.fullCalcOnLoad)
        doc.save(self.output)
        calc = Excel365Workbook.open(self.output).wb.calculation
        self.assertIsNone(calc.fullCalcOnLoad)
        self.assertIsNone(calc.forceFullCalc)
        self.assertIsNone(calc.calcMode)
        doc.save(self.output, recalculate_on_open=True)
        self.assertTrue(Excel365Workbook.open(self.output).wb.calculation.fullCalcOnLoad)

    def test_missing_calc_properties_can_be_opened(self):
        Workbook().save(self.source)
        root = self.xml(self.source, "xl/workbook.xml")
        root.remove(root.find(f"{{{MAIN_NS}}}calcPr"))
        self.rewrite({"xl/workbook.xml": tostring(root)})
        doc = Excel365Workbook.open(self.source)
        doc.save(self.output)
        self.assertTrue(self.output.exists())

    def test_changed_array_clears_metadata_and_keeps_explicit_edits(self):
        doc = Excel365Workbook.create()
        doc.set_dynamic_formula("Sheet", "A1", "=SEQUENCE(3)")
        doc.wb.active["A2"] = 2
        doc.wb.active["A3"] = 3
        doc.save(self.source)
        root = self.xml(self.source)
        self.cell(root, "A1").find(f"{{{MAIN_NS}}}f").set("ref", "A1:A3")
        self.rewrite({SHEET: tostring(root)})
        doc = Excel365Workbook.open(self.source)
        doc.wb.active["A1"] = 10
        doc.wb.active["A3"] = "edited"
        doc.save(self.output)
        result = load_workbook(self.output).active
        self.assertEqual(result["A1"].value, 10)
        self.assertIsNone(result["A2"].value)
        self.assertEqual(result["A3"].value, "edited")
        self.assertNotIn("cm", self.cell(self.xml(self.output), "A1").attrib)

    def shared_fixture(self):
        wb = Workbook()
        wb.active["A1"] = "=B1+1"
        wb.active["A2"] = "=B2+1"
        wb.save(self.source)
        root = self.xml(self.source)
        master = self.cell(root, "A1").find(f"{{{MAIN_NS}}}f")
        follower = self.cell(root, "A2").find(f"{{{MAIN_NS}}}f")
        master.attrib.update(t="shared", si="0", ref="A1:A2")
        follower.attrib.update(t="shared", si="0")
        follower.text = None
        self.rewrite({SHEET: tostring(root)})

    def test_shared_group_survives_unrelated_edit(self):
        self.shared_fixture()
        doc = Excel365Workbook.open(self.source)
        doc.wb.active["C1"] = 5
        doc.save(self.output)
        root = self.xml(self.output)
        self.assertEqual(self.cell(root, "A2").find(f"{{{MAIN_NS}}}f").get("t"), "shared")
        self.assertEqual(load_workbook(self.output).active["A2"].value, "=B2+1")

    def test_editing_shared_master_or_follower_expands_whole_group(self):
        for coordinate in ("A1", "A2"):
            with self.subTest(coordinate=coordinate):
                self.shared_fixture()
                doc = Excel365Workbook.open(self.source)
                doc.wb.active[coordinate] = "=42"
                doc.save(self.output)
                root = self.xml(self.output)
                self.assertFalse(root.findall(f'.//{{{MAIN_NS}}}f[@t="shared"]'))
                self.assertEqual(load_workbook(self.output).active[coordinate].value, "=42")

    def test_rich_text_spaces_survive_new_and_existing_workbooks(self):
        doc = Excel365Workbook.create()
        doc.wb.active["A1"] = CellRichText(TextBlock(InlineFont(b=True), " hello "), " world ")
        doc.save(self.source)
        doc = Excel365Workbook.open(self.source)
        doc.save(self.output)
        self.assertEqual(str(load_workbook(self.output, rich_text=True).active["A1"].value), " hello  world ")
        for node in self.xml(self.output).findall(f".//{{{MAIN_NS}}}is//{{{MAIN_NS}}}t"):
            self.assertEqual(node.get("{http://www.w3.org/XML/1998/namespace}space"), "preserve")

    def test_python_payload_relationship_and_content_type_survive(self):
        Workbook().save(self.source)
        part = "xl/python/python.xml"
        rels = self.xml(self.source, "xl/_rels/workbook.xml.rels")
        SubElement(rels, f"{{{REL}}}Relationship", Id="rId99", Type=PYTHON_REL, Target="python/python.xml")
        types = self.xml(self.source, "[Content_Types].xml")
        SubElement(types, f"{{{CT}}}Override", PartName="/" + part, ContentType="application/test+xml")
        payload = b"<python>synthetic payload</python>"
        self.rewrite({part: payload, "xl/_rels/workbook.xml.rels": tostring(rels), "[Content_Types].xml": tostring(types)})
        Excel365Workbook.open(self.source).save(self.output)
        with ZipFile(self.output) as archive:
            self.assertEqual(archive.read(part), payload)
        self.assertTrue(any(r.get("Type") == PYTHON_REL for r in self.xml(self.output, "xl/_rels/workbook.xml.rels")))
        self.assertTrue(any(r.get("PartName") == "/" + part for r in self.xml(self.output, "[Content_Types].xml")))

    def test_cloud_reference_reuse_and_relocation(self):
        doc = Excel365Workbook.create()
        url = "https://example.com/source.xlsx"
        for cell, sheet in (("A1", "Data"), ("A2", "Other")):
            self.assertEqual(doc.set_external_reference("Sheet", cell, source_url=url, source_sheet=sheet, source_cell="$B$2"), 1)
        doc.save(self.source)
        doc = Excel365Workbook.open(self.source)
        target = "https://example.com/moved/source.xlsx"
        doc.set_external_link_target(1, target)
        doc.save(self.output)
        reopened = Excel365Workbook.open(self.output)
        self.assertEqual(reopened.external_link_targets(), [{"index": 1, "targets": [target]}])
        self.assertEqual(reopened.wb.active["A2"].value, "='[1]Other'!$B$2")

    def test_shared_string_formula_cache_becomes_text(self):
        wb = Workbook()
        wb.active["A1"] = '="cached text"'
        wb.save(self.source)
        root = self.xml(self.source)
        cell = self.cell(root, "A1")
        cell.set("t", "s")
        cell.find(f"{{{MAIN_NS}}}v").text = "0"
        rels = self.xml(self.source, "xl/_rels/workbook.xml.rels")
        SubElement(rels, f"{{{REL}}}Relationship", Id="rId99",
                   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                   Target="sharedStrings.xml")
        types = self.xml(self.source, "[Content_Types].xml")
        SubElement(types, f"{{{CT}}}Override", PartName="/xl/sharedStrings.xml",
                   ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
        self.rewrite({SHEET: tostring(root), "xl/_rels/workbook.xml.rels": tostring(rels),
                      "[Content_Types].xml": tostring(types),
                      "xl/sharedStrings.xml": f'<sst xmlns="{MAIN_NS}" count="1" uniqueCount="1"><si><t>cached text</t></si></sst>'.encode()})
        Excel365Workbook.open(self.source).save(self.output)
        self.assertEqual(load_workbook(self.output, data_only=True).active["A1"].value, "cached text")
        self.assertEqual(self.cell(self.xml(self.output), "A1").get("t"), "str")

    def test_empty_formula_spill_child_is_a_cached_value(self):
        wb = Workbook()
        wb.active["A1"] = 12
        wb.save(self.source)
        root = self.xml(self.source)
        cell = self.cell(root, "A1")
        formula = SubElement(cell, f"{{{MAIN_NS}}}f", ca="1")
        cell.remove(formula)
        cell.insert(0, formula)
        self.rewrite({SHEET: tostring(root)})
        doc = Excel365Workbook.open(self.source)
        self.assertEqual(doc.wb.active["A1"].value, 12)
        doc.save(self.output)
        self.assertEqual(load_workbook(self.output).active["A1"].value, 12)

    def test_cloud_reference_rejects_invalid_input(self):
        doc = Excel365Workbook.create()
        for url in ("http://example.com/a.xlsx", "https://example.com/a.xlsx?token=secret", "https://user:pass@example.com/a.xlsx"):
            with self.subTest(url=url), self.assertRaises(ValueError):
                doc.set_external_reference("Sheet", "A1", source_url=url, source_sheet="Data", source_cell="A1")
        self.assertEqual(doc.external_link_targets(), [])

    def test_cross_volume_publish_cleans_temporary_file(self):
        self.source.write_bytes(b"complete workbook")
        self.output.write_bytes(b"old workbook")
        from os import replace
        calls = []
        def cross_volume(source, destination):
            calls.append(source)
            if len(calls) == 1:
                raise OSError(errno.EXDEV, "cross-volume")
            return replace(source, destination)
        with patch("openpyxl365.workbook.os.replace", side_effect=cross_volume):
            _publish(self.source, self.output)
        self.assertEqual(self.output.read_bytes(), b"complete workbook")
        self.assertFalse(list(self.root.glob("*.tmp")))

    def test_cross_volume_copy_failure_preserves_destination(self):
        self.source.write_bytes(b"complete workbook")
        self.output.write_bytes(b"old workbook")
        with patch("openpyxl365.workbook.os.replace", side_effect=OSError(errno.EXDEV, "cross-volume")), patch("openpyxl365.workbook.shutil.copyfile", side_effect=OSError("copy failed")):
            with self.assertRaises(OSError):
                _publish(self.source, self.output)
        self.assertEqual(self.output.read_bytes(), b"old workbook")
        self.assertFalse(list(self.root.glob("*.tmp")))
