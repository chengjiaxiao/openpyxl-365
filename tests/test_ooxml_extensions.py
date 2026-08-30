from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook
from openpyxl.xml.functions import SubElement, fromstring, tostring
from openpyxl365.ooxml_data_validations import (
    X14_NS,
    count_extended_data_validations,
)
from openpyxl365.ooxml_threaded_comments import (
    COMMENTS_REL,
    PERSON_REL,
    THREADED_COMMENT_REL,
    VML_DRAWING_REL,
)
from openpyxl365.workbook import (
    MAIN_NS,
    Excel365Workbook,
    _sheet_part_map,
)

DOCUMENT_REL_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"


class Excel365OoxmlExtensionTests(unittest.TestCase):
    def _rewrite_package(
        self,
        source: Path,
        destination: Path,
        replacements: dict[str, bytes],
        additions: dict[str, bytes] | None = None,
    ) -> None:
        with ZipFile(source) as archive, ZipFile(
            destination, "w", ZIP_DEFLATED
        ) as output:
            written: set[str] = set()
            for info in archive.infolist():
                output.writestr(
                    info,
                    replacements.get(info.filename, archive.read(info.filename)),
                )
                written.add(info.filename)
            for name, data in replacements.items():
                if name not in written:
                    output.writestr(name, data)
                    written.add(name)
            for name, data in (additions or {}).items():
                output.writestr(name, data)

    def test_x14_data_validation_extension_is_preserved(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            initial = Path(temp_dir) / "initial.xlsx"
            extended = Path(temp_dir) / "extended.xlsx"
            output = Path(temp_dir) / "roundtrip.xlsx"

            workbook = Workbook()
            workbook.active.title = "Parameters"
            workbook.active["A1"] = "Weekly"
            workbook.save(initial)

            with ZipFile(initial) as archive:
                sheet_part = _sheet_part_map(archive)["Parameters"]
                worksheet = fromstring(archive.read(sheet_part))
            extension_list = SubElement(
                worksheet, f"{{{MAIN_NS}}}extLst"
            )
            extension = SubElement(
                extension_list,
                f"{{{MAIN_NS}}}ext",
                {"uri": "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"},
            )
            validations = SubElement(
                extension,
                f"{{{X14_NS}}}dataValidations",
                {"count": "1"},
            )
            validation = SubElement(
                validations,
                f"{{{X14_NS}}}dataValidation",
                {"type": "list", "allowBlank": "1"},
            )
            formula = SubElement(validation, f"{{{X14_NS}}}formula1")
            SubElement(formula, f"{{{XM_NS}}}f").text = '"Daily,Weekly"'
            SubElement(validation, f"{{{XM_NS}}}sqref").text = "A1"
            self._rewrite_package(
                initial,
                extended,
                {sheet_part: tostring(worksheet)},
            )

            document = Excel365Workbook.open(extended)
            document.wb["Parameters"]["B1"] = "edited"
            document.save(output)

            self.assertEqual(
                {"Parameters": 1}, count_extended_data_validations(str(output))
            )
            with ZipFile(output) as archive:
                worksheet = fromstring(
                    archive.read(_sheet_part_map(archive)["Parameters"])
                )
                validation = worksheet.find(
                    f".//{{{X14_NS}}}dataValidation"
                )
                self.assertIsNotNone(validation)
                self.assertEqual(
                    '"Daily,Weekly"',
                    validation.findtext(f".//{{{XM_NS}}}f"),
                )
                self.assertEqual(
                    "A1", validation.findtext(f".//{{{XM_NS}}}sqref")
                )

    def test_threaded_comment_package_is_preserved_without_legacy_duplicates(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            initial = Path(temp_dir) / "initial.xlsx"
            threaded = Path(temp_dir) / "threaded.xlsx"
            output = Path(temp_dir) / "roundtrip.xlsx"

            workbook = Workbook()
            workbook.active.title = "Comments"
            workbook.active["A1"] = "value"
            workbook.save(initial)

            person_id = "{28A6D840-D9D8-4AAF-AB09-6A899745E014}"
            thread_id = "{356614D4-DB32-4713-B3C3-C6B37F2FA382}"
            comment_parts = {
                "xl/comments1.xml": (
                    f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="{MAIN_NS}"><authors><author>tc={thread_id}</author></authors><commentList><comment ref="A1" authorId="0"><text><t>[线程批注]\n\n注释:\n    keep threaded comment</t></text></comment></commentList></comments>'''
                ).encode(),
                "xl/drawings/vmlDrawing1.vml": b'''<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>''',
                "xl/threadedComments/threadedComment1.xml": (
                    f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A1" personId="{person_id}" id="{thread_id}"><text>keep threaded comment</text></threadedComment></ThreadedComments>'''
                ).encode(),
                "xl/persons/person.xml": (
                    f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person displayName="Tester" id="{person_id}" userId="tester" providerId="None"/></personList>'''
                ).encode(),
            }

            with ZipFile(initial) as archive:
                sheet_part = _sheet_part_map(archive)["Comments"]
                worksheet = fromstring(archive.read(sheet_part))
                SubElement(
                    worksheet,
                    f"{{{MAIN_NS}}}legacyDrawing",
                    {f"{{{DOCUMENT_REL_NS}}}id": "rId21"},
                )

                sheet_relationships = fromstring(
                    f'<Relationships xmlns="{PACKAGE_REL_NS}"/>'.encode()
                )
                for relationship_id, relationship_type, target in (
                    ("rId20", COMMENTS_REL, "../comments1.xml"),
                    ("rId21", VML_DRAWING_REL, "../drawings/vmlDrawing1.vml"),
                    (
                        "rId22",
                        THREADED_COMMENT_REL,
                        "../threadedComments/threadedComment1.xml",
                    ),
                ):
                    SubElement(
                        sheet_relationships,
                        f"{{{PACKAGE_REL_NS}}}Relationship",
                        {
                            "Id": relationship_id,
                            "Type": relationship_type,
                            "Target": target,
                        },
                    )

                workbook_relationships = fromstring(
                    archive.read("xl/_rels/workbook.xml.rels")
                )
                SubElement(
                    workbook_relationships,
                    f"{{{PACKAGE_REL_NS}}}Relationship",
                    {
                        "Id": "rId23",
                        "Type": PERSON_REL,
                        "Target": "persons/person.xml",
                    },
                )

                content_types = fromstring(
                    archive.read("[Content_Types].xml")
                )
                SubElement(
                    content_types,
                    f"{{{CONTENT_TYPES_NS}}}Default",
                    {
                        "Extension": "vml",
                        "ContentType": (
                            "application/vnd.openxmlformats-officedocument."
                            "vmlDrawing"
                        ),
                    },
                )
                for part_name, content_type in (
                    (
                        "/xl/comments1.xml",
                        (
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.comments+xml"
                        ),
                    ),
                    (
                        "/xl/threadedComments/threadedComment1.xml",
                        "application/vnd.ms-excel.threadedcomments+xml",
                    ),
                    (
                        "/xl/persons/person.xml",
                        "application/vnd.ms-excel.person+xml",
                    ),
                ):
                    SubElement(
                        content_types,
                        f"{{{CONTENT_TYPES_NS}}}Override",
                        {"PartName": part_name, "ContentType": content_type},
                    )

                replacements = {
                    sheet_part: tostring(worksheet),
                    "xl/worksheets/_rels/sheet1.xml.rels": tostring(
                        sheet_relationships
                    ),
                    "xl/_rels/workbook.xml.rels": tostring(
                        workbook_relationships
                    ),
                    "[Content_Types].xml": tostring(content_types),
                }
            self._rewrite_package(
                initial,
                threaded,
                replacements,
                comment_parts,
            )

            document = Excel365Workbook.open(threaded)
            document.wb["Comments"]["B1"] = "edited"
            document.save(output)

            with ZipFile(output) as archive:
                names = set(archive.namelist())
                for part, data in comment_parts.items():
                    self.assertEqual(data, archive.read(part))
                self.assertNotIn("xl/comments/comment1.xml", names)
                self.assertNotIn("xl/drawings/commentsDrawing1.vml", names)

                sheet_part = _sheet_part_map(archive)["Comments"]
                relationships = fromstring(
                    archive.read("xl/worksheets/_rels/sheet1.xml.rels")
                )
                relationships_by_type = {
                    item.get("Type"): item for item in relationships
                }
                self.assertEqual(
                    "../comments1.xml",
                    relationships_by_type[COMMENTS_REL].get("Target"),
                )
                self.assertEqual(
                    "../threadedComments/threadedComment1.xml",
                    relationships_by_type[THREADED_COMMENT_REL].get("Target"),
                )
                vml_relationship = relationships_by_type[VML_DRAWING_REL]
                worksheet = fromstring(archive.read(sheet_part))
                self.assertEqual(
                    vml_relationship.get("Id"),
                    worksheet.find(f"{{{MAIN_NS}}}legacyDrawing").get(
                        f"{{{DOCUMENT_REL_NS}}}id"
                    ),
                )

                workbook_relationships = fromstring(
                    archive.read("xl/_rels/workbook.xml.rels")
                )
                self.assertTrue(
                    any(
                        item.get("Type") == PERSON_REL
                        and item.get("Target") == "persons/person.xml"
                        for item in workbook_relationships
                    )
                )


if __name__ == "__main__":
    unittest.main()
