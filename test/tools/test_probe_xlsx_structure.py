import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "tools"))

from probe_xlsx_structure import inspect_workbook  # noqa: E402


WORKBOOK_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <sheets><sheet name="SensitiveSheetName" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''
RELS_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Target="xl/worksheets/sheet1.xml" Type="worksheet"/>
</Relationships>'''
SHARED_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <si><t>Product Secret Name</t></si><si><t>amg001_study</t></si>
 <si><t>sample</t></si><si><t>AMG002</t></si><si><t>Beyond E</t></si>
</sst>'''
SHEET_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <dimension ref="A1:Z1048576"/><sheetData>
  <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"/><c r="C1" t="s"><v>1</v></c>
   <c r="D1" t="s"><v>3</v></c><c r="F1" t="s"><v>4</v></c></row>
  <row r="2"><c r="A2" t="s"><v>2</v></c><c r="D2" t="s"><v>2</v></c></row>
  <row r="4"><c r="C4" t="s"><v>2</v></c><c r="D4" t="s"><v>2</v></c></row>
  <row r="20"><c r="A20" t="s"><v>2</v></c><c r="D20" t="s"><v>2</v></c></row>
  <row r="25"><c r="D25"><f>1+1</f><v>2</v></c></row>
 </sheetData></worksheet>'''


class ProbeTests(unittest.TestCase):
    def test_summary_is_anonymized_and_finds_anchorless_values(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "sample.xlsx"
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("xl/workbook.xml", WORKBOOK_XML)
                archive.writestr("xl/_rels/workbook.xml.rels", RELS_XML)
                archive.writestr("xl/sharedStrings.xml", SHARED_XML)
                archive.writestr("xl/worksheets/sheet1.xml", SHEET_XML)

            report = inspect_workbook(path)
            sheet = report["sheets"][0]

        self.assertEqual(report["worksheet_count"], 1)
        self.assertEqual(sheet["amg_columns"], [3, 4])
        self.assertEqual(sheet["product_columns_a_to_e"], [1])
        self.assertEqual(sheet["product_columns_after_e"], 1)
        self.assertEqual(sheet["last_anchor_row"], 20)
        self.assertEqual(sheet["last_amg_row"], 25)
        self.assertEqual(sheet["amg_rows_without_any_anchor"], 2)
        self.assertEqual(sheet["internal_fully_empty_rows"], 20)
        self.assertEqual(sheet["formula_cells_in_amg"], 1)
        self.assertEqual(sheet["xml_dimension_last_row"], 1048576)
        self.assertNotIn("SensitiveSheetName", repr(report))
        self.assertNotIn("Product Secret Name", repr(report))


if __name__ == "__main__":
    unittest.main()
