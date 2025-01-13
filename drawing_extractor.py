
import os
from typing import Dict, Any, List
import xml.etree.ElementTree as ET
import base64
from logger import Logger
import zipfile
from openpyxl.utils import get_column_letter

class DrawingExtractor:
    def __init__(self, logger: Logger):
        self.logger = logger
        self.ns = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            'sp': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
            'x': 'urn:schemas-microsoft-com:office:excel',
            'v': 'urn:schemas-microsoft-com:vml'
        }

    def get_sheet_drawing_relations(self, excel_zip) -> Dict[str, str]:
        self.logger.method_start("get_sheet_drawing_relations")
        sheet_drawing_map = {}
        try:
            with excel_zip.open('xl/workbook.xml') as wb_xml:
                wb_tree = ET.parse(wb_xml)
                wb_root = wb_tree.getroot()
                sheets = {
                    sheet.get(f'{{{self.ns["r"]}}}id'): sheet.get('name', '')
                    for sheet in wb_root.findall('.//sp:sheet', self.ns)
                }

            with excel_zip.open('xl/_rels/workbook.xml.rels') as rels_xml:
                rels_tree = ET.parse(rels_xml)
                rels_root = rels_tree.getroot()

                for rel in rels_root.findall('.//pr:Relationship', self.ns):
                    r_id = rel.get('Id')
                    if r_id in sheets:
                        sheet_name = sheets[r_id]
                        target = rel.get('Target')
                        target = target[1:] if target.startswith('/xl/') else f'xl/{target}' if not target.startswith('xl/') else target

                        sheet_base = os.path.splitext(target)[0]
                        sheet_rels_path = f"{sheet_base}.xml.rels"
                        sheet_rels_filename = f'xl/worksheets/_rels/{os.path.basename(sheet_rels_path)}'

                        if sheet_rels_filename in excel_zip.namelist():
                            with excel_zip.open(sheet_rels_filename) as sheet_rels:
                                sheet_rels_tree = ET.parse(sheet_rels)
                                sheet_rels_root = sheet_rels_tree.getroot()

                                for sheet_rel in sheet_rels_root.findall('.//pr:Relationship', self.ns):
                                    rel_target = sheet_rel.get('Target', '')
                                    if 'drawing' in rel_target.lower():
                                        drawing_path = rel_target.replace('..', 'xl')
                                        sheet_drawing_map[sheet_name] = drawing_path

        except Exception as e:
            self.logger.error(f"Error in get_sheet_drawing_relations: {str(e)}")

        self.logger.method_end("get_sheet_drawing_relations")
        return sheet_drawing_map

    def extract_drawing_info(self, sheet, excel_zip, drawing_path, openai_helper) -> List[Dict[str, Any]]:
        self.logger.method_start("extract_drawing_info")
        drawing_list = []
        try:
            vml_controls = self._get_vml_controls(excel_zip)

            with excel_zip.open(drawing_path) as xml_file:
                tree = ET.parse(xml_file)
                root = tree.getroot()

                anchors = (
                    root.findall('.//xdr:twoCellAnchor', self.ns) +
                    root.findall('.//xdr:oneCellAnchor', self.ns) +
                    root.findall('.//xdr:absoluteAnchor', self.ns)
                )

                for anchor in anchors:
                    self._process_shapes(anchor, vml_controls, drawing_list)
                    self._process_drawings(anchor, excel_zip, drawing_list, openai_helper)

        except Exception as e:
            self.logger.error(f"Error in extract_drawing_info: {str(e)}")

        return drawing_list
