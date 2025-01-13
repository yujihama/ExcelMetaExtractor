
import os
from typing import Dict, Any, List
import xml.etree.ElementTree as ET
import base64
from logger import Logger
import zipfile
from openpyxl.utils import get_column_letter
from openai_helper import OpenAIHelper

class DrawingExtractor:
    def __init__(self, logger: Logger):
        self.logger = logger
        self.openai_helper = OpenAIHelper()
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

    def _extract_shape_info(self, sp, anchor, vml_controls):
        """図形情報を抽出し、VMLコントロールとマッチングする"""
        try:
            shape_info = {
                "type": "shape",
                "name": "",
                "description": "",
                "hidden": False,
                "text_content": "",
                "coordinates": self._get_coordinates(anchor),
            }

            name_elem = sp.find('.//xdr:nvSpPr/xdr:cNvPr', self.ns)
            if name_elem is not None:
                shape_info["name"] = name_elem.get('name', '')
                shape_info["hidden"] = name_elem.get('hidden', '0') == '1'
                shape_info["description"] = name_elem.get('descr', '')
                shape_id = name_elem.get('id')

                if shape_id:
                    range_str = self._get_range_from_coordinates(shape_info["coordinates"])
                    shape_info["range"] = range_str

                    matching_control = None
                    for control in vml_controls:
                        if control.get('numeric_id') == shape_id:
                            matching_control = control
                            break

                    if matching_control:
                        shape_info.update({
                            "text_content": matching_control.get("text", ""),
                            "form_control_type": matching_control.get("type"),
                            "form_control_state": matching_control.get("checked", False),
                        })
                        if matching_control.get("is_first_button") is not None:
                            shape_info["is_first_button"] = matching_control["is_first_button"]
                    else:
                        txBody = sp.find('.//xdr:txBody//a:t', self.ns)
                        if txBody is not None and txBody.text:
                            shape_info["text_content"] = txBody.text

            return shape_info
        except Exception as e:
            self.logger.error(f"Error in _extract_shape_info: {str(e)}")
            self.logger.exception(e)
            return None

    def _get_coordinates(self, anchor):
        coords = {"from": {"col": 0, "row": 0}, "to": {"col": 0, "row": 0}}
        
        if anchor.tag.endswith('absoluteAnchor'):
            pos = anchor.find('.//xdr:pos', self.ns)
            ext = anchor.find('.//xdr:ext', self.ns)

            if pos is not None and ext is not None:
                from_col = int(int(pos.get('x', '0')) / 914400)
                from_row = int(int(pos.get('y', '0')) / 914400)
                to_col = from_col + int(int(ext.get('cx', '0')) / 914400)
                to_row = from_row + int(int(ext.get('cy', '0')) / 914400)

                coords = {
                    "from": {"col": from_col, "row": from_row},
                    "to": {"col": to_col, "row": to_row}
                }
        else:
            from_elem = anchor.find('.//xdr:from', self.ns)
            to_elem = anchor.find('.//xdr:to', self.ns) or anchor.find('.//xdr:ext', self.ns)

            if from_elem is not None:
                from_col = int(from_elem.find('xdr:col', self.ns).text)
                from_row = int(from_elem.find('xdr:row', self.ns).text)

                if to_elem is not None:
                    if anchor.tag.endswith('twoCellAnchor'):
                        to_col = int(to_elem.find('xdr:col', self.ns).text)
                        to_row = int(to_elem.find('xdr:row', self.ns).text)
                    else:  # oneCellAnchor
                        cx = int(to_elem.get('cx', '0'))
                        cy = int(to_elem.get('cy', '0'))
                        to_col = from_col + (cx // 914400)
                        to_row = from_row + (cy // 914400)
                else:
                    to_col = from_col + 1
                    to_row = from_row + 1

                coords = {
                    "from": {"col": from_col, "row": from_row},
                    "to": {"col": to_col, "row": to_row}
                }

        return coords

    def _get_range_from_coordinates(self, coords):
        from_col = get_column_letter(coords["from"]["col"] + 1)
        to_col = get_column_letter(coords["to"]["col"] + 1)
        return f"{from_col}{coords['from']['row'] + 1}:{to_col}{coords['to']['row'] + 1}"

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
def extract_picture_info(self, pic, excel_zip, ns): 
        try:
            name_elem = pic.find('.//xdr:nvPicPr/xdr:cNvPr', ns)
            if name_elem is not None:
                image_info = {
                    "type": "image",
                    "name": name_elem.get('name', ''),
                    "description": name_elem.get('descr', ''),
                }

                blip = pic.find('.//a:blip', ns)
                if blip is not None:
                    image_ref = blip.get(f'{{{ns["r"]}}}embed')
                    if image_ref:
                        image_info["image_ref"] = image_ref

                        try:
                            rels_path = f'xl/drawings/_rels/drawing1.xml.rels'
                            if rels_path in excel_zip.namelist():
                                with excel_zip.open(rels_path) as rels_file:
                                    rels_tree = ET.parse(rels_file)
                                    rels_root = rels_tree.getroot()

                                    for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                                        if rel.get('Id') == image_ref:
                                            image_path = rel.get('Target').replace('..', 'xl')
                                            if image_path in excel_zip.namelist():
                                                with excel_zip.open(image_path) as img_file:
                                                    image_data = img_file.read()
                                                    image_base64 = base64.b64encode(image_data).decode('utf-8')

                                                    analysis_result = self.openai_helper.analyze_image_with_gpt4o(image_base64)
                                                    if analysis_result:
                                                        image_info["gpt4o_analysis"] = analysis_result

                        except Exception as e:
                            self.logger.error(f"Error analyzing image: {str(e)}")
                            self.logger.exception(e)

                        return image_info
            return None
        except Exception as e:
            self.logger.error(f"Error in extract_picture_info: {str(e)}")
            return None
