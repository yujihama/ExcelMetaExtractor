import os
import json
import math
from datetime import datetime
import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from logger import Logger
from drawing_extractor import DrawingExtractor
import openpyxl.cell.cell
from openpyxl.utils import get_column_letter
from typing import Dict, Any, List, Optional, Tuple
from openai_helper import OpenAIHelper
import traceback
from pathlib import Path
import tempfile
import streamlit as st
import re
from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference
import matplotlib.pyplot as plt
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
import base64
import numpy as np

class ExcelMetadataExtractor:
    def __init__(self, file_obj):
        self.file_obj = file_obj
        self.workbook = load_workbook(file_obj, data_only=True)
        self.openai_helper = OpenAIHelper()
        self.MAX_CELLS_PER_ANALYSIS = 100
        self.logger = Logger()

        # Store excel_zip for later use
        temp_dir = tempfile.mkdtemp()
        temp_zip = os.path.join(temp_dir, 'temp.xlsx')
        with open(temp_zip, 'wb') as f:
            self.file_obj.seek(0)
            f.write(self.file_obj.read())
        self.excel_zip = zipfile.ZipFile(temp_zip, 'r')

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

    def __init__(self, file_obj):
        self.file_obj = file_obj
        self.workbook = load_workbook(file_obj, data_only=True)
        self.openai_helper = OpenAIHelper()
        self.MAX_CELLS_PER_ANALYSIS = 100
        self.logger = Logger()
        self.drawing_extractor = DrawingExtractor(self.logger)

        # Store excel_zip for later use
        temp_dir = tempfile.mkdtemp()
        temp_zip = os.path.join(temp_dir, 'temp.xlsx')
        with open(temp_zip, 'wb') as f:
            self.file_obj.seek(0)
            f.write(self.file_obj.read())
        self.excel_zip = zipfile.ZipFile(temp_zip, 'r')

    def get_sheet_drawing_relations(self, excel_zip) -> Dict[str, str]:
        return self.drawing_extractor.get_sheet_drawing_relations(excel_zip)

    def extract_chart_data(self, filepath, output_dir):
        self.logger.method_start("extract_chart_data")
        workbook = load_workbook(filepath, data_only=True)
        chart_data_list = []

        for sheetname in workbook.sheetnames:
            sheet = workbook[sheetname]
            for chart_index, chart in enumerate(sheet._charts):
                title = self._get_chart_title(chart)
                x_axis_title = self._get_axis_title(chart.x_axis) if chart.x_axis else None
                y_axis_title = self._get_axis_title(chart.y_axis) if chart.y_axis else None

                chart_data = {
                    "sheetname": sheetname,
                    "title": title,
                    "type": type(chart).__name__,
                    "data": [],
                    "categories": [],
                    "x_axis_title": x_axis_title,
                    "y_axis_title": y_axis_title,
                    "series_colors": []
                }

                if isinstance(chart, (BarChart, LineChart, PieChart, ScatterChart)):
                    self._extract_series_data(chart, sheet, chart_data)

                chart_data_list.append(chart_data)
        self.logger.method_end("extract_chart_data")
        return chart_data_list

    def _get_chart_title(self, chart):
        if not chart.title:
            return "Untitled"

        if isinstance(chart.title, str):
            return chart.title

        if hasattr(chart.title, 'tx') and chart.title.tx:
            if hasattr(chart.title.tx, 'rich') and chart.title.tx.rich:
                if len(chart.title.tx.rich.p) > 0:
                    p = chart.title.tx.rich.p[0]
                    if hasattr(p, 'r') and len(p.r) > 0 and hasattr(p.r[0], 't'):
                        return p.r[0].t
                    if hasattr(p, 'fld') and len(p.fld) > 0 and hasattr(p.fld[0], 't'):
                        return p.fld[0].t
            elif hasattr(chart.title.tx, 'strRef') and chart.title.tx.strRef:
                return chart.title.tx.strRef.f
        return "Untitled"

    def _get_axis_title(self, axis):
        if not axis or not axis.title:
            return None

        if isinstance(axis.title, str):
            return axis.title

        if hasattr(axis.title, 'tx') and axis.title.tx:
            if hasattr(axis.title.tx, 'rich') and axis.title.tx.rich:
                if len(axis.title.tx.rich.p) > 0 and axis.title.tx.rich.p[0].r:
                    return axis.title.tx.rich.p[0].r.t
            elif hasattr(axis.title.tx, 'strRef') and axis.title.tx.strRef:
                return axis.title.tx.strRef.f
        return None

    def _extract_series_data(self, chart, sheet, chart_data):
        for series in chart.series:
            chart_data["series_colors"].append(None)

            if series.val.numRef:
                values = self._get_cell_range(series.val.numRef.f, sheet)
                data = []
                for row_tuple in sheet.iter_rows(
                    min_col=values.min_col,
                    min_row=values.min_row,
                    max_col=values.max_col,
                    max_row=values.max_row
                ):
                    row_data = []
                    for cell in row_tuple:
                        value = 0 if cell.value == 'X' else float(cell.value) if cell.value is not None else 0
                        row_data.append(value)
                    data.extend(row_data)
                chart_data["data"].append(data)

            if series.cat and (series.cat.numRef or series.cat.strRef):
                ref = series.cat.numRef or series.cat.strRef
                categories = self._get_cell_range(ref.f, sheet)
                category_labels = []
                for row_tuple in sheet.iter_rows(
                    min_col=categories.min_col,
                    min_row=categories.min_row,
                    max_col=categories.max_col,
                    max_row=categories.max_row
                ):
                    category_labels.extend([cell.value for cell in row_tuple])
                chart_data["categories"].append(category_labels)

    def _get_cell_range(self, range_str, sheet):
        cell_range = range_str.split('!')[1]
        start, end = cell_range.replace('$', '').split(':')
        min_col, min_row = coordinate_from_string(start)
        max_col, max_row = coordinate_from_string(end)

        return Reference(
            sheet,
            min_col=column_index_from_string(min_col),
            min_row=int(min_row),
            max_col=column_index_from_string(max_col),
            max_row=int(max_row)
        )

    def recreate_charts(self, chart_data_list, output_dir):
        output_data = []
        for chart_data in chart_data_list:
            chart_info = {"chart_type": chart_data["type"]}

            if chart_data["categories"] and chart_data["data"]:
                categories = chart_data["categories"][0]
                data = chart_data["data"]

                if chart_data["type"] == "BarChart":
                    chart_info.update(self._process_bar_chart_data(categories, data))
                elif chart_data["type"] == "LineChart":
                    chart_info.update(self._process_line_chart_data(categories, data))
                elif chart_data["type"] == "PieChart":
                    chart_info.update(self._process_pie_chart_data(categories, data))
                elif chart_data["type"] == "ScatterChart":
                    chart_info.update(self._process_scatter_chart_data(categories, data))

            output_data.append(chart_info)

        return output_data

    def _process_bar_chart_data(self, categories, data):
        if len(data) > 1:
            return {
                "x": categories,
                "y": data
            }
        return {
            "x": categories,
            "y": data[0]
        }

    def _process_line_chart_data(self, categories, data):
        return {
            "x": categories,
            "y": data
        }

    def _process_pie_chart_data(self, categories, data):
        if len(data[0]) == len(categories):
            return {
                "labels": categories,
                "data": data[0]
            }
        return {}

    def _process_scatter_chart_data(self, categories, data):
        return {
            "x": categories,
            "y": data
        }

    def extract_drawing_info(self, sheet, excel_zip, drawing_path) -> List[Dict[str, Any]]:
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
                    self._process_drawings(anchor, excel_zip, drawing_list)

        except Exception as e:
            self.logger.error(f"Error in extract_drawing_info: {str(e)}")

        return drawing_list

    def _get_vml_controls(self, excel_zip):
        vml_controls = []
        vml_files = [f for f in excel_zip.namelist() if f.startswith('xl/drawings/') and f.endswith('.vml')]

        for vml_file in vml_files:
            try:
                with excel_zip.open(vml_file) as f:
                    vml_content = f.read().decode('utf-8')
                    controls = self._parse_vml_for_controls(vml_content)
                    vml_controls.extend(controls) #extend the list instead of overwriting.
            except Exception as e:
                self.logger.error(f"Error processing VML file {vml_file}: {str(e)}")
                self.logger.exception(e)

        return vml_controls

    def _process_shapes(self, anchor, vml_controls, drawing_list):
        for sp in anchor.findall('.//xdr:sp', self.ns):
            shape_info = self._extract_shape_info(sp, anchor, vml_controls)
            if shape_info:
                drawing_list.append(shape_info)

    def _process_drawings(self, anchor, excel_zip, drawing_list):
        coordinates = self._get_coordinates(anchor)
        range_str = self._get_range_from_coordinates(coordinates)

        # Process images
        for pic in anchor.findall('.//xdr:pic', self.ns):
            image_info = self._extract_picture_info(pic, excel_zip) #Pass excel_zip
            if image_info:
                image_info["coordinates"] = coordinates
                image_info["range"] = range_str
                drawing_list.append(image_info)

        # Process charts
        chart = anchor.find('.//c:chart', self.ns)
        if chart is not None:
            chart_info = self._extract_chart_info(chart, excel_zip)
            if chart_info:
                chart_info["coordinates"] = coordinates
                chart_info["range"] = range_str
                drawing_list.append(chart_info)

        # Process other elements
        for grp in anchor.findall('.//xdr:grpSp', self.ns):
            group_info = self._extract_group_info(grp)
            if group_info:
                group_info["coordinates"] = coordinates
                group_info["range"] = range_str
                drawing_list.append(group_info)

        for cxn in anchor.findall('.//xdr:cxnSp', self.ns):
            connector_info = self._extract_connector_info(cxn)
            if connector_info:
                connector_info["coordinates"] = coordinates
                connector_info["range"] = range_str
                drawing_list.append(connector_info)



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

            # 図形の基本情報を取得
            name_elem = sp.find('.//xdr:nvSpPr/xdr:cNvPr', self.ns)
            if name_elem is not None:
                shape_info["name"] = name_elem.get('name', '')
                shape_info["hidden"] = name_elem.get('hidden', '0') == '1'
                shape_info["description"] = name_elem.get('descr', '')
                shape_id = name_elem.get('id')

                if shape_id:
                    # 座標情報をセル範囲として保存
                    range_str = self._get_range_from_coordinates(shape_info["coordinates"])
                    shape_info["range"] = range_str

                    # IDに基づいてVMLコントロールを検索
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
                        # テキスト内容を直接取得
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

    def _extract_chart_info(self, chart, excel_zip):
        try:
            chart_ref = chart.get(f'{{{self.ns["r"]}}}id')
            output_dir = os.path.join(tempfile.gettempdir(), 'chart_images')
            os.makedirs(output_dir, exist_ok=True)

            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
                self.file_obj.seek(0)
                temp_file.write(self.file_obj.read())
                chart_data_list = self.extract_chart_data(temp_file.name, output_dir)
                chart_data_json = self.recreate_charts(chart_data_list, output_dir)
                os.unlink(temp_file.name)

            chart_info = {
                "type": "chart",
                "chart_ref": chart_ref,
                "chart_data_json": chart_data_json
            }

            ref_num = re.search(r'rId(\d+)', chart_ref)
            if ref_num:
                chart_xml_path = f'xl/charts/chart{ref_num.group(1)}.xml'
                if chart_xml_path in excel_zip.namelist():
                    with excel_zip.open(chart_xml_path) as chart_xml_file:
                        chart_tree = ET.parse(chart_xml_file)
                        chart_root = chart_tree.getroot()

                        chart_info.update(self._extract_chart_metadata(chart_root))

            return chart_info
        except Exception as e:
            self.logger.error(f"Error in _extract_chart_info: {str(e)}")
            return None

    def _extract_chart_metadata(self, chart_root):
        metadata = {}

        # チャートタイプの取得
        for elem in chart_root.findall('.//c:plotArea/*', self.ns):
            for chart_type in ['barChart', 'pieChart', 'lineChart']:
                if elem.tag.endswith(chart_type):
                    metadata["chartType"] = chart_type.replace('Chart', '')
                    break
            if "chartType" in metadata:
                break

        # タイトルの取得
        title = chart_root.find('.//c:title//c:tx//c:rich//a:t', self.ns)
        if title is not None and title.text:
            metadata["title"] = title.text

        # データ系列の取得
        series_list = []
        for series in chart_root.findall('.//c:ser', self.ns):
            series_info = {}

            series_name = series.find('.//c:tx//c:f', self.ns)
            if series_name is not None:
                series_info["name"] = series_name.text

            values = series.find('.//c:val//c:numRef//c:f', self.ns)
            if values is not None:
                series_info["data_range"] = values.text

            if series_info:
                series_list.append(series_info)

        if series_list:
            metadata["series"] = series_list

        return metadata

    def _extract_picture_info(self, pic, excel_zip): #Modified
        try:
            name_elem = pic.find('.//xdr:nvPicPr/xdr:cNvPr', self.ns)
            if name_elem is not None:
                image_info = {
                    "type": "image",
                    "name": name_elem.get('name', ''),
                    "description": name_elem.get('descr', ''),
                }

                blip = pic.find('.//a:blip', self.ns)
                if blip is not None:
                    image_ref = blip.get(f'{{{self.ns["r"]}}}embed')
                    if image_ref:
                        image_info["image_ref"] = image_ref

                        # Get image data and analyze with GPT-4 Vision
                        try:
                            # Get the relationship file for the current worksheet
                            rels_path = f'xl/drawings/_rels/drawing1.xml.rels'  # Assuming drawing1.xml
                            if rels_path in excel_zip.namelist():
                                with excel_zip.open(rels_path) as rels_file:
                                    rels_tree = ET.parse(rels_file)
                                    rels_root = rels_tree.getroot()

                                    # Find the target path for this image
                                    for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                                        if rel.get('Id') == image_ref:
                                            image_path = rel.get('Target').replace('..', 'xl')
                                            if image_path in excel_zip.namelist():
                                                with excel_zip.open(image_path) as img_file:
                                                    image_data = img_file.read()
                                                    image_base64 = base64.b64encode(image_data).decode('utf-8')

                                                    # Analyze image using GPT-4 Vision
                                                    analysis_result = self.openai_helper.analyze_image_with_gpt4o(image_base64)
                                                    if analysis_result:
                                                        image_info["gpt4o_analysis"] = analysis_result

                        except Exception as e:
                            self.logger.error(f"Error analyzing image: {str(e)}")
                            self.logger.exception(e)

                        return image_info
            return None
        except Exception as e:
            self.logger.error(f"Error in _extract_picture_info: {str(e)}")
            return None

    def _extract_group_info(self, grp):
        try:
            name_elem = grp.find('.//xdr:nvGrpSpPr/xdr:cNvPr', self.ns)
            if name_elem is not None:
                return {
                    "type": "group",
                    "name": name_elem.get('name', ''),
                    "description": name_elem.get('descr', '')
                }
            return None
        except Exception as e:
            self.logger.error(f"Error in _extract_group_info: {str(e)}")
            return None

    def _extract_connector_info(self, cxn):
        try:
            name_elem = cxn.find('.//xdr:nvCxnSpPr/xdr:cNvPr', self.ns)
            if name_elem is not None:
                return {
                    "type": "connector",
                    "name": name_elem.get('name', ''),
                    "description": name_elem.get('descr', '')
                }
            return None
        except Exception as e:
            self.logger.error(f"Error in _extract_connector_info: {str(e)}")
            return None

    def _parse_vml_for_controls(self, vml_content):
        """VMLコンテンツからコントロール情報を抽出"""
        controls = []
        try:
            namespaces = {
                'v': 'urn:schemas-microsoft-com:vml',
                'o': 'urn:schemas-microsoft-com:office:office',
                'x': 'urn:schemas-microsoft-com:office:excel'
            }

            root = ET.fromstring(vml_content)
            control_elements = root.findall('.//{urn:schemas-microsoft-com:vml}shape')


            for element in control_elements:
                try:
                    # テキスト内容を取得
                    textbox = element.find('.//v:textbox', namespaces)
                    text_content = ""
                    if textbox is not None:
                        div = textbox.find('.//div')
                        if div is not None:
                            text_content = "".join(div.itertext()).strip()

                    control_type = element.find('.//{urn:schemas-microsoft-com:office:excel}ClientData')
                    if control_type is not None:
                        control_type_value = control_type.get('ObjectType')

                        shape_id = element.get('id', '')
                        try:
                            # VML IDから数値部分を抽出（例：_x0000_s1027から1027を取得）
                            numeric_id = shape_id.split('_s')[-1]
                            numeric_id = int(numeric_id) if numeric_id.isdigit() else None

                        except (ValueError, IndexError) as e:
                            self.logger.error(f"Error extracting numeric ID from shape_id {shape_id}: {str(e)}")
                            continue

                        control = {
                            'id': shape_id,
                            'numeric_id': str(numeric_id) if numeric_id is not None else None,
                            'type': 'checkbox' if control_type_value == 'Checkbox' else 'radio',
                            'checked': False,
                            'position': '',
                            'text': text_content  # テキスト内容を設定
                        }

                        # チェックボックスの状態
                        checked = control_type.find('.//{urn:schemas-microsoft-com:office:excel}Checked')
                        if checked is not None and checked.text:
                            control['checked'] = checked.text == '1'

                        # アンカー情報の解析（セルの位置）
                        anchor = control_type.find('.//{urn:schemas-microsoft-com:office:excel}Anchor')
                        if anchor is not None and anchor.text:
                            try:
                                coords = [int(x) for x in anchor.text.split(',')]
                                from_col = coords[0]
                                from_row = coords[1]
                                to_col = coords[2]
                                to_row = coords[3]
                                control['position'] = f"{get_column_letter(from_col + 1)}{from_row + 1}:{get_column_letter(to_col + 1)}{to_row + 1}"
                            except (ValueError, IndexError) as e:
                                self.logger.error(f"Error processing anchor coordinates: {str(e)}")

                        # ラジオボタンの追加情報
                        if control_type_value == 'Radio':
                            first_button = control_type.find('.//{urn:schemas-microsoft-com:office:excel}FirstButton')
                            if first_button is not None:
                                control['is_first_button'] = first_button.text == '1'

                        controls.append(control)

                except Exception as control_error:
                    self.logger.error(f"Error processing individual control: {str(control_error)}")
                    continue

        except Exception as e:
            self.logger.error(f"Error parsing VML content: {str(e)}")
            self.logger.exception(e)

        return controls

    def detect_regions(self, sheet) -> List[Dict[str, Any]]:
        self.logger.method_start("detect_regions")
        regions = []
        drawing_regions = []
        cell_regions = []
        processed_cells = set()

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_zip = os.path.join(temp_dir, 'temp.xlsx')
                with open(temp_zip, 'wb') as f:
                    self.file_obj.seek(0)
                    f.write(self.file_obj.read())

                with zipfile.ZipFile(temp_zip, 'r') as excel_zip:
                    sheet_drawing_map = self.get_sheet_drawing_relations(excel_zip)
                    sheet_name = sheet.title
                    if sheet_name in sheet_drawing_map:
                        drawing_path = sheet_drawing_map[sheet_name]
                        drawings = self.extract_drawing_info(sheet, excel_zip, drawing_path)

                        for drawing in drawings:
                            self.logger.start_region_processing(drawing)
                            drawing_type = drawing["type"]

                            region_info = {
                                "regionType": drawing_type,
                                "type": drawing_type,
                                "range": drawing.get("range", ""),
                                "name": drawing.get("name", ""),
                                "description": drawing.get("description", ""),
                                "coordinates": drawing.get("coordinates", {}),
                                "text_content": drawing.get("text_content", ""),
                                "chartType": drawing.get("chartType", ""),
                                "series": drawing.get("series", ""),
                                "chart_data_json": drawing.get("chart_data_json", "")
                            }

                            if drawing_type == "image":
                                if "image_ref" in drawing:
                                    region_info["image_ref"] = drawing["image_ref"]
                                if "gpt4o_analysis" in drawing:
                                    region_info["gpt4o_analysis"] = drawing["gpt4o_analysis"]
                                else:
                                    self.logger.info("No GPT-4 Vision analysis found for image")

                            elif drawing_type == "smartart" and "diagram_type" in drawing:
                                region_info["diagram_type"] = drawing["diagram_type"]

                            if "form_control_type" in drawing:
                                region_info["form_control_type"] = drawing["form_control_type"]
                                region_info["form_control_state"] = drawing.get("form_control_state", False)
                                if "is_first_button" in drawing:
                                    region_info["is_first_button"] = drawing["is_first_button"]

                            drawing_regions.append(region_info)
                            self.logger.end_region_processing(region_info)

                            if "coordinates" in drawing:
                                from_col = drawing["coordinates"]["from"]["col"]
                                from_row = drawing["coordinates"]["from"]["row"]
                                to_col = drawing["coordinates"]["to"]["col"]
                                to_row = drawing["coordinates"]["to"]["row"]

                                for r in range(from_row, to_row + 1):
                                    for c in range(from_col, to_col + 1):
                                        processed_cells.add(f"{get_column_letter(c+1)}{r+1}")

            # セル領域の処理

            for row in range(1, min(sheet.max_row + 1, 100)):
                for col in range(1, min(sheet.max_column + 1, 20)):
                    try:
                        cell_coord = f"{get_column_letter(col)}{row}"
                        if cell_coord in processed_cells:
                            continue

                        cell = sheet.cell(row=row, column=col)
                        if cell.value is None:
                            continue

                        # 区切り文字のみのセルはスキップ
                        if isinstance(cell.value, str) and len(cell.value.strip()) == 1 and cell.value.strip() in '-_=':
                            continue

                        max_row, max_col = self.find_region_boundaries(sheet, row, col)
                        if max_row == row and max_col == col:  # 単一セルの場合はスキップ
                            continue

                        cells_data = self.extract_region_cells(sheet, row, col, max_row, max_col)
                        if not cells_data:  # 空のデータの場合はスキップ
                            continue

                        merged_cells = self.get_merged_cells_info(sheet, row, col, max_row, max_col)

                        # 処理済みのセルを記録
                        for r in range(row, max_row + 1):
                            for c in range(col, max_col + 1):
                                processed_cells.add(f"{get_column_letter(c)}{r}")

                        try:
                            region_analysis = self.openai_helper.analyze_region_type(
                                json.dumps({
                                    "cells": cells_data,
                                    "mergedCells": merged_cells
                                }))

                            if isinstance(region_analysis, str):
                                region_analysis = json.loads(region_analysis)

                            region_type = region_analysis.get("regionType", "unknown")
                            region_metadata = {
                                "regionType": region_type,
                                "range": f"{get_column_letter(col)}{row}:{get_column_letter(max_col)}{max_row}",
                                "sampleCells": cells_data,
                                "mergedCells": merged_cells
                            }

                            if region_type == "table":
                                try:
                                    header_analysis = self.openai_helper.analyze_table_structure(
                                        json.dumps(cells_data), json.dumps(merged_cells))

                                    if isinstance(header_analysis, str):
                                        header_analysis = json.loads(header_analysis)

                                    header_rows = header_analysis.get("headerStructure", {}).get("rows", [])
                                    header_range = "N/A"

                                    if header_rows:
                                        min_header_row = min(header_rows)
                                        max_header_row = max(header_rows)
                                        header_range = (f"{min_header_row}" if min_header_row == max_header_row
                                                       else f"{min_header_row}-{max_header_row}")

                                    region_metadata["headerStructure"] = {
                                        "headerType": header_analysis.get("headerStructure", {}).get("type", "none"),
                                        "headerRows": header_rows,
                                        "headerRange": header_range,
                                        "mergedCells": bool(merged_cells),
                                        "start_row": row
                                    }
                                except Exception as e:
                                    self.logger.error(f"Error analyzing table header: {str(e)}")
                                    continue

                            cell_regions.append(region_metadata)

                        except Exception as e:
                            self.logger.error(f"Error analyzing region at {cell_coord}: {str(e)}")
                            continue

                    except Exception as e:
                        self.logger.error(f"Error processing cell at row {row}, col {col}: {str(e)}")
                        continue

            # サマリーの生成

            for region in drawing_regions + cell_regions:
                try:
                    if "regionType" not in region:
                        region["regionType"] = region.get("type", "unknown")
                    region["summary"] = self.openai_helper.summarize_region(region)
                except Exception as e:
                    self.logger.error(f"Error generating summary for region: {str(e)}")
                    continue

            regions.extend(drawing_regions)
            regions.extend(cell_regions)

            if regions:
                try:
                    metadata = {
                        "type": "metadata",
                        "regionType": "metadata",
                        "totalRegions": len(regions),
                        "drawingRegions": len(drawing_regions),
                        "cellRegions": len(cell_regions)
                    }

                    sheet_data = {
                        "sheetName": sheet.title,
                        "regions": regions,
                        "drawingRegionsCount": len(drawing_regions),
                        "cellRegionsCount": len(cell_regions)
                    }

                    metadata["summary"] = self.openai_helper.generate_sheet_summary(sheet_data)
                    regions.append(metadata)
                except Exception as e:
                    self.logger.error(f"Error generating metadata: {str(e)}")

            return regions

        except Exception as e:
            self.logger.error(f"Error in detect_regions: {str(e)}")
            self.logger.exception(e)
            self.logger.method_end("detect_regions")
            return []
        finally:
            self.logger.method_end("detect_regions")

    def get_file_metadata(self) -> Dict[str, Any]:
        try:
            properties = self.workbook.properties

            return {
                "fileName": self.file_obj.name,
                "fileProperties": {
                    "createdTime": properties.created.isoformat() if properties.created else None,
                    "modifiedTime": properties.modified.isoformat() if properties.modified else None,
                    "fileSize": self.file_obj.size,
                    "author": properties.creator,
                    "lastModifiedBy": properties.lastModifiedBy,
                    "isPasswordProtected": False
                }
            }
        except Exception as e:
            self.logger.error(f"Error in get_file_metadata: {str(e)}")
            self.logger.exception(e)
            raise

    def analyze_cell_type(self, cell) -> str:
        if cell.value is None:
            return "empty"
        if isinstance(cell.value, (int, float)):
            return "numeric"
        if isinstance(cell.value, datetime):
            return "date"
        return "text"

    def find_region_boundaries(self, sheet, start_row: int, start_col: int) -> Tuple[int, int]:
        max_row = start_row
        max_col = start_col
        min_empty_rows = 1
        min_empty_cols = 1

        empty_row_count = 0
        for row in range(start_row, min(sheet.max_row + 1, start_row + 1000)):
            row_empty = True
            for col in range(start_col, min(start_col + 20, sheet.max_column + 1)):
                if sheet.cell(row=row, column=col).value is not None:
                    row_empty = False
                    break

            if row_empty:
                empty_row_count += 1
                if empty_row_count >= min_empty_rows:
                    break
            else:
                empty_row_count = 0
                max_row = row

        empty_col_count = 0
        for col in range(start_col, min(sheet.max_column + 1, start_col + 50)):
            col_empty = True
            for row in range(start_row, min(max_row + 1, start_row + 50)):
                if sheet.cell(row=row, column=col).value is not None:
                    col_empty = False
                    break

            if col_empty:
                empty_col_count += 1
                if empty_col_count >= min_empty_cols:
                    break
            else:
                empty_col_count = 0
                max_col = col

        max_row = max(max_row, start_row)
        max_col = max(max_col, start_col)

        return max_row, max_col

    def get_merged_cells_info(self, sheet, start_row: int, start_col: int, max_row: int, max_col: int) -> List[Dict[str, Any]]:
        merged_cells_info = []
        for merged_range in sheet.merged_cells.ranges:
            if (merged_range.min_row >= start_row and merged_range.max_row <= max_row and merged_range.min_col >= start_col and merged_range.max_col <= max_col):
                merged_cells_info.append({
                    "range": str(merged_range),
                    "value": sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
                })
        return merged_cells_info

    def extract_region_cells(self, sheet, start_row: int, start_col: int, max_row: int, max_col: int) -> List[List[Dict[str, Any]]]:
        cells_data = []
        actual_max_row = min(max_row, start_row + 5)
        actual_max_col = max_col

        for row in range(start_row, actual_max_row + 1):
            row_data = []
            for col in range(start_col, actual_max_col + 1):
                cell = sheet.cell(row=row, column=col)
                cell_type = self.analyze_cell_type(cell)

                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    for merged_range in sheet.merged_cells.ranges:
                        if merged_range.min_row <= row <= merged_range.max_row and \
                           merged_range.min_col <= col <= merged_range.max_col:
                            master_cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                            cell_info = {
                                "row": row,
                                "col": col,
                                "value": str(master_cell.value) if master_cell.value is not None else "",
                                "type": cell_type,
                                "isMerged": True,
                                "mergedRange": str(merged_range)
                            }
                            break
                    else:
                        cell_info = {
                            "row": row,
                            "col": col,
                            "value": "",
                            "type": cell_type,
                            "isMerged": True
                        }
                else:
                    cell_info = {
                        "row": row,
                        "col": col,
                        "value": str(cell.value) if cell.value is not None else "",
                        "type": cell_type
                    }

                row_data.append(cell_info)
            cells_data.append(row_data)

        if max_row > actual_max_row or max_col > actual_max_col:
            self.logger.info(f"Note: Region was truncated from {max_row}x{max_col} to {actual_max_row}x{actual_max_col}")

        return cells_data



    def get_sheet_metadata(self) -> list:
        try:
            sheets_metadata = []

            for sheet_name in self.workbook.sheetnames:
                sheet = self.workbook[sheet_name]

                merged_cells = [str(cell_range) for cell_range in sheet.merged_cells.ranges]
                regions = self.detect_regions(sheet)

                sheet_meta = {
                    "sheetName": sheet_name,
                    "isProtected": sheet.protection.sheet,
                    "rowCount": sheet.max_row,
                    "columnCount": sheet.max_column,
                    "hasPivotTables": bool(getattr(sheet, '_pivots', [])),
                    "hasCharts": bool(getattr(sheet, '_charts', [])),
                    "mergedCells": merged_cells,
                    "regions": regions
                }

                sheets_metadata.append(sheet_meta)

            return sheets_metadata
        except Exception as e:
            self.logger.error(f"Error in get_sheet_metadata: {str(e)}")
            self.logger.exception(e)
            raise

    def extract_all_metadata(self) -> Dict[str, Any]:
        self.logger.method_start("extract_all_metadata")
        try:
            file_metadata = self.get_file_metadata()
            sheets_metadata = self.get_sheet_metadata()

            for sheet in sheets_metadata:
                if "regions" in sheet:
                    for region in sheet["regions"]:
                        if "sampleCells" in region:
                            del region["sampleCells"]

            metadata = {
                **file_metadata,
                "worksheets": sheets_metadata,
                "crossSheetRelationships": []
            }

            return metadata
        except Exception as e:
            self.logger.error(f"Error in extract_all_metadata: {str(e)}")
            self.logger.exception(e)
            self.logger.method_end("extract_all_metadata")
            raise
        finally:
            self.logger.method_end("extract_all_metadata")