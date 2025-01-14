import os
from typing import Dict, Any, List
import xml.etree.ElementTree as ET
import base64
from logger import Logger
import zipfile
from openpyxl.utils import get_column_letter
from openai_helper import OpenAIHelper
from vml_processor import VMLProcessor


class DrawingExtractor:

    def __init__(self, logger: Logger):
        self.logger = logger
        self.openai_helper = OpenAIHelper()
        self.ns = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'xdr':
            'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'r':
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            'sp': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'pr':
            'http://schemas.openxmlformats.org/package/2006/relationships',
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
                        target = target[1:] if target.startswith(
                            '/xl/'
                        ) else f'xl/{target}' if not target.startswith(
                            'xl/') else target

                        sheet_base = os.path.splitext(target)[0]
                        sheet_rels_path = f"{sheet_base}.xml.rels"
                        sheet_rels_filename = f'xl/worksheets/_rels/{os.path.basename(sheet_rels_path)}'

                        if sheet_rels_filename in excel_zip.namelist():
                            with excel_zip.open(
                                    sheet_rels_filename) as sheet_rels:
                                sheet_rels_tree = ET.parse(sheet_rels)
                                sheet_rels_root = sheet_rels_tree.getroot()

                                for sheet_rel in sheet_rels_root.findall(
                                        './/pr:Relationship', self.ns):
                                    rel_target = sheet_rel.get('Target', '')
                                    if 'drawing' in rel_target.lower():
                                        drawing_path = rel_target.replace(
                                            '..', 'xl')
                                        sheet_drawing_map[
                                            sheet_name] = drawing_path

        except Exception as e:
            self.logger.error(
                f"Error in get_sheet_drawing_relations: {str(e)}")

        self.logger.method_end("get_sheet_drawing_relations")
        return sheet_drawing_map

    def _extract_shape_info(self, sp, anchor, vml_controls):
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
                    range_str = self._get_range_from_coordinates(
                        shape_info["coordinates"])
                    shape_info["range"] = range_str

                    matching_control = None
                    for control in vml_controls:
                        if control.get('numeric_id') == shape_id:
                            matching_control = control
                            break

                    if matching_control:
                        shape_info.update({
                            "text_content":
                            matching_control.get("text", ""),
                            "form_control_type":
                            matching_control.get("type"),
                            "form_control_state":
                            matching_control.get("checked", False),
                        })
                        if matching_control.get("is_first_button") is not None:
                            shape_info["is_first_button"] = matching_control[
                                "is_first_button"]
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
                    "from": {
                        "col": from_col,
                        "row": from_row
                    },
                    "to": {
                        "col": to_col,
                        "row": to_row
                    }
                }
        else:
            from_elem = anchor.find('.//xdr:from', self.ns)
            to_elem = anchor.find('.//xdr:to', self.ns) or anchor.find(
                './/xdr:ext', self.ns)

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
                    "from": {
                        "col": from_col,
                        "row": from_row
                    },
                    "to": {
                        "col": to_col,
                        "row": to_row
                    }
                }

        return coords

    def _get_range_from_coordinates(self, coords):
        from_col = get_column_letter(coords["from"]["col"] + 1)
        to_col = get_column_letter(coords["to"]["col"] + 1)
        return f"{from_col}{coords['from']['row'] + 1}:{to_col}{coords['to']['row'] + 1}"

    def _get_vml_controls(self, excel_zip):
        vml_controls = []
        vml_files = [
            f for f in excel_zip.namelist()
            if f.startswith('xl/drawings/') and f.endswith('.vml')
        ]

        for vml_file in vml_files:
            try:
                with excel_zip.open(vml_file) as f:
                    vml_content = f.read().decode('utf-8')
                    controls = self._parse_vml_for_controls(vml_content)
                    vml_controls.extend(controls)
            except Exception as e:
                self.logger.error(
                    f"Error processing VML file {vml_file}: {str(e)}")
                self.logger.exception(e)

        return vml_controls

    def _parse_vml_for_controls(self, vml_content):
        vml_processor = VMLProcessor(self.logger)
        return vml_processor.parse_vml_for_controls(vml_content)

    def _process_shapes(self, anchor, vml_controls, drawing_list):
        for sp in anchor.findall('.//xdr:sp', self.ns):
            shape_info = self._extract_shape_info(sp, anchor, vml_controls)
            if shape_info:
                drawing_list.append(shape_info)

    def _process_drawings(self, anchor, excel_zip, drawing_list, openai_helper,
                          drawing_path):
        coordinates = self._get_coordinates(anchor)
        range_str = self._get_range_from_coordinates(coordinates)

        # Process images
        for pic in anchor.findall('.//xdr:pic', self.ns):
            image_info = self.extract_picture_info(pic, excel_zip, self.ns,
                                                   drawing_path)
            if image_info:
                image_info["coordinates"] = coordinates
                image_info["range"] = range_str
                drawing_list.append(image_info)

        # Process charts using ChartProcessor
        from chart_processor import ChartProcessor
        chart = anchor.find('.//c:chart', self.ns)
        if chart is not None:
            chart_processor = ChartProcessor(self.logger)
            chart_info = chart_processor._extract_chart_info(chart, excel_zip)
            if chart_info:
                chart_info["coordinates"] = coordinates
                chart_info["range"] = range_str
                drawing_list.append(chart_info)

    def extract_drawing_info(self, sheet, excel_zip, drawing_path,
                             openai_helper) -> List[Dict[str, Any]]:
        self.logger.method_start("extract_drawing_info")
        drawing_list = []
        try:
            vml_controls = self._get_vml_controls(excel_zip)

            with excel_zip.open(drawing_path) as xml_file:
                tree = ET.parse(xml_file)
                root = tree.getroot()

                # SmartArt要素の詳細検索とログ出力
                self.logger.debug(
                    f"Starting SmartArt detection in file: {drawing_path}")
                self.logger.debug(f"XML Root tag: {root.tag}")

                ns = {
                    'mc':
                    'http://schemas.openxmlformats.org/markup-compatibility/2006',
                    'dgm':
                    'http://schemas.openxmlformats.org/drawingml/2006/diagram',
                    'a':
                    'http://schemas.openxmlformats.org/drawingml/2006/main'
                }

                # XML構造の全体をログ出力
                self.logger.debug("Full XML structure:")

                def print_elem(elem, level=0):
                    self.logger.debug(f"{'  ' * level}Element: {elem.tag}")
                    self.logger.debug(
                        f"{'  ' * level}Attributes: {elem.attrib}")
                    if elem.text and elem.text.strip():
                        self.logger.debug(
                            f"{'  ' * level}Text: {elem.text.strip()}")
                    for child in elem:
                        print_elem(child, level + 1)

                print_elem(root)

                # 複数のパターンで検索
                smartart_patterns = [
                    './/mc:AlternateContent', './/dgm:relIds',
                    './/dgm:dataModel',
                    './/a:graphicData[@uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"]'
                ]

                for pattern in smartart_patterns:
                    elements = root.findall(pattern, ns)
                    self.logger.debug(
                        f"Searching pattern '{pattern}' found {len(elements)} elements"
                    )
                    for elem in elements:
                        self.logger.debug(f"Element tag: {elem.tag}")
                        self.logger.debug(f"Element attributes: {elem.attrib}")
                        # 子要素も確認
                        for child in elem.iter():
                            self.logger.debug(
                                f"Child element: {child.tag} - {child.attrib}")

                anchors = (root.findall('.//xdr:twoCellAnchor', self.ns) +
                           root.findall('.//xdr:oneCellAnchor', self.ns) +
                           root.findall('.//xdr:absoluteAnchor', self.ns))

                for anchor in anchors:
                    self._process_shapes(anchor, vml_controls, drawing_list)
                    self._process_drawings(anchor, excel_zip, drawing_list,
                                           openai_helper, drawing_path)

                    # SmartArtの検出と処理
                    smartart_elem = anchor.find(
                        './/a:graphicData[@uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"]',
                        self.ns)
                    if smartart_elem is not None:
                        smartart_info = self._extract_smartart_info(
                            smartart_elem, excel_zip, drawing_path)
                        if smartart_info:
                            # 座標情報を設定
                            smartart_info["coordinates"] = self._get_coordinates(anchor)
                            smartart_info["range"] = self._get_range_from_coordinates(
                                smartart_info["coordinates"])
                            
                            # テキストコンテンツを文字列として結合
                            if "nodes" in smartart_info:
                                all_texts = []
                                for node in smartart_info["nodes"]:
                                    if "text_list" in node and node["text_list"]:
                                        all_texts.extend(node["text_list"])
                                smartart_info["text_content"] = " ".join(all_texts)
                            
                            drawing_list.append(smartart_info)

        except Exception as e:
            self.logger.error(f"Error in extract_drawing_info: {str(e)}")

        return drawing_list

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

    def extract_picture_info(self, pic, excel_zip, ns, drawing_path):
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
                            # シート固有のdrawing番号を使用
                            drawing_number = os.path.basename(
                                drawing_path).replace('drawing',
                                                      '').replace('.xml', '')
                            rels_path = f'xl/drawings/_rels/drawing{drawing_number}.xml.rels'
                            if rels_path in excel_zip.namelist():
                                with excel_zip.open(rels_path) as rels_file:
                                    rels_tree = ET.parse(rels_file)
                                    rels_root = rels_tree.getroot()

                                    for rel in rels_root.findall(
                                            './/{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'
                                    ):
                                        if rel.get('Id') == image_ref:
                                            image_path = rel.get(
                                                'Target').replace('..', 'xl')
                                            if image_path in excel_zip.namelist(
                                            ):
                                                with excel_zip.open(
                                                        image_path
                                                ) as img_file:
                                                    image_data = img_file.read(
                                                    )
                                                    image_base64 = base64.b64encode(
                                                        image_data).decode(
                                                            'utf-8')

                                                    analysis_result = None
                                                    if hasattr(
                                                            self,
                                                            'openai_helper'):
                                                        analysis_result = self.openai_helper.analyze_image_with_gpt4o(
                                                            image_base64)
                                                    if analysis_result:
                                                        image_info[
                                                            "gpt4o_analysis"] = analysis_result

                        except Exception as e:
                            self.logger.error(
                                f"Error analyzing image: {str(e)}")

                        return image_info
            return None
        except Exception as e:
            self.logger.error(f"Error in extract_picture_info: {str(e)}")
            return None

    def _extract_smartart_info(self, smartart_elem, excel_zip, drawing_path):
        try:
            self.logger.debug("Extracting SmartArt info")

            smartart_info = {
                "type": "smartart",
                "name": "",
                "description": "",
                "diagram_type": "",
                "layout_type": "",
                "text_contents": [],
                "style": {},
                "nodes": []
            }

            # SmartArtのリレーションシップIDを探す
            rel_ids = smartart_elem.find('.//dgm:relIds', {
                'dgm':
                'http://schemas.openxmlformats.org/drawingml/2006/diagram'
            })
            if rel_ids is not None:
                print(f"Found relIds: {rel_ids.attrib}")
                data_model_rel = rel_ids.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}dm'
                )
                style_rel = rel_ids.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}quickStyle'
                )
                color_rel = rel_ids.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}color'
                )

                # データモデルの解析
                if data_model_rel:
                    print(
                        f"Extracting diagram data for rel_id: {data_model_rel}"
                    )
                    diagram_data = self._extract_diagram_data(
                        excel_zip, data_model_rel, drawing_path)
                    if diagram_data:
                        print(f"Diagram data extracted: {diagram_data}")
                        smartart_info.update(diagram_data)

                # スタイル情報の解析
                if style_rel:
                    print(f"Extracting style data for rel_id: {style_rel}")
                    style_data = self._extract_style_data(excel_zip, style_rel)
                    if style_data:
                        print(f"Style data extracted: {style_data}")
                        smartart_info["style"] = style_data

            # レイアウトタイプの取得
            layout_elem = smartart_elem.find('.//dgm:layoutDef', {
                'dgm':
                'http://schemas.openxmlformats.org/drawingml/2006/diagram'
            })
            if layout_elem is not None:
                print(
                    f"Found layout element with uniqueId: {layout_elem.get('uniqueId', '')}"
                )
                smartart_info["layout_type"] = layout_elem.get('uniqueId', '')

            # テキスト内容の取得
            for text_elem in smartart_elem.findall('.//dgm:t', {
                    'dgm':
                    'http://schemas.openxmlformats.org/drawingml/2006/diagram'
            }):
                if text_elem is not None and text_elem.text:
                    print(f"Extracting text: {text_elem.text}")
                    smartart_info["text_contents"].append({
                        "text":
                        text_elem.text,
                        "level":
                        self._get_text_level(text_elem)
                    })

            # ノード構造の解析
            for pt_elem in smartart_elem.findall('.//dgm:pt', {
                    'dgm':
                    'http://schemas.openxmlformats.org/drawingml/2006/diagram'
            }):
                print("Extracting node info from element")
                node_info = self._extract_node_info(pt_elem)
                if node_info:
                    print(f"Node info extracted: {node_info}")
                    smartart_info["nodes"].append(node_info)

            return smartart_info
        except Exception as e:
            self.logger.error(f"Error in _extract_smartart_info: {str(e)}")
            self.logger.exception(e)
            return None

    def _extract_diagram_data(self, excel_zip, rel_id, drawing_path):
        try:
            # まずdiagrams直下のxmlファイルを確認
            diagram_files = [
                f for f in excel_zip.namelist()
                if f.startswith("xl/diagrams/") and f.endswith(".xml")
            ]
            
            diagram_path = None
            # データモデルファイルを探す
            for diag_file in diagram_files:
                if "data" in diag_file.lower() or "dm" in diag_file.lower():
                    diagram_path = diag_file
                    break
            
            # 見つからない場合は従来の方法でも探す
            if not diagram_path:
                drawing_number = os.path.basename(drawing_path).replace('drawing', '').replace('.xml', '')
                rels_path = f'xl/drawings/_rels/drawing{drawing_number}.xml.rels'
                
                if rels_path in excel_zip.namelist():
                    with excel_zip.open(rels_path) as rels_file:
                        rels_tree = ET.parse(rels_file)
                        rels_root = rels_tree.getroot()
                        
                        for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            if rel.get('Id') == rel_id:
                                diagram_path = 'xl/' + rel.get('Target').replace('..', '')
                                break

            if not diagram_path or diagram_path not in excel_zip.namelist():
                self.logger.debug("SmartArt(ダイアグラム)に相当するファイルが見つかりませんでした。")
                return None

            with excel_zip.open(diagram_path) as f:
                tree = ET.parse(f)
                root = tree.getroot()

                ns = {
                    'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
                    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
                }

                nodes = root.findall('.//dgm:pt', ns)
                diagram_data = {
                    "diagram_type": root.get('type', ''),
                    "name": root.get('name', ''),
                    "description": root.get('description', ''),
                    "diagram_file": diagram_path,
                    "nodes": []
                }

                for node in nodes:
                    node_id = node.get('modelId')
                    
                    # すべての a:t 要素を検索してテキストを抽出
                    all_a_t_elems = node.findall('.//a:t', ns)
                    texts = [el.text for el in all_a_t_elems if el.text]
                    
                    diagram_data['nodes'].append({
                        'id': node_id,
                        'text_list': texts,
                    })

                return diagram_data
            
        except Exception as e:
            self.logger.error(f"Error extracting diagram data: {str(e)}")
            return None

    def _extract_style_data(self, excel_zip, rel_id):
        try:
            style_path = f'xl/diagrams/quickStyle{rel_id}.xml'
            if style_path in excel_zip.namelist():
                with excel_zip.open(style_path) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()

                    return {
                        "style_id": root.get('id', ''),
                        "category": root.get('cat', ''),
                        "color_scheme": root.get('colorStyle', '')
                    }
            return None
        except Exception as e:
            self.logger.error(f"Error extracting style data: {str(e)}")
            return None

    def _get_text_level(self, text_elem):
        try:
            parent = text_elem.getparent()
            while parent is not None:
                if parent.tag.endswith('lvl'):
                    return int(parent.get('val', '0'))
                parent = parent.getparent()
            return 0
        except Exception:
            return 0

    def _extract_node_info(self, pt_elem):
        try:
            return {
                "node_id":
                pt_elem.get('modelId', ''),
                "node_type":
                pt_elem.get('type', ''),
                "text":
                pt_elem.findtext(
                    './/dgm:t', '', {
                        'dgm':
                        'http://schemas.openxmlformats.org/drawingml/2006/diagram'
                    })
            }
        except Exception:
            return None
        except Exception as e:
            self.logger.error(f"Error in _extract_smartart_info: {str(e)}")
            return None
