import os
import json
import math
from datetime import datetime
import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
import openpyxl.cell.cell
from openpyxl.utils import get_column_letter
from typing import Dict, Any, List, Optional, Tuple
from openai_helper import OpenAIHelper
import traceback
from pathlib import Path
import tempfile
import streamlit as st


class ExcelMetadataExtractor:

    def __init__(self, file_obj):
        self.file_obj = file_obj
        self.workbook = load_workbook(file_obj, data_only=True)
        self.openai_helper = OpenAIHelper()
        self.MAX_CELLS_PER_ANALYSIS = 100
        self.ns = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'xdr':
            'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'r':
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            'sp': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'pr':
            'http://schemas.openxmlformats.org/package/2006/relationships'
        }

    def get_sheet_drawing_relations(self, excel_zip) -> Dict[str, str]:
        """Get relationships between sheets and drawings from workbook.xml.rels"""
        sheet_drawing_map = {}

        try:
            # xl/workbook.xmlからシート情報を取得
            with excel_zip.open('xl/workbook.xml') as wb_xml:
                wb_tree = ET.parse(wb_xml)
                wb_root = wb_tree.getroot()

                # シート名の対応を取得
                sheets = {}
                for sheet in wb_root.findall('.//sp:sheet', self.ns):
                    r_id = sheet.get(f'{{{self.ns["r"]}}}id')
                    sheet_name = sheet.get('name', '')
                    sheets[r_id] = sheet_name
                    print(f"Found sheet: {sheet_name} (rId: {r_id})")

            print("\nProcessing workbook relationships...")
            # xl/_rels/workbook.xml.relsから関係性を解析
            with excel_zip.open('xl/_rels/workbook.xml.rels') as rels_xml:
                rels_tree = ET.parse(rels_xml)
                rels_root = rels_tree.getroot()

                # シートとターゲットの対応を取得
                for rel in rels_root.findall('.//pr:Relationship', self.ns):
                    r_id = rel.get('Id')
                    if r_id in sheets:
                        sheet_name = sheets[r_id]
                        target = rel.get('Target')
                        print(
                            f"\nProcessing relationship - Sheet: {sheet_name} (rId: {r_id})"
                        )
                        print(f"Original target path: {target}")

                        # パスの正規化
                        if target.startswith('/xl/'):
                            target = target[1:]
                        elif not target.startswith('xl/'):
                            target = f'xl/{target}'
                        print(f"Normalized target path: {target}")

                        # シートごとの_rels/sheet*.xml.relsを確認
                        sheet_base = os.path.splitext(target)[0]
                        sheet_rels_path = f"{sheet_base}.xml.rels"
                        sheet_rels_filename = f'xl/worksheets/_rels/{os.path.basename(sheet_rels_path)}'

                        print(f"Looking for rels file: {sheet_rels_filename}")
                        print(
                            f"Available files in zip: {[f for f in excel_zip.namelist() if 'rels' in f]}"
                        )

                        try:
                            if sheet_rels_filename in excel_zip.namelist():
                                print(
                                    f"Found rels file: {sheet_rels_filename}")
                                with excel_zip.open(
                                        sheet_rels_filename) as sheet_rels:
                                    sheet_rels_tree = ET.parse(sheet_rels)
                                    sheet_rels_root = sheet_rels_tree.getroot()

                                    # drawingへの参照を探す
                                    for sheet_rel in sheet_rels_root.findall(
                                            './/pr:Relationship', self.ns):
                                        rel_type = sheet_rel.get('Type', '')
                                        rel_target = sheet_rel.get(
                                            'Target', '')
                                        print(
                                            f"Found relationship - Type: {rel_type}, Target: {rel_target}"
                                        )

                                        if 'drawing' in rel_target.lower():
                                            drawing_path = rel_target.replace(
                                                '..', 'xl')
                                            sheet_drawing_map[
                                                sheet_name] = drawing_path
                                            print(
                                                f"Found drawing for sheet '{sheet_name}': {drawing_path}"
                                            )
                            else:
                                print(
                                    f"Rels file not found: {sheet_rels_filename}"
                                )

                        except KeyError as e:
                            print(
                                f"No drawing relations found for sheet {sheet_name}: {str(e)}"
                            )
                            continue
                        except Exception as e:
                            print(
                                f"Error processing sheet relations for sheet {sheet_name}: {str(e)}"
                            )
                            continue

        except Exception as e:
            print(
                f"Error getting sheet-drawing relations: {str(e)}\n{traceback.format_exc()}"
            )

        print(
            f"\nFinal sheet-drawing map: {json.dumps(sheet_drawing_map, indent=2)}"
        )
        return sheet_drawing_map

    def extract_drawing_info(self, sheet, excel_zip,
                             drawing_path) -> List[Dict[str, Any]]:
        """Extract information about images and shapes from drawing.xml"""
        drawing_list = []

        try:
            print(f"\nAttempting to open drawing file: {drawing_path}")
            print(f"Available files in zip: {excel_zip.namelist()}")

            # drawing.xmlを解析
            with excel_zip.open(drawing_path) as xml_file:
                print(f"Successfully opened {drawing_path}")
                tree = ET.parse(xml_file)
                root = tree.getroot()
                print(f"XML Root tag: {root.tag}")
                print(
                    f"XML Namespaces: {root.nsmap if hasattr(root, 'nsmap') else 'Default namespace'}"
                )

                # すべてのアンカー要素を検索して、その中の図形を処理
                anchors = (root.findall('.//xdr:twoCellAnchor', self.ns) +
                           root.findall('.//xdr:oneCellAnchor', self.ns) +
                           root.findall('.//xdr:absoluteAnchor', self.ns))

                shape_count = 0
                for anchor in anchors:
                    sp_list = anchor.findall('.//xdr:sp', self.ns)
                    for idx, sp in enumerate(sp_list, 1):
                        shape_count += 1
                        print(f"\nProcessing shape #{shape_count}")
                        shape_info = self._extract_shape_info(sp, anchor)
                        if shape_info:
                            print(
                                f"Extracted shape info: {json.dumps(shape_info, indent=2)}"
                            )
                            drawing_list.append(shape_info)
                            print(f"Added shape at {shape_info['range']}")
                            if shape_info.get('text_content'):
                                print(
                                    f"  Text content: {shape_info['text_content']}"
                                )
                        else:
                            print(
                                f"Failed to extract info for shape #{shape_count}"
                            )

                print(f"\nFound {shape_count} shape elements")

                # 残りのアンカー要素を処理
                anchors = (root.findall('.//xdr:twoCellAnchor', self.ns) +
                           root.findall('.//xdr:oneCellAnchor', self.ns) +
                           root.findall('.//xdr:absoluteAnchor', self.ns))

                for anchor in anchors:
                    try:
                        # 図形の位置情報を取得
                        from_elem = anchor.find('.//xdr:from', self.ns)
                        to_elem = anchor.find('.//xdr:to', self.ns) or \
                                    anchor.find('.//xdr:ext', self.ns)  # oneCellAnchorの場合

                        # absoluteAnchorの場合は位置情報を変換
                        if anchor.tag.endswith('absoluteAnchor'):
                            pos = anchor.find('.//xdr:pos', self.ns)
                            ext = anchor.find('.//xdr:ext', self.ns)
                            if pos is not None and ext is not None:
                                # EMUからセル座標への概算変換
                                from_col = int(
                                    int(pos.get('x', '0')) /
                                    914400)  # 1インチ = 914400 EMU
                                from_row = int(int(pos.get('y', '0')) / 914400)
                                to_col = from_col + int(
                                    int(ext.get('cx', '0')) / 914400)
                                to_row = from_row + int(
                                    int(ext.get('cy', '0')) / 914400)
                        else:
                            if from_elem is None:
                                continue

                            # 通常の座標情報を取得
                            from_col = int(
                                from_elem.find('xdr:col', self.ns).text)
                            from_row = int(
                                from_elem.find('xdr:row', self.ns).text)

                            if to_elem is not None:
                                if anchor.tag.endswith('twoCellAnchor'):
                                    to_col = int(
                                        to_elem.find('xdr:col', self.ns).text)
                                    to_row = int(
                                        to_elem.find('xdr:row', self.ns).text)
                                else:  # oneCellAnchor
                                    cx = int(to_elem.get('cx', '0'))
                                    cy = int(to_elem.get('cy', '0'))
                                    # EMUからセル数への変換
                                    to_col = from_col + (cx // 914400)
                                    to_row = from_row + (cy // 914400)
                            else:
                                to_col = from_col + 1
                                to_row = from_row + 1

                        # 図形要素の検出と解析
                        drawing_elements = []

                        # 2. グループ形状 (xdr:grpSp)
                        for grp in anchor.findall('.//xdr:grpSp', self.ns):
                            group_info = self._extract_group_info(grp)
                            if group_info:
                                drawing_elements.append(group_info)

                        # 3. コネクタ形状 (xdr:cxnSp)
                        for cxn in anchor.findall('.//xdr:cxnSp', self.ns):
                            connector_info = self._extract_connector_info(cxn)
                            if connector_info:
                                drawing_elements.append(connector_info)

                        # 4. 画像 (xdr:pic)
                        for pic in anchor.findall('.//xdr:pic', self.ns):
                            picture_info = self._extract_picture_info(pic)
                            if picture_info:
                                drawing_elements.append(picture_info)

                        # 5. グラフ (c:chart)
                        chart = anchor.find('.//c:chart', self.ns)
                        if chart is not None:
                            chart_info = {
                                "type": "chart",
                                "chart_ref": chart.get(f'{{{self.ns["r"]}}}id')
                            }
                            drawing_elements.append(chart_info)

                        # 各図形要素に共通の座標情報を追加
                        for element in drawing_elements:
                            element.update({
                                "range":
                                f"{get_column_letter(from_col + 1)}{from_row + 1}:"
                                f"{get_column_letter(to_col + 1)}{to_row + 1}",
                                "coordinates": {
                                    "from": {
                                        "col": from_col + 1,
                                        "row": from_row + 1
                                    },
                                    "to": {
                                        "col": to_col + 1,
                                        "row": to_row + 1
                                    }
                                }
                            })
                            drawing_list.append(element)

                    except Exception as e:
                        print(f"Error processing anchor: {str(e)}")
                        continue

        except KeyError as e:
            print(
                f"KeyError in extract_drawing_info: {str(e)}\n{traceback.format_exc()}"
            )
        except Exception as e:
            print(
                f"Error extracting drawing info: {str(e)}\n{traceback.format_exc()}"
            )

        print(f"\nTotal drawings extracted: {len(drawing_list)}")
        return drawing_list

    def _extract_shape_info(self, sp_elem, anchor) -> Optional[Dict[str, Any]]:
        """Extract information from a shape element (xdr:sp)"""
        try:
            # 非表示情報を取得 (xdr:nvSpPr)
            nv_sp_pr = sp_elem.find('.//xdr:nvSpPr', self.ns)
            if nv_sp_pr is None:
                return None

            # 形状情報を取得 (xdr:spPr)
            sp_pr = sp_elem.find('.//xdr:spPr', self.ns)
            if sp_pr is None:
                return None

            # テキスト情報の取得（日本語対応、段落区切り対応）
            texts = []
            for p_elem in sp_elem.findall('.//a:p', self.ns):
                paragraph_texts = []
                # テキスト要素を一度だけ処理
                for t_elem in p_elem.findall('.//a:t', self.ns):
                    if t_elem is not None and t_elem.text:
                        paragraph_texts.append(t_elem.text)
                # 段落のテキストを結合
                if paragraph_texts:
                    texts.append(''.join(paragraph_texts))

            # 段落を改行で結合
            text_content = '\n'.join(texts) if texts else ''

            # 基本情報の構築
            shape_info = {
                "type": "shape",
                "name": nv_sp_pr.find('.//xdr:cNvPr', self.ns).get('name', ''),
                "description": nv_sp_pr.find('.//xdr:cNvPr', self.ns).get('descr', ''),
                "hidden": nv_sp_pr.find('.//xdr:cNvSpPr',
                                      self.ns).get('hidden', 'false') == 'true',
                "text_content": text_content
            }

            # プリセット形状の情報を追加
            preset_geom = sp_pr.find('.//a:prstGeom', self.ns)
            if preset_geom is not None and preset_geom.get('prst'):
                shape_info["shape_type"] = preset_geom.get('prst')

            # アンカーから座標情報を取得
            from_elem = anchor.find('xdr:from', self.ns)
            to_elem = anchor.find('xdr:to', self.ns)

            if from_elem is not None and to_elem is not None:
                # セル座標の取得
                from_col = int(from_elem.find('xdr:col', self.ns).text)
                from_row = int(from_elem.find('xdr:row', self.ns).text)
                to_col = int(to_elem.find('xdr:col', self.ns).text)
                to_row = int(to_elem.find('xdr:row', self.ns).text)

                # オフセットの取得（EMU単位）
                from_col_off = int(from_elem.find('xdr:colOff', self.ns).text)
                from_row_off = int(from_elem.find('xdr:rowOff', self.ns).text)
                to_col_off = int(to_elem.find('xdr:colOff', self.ns).text)
                to_row_off = int(to_elem.find('xdr:rowOff', self.ns).text)

                # EMUをセル単位に変換（1セル = 914400 EMU）
                EMU_PER_CELL = 914400

                # オフセットを小数点以下まで計算
                from_col_adj = from_col + (from_col_off / EMU_PER_CELL)
                from_row_adj = from_row + (from_row_off / EMU_PER_CELL)
                to_col_adj = to_col + (to_col_off / EMU_PER_CELL)
                to_row_adj = to_row + (to_row_off / EMU_PER_CELL)

                # 小数点以下を切り捨てて座標を計算
                from_col_final = math.floor(from_col_adj)
                from_row_final = math.floor(from_row_adj)
                to_col_final = math.floor(to_col_adj)
                to_row_final = math.floor(to_row_adj)

                shape_info["coordinates"] = {
                    "from": {
                        "col": from_col_final + 1,
                        "row": from_row_final + 1
                    },
                    "to": {
                        "col": to_col_final + 1,
                        "row": to_row_final + 1
                    }
                }

                # レンジ文字列を生成（実際のセル座標を使用）
                shape_info["range"] = (
                    f"{get_column_letter(from_col + 1)}{from_row + 1}:"
                    f"{get_column_letter(to_col + 1)}{to_row + 1}")

            return shape_info

        except Exception as e:
            print(
                f"Error extracting shape info: {str(e)}\n{traceback.format_exc()}"
            )
            return None

    def _emu_to_cell_coordinates(self, x: Optional[str], y: Optional[str],
                                 cx: Optional[str],
                                 cy: Optional[str]) -> Dict[str, Any]:
        """Convert EMU coordinates to cell coordinates"""
        try:
            # EMUからセル座標への変換（1インチ = 914400 EMU）
            EMU_PER_INCH = 914400
            CELLS_PER_INCH = 6  # おおよその値

            if all(v is not None for v in [x, y, cx, cy]):
                from_col = int(int(x) / EMU_PER_INCH * CELLS_PER_INCH)
                from_row = int(int(y) / EMU_PER_INCH * CELLS_PER_INCH)
                width_cells = int(int(cx) / EMU_PER_INCH * CELLS_PER_INCH)
                height_cells = int(int(cy) / EMU_PER_INCH * CELLS_PER_INCH)

                return {
                    "from": {
                        "col": from_col + 1,
                        "row": from_row + 1
                    },
                    "to": {
                        "col": from_col + width_cells + 1,
                        "row": from_row + height_cells + 1
                    }
                }
            else:
                return {
                    "from": {
                        "col": 1,
                        "row": 1
                    },
                    "to": {
                        "col": 2,
                        "row": 2
                    }
                }

        except Exception as e:
            print(f"Error converting EMU coordinates: {str(e)}")
            return {"from": {"col": 1, "row": 1}, "to": {"col": 2, "row": 2}}

    def _extract_group_info(self, grp_elem) -> Optional[Dict[str, Any]]:
        """Extract information from a group shape element (xdr:grpSp)"""
        try:
            nv_grp_sp_pr = grp_elem.find('.//xdr:nvGrpSpPr', self.ns)
            if nv_grp_sp_pr is None:
                return None

            # グループ内の図形数をカウント
            shapes_count = len(grp_elem.findall('.//xdr:sp', self.ns))
            pics_count = len(grp_elem.findall('.//xdr:pic', self.ns))

            return {
                "type":
                "group",
                "name":
                nv_grp_sp_pr.find('.//xdr:cNvPr', self.ns).get('name', ''),
                "description":
                nv_grp_sp_pr.find('.//xdr:cNvPr', self.ns).get('descr', ''),
                "shapes_count":
                shapes_count,
                "pictures_count":
                pics_count
            }
        except Exception as e:
            print(f"Error extracting group info: {str(e)}")
            return None

    def _extract_connector_info(self, cxn_elem) -> Optional[Dict[str, Any]]:
        """Extract information from a connector shape element (xdr:cxnSp)"""
        try:
            nv_cxn_sp_pr = cxn_elem.find('.//xdr:nvCxnSpPr', self.ns)
            if nv_cxn_sp_pr is None:
                return None

            return {
                "type":
                "connector",
                "name":
                nv_cxn_sp_pr.find('.//xdr:cNvPr', self.ns).get('name', ''),
                "description":
                nv_cxn_sp_pr.find('.//xdr:cNvPr', self.ns).get('descr', '')
            }
        except Exception as e:
            print(f"Error extracting connector info: {str(e)}")
            return None

    def _extract_picture_info(self, pic_elem) -> Optional[Dict[str, Any]]:
        """Extract information from a picture element (xdr:pic)"""
        try:
            nv_pic_pr = pic_elem.find('.//xdr:nvPicPr', self.ns)
            if nv_pic_pr is None:
                return None

            # 画像参照情報を取得
            blip = pic_elem.find('.//a:blip', self.ns)
            image_ref = blip.get(
                f'{{{self.ns["r"]}}}embed') if blip is not None else None

            return {
                "type":
                "image",
                "name":
                nv_pic_pr.find('.//xdr:cNvPr', self.ns).get('name', ''),
                "description":
                nv_pic_pr.find('.//xdr:cNvPr', self.ns).get('descr', ''),
                "image_ref":
                image_ref
            }
        except Exception as e:
            print(f"Error extracting picture info: {str(e)}")
            return None

    def detect_regions(self, sheet) -> List[Dict[str, Any]]:
        """Enhanced region detection including drawings and overlapping regions"""
        regions = []
        drawing_regions = []  # 描画オブジェクト由来の領域
        cell_regions = []  # セル由来の領域
        processed_cells = set()

        try:
            # まず画像・図形領域を検出
            print(f"\nProcessing drawings in sheet: {sheet.title}")

            # 一時ディレクトリを作成してZipファイルを展開
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_zip = os.path.join(temp_dir, 'temp.xlsx')
                with open(temp_zip, 'wb') as f:
                    self.file_obj.seek(0)
                    f.write(self.file_obj.read())

                with zipfile.ZipFile(temp_zip, 'r') as excel_zip:
                    # シートとdrawingの関連付けを取得
                    sheet_drawing_map = self.get_sheet_drawing_relations(
                        excel_zip)

                    # 現在のシートに対応するdrawingを探す
                    sheet_name = sheet.title
                    if sheet_name in sheet_drawing_map:
                        drawing_path = sheet_drawing_map[sheet_name]
                        drawings = self.extract_drawing_info(
                            sheet, excel_zip, drawing_path)

                        for drawing in drawings:
                            drawing_type = drawing[
                                "type"]  # "image", "shape", "smartart", "chart"
                            region_info = {
                                "regionType": drawing_type,
                                "type": drawing_type,
                                "range": drawing["range"],
                                "name": drawing.get("name", ""),
                                "description": drawing.get("description", ""),
                                "coordinates": drawing["coordinates"],
                                "text_content":
                                drawing.get("text_content", "")
                            }

                            # 図形タイプ別の追加情報
                            if drawing_type == "image" and "image_ref" in drawing:
                                region_info["image_ref"] = drawing["image_ref"]
                            elif drawing_type == "smartart" and "diagram_type" in drawing:
                                region_info["diagram_type"] = drawing[
                                    "diagram_type"]
                            elif drawing_type == "chart" and "chart_ref" in drawing:
                                region_info["chart_ref"] = drawing["chart_ref"]

                            drawing_regions.append(region_info)
                            print(
                                f"Added drawing region: {region_info['range']}"
                            )

                            # 図形が占める領域をprocessed_cellsに追加
                            from_col = drawing["coordinates"]["from"]["col"]
                            from_row = drawing["coordinates"]["from"]["row"]
                            to_col = drawing["coordinates"]["to"]["col"]
                            to_row = drawing["coordinates"]["to"]["row"]

                            for r in range(from_row, to_row + 1):
                                for c in range(from_col, to_col + 1):
                                    processed_cells.add(
                                        f"{get_column_letter(c)}{r}")

            # 残りのセル領域を検出
            print("\nProcessing remaining cell regions")
            for row in range(1, min(sheet.max_row + 1, 100)):
                for col in range(1, min(sheet.max_column + 1, 20)):
                    cell_coord = f"{get_column_letter(col)}{row}"

                    if cell_coord in processed_cells or sheet.cell(
                            row=row, column=col).value is None:
                        continue

                    # Find region boundaries
                    max_row, max_col = self.find_region_boundaries(
                        sheet, row, col)

                    with st.expander(f"Region: row:{max_row},col:{max_col}"):
                        st.write(
                            f"Region end : {get_column_letter(max_col)}{max_row}"
                        )

                    # 不要な区切り文字のみの場合はスキップ
                    cell_value = sheet.cell(row=row, column=col).value
                    if isinstance(cell_value, str) and len(cell_value.strip(
                    )) == 1 and cell_value.strip() in '-_=':
                        continue

                    # Extract cells data with limits
                    cells_data = self.extract_region_cells(
                        sheet, row, col, max_row, max_col)

                    # Get merged cells information
                    merged_cells = self.get_merged_cells_info(
                        sheet, row, col, max_row, max_col)

                    # Mark cells as processed
                    for r in range(row, max_row + 1):
                        for c in range(col, max_col + 1):
                            processed_cells.add(f"{get_column_letter(c)}{r}")

                    # Analyze region type
                    region_analysis = self.openai_helper.analyze_region_type(
                        json.dumps({
                            "cells": cells_data[:5],
                            "mergedCells": merged_cells[:3]
                        }))
                    if isinstance(region_analysis, str):
                        region_analysis = json.loads(region_analysis)

                    region_type = region_analysis.get("regionType", "unknown")

                    # Create basic region metadata
                    region_metadata = {
                        "regionType": region_type,
                        "range":
                        f"{get_column_letter(col)}{row}:{get_column_letter(max_col)}{max_row}",
                        "sampleCells": cells_data[:3],
                        "mergedCells": merged_cells
                    }

                    # Add table-specific metadata if it's a table region
                    if region_type == "table":
                        # Analyze header structure
                        header_structure = self.detect_header_structure(
                            cells_data, merged_cells)
                        if isinstance(header_structure, str):
                            header_structure = json.loads(header_structure)

                        # Calculate header range only if header rows were found
                        if header_structure.get("headerRows"):
                            header_rows = header_structure["headerRows"]
                            if header_rows:
                                min_header_row = min(header_rows)
                                max_header_row = max(header_rows)
                                # ヘッダーのタイプに応じて範囲を計算
                                if header_structure.get(
                                        "headerType") == "single":
                                # 単一ヘッダーの場合は同じ行を指定
                                    header_range = f"{min_header_row}"
                            else:
                                # 複合ヘッダーの場合は範囲を指定
                                header_range = f"{min_header_row}-{max_header_row}"
                        else:
                            header_range = "N/A"

                        region_metadata["headerStructure"] = {
                            "headerType":
                            header_structure.get("headerType", "none"),
                            "headerRowsCount":
                            header_structure.get("headerRowsCount", 0),
                            "headerRange":
                            header_range,
                            "mergedCells":
                            bool(merged_cells)
                        }

                    cell_regions.append(region_metadata)
                    print(f"Added cell region: {region_metadata['range']}")

                    # 領域数も制限する
                    if len(regions) >= 10:  # 最大10領域まで
                        print(
                            "Warning: Maximum number of regions reached, stopping analysis"
                        )
                        return regions

            # 描画オブジェクトとセル領域を両方保持（重複を許可）
            regions.extend(drawing_regions)
            regions.extend(cell_regions)

            print(
                f"\nTotal regions detected: {len(regions)} (Drawings: {len(drawing_regions)}, Cells: {len(cell_regions)})"
            )
            return regions
        except Exception as e:
            print(
                f"Error in detect_regions: {str(e)}\n{traceback.format_exc()}"
            )
            raise

    def get_file_metadata(self) -> Dict[str, Any]:
        """Extract file-level metadata"""
        try:
            properties = self.workbook.properties

            return {
                "fileName": self.file_obj.name,
                "fileProperties": {
                    "createdTime":
                    properties.created.isoformat()
                    if properties.created else None,
                    "modifiedTime":
                    properties.modified.isoformat()
                    if properties.modified else None,
                    "fileSize":
                    self.file_obj.size,
                    "author":
                    properties.creator,
                    "lastModifiedBy":
                    properties.lastModifiedBy,
                    "isPasswordProtected":
                    False  # Basic implementation
                }
            }
        except Exception as e:
            print(
                f"Error in get_file_metadata: {str(e)}\n{traceback.format_exc()}"
            )
            raise

    def analyze_cell_type(self, cell) -> str:
        """Analyze the type of a cell's content"""
        if cell.value is None:
            return "empty"
        if isinstance(cell.value, (int, float)):
            return "numeric"
        if isinstance(cell.value, datetime):
            return "date"
        return "text"

    def find_region_boundaries(self, sheet, start_row: int,
                               start_col: int) -> Tuple[int, int]:
        """Find the boundaries of a contiguous region with improved detection"""
        max_row = start_row
        max_col = start_col
        min_empty_rows = 1  # 空白行が1行以上続いたら領域の終わりとみなす
        min_empty_cols = 1  # 空白列が1列以上続いたら領域の終わりとみなす

        # 下方向のスキャン
        empty_row_count = 0
        for row in range(start_row, min(sheet.max_row + 1,
                                        start_row + 1000)):  # 1000行を上限に
            # 現在の行が空かどうかチェック
            row_empty = True
            for col in range(start_col,
                             min(start_col + 20,
                                 sheet.max_column + 1)):  # 20列をサンプルに
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

        # 右方向のスキャン
        empty_col_count = 0
        for col in range(start_col, min(sheet.max_column + 1,
                                        start_col + 50)):  # 50列を上限に
            # 現在の列が空かどうかチェック
            col_empty = True
            for row in range(start_row, min(max_row + 1,
                                            start_row + 50)):  # 20行をサンプルに
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

        # 最小でも元の位置を維持
        max_row = max(max_row, start_row)
        max_col = max(max_col, start_col)

        return max_row, max_col

    def get_merged_cells_info(self, sheet, start_row: int, start_col: int,
                              max_row: int,
                              max_col: int) -> List[Dict[str, Any]]:
        """Get information about merged cells in the region"""
        merged_cells_info = []
        for merged_range in sheet.merged_cells.ranges:
            if (merged_range.min_row >= start_row
                    and merged_range.max_row <= max_row
                    and merged_range.min_col >= start_col
                    and merged_range.max_col <= max_col):
                merged_cells_info.append({
                    "range":
                    str(merged_range),
                    "value":
                    sheet.cell(row=merged_range.min_row,
                               column=merged_range.min_col).value
                })
        return merged_cells_info

    def extract_region_cells(self, sheet, start_row: int, start_col: int,
                             max_row: int,
                             max_col: int) -> List[List[Dict[str, Any]]]:
        """Extract cell information from a region with limits"""
        cells_data = []
        # 範囲が大きすぎる場合は制限する
        actual_max_row = min(
            max_row, start_row + self.MAX_CELLS_PER_ANALYSIS //
            (max_col - start_col + 1))
        actual_max_col = min(
            max_col, start_col + self.MAX_CELLS_PER_ANALYSIS //
            (max_row - start_row + 1))

        for row in range(start_row, actual_max_row + 1):
            row_data = []
            for col in range(start_col, actual_max_col + 1):
                cell = sheet.cell(row=row, column=col)
                cell_type = self.analyze_cell_type(cell)

                # Handle merged cells
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    for merged_range in sheet.merged_cells.ranges:
                        if merged_range.min_row <= row <= merged_range.max_row and \
                           merged_range.min_col <= col <= merged_range.max_col:
                            master_cell = sheet.cell(
                                row=merged_range.min_row,
                                column=merged_range.min_col)
                            cell_info = {
                                "row":
                                row,
                                "col":
                                col,
                                "value":
                                str(master_cell.value)
                                if master_cell.value is not None else "",
                                "type":
                                cell_type,
                                "isMerged":
                                True,
                                "mergedRange":
                                str(merged_range)
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
                    # Regular cell - 必要最小限の情報のみを含める
                    cell_info = {
                        "row": row,
                        "col": col,
                        "value":
                        str(cell.value) if cell.value is not None else "",
                        "type": cell_type
                    }

                row_data.append(cell_info)
            cells_data.append(row_data)

        if max_row > actual_max_row or max_col > actual_max_col:
            print(
                f"Note: Region was truncated from {max_row}x{max_col} to {actual_max_row}x{actual_max_col}"
            )

        return cells_data

    def detect_header_structure(
            self, cells_data: List[List[Dict[str, Any]]],
            merged_cells: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Analyze the header structure of a table region using LLM."""

        def json_to_markdown_table(json_data, include_row_col=False):
            """JSONデータをマークダウンテーブルに変換する。

            Args:
                json_data (dict): JSONデータ（cellsキーを持つ辞書）。
                include_row_col (bool, optional): rowとcolの情報を含めるかどうか。デフォルトはFalse。

            Returns:
                str: マークダウンテーブルの文字列。エラー時はNoneを返す。
            """
            try:
                cells = json_data["cells"]
            except (KeyError, TypeError):
                return None

            # 最大の列数を取得
            max_cols = 0
            for row in cells:
                max_cols = max(
                    max_cols,
                    max(cell["col"] for cell in row) + 1 if row else 0)

            markdown = ""

            if include_row_col:
                # ヘッダー行（列番号）
                markdown += "|       |"  # 左上の空白セル
                for col in range(max_cols):
                    markdown += f" col {col} |"
                markdown += "\n"

            for i, row_data in enumerate(cells):
                if include_row_col:
                    markdown += f"| row {row_data[0]['row'] if row_data else ''} |"  # 行番号
                for col in range(max_cols):
                    cell_value = ""
                    for cell in row_data:
                        if cell["col"] == col:
                            cell_value = cell.get("value", "")
                            break
                    markdown += f" {cell_value} |"
                markdown += "\n"

            if include_row_col:
                markdown += "\n"  # テーブルと注釈の間に空行を入れる
                markdown += "※Reference:"
                for row_data in cells:
                    for cell in row_data:
                        markdown += f"* Row {cell['row']}, Col {cell['col']}: {cell.get('value', '')}\n"

            return markdown

        try:
            markdown_table = json_to_markdown_table({"cells": cells_data[:5]},
                                                    include_row_col=True)
            analysis_str = self.openai_helper.analyze_table_structure(
                markdown_table)

            if isinstance(analysis_str, str):
                analysis = json.loads(analysis_str)
            else:
                return {
                    "headerType": "none",
                    "headerRowsCount": 0,
                    "headerRows": [],
                    "headerRange": "N/A",
                    "confidence": 0
                }

            header_type = analysis.get("headerStructure",
                                       {}).get("type", "none")
            header_rows_zero_based = analysis.get("headerStructure",
                                                  {}).get("rows", [])
            header_rows = [row + 1
                           for row in header_rows_zero_based]  # 0始まりを1始まりに変換

            # ヘッダー行の範囲を計算
            if header_rows:
                min_row = min(header_rows)
                max_row = max(header_rows)
                header_range = f"{min_row}-{max_row}"
            else:
                header_range = "N/A"

            return {
                "headerType": header_type,
                "headerRowsCount": len(header_rows),
                "headerRows": header_rows,
                "headerRange": header_range,
                "confidence": analysis.get("confidence", 0)
            }

        except Exception as e:
            print(f"Error in detect_header_structure: {str(e)}")
            return {
                "headerType": "none",
                "headerRowsCount": 0,
                "headerRows": [],
                "headerRange": "N/A",
                "confidence": 0
            }

    def get_sheet_metadata(self) -> list:
        """Extract enhanced sheet-level metadata"""
        try:
            sheets_metadata = []

            for sheet_name in self.workbook.sheetnames:
                sheet = self.workbook[sheet_name]

                # Get merged cells
                merged_cells = [
                    str(cell_range) for cell_range in sheet.merged_cells.ranges
                ]

                # Detect regions with enhanced analysis
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
            print(
                f"Error in get_sheet_metadata: {str(e)}\n{traceback.format_exc()}"
            )
            raise

    def extract_all_metadata(self) -> Dict[str, Any]:
        """Extract all metadata with enhanced analysis"""
        try:
            file_metadata = self.get_file_metadata()
            sheets_metadata = self.get_sheet_metadata()

            metadata = {
                **file_metadata,
                "worksheets": sheets_metadata,
                "crossSheetRelationships": []  # Basic implementation
            }

            return metadata
        except Exception as e:
            print(
                f"Error in extract_all_metadata: {str(e)}\n{traceback.format_exc()}"
            )
            raise
