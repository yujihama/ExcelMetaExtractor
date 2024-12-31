import os
import json
from datetime import datetime
from openpyxl import load_workbook
import openpyxl.cell.cell
from openpyxl.utils import get_column_letter
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from typing import Dict, Any, List, Optional, Tuple
from openai_helper import OpenAIHelper
import traceback

class ExcelMetadataExtractor:
    def __init__(self, file_obj):
        self.file_obj = file_obj
        self.workbook = load_workbook(file_obj, data_only=True)
        self.openai_helper = OpenAIHelper()
        self.MAX_CELLS_PER_ANALYSIS = 100

    def extract_drawing_info(self, sheet) -> List[Dict[str, Any]]:
        """Extract information about images and shapes from drawing.xml"""
        drawing_list = []

        try:
            if not hasattr(sheet, '_drawings') or not sheet._drawings:
                return drawing_list

            for drawing in sheet._drawings:
                if isinstance(drawing, SpreadsheetDrawing):
                    for shape in drawing.shapes:
                        # 画像や図形の座標情報を取得
                        from_col, from_row = shape.anchor._from.col, shape.anchor._from.row
                        to_col, to_row = shape.anchor._to.col, shape.anchor._to.row

                        drawing_info = {
                            "type": "image" if hasattr(shape, 'image') else "shape",
                            "range": f"{get_column_letter(from_col + 1)}{from_row + 1}:"
                                    f"{get_column_letter(to_col + 1)}{to_row + 1}",
                            "description": getattr(shape, 'description', ''),
                            "name": getattr(shape, 'name', ''),
                            "coordinates": {
                                "from": {"col": from_col + 1, "row": from_row + 1},
                                "to": {"col": to_col + 1, "row": to_row + 1}
                            }
                        }

                        # 画像特有の情報を追加
                        if hasattr(shape, 'image'):
                            drawing_info.update({
                                "image_format": shape.image.format,
                                "image_ref": shape.image.ref,
                            })

                        drawing_list.append(drawing_info)

        except Exception as e:
            print(f"Error extracting drawing info: {str(e)}\n{traceback.format_exc()}")

        return drawing_list

    def detect_regions(self, sheet) -> List[Dict[str, Any]]:
        """Enhanced region detection including drawings"""
        regions = []
        processed_cells = set()

        try:
            # まず画像・図形領域を検出
            drawings = self.extract_drawing_info(sheet)
            for drawing in drawings:
                regions.append({
                    "regionType": "image" if drawing["type"] == "image" else "shape",
                    "range": drawing["range"],
                    "description": drawing["description"],
                    "name": drawing["name"],
                    "coordinates": drawing["coordinates"]
                })

                # 画像・図形が占める領域をprocessed_cellsに追加
                from_col = drawing["coordinates"]["from"]["col"]
                from_row = drawing["coordinates"]["from"]["row"]
                to_col = drawing["coordinates"]["to"]["col"]
                to_row = drawing["coordinates"]["to"]["row"]

                for r in range(from_row, to_row + 1):
                    for c in range(from_col, to_col + 1):
                        processed_cells.add(f"{get_column_letter(c)}{r}")

            # 残りのセル領域を検出（既存のコード）
            for row in range(1, min(sheet.max_row + 1, 100)):
                for col in range(1, min(sheet.max_column + 1, 20)):
                    cell_coord = f"{get_column_letter(col)}{row}"

                    if cell_coord in processed_cells or sheet.cell(row=row, column=col).value is None:
                        continue

                    # Find region boundaries
                    max_row, max_col = self.find_region_boundaries(sheet, row, col)

                    # Skip if region is too small
                    if max_row - row < 1 or max_col - col < 1:
                        continue

                    # Extract cells data with limits
                    cells_data = self.extract_region_cells(sheet, row, col, max_row, max_col)

                    # Get merged cells information (制限付き)
                    merged_cells = self.get_merged_cells_info(sheet, row, col, max_row, max_col)

                    # Mark cells as processed
                    for r in range(row, max_row + 1):
                        for c in range(col, max_col + 1):
                            processed_cells.add(f"{get_column_letter(c)}{r}")

                    # Analyze region type
                    region_analysis = self.openai_helper.analyze_region_type(json.dumps({
                        "cells": cells_data[:5],
                        "mergedCells": merged_cells[:3]
                    }))
                    if isinstance(region_analysis, str):
                        region_analysis = json.loads(region_analysis)

                    region_type = region_analysis.get("regionType", "unknown")

                    # Create basic region metadata
                    region_metadata = {
                        "regionType": region_type,
                        "range": f"{get_column_letter(col)}{row}:{get_column_letter(max_col)}{max_row}",
                        "sampleCells": cells_data[:3],
                        "mergedCells": merged_cells
                    }

                    # Add table-specific metadata if it's a table region
                    if region_type == "table":
                        # Analyze header structure
                        header_structure = self.detect_header_structure(cells_data)
                        if isinstance(header_structure, str):
                            header_structure = json.loads(header_structure)

                        # Calculate header range only if header rows were found
                        if header_structure.get("headerRows"):
                            header_rows = header_structure["headerRows"]
                            if header_rows:
                                min_header_row = min(header_rows)
                                max_header_row = max(header_rows)
                                header_range = f"{row + min_header_row}-{row + max_header_row}"
                            else:
                                header_range = "N/A"
                        else:
                            header_range = "N/A"

                        region_metadata["headerStructure"] = {
                            "headerType": header_structure.get("headerType", "none"),
                            "headerRowsCount": header_structure.get("headerRowsCount", 0),
                            "headerRange": header_range,
                            "mergedCells": bool(merged_cells)
                        }

                    regions.append(region_metadata)

                    # 領域数も制限する
                    if len(regions) >= 10:  # 最大10領域まで
                        print("Warning: Maximum number of regions reached, stopping analysis")
                        return regions

            return regions
        except Exception as e:
            print(f"Error in detect_regions: {str(e)}\n{traceback.format_exc()}")
            raise

    def get_file_metadata(self) -> Dict[str, Any]:
        """Extract file-level metadata"""
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
                    "isPasswordProtected": False  # Basic implementation
                }
            }
        except Exception as e:
            print(f"Error in get_file_metadata: {str(e)}\n{traceback.format_exc()}")
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

    def find_region_boundaries(self, sheet, start_row: int, start_col: int) -> Tuple[int, int]:
        """Find the boundaries of a contiguous region"""
        max_row = start_row
        max_col = start_col

        # Scan downward
        for row in range(start_row, sheet.max_row + 1):
            if all(sheet.cell(row=row, column=col).value is None
                  for col in range(start_col, min(start_col + 3, sheet.max_column + 1))):
                break
            max_row = row

        # Scan rightward
        for col in range(start_col, sheet.max_column + 1):
            if all(sheet.cell(row=row, column=col).value is None
                  for row in range(start_row, min(start_row + 3, max_row + 1))):
                break
            max_col = col

        return max_row, max_col

    def get_merged_cells_info(self, sheet, start_row: int, start_col: int, max_row: int, max_col: int) -> List[Dict[str, Any]]:
        """Get information about merged cells in the region"""
        merged_cells_info = []
        for merged_range in sheet.merged_cells.ranges:
            if (merged_range.min_row >= start_row and merged_range.max_row <= max_row and
                merged_range.min_col >= start_col and merged_range.max_col <= max_col):
                merged_cells_info.append({
                    "range": str(merged_range),
                    "value": sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
                })
        return merged_cells_info

    def extract_region_cells(self, sheet, start_row: int, start_col: int, max_row: int, max_col: int) -> List[List[Dict[str, Any]]]:
        """Extract cell information from a region with limits"""
        cells_data = []
        # 範囲が大きすぎる場合は制限する
        actual_max_row = min(max_row, start_row + self.MAX_CELLS_PER_ANALYSIS // (max_col - start_col + 1))
        actual_max_col = min(max_col, start_col + self.MAX_CELLS_PER_ANALYSIS // (max_row - start_row + 1))

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
                    # Regular cell - 必要最小限の情報のみを含める
                    cell_info = {
                        "row": row,
                        "col": col,
                        "value": str(cell.value) if cell.value is not None else "",
                        "type": cell_type
                    }

                row_data.append(cell_info)
            cells_data.append(row_data)

        if max_row > actual_max_row or max_col > actual_max_col:
            print(f"Warning: Region was truncated from {max_row-start_row+1}x{max_col-start_col+1} to {actual_max_row-start_row+1}x{actual_max_col-start_col+1} cells due to size limits")

        return cells_data

    def detect_header_structure(self, cells_data: List[List[Dict[str, Any]]]) -> Dict[str, Any]:
        """Analyze header structure using pattern recognition and LLM"""
        try:
            if not cells_data or not cells_data[0]:
                return {
                    "headerType": "none",
                    "headerRowsCount": 0,
                    "confidence": 0.0
                }

            # 結合セルパターンの分析
            merged_cells_in_first_rows = any(
                cell.get('isMerged', False)
                for row in cells_data[:2]  # 最初の2行を確認
                for cell in row
            )

            # ヘッダー候補行のデータ型分析
            header_rows = []
            data_rows = []

            for i, row in enumerate(cells_data[:4]):  # 最初の4行まで分析
                # 行の特徴を分析
                cell_types = [cell['type'] for cell in row]
                cell_values = [str(cell['value']).strip() for cell in row]

                # ヘッダーらしい特徴をチェック
                is_header_like = (
                    all(t in ['text', 'empty'] for t in cell_types) and  # テキストか空セルのみ
                    any(v != '' for v in cell_values) and  # 少なくとも1つは値がある
                    not any(v.isdigit() for v in cell_values if v)  # 数値のみの値がない
                )

                if is_header_like:
                    header_rows.append(i)
                elif any(v != '' for v in cell_values):
                    data_rows.append(i)
                    if header_rows:  # ヘッダー行が見つかった後にデータ行が来たら終了
                        break

            # ヘッダー構造の判定
            if not header_rows:
                return {
                    "headerType": "none",
                    "headerRowsCount": 0,
                    "confidence": 0.8
                }

            header_type = "multiple" if (len(header_rows) > 1 or merged_cells_in_first_rows) else "single"
            header_rows_count = max(header_rows) - min(header_rows) + 1

            return {
                "headerType": header_type,
                "headerRowsCount": header_rows_count,
                "headerRows": header_rows,
                "confidence": 0.9
            }

        except Exception as e:
            print(f"Error in detect_header_structure: {str(e)}")
            return {
                "headerType": "none",
                "headerRowsCount": 0,
                "confidence": 0.0
            }

    def get_sheet_metadata(self) -> list:
        """Extract enhanced sheet-level metadata"""
        try:
            sheets_metadata = []

            for sheet_name in self.workbook.sheetnames:
                sheet = self.workbook[sheet_name]

                # Get merged cells
                merged_cells = [str(cell_range) for cell_range in sheet.merged_cells.ranges]

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
            print(f"Error in get_sheet_metadata: {str(e)}\n{traceback.format_exc()}")
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
            print(f"Error in extract_all_metadata: {str(e)}\n{traceback.format_exc()}")
            raise