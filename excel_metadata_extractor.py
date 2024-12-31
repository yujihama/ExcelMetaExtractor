import os
import json
from datetime import datetime
from openpyxl import load_workbook
import openpyxl.cell.cell  # Add this import for MergedCell type checking
from openpyxl.utils import get_column_letter
from typing import Dict, Any, List, Optional, Tuple
from openai_helper import OpenAIHelper
import traceback

class ExcelMetadataExtractor:
    def __init__(self, file_obj):
        self.file_obj = file_obj
        self.workbook = load_workbook(file_obj, data_only=True)
        self.openai_helper = OpenAIHelper()

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
        """Extract cell information from a region"""
        cells_data = []
        for row in range(start_row, max_row + 1):
            row_data = []
            for col in range(start_col, max_col + 1):
                cell = sheet.cell(row=row, column=col)
                cell_type = self.analyze_cell_type(cell)

                # Handle merged cells
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    # Find the master cell (top-left) of the merged range
                    for merged_range in sheet.merged_cells.ranges:
                        if merged_range.min_row <= row <= merged_range.max_row and \
                           merged_range.min_col <= col <= merged_range.max_col:
                            master_cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                            cell_info = {
                                "row": row,
                                "col": col,
                                "value": str(master_cell.value) if master_cell.value is not None else "",
                                "type": cell_type,
                                "formula": None,  # Merged cells don't have formulas
                                "isMerged": True,
                                "mergedRange": str(merged_range)
                            }
                            break
                    else:
                        cell_info = {"row": row, "col": col, "value": "", "type": cell_type, "formula": None, "isMerged": True}

                else:
                    # Regular cell
                    formula = cell.value if isinstance(cell.value, str) and cell.value.startswith('=') else None
                    cell_info = {
                        "row": row,
                        "col": col,
                        "value": str(cell.value) if cell.value is not None else "",
                        "type": cell_type,
                        "formula": formula,
                        "isMerged": False
                    }

                row_data.append(cell_info)
            cells_data.append(row_data)
        return cells_data

    def detect_header_structure(self, cells_data: List[List[Dict[str, Any]]]) -> Dict[str, Any]:
        """Analyze header structure using pattern recognition and LLM"""
        # Convert cells data to a format suitable for LLM analysis
        header_analysis = self.openai_helper.analyze_table_structure(json.dumps(cells_data))
        if isinstance(header_analysis, str):
            header_analysis = json.loads(header_analysis)

        return {
            "headerType": header_analysis.get("headerType", "single"),
            "headerRowsCount": header_analysis.get("headerRowsCount", 1),
            "confidence": header_analysis.get("confidence", 0.0)
        }

    def detect_regions(self, sheet) -> List[Dict[str, Any]]:
        """Enhanced region detection with LLM assistance"""
        try:
            regions = []
            processed_cells = set()

            for row in range(1, sheet.max_row + 1):
                for col in range(1, sheet.max_column + 1):
                    cell_coord = f"{get_column_letter(col)}{row}"

                    if cell_coord in processed_cells or sheet.cell(row=row, column=col).value is None:
                        continue

                    # Find region boundaries
                    max_row, max_col = self.find_region_boundaries(sheet, row, col)

                    # Skip if region is too small
                    if max_row - row < 1 or max_col - col < 1:
                        continue

                    # Extract cells data
                    cells_data = self.extract_region_cells(sheet, row, col, max_row, max_col)

                    # Get merged cells information
                    merged_cells = self.get_merged_cells_info(sheet, row, col, max_row, max_col)

                    # Mark cells as processed
                    for r in range(row, max_row + 1):
                        for c in range(col, max_col + 1):
                            processed_cells.add(f"{get_column_letter(c)}{r}")

                    # Analyze region type and structure
                    region_analysis = self.openai_helper.analyze_region_type(json.dumps({
                        "cells": cells_data,
                        "mergedCells": merged_cells
                    }))

                    if isinstance(region_analysis, str):
                        region_analysis = json.loads(region_analysis)

                    region_type = region_analysis.get("regionType", "unknown")

                    # Create region metadata based on type
                    region_metadata = {
                        "regionType": region_type,
                        "range": f"{get_column_letter(col)}{row}:{get_column_letter(max_col)}{max_row}",
                        "cells": cells_data,
                        "mergedCells": merged_cells
                    }

                    if region_type == "table":
                        header_structure = self.detect_header_structure(cells_data)
                        region_metadata.update({
                            "headerStructure": header_structure,
                            "purpose": region_analysis.get("purpose", "")
                        })
                    elif region_type == "text":
                        text_content = "\n".join(
                            str(cell["value"]) for row in cells_data for cell in row if cell["value"]
                        )
                        text_analysis = self.openai_helper.analyze_text_block(text_content)
                        if isinstance(text_analysis, str):
                            text_analysis = json.loads(text_analysis)
                        region_metadata.update({
                            "content": text_content,
                            "classification": text_analysis.get("contentType", "unknown"),
                            "importance": text_analysis.get("importance", "medium"),
                            "summary": text_analysis.get("summary", ""),
                            "keyPoints": text_analysis.get("keyPoints", [])
                        })
                    elif region_type == "chart":
                        # Handle charts and images
                        chart_elements = region_analysis.get("chartElements", {})
                        chart_analysis = self.openai_helper.analyze_chart(json.dumps(chart_elements))
                        if isinstance(chart_analysis, str):
                            chart_analysis = json.loads(chart_analysis)
                        region_metadata.update({
                            "chartType": chart_analysis.get("chartType", "unknown"),
                            "purpose": chart_analysis.get("purpose", ""),
                            "dataRelations": chart_analysis.get("dataRelations", []),
                            "suggestedUsage": chart_analysis.get("suggestedUsage", "")
                        })

                    regions.append(region_metadata)

            return regions
        except Exception as e:
            print(f"Error in detect_regions: {str(e)}\n{traceback.format_exc()}")
            raise

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