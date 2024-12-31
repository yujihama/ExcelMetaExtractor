from datetime import datetime
import os
import json
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from typing import Dict, Any, List
from openai_helper import OpenAIHelper

class ExcelMetadataExtractor:
    def __init__(self, file_obj):
        self.file_obj = file_obj
        self.workbook = load_workbook(file_obj, data_only=True)
        self.openai_helper = OpenAIHelper()

    def get_file_metadata(self) -> Dict[str, Any]:
        """Extract file-level metadata"""
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

    def get_sheet_metadata(self) -> list:
        """Extract sheet-level metadata"""
        sheets_metadata = []

        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]

            # Get merged cells
            merged_cells = [str(cell_range) for cell_range in sheet.merged_cells.ranges]

            # Get sheet dimensions
            max_row = sheet.max_row
            max_col = sheet.max_column

            # Get pivots and charts
            pivot_tables = getattr(sheet, '_pivots', [])
            charts = getattr(sheet, '_charts', [])

            sheet_meta = {
                "sheetName": sheet_name,
                "isProtected": sheet.protection.sheet,
                "rowCount": max_row,
                "columnCount": max_col,
                "hasPivotTables": len(pivot_tables) > 0,
                "hasCharts": len(charts) > 0,
                "mergedCells": merged_cells,
                "regions": self.detect_regions(sheet)
            }

            sheets_metadata.append(sheet_meta)

        return sheets_metadata

    def detect_regions(self, sheet) -> List[Dict[str, Any]]:
        """Detect and analyze regions in the sheet"""
        regions = []
        max_row = sheet.max_row
        max_col = sheet.max_column

        # Track processed cells to avoid duplicate analysis
        processed_cells = set()

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell_coord = f"{get_column_letter(col)}{row}"

                if cell_coord in processed_cells:
                    continue

                # Get current cell
                cell = sheet[cell_coord]

                if cell.value is None:
                    continue

                # Detect table regions
                table_region = self.detect_table_region(sheet, row, col, processed_cells)
                if table_region:
                    regions.append(table_region)
                    continue

                # Detect text regions
                text_region = self.detect_text_region(sheet, row, col, processed_cells)
                if text_region:
                    regions.append(text_region)
                    continue

                # Detect image/chart regions
                image_region = self.detect_image_region(sheet, row, col)
                if image_region:
                    regions.append(image_region)
                    continue

        return regions

    def detect_table_region(self, sheet, start_row: int, start_col: int, processed_cells: set) -> Dict[str, Any]:
        """Detect and analyze table regions"""
        # Check surrounding cells to determine if it's a table
        max_row = start_row
        max_col = start_col

        # Find table boundaries
        for row in range(start_row, sheet.max_row + 1):
            if all(sheet.cell(row=row, column=col).value is None for col in range(start_col, start_col + 2)):
                break
            max_row = row

        for col in range(start_col, sheet.max_column + 1):
            if all(sheet.cell(row=row, column=col).value is None for row in range(start_row, start_row + 2)):
                break
            max_col = col

        # If region is too small, it's not a table
        if max_row - start_row < 1 or max_col - start_col < 1:
            return None

        # Extract cells data for analysis
        cells_data = []
        for row in range(start_row, max_row + 1):
            for col in range(start_col, max_col + 1):
                cell = sheet.cell(row=row, column=col)
                cell_coord = f"{get_column_letter(col)}{row}"
                processed_cells.add(cell_coord)

                cells_data.append({
                    "row": row,
                    "col": col,
                    "value": str(cell.value) if cell.value is not None else "",
                    "formula": cell.formula if cell.formula else None
                })

        # Analyze table structure using LLM
        analysis = json.loads(self.openai_helper.analyze_table_structure(json.dumps(cells_data)))

        if not analysis.get("isTable", False):
            return None

        return {
            "regionType": "table",
            "range": f"{get_column_letter(start_col)}{start_row}:{get_column_letter(max_col)}{max_row}",
            "headerStructure": {
                "headerType": analysis.get("headerType", "single"),
                "headerRowsCount": analysis.get("headerRowsCount", 1),
                "mergedCells": any(cell.coordinate in sheet.merged_cells for cell in sheet[f"{get_column_letter(start_col)}{start_row}:{get_column_letter(max_col)}{start_row}"])
            },
            "cells": cells_data,
            "notes": analysis.get("purpose", "")
        }

    def detect_text_region(self, sheet, start_row: int, start_col: int, processed_cells: set) -> Dict[str, Any]:
        """Detect and analyze text regions"""
        text_content = []
        max_row = start_row

        # Find text block boundaries
        for row in range(start_row, min(start_row + 10, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=start_col)
            if cell.value is None:
                break

            cell_coord = f"{get_column_letter(start_col)}{row}"
            if cell_coord in processed_cells:
                continue

            text_content.append(str(cell.value))
            processed_cells.add(cell_coord)
            max_row = row

        if not text_content:
            return None

        # Analyze text content using LLM
        analysis = json.loads(self.openai_helper.analyze_text_block("\n".join(text_content)))

        return {
            "regionType": "text",
            "range": f"{get_column_letter(start_col)}{start_row}:{get_column_letter(start_col)}{max_row}",
            "content": "\n".join(text_content),
            "classification": analysis.get("contentType", "unknown"),
            "importance": analysis.get("importance", "medium"),
            "summary": analysis.get("summary", ""),
            "keyPoints": analysis.get("keyPoints", [])
        }

    def detect_image_region(self, sheet, row: int, col: int) -> Dict[str, Any]:
        """Detect and analyze image/chart regions"""
        charts = getattr(sheet, '_charts', [])
        for chart in charts:
            if hasattr(chart, '_chart_type'):
                chart_elements = {
                    "type": chart._chart_type,
                    "title": chart.title.text if hasattr(chart.title, 'text') else None,
                    "hasLegend": chart.has_legend if hasattr(chart, 'has_legend') else False
                }

                # Analyze chart using LLM
                analysis = json.loads(self.openai_helper.analyze_chart(chart_elements))

                return {
                    "regionType": "chart",
                    "chartType": analysis.get("chartType", chart._chart_type),
                    "purpose": analysis.get("purpose", ""),
                    "dataRelations": analysis.get("dataRelations", []),
                    "suggestedUsage": analysis.get("suggestedUsage", "")
                }
        return None

    def extract_all_metadata(self) -> Dict[str, Any]:
        """Extract all metadata and return in specified JSON format"""
        file_metadata = self.get_file_metadata()
        sheets_metadata = self.get_sheet_metadata()

        metadata = {
            **file_metadata,
            "worksheets": sheets_metadata,
            "crossSheetRelationships": []  # Basic implementation
        }

        return metadata