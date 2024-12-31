from datetime import datetime
import os
from openpyxl import load_workbook
from typing import Dict, Any

class ExcelMetadataExtractor:
    def __init__(self, file_obj):
        self.file_obj = file_obj
        self.workbook = load_workbook(file_obj, data_only=True)

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
            
            sheet_meta = {
                "sheetName": sheet_name,
                "isProtected": sheet.protection.sheet,
                "rowCount": max_row,
                "columnCount": max_col,
                "hasPivotTables": len(sheet._pivots) > 0,
                "hasCharts": len(sheet._charts) > 0,
                "mergedCells": merged_cells
            }
            
            sheets_metadata.append(sheet_meta)
        
        return sheets_metadata

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
