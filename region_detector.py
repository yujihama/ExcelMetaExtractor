
from typing import Tuple, List, Dict, Any
from openpyxl.utils import get_column_letter

class RegionDetector:
    def __init__(self):
        self.logger = Logger()

    def find_region_boundaries(self, sheet, start_row: int, start_col: int) -> Tuple[int, int]:
        max_row = start_row
        max_col = start_col
        current_row = start_row
        self.logger.debug_boundaries(start_row, start_col, sheet.max_row, sheet.max_column)
        
        # Find the bottom boundary
        while current_row <= sheet.max_row:
            if sheet.cell(row=current_row, column=start_col).value is not None:
                max_row = current_row
                current_row += 1
            else:
                # Check next few rows to confirm end of region
                empty_rows = 0
                for i in range(3):
                    if current_row + i <= sheet.max_row and sheet.cell(row=current_row + i, column=start_col).value is None:
                        empty_rows += 1
                if empty_rows == 3:
                    break
                current_row += 1

        # Find the right boundary
        current_col = start_col
        while current_col <= sheet.max_column:
            if sheet.cell(row=start_row, column=current_col).value is not None:
                max_col = current_col
                current_col += 1
            else:
                # Check next few columns to confirm end of region
                empty_cols = 0
                for i in range(3):
                    if current_col + i <= sheet.max_column and sheet.cell(row=start_row, column=current_col + i).value is None:
                        empty_cols += 1
                if empty_cols == 3:
                    break
                current_col += 1

        return max_row, max_col

    def get_merged_cells_info(self, sheet, start_row: int, start_col: int, max_row: int, max_col: int) -> List[Dict[str, Any]]:
        merged_cells_info = []
        for merged_range in sheet.merged_cells.ranges:
            if (merged_range.min_row >= start_row and merged_range.max_row <= max_row and
                merged_range.min_col >= start_col and merged_range.max_col <= max_col):
                merged_cells_info.append({
                    "range": str(merged_range),
                    "value": sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
                })
        return merged_cells_info
