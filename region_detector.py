
from typing import Tuple, List, Dict, Any
from openpyxl.utils import get_column_letter
from logger import Logger

class RegionDetector:
    def __init__(self):
        self.logger = Logger()

    def find_region_boundaries(self, sheet, start_row: int, start_col: int) -> Tuple[int, int]:
        max_row = start_row
        max_col = start_col
        current_row = start_row
        self.logger.debug_boundaries(start_row, start_col, sheet.max_row, sheet.max_column)
        self.logger.info(f"Starting boundary detection from cell ({start_row}, {start_col})")
        
        # Find the bottom boundary
        while current_row <= sheet.max_row:
            row_empty = True
            # 行全体が空かどうかをチェック
            for check_col in range(start_col, sheet.max_column + 1):
                if sheet.cell(row=current_row, column=check_col).value is not None:
                    row_empty = False
                    break

            if not row_empty:
                max_row = current_row
                current_row += 1
            else:
                # 次の行も確認
                next_row_empty = True
                if current_row + 1 <= sheet.max_row:
                    for check_col in range(start_col, sheet.max_column + 1):
                        if sheet.cell(row=current_row + 1, column=check_col).value is not None:
                            next_row_empty = False
                            break
                    if not next_row_empty:
                        current_row += 1
                        continue
                break

        # Find the right boundary
        current_col = start_col
        while current_col <= sheet.max_column:
            col_empty = True
            # 現在の列が空かどうかをチェック
            for check_row in range(start_row, max_row + 1):
                if sheet.cell(row=check_row, column=current_col).value is not None:
                    col_empty = False
                    max_col = current_col
                    break
            
            if col_empty:
                # 次の列も確認
                next_col_empty = True
                if current_col + 1 <= sheet.max_column:
                    for check_row in range(start_row, max_row + 1):
                        if sheet.cell(row=check_row, column=current_col + 1).value is not None:
                            next_col_empty = False
                            break
                    if next_col_empty:
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
