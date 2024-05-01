#————————————————————测试openyxl————————————————————————#

import os
import openpyxl
import openpyxl.cell
import openpyxl.workbook
import openpyxl.worksheet
import openpyxl.worksheet.cell_range
import openpyxl.worksheet.merge
import openpyxl.worksheet.worksheet

cur_dir = os.path.dirname(os.path.realpath(__file__))
filedir = os.path.join(cur_dir, "xlsx_file", "test.xlsx")

# print(filedir)
workbook = openpyxl.load_workbook(filedir)

worksheet = workbook.active
merge_ranges = worksheet.merged_cells.ranges
# print(merge_ranges)
# for merge_range in merge_ranges:
#     print(merge_range)

for rows in worksheet.rows:
    for cell in rows:
        if type(cell) == openpyxl.cell.cell.Cell:
            print(cell.value, end=", ")
        elif type(cell) == openpyxl.cell.cell.MergedCell:
            for merge_range in merge_ranges:
                if cell.coordinate in merge_range:
                    # print(f"({merge_range.min_row}, {merge_range.min_col})", end=", ")
                    print(worksheet.cell(row = merge_range.min_row, column = merge_range.min_col).value, end=', ')
                    break
    print("")