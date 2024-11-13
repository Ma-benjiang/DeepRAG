import openpyxl
import pandas as pd
from openpyxl import load_workbook

class ExcelParser:
    def __call__(self, file_path) -> str:
        wb = load_workbook(file_path)
        tb_chunks = []

        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            merged_cells_dict = {}

            # 遍历所有合并区域，记录每个区域的左上角单元格的值
            for range_ in sheet.merged_cells.ranges:
                min_col, min_row, max_col, max_row = range_.min_col, range_.min_row, range_.max_col, range_.max_row
                top_left_cell_value = sheet.cell(row=min_row, column=min_col).value
                for row in range(min_row, max_row + 1):
                    for col in range(min_col, max_col + 1):
                        cell = sheet.cell(row=row, column=col)
                        merged_cells_dict[cell.coordinate] = top_left_cell_value

            tb = f"<table><caption>{sheet_name}</caption>"

            # 获取标题行
            header_row = list(sheet.iter_rows(min_row=1, max_row=1, values_only=False))[0]
            tb += "<tr>"
            for header_cell in header_row:
                tb += f"<th>{header_cell.value}</th>"
            tb += "</tr>"

            rows = list(sheet.iter_rows(min_row=2, values_only=False))
            for r in rows:
                tb += "<tr>"
                for c in r:
                    # 使用合并单元格字典检查是否需要替换值
                    cell_value = merged_cells_dict.get(c.coordinate, c.value)
                    tb += f"<td>{cell_value if cell_value is not None else ''}</td>"
                tb += "</tr>"

            tb += "</table>"
            tb_chunks.append(tb)

        return ''.join(tb_chunks)
if __name__ == "__main__":
    
    psr = ExcelParser()
    text = psr("test.xlsx")
    print(text)