from openpyxl import load_workbook
from docx import Document

# 加载Excel文件
workbook = load_workbook('input.xlsx')
sheet = workbook.active

# 加载Word模板
document = Document('template.docx')

# 获取Excel表格数据
data = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    data.append(row)

# 获取Excel表头
headers = [cell.value for cell in sheet[1]]

# 根据管道编号/单线号填充Word模板
target_column = '管道编号/单线号'
row_index = 0
page_index = 0

for row in data:
    if row[headers.index(target_column)]:
        # 将数据填入Word模板的对应位置
        table = document.tables[page_index]
        current_row = table.rows[row_index]
        for cell_index, value in enumerate(row):
            current_row.cells[cell_index].text = str(value)

        row_index += 1

        # 每页最多填写13行数据
        if row_index >= 13:
            row_index = 0
            page_index += 1
            if page_index >= len(document.tables):
                document.add_page_break()
                document.add_table(rows=14, cols=len(headers))

# 保存填充后的Word文档
document.save('output.docx')