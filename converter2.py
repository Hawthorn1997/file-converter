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
current_page = document.sections[0].first_page_header.tables[0]

for row in data:
    if row[headers.index(target_column)]:
        current_row = current_page.add_row().cells
        for cell_index, value in enumerate(row):
            current_row[cell_index].text = str(value)

# 保存填充后的Word文档
document.save('output.docx')
