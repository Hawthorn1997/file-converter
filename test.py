from docx import Document

# 打开原始的 Word 模板
template_path = 'template.docx'
template_doc = Document(template_path)

# 创建新的 Word 文档
new_doc = Document()

# 复制原始模板到新文档
for element in template_doc.element.body:
    new_doc.element.body.append(element)

# 保存新文档
new_doc.save('output.docx')



def copy(template_doc, new_doc):
    for element in template_doc.element.body:
        new_doc.element.body.append(element)