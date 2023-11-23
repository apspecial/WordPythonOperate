from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt

# 创建一个新的文档对象
doc = Document()

# 添加一个自定义标题
title = '这是一个自定义标题'
paragraph = doc.add_paragraph(title)

# 设置标题的样式
run = paragraph.runs[0]
font = run.font
font.name = 'Arial'
font.size = Pt(24)
font.bold = True
font.italic = False

# 设置标题的对齐方式
paragraph_format = paragraph.paragraph_format
paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

# 保存文档
doc.save('custom_title.docx')