from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# 打开Word文档
doc = Document("temp_empty.docx")

# 选择要修改的段落
paragraph = doc.paragraphs[0]

# 修改段落样式
paragraph.style = 'Heading 1'
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
# paragraph.font.size = Pt(24)

# 保存修改后的文档
doc.save("modified_example.docx")