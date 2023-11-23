import docx

def extract_paragraph_styles(docx_file):
    # 打开Word文档
    doc = docx.Document(docx_file)

    # 存储段落样式的字典
    paragraph_styles = {}

    # 遍历文档中的段落
    for paragraph in doc.paragraphs:
        # 获取段落样式
        style = paragraph.style.name

        print(style)

        # 如果样式不在字典中，则添加到字典中
        if style not in paragraph_styles:
            paragraph_styles[style] = []

        # 将段落添加到对应样式的列表中
        paragraph_styles[style].append(paragraph.text)

    return paragraph_styles

# 使用示例
docx_file = 'input.docx'
result = extract_paragraph_styles(docx_file)
print(result)
