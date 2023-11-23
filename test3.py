# # from docx import Document
# #
# # # doc = Document("temp_empty.docx")
# # #
# # # doc._body.clear_content()
# # #
# # # doc.add_heading("段落1", level=3)
# # #
# # # doc.add_heading("段落2", level=4)
# # #
# # # doc.add_heading("段落3", level=4)
# # #
# # # doc.add_heading("段落4", level=3)
# # #
# # # doc.add_heading("段落5", level=4)
# # #
# # # doc.add_heading("段落6", level=3)
# # #
# # # doc.save("result.docx")
# #
# #
# # # def create_all_style():
# # #     title_style = []
# # #     for i in range(9):
# # #         title_style.append(str(i+1)+'级标题')
# # #     return title_style
# # #
# # # print(create_all_style())
# #
# #
# # import re
# #
# # def remove_prefix(s):
# #     return re.sub(r'^(\d+\.*)+', '', s)
# #
# # # 测试
# # s = '123 abc'
# # result = remove_prefix(s)
# # print(result)
#
# from docx import Document
#
# def delete_content_before_line(doc_path, line_number):
#     doc = Document(doc_path)
#     for paragraph in doc.paragraphs:
#         if len(paragraph.text.split(' ')) >= line_number:
#             paragraph.clear()
#     doc.save('output.docx')
#
# # 使用示例
# # delete_content_before_line('input.docx', 3)
# doc = Document("input.docx")
# tables = doc.tables[1]
# print(tables.cell(0,0).text)

def remove_last_colon(s):
    return s.rstrip(':')

string = "example: string:"
result = remove_last_colon(string)
print(result)
