from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 加载文档
doc = Document('./doc/ACM_CH01_MM01_TEST.docx')

# 遍历文档中的每个段落
for para in doc.paragraphs:
    # 检查段落对齐方式
    alignment = para.paragraph_format.alignment
    is_centered = alignment == WD_ALIGN_PARAGRAPH.CENTER
    print(f'这段文字{"是" if is_centered else "不是"}居中对齐: {para.text}')

    # 遍历段落中的每个运行
    for run in para.runs:
        # 获取字体对象
        font = run.font

        # 检查字体大小
        if font.size:
            print(f'字号: {font.size.pt}')

        # 检查字体颜色
        if font.color.rgb:
            print(f'字体颜色: RGB({font.color.rgb[0]}, {font.color.rgb[1]}, {font.color.rgb[2]})')
        else:
            print('字体颜色: 无')

        # 检查是否粗体、斜体或下划线
        print(f'粗体: {"是" if font.bold else "否"}')
        print(f'斜体: {"是" if font.italic else "否"}')
        print(f'下划线: {"是" if font.underline else "否"}')