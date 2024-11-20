from docx import Document

def split_text_by_length(text, max_length=150):
    """
    将一段文本切分为多个子段，每段小于等于 max_length 个字符。
    """
    return [text[i:i + max_length] for i in range(0, len(text), max_length)]

def split_docx_by_paragraph(input_file, max_length=150):
    # 打开 .docx 文档
    doc = Document(input_file)
    
    # 存储最终分割后的段落数组
    result_paragraphs = []

    # 遍历每个段落并处理
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:  # 忽略空白段落
            # 如果段落长度超过 max_length，则切分为多个子段落
            if len(text) > max_length:
                result_paragraphs.extend(split_text_by_length(text, max_length))
            else:
                result_paragraphs.append(text)
    
    return result_paragraphs

if __name__ == "__main__":
    input_file = r"C:\Users\HP\Desktop\大创\output_step1\06睿恒化工--控评--修改版2.docx"  # 你的 .docx 文件路径

    # 获取切分后的段落数组
    paragraph_array = split_docx_by_paragraph(input_file, max_length=150)

    # 打印数组中的段落
    for i, paragraph in enumerate(paragraph_array, 1):
        print(f"Paragraph {i}: {paragraph}")
    
    # 进一步操作该数组，例如：
    # print(paragraph_array)
