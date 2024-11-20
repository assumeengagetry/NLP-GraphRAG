from docx import Document
import os

def split_by_heading(input_file, output_folder):
    # 读取文档
    doc = Document(input_file)
    current_heading = None
    content = []

    # 创建输出文件夹
    os.makedirs(output_folder, exist_ok=True)
    
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            # 如果遇到新标题，保存前面的内容（按200字符切分）
            if current_heading:
                save_to_file(current_heading, content, output_folder)
            current_heading = para.text
            content = []
        else:
            content.append(para.text)

    # 保存最后一段内容
    if current_heading:
        save_to_file(current_heading, content, output_folder)

def save_to_file(heading, content, output_folder):
    # 合并内容为一个字符串
    text = "\n".join(content)
    parts = split_text(text, 200)
    
    for i, part in enumerate(parts, 1):
        # 创建文件名，去除无效字符
        safe_heading = heading.replace('/', '-').strip()
        filename = os.path.join(output_folder, f"{safe_heading}_{i}.txt")
        
        # 写入文件
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(part)

def split_text(text, limit):
    parts = []
    while len(text) > limit:
        # 找到从尾部开始的最近的句号或分号的位置
        cut_index_period = text.rfind('。', 0, limit)
        cut_index_semicolon = text.rfind('；', 0, limit)
        
        # 选择最靠近尾部的句号或分号
        cut_index = max(cut_index_period, cut_index_semicolon)
        if cut_index == -1:  # 如果没有找到句号或分号
            cut_index = limit
        
        # 添加切分部分并更新剩余内容
        parts.append(text[:cut_index + 1].strip())
        text = text[cut_index + 1:].strip()
    
    # 添加剩余部分
    if text:
        parts.append(text)
    return parts

# 示例调用
input_file = r"C:\Users\HP\Desktop\大创\output_step1\06睿恒化工--控评--修改版2.docx"
output_folder = r"C:\Users\HP\Desktop\大创\output_step3"
split_by_heading(input_file, output_folder)
