import os
import re

def split_markdown_by_custom_headers_to_array(markdown_folder):
    """按指定格式的标题（如 **1.2****评价目的**）切分文件夹中的Markdown文件，并将结果存储到数组中"""
    
    # 存储切分后的内容的数组
    split_content_array = []
    
    # 遍历文件夹中的所有 Markdown 文件
    for filename in os.listdir(markdown_folder):
        if filename.endswith('.md'):
            markdown_path = os.path.join(markdown_folder, filename)
            
            with open(markdown_path, 'r', encoding='utf-8') as md_file:
                content = md_file.read()
            
            # 使用正则表达式匹配标题（如 **1.2****评价目的**）
            header_pattern = re.compile(r'\*\*(\d+\.\d+)\*\*\*\*(.+?)\*\*')
            matches = list(header_pattern.finditer(content))
            
            if not matches:
                print(f"文件 {filename} 未找到任何匹配的标题，跳过")
                continue
            
            # 遍历匹配到的标题，将内容按标题切分并存入数组
            for i in range(len(matches)):
                start = matches[i].start()
                end = matches[i+1].start() if i + 1 < len(matches) else len(content)
                header_number = matches[i].group(1).strip()  # 获取标题编号
                header_text = matches[i].group(2).strip()    # 获取标题内容
                section_content = content[start:end].strip() # 切分对应的部分内容
                
                # 将切分出的部分以字典形式存储到数组中
                split_content_array.append({
                    'file': filename,
                    'header_number': header_number,
                    'header_text': header_text,
                    'content': section_content
                })

    return split_content_array

# 使用示例
markdown_folder = r"C:\Users\HP\Desktop\大创\output_step2\markdown"  # 输入的Markdown文件夹路径
split_content_array = split_markdown_by_custom_headers_to_array(markdown_folder)

# 打印结果，输出全部内容
for i, section in enumerate(split_content_array):
    print(f"Section {i+1}:")
    print(f"File: {section['file']}")
    print(f"Header Number: {section['header_number']}")
    print(f"Header Text: {section['header_text']}")
    print(f"Content:\n{section['content']}")
    print('-' * 40)
