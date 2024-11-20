import os
import re

def split_markdown_by_double_asterisk_headers(markdown_path, output_folder):
    """将Markdown文件按以**符号表示的小标题切分，并将每部分保存到指定文件夹中"""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    with open(markdown_path, 'r', encoding='utf-8') as md_file:
        content = md_file.read()
    
    # 正则表达式匹配以 ** 开头的标题
    header_pattern = re.compile(r'^\*\*\s+(?P<header>.+)$', re.MULTILINE)
    matches = list(header_pattern.finditer(content))
    
    if not matches:
        print("未找到任何标题，无法分割文件")
        return
    
    # 遍历匹配到的标题，将内容按标题切分
    for i in range(len(matches)):
        start = matches[i].start()
        end = matches[i+1].start() if i + 1 < len(matches) else len(content)
        header = matches[i].group('header').strip()
        section_content = content[start:end].strip()
        
        # 生成一个合法的文件名
        filename = f"{i+1}_{re.sub(r'[^a-zA-Z0-9_-]', '_', header)}.md"
        file_path = os.path.join(output_folder, filename)
        
        # 保存切分后的部分
        with open(file_path, 'w', encoding='utf-8') as output_file:
            output_file.write(section_content)

    print(f"Markdown文件已成功按小标题切分并保存到 {output_folder}")

# 使用示例
markdown_path = r"C:\Users\HP\Desktop\大创\output_step2\markdown\成都昇科智能 专篇（正式版).md"  # 输入的Markdown文件路径
output_folder = r"C:\Users\HP\Desktop\大创\output_step3"  # 输出文件夹路径

split_markdown_by_double_asterisk_headers(markdown_path, output_folder)
