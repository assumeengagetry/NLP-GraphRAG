import os
import re

def split_markdown_by_custom_headers(markdown_path, output_folder):
    """按指定格式的标题（如 **1.2**）切分Markdown文件，并保存到指定文件夹中"""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    with open(markdown_path, 'r', encoding='utf-8') as md_file:
        content = md_file.read()
    
    # 匹配标题的正则表达式，如 **1.2**
    header_pattern = re.compile(r'\*\*(\d+\.\d+)\*\*', re.MULTILINE)
    matches = list(header_pattern.finditer(content))
    
    if not matches:
        print("未找到任何标题，无法分割文件")
        return
    
    # 初始化变量
    current_pos = 0
    # 遍历匹配到的标题，将内容按标题切分
    for i, match in enumerate(matches):
        start = match.start()
        header_number = match.group(1).strip()
        
        # 查找下一个匹配的标题位置
        if i < len(matches) - 1:
            next_start = matches[i + 1].start()
        else:
            next_start = len(content)
        
        section_content = content[start:next_start].strip()
        
        # 生成一个合法的文件名
        filename = f"{header_number}.md"
        file_path = os.path.join(output_folder, filename)
        
        # 保存切分后的部分
        with open(file_path, 'w', encoding='utf-8') as output_file:
            output_file.write(section_content)
        
        print(f'章节 {header_number} 已保存到 {file_path}')
        current_pos = next_start
    
    print(f"Markdown文件已成功按标题切分并保存到 {output_folder}")

# 使用示例
markdown_path = r""  # 输入的Markdown文件路径
output_folder = r""  # 输出文件夹路径

split_markdown_by_custom_headers(markdown_path, output_folder)
