import os
import re

def split_markdown_by_custom_headers(markdown_path, output_folder):
    """按指定格式的标题（如 **1** **.** **3** 评价目的）切分Markdown文件，并保存到指定文件夹中"""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    try:
        with open(markdown_path, 'r', encoding='utf-8') as md_file:
            content = md_file.read()
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return
    
    # 更新正则表达式匹配标题（如 **1** **.** **3** 评价目的）
    header_pattern = re.compile(r'\*\*(\d+)\*\*\s*\*\*\.\*\*\s*\*\*(.+?)\*\*')
    matches = list(header_pattern.finditer(content))
    
    if not matches:
        print("未找到任何标题，无法分割文件")
        return
    
    # 遍历匹配到的标题，将内容按标题切分
    for i in range(len(matches)):
        start = matches[i].start()
        end = matches[i+1].start() if i + 1 < len(matches) else len(content)
        header_number = matches[i].group(1).strip()
        header_text = matches[i].group(2).strip()
        section_content = content[start:end].strip()
        
        # 检查内容是否为空
        if not section_content:
            print(f"警告：分割内容为空，跳过 {header_number} {header_text}")
            continue
        
        # 生成一个合法的文件名
        filename = f"{i+1}_{header_number}_{re.sub(r'[^a-zA-Z0-9_-]', '_', header_text)}.md"
        file_path = os.path.join(output_folder, filename)
        
        try:
            # 保存切分后的部分
            with open(file_path, 'w', encoding='utf-8') as output_file:
                output_file.write(section_content)
        except Exception as e:
            print(f"保存文件时出错: {e}")

    print(f"Markdown文件已成功按标题切分并保存到 {output_folder}")

# 使用示例
markdown_path = r"C:\Users\HP\Desktop\大创\output_step2\06睿恒化工--控评--修改版2\markdown\06睿恒化工--控评--修改版2.md"  # 输入的Markdown文件路径
output_folder = r"C:\Users\HP\Desktop\大创\output_step3"  # 输出文件夹路径

split_markdown_by_custom_headers(markdown_path, output_folder)
