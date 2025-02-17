import re
import shutil
from pathlib import Path

def update_markdown_content(md_file_path, image_name, new_content):
    try:
        md_file_path = Path(md_file_path)
        
        # 读取Markdown文件
        with open(md_file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        
        # 创建备份
        backup_path = md_file_path.with_suffix(md_file_path.suffix + '.backup')
        shutil.copy2(md_file_path, backup_path)
        
        # 构建匹配模式 - 适配 ![](images/filename.jpg) 格式
        pattern = f"!\\[\\]\\(images/{re.escape(image_name)}\\)"
        if not re.search(pattern, content):
            raise ValueError(f"Image reference '{image_name}' not found in the Markdown file")
            
        updated_content = re.sub(pattern, new_content, content)
        
        # 写入更新后的内容
        with open(md_file_path, 'w', encoding='utf-8') as file:
            file.write(updated_content)
            
        print(f"Successfully updated {md_file_path}")
        print(f"Backup created at {backup_path}")
        
    except FileNotFoundError:
        print(f"Error: File not found - {md_file_path}")
    except Exception as e:
        print(f"Error occurred: {str(e)}")

if __name__ == "__main__":
    # 示例使用:
    md_file_path = r"C:\Users\HP\Desktop\test3.md"
    # 只需要文件名部分
    image_name = "7706d6ed5ddd1cef4b715ee682cf2eb945cab3d11fb97dbe7a869b16f0e7fda3.jpg"
    new_content = "这里是图片的描述文本"
    
    update_markdown_content(md_file_path, image_name, new_content)
