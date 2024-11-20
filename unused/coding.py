import os

def get_file_type(file_path):
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()  # 转为小写以进行不区分大小写的比较

    if ext in ['.xls', '.xlsx']:
        return 'Excel'
    elif ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']:
        return 'Image'
    elif ext in ['.doc', '.docx']:
        return 'DOCX'
    else:
        return 'Unknown'

# 示例用法
file_path = 'example.xlsx'
file_type = get_file_type(file_path)
print(f'文件类型: {file_type}')
