#这是使用LibreOffice的文件转化 doc to docx
import os
import subprocess

def convert_doc_to_docx(input_directory, output_directory):
    # 确保输出目录存在
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    # 遍历输入目录中的所有文件
    for filename in os.listdir(input_directory):
        if filename.endswith(".doc"):
            input_path = os.path.join(input_directory, filename)
            output_path = os.path.join(output_directory, os.path.splitext(filename)[0] + '.docx')
            
            # 构建 LibreOffice 命令
            command = ["soffice", "--headless", "--convert-to", "docx", "--outdir", output_directory, input_path,libreoffice_path]
            
            # 调用 LibreOffice 进行转换
            subprocess.run(command, check=True)
            print(f"Converted {input_path} to {output_path}")

# 指定包含 .doc 文件的目录
input_directory = "C:\\Users\\HP\\Desktop\\大创\\input"
# 指定输出 .docx 文件的目录
output_directory = "C:\\Users\\HP\\Desktop\\大创\\output_step1"
# LibreOffice 的安装路径
libreoffice_path = "C:\\Program Files\\LibreOffice\\program\\soffice.com"

convert_doc_to_docx(input_directory, output_directory)