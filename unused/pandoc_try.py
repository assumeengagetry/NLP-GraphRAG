import os
import subprocess

def convert_docx_to_md(input_folder, output_folder):
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 遍历输入文件夹中的所有文件
    for filename in os.listdir(input_folder):
        if filename.endswith(".docx"):
            # 设置源文件和目标文件的路径
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.md')
            
            # 调用 pandoc 进行转换
            subprocess.run(["pandoc", input_path, "-o", output_path], check=True)

            print(f"Converted '{input_path}' to '{output_path}'")

# 设置输入和输出文件夹路径
input_folder = r"C:\Users\HP\Desktop\大创\output_step1" # 替换为实际的输入文件夹路径
output_folder = r"C:\Users\HP\Desktop\大创\output_step2_md"  # 替换为实际的输出文件夹路径

# 执行转换
convert_docx_to_md(input_folder, output_folder)