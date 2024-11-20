import subprocess
import logging

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def convert_docx_to_markdown(docx_file, output_md_file):
    """
    使用 pandoc 将 .docx 文件转换为 .md 文件，仅处理文本内容。
    """
    try:
        subprocess.run([
            "pandoc", docx_file, "-o", output_md_file, "-t", "markdown", "--extract-media=unused"
        ], check=True)
        logging.info(f"Converted '{docx_file}' to '{output_md_file}'")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error converting '{docx_file}': {e}")

def main():
    docx_file = r"C:\Users\HP\Desktop\大创\output_step1\3-正宏商混-现评-正式版.docx" # 输入的 .docx 文件路径
    output_md_file = r"C:\Users\HP\Desktop\大创\output.md"  # 输出的 .md 文件路径

    convert_docx_to_markdown(docx_file, output_md_file)

if __name__ == "__main__":
    main()
