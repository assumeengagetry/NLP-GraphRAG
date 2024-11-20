from docx import Document
import os

def extract_images(docx_file, image_folder):
    doc = Document(docx_file)
    if not os.path.exists(image_folder):
        os.makedirs(image_folder)
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            image = rel.target_part.blob
            # 获取实际的文件名，去掉'word/media/'前缀
            image_filename = os.path.basename(rel.target_ref)
            image_path = os.path.join(image_folder, image_filename)
            with open(image_path, 'wb') as img_file:
                img_file.write(image)
    print(f"Images extracted to {image_folder}")

# 使用函数
docx_file_path = r'C:\\Users\\HP\\Desktop\\大创\\ke_lun_yuan_cai_liao.docx'  # docx文件路径
image_folder = r'C:\\Users\\HP\\Desktop\\大创\\pictures'  # 提取的图片将保存在此文件夹
extract_images(docx_file_path, image_folder)
def create_image_markdown(image_folder, markdown_file):
    image_filenames = [f for f in os.listdir(image_folder) if os.path.isfile(os.path.join(image_folder, f))]
    markdown_content = ""
    for image_filename in image_filenames:
        image_path = os.path.join(image_folder, image_filename)
        markdown_content += f"![]({image_path})\n\n"
    with open(markdown_file, 'w', encoding='utf-8') as md_file:
        md_file.write(markdown_content)
    print(f"Markdown file with images created at {markdown_file}")

# 使用函数
markdown_file_path = "C:\\Users\\HP\\Desktop\\大创\\test_pictures.md"  # 输出的Markdown文件路径
create_image_markdown(image_folder, markdown_file_path)