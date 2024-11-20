import os
import spacy

# 加载英文模型
nlp = spacy.load("zh_core_web_sm")

def extract_named_entities_from_markdown(markdown_text):
    doc = nlp(markdown_text)
    named_entities = []
    for ent in doc.ents:
        named_entities.append((ent.text, ent.label_))
    return named_entities

def process_markdown_files(input_folder, output_folder):
    # 确保输出目录存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 遍历输入目录中的所有文件
    for filename in os.listdir(input_folder):
        if filename.endswith(".md"):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '_entities.md')

            # 读取Markdown文件
            with open(input_path, 'r', encoding='utf-8') as file:
                markdown_text = file.read()

            # 提取命名实体
            named_entities = extract_named_entities_from_markdown(markdown_text)

            # 将原始文本和命名实体写入新的Markdown文件
            with open(output_path, 'w', encoding='utf-8') as output_file:
                output_file.write(markdown_text + '\n\n')
                output_file.write('# Named Entities\n')
                for entity, entity_type in named_entities:
                    output_file.write(f'- {entity} ({entity_type})\n')

            print(f"Processed {input_path} to {output_path}")

# 指定输入和输出文件夹
input_folder = r"C:\Users\HP\Desktop\大创\output_step2_md"  # 输入Markdown文件夹路径
output_folder = r"C:\Users\HP\Desktop\大创\output_step3_md"  # 输出Markdown文件夹路径

process_markdown_files(input_folder, output_folder)