import os
import re
import spacy

# 加载spaCy的NER模型
nlp = spacy.load("zh_core_web_sm")

def split_markdown(file_path, output_folder):
    # 读取Markdown文件内容
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # 正则表达式匹配Markdown标题
    chapters = re.split(r'(##\s+[^#]+)', content)[1:]  # 跳过第一个空字符串

    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 保存切分后的章节为新的Markdown文件
    for i, chapter in enumerate(chapters):
        title = re.search(r'^##\s+(.+)', chapter)
        if title:
            title = title.group(1)
            chapter_content = chapter.replace(title + '\n', '', 1).strip()
            chapter_file_path = os.path.join(output_folder, f"{os.path.splitext(os.path.basename(file_path))[0]}_chapter_{i+1}.md")
            with open(chapter_file_path, 'w', encoding='utf-8') as new_file:
                new_file.write(f"## {title}\n\n{chapter_content}")
    return len(chapters)  # 返回切分后的文件数量

def ner_on_text(file_path, output_folder):
    # 读取文件内容
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()

    # 使用spaCy进行NER
    doc = nlp(text)
    entities = [(ent.text, ent.label_) for ent in doc.ents]
    
    # 保存NER结果
    ner_file_path = os.path.splitext(file_path)[0] + '_entities.txt'
    with open(ner_file_path, 'w', encoding='utf-8') as ner_file:
        for entity in entities:
            ner_file.write(f"{entity[0]} ({entity[1]})\n")
    os.replace(ner_file_path, os.path.join(output_folder, os.path.basename(ner_file_path)))  # 移动文件到NER输出文件夹

def process_markdown_files(input_folder, markdown_output_folder, ner_output_folder):
    # 确保输出文件夹存在
    if not os.path.exists(markdown_output_folder):
        os.makedirs(markdown_output_folder)
    if not os.path.exists(ner_output_folder):
        os.makedirs(ner_output_folder)

    # 遍历输入文件夹中的所有Markdown文件
    for filename in os.listdir(input_folder):
        if filename.endswith(".md"):
            file_path = os.path.join(input_folder, filename)
            chapters_count = split_markdown(file_path, markdown_output_folder)
            for i in range(1, chapters_count + 1):
                chapter_file_path = os.path.join(markdown_output_folder, f"{os.path.splitext(filename)[0]}_chapter_{i}.md")
                ner_on_text(chapter_file_path, ner_output_folder)

# 设置输入和输出文件夹路径
input_folder = r"C:\Users\HP\Desktop\大创\output_step2_md"
markdown_output_folder = r"C:\Users\HP\Desktop\大创\output_step3"
ner_output_folder = r"C:\Users\HP\Desktop\大创\output_step4"

# 处理Markdown文件
process_markdown_files(input_folder, markdown_output_folder, ner_output_folder)