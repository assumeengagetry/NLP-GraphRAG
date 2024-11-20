import os
import spacy

# 加载spaCy的NER模型
nlp = spacy.load("zh_core_web_sm")

def perform_ner_on_folder(input_folder, output_folder):
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 遍历输入文件夹中的所有Markdown文件
    for filename in os.listdir(input_folder):
        if filename.endswith(".md"):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '_ner.txt')
            
            # 读取Markdown文件内容
            with open(input_path, 'r', encoding='utf-8') as file:
                content = file.read()

            # 使用spaCy进行NER
            doc = nlp(content)
            entities = [(ent.text, ent.label_) for ent in doc.ents]

            # 保存NER结果
            with open(output_path, 'w', encoding='utf-8') as ner_file:
                for entity in entities:
                    ner_file.write(f"{entity[0]} ({entity[1]})\n")
            print(f"NER results saved to {output_path}")

# 设置输入和输出文件夹路径
input_folder =r"C:\Users\HP\Desktop\大创\output_step2_md"  # 替换为你的输入文件夹路径
output_folder = r"C:\Users\HP\Desktop\大创\output_step3"  # 替换为你的输出文件夹路径

# 执行NER并保存结果
perform_ner_on_folder(input_folder, output_folder)