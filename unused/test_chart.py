from docx import Document
import pandas as pd
import os

def extract_tables(docx_file):
    doc = Document(docx_file)
    tables_data = []
    for table in doc.tables:
        table_content = []
        for row in table.rows:
            row_content = [cell.text for cell in row.cells]
            table_content.append(row_content)
        tables_data.append(table_content)
    return tables_data

def save_tables_to_excel(tables_data, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for i, table in enumerate(tables_data):
            df = pd.DataFrame(table)
            df.to_excel(writer, sheet_name=f'Table {i+1}', index=False)

def excel_to_markdown(excel_file):
    xls = pd.ExcelFile(excel_file)
    markdown_content = ""
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        markdown_table = df.to_markdown(index=False)
        markdown_content += f"## {sheet_name}\n\n{markdown_table}\n\n"
    return markdown_content

def main(docx_file, excel_file, markdown_file):
    # 提取表格
    tables_data = extract_tables(docx_file)
    
    # 保存为Excel文件
    save_tables_to_excel(tables_data, excel_file)
    
    # 将Excel转换为Markdown
    markdown_content = excel_to_markdown(excel_file)
    
    # 保存Markdown内容
    with open(markdown_file, 'w', encoding='utf-8') as file:
        file.write(markdown_content)
    print(f"Markdown文件已保存：{markdown_file}")

if __name__ == "__main__":
    docx_file_path = "C:\\Users\\HP\\Desktop\\大创\\ke_lun_yuan_cai_liao.docx"  # docx文件路径
    excel_file_path ="C:\\Users\\HP\\Desktop\\chart_test.xlsx" # 输出的Excel文件路径
    markdown_file_path ="C:\\Users\\HP\\Desktop\\chart_test.md"  # 输出的Markdown文件路径
    main(docx_file_path, excel_file_path, markdown_file_path)