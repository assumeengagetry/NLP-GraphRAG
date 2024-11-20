import os
import pdfplumber
import pandas as pd

def extract_text_and_tables_from_pdf(pdf_file, text_output_folder, excel_output_folder):
    with pdfplumber.open(pdf_file) as pdf:
        text_content = ""
        tables = []
        
        for page in pdf.pages:
            text_content += page.extract_text() + "\n"
            tables.extend(page.extract_tables())

        # 保存文本内容到Markdown文件
        markdown_file = os.path.splitext(os.path.basename(pdf_file))[0] + '.md'
        with open(os.path.join(text_output_folder, markdown_file), "w", encoding="utf-8") as md_file:
            md_file.write(text_content)
        print(f"Text extracted to {markdown_file}")

        # 保存表格数据到Excel文件
        if tables:
            excel_file = os.path.splitext(os.path.basename(pdf_file))[0] + '.xlsx'
            df_list = [pd.DataFrame(table) for table in tables]
            with pd.ExcelWriter(os.path.join(excel_output_folder, excel_file), engine='openpyxl') as writer:
                for i, df in enumerate(df_list):
                    df.to_excel(writer, sheet_name=f'Table {i+1}', index=False)
            print(f"Tables extracted to {excel_file}")

def excel_to_markdown(excel_file, markdown_file):
    if os.path.exists(excel_file):
        xls = pd.ExcelFile(excel_file)
        markdown_content = ""
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            markdown_table = df.to_markdown(index=False)
            markdown_content += f"## {sheet_name}\n\n{markdown_table}\n\n"

        with open(markdown_file, "a", encoding="utf-8") as md_file:
            md_file.write(markdown_content)
        print(f"Markdown tables added to {markdown_file}")
    else:
        print(f"Excel file not found: {excel_file}")

def main(pdf_folder, text_output_folder, excel_output_folder):
    if not os.path.exists(excel_output_folder):
        os.makedirs(excel_output_folder)
    if not os.path.exists(text_output_folder):
        os.makedirs(text_output_folder)

    for filename in os.listdir(pdf_folder):
        if filename.endswith(".pdf"):
            pdf_file_path = os.path.join(pdf_folder, filename)
            markdown_file = os.path.splitext(filename)[0] + '.md'
            excel_file = os.path.splitext(filename)[0] + '.xlsx'
            extract_text_and_tables_from_pdf(pdf_file_path, text_output_folder, excel_output_folder)
            excel_file_path = os.path.join(excel_output_folder, excel_file)
            markdown_file_path = os.path.join(text_output_folder, markdown_file)
            if os.path.exists(excel_file_path):
                excel_to_markdown(excel_file_path, markdown_file_path)
            else:
                print(f"Excel file not found: {excel_file_path}")

if __name__ == "__main__":
    pdf_folder = r"C:\Users\HP\Desktop\大创\output_step1"  # 输入PDF文件夹路径
    text_output_folder = r"C:\Users\HP\Desktop\大创\output_step2_md"  # 输出Markdown文件夹路径
    excel_output_folder = r"C:\Users\HP\Desktop\大创\chart_excel"  # 输出Excel文件夹路径

    main(pdf_folder, text_output_folder, excel_output_folder)