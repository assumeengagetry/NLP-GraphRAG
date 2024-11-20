#这是使用python的某些库进行的doc to docx文件转化
import comtypes.client

def doc_to_docx(doc_path, docx_path):
    # 启动Word应用程序
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    
    try:
        # 打开.doc文件
        doc = word.Documents.Open(doc_path)
        # 保存为.docx格式
        doc.SaveAs(docx_path, FileFormat=16)  # 16代表wdFormatXMLDocument
        print(f"文件已转换并保存为：{docx_path}")
    except Exception as e:
        print(f"转换过程中发生错误：{e}")
    finally:
        # 关闭文档
        doc.Close()
        # 关闭Word应用程序
        word.Quit()

# 使用函数
doc_path = "C:\\Users\\HP\\Desktop\\大创\\ren_an_yao_ye.doc" # .doc文件路径
docx_path = "C:\\Users\\HP\\Desktop\\大创\\ren_an1.docx"  # 输出的.docx文件路径
doc_to_docx(doc_path, docx_path)