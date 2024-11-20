import os
import requests
import json

# 定义API信息
knowledge_base_api_url = "你的知识库API URL"
knowledge_base_api_key = "你的知识库API key"
deepseek_api_url = "https://api.deepseek.com"  # DeepSeek API的基础URL
deepseek_api_key = "sk-810c50286291463f963c6af4e776ec44"

# 设置API请求头
headers = {
    "Authorization": f"Bearer {deepseek_api_key}",
    "Content-Type": "application/json"
}

# 读取文件夹内的结构体文件
def read_structures_from_folder(folder_path):
    structures = []
    for filename in os.listdir(folder_path):
        if filename.endswith(".json"):  # 假设结构体以JSON格式存储
            with open(os.path.join(folder_path, filename), 'r') as file:
                structures.append(json.load(file))
    return structures

# 调用DeepSeek API进行比对
def compare_with_deepseek(structures, knowledge_base_content):
    comparisons = []
    for structure in structures:
        # 构建DeepSeek API请求体
        payload = {
            "model": "deepseek-chat",  # 或者其他DeepSeek支持的模型
            "messages": [
                {"role": "system", "content": "Compare the following structure with the knowledge base."},
                {"role": "user", "content": json.dumps(structure)},
                {"role": "user", "content": knowledge_base_content}
            ],
            "stream": False
        }
        # 发送POST请求
        response = requests.post(f"{deepseek_api_url}/chat/completions", headers=headers, json=payload)
        if response.status_code == 200:
            comparison_result = response.json()
            comparisons.append(comparison_result)
        else:
            comparisons.append({"error": "Failed to compare"})
    return comparisons

# 主函数
def main():
    folder_path = "path_to_your_structures_folder"  # 结构体文件所在的文件夹路径
    structures = read_structures_from_folder(folder_path)
    
    # 假设你已经有了知识库的内容
    knowledge_base_content = "your_knowledge_base_content"
    
    # 进行比对
    comparison_results = compare_with_deepseek(structures, knowledge_base_content)
    
    # 输出比对结果
    for result in comparison_results:
        if "error" in result:
            print(result["error"])
        else:
            print(result)

if __name__ == "__main__":
    main()