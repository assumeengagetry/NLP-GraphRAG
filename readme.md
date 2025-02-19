以下是该API的完整使用指南：

---

### **API 使用流程**

#### **1. 启动服务**
```bash
python app.py
# 成功启动后输出：
# [系统启动] 上传目录已创建: {绝对路径}/uploads
# * Running on http://0.0.0.0:5000
```

---

### **2. 接口说明**

| 接口方法 | 路径 | 功能 | Content-Type | 参数 |
|---------|------|-----|--------------|------|
| POST    | /convert | 上传DOCX文件 | multipart/form-data | 文件字段名：file |
| GET     | /result/{task_id} | 下载结果文件 | - | URL路径参数：task_id |

---

### **3. 完整调用示例**

#### **步骤1：上传DOCX文件**
```bash
curl -X POST \
  http://localhost:5000/convert \
  -H 'Content-Type: multipart/form-data' \
  -F 'file=@/path/to/your.docx'
```

**成功响应**：
```json
{
    "task_id": "550e8400-e29b-41d4-a716-446655440000",
    "download_url": "/result/550e8400-e29b-41d4-a716-446655440000"
}
```

**错误处理**：
| 状态码 | 错误信息 | 解决方案 |
|--------|----------|----------|
| 400    | 未检测到文件上传 | 检查是否添加了文件字段 |
| 400    | 仅支持DOCX格式 | 确保文件扩展名为.docx |
| 500    | 处理失败: {错误详情} | 根据错误信息检查服务日志 |

---

#### **步骤2：下载结果文件**
```bash
curl -OJ http://localhost:5000/result/550e8400-e29b-41d4-a716-446655440000
```

**参数说明**：
- `-O`：保留服务器返回的文件名
- `-J`：使用服务器建议的文件名

**响应结果**：
- 自动下载生成好的Markdown文件（如：`test3.md`）

**错误处理**：
| 状态码 | 错误信息 | 解决方案 |
|--------|----------|----------|
| 404    | 无效任务ID | 检查task_id是否正确 |
| 404    | 文件不存在 | 等待处理完成后再下载 |
| 500    | 下载失败: {错误详情} | 检查文件权限或路径 |

---

### **4. 使用Postman测试**

#### **POST请求配置**
![Postman上传配置](https://assets.postman.com/postman-docs/v10/upload-config.png)
1. Method选择 **POST**
2. URL填写 `http://localhost:5000/convert`
3. Body选择 **form-data**
4. Key填写 `file` → 类型选择 **File**
5. Value点击选择本地.docx文件

#### **GET请求配置**
![Postman下载配置](https://assets.postman.com/postman-docs/v10/download-config.png)
1. Method选择 **GET**
2. URL填写 `http://localhost:5000/result/{task_id}`
3. 点击Send后会自动下载文件

---

### **5. 服务日志解读**

| 日志级别 | 示例 | 说明 |
|----------|------|------|
| INFO | [任务启动] 新建任务目录 | 新任务开始处理 |
| DEBUG | [转换日志] 开始转换 DOCX 到 PDF | 转换进度跟踪 |
| ERROR | [图片错误] 处理失败 figure1.jpg: API超时 | 单个图片处理失败不影响整体流程 |
| WARNING | [分析跳过] 未引用图片: temp.png | 跳过未被Markdown引用的图片 |

---

### **6. 开发注意事项**

1. **文件大小限制**：最大支持100MB文件
2. **中文路径处理**：已配置中文本地化环境
3. **安全建议**：
   - 生产环境需添加身份验证
   - 建议配置HTTPS
   - 定期清理uploads目录
4. **性能优化**：
   ```python
   # 调整Flask配置
   app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 增大到200MB
   app.config['UPLOAD_FOLDER'] = '/data/uploads'  # 使用独立存储分区
   ```

---

### **7. 常见问题排查**

**Q1：长时间无响应**
```bash
# 检查服务是否运行
ps aux | grep app.py

# 查看端口监听
netstat -tuln | grep 5000
```

**Q2：文件权限问题**
```bash
# 给上传目录赋权
chmod -R 777 /path/to/uploads
```

**Q3：中文文件名乱码**
```python
# 在代码开头添加编码声明
import sys
sys.setdefaultencoding('utf-8')
```

---

通过以上指南即可完成DOCX到Markdown的转换流程。建议首次使用时先用小于10MB的文件进行测试，待流程验证通过后再处理大文件。
