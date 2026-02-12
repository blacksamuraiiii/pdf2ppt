---
name: skill-pdf2ppt
description: 将 PDF 文档转换为可编辑的 PowerPoint (PPTX) 演示文稿。支持智能版面分析、去水印及表格还原。
---

# PDF 转 PPT 转换技能

## 触发条件
当用户有以下需求时加载此 Skill：
1. "把这个 PDF 转成 PPT"
2. "Convert this document to slides"
3. "将会议记录 PDF 变成演示文稿"

## 前置检查
在执行转换前，请确认：
1. **PDF路径**：用户是否提供了目标 PDF 文件？
2. **API Token**：环境中是否已配置 MinerU Token？（如果未配置，需询问用户）
3. **输出偏好**：用户是否指定了 16:9 或 4:3比例？（默认 16:9）

## 执行流程

### 1. 准备环境
确保 Python 依赖已安装：`pip install -r skill-pdf2ppt/requirements.txt`

### 2. 执行转换
使用 `scripts/main.py` 执行转换任务。

#### 基本用法
```python
# 伪代码示例
command = f"python skill-pdf2ppt/scripts/main.py --pdf_path '{pdf_path}'"
```

#### 完整参数示例
```python
# 如果用户指定了 Token 和去水印
command = f"python skill-pdf2ppt/scripts/main.py --pdf_path '{pdf_path}' --token '{user_token}' --remove_watermark"
```

### 3. 结果验证
- 检查脚本输出是否包含 "Conversion Success" 或类似成功标志。
- 确认输出路径下是否存在 `.pptx` 文件。

## 常见错误处理
- **Token 无效**：提示用户检查 MinerU Token。
- **解析超时**：MinerU 服务可能繁忙，建议稍后重试。
