# PPT 翻译工具

这是一个使用DeepSeek API的PPT翻译工具，支持资深专家级翻译，自带计算机类、操作系统类等等专业词库。工具可以导入本地PPT进行翻译，并在不改变原版PPT格式的情况下输出翻译后的PPT。

## 功能特点

- 专业领域词汇支持（计算机科学、操作系统等）
- 保留原PPT格式和布局
- 支持批量翻译
- 可自定义翻译指令

## 安装

1. 克隆此仓库
2. 安装依赖包：
   ```
   pip install -r requirements.txt
   ```
3. 复制`.env.example`文件为`.env`并添加你的DeepSeek API密钥

## 使用方法

```
python translator.py --input your_ppt_file.pptx --output translated_file.pptx --source en --target zh
```

### 参数说明

- `--input`: 输入PPT文件路径
- `--output`: 输出PPT文件路径
- `--source`: 源语言代码（默认为'en'）
- `--target`: 目标语言代码（默认为'zh'）
- `--model`: 使用的模型名称（默认为'deepseek-chat'）
- `--domain`: 专业领域（可选：'computer', 'os', 'general'等）

## 许可

MIT 