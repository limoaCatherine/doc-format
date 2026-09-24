# 办公文件排版说明

唯一对外入口：`NumericalTools/启动_文档格式.py`。  
包内模块是内部实现，**不要**当第二入口去跑。

扫描默认根目录：`D:\Limoa`。  
只处理已有 **Word（.docx）** 与 **Excel（.xlsx）** 的格式；**不做** Markdown 转换。  
**不写本地 .bak**；跑之前请先 `git commit`。

## 怎么跑

```bash
cd NumericalTools
python 启动_文档格式.py
```

| 参数 | 作用 |
|------|------|
| （无参） | 扫根目录下全部 `.docx` / `.xlsx` |
| `--root 路径` | 改扫描根目录 |
| `--docx-only` | 只处理 Word |
| `--xlsx-only` | 只处理 Excel |

## 会做什么

- **Word**：微软雅黑、标题/正文行距、表格边框斑马、页眉页码  
- **Excel**：统一雅黑；非框架表可加浅灰表头/列宽；文件名含「战斗数值框架」「标模平衡」或体积 ≥1.5MB 时只改字体  

## 包内结构（内部，不用手点）

```text
文档格式/
  config.py / scan.py / pipeline.py
  word.py / excel.py
```
