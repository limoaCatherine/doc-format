# 文档排版（doc-format）

对已有 Word（`.docx`）与 Excel（`.xlsx`）做轻量、就地排版：统一字体、标题层级、表格样式与页眉页码。

- **不做** Markdown → 办公文件转换
- **不写** `.bak`；请先提交或自行备份
- 文件名含「战斗数值框架」「标模平衡」，或体积 ≥ 1.5MB 的 Excel，只改字体，不改表头/列宽/标签

## 布局

```
doc-format/
  启动_文档格式.py
  文档格式/
    config.py / scan.py / pipeline.py
    word.py / excel.py
```

## 运行

需要 Python 3.11+。

```bash
pip install -r requirements.txt
python 启动_文档格式.py --root path\to\docs
python 启动_文档格式.py --root path\to\docs --docx-only
python 启动_文档格式.py --root path\to\docs --xlsx-only
```

| 参数 | 说明 |
|------|------|
| `--root` | 扫描根目录（强烈建议显式传入；默认当前工作目录） |
| `--docx-only` | 只处理 Word |
| `--xlsx-only` | 只处理 Excel |

## 许可

[MIT](LICENSE)
