# 文档排版

给已有的 Word 和 Excel 统一字体和版式。不把 Markdown 转成办公文件，也不改数值框架表里的表头、列宽和标签。文件名含「战斗数值框架」或「标模平衡」，或体积达到 1.5MB 的表，只改字体。

不写 `.bak`。跑之前先提交或另存。

## 运行

```bash
pip install -r requirements.txt
python 启动_文档格式.py --root 目录
```

| 参数 | 作用 |
|------|------|
| `--root` | 扫描根目录 |
| `--docx-only` | 只处理 Word |
| `--xlsx-only` | 只处理 Excel |

不传 `--root` 时，扫描位置由脚本所在目录向上推算，建议显式传入。

## 许可

[MIT](LICENSE)。
