# -*- coding: utf-8 -*-
"""扫描根目录并批量排版 Word / Excel。"""
from __future__ import annotations

import os
from typing import Optional

from 文档格式.config import TARGET_DIR
from 文档格式.excel import process_xlsx, xlsx_font_only
from 文档格式.scan import iter_files
from 文档格式.word import process_docx


def run(
    *,
    root: Optional[str] = None,
    docx_only: bool = False,
    xlsx_only: bool = False,
) -> int:
    root = os.path.abspath(root or TARGET_DIR)
    if not os.path.isdir(root):
        print(f"[错误] 找不到目录: {root}")
        return 1

    do_all = not (docx_only or xlsx_only)
    do_docx = do_all or docx_only
    do_xlsx = do_all or xlsx_only

    print(f"[扫描] {root}")
    print("[提示] 不写 .bak；回滚请用 git。")
    files = iter_files(root)
    print(f"[发现] docx={len(files['.docx'])}  xlsx={len(files['.xlsx'])}\n")

    if do_docx:
        print("=== Word (.docx) ===")
        ok = err = 0
        for path in files[".docx"]:
            rel = os.path.relpath(path, root)
            print(f"  {rel} ...", end=" ", flush=True)
            try:
                n = process_docx(path)
                print(f"[OK] {n} 表")
                ok += 1
            except PermissionError:
                print("[ERR] 文件被占用，请先关闭")
                err += 1
            except Exception as e:
                print(f"[ERR] {e}")
                err += 1
        print(f"  → Word 完成：成功 {ok}，失败 {err}\n")

    if do_xlsx:
        print("=== Excel (.xlsx) ===")
        ok = err = 0
        for path in files[".xlsx"]:
            rel = os.path.relpath(path, root)
            tag = "字体-only" if xlsx_font_only(path) else "轻美化"
            print(f"  {rel} ({tag}) ...", end=" ", flush=True)
            try:
                n = process_xlsx(path)
                print(f"[OK] 改字体 {n} 格")
                ok += 1
            except PermissionError:
                print("[ERR] 文件被占用，请先关闭")
                err += 1
            except Exception as e:
                print(f"[ERR] {e}")
                err += 1
        print(f"  → Excel 完成：成功 {ok}，失败 {err}\n")

    print("[完成] 办公文件排版结束。")
    return 0
