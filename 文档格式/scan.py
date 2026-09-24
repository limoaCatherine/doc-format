# -*- coding: utf-8 -*-
"""扫描与原子写盘。"""
from __future__ import annotations

import os
import tempfile
from typing import Dict, List

from docx.document import Document as DocumentType

from 文档格式.config import SKIP_DIR_NAMES


def iter_files(root: str) -> Dict[str, List[str]]:
    """一次 walk，按后缀分流（仅 Word / Excel）。"""
    out: Dict[str, List[str]] = {".docx": [], ".xlsx": []}
    root = os.path.abspath(root)
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [
            d for d in dirnames
            if d not in SKIP_DIR_NAMES and not d.endswith(".disabled")
        ]
        for name in filenames:
            ext = os.path.splitext(name)[1].lower()
            if ext not in out:
                continue
            if name.startswith("~$") or name.startswith(".~"):
                continue
            path = os.path.join(dirpath, name)
            if path.replace("\\", "/").find("/ROINF文档备份/") >= 0:
                continue
            out[ext].append(path)
    for k in out:
        out[k].sort()
    return out


def atomic_save_docx(doc: DocumentType, path: str) -> None:
    """先写临时文件再替换，规避 Windows 上部分路径直接 save 的 Errno 22。"""
    folder = os.path.dirname(path) or "."
    fd, tmp = tempfile.mkstemp(suffix=".docx", dir=folder)
    os.close(fd)
    try:
        doc.save(tmp)
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        raise


def atomic_save_xlsx(wb, path: str) -> None:
    folder = os.path.dirname(path) or "."
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=folder)
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        raise
