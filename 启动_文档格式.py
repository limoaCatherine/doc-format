#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ROINF 文档格式。办公文件轻量排版（Word / Excel）。

用法:
  python 启动_文档格式.py
  python 启动_文档格式.py --docx-only
  python 启动_文档格式.py --xlsx-only
"""
from __future__ import annotations

import argparse
import io
import sys
from pathlib import Path
from typing import Optional, cast

if hasattr(sys.stdout, "reconfigure"):
    cast(io.TextIOWrapper, sys.stdout).reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    cast(io.TextIOWrapper, sys.stderr).reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

_SIBLING_REPOS = ("combat-sim", "scene-coverage", "attr-value", "scene-balance", "numeric-ssot", "doc-format")
for _name in _SIBLING_REPOS:
    _p = ROOT.parent / _name
    if _p.is_dir() and str(_p) not in sys.path:
        sys.path.insert(0, str(_p))


def main(argv: Optional[list[str]] = None) -> int:
    from 文档格式.config import TARGET_DIR
    from 文档格式.pipeline import run

    ap = argparse.ArgumentParser(
        description="办公文件轻量排版（Word .docx / Excel .xlsx）。无本地备份，请先 git commit。"
    )
    ap.add_argument("--root", default=TARGET_DIR, help="扫描根目录，默认 Limoa")
    ap.add_argument("--docx-only", action="store_true", help="只处理 Word")
    ap.add_argument("--xlsx-only", action="store_true", help="只处理 Excel")
    args = ap.parse_args(argv)
    return run(
        root=args.root,
        docx_only=args.docx_only,
        xlsx_only=args.xlsx_only,
    )


if __name__ == "__main__":
    raise SystemExit(main())
