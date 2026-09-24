# -*- coding: utf-8 -*-
"""办公文件排版：路径与样式常量。"""
from __future__ import annotations

from pathlib import Path

from docx.shared import Pt, RGBColor

_PKG_DIR = Path(__file__).resolve().parent
_TOOLS_DIR = _PKG_DIR.parent
_ROINF_DIR = _TOOLS_DIR.parent
_LIMOA_DIR = _ROINF_DIR.parent  # D:\Limoa

TARGET_DIR = str(_LIMOA_DIR)
SKIP_DIR_NAMES = {
    ".git", ".venv", "venv", "node_modules", "__pycache__", ".cursor",
    "cursor",
    "ROINF文档备份", "游戏包体", "ComfyUI", "ro_lora_training",
    ".pytest_cache", "dist", "build", ".idea",
}

# 框架配置表：只改字体名，不做表头/列宽/标签美化
# framework layout is owned by numerical design workbook / opt scripts; formatter only touches fonts
XLSX_FONT_ONLY_KEYWORDS = ("战斗数值框架", "标模平衡")

FONT_NAME = "微软雅黑"

SIZE_H1 = Pt(16)
SIZE_H2 = Pt(14)
SIZE_H3 = Pt(12)
SIZE_BODY = Pt(11)
SIZE_HEADER = Pt(11)

COLOR_BORDER = "B0B0B0"
COLOR_HEADER_BG = "F2F2F2"
COLOR_HEADER_FG = RGBColor(0x33, 0x33, 0x33)
COLOR_ZEBRA = "FAFAFA"
COLOR_TOTAL = "FFF2CC"
COLOR_H1 = RGBColor(0x22, 0x22, 0x22)
COLOR_H2 = RGBColor(0x33, 0x33, 0x33)
COLOR_H3 = RGBColor(0x44, 0x44, 0x44)

HIGHLIGHT_KEYWORDS = ("合计", "总计", "Total", "总和", "平均", "Average", "TOTAL")
