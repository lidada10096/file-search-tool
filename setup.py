#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包脚本 - 使用 cx_Freeze 打包GUI版本
"""

import sys
import os
import shutil

# 清理旧的build目录
for d in ['build', 'dist', 'output']:
    if os.path.exists(d):
        try:
            shutil.rmtree(d)
        except:
            pass

from cx_Freeze import setup, Executable

# 依赖项
build_exe_options = {
    "packages": ["os", "json", "shutil", "re", "datetime", "difflib", "pathlib", "collections", "tkinter", "threading"],
    "includes": ["openpyxl", "openpyxl.styles", "openpyxl.utils"],
    "excludes": [],
    "include_files": ["config.json", "name.txt", "README.md"],
    "build_exe": "output"  # 使用不同的输出目录
}

# 基础配置
base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable(
        "gui_main.py",
        base=base,
        target_name="FileSearchTool.exe",
    )
]

setup(
    name="FileSearchTool",
    version="2.0.0",
    description="智能文件搜索与复制工具",
    options={"build_exe": build_exe_options},
    executables=executables,
)
