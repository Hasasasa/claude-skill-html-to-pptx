"""font_paths.py — 字体缓存目录的单一权威源。

跨项目共享，不随 skill 删除 / 重装而丢。Windows 用 LOCALAPPDATA，
其他系统按 XDG 规范走 ~/.cache。
"""
import os
from pathlib import Path


def user_cache_dir() -> Path:
    """返回 user-home 字体缓存目录，确保存在。"""
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA")
        if not base:
            base = str(Path.home() / "AppData" / "Local")
        root = Path(base) / "html-to-pptx" / "fonts"
    else:
        base = os.environ.get("XDG_CACHE_HOME") or str(Path.home() / ".cache")
        root = Path(base) / "html-to-pptx" / "fonts"
    root.mkdir(parents=True, exist_ok=True)
    return root


# 模块级常量，给 embed_fonts.py / font_resolver.py 用
CACHE_DIR = user_cache_dir()
