"""local_config.py — skill 本机状态加载 + 首次 seed。

- `.config.local.toml`（gitignored）：用户偏好（fonts.auto_install / cleanup.default）
- `references/lessons-learned.md`（gitignored）：本地沉淀工作副本；首次缺失从
  `lessons-learned.md.example`（committed 模板）seed

文件缺失 / 解析失败 / key 缺 → 走 `_DEFAULTS`。永不抛。
"""
from pathlib import Path

try:
    import tomllib  # py 3.11+
except ImportError:
    tomllib = None

SKILL_ROOT = Path(__file__).parent.parent
CONFIG_PATH = SKILL_ROOT / ".config.local.toml"
LESSONS_LOCAL = SKILL_ROOT / "references" / "lessons-learned.md"
LESSONS_TEMPLATE = SKILL_ROOT / "references" / "lessons-learned.md.example"

_DEFAULTS: dict = {
    "fonts": {"auto_install": "ask"},   # "yes" / "no" / "ask"
    "cleanup": {"default": "clean"},    # "clean" / "keep"
}


def load() -> dict:
    if not CONFIG_PATH.exists() or tomllib is None:
        return {k: {**v} for k, v in _DEFAULTS.items()}
    try:
        with CONFIG_PATH.open("rb") as f:
            user = tomllib.load(f) or {}
    except Exception as e:
        print(f"[config] 解析 {CONFIG_PATH.name} 失败（用默认）: {e}")
        return {k: {**v} for k, v in _DEFAULTS.items()}
    merged = {k: {**v} for k, v in _DEFAULTS.items()}
    for section, kvs in user.items():
        if section in merged and isinstance(kvs, dict):
            merged[section].update(kvs)
    return merged


def fonts_auto_install() -> str:
    return str(load().get("fonts", {}).get("auto_install", "ask")).lower()


def cleanup_default() -> str:
    return str(load().get("cleanup", {}).get("default", "clean")).lower()


def seed_lessons_learned() -> None:
    """首次运行：本地 lessons-learned.md 缺失就从 .example 模板复制一份。

    模板上游会被覆盖，本地副本不会——agent 在本地副本上自由加 / 改 / 整理，
    打算上游的条目作者手动复制回 .example 模板再提交。
    """
    if LESSONS_LOCAL.exists() or not LESSONS_TEMPLATE.exists():
        return
    LESSONS_LOCAL.write_text(LESSONS_TEMPLATE.read_text(encoding="utf-8"), encoding="utf-8")
    print(f"[lessons-learned] 首次运行：从模板 seed 本地工作副本 → {LESSONS_LOCAL.name}")
