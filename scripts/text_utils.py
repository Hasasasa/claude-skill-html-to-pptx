"""text_utils.py — 文本处理共享工具。

CJK 字符范围的单一定义。assemble.py / convert.py / preflight.py 都从这里派生。
preflight 内嵌的 JS 用同样的 Unicode 范围（注释里手动同步）。
"""
import re

# CJK 字符 Unicode 范围（与 preflight SCAN_JS 的 cjkRe 保持一致）：
# - U+3000-U+303F  CJK 符号与标点（全角空格、句号、引号 等）
# - U+3040-U+30FF  日语 平假名 + 片假名
# - U+3400-U+4DBF  CJK 扩展 A
# - U+4E00-U+9FFF  CJK 统一表意（最常用区）
# - U+F900-U+FAFF  CJK 兼容表意
# - U+FF00-U+FFEF  半宽 / 全宽（中文标点常居此区）
CJK_RE = re.compile(
    "["
    "　-〿"
    "぀-ヿ"
    "㐀-鿿"
    "豈-﫿"
    "＀-￯"
    "]"
)


def is_cjk_text(s: str | None) -> bool:
    """文本里任一字符属于 CJK 范围则返回 True。空串 / None → False。"""
    return bool(s) and bool(CJK_RE.search(s))


# 给 JS 端用的等价范围字符串（手抄需要 sync）。preflight.SCAN_JS 用此模板插值。
CJK_JS_RANGE = r"[　-〿぀-ヿ㐀-鿿豈-﫿＀-￯]"
