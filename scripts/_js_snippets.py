"""_js_snippets.py — measure.py 与 preflight.py 共享的浏览器端 JS 片段。

抽这一份是因为 deco 检测要在两边保持完全一致——preflight 报告"会走 deco_snapshot
的元素"，measure 实际走 deco_snapshot 的判定，两者必须用同一组判定函数。

每个 const 是一段可注入 IIFE 内的函数定义；用 `from _js_snippets import DECO_HELPERS`
然后拼到 JS 字符串前面即可。
"""

# 注入在 IIFE 顶部的工具函数集合：
# - isNonTranslateTransform(transformStr) → bool
# - isClippingContainerWithTransformedChildren(s, el) → bool
# - hasPseudoDecoration(el, pseudo) → bool（::before/::after 是否真的画了装饰）
# - hasComplexDecoration(s, el) → bool（命中即 measure 走 deco_snapshot 路径）
DECO_HELPERS = r"""
  // CSS transform 非平移（含 rotate / skew / scale）。matrix(a,b,c,d,tx,ty)：
  // 纯平移要求 a=d=1 && b=c=0。matrix3d / 关键字形式当成非平移处理。
  const isNonTranslateTransform = (transformStr) => {
    if (!transformStr || transformStr === 'none') return false;
    const m = transformStr.match(/^matrix\(([^)]+)\)$/);
    if (m) {
      const v = m[1].split(',').map(parseFloat);
      const [a, b, c, d] = v;
      return Math.abs(b) > 0.001 || Math.abs(c) > 0.001 ||
             Math.abs(a - 1) > 0.001 || Math.abs(d - 1) > 0.001;
    }
    return true;
  };

  // overflow:hidden/clip 容器 + 含 transformed 子 → 容器是裁切框，子被裁。
  // 这种模式必须把容器整块截图，子的旋转 AABB 远大于裁切框，单独画会溢出。
  const isClippingContainerWithTransformedChildren = (s, el) => {
    const ov = s.overflow, ovx = s.overflowX, ovy = s.overflowY;
    const clipped = ov === 'hidden' || ov === 'clip' ||
                    ovx === 'hidden' || ovx === 'clip' ||
                    ovy === 'hidden' || ovy === 'clip';
    if (!clipped || !el.children.length) return false;
    for (const ch of el.children) {
      const cs = getComputedStyle(ch);
      if (isNonTranslateTransform(cs.transform)) return true;
    }
    return false;
  };

  // 伪元素装饰：non-empty content 或 空 content+background-image。
  // 单独抽出来给 preflight 复用（preflight 不需要走截图，只需要知道有装饰）。
  const hasPseudoDecoration = (el, pseudo) => {
    const ps = getComputedStyle(el, pseudo);
    const content = ps.content;
    const hasContent = content && content !== 'none' && content !== 'normal'
                       && content !== '""' && content !== "''";
    if (hasContent) return true;
    if (content === '""' || content === "''") {
      if (ps.backgroundImage && ps.backgroundImage !== 'none') return true;
    }
    return false;
  };

  // 通用复杂装饰：命中任何一项 measure 就走"整块截图嵌入"路径
  const hasComplexDecoration = (s, el) => {
    if (s.backgroundImage && s.backgroundImage !== 'none') return true;
    if (s.boxShadow && s.boxShadow !== 'none') return true;
    if (s.outlineStyle && s.outlineStyle !== 'none' && parseFloat(s.outlineWidth) > 0) return true;
    if (hasPseudoDecoration(el, '::before')) return true;
    if (hasPseudoDecoration(el, '::after')) return true;
    if (isClippingContainerWithTransformedChildren(s, el)) return true;
    return false;
  };
"""
