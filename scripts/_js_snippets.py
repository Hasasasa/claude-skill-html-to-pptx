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
# - isSimplePseudoLineDecoration(el, pseudo) → bool（可用 PPT shape 表达的伪元素线框）
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

  // 伪元素装饰：non-empty content（string/url/counter 等）或
  // 空 content + 任何可见装饰（背景图 / 实色背景 / 边框）。
  // 漏检 background-color / border：doodle-frame::after 这种实色填充 + asym border-radius
  // 会整个丢失（host 走普通 shape 档不带 ::after 渲染）。
  // 单独抽出来给 preflight 复用（preflight 不需要走截图，只需要知道有装饰）。
  const hasPseudoDecoration = (el, pseudo) => {
    const ps = getComputedStyle(el, pseudo);
    const content = ps.content;
    const hasContent = content && content !== 'none' && content !== 'normal'
                       && content !== '""' && content !== "''";
    if (hasContent) return true;
    if (content === '""' || content === "''") {
      if (ps.backgroundImage && ps.backgroundImage !== 'none') return true;
      const bg = ps.backgroundColor;
      if (bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent') return true;
      if (parseFloat(ps.borderTopWidth) > 0 || parseFloat(ps.borderBottomWidth) > 0 ||
          parseFloat(ps.borderLeftWidth) > 0 || parseFloat(ps.borderRightWidth) > 0) return true;
    }
    return false;
  };

  const isZeroCssLength = (value) => {
    if (!value) return true;
    const n = parseFloat(value);
    return Number.isFinite(n) && Math.abs(n) < 0.001;
  };

  // 可矢量化的伪元素线框：空 content + 透明背景 + 简单 border。
  // 这类应转成 PPT shape，保持可编辑；不要走 deco_snapshot。
  // 只覆盖保守子集：无 transform / filter / shadow / radius / background-image。
  const isSimplePseudoLineDecoration = (el, pseudo) => {
    const ps = getComputedStyle(el, pseudo);
    const content = ps.content;
    const emptyContent = content === '""' || content === "''";
    if (!emptyContent) return false;
    if (ps.display === 'none' || ps.visibility === 'hidden') return false;
    if (ps.backgroundImage && ps.backgroundImage !== 'none') return false;
    const bg = ps.backgroundColor;
    if (bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent') return false;
    if (ps.boxShadow && ps.boxShadow !== 'none') return false;
    if (ps.filter && ps.filter !== 'none') return false;
    if (ps.mixBlendMode && ps.mixBlendMode !== 'normal') return false;
    if (ps.transform && ps.transform !== 'none') return false;
    if (!isZeroCssLength(ps.borderTopLeftRadius) ||
        !isZeroCssLength(ps.borderTopRightRadius) ||
        !isZeroCssLength(ps.borderBottomRightRadius) ||
        !isZeroCssLength(ps.borderBottomLeftRadius)) return false;
    return parseFloat(ps.borderTopWidth) > 0 || parseFloat(ps.borderBottomWidth) > 0 ||
           parseFloat(ps.borderLeftWidth) > 0 || parseFloat(ps.borderRightWidth) > 0;
  };

  const hasRasterPseudoDecoration = (el, pseudo) =>
    hasPseudoDecoration(el, pseudo) && !isSimplePseudoLineDecoration(el, pseudo);

  // CSS transform 矩阵是否"非纯旋转+平移"（含 skew / 非均匀 scale）。
  // 纯 rotate+translate 的 2D matrix 满足正交：a²+b² ≈ 1 且 c²+d² ≈ 1 且 a*c+b*d ≈ 0。
  // 不满足 → OOXML 没原语表达（OOXML 只有 rot），必须走截图。
  // 仅 translate 早就走 layout 路径，不在这里判定。
  const isUnrepresentableTransform = (transformStr) => {
    if (!transformStr || transformStr === 'none') return false;
    const m = transformStr.match(/^matrix\(([^)]+)\)$/);
    if (!m) return true;  // matrix3d 或关键字形式都当不可表达
    const v = m[1].split(',').map(parseFloat);
    const [a, b, c, d] = v;
    const len1 = a*a + b*b;
    const len2 = c*c + d*d;
    const dot  = a*c + b*d;
    const eps  = 0.005;
    return Math.abs(len1 - 1) > eps || Math.abs(len2 - 1) > eps || Math.abs(dot) > eps;
  };

  // 元素 filter 是否"非平凡"：none / 空都不算。drop-shadow / blur / saturate 等都算。
  const hasNontrivialFilter = (filterStr) => {
    if (!filterStr) return false;
    const t = filterStr.trim();
    return t !== '' && t !== 'none';
  };

  // 通用复杂装饰：命中任何一项 measure 就走"整块截图嵌入"路径
  const hasComplexDecoration = (s, el) => {
    if (s.backgroundImage && s.backgroundImage !== 'none') return true;
    if (s.boxShadow && s.boxShadow !== 'none') return true;
    if (s.outlineStyle && s.outlineStyle !== 'none' && parseFloat(s.outlineWidth) > 0) return true;
    if (hasRasterPseudoDecoration(el, '::before')) return true;
    if (hasRasterPseudoDecoration(el, '::after')) return true;
    if (isClippingContainerWithTransformedChildren(s, el)) return true;
    // backdrop-filter（毛玻璃）— OOXML 无原语，走截图。截图天然包含被模糊的背景
    if (s.backdropFilter && s.backdropFilter !== 'none') return true;
    // filter（blur / drop-shadow / saturate / 等）— 同上
    if (hasNontrivialFilter(s.filter)) return true;
    // mix-blend-mode：浏览器截元素 BCR 时拿到的就是已混合的像素，OOXML 表达不了
    if (s.mixBlendMode && s.mixBlendMode !== 'normal') return true;
    // skew / 非均匀 scale：OOXML 只有 rot，其它非线性变换走截图
    if (isUnrepresentableTransform(s.transform)) return true;
    return false;
  };
"""
