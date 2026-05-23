"""measure.py — 用浏览器实测 HTML 中所有 slide 的可见元素，输出 measurement JSON。

Usage:
    python measure.py <html_path> <out_json> [slide_index]
    - 不传 slide_index：抽取全部 slide，输出 { "slides": [ {slide:..., records:[...]}, ... ] }
    - 传 slide_index：抽取单页，输出 { "slide":..., "records":[...] }（兼容旧 API）

约定：
- 视口固定 1920x1080
- slide_index 从 0 开始（对应 document.querySelectorAll('section.slide')[i]）
"""
import json
import sys
import tempfile
from pathlib import Path

from playwright.sync_api import sync_playwright

from _js_snippets import DECO_HELPERS

VIEWPORT = {"width": 1920, "height": 1080}

# 在浏览器上下文里执行：抽取当前被 adapter 标记的 slide 的所有可视绘制单元
# 约定：adapter 的 activate_js 必须给目标 slide 设置 data-pptx-target 属性
# DECO_HELPERS 注入 isNonTranslateTransform / isClippingContainerWithTransformedChildren /
# hasPseudoDecoration / hasComplexDecoration 四个工具；measure 与 preflight 共享，
# 避免两边漂移（修在 _js_snippets.py 单点）。
EXTRACT_JS = r"""
(slideIndex) => {
  const slide = document.querySelector('[data-pptx-target]');
  if (!slide) return { error: 'no slide tagged with data-pptx-target' };

  // 把目标 slide 滚动到视口内（容器使用 scroll-snap）
  slide.scrollIntoView({block:'start', inline:'start', behavior:'instant'});

  // 强制等一帧让 layout 稳定（同步：DOM 读取已经触发回流）
  const slideRect = slide.getBoundingClientRect();

  // 我们要抽：
  // 1) text leaves（含 textContent 的最深节点，且没有可见子文本节点冲突）
  // 2) <img>
  // 3) <svg>（整体序列化）
  // 4) 装饰节点（有 border / 非透明背景 / 非默认）

  const css = (el, prop) => getComputedStyle(el).getPropertyValue(prop);
  const isHidden = (el) => {
    const s = getComputedStyle(el);
    if (s.display === 'none' || s.visibility === 'hidden' || parseFloat(s.opacity) === 0) return true;
    const r = el.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) return true;
    return false;
  };

""" + DECO_HELPERS + r"""

  // 累计 transform 旋转角度（度数；含祖先节点的 rotate）
  // 用 matrix(a, b, ...) 反解 atan2(b, a)。仅处理纯旋转分量；
  // skew / 非均匀 scale 暂不还原（影响极少数模板）
  const cumulativeRotation = (el) => {
    let total = 0;
    let cur = el;
    while (cur && cur !== document.body) {
      const t = getComputedStyle(cur).transform;
      if (t && t !== 'none') {
        // matrix(a, b, c, d, e, f) 或 matrix3d(...)
        const m = t.match(/matrix(?:3d)?\(([^)]+)\)/);
        if (m) {
          const v = m[1].split(',').map(parseFloat);
          // 取 a, b（2D 矩阵） 或 m11, m12（3D 矩阵）
          const a = v[0], b = v[1];
          total += Math.atan2(b, a) * 180 / Math.PI;
        }
      }
      cur = cur.parentElement;
    }
    return total;
  };

  const records = [];
  let nodeId = 0;

  // 标记一个节点是否为 "text leaf"：包含 textContent 但所有子节点要么是文本节点，要么是 inline 装饰（em/span 等没有进一步分割结构的）
  // 简化：只要这个元素的 children 中没有任何 block 级元素，就算 text leaf。
  const BLOCK_TAGS = new Set(['DIV','SECTION','ARTICLE','HEADER','FOOTER','MAIN','NAV','P','H1','H2','H3','H4','H5','H6','UL','OL','LI','FIGURE','FIGCAPTION','TABLE','PRE','SVG','IMG','CANVAS','VIDEO']);
  const isAtomicInline = (node) => {
    if (!node || node.nodeType !== 1) return false;
    if (BLOCK_TAGS.has(node.tagName.toUpperCase())) return false;
    const s = getComputedStyle(node);
    const bg = s.backgroundColor;
    const hasBg = bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent';
    const hasBorder = parseFloat(s.borderTopWidth) > 0 || parseFloat(s.borderBottomWidth) > 0 ||
                      parseFloat(s.borderLeftWidth) > 0 || parseFloat(s.borderRightWidth) > 0;
    return s.position === 'absolute' || s.position === 'fixed' ||
           s.display === 'inline-block' || s.display === 'inline-flex' ||
           hasBg || hasBorder;
  };
  // svg / img 等元素即便不是 HTML 块级，也要阻断 text-leaf 判定，
  // 否则容器里同时存在 <svg> 与 <span> 文本时,会被错误地当作纯文本叶子整体吞掉。
  const isTextLeaf = (el) => {
    if (!el.textContent || !el.textContent.trim()) return false;
    for (const ch of el.children) {
      if (BLOCK_TAGS.has(ch.tagName.toUpperCase())) return false;
      if (isAtomicInline(ch)) return false;
    }
    return true;
  };

  // 富文本 runs：把一个 text leaf 拆成多个 run，每个 run 携带自己的 computed style
  // 这样 <em> 等内嵌强调可以保留独立字体
  const extractRuns = (el) => {
    const runs = [];
    const walk = (n) => {
      if (n.nodeType === 3) {
        const text = n.nodeValue;
        if (!text) return;
        const parent = n.parentElement;
        const s = getComputedStyle(parent);
        runs.push({
          text,
          fontFamily: s.fontFamily,
          fontSize: parseFloat(s.fontSize),
          fontWeight: s.fontWeight,
          fontStyle: s.fontStyle,
          color: s.color,
          letterSpacing: s.letterSpacing,
          textDecoration: s.textDecorationLine,
          textShadow: s.textShadow,
        });
      } else if (n.nodeType === 1) {
        // skip <br>: emit a soft break marker
        if (n.tagName === 'BR') {
          runs.push({ text: '\n', linebreak: true });
          return;
        }
        for (const ch of n.childNodes) walk(ch);
      }
    };
    walk(el);
    return runs;
  };

  // 选择需要导出的节点
  const walk = (el) => {
    if (!el || el.nodeType !== 1) return;
    if (isHidden(el)) return;

    // SVG 整体作为一个节点导出
    if (el.tagName.toLowerCase() === 'svg') {
      const r = el.getBoundingClientRect();
      // 给元素打 marker，便于 Playwright 后续按 marker 截图
      const svgIndex = records.filter(x => x.kind === 'svg').length;
      el.setAttribute('data-pptx-svg-id', `slide${slideIndex+1}-svg${svgIndex+1}`);
      records.push({
        id: nodeId++,
        kind: 'svg',
        tag: 'svg',
        rect: rectRel(r),
        marker: `slide${slideIndex+1}-svg${svgIndex+1}`,
        outerHTML: el.outerHTML,
        color: css(el, 'color'),
      });
      return; // SVG 内部不再下钻
    }

    if (el.tagName.toLowerCase() === 'img') {
      const r = el.getBoundingClientRect();
      const imgIndex = records.filter(x => x.kind === 'img').length;
      el.setAttribute('data-pptx-img-id', `slide${slideIndex+1}-img${imgIndex+1}`);
      records.push({
        id: nodeId++,
        kind: 'img',
        tag: 'img',
        rect: rectRel(r),
        marker: `slide${slideIndex+1}-img${imgIndex+1}`,
        src: el.currentSrc || el.src,
      });
      return;
    }

    // canvas（Chart.js / WebGL / 自绘图）— 整体作为 picture 截图嵌入
    // canvas 像素无法用 OOXML 表达，但 measure 阶段已经等过动画稳定，
    // 截首帧足以还原静态呈现（图表 / 背景纹理 / IP 装饰）。
    if (el.tagName.toLowerCase() === 'canvas') {
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) {
        const canvasIndex = records.filter(x => x.kind === 'canvas').length;
        const marker = `slide${slideIndex+1}-canvas${canvasIndex+1}`;
        el.setAttribute('data-pptx-canvas-id', marker);
        records.push({
          id: nodeId++,
          kind: 'canvas',
          tag: 'canvas',
          rect: rectRel(r),
          naturalSize: { w: el.offsetWidth, h: el.offsetHeight },
          rotation: cumulativeRotation(el),
          marker,
        });
      }
      return;
    }

    const s = getComputedStyle(el);

    // 装饰：边框或非透明背景色（且不是默认透明）
    const bg = s.backgroundColor;
    const hasBg = bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent';
    const borderTop = parseFloat(s.borderTopWidth) > 0;
    const borderBottom = parseFloat(s.borderBottomWidth) > 0;
    const borderLeft = parseFloat(s.borderLeftWidth) > 0;
    const borderRight = parseFloat(s.borderRightWidth) > 0;
    const hasBorder = borderTop || borderBottom || borderLeft || borderRight;

    // 通用装饰捕获：任何元素只要有 background-image / box-shadow / 伪元素装饰
    // → 整块截图嵌入作为底层；子节点（文字 / 子装饰）按原流程继续处理画在之上
    // 这套机制覆盖所有 PPT 几何原语无法表达的 CSS 装饰，无需为每种新增加 patch
    if (hasComplexDecoration(s, el)) {
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) {
        const decoIndex = records.filter(x => x.kind === 'deco_snapshot').length;
        const marker = `slide${slideIndex+1}-deco${decoIndex+1}`;
        el.setAttribute('data-pptx-deco-id', marker);
        records.push({
          id: nodeId++,
          kind: 'deco_snapshot',
          tag: el.tagName.toLowerCase(),
          rect: rectRel(r),
          naturalSize: { w: el.offsetWidth, h: el.offsetHeight },
          rotation: cumulativeRotation(el),
          marker,
        });
        // overflow:hidden 裁切容器：容器 PNG 已经包含被裁后的子装饰，子不再单独处理
        // （否则旋转子的 AABB 远大于裁切框，单独画会变成超大色块覆盖周围）
        // 例外：slide 根节点。slide 根的 overflow:hidden 是布局结构（裁视口），
        // 不是"装饰裁切意图"。若 slide 根本身命中此分支会吞掉所有 text/svg 子记录。
        if (el !== slide && isClippingContainerWithTransformedChildren(s, el)) {
          return;
        }
        // 其他情况（背景图 / box-shadow / 伪元素装饰）：子节点继续画在截图之上
      }
    }

    // text leaf：直接导出文本节点
    if (isTextLeaf(el)) {
      const r = el.getBoundingClientRect();
      const runs = extractRuns(el);
      const rotDeg = cumulativeRotation(el);
      records.push({
        id: nodeId++,
        kind: 'text',
        tag: el.tagName.toLowerCase(),
        className: el.className || '',
        rect: rectRel(r),
        // 元素未旋转的尺寸（不含 transform 效果），用于旋转还原
        naturalSize: { w: el.offsetWidth, h: el.offsetHeight },
        rotation: rotDeg,
        runs,
        style: {
          color: s.color,
          fontFamily: s.fontFamily,
          fontSize: parseFloat(s.fontSize),
          fontWeight: s.fontWeight,
          fontStyle: s.fontStyle,
          lineHeight: s.lineHeight,
          letterSpacing: s.letterSpacing,
          textAlign: s.textAlign,
          textTransform: s.textTransform,
          opacity: s.opacity,
          // CSS padding（textbox 内 margin 用，让 textbox 几何严格 = BCR 时不丢 padding）
          paddingTop: parseFloat(s.paddingTop) || 0,
          paddingRight: parseFloat(s.paddingRight) || 0,
          paddingBottom: parseFloat(s.paddingBottom) || 0,
          paddingLeft: parseFloat(s.paddingLeft) || 0,
          // flex / grid 居中：HTML 用 align-items / justify-content 在容器内居中
          // 文字时要翻译成 OOXML 的 anchor (垂直) + algn (水平)
          display: s.display,
          alignItems: s.alignItems,
          justifyContent: s.justifyContent,
        },
        deco: { hasBg, bg, borderTop, borderBottom, borderLeft, borderRight,
                borderColor: s.borderTopColor,
                borderTopWidth: parseFloat(s.borderTopWidth),
                borderBottomWidth: parseFloat(s.borderBottomWidth),
                borderLeftWidth: parseFloat(s.borderLeftWidth),
                borderRightWidth: parseFloat(s.borderRightWidth),
                borderRadius: s.borderTopLeftRadius,  // 用 top-left 代表整体；'50%' 或 px 值
              },
        text: el.innerText,
      });
      return;
    }

    // 非 text leaf 但有装饰 → 单独导出形状（不带文本，文本由子节点输出）
    if (hasBg || hasBorder) {
      const r = el.getBoundingClientRect();
      const rotDeg = cumulativeRotation(el);
      records.push({
        id: nodeId++,
        kind: 'shape',
        tag: el.tagName.toLowerCase(),
        rect: rectRel(r),
        // 元素未旋转的尺寸（不含 transform 效果），用于旋转还原
        naturalSize: { w: el.offsetWidth, h: el.offsetHeight },
        rotation: rotDeg,
        deco: { hasBg, bg, borderTop, borderBottom, borderLeft, borderRight,
                borderColor: s.borderTopColor,
                borderTopWidth: parseFloat(s.borderTopWidth),
                borderBottomWidth: parseFloat(s.borderBottomWidth),
                borderLeftWidth: parseFloat(s.borderLeftWidth),
                borderRightWidth: parseFloat(s.borderRightWidth),
                borderRadius: s.borderTopLeftRadius,  // 用 top-left 代表整体；'50%' 或 px 值
              },
      });
    }

    // 处理"混合容器"：当 el 既有 block 子，又有直接挂着的 text node / inline 元素，
    // 直接文本节点不会进入 el.children 遍历也不属于 isTextLeaf 分支，
    // 必须单独抓取为 inline-group text 记录，否则会丢字（典型：callout）
    emitInlineGroupsAround(el);

    // 继续下钻 block 子
    for (const ch of el.children) {
      if (BLOCK_TAGS.has(ch.tagName.toUpperCase())) walk(ch);
    }
  };

  // 把 el 的子节点按 block 边界切成若干 inline group；每个 group 单独发一个 text 记录
  const emitInlineGroupsAround = (el) => {
    const groups = [];
    let cur = null;
    for (const ch of el.childNodes) {
      const isElem = ch.nodeType === 1;
      const isBlock = isElem && BLOCK_TAGS.has(ch.tagName.toUpperCase());
      if (isBlock || isAtomicInline(ch)) {
        if (cur) { groups.push(cur); cur = null; }
        if (isAtomicInline(ch)) groups.push([ch]);
        continue;
      }
      // text node 或 inline element
      if (ch.nodeType === 3 && (!ch.nodeValue || !ch.nodeValue.trim())) {
        // 单独空白文本节点：若已有 cur，把它纳入；否则忽略
        if (cur) cur.push(ch);
        continue;
      }
      if (!cur) cur = [];
      cur.push(ch);
    }
    if (cur) groups.push(cur);

    for (const group of groups) {
      if (!group.length) continue;
      // 剥掉 group 首尾的 <br>：它们没几何意义（零宽软换行），
      // 留着会让 range BCR 跨进上一行/下一行 → record h 翻倍 → 后续 textbox 撞下方相邻段
      let trimmed = group;
      while (trimmed.length && trimmed[0].nodeType === 1 && trimmed[0].tagName === 'BR') {
        trimmed = trimmed.slice(1);
      }
      while (trimmed.length && trimmed[trimmed.length - 1].nodeType === 1 && trimmed[trimmed.length - 1].tagName === 'BR') {
        trimmed = trimmed.slice(0, -1);
      }
      if (!trimmed.length) continue;
      const range = document.createRange();
      range.setStartBefore(trimmed[0]);
      range.setEndAfter(trimmed[trimmed.length - 1]);
      const rects = range.getClientRects();
      if (!rects.length) continue;
      // 取整组并集 rect（多行时 getBoundingClientRect 会给到完整范围）
      const atomicEl = trimmed.length === 1 && isAtomicInline(trimmed[0]) ? trimmed[0] : null;
      const r = atomicEl ? atomicEl.getBoundingClientRect() : range.getBoundingClientRect();
      if (r.width === 0 || r.height === 0) continue;
      // 跳过纯空白（trim 后没文本）
      const txt = (atomicEl ? atomicEl.textContent : range.toString()).replace(/\s+/g, ' ').trim();
      if (!txt) continue;

      // 用 trimmed 组的代表元素取样式：第一个 element 节点；找不到就用父 el
      // （用 trimmed 而非原 group：原 group 里如果首节点是 <br> 会把 styleHost 错设成 br）
      let styleHost = el;
      for (const n of trimmed) { if (n.nodeType === 1) { styleHost = n; break; } }
      const gs = getComputedStyle(styleHost);
      const bg = gs.backgroundColor;
      const hasBg = bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent';
      const borderTop = parseFloat(gs.borderTopWidth) > 0;
      const borderBottom = parseFloat(gs.borderBottomWidth) > 0;
      const borderLeft = parseFloat(gs.borderLeftWidth) > 0;
      const borderRight = parseFloat(gs.borderRightWidth) > 0;

      // 抽取 runs（遍历每个 group 成员）
      const runs = [];
      const walkInline = (n) => {
        if (n.nodeType === 3) {
          if (!n.nodeValue) return;
          const p = n.parentElement;
          const ps = getComputedStyle(p);
          runs.push({
            text: n.nodeValue,
            fontFamily: ps.fontFamily,
            fontSize: parseFloat(ps.fontSize),
            fontWeight: ps.fontWeight,
            fontStyle: ps.fontStyle,
            color: ps.color,
            letterSpacing: ps.letterSpacing,
            textDecoration: ps.textDecorationLine,
            textShadow: ps.textShadow,
          });
        } else if (n.nodeType === 1) {
          if (n.tagName === 'BR') { runs.push({ text: '\n', linebreak: true }); return; }
          for (const c of n.childNodes) walkInline(c);
        }
      };
      // 用 trimmed 抽 runs：首尾 br 不进 runs（避免 OOXML 多出空行把单行文本撑高错位）
      for (const n of trimmed) walkInline(n);

      records.push({
        id: nodeId++,
        kind: 'text',
        tag: styleHost.tagName.toLowerCase() + '#inline',
        className: styleHost.className || '',
        rect: rectRel(r),
        runs,
        style: {
          color: gs.color,
          fontFamily: gs.fontFamily,
          fontSize: parseFloat(gs.fontSize),
          fontWeight: gs.fontWeight,
          fontStyle: gs.fontStyle,
          lineHeight: gs.lineHeight,
          letterSpacing: gs.letterSpacing,
          textAlign: gs.textAlign,
          textTransform: gs.textTransform,
          opacity: gs.opacity,
          paddingTop: parseFloat(gs.paddingTop) || 0,
          paddingRight: parseFloat(gs.paddingRight) || 0,
          paddingBottom: parseFloat(gs.paddingBottom) || 0,
          paddingLeft: parseFloat(gs.paddingLeft) || 0,
          display: gs.display,
          alignItems: gs.alignItems,
          justifyContent: gs.justifyContent,
        },
        deco: { hasBg, bg, borderTop, borderBottom, borderLeft, borderRight,
                borderColor: gs.borderTopColor,
                borderTopWidth: parseFloat(gs.borderTopWidth),
                borderBottomWidth: parseFloat(gs.borderBottomWidth),
                borderLeftWidth: parseFloat(gs.borderLeftWidth),
                borderRightWidth: parseFloat(gs.borderRightWidth),
                borderRadius: gs.borderTopLeftRadius },
        text: txt,
      });
    }
  };

  const rectRel = (r) => ({
    x: r.left - slideRect.left,
    y: r.top - slideRect.top,
    w: r.width,
    h: r.height,
  });

  walk(slide);

  // 找一个有"实际"背景色的祖先：slide 自己若 transparent，向上回退到 body
  const opaqueBg = (el) => {
    let cur = el;
    while (cur) {
      const bg = getComputedStyle(cur).backgroundColor;
      if (bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent') return bg;
      cur = cur.parentElement;
    }
    return 'rgb(255, 255, 255)';
  };

  return {
    slide: {
      width: slideRect.width,
      height: slideRect.height,
      theme: slide.className,  // hero dark / light 等
      background: opaqueBg(slide),
      color: getComputedStyle(slide).color,
    },
    records,
  };
}
"""


_DECO_HIDE_FOREGROUND_JS = r"""(marker) => {
    const deco = document.querySelector(`[data-pptx-deco-id='${marker}']`);
    if (!deco) return;
    const slide = deco.closest('[data-pptx-target]') || document.querySelector('[data-pptx-target]');
    if (!slide) return;
    const isAncestor = (el) => {
        let cur = deco.parentElement;
        while (cur) { if (cur === el) return true; cur = cur.parentElement; }
        return false;
    };
    const hasDirectText = (el) => {
        for (const ch of el.childNodes) {
            if (ch.nodeType === 3 && ch.nodeValue && ch.nodeValue.trim()) return true;
        }
        return false;
    };
    const MEDIA = new Set(['SVG','IMG','CANVAS','VIDEO']);
    const hidden = [];
    for (const el of slide.querySelectorAll('*')) {
        if (el === deco) continue;
        if (isAncestor(el)) continue;
        const otherDeco = el.hasAttribute('data-pptx-deco-id');
        const media = MEDIA.has(el.tagName.toUpperCase());
        if (otherDeco || media || hasDirectText(el)) {
            hidden.push([el, el.style.visibility, el.style.getPropertyPriority('visibility')]);
            // inline !important 才能 beat adapters 注入的 [data-anim]{visibility:visible!important}
            el.style.setProperty('visibility', 'hidden', 'important');
        }
    }
    window.__pptx_deco_hidden = hidden;
    // deco 自身含直接文本时，临时清 color/text-shadow 避免烘进 PNG
    window.__pptx_deco_self = null;
    if (hasDirectText(deco)) {
        window.__pptx_deco_self = {
            el: deco,
            color: deco.style.color,
            shadow: deco.style.textShadow
        };
        deco.style.setProperty('color', 'transparent', 'important');
        deco.style.setProperty('text-shadow', 'none', 'important');
    }
}"""

_DECO_RESTORE_FOREGROUND_JS = r"""() => {
    for (const item of (window.__pptx_deco_hidden || [])) {
        const [el, v, prio] = item;
        el.style.removeProperty('visibility');
        if (v) el.style.setProperty('visibility', v, prio || '');
    }
    window.__pptx_deco_hidden = null;
    const s = window.__pptx_deco_self;
    if (s) {
        s.el.style.removeProperty('color');
        s.el.style.removeProperty('text-shadow');
        if (s.color) s.el.style.color = s.color;
        if (s.shadow) s.el.style.textShadow = s.shadow;
    }
    window.__pptx_deco_self = null;
}"""

_SVG_HIDE_SIBLINGS_JS = r"""(marker) => {
    const svg = document.querySelector(`[data-pptx-svg-id='${marker}']`);
    if (!svg) return;
    const parent = svg.parentElement;
    if (!parent) return;
    window.__pptx_hidden = [];
    for (const ch of parent.children) {
        if (ch === svg) continue;
        window.__pptx_hidden.push([ch, ch.style.visibility]);
        ch.style.visibility = 'hidden';
    }
}"""

_SVG_RESTORE_SIBLINGS_JS = r"""() => {
    for (const [ch, v] of (window.__pptx_hidden || [])) ch.style.visibility = v;
    window.__pptx_hidden = [];
}"""


# (kind, attr, omit_bg, pre_js, post_js)
_MARKER_SHOOT_SPECS = {
    "deco_snapshot": ("data-pptx-deco-id", False, _DECO_HIDE_FOREGROUND_JS, _DECO_RESTORE_FOREGROUND_JS),
    "svg":           ("data-pptx-svg-id",  True,  _SVG_HIDE_SIBLINGS_JS,    _SVG_RESTORE_SIBLINGS_JS),
    "canvas":        ("data-pptx-canvas-id", False, None, None),
    "img":           ("data-pptx-img-id",  True,  None, None),
}


def _shoot_marker_records(page, records, out_dir: Path):
    """统一处理 deco_snapshot / svg / canvas / img 四类 marker 截图。

    各类型差异封装在 _MARKER_SHOOT_SPECS：
    - deco/svg 截图前后需要 JS 隐藏 / 恢复前景或兄弟节点
    - canvas/img 直接截图，无前后处理
    截图成功的 record 写入 rec["screenshot"]；失败的 print warning，rec 不变。
    """
    for rec in records:
        kind = rec.get("kind")
        spec = _MARKER_SHOOT_SPECS.get(kind)
        if spec is None or not rec.get("marker"):
            continue
        attr, omit_bg, pre_js, post_js = spec
        marker = rec["marker"]
        sel = f"[{attr}='{marker}']"
        out_png = out_dir / f"{marker}.png"
        try:
            if pre_js is not None:
                page.evaluate(pre_js, marker)
            page.locator(sel).screenshot(path=str(out_png), omit_background=omit_bg)
            rec["screenshot"] = str(out_png)
            if post_js is not None:
                page.evaluate(post_js)
        except Exception as e:
            print(f"    [warn] {kind} shoot fail {marker}: {e}")


def measure(html_path: Path, out_json: Path | None = None, *,
            single_index: int | None = None, no_screenshots: bool = True,
            verbose: bool = True) -> dict:
    """实测 HTML 中所有 slide。返回 measurement dict。
    out_json 不为 None 时同步写盘；svg 截图始终落盘到 out_json 旁的 _svg_assets/。
    """
    html_path = Path(html_path).resolve()
    url = html_path.as_uri()

    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport=VIEWPORT, device_scale_factor=1)
        page = ctx.new_page()
        page.goto(url, wait_until="networkidle")
        # custom-element upgrade（<deck-stage> 等）有时还没 settle，多等一拍
        page.evaluate("() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))")

        # canvas 内容（Chart.js / WebGL / 自绘图）的入场动画用 JS rAF 驱动，
        # 不受 CSS animation kill 影响。这里用通用稳定性检测：
        # 1. 对所有 canvas 取像素 hash，等到连续两次 hash 一致即认为稳定
        # 2. 最多等 2.0s（覆盖 Chart.js 默认 1000ms + 安全余量），不针对任何特定库
        # 不再硬编码 window.Chart 等库名 — 任何用 rAF 渲染的库都受益
        has_canvas = page.evaluate("() => document.querySelectorAll('canvas').length > 0")
        if has_canvas:
            page.wait_for_function(r"""
                () => {
                    const cans = document.querySelectorAll('canvas');
                    if (!cans.length) return true;
                    const sample = (cv) => {
                        // 抽样像素 hash，避免 readback 大数据
                        try {
                            const ctx = cv.getContext('2d', { willReadFrequently: true });
                            if (!ctx) return cv.width + 'x' + cv.height;
                            const w = Math.max(1, Math.min(cv.width, 32));
                            const h = Math.max(1, Math.min(cv.height, 32));
                            const data = ctx.getImageData(0, 0, w, h).data;
                            let acc = 0;
                            for (let i = 0; i < data.length; i += 16) acc = (acc * 31 + data[i]) | 0;
                            return acc.toString();
                        } catch (e) {
                            // WebGL canvas getContext('2d') 会 fail；用 toDataURL 短前缀
                            try { return cv.toDataURL().slice(0, 64); } catch (_) { return ''; }
                        }
                    };
                    const snap = Array.from(cans).map(sample).join('|');
                    window.__pptxCanvasSnap = window.__pptxCanvasSnap || '';
                    const prev = window.__pptxCanvasSnap;
                    window.__pptxCanvasSnap = snap;
                    return prev === snap && snap !== '';
                }
            """, timeout=2000)

        # 一次性准备：disable 动画 / 注入 force-position CSS / 跑 slide 发现
        from adapters import PREPARE_JS, ENUMERATE_JS, ACTIVATE_JS
        page.evaluate(PREPARE_JS)
        page.evaluate("() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))")

        total = page.evaluate(ENUMERATE_JS)
        if verbose:
            print(f"[measure] 共 {total} 张 slide")

        # svg / 参考截图的落盘位置：依附于 out_json；无 out_json 时落到临时目录
        if out_json is not None:
            anchor = Path(out_json)
            anchor.parent.mkdir(parents=True, exist_ok=True)
        else:
            anchor = Path(tempfile.mkdtemp(prefix="h2p_meas_")) / "measurements.json"

        screenshots_dir = Path(str(anchor.with_suffix("")) + "_screenshots")
        if not no_screenshots:
            screenshots_dir.mkdir(exist_ok=True, parents=True)

        svg_dir = anchor.parent / (anchor.stem + "_svg_assets")
        svg_dir.mkdir(parents=True, exist_ok=True)

        if single_index is not None:
            indices = [single_index]
        else:
            indices = list(range(total))

        slides_data = []
        for i in indices:
            page.evaluate(ACTIVATE_JS, i)
            # 等一帧让 .active 类等切换后的 computed style / transform 生效
            page.evaluate("() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))")

            # counter 动画（[data-target] 元素从 0 渐变到 dataset.target）通用模式：
            # 等到所有 counter 的 textContent 包含其 dataset.target 数值字串再抓。
            # 不等的话会拿到动画中段帧（如 "1+ 0+ 0+ 3" 而非终值 "42+ 20+ 7+ 100"）。
            # 没有 counter 的 slide 立刻 return true，不影响速度。
            has_counter = page.evaluate("() => document.querySelectorAll('[data-target]').length > 0")
            if has_counter:
                try:
                    page.wait_for_function(
                        """() => Array.from(document.querySelectorAll('[data-target]')).every(c => {
                            const t = c.dataset.target;
                            return t && c.textContent.includes(t);
                        })""",
                        timeout=2500,
                    )
                except Exception:
                    # IntersectionObserver 没触发或别的原因没跑完——强制设终值兜底
                    page.evaluate("""() => {
                        for (const c of document.querySelectorAll('[data-target]')) {
                            if (c.dataset.target) c.textContent = c.dataset.target;
                        }
                    }""")
            data = page.evaluate(EXTRACT_JS, i)

            # 统一截图四类 marker 元素：deco_snapshot / svg / canvas / img
            # 类型差异封装在 _MARKER_SHOOT_SPECS（pre/post JS、omit_background）
            _shoot_marker_records(page, data.get("records", []), svg_dir)

            slides_data.append(data)
            if verbose:
                if no_screenshots:
                    print(f"  slide {i+1:02d}: {len(data.get('records', []))} records")
                else:
                    ss = screenshots_dir / f"slide_{i+1:02d}.png"
                    # 当前 active slide 总是带 data-pptx-target；不依赖 adapter selector
                    page.locator("[data-pptx-target]").first.screenshot(path=str(ss))
                    print(f"  slide {i+1:02d}: {len(data.get('records', []))} records → {ss.name}")

        if single_index is not None:
            payload = slides_data[0]
        else:
            payload = {"slides": slides_data}

        if out_json is not None:
            Path(out_json).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            if verbose:
                print(f"wrote {out_json}")
        browser.close()

    return payload


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    no_screenshots = "--no-screenshots" in flags
    if len(args) < 2:
        print(__doc__)
        sys.exit(1)
    html_path = Path(args[0]).resolve()
    out_json = Path(args[1]).resolve()
    single_index = int(args[2]) if len(args) >= 3 else None
    measure(html_path, out_json, single_index=single_index, no_screenshots=no_screenshots)


if __name__ == "__main__":
    main()
