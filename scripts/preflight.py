"""preflight.py — 输入识别 / 风险预扫（流水线 Stage 1）。

在 measure / assemble / self-check 之前先跑一遍：
- 打开 HTML，识别每张 slide 上**已知会踩坑**的 CSS 模式
- 输出结构化 preflight.json + 终端摘要
- 高风险 slide 会被传给 self_check.py 强制纳入"需人工 WPS 复核"列表

只识别静态可预判的模式；不替代 measure 的实际抽取。

风险码（risk codes）：
- R001 deco-on-slide-root         slide 根节点带 deco_snapshot 触发模式（最常见双层文字 bug）
- R002 multi-layer-text-shadow    任意元素 text-shadow 含 2+ 阴影（OOXML 只能表达单 outerShdw）
- R003 backdrop-filter            已自动走 deco_snapshot 兜底；保留信息性提示，不算阻塞风险
- R004 video-element              <video> 帧不会转换
- R005 webgl-or-canvas            canvas/WebGL 只能取静态帧
- R006 tight-line-height          line-height < 0.9 的多行文本（PPT 行距会偏松）
- R007 chinese-italic             CJK 字符使用 font-style: italic（PPT 会回退正体）
- R008 external-font              HTML 用到外部字体家族（font_resolver 会尝试从 GF 拉，命中即嵌入；GF 没有的就回退到 viewer 系统字体）
- R009 large-rotated-band         占整页大色带 + 旋转/倾斜（几何与截图边界情况）

每条风险有 severity: high / medium / low。
slide.confidence 由该 slide 的最高 severity 决定（high → low confidence）。
"""
import argparse
import json
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

from _js_snippets import DECO_HELPERS
from text_utils import CJK_JS_RANGE

VIEWPORT = {"width": 1920, "height": 1080}

# CSS 通用关键字 + font_resolver.VARIABLE_RECIPES 里 CJK 直链家族（preflight 时 FONT_PLAN
# 还是空的，但这些家族 convert 会自动种子，不算"外部字体"）
from font_resolver import VARIABLE_RECIPES
KNOWN_AUTO_RESOLVED = {k for k in VARIABLE_RECIPES}  # 已是小写
KNOWN_AUTO_RESOLVED |= {a for r in VARIABLE_RECIPES.values() for a in r.get("aliases", [])}
GENERIC_FAMILIES = {
    "sans-serif", "serif", "monospace", "system-ui", "ui-sans-serif",
    "ui-serif", "ui-monospace", "cursive", "fantasy", "math", "emoji",
    "inherit", "initial", "unset", "revert",
}

SEVERITY_RANK = {"high": 3, "medium": 2, "low": 1}


SCAN_JS = r"""
() => {
  const slide = document.querySelector('[data-pptx-target]');
  if (!slide) return { error: 'no slide tagged with data-pptx-target' };

  slide.scrollIntoView({block:'start', inline:'start', behavior:'instant'});
  const slideRect = slide.getBoundingClientRect();
  const slideArea = Math.max(1, slideRect.width * slideRect.height);

  const risks = [];
  const css = (el, prop) => getComputedStyle(el).getPropertyValue(prop);

  const describe = (el) => {
    const id = el.id ? '#' + el.id : '';
    const cls = (el.className && typeof el.className === 'string')
      ? '.' + el.className.trim().split(/\s+/).slice(0, 3).join('.')
      : '';
    return (el.tagName.toLowerCase() + id + cls).slice(0, 80);
  };

""" + DECO_HELPERS + r"""

  // R001 用的 "命中 deco_snapshot 的具体原因"：与 hasComplexDecoration 必须保持完全同步
  // （hasComplexDecoration 只返 bool；这里把命中的项拆开给 risk.detail 描述）
  const decoReasons = (el) => {
    const s = getComputedStyle(el);
    const r = [];
    if (s.backgroundImage && s.backgroundImage !== 'none') r.push('background-image');
    if (s.boxShadow && s.boxShadow !== 'none') r.push('box-shadow');
    if (s.outlineStyle && s.outlineStyle !== 'none' && parseFloat(s.outlineWidth) > 0) r.push('outline');
    if (hasRasterPseudoDecoration(el, '::before')) r.push('::before');
    if (hasRasterPseudoDecoration(el, '::after')) r.push('::after');
    if (isClippingContainerWithTransformedChildren(s, el)) r.push('clip-with-transformed-children');
    return r;
  };

  // R001 slide 范围内 ≥50% 覆盖的 deco_snapshot + 有文字共存 → HIGH
  //     (覆盖 [[project-html-to-pptx-deco-snapshot-bug]] 已知触发条件)
  //     slide 根有 deco 也算 HIGH（一定有文字共存）
  //     <50% 覆盖的小卡片 deco → 不报，避免误伤
  {
    const rootReasons = decoReasons(slide);
    if (rootReasons.length) {
      risks.push({
        code: 'R001',
        name: 'deco-on-slide-root',
        severity: 'high',
        detail: rootReasons.join(' + '),
        where: describe(slide),
      });
    }
    // 再扫子元素：找大覆盖 deco
    for (const el of slide.querySelectorAll('*')) {
      if (el === slide) continue;
      const sty = getComputedStyle(el);
      if (sty.display === 'none' || sty.visibility === 'hidden') continue;
      const r = el.getBoundingClientRect();
      if (r.width === 0 || r.height === 0) continue;
      const area = r.width * r.height;
      if (area / slideArea < 0.5) continue;
      const reasons = decoReasons(el);
      if (!reasons.length) continue;

      // 有没有"文字兄弟"（同父或后代里有其他含直接文字的元素）
      const hasTextNeighbor = (() => {
        for (const cand of slide.querySelectorAll('*')) {
          if (cand === el) continue;
          if (el.contains(cand) || cand.contains(el)) continue;
          for (const ch of cand.childNodes) {
            if (ch.nodeType === 3 && ch.nodeValue && ch.nodeValue.trim()) return true;
          }
        }
        return false;
      })();
      if (hasTextNeighbor) {
        risks.push({
          code: 'R001',
          name: 'large-deco-with-text-neighbor',
          severity: 'high',
          detail: `覆盖 ${(area/slideArea*100).toFixed(0)}% · ${reasons.join(' + ')} · 与文字兄弟共存`,
          where: describe(el),
        });
        break;  // 一页一条
      }
    }
  }

  // 收集所有可见元素一次扫描
  const allEls = Array.from(slide.querySelectorAll('*'));
  const fontFamilies = new Set();
  let textShadowCount = 0;
  let backdropCount = 0;
  let videoCount = 0;
  let canvasCount = 0;
  let chineseItalicCount = 0;
  let tightLhCount = 0;
  let rotatedBandCount = 0;
  const cjkRe = /""" + CJK_JS_RANGE + r"""/;

  for (const el of allEls) {
    const s = getComputedStyle(el);
    if (s.display === 'none' || s.visibility === 'hidden') continue;
    const r = el.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) continue;

    // R002 多层 text-shadow（在元素含直接文字时才计入）
    const ts = s.textShadow;
    if (ts && ts !== 'none') {
      // 简单按 rgb()/rgba()/颜色关键字数量数 shadow 段
      const segs = ts.split(/,(?![^()]*\))/).filter(Boolean);
      if (segs.length >= 2) {
        let hasDirectText = false;
        for (const ch of el.childNodes) {
          if (ch.nodeType === 3 && ch.nodeValue && ch.nodeValue.trim()) {
            hasDirectText = true;
            break;
          }
        }
        if (hasDirectText) {
          textShadowCount++;
          if (textShadowCount === 1) {
            risks.push({
              code: 'R002',
              name: 'multi-layer-text-shadow',
              severity: 'high',
              detail: `${segs.length} 层阴影 (${ts.slice(0, 80)}...)`,
              where: describe(el),
            });
          }
        }
      }
    }

    // R003 backdrop-filter — 已自动走 deco_snapshot 兜底（_js_snippets.hasComplexDecoration）
    // 保留为信息性提示，severity=low；audit agent 不应据此做 HTML workaround
    if (s.backdropFilter && s.backdropFilter !== 'none') {
      backdropCount++;
      if (backdropCount === 1) {
        risks.push({
          code: 'R003',
          name: 'backdrop-filter',
          severity: 'low',
          detail: s.backdropFilter.slice(0, 80),
          where: describe(el),
        });
      }
    }

    // R004 video
    if (el.tagName === 'VIDEO') {
      videoCount++;
      if (videoCount === 1) {
        risks.push({
          code: 'R004',
          name: 'video-element',
          severity: 'high',
          detail: '<video> 帧不会转换',
          where: describe(el),
        });
      }
    }

    // R005 canvas（可能是 WebGL/Chart.js，至少能取静态帧）
    if (el.tagName === 'CANVAS') {
      canvasCount++;
      if (canvasCount === 1) {
        let kind = 'canvas';
        try {
          if (el.getContext('webgl') || el.getContext('webgl2')) kind = 'WebGL canvas';
        } catch (e) {}
        risks.push({
          code: 'R005',
          name: 'webgl-or-canvas',
          severity: 'low',
          detail: `${kind}（仅静态帧）`,
          where: describe(el),
        });
      }
    }

    // R006 紧 line-height（多行文本）
    const lh = parseFloat(s.lineHeight);
    const fs = parseFloat(s.fontSize);
    if (lh && fs && lh / fs < 0.9 && r.height > fs * 1.4) {
      // 行高比 < 0.9 且高度大于 ~1.4 行（多行）
      let hasDirectText = false;
      for (const ch of el.childNodes) {
        if (ch.nodeType === 3 && ch.nodeValue && ch.nodeValue.trim()) {
          hasDirectText = true;
          break;
        }
      }
      if (hasDirectText) {
        tightLhCount++;
        if (tightLhCount === 1) {
          risks.push({
            code: 'R006',
            name: 'tight-line-height',
            severity: 'low',
            detail: `line-height/font-size ≈ ${(lh/fs).toFixed(2)}`,
            where: describe(el),
          });
        }
      }
    }

    // R007 中文斜体
    if (s.fontStyle === 'italic') {
      let txt = '';
      for (const ch of el.childNodes) {
        if (ch.nodeType === 3 && ch.nodeValue) txt += ch.nodeValue;
      }
      if (cjkRe.test(txt)) {
        chineseItalicCount++;
        if (chineseItalicCount === 1) {
          risks.push({
            code: 'R007',
            name: 'chinese-italic',
            severity: 'medium',
            detail: 'CJK 字符 + font-style:italic（PPT 会显示正体）',
            where: describe(el),
          });
        }
      }
    }

    // R008 收集字体家族（最后统一判断）
    const ff = s.fontFamily;
    if (ff) {
      ff.split(',').forEach(name => {
        const cleaned = name.trim().replace(/^["']|["']$/g, '').toLowerCase();
        if (cleaned) fontFamilies.add(cleaned);
      });
    }

    // R009 大旋转色带：占 slide 面积 > 25% + 含 transform rotate/skew
    const t = s.transform;
    if (t && t !== 'none' && /matrix/.test(t)) {
      const m = t.match(/matrix(?:3d)?\(([^)]+)\)/);
      if (m) {
        const v = m[1].split(',').map(parseFloat);
        const angle = Math.atan2(v[1], v[0]) * 180 / Math.PI;
        const area = r.width * r.height;
        if (Math.abs(angle) > 5 && area / slideArea > 0.25) {
          const bg = s.backgroundColor;
          const hasBg = bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent';
          const hasBgImg = s.backgroundImage && s.backgroundImage !== 'none';
          if (hasBg || hasBgImg) {
            rotatedBandCount++;
            if (rotatedBandCount === 1) {
              risks.push({
                code: 'R009',
                name: 'large-rotated-band',
                severity: 'medium',
                detail: `面积 ${(area/slideArea*100).toFixed(0)}% / 旋转 ${angle.toFixed(0)}°`,
                where: describe(el),
              });
            }
          }
        }
      }
    }
  }

  return {
    ok: true,
    theme: slide.className,
    risks,
    fonts: Array.from(fontFamilies),
    counts: {
      text_shadow_multi: textShadowCount,
      backdrop_filter: backdropCount,
      video: videoCount,
      canvas: canvasCount,
      chinese_italic: chineseItalicCount,
      tight_line_height: tightLhCount,
      rotated_band: rotatedBandCount,
    },
  };
}
"""


def _confidence_from_risks(risks):
    if not risks:
        return "high"
    top = max(SEVERITY_RANK[r["severity"]] for r in risks)
    if top >= SEVERITY_RANK["high"]:
        return "low"
    if top >= SEVERITY_RANK["medium"]:
        return "medium"
    return "high"


def preflight(html_path: Path, out_json: Path | None = None,
              verbose: bool = True) -> dict:
    """Stage 1：扫风险。返回 dict（也会写到 out_json 如指定）。"""
    html_path = Path(html_path).resolve()
    url = html_path.as_uri()

    result = {
        "html_path": str(html_path),
        "slide_count": 0,
        "slides": [],
        "global_fonts": [],
        "summary": {
            "total_risks": 0,
            "by_severity": {"high": 0, "medium": 0, "low": 0},
            "manual_review_slides": [],
            "manual_review_reason": {},
        },
    }

    t0 = time.perf_counter()
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport=VIEWPORT, device_scale_factor=1)
        page = ctx.new_page()
        page.goto(url, wait_until="networkidle")
        page.evaluate(
            "() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))"
        )

        from adapters import PREPARE_JS, ENUMERATE_JS, ACTIVATE_JS
        page.evaluate(PREPARE_JS)
        page.evaluate(
            "() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))"
        )
        total = page.evaluate(ENUMERATE_JS)
        result["slide_count"] = total

        global_fonts = set()
        for i in range(total):
            page.evaluate(ACTIVATE_JS, i)
            page.evaluate(
                "() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))"
            )
            data = page.evaluate(SCAN_JS)
            if data.get("error"):
                result["slides"].append({
                    "index": i + 1, "confidence": "low",
                    "error": data["error"], "risks": [],
                })
                continue

            risks = data.get("risks", [])
            for r in risks:
                result["summary"]["total_risks"] += 1
                result["summary"]["by_severity"][r["severity"]] += 1

            for fam in data.get("fonts", []):
                global_fonts.add(fam)

            confidence = _confidence_from_risks(risks)
            slide_rec = {
                "index": i + 1,
                "theme": data.get("theme", ""),
                "confidence": confidence,
                "risks": risks,
                "counts": data.get("counts", {}),
            }
            result["slides"].append(slide_rec)

            if confidence == "low":
                result["summary"]["manual_review_slides"].append(i + 1)
                result["summary"]["manual_review_reason"][str(i + 1)] = [
                    r["code"] for r in risks if r["severity"] == "high"
                ]

        # R008 全局：外部字体（信息性，font_resolver 会试图从 GF 拉）
        external = sorted(
            f for f in global_fonts
            if f and f.lower() not in GENERIC_FAMILIES
            and f.lower() not in KNOWN_AUTO_RESOLVED
            and not f.startswith('-')
        )
        result["global_fonts"] = sorted(global_fonts)
        if external:
            result["summary"]["external_fonts"] = external
            result["summary"]["by_severity"]["low"] += 1
            result["summary"]["total_risks"] += 1

        browser.close()

    if out_json is not None:
        Path(out_json).parent.mkdir(parents=True, exist_ok=True)
        Path(out_json).write_text(
            json.dumps(result, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    if verbose:
        _print_summary(result, time.perf_counter() - t0)

    return result


def _print_summary(result: dict, elapsed: float):
    s = result["summary"]
    print(f"\n[preflight] {result['slide_count']} 页  ·  {elapsed:.2f}s")
    print(f"[preflight] 风险 {s['total_risks']} 项  "
          f"(high={s['by_severity']['high']}, "
          f"medium={s['by_severity']['medium']}, "
          f"low={s['by_severity']['low']})")

    if s.get("external_fonts"):
        print(f"[preflight] 外部字体: {', '.join(s['external_fonts'])}  "
              f"→ 试图从 GF 自动拉取，未命中则回退 viewer 系统字体")

    if s["manual_review_slides"]:
        print(f"[preflight] !! {len(s['manual_review_slides'])} 页强制人工 WPS 复核: "
              f"{s['manual_review_slides']}")
        for slide in result["slides"]:
            if slide["confidence"] != "low":
                continue
            print(f"           第 {slide['index']:02d} 页 ({slide.get('theme','')[:40]})")
            for r in slide["risks"]:
                if r["severity"] == "high":
                    print(f"             [HIGH] {r['code']} {r['name']}: {r['detail']}  @ {r['where']}")
    else:
        print("[preflight] OK 无高风险页")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("input", help="HTML 路径")
    ap.add_argument("--out", help="preflight.json 输出路径")
    args = ap.parse_args()
    in_path = Path(args.input)
    out = Path(args.out) if args.out else in_path.with_name(in_path.stem + "_preflight.json")
    preflight(in_path, out, verbose=True)


if __name__ == "__main__":
    sys.dont_write_bytecode = True
    SCRIPTS = Path(__file__).parent
    sys.path.insert(0, str(SCRIPTS))
    main()
