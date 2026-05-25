"""visual_audit.py — Stage 5b 视觉审计物料。

产出 `<out>_audit/`：
- `slide_NN_compare.png` × N — HTML | PPT 双栏拼图
- `audit_index.json` — 每页元数据 + 结构化告警
- `audit_prompt.md` — 给上游 VLM agent 的审计指南
"""
import json
from pathlib import Path


AUDIT_PROMPT_MD = """# Visual Audit

逐页视觉对比 HTML 参考图与 PPT 输出图。本文件专注 sub-agent 调用细节和 findings 格式；**完整工作流（修复纪律 / sticky 规则 / 增量重跑 / 终止条件）见 SKILL.md "工作流" 与 "修复纪律" 章节**。

起手：读 `audit_index.json` 拿页清单。`incremental_mode=true` 时只看 `fresh_indices` 列出的页。

## 并行执行（强制，不论页数）

按 SKILL.md "并行 audit" 页数策略拆 batch：≤4 页 = 1 个 batch；5-20 页 = 每 batch 4 页；>20 页 = 每 batch 4-5 页。即使只有 1 页也走 sub-agent，主 agent 不常规自己 Read compare 图。

Sub-agent 调用模板（Claude `Agent(...)` 风格；每个 Agent 一个 batch，全部塞在主 agent 同一条 message 里才并行）：

```
Agent(
  description="Audit slides N1-N2",
  subagent_type="general-purpose",
  run_in_background=True,
  prompt='''你是视觉审计员。看以下这批 compare 图（左=HTML 参考，右=PPT 输出），按顺序逐张应用检查清单。

本批图（每张是一页）：
- <full path to slide_N1_compare.png>
- <full path to slide_N2_compare.png>
- <full path to slide_N3_compare.png>
- <full path to slide_N4_compare.png>

可读参考（只读，不编辑；不要假设主 agent 的 Read/Grep 输出会自动传给你）：
- deck HTML: <full path to template.audited.html or template.html>
- lessons learned: <skill_dir>/references/lessons-learned.md

仅审本批 compare 图，不要读其它页的 compare 图，不要改代码/HTML。只有当定位同一页元素需要时，才只读上面的 HTML / lessons 文件。

**看图须知（避免常见误判）**：
- 每张图是**左右拼图**：左半 = HTML 参考渲染，右半 = PPT 输出，两半之间有窄分隔留白
- 顶部标题栏（"HTML 参考 / PPT 输出 / slide NN"）和中间灰色分隔线是审计 UI，**不是 slide 内容**，不要报告这些区域的问题
- 判断居中 / 对齐 / 偏移**只看每一半内部**的相对位置——不要跨左右半比较 x 坐标
- 例：卡片在 PPT 半图的中部、HTML 半图也在中部 → PPT 和 HTML 都居中 ✓；**不能**因为卡片在整张拼图的 x≈2880 就说"内容右移"
- 同理判断字号 / 间距 / 段宽时，只比"HTML 半图里的元素 vs PPT 半图里**同一**元素"，不要拿一半的元素去和另一半别的元素比
- 只报告 PPT 半图相对 HTML 半图**新增或放大**的视觉问题；如果 HTML 半图本身也有同样溢出 / 重叠 / 错位，不要作为转换 finding 报告
- **报 finding 前必须先看 HTML 半图、描述它的实际像素态**（颜色 / 填充 / 描边 / 形状），再讲 PPT 半图差异。写不出 HTML 半图实际像素 = 没看清 = 不报。不要靠"HTML 应该是 X 因为它叫 .filled / 命名暗示 X"脑补——只看左半图实际像素值。
  - ✅ 报："HTML 半图 step 2 是浅米色填充 + 深色数字，PPT 半图 step 2 是深色填充 + 浅色数字"（描述了 HTML 半图实际颜色）
  - ❌ 不报："HTML 里 5 个圆都深色，PPT 把 2/4 改成浅色"（只说"HTML 里"，没描述左半图像素，是脑补）
- **严重度量化阈值**（不允许"显著 / 明显 / 大幅"等模糊词单独构成 HIGH）：
  - **HIGH**：装饰完全丢失（HTML 有 PPT 没）/ 完全错形（圆变方块、有色变黑等取代级）/ 文字被遮 / 文字溢出 slide / 字号差 ≥ 50% / 颜色取反
  - **MID**：位置 / 尺寸 / 字号差 20-50%，颜色明显偏移但同色系
  - **LOW**：差异 < 20%、原设计本身允许的微旋转 / 微 gap、可接受的设计美化

检查清单（按重要度排序，每页都过一遍）：
1. 文字被线条 / 形状边界 / 图片角穿过 / 覆盖
2. 文字之间不该有的重叠 / 叠压
3. 文字溢出 slide 边界 / 被裁切 / 溢入相邻列
4. 元素相对 HTML 参考图大幅错位
5. 字体回退 / 字号变形
6. 图片拉伸 / 错位 / 缺失，装饰色块变形
7. 颜色错误（明显偏离 HTML）

输出**纯文本**（不要 markdown code fence 包装）：为本批每一页输出一个块，严格用以下格式：

有问题的页：
## page NN
- [HIGH] <稳定元素短名>：HTML 半图 <实际状态>；PPT 半图 <差异>
- [MID]  <稳定元素短名>：HTML 半图 <实际状态>；PPT 半图 <差异>
- [LOW]  <稳定元素短名>：HTML 半图 <实际状态>；PPT 半图 <差异>

无问题的页：
## page NN · OK

HIGH=用户一眼能看出 / MID=细看才发现 / LOW=设计美化建议。**每页都必须有一个块（OK 或带 finding），不要漏页**。每条 finding 必须点名稳定元素短名，方便主 agent 做 sticky key；不要写原因猜测、修复方案或总结。'''
)
```

sub-agent 只返回 findings 文本，**不要让它直接写 audit_findings.md**——并发写会互相覆盖，主 agent 收齐后统一合并。

同一条 message 里跟着发并行准备：

```
Read(file_path="<deck>/template.html")
Read(file_path="<skill_dir>/references/lessons-learned.md")
Grep(pattern='class="slide |data-slide=', path="<deck>/template.html",
     output_mode="content", -n=true)
```

只发"无论 findings 是什么都用得上"的准备，不要预测 findings 提前改 HTML。详见 SKILL.md "与 sub-agent 并行的主 agent 准备"。

## findings 格式

```
## page 08
- [HIGH] 右列 lead 字：HTML 半图 lead 字完整未被穿过；PPT 半图 border-left 竖线穿过"实"字
- [LOW] 副标题间距：HTML 半图副标题与正文留白约 24px；PPT 半图留白约 14px

## page 09 · OK
```

- **HIGH**：用户一眼能看出，必须修
- **MID**：细看才发现，强烈建议修
- **LOW**：设计美化建议，可选
"""


def build_compare_image(html_png: Path, ppt_png: Path, out_path: Path, page_idx: int):
    """生成单页 HTML | PPT 双栏拼图。"""
    from PIL import Image, ImageDraw, ImageFont
    # 跨平台 title 字体兜底：arial=Windows、DejaVuSans=Linux/PIL bundled、Helvetica=macOS。
    # 都找不到时 load_default() 是 bitmap，36pt 显示效果差但不阻塞 audit。
    title_font = None
    for name in ("arial.ttf", "DejaVuSans.ttf", "Helvetica.ttc"):
        try:
            title_font = ImageFont.truetype(name, 36)
            break
        except Exception:
            continue
    if title_font is None:
        title_font = ImageFont.load_default()
    try:
        html_img = Image.open(html_png).convert("RGB").resize((1920, 1080))
        ppt_img = Image.open(ppt_png).convert("RGB").resize((1920, 1080))
    except Exception as e:
        print(f"  [warn] compare build fail page {page_idx}: {e}")
        return None

    bar_h = 60
    composite = Image.new("RGB", (1920 * 2 + 8, 1080 + bar_h), (255, 255, 255))
    d = ImageDraw.Draw(composite)
    # 标题栏
    d.rectangle((0, 0, 1920, bar_h), fill=(245, 245, 247))
    d.rectangle((1928, 0, 3848, bar_h), fill=(255, 245, 235))
    d.text((28, 12), f"HTML 参考  ·  slide {page_idx:02d}", fill=(20, 20, 20), font=title_font)
    d.text((1956, 12), f"PPT 输出  ·  slide {page_idx:02d}", fill=(20, 20, 20), font=title_font)
    # 中间分隔
    d.rectangle((1920, 0, 1928, 1080 + bar_h), fill=(200, 200, 200))
    composite.paste(html_img, (0, bar_h))
    composite.paste(ppt_img, (1928, bar_h))
    composite.save(out_path, optimize=True)
    return out_path


def build_audit_package(pptx_path: Path, html_screenshots_dir: Path, ppt_screenshots_dir: Path,
                        self_check_result: dict, preflight_result: dict | None,
                        out_dir: Path,
                        only_indices: set[int] | None = None) -> dict:
    """产出 audit 物料包：compare 图 × N + audit_index.json + audit_prompt.md。

    only_indices 给定时走增量：只对列出的页重建 compare 图，其它页保留 out_dir 里上轮的
    slide_NN_compare.png（缓存缺失则兜底重建）。audit_index.json 标记每页 fresh=true/false，
    上游 agent 看 fresh_indices 决定本轮要复审哪些页。

    返回 dict 描述包内容（方便 convert.py 在终端打印）。
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    html_pngs = sorted(html_screenshots_dir.glob("slide_*.png"))
    ppt_pngs = sorted(ppt_screenshots_dir.glob("slide_*.png"))
    n = min(len(html_pngs), len(ppt_pngs))

    pages_meta = []
    pages_by_idx = {p["idx"]: p for p in self_check_result.get("pages", [])}
    preflight_by_idx = {s["index"]: s for s in (preflight_result or {}).get("slides", [])}

    fresh_set: set[int] = set()
    skipped_set: set[int] = set()
    for i in range(n):
        idx = i + 1
        compare_path = out_dir / f"slide_{idx:02d}_compare.png"
        must_rebuild = (only_indices is None
                        or idx in only_indices
                        or not compare_path.exists())
        if must_rebuild:
            build_compare_image(html_pngs[i], ppt_pngs[i], compare_path, idx)
            fresh_set.add(idx)
        else:
            skipped_set.add(idx)
        page_info = pages_by_idx.get(idx, {})
        preflight_info = preflight_by_idx.get(idx, {})
        risks = [r["code"] for r in preflight_info.get("risks", [])]
        pages_meta.append({
            "index": idx,
            "compare_image": str(compare_path.name),
            "html_screenshot": str(html_pngs[i].name),
            "ppt_screenshot": str(ppt_pngs[i].name),
            "structural_level": page_info.get("level"),
            "preflight_risks": risks,
            "preflight_confidence": preflight_info.get("confidence"),
            "fresh": idx in fresh_set,
        })

    index_data = {
        "pptx": str(pptx_path.name),
        "pptx_path": str(pptx_path),
        "total_pages": n,
        "instructions_file": "audit_prompt.md",
        "findings_output": "audit_findings.md",
        "incremental_mode": only_indices is not None,
        "fresh_indices": sorted(fresh_set),
        "cached_indices": sorted(skipped_set),
        "pages": pages_meta,
        "self_check_summary": {
            "engine": self_check_result.get("engine"),
            "structural_warnings_count": len(self_check_result.get("warnings", [])),
        },
    }

    (out_dir / "audit_index.json").write_text(
        json.dumps(index_data, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "audit_prompt.md").write_text(AUDIT_PROMPT_MD, encoding="utf-8")

    return {
        "out_dir": str(out_dir),
        "pages": n,
        "fresh": sorted(fresh_set),
        "cached": sorted(skipped_set),
        "incremental": only_indices is not None,
        "index": str(out_dir / "audit_index.json"),
        "prompt": str(out_dir / "audit_prompt.md"),
    }
