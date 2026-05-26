"""visual_audit.py — Stage 5b 视觉审计物料。

产出 `<out>_audit/`：
- `slide_NN_compare.png` × N — HTML | PPT 双栏拼图
- `audit_index.json` — 每页元数据 + 结构化告警
- `audit_prompt.md` — 给上游 VLM agent 的审计指南
"""
import json
from pathlib import Path

from local_config import audit_mode


AUDIT_PROMPT_MD = """# Visual Audit

视觉对比 HTML 参考图与 PPT 输出图。本文件专注 sub-agent 调用细节和 findings 格式；**完整工作流（修复纪律 / sticky 规则 / 增量重跑 / 终止条件）见 SKILL.md "工作流" 与 "修复纪律" 章节**。

起手：读 `audit_index.json` 拿页清单、`audit_mode`、`contact_sheets`。`incremental_mode=true` 时只看 `fresh_indices` 列出的页。

## page/detail 审查执行

适用范围：`page` 模式所有目标页、`triage` 模式第 2 轮入围页（`manual` 模式不进入本节）。在适用范围内即使本批只有 1 页也走 sub-agent。

按 SKILL.md "audit 分发与 compare 图读取规则" 章节的页数策略拆 batch。Page/detail sub-agent 调用模板（Claude `Agent(...)` 风格；每个 Agent 一个 batch，全部塞在主 agent 同一条 message 里才并行）：

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

**默认 OK 偏置（最重要）**

这是 HTML → PPT 视觉转换审计，不是 pixel-perfect 还原审计。OOXML / PowerPoint / 浏览器之间必然有底层差异（字体 hinting、anti-aliasing、sub-pixel positioning、装饰栅格化、行宽细微差异），**这些不是转换 bug，不要报告**。

期望基线：大多数 deck 的**多数页**应该 `## page NN · OK`。如果你这批 4 页里有 ≥ 3 页有 MID/HIGH finding，几乎一定是你在过度解读细节——停下来重审，把"细看才发现"的全部删掉再交。

**不报告清单（HTML→PPT 自然差异 + 设计偏好，永不报告）**

下面这些是渲染底层自然差异或主观偏好，**永不报告**（哪怕降级到 LOW 也别报）：
- 字体 anti-aliasing / hinting / 字偶距 / 锯齿边 / sub-pixel 渲染差异
- 位置漂移 < 5 px / 装饰线条 / 边框相差几像素（形状、层级、视觉含义一致前提下）
- 字号 / 字距 / 行高 / 字重差异 < 10%（不同字体度量不同，10% 是正常容差）
- 字体 weight 看着略不同但字形一致
- 颜色饱和度 / 色温微差（同色系内）
- 文字行宽多换 1 行 / 标题换行位置不同（除非真溢出 slide 或真叠压邻元素）
- 装饰阴影 / blur / 柔和度的轻微差异
- 上标 / 下标 / 标点的基线漂移 < 5 px
- HTML 半图本身已模糊 / 挤压 / 重叠 / 看不清的区域
- 用"略粗 / 略细 / 略松 / 略紧 / 略偏"形容的无锚定差异
- 设计偏好（"我觉得这里更应该居中 / 更应该大一点"）

**看图须知（避免常见误判）**：
- 每张图是**左右拼图**：左半 = HTML 参考渲染，右半 = PPT 输出，两半之间有窄分隔留白
- 顶部标题栏（"HTML 参考 / PPT 输出 / slide NN"）和中间灰色分隔线是审计 UI，**不是 slide 内容**，不要报告这些区域的问题
- 判断居中 / 对齐 / 偏移**只看每一半内部**的相对位置——不要跨左右半比较 x 坐标
- 同理判断字号 / 间距 / 段宽时，只比"HTML 半图里的元素 vs PPT 半图里**同一**元素"
- 只报告 PPT 半图相对 HTML 半图**新增或放大**的视觉问题；HTML 半图本身已有的问题不算转换 finding
- **报 finding 前必须先看 HTML 半图、描述它的实际像素态**（颜色 / 填充 / 形状 / 位置），再讲 PPT 半图差异。写不出 HTML 半图实际像素 = 没看清 = 不报。不要靠"HTML 应该是 X 因为它叫 .filled / 命名暗示 X"脑补。

**严重度阈值**（不允许"显著 / 明显 / 大幅 / 略小"等模糊词单独构成级别）

- **HIGH**（用户瞄一眼就能看出来，必须修）：
  - 关键元素缺失或被错误形状 / 错误颜色取代
  - 文字被遮盖、裁切、溢出 slide，或压住其它正文
  - 字号 / 尺寸差 ≥ 50%
  - 颜色跨明暗关系或跨色相，影响阅读 / 含义

- **MID**（细看 5 秒能确认，**且必须能量化**）：
  - 字号 / 尺寸差 20-50%
  - 位置差 20-50 px
  - 颜色明显偏移但仍同色系
  - HTML 不叠压而 PPT 产生文字 / 元素叠压

- **LOW**（设计师 nice-to-have，可选；多数情况下应**不报**）：
  - 量化差异 10-20%，且确实影响局部观感
  - 不确定 LOW 还是不报时，**选不报**

检查清单（按重要度排序）：
1. 文字被线条 / 形状边界 / 图片角穿过 / 覆盖
2. 文字之间不该有的重叠 / 叠压（HTML 不叠压前提下）
3. 文字溢出 slide 边界 / 被裁切 / 溢入相邻列
4. 元素相对 HTML 参考图大幅错位（≥ 20 px）
5. 字体回退到完全不同的字形家族（如 serif → 等宽 / 衬线 → 无衬线）
6. 图片拉伸 / 错位 / 缺失，装饰色块变形
7. 颜色错误（取反 / 跨色系）

输出**纯文本**（不要 markdown code fence 包装）：为本批每一页输出一个块，严格用以下格式：

有问题的页：
## page NN
- [HIGH] <稳定元素短名>：HTML 半图 <实际状态>；PPT 半图 <差异 + 量化>
- [MID]  <稳定元素短名>：HTML 半图 <实际状态>；PPT 半图 <差异 + 量化>

无问题的页：
## page NN · OK

每条 finding 必须点名稳定元素短名，方便主 agent 做 sticky key；不要写原因猜测、修复方案或总结。

**输出前自检（必须）**

写好 findings 后，对每一条问自己 5 题（任一不过都删）：
1. 全页可见、且影响阅读 / 含义 / 布局？看不出 / 不影响 → 删
2. 落入上方"不报告清单"？是 → 删
3. HIGH：普通观众瞄一眼能看出？不能 → 降到 MID 或删
4. MID：细看 5 秒能确认且能量化 20-50%？不能 → 删
5. HTML 半图实际像素状态写清楚了吗（颜色 / 填充 / 形状 / 位置）？没有 → 可能在脑补 → 删

批量 sanity check：每 4 页 batch 通常 0-2 条 finding。超过 4 条 → 先假设自己过度审查，合并同类项并删掉所有不影响阅读 / 结构 / 含义的细节。'''
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

findings 格式与级别口径见上方 sub-agent prompt 模板（"严重度阈值" + "输出前自检"）。
"""


def build_audit_prompt(mode: str, contact_sheets: list[dict]) -> str:
    contact_lines = "\n".join(
        f"- {item['path']}  (pages {item['pages'][0]:02d}-{item['pages'][-1]:02d})"
        for item in contact_sheets
    ) or "- none"

    if mode == "page":
        mode_instructions = """## 当前审查模式：page

严格逐页审查。忽略 contact sheet 的省 token 分流作用，按 page/detail 模板把所有目标页的 `slide_NN_compare.png` 分 batch 交给 sub-agent。适合最终交付前、客户高风险 deck、或用户明确要求最高覆盖率。"""
    elif mode == "manual":
        mode_instructions = """## 当前审查模式：manual

只生成审查物料，不默认调用 sub-agent。把 contact sheet、单页 compare 图目录、`audit_index.json` 交给用户人工核查；用户点名反馈某页后，再打开对应 `slide_NN_compare.png` 修复。"""
    elif mode == "ask":
        mode_instructions = """## 当前审查模式：ask（未设置）

`.config.local.toml` 的 `[audit].mode` 是 `"ask"` 或缺失——主 agent 应该在 convert **之前** AskUserQuestion 收集用户偏好（triage / page / manual）并 Write config，本次执行没走那步。

请立即 AskUserQuestion 收集用户对本次审查的偏好（建议同时把答复写到 `.config.local.toml` 的 `[audit].mode`，下次会话不再问），然后按所选模式继续。三模式行为简述（详细规则见 SKILL.md "视觉 audit 模式" 章节）：

- 选 `triage` → 主 agent 自己读 `audit_contact_*.png` 分流 + sub-agent 看入围页单图
- 选 `page` → 所有目标页都派 sub-agent 看 `slide_NN_compare.png`
- 选 `manual` → 把 audit 目录交给用户人工核查

收到答复前**不要**派 sub-agent，也不要把 contact sheet 当成最终证据。"""
    else:
        mode_instructions = """## 当前审查模式：triage

默认省 token 审查。**第 1 轮 contact pass 只做分流，第 2 轮 detail pass 才做最终 finding**。

## 第 1 轮：contact pass（主 agent 自己执行）

直接 Read 下方列出的 `audit_contact_*.png`（缩略总览，多页一张图）。不要把 contact sheet 派给 sub-agent；上下文里已有源 HTML / Stage 5a / preflight，主 agent 是分流决策的最低成本执行者。

contact pass 只回答"哪些页要打开单页大图复核"，不做最终 bug 审计，不修 HTML。

只排查这些全页级 / 高风险信号：
- 页面空白、缺大块内容、整页偏移 / 缩放错误
- 背景色 / 主题色明显错，或出现明显整页截图叠层
- 标题字号 / 字体族明显错，或大段文字换行导致版式结构变了
- 文字重叠、被裁切、溢出 slide、压住其它内容
- 关键图形 / 图表 / 装饰缺失或形态错误（实心变空心、双线框变单线框、圆变方等）
- 列表、图表、细线 / 边框 / 双线框、小字密集、复杂装饰、`deco_snapshot`、SVG、pseudo-element、Stage 5a / preflight 告警页必须入围
- 其余缩略图看似 OK 的页，每 4-5 页至少抽 1 页入围

不要因为这些细节入围：
- 抗锯齿、hinting、字偶距、锯齿边
- < 5 px 位置漂移、细线端点、几像素边框差异
- 字重略粗 / 略细、轻微字距 / 行高 / 阴影柔和度
- 标题换行位置不同但没有遮挡、溢出、叠压、阅读顺序变化

contact pass 输出下面三段，供自己进入第 2 轮：

```
zoom_pages:
- page NN: <必须打开单页 compare 图的原因>

obvious_blockers:
- page NN: <缩略图已经能确认的页面级严重问题；没有就写 none>

likely_ok_pages:
- page NN, page NN, ...
```

## 第 2 轮：detail pass（sub-agent 执行）

把 `zoom_pages` 中的页号按 batch 派给 sub-agent，用下方 page/detail 模板让 sub-agent 看对应 `slide_NN_compare.png` 输出最终 findings。没进 `zoom_pages` 的页主 agent 标 `## page NN · OK`（contact pass 判定）。

不能只凭 contact sheet 判 HIGH/MID finding；要报 finding 必须经第 2 轮 sub-agent 看单页确认。"""

    return f"""{mode_instructions}

## Contact Sheets

{contact_lines}

{AUDIT_PROMPT_MD}"""


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


def build_contact_sheets(out_dir: Path, page_indices: list[int], per_sheet: int = 9) -> list[dict]:
    """Build compact overview sheets from compare images for triage/manual audit."""
    from PIL import Image, ImageDraw, ImageFont

    for old in out_dir.glob("audit_contact_*.png"):
        try:
            old.unlink()
        except OSError:
            pass

    title_font = None
    for name in ("arial.ttf", "DejaVuSans.ttf", "Helvetica.ttc"):
        try:
            title_font = ImageFont.truetype(name, 20)
            break
        except Exception:
            continue
    if title_font is None:
        title_font = ImageFont.load_default()

    contact_sheets: list[dict] = []
    cols = 3
    tile_w = 760
    label_h = 28
    gap = 28
    margin = 18

    for start in range(0, len(page_indices), per_sheet):
        chunk = page_indices[start:start + per_sheet]
        images = []
        for idx in chunk:
            path = out_dir / f"slide_{idx:02d}_compare.png"
            if not path.exists():
                continue
            img = Image.open(path).convert("RGB")
            tile_h = int(round(tile_w * img.height / img.width))
            images.append((idx, img.resize((tile_w, tile_h))))
        if not images:
            continue

        rows = (len(images) + cols - 1) // cols
        tile_h = images[0][1].height
        sheet_w = margin * 2 + cols * tile_w + (cols - 1) * gap
        sheet_h = margin * 2 + rows * (label_h + tile_h) + (rows - 1) * gap
        sheet = Image.new("RGB", (sheet_w, sheet_h), (255, 255, 255))
        draw = ImageDraw.Draw(sheet)

        for pos, (idx, img) in enumerate(images):
            col = pos % cols
            row = pos // cols
            x = margin + col * (tile_w + gap)
            y = margin + row * (label_h + tile_h + gap)
            draw.text((x, y), f"slide_{idx:02d}", fill=(20, 20, 20), font=title_font)
            sheet.paste(img, (x, y + label_h))

        out_path = out_dir / f"audit_contact_{chunk[0]:02d}_{chunk[-1]:02d}.png"
        sheet.save(out_path, optimize=True)
        contact_sheets.append({
            "file": out_path.name,
            "path": str(out_path),
            "pages": chunk,
        })

    return contact_sheets


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

    mode = audit_mode()
    contact_sheets = build_contact_sheets(out_dir, [p["index"] for p in pages_meta])

    index_data = {
        "pptx": str(pptx_path.name),
        "pptx_path": str(pptx_path),
        "total_pages": n,
        "audit_mode": mode,
        "contact_sheets": contact_sheets,
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
    (out_dir / "audit_prompt.md").write_text(build_audit_prompt(mode, contact_sheets), encoding="utf-8")

    return {
        "out_dir": str(out_dir),
        "pages": n,
        "audit_mode": mode,
        "contact_sheets": contact_sheets,
        "fresh": sorted(fresh_set),
        "cached": sorted(skipped_set),
        "incremental": only_indices is not None,
        "index": str(out_dir / "audit_index.json"),
        "prompt": str(out_dir / "audit_prompt.md"),
    }
