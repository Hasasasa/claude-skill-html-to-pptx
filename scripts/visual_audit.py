"""visual_audit.py — Stage 5b 视觉审计物料。

产出 `<out>_audit/`：
- `slide_NN_compare.png` × N — HTML | PPT 双栏拼图
- `audit_index.json` — 每页元数据 + 结构化告警
- `audit_prompt.md` — 给上游 VLM agent 的审计指南
"""
import json
from pathlib import Path


AUDIT_PROMPT_MD = """# Visual Audit

逐页视觉对比 HTML 参考图与 PPT 输出图，识别 PPT 这边的视觉问题，迭代修复到交付。

## 流程

1. 读 `audit_index.json` 拿到页清单。若 `incremental_mode=true`，只看 `fresh_indices` 列出的页
2. 按 batch 并行 dispatch sub-agent 看图（强制，不论页数；详见下"并行执行"）
3. 主 agent 收回各 sub-agent 的 findings 文本（每个含若干 `## page NN` 块），按页号拼成 `audit_findings.md`（首轮）或 `audit_findings_round_N.md`（迭代轮）
4. 每个 finding 内部做最小局部 HTML 修改（只改让它消失的那一处）：不追溯根因、不做 finding 列表外的"顺手优化"、不跨 finding 做结构性重构。同一 finding 试改 ≥ 2 次仍不 OK，停下来告诉用户。判定标准：diff 行数 ≤ findings 数 × 3 行
5. 本轮所有 finding 改完一次性 `convert.py <html> --only-slides N1,N2,...`（本轮被改过的页号）。例外：改了全局 CSS / 字体 / deck-level 样式 → 不带 `--only-slides`，全量重跑。回到第 1 步
6. 所有页 OK 或仅剩 LOW 才交付

## 并行执行（强制，不论页数）

每 batch 3-4 页一个 sub-agent。即使只剩 1-3 页也走 sub-agent，主 agent 不自己 Read compare 图。何时拆 batch 见 SKILL.md "并行 audit" 页数策略表。

Sub-agent 调用模板（每个 Agent 一个 batch，全部塞在主 agent 同一条 message 里才并行）：

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

仅看这几页，不要试图读其它页或改代码。

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
- [HIGH] <一句话描述>
- [MID]  <一句话描述>
- [LOW]  <一句话描述>

无问题的页：
## page NN · OK

HIGH=用户一眼能看出 / MID=细看才发现 / LOW=设计美化建议。**每页都必须有一个块（OK 或带 finding），不要漏页**。'''
)
```

sub-agent 只返回 findings 文本，**不要让它直接写 audit_findings.md**——并发写会互相覆盖，主 agent 统一合并。

同一条 message 里跟着发并行准备：

```
Read(file_path="<deck>/template.html")
Read(file_path="<skill_dir>/references/lessons-learned.md")
Grep(pattern='class="slide |data-slide=', path="<deck>/template.html",
     output_mode="content", -n=true)
```

只发"无论 findings 是什么都用得上"的准备，不要预测 findings 提前改 HTML。详见 SKILL.md "与 sub-agent 并行的主 agent 准备"。

## 检查清单（同上，给主 agent 复核用）

1. 文字被线条 / 形状边界 / 图片角穿过 / 覆盖
2. 文字之间不该有的重叠 / 叠压
3. 文字溢出 slide 边界 / 被裁切 / 溢入相邻列
4. 元素相对 HTML 参考图大幅错位
5. 字体回退 / 字号变形
6. 图片拉伸 / 错位 / 缺失，装饰色块变形
7. 颜色错误（明显偏离 HTML）

## findings 格式

```
## page 08
- [HIGH] 右列 border-left 竖线穿过左列 lead "实"字
- [LOW] 副标题与正文间距偏小

## page 09 · OK
```

- **HIGH**：用户一眼能看出，必须修
- **MID**：细看才发现，强烈建议修
- **LOW**：设计美化建议，可选

## 修复路径

所有 finding 都改源 HTML（局部 inline style / 替换字体声明 / 替换 emoji 字符等）。不改 skill 代码。

## 终止条件

- 通常 2-3 轮收敛
- 超过 5 轮还有 HIGH → 停，告诉用户，可能是 OOXML 边界或 skill 抽象不够
- 剩余 HIGH/MID 若属于 OOXML 边界（lessons-learned 的 OOXML Limits 段），列入 boundary 清单交付时告知用户
"""


def build_compare_image(html_png: Path, ppt_png: Path, out_path: Path, page_idx: int):
    """生成单页 HTML | PPT 双栏拼图。"""
    from PIL import Image, ImageDraw, ImageFont
    try:
        title_font = ImageFont.truetype("arial.ttf", 36)
    except Exception:
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
