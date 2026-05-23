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

1. 读 `audit_index.json` 拿到页清单
2. **并行 dispatch sub-agent 看图**（详见下"并行执行"）—— 每页一个 Agent，不要主 agent 一张张串行 Read
3. 主 agent 收回各 sub-agent 的 findings 文本，按页号拼成 `audit_findings.md`（首轮）或 `audit_findings_round_N.md`（迭代轮）
4. 按 finding 修源 HTML 或 skill 代码
5. 重跑 `convert.py`，回到第 2 步
6. 所有页 OK 或仅剩 LOW 才交付

## 并行执行（强制）

每页一个 sub-agent。理由：VLM 单图 100% 注意力比同时看 3-4 张漏判更少；并行墙钟也快 3-5×。

按页数选策略：

| slide 数 | 策略 |
|---|---|
| ≤ 16 | **每页一个 sub-agent**，一条消息里全部 Agent 调用并行 dispatch |
| > 16 | 分 4-5 个 batch 并行（避免撞 API 限速 / 节省 token） |

**Sub-agent 调用模板**（每页一个 Agent，全部塞在主 agent 同一条 message 里——多个 Agent 一次发才并行）：

```
Agent(
  description="Audit slide NN",
  subagent_type="general-purpose",
  prompt='''你是单页视觉审计员。

读这一张图：<full path to slide_NN_compare.png>（左=HTML 参考，右=PPT 输出）。
按下面检查清单逐项识别 PPT 这边相对 HTML 的视觉问题，仅看这一页，不要试图读其他页或修代码。

检查清单（按重要度排序）：
1. 文字被线条 / 形状边界 / 图片角穿过 / 覆盖
2. 文字之间不该有的重叠 / 叠压
3. 文字溢出 slide 边界 / 被裁切 / 溢入相邻列
4. 元素相对 HTML 参考图大幅错位
5. 字体回退 / 字号变形
6. 图片拉伸 / 错位 / 缺失，装饰色块变形
7. 颜色错误（明显偏离 HTML）

输出**纯文本**（不要 markdown code fence 包装），严格用以下格式：

如果有问题：
## page NN
- [HIGH] <一句话描述>
- [MID]  <一句话描述>
- [LOW]  <一句话描述>

如果无问题：
## page NN · OK

HIGH=用户一眼能看出 / MID=细看才发现 / LOW=设计美化建议。'''
)
```

**关键**：sub-agent 只返回 findings 文本，**不要让它直接写 audit_findings.md**——多个并发写会互相覆盖。主 agent 统一合并。

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

| 问题 | 改哪里 |
|---|---|
| 单页特有（间距 / 宽度 / 字号） | 源 HTML inline style |
| 多页同类问题 | skill 的 measure / assemble / adapters |
| 字体 / 字重 | FONT_PLAN 或源 HTML font-family |
| 颜色 / 装饰 | 定位是 measure 抽错还是 assemble 写错 |

不要在 assemble 加 collision avoidance 死规则——规则覆盖窄、副作用大、跟反假设原则冲突。

## 终止条件

- 通常 2-3 轮收敛
- 超过 5 轮还有 HIGH → 思路错了，重新评估布局选型
- 剩余 HIGH/MID 若属于 known-boundaries（OOXML 表达不了），列入 boundary 清单交付时告知用户
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
                        out_dir: Path) -> dict:
    """产出 audit 物料包：compare 图 × N + audit_index.json + audit_prompt.md。

    返回 dict 描述包内容（方便 convert.py 在终端打印）。
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    html_pngs = sorted(html_screenshots_dir.glob("slide_*.png"))
    ppt_pngs = sorted(ppt_screenshots_dir.glob("slide_*.png"))
    n = min(len(html_pngs), len(ppt_pngs))

    pages_meta = []
    pages_by_idx = {p["idx"]: p for p in self_check_result.get("pages", [])}
    preflight_by_idx = {s["index"]: s for s in (preflight_result or {}).get("slides", [])}

    for i in range(n):
        idx = i + 1
        compare_path = out_dir / f"slide_{idx:02d}_compare.png"
        ok = build_compare_image(html_pngs[i], ppt_pngs[i], compare_path, idx)
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
        })

    index_data = {
        "pptx": str(pptx_path.name),
        "pptx_path": str(pptx_path),
        "total_pages": n,
        "instructions_file": "audit_prompt.md",
        "findings_output": "audit_findings.md",
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
        "index": str(out_dir / "audit_index.json"),
        "prompt": str(out_dir / "audit_prompt.md"),
    }
