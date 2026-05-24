---
name: html-to-pptx
description: 把 HTML 制作的演示文稿（slide deck）转换为 .pptx。保留多 run 富文本（中-英斜体）、矢量定位、字体子集嵌入（换机不掉字）、SVG 图形、半透明背景、圆形/椭圆装饰。通过启发式定位 slide + force-position 强制激活，不依赖具体框架 / 类名 / 库约定，对 transform 平移 / .active 类切换 / scroll-snap 等切页机制统一兜底。当用户提到"HTML 转 PPT"、"网页幻灯片转 pptx"、"想给同事一份 ppt 副本"、"汇报不方便放浏览器"时触发。
---

# html-to-pptx

## 何时触发

用户说出下列任一意图：
- "把这个 HTML / 网页 deck 转成 PPT / pptx"
- "做了 HTML 幻灯片，要给同事一份 ppt"
- "汇报现场不方便放浏览器，想要 ppt 文件"
- 已有 HTML 文件路径 + 提到 ppt / pptx / 演示 / 幻灯片

## 调用

`<skill_dir>` 是这个 skill 安装路径（通常 `~/.claude/skills/html-to-pptx/`，Windows 上是 `%USERPROFILE%\.claude\skills\html-to-pptx\`）。下面命令里的 `<skill_dir>` 替换成实际路径，或者先 `cd <skill_dir>` 再直接 `python convert.py …`。

```bash
python <skill_dir>/convert.py <input.html>
```

- 默认输出到与输入同目录的 `<input>.pptx`
- 字体完全按需：HTML 用到的字体在 convert 时从 Google Fonts 拉取并 subset 嵌入，缓存到 `%LOCALAPPDATA%\html-to-pptx\fonts\` 或 `~/.cache/html-to-pptx/fonts/`，下次同名字体秒复用
- HTML 含 CJK 字符会自动种子 Noto Sans SC + Noto Serif SC（首次约 40 MB / ~30s 下载并 instance 静态 Regular/Bold，之后命中 cache）
- GF 没有的家族会回退到 viewer 系统字体并打印 warning

| 选项 | 含义 |
|---|---|
| `--out <path>` | 自定义输出 .pptx 路径 |
| `--keep-screenshots` | 同时保留每页 HTML 参考截图 + `preflight.json` |
| `--no-embed-fonts` | 跳过字体嵌入。文件更小但换机会回退到系统字体 |
| `--no-preflight` | 关闭 Stage 1 风险预扫 |
| `--no-verify` | 关闭 Stage 5a 结构化自检 |
| `--no-visual-audit` | 关闭 Stage 5b 视觉 audit 物料产出。日常不要关 |
| `--install-user-fonts` | 把自动解析到的非 CJK 字体装到用户字体目录（让 WPS 能正确渲染）。Win/macOS/Linux 都支持。**必须先问用户**，见下方"字体安装确认"章节 |
| `--only-slides N,N,N` | **增量重跑**。逗号分隔的页号（1-based）。measure 只跑指定页 + 与上轮 cached measurement 合并；assemble/embed 仍全量；Stage 5a 只渲指定页；Stage 5b 只重建指定页的 compare 图——其它页全部复用上轮缓存。audit 迭代轮专用，详见下方"增量重跑"章节 |
| `--cleanup` | 不做转换。删 input.pptx 旁的 audit / measurement / preflight 工作物，只保留 .pptx 本身。**最终交付前用**，见下方"工作流"末步 |

## 字体安装确认（第一次 convert 之前完成）

PowerPoint COM `slide.Export()` 和 WPS Office 都不读 pptx 内嵌的裸 TTF——audit 渲染和 WPS 都回退到系统字体。装到用户字体目录后两者都按系统字体走。

- Windows: `%LOCALAPPDATA%\Microsoft\Windows\Fonts\` + HKCU 注册
- macOS: `~/Library/Fonts/`
- Linux: `~/.local/share/fonts/` + `fc-cache -f`

用户级，无需管理员，可手工删。

### 触发判定

看 HTML `<head>` 的 `<link href="...fonts.googleapis.com/...">` / `@import url(https://fonts.googleapis.com/...)` / `<style>` 里 `font-family:` 出现的字体名：

- **触发**：任何 GF / 自托管字体（Bricolage Grotesque、DM Sans、Inter、Space Grotesk、Caveat 等）
- **不触发**：只用系统字体白名单（Arial、Times New Roman、Helvetica、Helvetica Neue、Courier、system-ui、-apple-system、BlinkMacSystemFont、SF Pro、Microsoft YaHei、SimSun、PingFang、Hiragino）

### 触发后

第一次 convert **之前**用 AskUserQuestion 问一次。同意 → 之后**每次** convert 加 `--install-user-fonts`；拒绝 → 之后**都不**加。**整个会话不再问**。

CLAUDE.md / 全局指令里写过"以后字体自动装无需再问"的，直接当同意，跳过 ask。

## 工作流（强制）

convert.py 跑完不等于交付完成。完整流程：

```
convert.py
  → Stage 5a 结构化自检（OOXML 扫描，给提示）
  → Stage 5b 视觉 audit 物料 — 需要 PowerPoint COM 或 LibreOffice 渲染器
  → 【并行 audit】按 batch（4 页/batch）并行 sub-agent 看 compare 图返回 findings，主 agent 合并写 audit_findings.md
  → 按 finding 逐项做最小局部 HTML 修改（见下"修复纪律"）
  → 一批 finding 改完后 `convert.py <html> --only-slides <被改页号>` + 只对 fresh 页重审（见下"增量重跑"）
  → 所有页 OK 或仅剩 LOW
  → 【若撞到 HTML 反模式 / 新 OOXML 边界】沉淀到 lessons-learned，见下方"沉淀 HTML 问题与 OOXML 边界"
  → python convert.py <out>.pptx --cleanup  ← 删 audit/measurement/preflight 工作物
  → 把 .pptx 路径交付给用户
```

### 渲染器要求（Stage 5b 前置条件）

Stage 5b 视觉 audit 需要把 .pptx 渲染成 PNG，依赖：
- **PowerPoint COM**（Windows + Office + `pywin32`），或
- **LibreOffice**（跨平台，配 `pip install pdf2image`）

任一可用就能跑 audit。convert 输出里看到 `[self-check] 跳过：找不到可用的 pptx 渲染器`，说明两个都没装——此时 audit **不会**产出 compare 图，Stage 5b 直接跳过。

**这种情况下 agent 必须 ask 用户**（不要静默交付未审计的 pptx）：

> "你机器上没装 PowerPoint 也没装 LibreOffice，视觉 audit 跑不了，PPT 可能有看不出的视觉 bug。三个选择：
> 1. 装 LibreOffice（推荐，跨平台，2-3 分钟）：`winget install LibreOffice.LibreOffice`（Windows）/ `brew install --cask libreoffice`（mac）/ `apt install libreoffice`（Linux），然后 `pip install pdf2image`，重跑 convert
> 2. 跳过 audit 直接交付（接受 PPT 可能有视觉 bug 的风险）
> 3. 在已经装了 Office 的另一台机器上重跑"

用户选 1 → 等他装完重跑；选 2 → 加 `--no-visual-audit` 跑一遍，把告知风险后交付；选 3 → 把当前目录 + HTML 发给他。

### 并行 audit（强制，前提：5b 跑起来了）

所有页数都走 sub-agent 并行 dispatch；主 agent 不自己 Read compare 图。

| slide 数 | 策略 |
|---|---|
| ≤ 4 | 1 个 batch |
| 5-20 | 每 batch 4 页（向上取整，例：9 页 = 4+5） |
| > 20 | 每 batch 4-5 页 |

完整 sub-agent 调用模板 + 检查清单 + findings 格式见 `<out>_audit/audit_prompt.md`。多个 batch 塞在主 agent 同一条 message 里才并行。

sub-agent 只返回 findings 文本（含每页 `## page NN` 块），主 agent 统一合并写 `audit_findings.md`——并发写会互相覆盖。

### 与 sub-agent 并行的主 agent 准备（强制）

Agent 调用用 `run_in_background: true`，并在**同一条 message** 里同时发：

```
Agent(run_in_background: true, description: "Audit slides 1-4", prompt: ...)
Agent(run_in_background: true, description: "Audit slides 5-9", prompt: ...)
Read(file_path: "<deck>/template.html")
Read(file_path: "<skill_dir>/references/lessons-learned.md")
Grep(pattern: "class=\"slide |data-slide=", path: "<deck>/template.html",
     output_mode: "content", -n: true)
```

只发"无论 findings 是什么都用得上"的准备——不要预测 findings 提前改 HTML。

### 修复纪律（针对 finding，不追根因）

收到 `audit_findings.md` 之后，**一轮里把本轮所有 finding 都改完，再一次性重跑 convert + audit**——不是改一个跑一次。一轮 = 一批 HTML 编辑 + 一次 convert + 一次 audit。

每个 finding 内部做最小局部 HTML 修改：

- 每个 finding 只改"让它消失"的那一处，改完进下一个 finding（**继续改 HTML**，不是重跑 convert）
- **不**追溯"为什么字体回退 / 为什么布局偏移"的深层根因
- **不**做"我顺手把 .footer 也改成绝对定位"这种 finding 列表外的优化
- **不**跨 finding 做结构性重构（"统一把所有 slide 的 font-family 显式声明"≠ 单个 finding 的最小修复）
- 同一 finding 反复试改 ≥ 2 次仍不 OK，停下来告诉用户，**不要**继续扩大改动面

本批 finding 改完一次性 `python convert.py <html> --only-slides <被改过的页号>`。

判定标准：你的 diff 行数 ≤ findings 数 × 3 行。超过这个量级说明在"乱发挥"。

### 增量重跑（`--only-slides`）

audit 第 2 轮起用，首轮全量。

```bash
python convert.py <html> --only-slides 2,7,12
```

增量覆盖 measure / Stage 5a 渲染 / Stage 5b compare 图，仅对列出的页 + 缓存缺失的页执行；assemble / embed_fonts 仍全量。`audit_index.json` 标 `incremental_mode: true`、`fresh_indices`、`cached_indices`、每页 `fresh: true/false`。

并行 audit 只对 `fresh_indices` 分 batch 分发 sub-agent。

前提：上轮的 `<out>_audit/` 目录还在。cache 缺失自动回退全量；HTML 页数变化自动检测并回退全量 measure。

**不能用 `--only-slides` 的情况**（视觉溢出列出的页，缓存假阴性）：
- 全局 CSS（`<style>`、根选择器、`.slide` 通用样式）变更
- 新增 / 删除字体（@import、font-family 声明变更）
- deck-level token（背景、主题色、间距）变更
- 任何影响其它页布局的全局变量变更

这几类不带 flag 全量重跑。判断不准 → 全量重跑总是安全的。

**交付前必须 cleanup**：audit 物料是 agent 工作用的中间产物，用户只要 .pptx。最终一行命令 `python convert.py <out>.pptx --cleanup` 会把同目录下 `<out>_audit/`、`<out>_measurements*`、`<out>_preflight.json` 全部删干净，目录里只剩 `<out>.pptx`。

## 沉淀 HTML 问题与 OOXML 边界（强制）

你的角色是**调用 + 用法沉淀**，不是修 skill。发现 skill 内部 bug 时**不要原地改 measure.py / assemble.py**。

只沉淀两类：

### 1. HTML 写法问题（用户的 HTML 让转换失真）

| 类别 | 判定 | 做什么 |
|---|---|---|
| 单次 case（只在这 deck 出现，改一两行 HTML 就好） | 单页特例 | 直接改 HTML 源；不沉，不通报 |
| 通用 HTML 反模式（任何人写类似 HTML 都会踩） | 跨 deck | 改 HTML + 沉到 `references/lessons-learned.md` 的 "HTML 写法规避" 区 |

### 2. OOXML 表达力边界（PowerPoint / OOXML 天然不支持）

典型：彩色 emoji、CSS `background-clip: text` 文字渐变、`filter: blur` / 复杂 mask、`backdrop-filter` 之类。

改 HTML 走替代通路（Twemoji SVG `<img>` / inline `<svg><text fill="url(#grad)">` / 让容器走 deco_snapshot 截图）+ 沉到 `references/lessons-learned.md` 的 "OOXML 边界" 区。

已沉的边界（搜 lessons-learned 验证最新版）：
- 彩色 emoji → Twemoji SVG `<img>`
- CSS 文字渐变 → inline `<svg><text fill="url(#grad)">`
- backdrop-filter / 复杂 filter → 让该容器走 deco_snapshot 截图

### 看似 skill 内 bug 怎么处理（**不修不沉**）

发现某症状用同种 CSS 模式换 deck 还会撞、且不在已知 OOXML 边界内 → 这是 skill 抽象不到位。**不要原地改 skill 代码**。做这两件事：

1. 当前 deck **走 HTML 端 workaround** 把 finding 修掉，让用户拿到能交付的 PPT
2. **明确告知用户**："发现一个看起来是 skill 内的通用 bug：[症状 + 触发 CSS 模式 + 当前 workaround]。建议作者按 issue 收录，让 skill 在 measure/assemble 里更通用地处理这种模式"

### 共享 vs 本地：HTML/OOXML 沉淀到哪一边

| 内容 | 写到哪 | 上游 skill 更新时 |
|---|---|---|
| 通用 HTML 反模式 / OOXML 边界（让所有用户受益） | `references/lessons-learned.md` | **会被覆盖**——PR 上游，或备份后手动合并 |
| 本地业务 / 特定客户专有写法 | `references/lessons-learned-local.md`（gitignored） | **不会**被覆盖。文件不存在就 new 一个，schema 跟 lessons-learned 一致 |

**判定**：任何人写类似 HTML 都会踩？是 → `lessons-learned.md`；只在我们业务/客户场景出现 → `lessons-learned-local.md`。

排查时按顺序搜：`lessons-learned.md` → `lessons-learned-local.md`（如果存在）。两者命中任一即可。

## 流水线

```
[1 输入识别/预扫] → [2 测量] → [3 组装] → [4 字体嵌入] → [5a 结构化自检] → [5b 视觉 audit]
   preflight.py    measure.py   assemble.py   embed_fonts.py  self_check.py    visual_audit.py
```

**Stage 1 preflight**：扫已知风险模式（slide 根 deco、多层 text-shadow、backdrop-filter、video、非内置字体等），输出 `preflight.json`。

**Stage 2-4 转换**：measure 用 Playwright 实测 DOM → assemble 写 OOXML → embed_fonts 跨页 subset 嵌入。

**Stage 5a self_check**：解压 pptx 扫结构化告警 ——
- 全屏 `<p:pic>` 嫌疑（`FULL-PIC`，潜在 deco_snapshot 双层 bug）
- 文本框横向重叠（`LAYOUT`）
- 合并 preflight 高风险页（`PREFLIGHT`）

**不做像素 diff** —— 像素 diff 数字对局部 bug 不敏感，会给假信心。视觉判断完全交 5b。

**Stage 5b visual_audit**（**需要 PowerPoint COM 或 LibreOffice 渲染器**，见上"渲染器要求"）：产出 `<out>_audit/`：
- `slide_NN_compare.png` × N（HTML | PPT 双栏拼图）
- `audit_index.json`（每页元数据 + 结构化告警）
- `audit_prompt.md`（给你的检查清单 + sub-agent 调用模板）

照着 audit_prompt.md 做，写 `audit_findings.md`。无渲染器时 5b 跳过——按上面"渲染器要求"小节 ask 用户。

## 调用前确认

1. 输入是单文件 HTML，含若干 slide 元素（`<section class="slide">` / `<div class="slide">` / `<deck-stage>` 子节点等都支持）
2. 不需要为用户配置切页机制 —— skill 自动识别
3. 用户已写好 HTML 才来调用，按要求转就行

## 报告

| 用户反馈 | 处理 |
|---|---|
| 排版问题 / 字溢出 / 撑出框 | 先看 HTML 参考图是否也溢出；若 HTML 正常、PPT 异常，按 `references/lessons-learned.md` 排查 |
| 装饰显示成方块 | 看 compare 图判断是几何、画面捕获还是渲染器差异，查 `references/lessons-learned.md` |
| 字体不对（PowerPoint 里看） | 确认 OOXML typeface= 和 cache 里 TTF 的 nameID=1 严格一致；如果用 GF 自动解析的字体出问题，参考 `lessons-learned.md` 的 "GF Medium-as-Regular" 条目 |
| 字体不对（WPS 里看） | 99% 是 WPS 不读裸 TTF 嵌入字体。问用户是否同意装到用户字体目录后用 `--install-user-fonts` 重转 |
| 中文显示方框 □□□ | 检查 OOXML rPr 的 `<a:ea>` 是否走了 CJK 字体，参考 `lessons-learned.md` |
| 居中文字位置偏下 / 偏出容器 | 检查 assemble.py 里 `_text_box_size_px` 对 flex/grid + align-items:center 是否跳过了 h*1.4 撑高，参考 `lessons-learned.md` 的 "anchor=ctr 双重撑高" 条目 |
| 想加页 / 改文字 | 改 HTML 源文件重新转，不要在 pptx 里改 |

## 不要

- 不向用户解释完整管线（除非问），只报告转换结果与告警
- 不在未看 audit compare 图前就归因
- 不承诺 1:1 视觉还原 —— OOXML 表达力有限

## 排查路径

1. 看 Stage 5a 自检报告：哪几页被告警
2. 看 Stage 5b audit compare 图：用 Read 工具逐页对照
3. 搜 `references/lessons-learned.md` 已知症状（HTML 反模式 / OOXML 边界）
4. HTML 正常、PPT 异常且不符合已知边界 → 看是不是 skill 抽象不到位（按"看似 skill 内 bug 怎么处理"流程：HTML workaround 修当前 deck + 告知用户提 issue）

## 引用

- 维护扩展必读 → [`references/methodology.md`](./references/methodology.md) — 五步反假设流水线 checklist（作者扩展 skill 时读，agent 调用不必读）
- 历史踩坑修复 + HTML 反模式 + OOXML 边界 → [`references/lessons-learned.md`](./references/lessons-learned.md)（agent 排查必读）
