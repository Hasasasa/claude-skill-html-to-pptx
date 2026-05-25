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
| `--cleanup` | 不做转换。删 input.pptx 旁的 audit / measurement / preflight 工作物，只保留 .pptx 和 audited.html。**最终交付前用**，见下方"工作流"末步 |

## 工作副本（原 HTML 不动）

第一次 `python convert.py <input>.html` 跑完，convert.py 自动 cp 一份 `<input>.audited.html` 到源 HTML 同目录。机制：

- **所有 audit 修复改 `audited.html`，不改源 HTML**
- 后续轮 convert（含 `--only-slides`）input 用 `<input>.audited.html`，输出仍是 `<input>.pptx`（自动去掉 `.audited` 后缀）
- `audited.html` 已存在 → 复用，不覆盖（保留上轮修复）
- agent 误传源 HTML 当 input → convert.py 内部检测后切到 audited.html 兜底
- cleanup 不删 audited.html（视为可交付副产物，跟 `.pptx` 一组交付）
- 想从头开始 → 手动删 audited.html 后重跑

## 配置（`.config.local.toml`）

skill 根目录下可选的本机偏好文件，gitignored，每用户一份。两个 key：

```toml
[fonts]
auto_install = "ask"   # "yes" 自动装、"no" 永不装、"ask" 由 agent 询问

[cleanup]
default = "clean"      # "clean" 工作流末步自动 cleanup、"keep" 保留 audit 产物
```

- 文件 / key 缺失 → 走默认（同上）
- 模板：`<skill_dir>/.config.local.toml.example`
- `fonts.auto_install` 由 convert.py 强制读取并应用（=yes 自动加 `--install-user-fonts`、=no 自动忽略）
- `cleanup.default` 由 agent 在工作流末步读取决定行为
- **首次使用**：agent 通过 AskUserQuestion 收集偏好后**自己 Write 这个文件**——用户不用手动建

## 字体安装确认（第一次 convert 之前完成）

PowerPoint COM `slide.Export()` 和 WPS Office 都不读 pptx 内嵌的裸 TTF——audit 渲染和 WPS 都回退到系统字体。装到用户字体目录后两者都按系统字体走。

- Windows: `%LOCALAPPDATA%\Microsoft\Windows\Fonts\` + HKCU 注册
- macOS: `~/Library/Fonts/`
- Linux: `~/.local/share/fonts/` + `fc-cache -f`

用户级，无需管理员，可手工删。

### 询问顺序（先读 config）

每次会话第一次要 convert 之前，按下面顺序判定：

1. Read `<skill_dir>/.config.local.toml`，看 `fonts.auto_install`：
   - `"yes"` → 不问，CLI 不用加 flag（convert.py 自己加），后续整个会话也不再问
   - `"no"` → 不问，不加 flag，整个会话不再问
   - `"ask"` 或 config 文件不存在 / 字段缺失 → 进入下面"触发判定 + ask"
2. CLAUDE.md / 全局指令里写过"以后字体自动装无需再问"的，等同 `"yes"`，并补写 config（保持单一事实源）

### 触发判定（仅当 config = ask 时）

看 HTML `<head>` 的 `<link href="...fonts.googleapis.com/...">` / `@import url(https://fonts.googleapis.com/...)` / `<style>` 里 `font-family:` 出现的字体名：

- **触发**：任何 GF / 自托管字体（Bricolage Grotesque、DM Sans、Inter、Space Grotesk、Caveat 等）
- **不触发**：只用系统字体白名单（Arial、Times New Roman、Helvetica、Helvetica Neue、Courier、system-ui、-apple-system、BlinkMacSystemFont、SF Pro、Microsoft YaHei、SimSun、PingFang、Hiragino）

### 触发后

AskUserQuestion 问一次，**两条都要做**：

- **写入** `<skill_dir>/.config.local.toml`：同意 → `fonts.auto_install = "yes"`、拒绝 → `= "no"`（不写 ask，否则下次会话还会再问一次）
- 本次会话同时按答案走（同意 → 加 `--install-user-fonts`，拒绝 → 不加）

写 config 时：文件**不存在** → Write 整文件落盘（两段 [fonts] [cleanup] 模板内容）；文件**已存在** → 用 Edit 做单行 replace（只改 `auto_install` 一行），**不要** Write 重写整文件，否则会吞掉用户其它自定义 key（包括以后可能新增的章节）。

## 工作流（强制）

convert.py 跑完不等于交付完成。完整流程：

```
convert.py（首轮自动 cp <input>.audited.html，所有修复改这个副本）
  → Stage 5a 结构化自检（OOXML 扫描，给提示）
  → Stage 5b 视觉 audit 物料 — 需要 PowerPoint COM 或 LibreOffice 渲染器
  → 【并行 audit】按 batch（4 页/batch）并行 sub-agent 看 compare 图返回 findings，主 agent 合并写 audit_findings.md
  → 按 finding 逐项做最小局部修改，目标是 audited.html（见下"修复纪律"）
  → 一批 finding 改完后 `convert.py <input>.audited.html --only-slides <被改页号>` + 只对 fresh 页重审（见下"增量重跑"）
  → 所有页 OK 或仅剩 LOW
  → 【若撞到 HTML 反模式 / 新 OOXML 边界】沉淀到 lessons-learned，见下方"沉淀 HTML 问题与 OOXML 边界"
  → 读 `<skill_dir>/.config.local.toml` 的 `cleanup.default`：
        "clean"（缺省）→ 直接 `python convert.py <out>.pptx --cleanup`，不问用户
                          清完目录剩：<input>.html（源，未动）+ <input>.audited.html（修复版）+ <out>.pptx
        "keep"        → 跳过 cleanup，audit 产物全部保留
  → 把 .pptx 路径交付给用户（audited.html 作为修复版 HTML 一同交付）
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

拿到 findings 之后：

- **不**自己 Read compare 图（findings 已经把症状写明）
- **不**读 `measure.py` / `assemble.py` / 任何 skill 内代码找根因
- 直接打开 `audited.html` 按 finding 逐条改
- 例外：finding 描述含糊到无法定位 HTML 修改点 → 只 Read 该单页 compare 图复核，且只看那一张

每个 finding 内部做最小局部 HTML 修改：

- 每个 finding 只改"让它消失"的那一处，改完进下一个 finding（**继续改 HTML**，不是重跑 convert）
- **不**追溯"为什么字体回退 / 为什么布局偏移"的深层根因
- **不**做"我顺手把 .footer 也改成绝对定位"这种 finding 列表外的优化
- **不**跨 finding 做结构性重构（"统一把所有 slide 的 font-family 显式声明"≠ 单个 finding 的最小修复）
- 同一 finding 在**连续两轮**（上一轮 + 本轮）都出现 → sticky 命中，停下来告诉用户，**不要**继续扩大改动面。判定 key = `(页号, 视觉元素短名)`：
  - 措辞 / 严重度 / HIGH↔MID 切换都**不重置**判定。例：`(page 5, take 字号)` 在 round 6 被描述成"偏小"、round 7 被描述成"严重不足" → 同一 sticky key，round 7 即命中
  - 只有"页号 + 元素"都不重合才算新 sticky key
  - **隔轮恢复**：若 round 6 出现、round 7 消失、round 8 又出现 → 不算 sticky（round 7 缺席视为问题已修，round 8 是新回归）
  - 命中 = 本轮**不修**该 finding、列入 skill 边界候选告诉用户，本轮其他 finding 照修

  实施：每轮合并完 `audit_findings.md`、动 Edit 之前，必须先 Read 上一轮 `audit_findings.md`（首轮无）做两两比对——本轮和上轮**都**出现的 key 即命中。命中的 finding 单独列出来交给用户判断，不再进本轮编辑队列。

本批 finding 改完一次性 `python convert.py <input>.audited.html --only-slides <被改过的页号>`。

判定标准：你的 diff 行数 ≤ findings 数 × 3 行。超过这个量级说明在"乱发挥"。

### 增量重跑（`--only-slides`）

audit 第 2 轮起用，首轮全量。

```bash
python convert.py <input>.audited.html --only-slides 2,7,12
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

## 沉淀 HTML 问题与 OOXML 边界（强制）

你的角色是**调用 + 用法沉淀**，不是修 skill。发现 skill 内部 bug 时**不要原地改 measure.py / assemble.py**。

只沉淀两类：

### 1. HTML 写法问题（用户的 HTML 让转换失真）

| 类别 | 判定 | 做什么 |
|---|---|---|
| 单次 case（只在这 deck 出现，改一两行 HTML 就好） | 单页特例 | 直接改 HTML 源；不沉，不通报 |
| 通用 HTML 反模式（任何人写类似 HTML 都会踩） | 跨 deck | 改 HTML + 沉到 `references/lessons-learned.md` 的 "HTML 写法规避" 区 |

### 2. OOXML 表达力边界（PowerPoint / OOXML 天然不支持）

绝大多数装饰类 CSS（backdrop-filter / filter / mix-blend-mode / skew / box-shadow / gradient / `::before` `::after` **装饰用法**——空 content + bg-image / box-shadow 等）已经被"覆盖策略"里的 `hasComplexDecoration` 自动捕获走 deco_snapshot——**HTML 不用改**。`::before` `::after` **string content**（`content: "↑"` 等）走文字通道，详见 [`references/supported-css.md`](./references/supported-css.md) 第 1 档。

仍需 HTML workaround 的真边界（OOXML 文字层原生表达不了的）：

| 边界 | 替代通路 |
|---|---|
| 彩色 emoji（COLR/CPAL 字体） | Twemoji SVG `<img>`，走 img 通道嵌入 |
| `background-clip: text` 文字渐变 | inline `<svg><text fill="url(#grad)">`，走 SVG 通道 |

撞到这两类 → 改 HTML 走替代通路 + 沉到 `references/lessons-learned.md` 的 "OOXML 边界" 区。

撞到不在上表的新 CSS 不支持 → **先看 `hasComplexDecoration` 是不是漏触发**（加一行触发条件即可），而不是改 HTML。详见 [`references/supported-css.md`](./references/supported-css.md) 的"加新支持的标准做法"。

### 看似 skill 内 bug 怎么处理（**不修不沉**）

发现某症状用同种 CSS 模式换 deck 还会撞、且不在已知 OOXML 边界内 → 这是 skill 抽象不到位。**不要原地改 skill 代码**。做这两件事：

1. 当前 deck **走 HTML 端 workaround** 把 finding 修掉，让用户拿到能交付的 PPT
2. **明确告知用户**："发现一个看起来是 skill 内的通用 bug：[症状 + 触发 CSS 模式 + 当前 workaround]。建议作者按 issue 收录，让 skill 在 measure/assemble 里更通用地处理这种模式"

### 沉淀写到哪：本地副本 + 上游模板

`references/lessons-learned.md` 是**每用户独立的本地工作副本**（gitignored）。首次跑 `convert.py` 时 `local_config.seed_lessons_learned()` 自动从模板 `lessons-learned.md.example`（committed）拷贝一份。机制和 `.config.local.toml` 一样。

- **agent 排查 / 沉淀都读写本地 `lessons-learned.md`**——自由加 / 改 / 整理，无 git 噪声
- 本地副本**永不被上游覆盖**——`git pull` 拉的是 `.example` 模板，不动你的工作副本
- 想把某条沉淀上游：**作者手动把 entry 复制回 `lessons-learned.md.example`** 再 commit；选择性 curate，不是全推
- 本地业务 / 特定客户专有写法 → 直接写在本地副本里就行，不会上游

**触发**：撞到通用 HTML 反模式 / OOXML 边界 → 加到本地 `lessons-learned.md` 对应表。每次 convert 都重读，立即生效。

## 覆盖策略（核心架构）

把"支持哪些 CSS"翻译成"输出走哪条路径"——四档兜底，命中即捕获，不依赖逐属性翻译：

| 档位 | 输出形式 | 触发条件 | 文字状态 |
|---|---|---|---|
| **矢量文字** | OOXML textbox（多 run 富文本） | 所有 text leaf 永远走这条 | 可编辑、可搜索、可缩放 |
| **矢量形状** | OOXML 几何原语（圆 / 圆角矩形 / 线） | 仅 `bg-color` / `border` / `border-radius` 简单组合 | — |
| **栅格装饰** | PNG snapshot（仅装饰像素） | `_js_snippets.py:hasComplexDecoration` 命中 | **仍走矢量** — 文字 hide 后截图，再用 textbox 画在截图之上 |
| **媒体直传** | 原始 `<svg>` / `<img>` / `<canvas>` 截图 | 元素本身就是媒体 | — |

### 关键不变量

**栅格档位只栅格化装饰像素，文字始终矢量。** `_DECO_HIDE_FOREGROUND_JS` 在截图前会：

1. `visibility:hidden` 所有带直接文字的子元素
2. `color:transparent` deco 自身的文字
3. 截图——拿到"只有装饰、没有文字"的 PNG
4. 截图后恢复

之后矢量 textbox 画在截图之上，所以**任何走栅格档位的 CSS 都不影响文字可编辑性**。

### `hasComplexDecoration` 当前触发条件

任何一条命中即走截图（`scripts/_js_snippets.py`）：

- `background-image`（含 linear/radial/repeating/conic-gradient、`url(...)`）
- `box-shadow`
- `outline`
- `::before` / `::after` 装饰内容
- `overflow:hidden` 容器 + 含 transformed 子（裁切框）
- `backdrop-filter`
- `filter`（blur / drop-shadow / saturate / 等）
- `mix-blend-mode`
- 不可表达 transform（skew / 非均匀 scale）

### 加新支持的标准做法

发现某种 CSS 没被捕获、视觉缺失：

1. **先判断它是装饰类还是文字类**：
   - 装饰类（视觉效果作用于背景 / 边框 / 整块） → **只在 `hasComplexDecoration` 加一行触发条件**，1 行代码完事
   - 文字类（影响 run-level 样式，如 writing-mode / text-stroke）→ measure 抓 style 字段 + assemble 翻译，约 5-10 行
   - 几何变换类 → 走第一条，截图兜底
2. **绝大多数都属于第一条**——除非视觉效果必须作用于文字本身，否则统一走截图，不写新代码。

详见 [`references/supported-css.md`](./references/supported-css.md)。

## 流水线

```
[1 输入识别/预扫] → [2 测量] → [3 组装] → [4 字体嵌入] → [5a 结构化自检] → [5b 视觉 audit]
   preflight.py    measure.py   assemble.py   embed_fonts.py  self_check.py    visual_audit.py
```

**Stage 1 preflight**：扫已知风险模式（slide 根 deco、多层 text-shadow、video / canvas、非内置字体等），输出 `preflight.json`。注：`backdrop-filter` / `filter` / `mix-blend-mode` / `skew` 已自动走 deco_snapshot，preflight 即便扫到也只是信息性提示，不需要 HTML workaround。

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
| 想加页 / 改文字 | 改源 `<input>.html`（不是 audited.html，那是 audit 修复版）→ 删 audited.html → 重跑 convert |
| 排版问题 / 字溢出 / 撑出框 | 先看 HTML 参考图是否也溢出；若 HTML 正常、PPT 异常，按 `references/lessons-learned.md` 排查 |
| 装饰显示成方块 | 看 compare 图判断是几何、画面捕获还是渲染器差异，查 `lessons-learned.md` 的 OOXML 边界表 |
| 字体不对（PowerPoint 里看） | 确认 OOXML typeface= 和 cache 里 TTF 的 nameID=1 严格一致；GF 自动解析问题查 `lessons-learned.md` "已知 regression checkpoint" |
| 字体不对（WPS 里看） | 99% 是 WPS 不读裸 TTF 嵌入字体。问用户是否同意装到用户字体目录后用 `--install-user-fonts` 重转 |
| 中文显示方框 □□□ | 检查 OOXML rPr 的 `<a:ea>` 是否走了 CJK 字体 |
| 居中文字位置偏下 / 偏出容器 | 检查 `assemble.py:_text_box_size_px` 对 flex/grid + align-items:center 是否跳过了 h*1.4 撑高（详见 `lessons-learned.md` "已知 regression checkpoint"） |

## 不要

- 不向用户解释完整管线（除非问），只报告转换结果与告警
- 不在未看 audit compare 图前就归因
- 不承诺 1:1 视觉还原 —— OOXML 表达力有限

## 排查路径

**先判档位**：80% 的 finding 不是档位内 bug，是档位选错了。看 `<out>_preflight.json` 或 `<out>_measurements.json` 里这个元素走了 text / shape / deco_snapshot / svg 哪条路径，对照症状：

| 症状 | 大概率原因 |
|---|---|
| 走 `deco_snapshot` 但视觉缺装饰 | 截图前的 hide JS 漏掉了什么；或 deco 区域比预期小 |
| 走 `text` 但样式没对上（颜色 / 字号 / 行距） | measure 抓 style 字段不全 / assemble 翻译漏字段 |
| 走 `shape` 但圆角 / border 不对 | border-radius 数值或 `_round_kind` 阈值 |
| 视觉效果（gradient / shadow / filter）完全丢 | **该走 deco_snapshot 但没走** — `hasComplexDecoration` 漏触发，加一行 |
| 文字消失 | 走了截图档但没发对应 text 记录；或 hide JS 把不该 hide 的也 hide 了 |

然后再走通常流程：

1. 看 Stage 5a 自检报告：哪几页被告警
2. 看 Stage 5b audit compare 图：用 Read 工具逐页对照
3. 搜 `references/lessons-learned.md` 已知症状（HTML 反模式 / OOXML 边界）
4. HTML 正常、PPT 异常且不符合已知边界 → 看是不是 skill 抽象不到位（按"看似 skill 内 bug 怎么处理"流程：HTML workaround 修当前 deck + 告知用户提 issue）

## 引用

- 维护扩展必读 → [`references/methodology.md`](./references/methodology.md) — 五步反假设流水线 checklist（作者扩展 skill 时读，agent 调用不必读）
- 历史踩坑修复 + HTML 反模式 + OOXML 边界 → `references/lessons-learned.md`（agent 排查必读）。本地工作副本（gitignored，首次 convert 自动从 `lessons-learned.md.example` seed），自由加 / 改 / 整理；想上游某条 → 作者手动复制回 `.example` 再 commit
