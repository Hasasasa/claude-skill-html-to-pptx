# Symptom Index

Use this file to choose the next debugging step after checking the self-check report and compare image. Entries are not final conclusions; they are prior fixes and likely inspection points.

## What goes here (and what doesn't)

This file is the **shared knowledge base** for HTML 反模式 (`HTML Anti-Patterns`) and OOXML 表达力边界 (`OOXML Limits`)—patterns where the right fix is **changing the source HTML**, not the skill code.

Existing sections below (Adapter, Text Extraction, Shapes, Fonts, Layout, etc.) are the **skill's debug history**—they record where past skill changes happened so future debugging can locate similar root causes faster. They are kept for diagnostic value, but **do not add new "skill bug → patched in measure.py" entries**. Per the SKILL.md "沉淀 HTML 问题与 OOXML 边界" policy:

- New shared entries go into the **HTML Anti-Patterns** or **OOXML Limits** sections at the bottom of this file.
- Skill-internal bugs (same CSS pattern breaks across decks) are **not patched ad-hoc** by agents. The current deck gets an HTML workaround; the agent reports the pattern to the user so the maintainer can decide whether to raise skill abstraction.
- Project / customer-specific patterns go into `lessons-learned-local.md` (gitignored).

## Quick Triage

1. Missing or blank slide: check adapter activation and hidden-state cleanup.
2. Missing text: check text-leaf detection and mixed inline/block containers.
3. Missing SVG/canvas/decorations: check screenshot marker records and generated PNG assets.
4. Wrong CJK/font rendering: check OOXML `a:latin` / `a:ea` and embedded font output.
5. Unexpected wrapping: check single-line detection and `bodyPr wrap`.
6. Wrong color or transparency: check rgba parsing and alpha propagation.
7. Wrong rotation: distinguish OOXML-drawn shapes/text from screenshot-based records.

## Adapter And Visibility

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| Later slides are blank or show slide 1 | `scrollIntoView` does not work for transform-based decks with `overflow:hidden` | Use the deck's own activation path, e.g. set `deck.style.transform = translateX(-idx*100vw)`. Check `adapter: xxx` output. |
| Non-current slides have only chrome/footer records | Animation CSS keeps `[data-anim]` at `opacity: 0` | Ensure measure adds low-power/no-animation state before extracting slides. |
| Wrong slide is measured for a framework | Adapter selector and activation behavior disagree | Each adapter should activate the target slide and mark it with `data-pptx-target`; extraction should read only that marker. |

## Text Extraction

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| SVG/IMG/CANVAS and nearby labels merge into one text record | `isTextLeaf` treats media tags as inline | Keep `SVG`, `IMG`, `CANVAS`, `VIDEO` in block-tag handling. |
| Direct text in a container disappears when the container also has block children | Walk recurses over element children and skips direct text nodes | Use range-based inline groups around block children to emit direct text. |
| Colored card/bar keeps text but loses its background | Text-leaf path draws only text and skips decoration | In text-box assembly, draw a synthesized background/border shape before drawing text. |
| Short labels like `01` wrap into two lines | PPT metrics are slightly wider than browser metrics | Disable wrapping for single-line records and add width buffer. |
| Hero word wraps despite `tf.word_wrap = False` | Later lxml code overwrites `<a:bodyPr wrap>` back to `square` | Set `wrap="none"` for single-line records and avoid overriding it later. |

## Shapes, Borders, Rotation

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| Single top border becomes a full rectangle border | Rectangle geometry cannot express one-sided border | Draw active border sides with connector lines; use rectangle border only when all sides are present. |
| CSS `border-radius: 50%` becomes square | Percent parsing or oval detection failed | Use `[:-1]` for `%`, `[:-2]` for `px`; emit oval geometry when radius implies circle/pill. |
| Rotated text or shape becomes axis-aligned | Rotation was not carried into OOXML | For OOXML-drawn shapes/text, compute cumulative rotation and apply `xfrm rot` around the AABB center using natural size. |
| Screenshot-based decoration looks double-rotated or squeezed | Pixel screenshot already includes rotation, then `_apply_rotation` rotated it again | Do not apply OOXML rotation to `deco_snapshot`, SVG picture, or canvas picture records. |
| Large rotated ribbon remains visibly different | Full-width rotated/skewed elements stress both geometry and element screenshot AABB | Treat as a known boundary only after compare image confirms this exact pattern. |
| HTML 里 `overflow:hidden` 容器装一堆 `transform:rotate` 子（cover 彩条、diagonal stripe 装饰等）在 PPT 里变成超大色块覆盖周围（旋转后子的 BCR AABB 远大于裁切框） | 子 ribbon 走普通 shape 路径用 BCR rectangle 画进 PPT，BCR 已经是 rotate 后的 AABB（一个 30px×600px 的 ribbon rotate -22° 后 AABB 变成 1528×714） | `hasComplexDecoration` 新加 `isClippingContainerWithTransformedChildren(s, el)`：`overflow:hidden/clip` + 至少一个子 transform 非平移 → 容器整体走 deco_snapshot；命中这条时 walker **return**（不下钻），否则子又会被单独画一遍。判定 transform 非平移用 `matrix(a,b,c,d,...)` 反解：`b≠0 ‖ c≠0 ‖ a≠1 ‖ d≠1` |
| PPT 端粗体文字出现"字符堆积 / 双绘"（每字符像被重叠画了两遍）；OOXML 里 typeface 是 `-apple-system` / `BlinkMacSystemFont` | CSS 系统字体别名（`-apple-system`、`BlinkMacSystemFont`、`-webkit-system-font`）是 CSS 端关键字解析为当前 OS UI 字体，不是真实字体名。写进 OOXML typeface= 后 PowerPoint / WPS 找不到 → 默认替换字体 → b="1" run 没粗体变体就靠"画两遍偏移"做 faux-bold。典型触发：HTML 里这些关键字出现在 `font-family` 栈最前面（包括 web component shadow CSS 通过继承泄漏到 body 文字的情形） | `scripts/assemble.py` 的 `GENERIC_FONT_KEYWORDS` 集合加 `-apple-system` / `blinkmacsystemfont` / `-webkit-system-font`——和 `ui-sans-serif` 等关键字一样跳过；first_font 会落到下一项真实字体（Helvetica Neue / Arial 等）。验证：dump_records.py 看 runs[i]['fontFamily'] 是否以这三关键字开头；解压 pptx 看 slide.xml `<a:latin typeface=...>` |
| slide 整页变成空背景（只剩底色，所有文字 / SVG / 装饰都不见）；measure 该页只有 1 条 deco_snapshot 记录 | slide 根元素本身命中 `isClippingContainerWithTransformedChildren` 早返回：slide 通常 `overflow:hidden`（裁视口）+ `::before` 伪元素装饰（纸面噪点 / 滤镜等）→ `hasComplexDecoration` 成立 → 再加任一直接子是 transform:rotate 的小装饰（pin SVG、scribble、徽章等），整片 slide 被吞成一个 PNG，deco hide-foreground 又把文字 SVG 全部隐藏，PNG 只剩底色 | `measure.py` walk 里 `isClippingContainerWithTransformedChildren` 早返回前加 `el !== slide` 守卫——slide 根的 overflow:hidden 是布局结构（裁视口），不是"装饰裁切意图"。修复后 slide 根仍发 deco_snapshot 抓底色装饰，但继续下钻子节点。验证：受影响的页 record 数应从 1 跳到正常多条（通常 8-20+），文字 / SVG 全回归 |

## Media And Decorations

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| SVG picture missing | Media was swallowed by text-leaf logic or marker screenshot failed | Confirm the record kind is `svg` and the referenced PNG exists. |
| SVG labels appear twice | Element screenshot captured absolutely positioned sibling labels | Hide non-SVG siblings before SVG screenshot, then restore them. |
| Background images, gradients, pseudo-elements, or shadows disappear | Complex CSS decoration has no OOXML mapping | Use `deco_snapshot` for elements with `background-image`, `box-shadow`, outline, or meaningful pseudo-elements; continue extracting children on top. |
| Canvas chart/WebGL area is blank | No canvas record or animation frame not stabilized | Add/verify canvas screenshot path and disable Chart.js animation before capture. |
| `<img>` 完全不出现在 PPT 里（卡片 emoji 图标、SVG `<img>` 等） | `assemble.py` 历史上 `kind=='img'` 分支写着 "第一版先跳过"，直接 `pass` | walker 给 `<img>` 打 `data-pptx-img-id` marker；measure 主循环按 marker `locator.screenshot(omit_background=True)`；`assemble.add_img_picture` 跟 svg picture 一样用 `add_picture(png, x, y, w, h)`。截图通道天然兼容远程 URL / cross-origin / SVG（OOXML 不能直嵌 SVG）/ data URI |
| 彩色 emoji（COLR/CPAL）在 PPT 里变黑白单色 | PowerPoint / WPS 字形渲染都不支持彩色字体，emoji 字符走文字通道必丢色 | 在 HTML 源里把 emoji 字符换成 `<img src=".../twemoji/{cp}.svg">`（CDN 用 `cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/svg/`；npm 路径已 404）。配 `.emoji-img { width:1em; height:1em; vertical-align:-0.125em; }` 让图跟字号。依赖上一行的 img marker screenshot 通道 |
| CSS `background-clip: text + linear-gradient`（文字渐变）在 PPT 端退化成单色 | OOXML 文字 fill 只支持纯色/图案，没有 text gradient clip 原语 | HTML 源里把渐变文字换成 inline `<svg>` + `<text fill="url(#grad)">` + `<linearGradient>`。skill 会把 SVG 走 svg picture 通道截图嵌入，渐变完美保留。SVG `viewBox` 给宽裕（如 320×64），`text-anchor="middle"` + `x="50%"` 让字居中对齐 |

## Fonts And Color

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| Chinese renders as square boxes | CJK characters are using the Latin font slot | Set OOXML `a:latin` and `a:ea` separately; CJK should route to Noto Sans/Serif SC. |
| Font weight is wrong after embedding | Variable fonts were embedded directly | Instance variable fonts into static TTFs before embedding. |
| PPTX is huge | Full CJK fonts were embedded instead of subsets | Subset fonts using characters found in measurement records. |
| Semi-transparent fills/lines become solid | Alpha channel was dropped during parse or OOXML emission | Preserve rgba alpha and write `<a:alpha val="...">` for fills and lines. |
| Text shadow is missing | Text shadow was not captured per run or OOXML effect order is wrong | Emit `a:effectLst` after fill and before font elements; remember hard shadows remain a rendering boundary. |
| WPS 打开字体不对，PowerPoint 打开字体对（嵌入字体在 WPS 失效） | WPS 不读 pptx 里裸 TTF 嵌入字体（只认 ECMA-376 obfuscated EOT） | 让用户用 `--install-user-fonts` 重转——字体装到 `%LOCALAPPDATA%\Microsoft\Windows\Fonts\` 后 WPS 直接拿系统字体。先 ask 用户授权再调，见 SKILL.md "字体安装确认"章节 |
| 自动解析的字体（GF）嵌入后 PowerPoint 显示成另一种字体 | GF 给 `wght@400` 的 src 可能指向 Medium(500) 文件，文件 nameID=1 = "Family Medium" 与 OOXML `typeface="Family"` 不匹配 → PowerPoint 拒绝加载嵌入字体回退系统 | `font_resolver._normalize_to_slot` 强制把 cache 里 TTF 的 nameID=1/4 改写成纯 family + 标准 slot 名，OS/2.usWeightClass 改成 400/700。`_cached_font_matches` 验证旧 cache 命中也得过这一关，不然死循环 |
| HTML 用 weight 500 + 600（无 400/700），PPT 里两个权重看起来一样 | resolver 早期把 500/600 都按 slot=regular 命名 → 后下载的覆盖前一个 | resolver 现在全局分配：non-italic faces 里 closest-to-400 进 regular slot，剩下 closest-to-700 进 bold slot。assemble.py 的 b="1" 阈值用 `weight >= 600` 兼容这种情况 |

## Layout / Anchor

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| HTML 用 flex+align-items:center 居中的文字，PPT 里跑到容器底部以下（"居中文字怎么靠下了"） | measure 抓的是容器 BCR（不是文字 BCR），assemble 又对 cy 做 h*1.4 撑高再设 anchor=ctr → 文字中心落到 1.4 倍高的中线 = 容器外 | `_text_box_size_px` 检测 `style.display in (flex,grid) && alignItems in (center, flex-end, end)` → cy 严格 = r.h，不再撑高。**不要**也加 anchor=t 兜底，那样真居中场景会偏上 |
| 文字横向居中跑到右边偏移 | textbox.w 加了 buffer 但 textbox.x 不变（buffer 全加到右边） | textbox.w 严格 = HTML BCR.w，buffer 通过 OOXML `tf.margin_left/right` 吸收 |
| 含 `<br>` 显式分行的多行段下方相邻段被盖住（如尾页"关键是..."段把下方链接段叠住） | 浏览器 BCR 已精确测出多行实际高度，assemble 又走"非单行 → r.h * 1.3"撑高 → cy 比 BCR 高 30%，覆盖下段 | `_text_box_size_px` 检测 `_has_explicit_break(rec)` → cy 严格 = r.h，与 flex/grid 居中分支同处理。自然换行（无 `<br>`）仍走 1.3x 防裁 |
| 含 `<br><br>` 空段的多行 `<p>`，PPT 里"关键是..."等末段被推出 textbox 砸到下方相邻元素 | OOXML 空 `<a:p>`（无 `<a:r>`）的 endParaRPr 没设 sz → PPT 用默认 18pt × 行距撑出 ≈43px 空行高，远大于正文 14-16pt × 行距 | `add_text_box` 写完所有 runs 后扫一遍空段，对每个空 `<a:p>` 设 `endParaRPr sz = rec.style.fontSize * 50`（OOXML sz 单位 = pt × 100，px → pt 乘 0.5）。让空段高度跟正文一致 |
| h1/h2 含 `<br>` + 内嵌 block（如 inline SVG），PPT 里 br 后的第二行文字被推下去 ~1 行（跟下方段重叠） | `emitInlineGroupsAround` 按 block 边界切 inline group 时把 `<br>` 当成普通 inline 元素并入下一组 → group=[br, textNode]，`range.setStartBefore(br)` 让 BCR 跨进上一行 → record h 翻倍 ≈ 2 行；同时 styleHost=br 让 record.tag 错标成 'br#inline' | `emitInlineGroupsAround` 在 range/styleHost/runs 步骤前，先把 group 首尾的 `<br>` 节点剥掉（`trimmed = group` 去 leading/trailing BR）。trimmed 进 range（修 BCR）+ 找 styleHost（修 tag）+ 喂 walkInline（避免 OOXML 多出空行 run）。**判据**：record.tag === 'br#inline' 永远是 bug |

## Counter / Animation Synchronization

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| 大数字标语段 PPT 里数字不对（如 HTML 终值 `42+ 20+ 7+ 100`，PPT 显示 `1+ 0+ 0+ 3`） | HTML 用 JS counter 动画（通用 `[data-target]` 元素从 0 渐变到 dataset.target，rAF 驱动 ≈1500ms），measure 在 activate_js 后立刻 evaluate(EXTRACT_JS) 抓 DOM，counter 还在中段 | `measure.py` 主循环 activate 后判断 `has_counter = document.querySelectorAll('[data-target]').length > 0`；有则 wait_for_function 等 `every(c => c.textContent.includes(c.dataset.target))` timeout 2500ms；超时兜底直接 `c.textContent = c.dataset.target`。空页 0 开销 |

## Self-Check Notes

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| Stage 5a skipped — "找不到可用的 pptx 渲染器" | PowerPoint COM and LibreOffice both unavailable | Install Office (Windows) or `apt install libreoffice` + `pip install pdf2image`. Without a renderer, Stage 5b audit cannot run either. |
| Stage 5a FULL-PIC warning | A `<p:pic>` covers ≥ ~98% of the slide — likely `deco_snapshot` double-layer bug | Inspect via `dump_records.py`; if confirmed, this is a skill-abstraction gap—apply an HTML workaround for the current deck and report to the maintainer. |
| Stage 5a LAYOUT warning | Two PPT text boxes overlap horizontally while HTML measurements show them apart | Inspect the slide's records via `dump_records.py`; usually a flex/grid gap got collapsed or font metrics differ enough to trigger. |

## HTML Anti-Patterns（HTML 反模式）

写源 HTML 时直接避开的通用模式。**新增 HTML 反模式条目放这里**。

| Pattern | Why it breaks PPT | HTML rewrite |
|---|---|---|
| _（待补充——任何写法导致跨 deck 转换失真的，沉到此处）_ | | |

## OOXML Limits（OOXML 表达力边界）

OOXML / PPT 渲染器天然无法精确表达的 CSS / DOM 模式。**新增 OOXML 边界条目放这里**。HTML 端必须走替代通路。

| CSS / DOM 模式 | OOXML 缺什么 | HTML 替代通路 |
|---|---|---|
| 彩色 emoji（COLR/CPAL 字体） | PowerPoint / WPS 字形渲染不支持彩色字体 | 换 Twemoji SVG `<img src="cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/svg/{cp}.svg">` + `.emoji-img { width:1em; height:1em; vertical-align:-0.125em; }`；走 img screenshot 通道 |
| CSS `background-clip: text + linear-gradient` 文字渐变 | OOXML 文字 fill 只支持纯色/图案，没有 text gradient clip 原语 | 换 inline `<svg>` + `<text fill="url(#grad)">` + `<linearGradient>`；走 SVG screenshot 通道。viewBox 给宽裕（如 320×64），`text-anchor="middle"` + `x="50%"` 居中 |
| `backdrop-filter: blur` / 复杂 filter | OOXML 无对应原语 | skill 已自动让该容器走 deco_snapshot 像素截图（无需 HTML 改）；若效果不接受则改 HTML 用纯背景色替代 |

## Debug Commands

`<skill_dir>` 是 skill 安装路径（通常 `~/.claude/skills/html-to-pptx/`）。

```bash
python <skill_dir>/convert.py <input.html> --keep-screenshots
python <skill_dir>/scripts/dump_records.py <measurements.json> <slide_idx>
```

`--keep-screenshots` retains HTML reference PNGs + `measurements.json` + `preflight.json` next to the output `.pptx`. Stage 5b audit material lands in `<out>_audit/` when a PPT renderer (PowerPoint COM / LibreOffice) is available and `--no-visual-audit` is not set; otherwise 5b is skipped (see SKILL.md "渲染器要求").
