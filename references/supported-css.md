# 支持的 CSS 子集

> 本项目的策略：让 Chromium（Playwright）先把 HTML 完整渲染好，再读 **computed style** 和 **bounding rect** 装配进 PPTX。
> 所以"支持的 CSS"实际上分四档：矢量保留 / 栅格兜底 / 已驯化 / 已知缺口。

## 第 1 档：矢量保留（OOXML 原生表达）

读出 computed 值，直接画成 PPT 形状/文字。这一档保留可编辑、可搜索、可缩放。

### 文字（`measure.py:EXTRACT_JS` → `assemble.py:add_text_box`）

| CSS 属性 | OOXML 落点 |
|---|---|
| `font-family`（含 Google Fonts / CJK fallback） | `<a:rFont>` + `<a:ea>` 中英分立 |
| `font-size` | `sz=` 半磅 |
| `font-weight` 300-900 | `b="1"` 阈值 ≥600；细字按字体变体 |
| `font-style: italic` | `i="1"` |
| `color`（hex / rgb / rgba） | `<a:solidFill>` + alpha |
| `line-height` | `<a:spcPct>` |
| `letter-spacing` | `spc=` 百分点 |
| `text-align: left/center/right` | `algn="l/ctr/r"` |
| `text-transform: uppercase/lowercase` | run 文本预处理 |
| `text-decoration: underline` | `u="sng"` |
| `text-shadow`（第一层） | `<a:effectLst><a:outerShdw>` |
| `padding`（textbox 内边距） | `marginLeft/Right/Top/Bottom` |
| `opacity`（文字层） | run/fill alpha |
| `display: flex/grid` + `align-items` | `anchor="t/ctr/b"`（垂直锚点）|
| `display: flex/grid` + `justify-content` | 覆盖 text-align 当水平锚点 |
| **多 run 富文本** | 同段内多个 `<a:r>`，各带独立字体颜色 |
| `::before` / `::after` `content: "..."`（string literal） | 通过 `getComputedStyle(el, pseudo).content` 抽取，作为 run 拼在父 textbox 头/尾。常见用途：装饰前缀（↑↓ 箭头、列表 marker、引号、徽章"NEW"） |
| `writing-mode: vertical-rl` / `vertical-lr` / `sideways-*` | `<a:bodyPr vert="eaVert">`，CJK 竖排可用；纯 Latin 竖排视觉可能略偏 |

### 形状（`measure.py` shape 分支 → `assemble.py:add_shape_box`）

| CSS 属性 | OOXML 落点 |
|---|---|
| `background-color`（纯色 + alpha） | `<a:solidFill>` |
| `border-{top/right/bottom/left}-{width/color/style}` | 四边分别画线，宽度+颜色+样式 |
| `border-radius`（50% → 圆 / 椭圆；px → 圆角矩形） | `MSO_SHAPE.OVAL` / `ROUND_RECTANGLE` |
| `transform: rotate`（含祖先链累加，从 matrix 反解） | `<a:xfrm rot=...>` |

## 第 2 档：栅格兜底（`deco_snapshot` — 截图嵌入 + 文字叠上）

`_js_snippets.py:hasComplexDecoration` 命中以下任一条件，浏览器截下元素整块图作为底层，子节点的文字 / SVG 再画在之上：

1. `background-image` 任何值（linear/radial/repeating/conic-gradient、`url(...)`、SVG data URI）
2. `box-shadow` 任何值（含硬阴影 `8px 8px 0 #color`）
3. `outline`
4. `::before` / `::after` **装饰**内容（空 content + background-image / box-shadow 等；纯 string content 走文字通道见第 1 档表）
5. `overflow:hidden` + 含 transformed 子（裁切框 + 旋转子）
6. **`backdrop-filter`** 任何值（毛玻璃）— 2025-05 新增
7. **`filter`** 任何值（blur / drop-shadow / saturate / 等）— 2025-05 新增
8. **`mix-blend-mode`** 非 normal — 2025-05 新增
9. **不可矢量表达的 transform**（skew / 非均匀 scale）— 2025-05 新增

栅格兜底的工作机制（`_DECO_HIDE_FOREGROUND_JS`）：

1. 截图**前**，把元素内所有带文字的子元素 `visibility: hidden`，把元素自身的文字 `color: transparent`
2. 截图——拿到的是"只有装饰、没有文字"的 PNG
3. 截图**后**恢复
4. 文字在装配阶段以**矢量 textbox 画在截图之上**——保持可编辑、可搜索、可缩放

**所以"栅格兜底"只栅格化装饰像素，文字始终是矢量。**

代价：filter / mix-blend-mode / skew 这些"应同时作用于文字"的效果，在 PPT 里只作用于背景，文字保持 crisp。客户绝大多数场景偏好可编辑文字而不是"完全保真但锁死成图片"。如果某个 slide 需要后者，可以直接把 HTML 整页截图。

## 第 3 档：已驯化（不需要"支持"，主动消解）

| 难题 | 消解策略 | 位置 |
|---|---|---|
| `@keyframes` / `animation` / `transition` | preflight 阶段全局 kill，永远拿终止帧 | `adapters.PREPARE_JS` |
| `scroll-snap` 翻页容器 | force-position CSS 把每页强制定位回视口 | `adapters.PREPARE_JS` |
| `transform: translateX(N*100vw)` 翻页 deck | 同上 | 同上 |
| 手写 JS counter（`data-target` / `data-count-to` / `data-counter`） | 等终值或强制设终值再 measure | `measure.py` |
| canvas 入场动画（Chart.js / WebGL） | 等连续两帧像素 hash 一致再截首帧 | `measure.py:has_canvas` 块 |

## 第 4 档：媒体元素

| 元素 | 处理 |
|---|---|
| `<svg>` | 整块截图（PNG，omit_background=True），同时保留 outerHTML 作后备 |
| `<canvas>` | 等稳定后截首帧 |
| `<video>` | 截首帧像素作静态 picture 嵌入（OOXML 无原生播放表达；走 canvas 同一通路，kind='canvas' tag='video'）|
| `<img>` | 直接取 `currentSrc` 嵌入原图（保持原分辨率）|

## 第 5 档：已知缺口（暂不支持，遇到只能裸奔）

| 缺口 | 影响 | 处置 |
|---|---|---|
| `transform: matrix3d` 真 3D | 反解只取 2D rot 分量，3D 透视丢 | 不修复（PPT 无对应原语）|
| SVG `<filter>` 复杂滤镜（feTurbulence / 多滤镜叠加）| 浏览器渲染失败时整块 SVG 模糊或缺失 | 等浏览器修复 |
| `clip-path: polygon` 复杂形状 | 走栅格兜底没问题，但若元素本身没触发 deco（极少见）会显示完整矩形 | 添加 clip-path 到 hasComplexDecoration（待）|
| iframe / audio | 不抽取 | 显式忽略 |

## 落地实战清单（给开发者）

1. **遇到 bug 报告先问"哪个 slide、哪个元素"**：用 `out/<deck>/preflight.json` 看该元素是不是走了 deco_snapshot。
2. **新发现的不支持 CSS**：
   - 视觉装饰类（gradient、blend、filter）→ 加进 `hasComplexDecoration` 触发集即可，不写新代码
   - 文字属性类（writing-mode、列布局）→ 在 `measure.py` style 字典里加字段，在 `assemble.py:add_text_box` 里翻译
   - 几何变换类（skew / 3D）→ 走 deco_snapshot 截图兜底，几乎不用专门修
