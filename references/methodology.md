# 反假设流水线

维护扩展 skill 时的工作准则。每加一个新模板出问题，先对照这份 checklist 定位到哪一步偷了假设。

```
[1 探测] → [2 强制] → [3 抽取] → [4 兜底栅格] → [5 验证]
   inspect    force      extract     fallback       verify
```

每步三件事：**动作 / 反假设规则 / 退出前断言**。

---

## 1 探测（preflight / discover）

**动作**：进入 DOM 实测一切——
- slide 元素是谁
- 翻页 / 移动容器在哪
- 各 slide 的自然 display
- 字体实际渲染回退到了什么
- 已知风险模式（多层 text-shadow / backdrop-filter / video / canvas / 非内置字体）

**反假设规则**：决策只能依赖运行时可观察事实——
- ✅ `getComputedStyle` / `getBoundingClientRect` / 子树结构相似性 / 祖先链 CSS 属性
- ❌ 写死的 selector / 类名 / ID / `window.<libName>` / body 特定 class 约定

**断言**：slide 数 ≥ 1；每张可用 `data-pptx-target` 定位。

---

## 2 强制（force / activate）

**动作**：主动把目标置于可测量状态——
- force-position 到视窗：`position:fixed; inset:0; z-index:max; !important`
- 清空所有祖先 transform（避免 fixed 不相对视窗）
- 注入入场动画兜底 CSS：`[data-anim], [data-aos], [data-reveal], .fade-in` 等强制 `opacity:1; transform:none; visibility:visible; filter:none`
- 关掉所有 transition / animation 时长

**反假设规则**：不依赖页面"原本怎么切页 / 怎么揭示动画"——要它在哪儿就推到哪儿。

**断言**：目标 slide BCR 在视窗内；大小 ≥ 50% viewport；opacity > 0；visibility ≠ hidden；不被祖先 overflow hidden 裁。

---

## 3 抽取（measure / extract）

**动作**：走 DOM 拿能用 OOXML 表达的部分——
- 文字 runs（font / color / size / shadow）
- 形状（背景 / 边框 / 圆角）
- 几何（位置 / 尺寸 / 旋转）
- 媒体引用（img / svg / canvas marker）

**反假设规则**：
- BCR / offsetWidth 直接当真
- 不替源 HTML 判断布局意图

**断言**：可编辑记录数 ≥ 合理下界（hero 页 ≥ 3，内容页 ≥ 5）；每条 rect 在 slide 范围内。

---

## 4 兜底栅格（snapshot / rasterize fallback）

**动作**：OOXML 表达不了的（多层 text-shadow / backdrop-filter / 复杂渐变 / canvas / SVG），截图嵌入。

**反假设规则**：
- 截图前显式隐藏所有"将在第 3 步独立绘制的元素"（文字 / 媒体 / 其他装饰），否则被烘焙进 PNG → 双层
- 隐藏用 inline `!important`：`el.style.setProperty('visibility', 'hidden', 'important')`。inline + `!important` 是 cascade 最高优先级，beat 任何 CSS `!important`
- 还原要按记录的优先级恢复：`style.setProperty(name, value, prio)`
- 跳过 deco 自身和祖先链（祖先背景要透到 deco 截图）

**断言**：PNG 不含会被第 3 步重画的内容；尺寸符合 deco 元素 BCR。

---

## 5 验证（verify）

两层核查：

**5a 结构化扫描**（廉价规则化）
- 全屏 `<p:pic>` 嫌疑（cx ≥ 12M EMU 且 cy ≥ 6.8M EMU）→ FULL-PIC
- 文本框横向重叠 → LAYOUT
- 合并 preflight 高风险页 → PREFLIGHT
- 不做像素 diff —— 数字给假信心，对局部 bug 不敏感

**5b 视觉 audit**（VLM 判断）
- 产出每页 HTML | PPT 双栏对比图 + `audit_prompt.md`
- 上游 agent 逐页 Read 图，按 prompt 检查清单写 `audit_findings.md`
- 迭代修复直到全 OK 或仅 LOW

**反假设规则**：
- 不在 skill 里写 collision avoidance 死规则——规则覆盖窄、副作用大
- 视觉判断外包给 VLM
- 5a 通过 ≠ 视觉可用，必须人工或 VLM 复核

**断言**：5a 三类告警都记录；5b findings 写完每页要么 OK 要么有问题描述；HIGH/MID 全修完或归入 known-boundaries 告知用户。

---

## 新模板 checklist

- [ ] 1 探测：slide 数对吗？mover 找对了吗？BCR 能读吗？
- [ ] 2 强制：activate 后 BCR 在视窗？`[data-anim]` 类元素强制可见？
- [ ] 3 抽取：记录数和肉眼数大致对得上？没漏抽？
- [ ] 4 兜底栅格：deco PNG 里有没有文字 / 图标 / 前景元素？
- [ ] 5 验证：5a 告警对得上 HTML 实际状态？5b audit 跑了吗？

任一项答 No，定位到对应步骤补反假设规则。**不要在下游打补丁。**

---

## 适用范围

这套不只 HTML → PPT 适用。任何"源格式能力 > 目标格式能力"的转换都按这五步走：
- 源能力无限（CSS / SVG / Figma），目标能力有限（OOXML / PDF / Markdown）
- 中间显式承认落差：哪些可编辑提取（3）、哪些只能栅格兜底（4）、哪些根本表达不了（5 告警）
