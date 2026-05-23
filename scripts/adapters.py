"""adapters.py — slide 发现 + 强制激活。

纯能力驱动，单一启发式通路：
- 发现：用户显式 `[data-pptx-slide]` 优先；否则按"同 tag + 视窗尺寸级别"启发式找兄弟组
- 激活：打通用类 + 自动找 transition:transform mover translate + force-position CSS 兜底
- 不依赖标签名 / 类名 / ID / 库名 / body 类约定

详见 references/methodology.md 第 1-2 步。
"""


# 通用动画杀手：把 transition / animation 时长归零 + 强制业界主流入场动画约定可见
ZERO_ANIMATIONS_CSS = r"""
    {
        const styleEl = document.createElement('style');
        styleEl.id = 'pptx-zero-anim';
        styleEl.textContent = `
            *, *::before, *::after {
                animation-duration: 0.0001s !important;
                animation-delay: 0s !important;
                transition-duration: 0s !important;
                transition-delay: 0s !important;
            }
            /* 业界主流"入场动画前"约定全部强制还原为最终可见态 */
            [data-anim], [data-animate], [data-aos], [data-scroll],
            [data-motion], [data-reveal], [data-fade], [data-stagger],
            .reveal, .fade-in, .animate-in, .aos-init, .motion-element {
                opacity: 1 !important;
                transform: none !important;
                visibility: visible !important;
                filter: none !important;
            }
        `;
        document.head.appendChild(styleEl);
    }
"""


# 统一发现 slide：跑一次得出 window.__pptxSlides 数组 + window.__pptxNaturalDisplay
DISCOVER_JS = r"""
() => {
    // 1) 用户显式标注优先：[data-pptx-slide]
    const explicit = document.querySelectorAll('[data-pptx-slide]');
    let group = explicit.length >= 1 ? Array.from(explicit) : null;

    // 2) 自动启发：找一个 parent，使其同 tag 兄弟组中：
    //      - 至少 1 个 BCR ≥ 50% viewport（说明它们是"页级"元素）
    //      - 数量越多越优先
    //      - 同 tag 限制是为了排除 header / footer 等异类
    if (!group) {
        const vw = window.innerWidth, vh = window.innerHeight;
        const minW = vw * 0.5, minH = vh * 0.5;
        const candidates = [];
        for (const el of document.body.querySelectorAll('*')) {
            const r = el.getBoundingClientRect();
            if (r.width >= minW && r.height >= minH) candidates.push(el);
        }
        let bestGroup = [], bestScore = 1;
        for (const cand of candidates) {
            const parent = cand.parentElement;
            if (!parent || parent === document.documentElement) continue;
            const sameTag = Array.from(parent.children).filter(
                ch => ch.tagName === cand.tagName
            );
            if (sameTag.length > bestScore) {
                bestScore = sameTag.length;
                bestGroup = sameTag;
            }
        }
        if (bestGroup.length === 0 && candidates.length >= 1) {
            // 单页兜底：取最深的视窗尺寸元素
            candidates.sort((a, b) => {
                let da = 0, db = 0;
                for (let c = a; c; c = c.parentElement) da++;
                for (let c = b; c; c = c.parentElement) db++;
                return db - da;
            });
            bestGroup = [candidates[0]];
        }
        group = bestGroup;
    }

    // 探测"自然 display"：找当前可见（display!=none）的 slide 的 display 值
    // 用于 activate 时覆盖 `display:none` 的隐藏，但保留模板的 flex/grid 等布局
    let naturalDisplay = 'block';
    for (const s of group) {
        const d = getComputedStyle(s).display;
        if (d && d !== 'none') { naturalDisplay = d; break; }
    }

    window.__pptxSlides = group;
    window.__pptxNaturalDisplay = naturalDisplay;
    return group;
}
"""


# 注入"force-position"覆盖 CSS（仅适用于 [data-pptx-target] 的 slide）
# 用高优先级 !important 把目标 slide 钉到 (0,0,100vw,100vh)，beat 任何 .slide:not(.active) 之类规则
FORCE_POSITION_CSS = r"""
    {
        const s = document.createElement('style');
        s.id = 'pptx-force-position';
        s.textContent = `
            [data-pptx-target] {
                position: fixed !important;
                top: 0 !important;
                left: 0 !important;
                right: 0 !important;
                bottom: 0 !important;
                width: 100vw !important;
                height: 100vh !important;
                margin: 0 !important;
                transform: none !important;
                opacity: 1 !important;
                visibility: visible !important;
                z-index: 2147483647 !important;
            }
        `;
        document.head.appendChild(s);
    }
"""


# 激活第 idx 张 slide（不假设模板，多手段并用）：
#  1. 清空所有 slide 的 data-pptx-target / .is-active / .active；目标加上
#  2. 若有 mover（祖先 transition:transform），自动判断 X/Y 轴翻页方向并 translate
#  3. raf×2 后检查目标 BCR 是否在视窗；不在则 force-position 兜底
ACTIVATE_JS = r"""
(idx) => {
    const slides = window.__pptxSlides || [];
    if (!slides[idx]) return { error: 'index out of range' };
    const target = slides[idx];
    const naturalDisplay = window.__pptxNaturalDisplay || 'block';

    // Step 1: 还原上一次 activate 留下的标记 / 内联样式
    for (const s of slides) {
        s.removeAttribute('data-pptx-target');
        s.classList.remove('is-active', 'active');
        s.style.removeProperty('display');
    }
    // 还原上一次 activate 清空过的祖先 transform
    if (window.__pptxClearedAncestors) {
        for (const [el, prev] of window.__pptxClearedAncestors) {
            if (prev === '') el.style.removeProperty('transform');
            else el.style.setProperty('transform', prev);
        }
    }
    window.__pptxClearedAncestors = [];

    // Step 2: 给目标打通用激活类 + force-position 触发标记
    target.setAttribute('data-pptx-target', '');
    target.classList.add('is-active', 'active');
    target.style.setProperty('display', naturalDisplay, 'important');

    // Step 3: 关键——清空目标 slide 所有祖先的 transform。
    // 原因：CSS spec 规定 position:fixed 的"包含块"是最近一个有 transform / filter /
    // perspective / will-change 等属性的祖先；没有这些时才是 viewport。
    // 如果不清，force-position 的 inset:0 会被祖先 transform 当成参考系而不是视窗。
    let cur = target.parentElement;
    while (cur && cur !== document.body) {
        const cs = getComputedStyle(cur);
        if (cs.transform && cs.transform !== 'none') {
            window.__pptxClearedAncestors.push([cur, cur.style.transform || '']);
            cur.style.setProperty('transform', 'none', 'important');
        }
        cur = cur.parentElement;
    }

    // Step 4: 强制 reflow 以便后续 BCR 读取拿到最新位置
    target.getBoundingClientRect();
    return { ok: true };
}
"""


# 一次性准备：注入动画兜底 CSS + force-position CSS + 跑 slide 发现
PREPARE_JS = """
    () => {
        """ + ZERO_ANIMATIONS_CSS + """
        """ + FORCE_POSITION_CSS + """
        (""" + DISCOVER_JS + """)();
    }
"""

ENUMERATE_JS = "() => (window.__pptxSlides || []).length"
