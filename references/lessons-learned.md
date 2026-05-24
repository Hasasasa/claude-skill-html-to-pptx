# Symptom Index

Use this file to pick the next debugging step after reading the self-check report and the compare image. Entries are not final answers; they record prior fixes and likely inspection points.

## What Goes Here

- Add new entries to **HTML Anti-Patterns** or **OOXML Limits** at the bottom.
- Project- / customer-specific patterns go to `lessons-learned-local.md` (gitignored).
- The earlier sections (Adapter, Text Extraction, Shapes, Fonts, Layout, etc.) are diagnostic history — read them when debugging, don't extend them.

## Quick Triage

1. Missing or blank slide: check adapter activation and hidden-state cleanup.
2. Missing text: check text-leaf detection and mixed inline/block containers.
3. Missing SVG / canvas / decorations: check screenshot marker records and generated PNG assets.
4. Wrong CJK / font rendering: check OOXML `a:latin` / `a:ea` and embedded font output.
5. Unexpected wrapping: check single-line detection and `bodyPr wrap`.
6. Wrong color or transparency: check rgba parsing and alpha propagation.
7. Wrong rotation: distinguish OOXML-drawn shapes / text from screenshot-based records.

## Adapter And Visibility

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| Later slides are blank or show slide 1 | `scrollIntoView` does not work for transform-based decks with `overflow:hidden` | Use the deck's own activation path, e.g. set `deck.style.transform = translateX(-idx*100vw)`. |
| Non-current slides have only chrome / footer records | Animation CSS keeps `[data-anim]` at `opacity: 0` | Ensure measure adds low-power / no-animation state before extracting slides. |
| Wrong slide is measured for a framework | Adapter selector and activation behavior disagree | Each adapter should activate the target slide and mark it with `data-pptx-target`; extraction should read only that marker. |

## Text Extraction

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| SVG / IMG / CANVAS and nearby labels merge into one text record | `isTextLeaf` treats media tags as inline | Keep `SVG`, `IMG`, `CANVAS`, `VIDEO` in block-tag handling. |
| Direct text in a container disappears when the container also has block children | Walk recurses over element children and skips direct text nodes | Use range-based inline groups around block children to emit direct text. |
| Colored card / bar keeps text but loses its background | Text-leaf path draws only text and skips decoration | In text-box assembly, draw a synthesized background / border shape before drawing text. |
| Short labels like `01` wrap into two lines | PPT metrics are slightly wider than browser metrics | Disable wrapping for single-line records and add width buffer. |
| Hero word wraps despite `tf.word_wrap = False` | Later lxml code overwrites `<a:bodyPr wrap>` back to `square` | Set `wrap="none"` for single-line records and avoid overriding it later. |

## Shapes, Borders, Rotation

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| Single top border becomes a full rectangle border | Rectangle geometry cannot express one-sided border | Draw active border sides with connector lines; use rectangle border only when all sides are present. |
| CSS `border-radius: 50%` becomes square | Percent parsing or oval detection failed | Use `[:-1]` for `%`, `[:-2]` for `px`; emit oval geometry when radius implies circle / pill. |
| Rotated text or shape becomes axis-aligned | Rotation was not carried into OOXML | For OOXML-drawn shapes / text, compute cumulative rotation and apply `xfrm rot` around the AABB center using natural size. |
| Screenshot-based decoration looks double-rotated or squeezed | Pixel screenshot already includes rotation, then `_apply_rotation` rotated it again | Do not apply OOXML rotation to `deco_snapshot`, SVG picture, or canvas picture records. |
| Large rotated ribbon remains visibly different | Full-width rotated / skewed elements stress both geometry and element-screenshot AABB | Treat as a known boundary only after the compare image confirms this exact pattern. |
| An `overflow:hidden` container holding many `transform:rotate` children (cover ribbons, diagonal stripes, etc.) becomes one huge color block in PPT that covers nearby content | Each child ribbon takes the normal shape path; its BCR is already the rotated AABB (a 30 px × 600 px ribbon rotated -22° has an AABB of 1528 × 714), so the painted rectangle leaks far beyond the clip box | In `hasComplexDecoration`, treat `isClippingContainerWithTransformedChildren(s, el)` as a deco trigger: `overflow:hidden/clip` plus at least one non-translation transform child → snapshot the container as one PNG and **`return`** from the walker (skip descent) so the children are not painted again. Detect a non-translation transform via `matrix(a,b,c,d,...)`: `b≠0 ‖ c≠0 ‖ a≠1 ‖ d≠1`. |
| PPT bold text shows "character pile-up / double draw" (each glyph looks painted twice); OOXML `typeface` is `-apple-system` / `BlinkMacSystemFont` | These CSS system aliases are keywords resolved by the browser to the host OS UI font; they are not real font names. Written to OOXML `typeface=`, PowerPoint / WPS cannot find them, falls back to a default, and fakes bold by painting the glyph twice at an offset. Typical trigger: those keywords appear at the head of a `font-family` stack (web component shadow CSS can leak them into body text via inheritance). | Add `-apple-system` / `blinkmacsystemfont` / `-webkit-system-font` to `scripts/assemble.py:GENERIC_FONT_KEYWORDS` so they are skipped like `ui-sans-serif`; `first_font` falls through to the next real family (Helvetica Neue / Arial / etc.). Verify with `dump_records.py` (check whether any `runs[i].fontFamily` starts with those keywords) and by unzipping the pptx (`slide.xml` `<a:latin typeface=...>`). |
| A whole slide goes empty (only background color; all text / SVG / decoration gone); `measure` reports a single `deco_snapshot` record for the page | The slide root itself triggers `isClippingContainerWithTransformedChildren` and the walker returns early. Slides typically have `overflow:hidden` (clipping the viewport) plus a `::before` pseudo-decoration (paper noise / filter), so `hasComplexDecoration` is satisfied; if any direct child also has `transform:rotate` (a pin SVG, scribble, badge), the whole slide is swallowed into one PNG. The deco hide-foreground pass then hides text / SVG, so the PNG only shows the background. | Add a `el !== slide` guard in `measure.py`'s walk before the early return for `isClippingContainerWithTransformedChildren`. The slide root's `overflow:hidden` is layout structure (viewport clipping), not "decorative clipping intent". After the fix, the slide root still emits a `deco_snapshot` for the background but continues descending. Verify: affected pages should jump from 1 record to a normal count (typically 8–20+) with text / SVG restored. |

## Media And Decorations

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| SVG picture missing | Media was swallowed by text-leaf logic or marker screenshot failed | Confirm the record kind is `svg` and the referenced PNG exists. |
| SVG labels appear twice | Element screenshot captured absolutely positioned sibling labels | Hide non-SVG siblings before the SVG screenshot, then restore them. |
| Background images, gradients, pseudo-elements, or shadows disappear | Complex CSS decoration has no OOXML mapping | Use `deco_snapshot` for elements with `background-image`, `box-shadow`, `outline`, or meaningful pseudo-elements; continue extracting children on top. |
| Canvas chart / WebGL area is blank | No canvas record or the animation frame was not yet stable | Add / verify the canvas screenshot path and disable Chart.js animation before capture. |
| `<img>` does not appear in PPT at all (card icon emoji, SVG `<img>`, etc.) | Historically `assemble.py`'s `kind=='img'` branch was a `pass` ("first pass — skip for now") | The walker marks `<img>` with `data-pptx-img-id`; the measure loop screenshots by marker (`omit_background=True`); `assemble.add_img_picture` uses `add_picture(png, x, y, w, h)` like SVG. The screenshot channel handles remote URLs, cross-origin, SVG (OOXML cannot embed SVG directly), and data URIs uniformly. |

## Fonts And Color

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| Chinese renders as square boxes | CJK characters use the Latin font slot | Set OOXML `a:latin` and `a:ea` separately; CJK should route to Noto Sans / Serif SC. |
| Font weight is wrong after embedding | Variable fonts were embedded directly | Instance variable fonts into static TTFs before embedding. |
| PPTX is huge | Full CJK fonts were embedded instead of subsets | Subset fonts using the characters found in measurement records. |
| Semi-transparent fills / lines become solid | Alpha channel was dropped during parse or OOXML emission | Preserve rgba alpha and write `<a:alpha val="...">` for fills and lines. |
| Text shadow is missing | Text shadow was not captured per run or OOXML effect order is wrong | Emit `a:effectLst` after fill and before font elements; remember hard shadows remain a rendering boundary. |
| WPS shows the wrong font even though PowerPoint shows the correct one (embedded fonts ignored by WPS) | WPS does not load raw TTFs embedded in pptx (it only recognizes ECMA-376 obfuscated EOT) | Re-run convert with `--install-user-fonts` — fonts go to `%LOCALAPPDATA%\Microsoft\Windows\Fonts\` so WPS reads them as system fonts. Ask the user first (see SKILL.md "字体安装确认"). |
| GF-resolved font is embedded but PowerPoint renders a different face | GF's `wght@400` src may point to a Medium (500) file whose `nameID=1 = "Family Medium"`. PowerPoint then matches `typeface="Family"` against the embedded font's name, fails, and falls back to a system font. | `font_resolver._normalize_to_slot` forcibly rewrites the cached TTF's `nameID=1/4` to a clean family + standard slot name and sets `OS/2.usWeightClass` to 400 / 700. `_cached_font_matches` also has to pass this check before reuse, otherwise it loops. |
| HTML uses weight 500 + 600 (no 400 / 700); both weights look the same in PPT | The resolver originally named both 500 and 600 with `slot=regular`, so later downloads overwrote earlier ones | The resolver now allocates globally: among non-italic faces, closest-to-400 fills the regular slot, then closest-to-700 fills bold. `assemble.py`'s `b="1"` threshold uses `weight >= 600` so 600-only stacks still get bolded. |

## Layout / Anchor

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| Text centered with `flex + align-items: center` in HTML drops below the container in PPT ("why is the centered text sitting at the bottom?") | `measure` captures the container's BCR (not the text's BCR), and `assemble` then expands `cy` to `h * 1.4` while applying `anchor=ctr`, so the text center lands on the midline of a 1.4× tall box — outside the container. | Detect `style.display in (flex, grid)` and `alignItems in (center, flex-end, end)` inside `_text_box_size_px`, then set `cy = r.h` exactly (no expansion). **Do not** add an `anchor=t` fallback; that breaks genuinely centered text upward. |
| Horizontally centered text drifts to the right | `textbox.w` was widened with a buffer but `textbox.x` was not adjusted (all the buffer landed on the right side) | Keep `textbox.w = HTML BCR.w`; absorb the buffer through OOXML `tf.margin_left / right` instead. |
| Multi-line paragraphs with explicit `<br>` cover the next sibling (e.g. the "key is…" paragraph covers the link beneath it) | The browser BCR already measures the true multi-line height; `assemble` then re-expands non-single-line content to `r.h * 1.3`, so `cy` exceeds the BCR by 30 % and overlaps the next paragraph. | In `_text_box_size_px`, detect `_has_explicit_break(rec)` and set `cy = r.h` exactly, same as the flex / grid centering branch. Natural wrapping (no `<br>`) still gets the 1.3× safety expansion. |
| A `<p>` with `<br><br>` empty paragraphs pushes its tail content out of the textbox and onto the next sibling | An empty OOXML `<a:p>` (no `<a:r>`) without an `endParaRPr.sz` falls back to PPT's default 18 pt × line-height (≈ 43 px), much taller than the body's 14–16 pt | After writing runs, `add_text_box` scans for empty paragraphs and sets `endParaRPr.sz = rec.style.fontSize * 50` on each (OOXML `sz` unit = pt × 100; px → pt is × 0.5), making empty paragraph height match the body. |
| An `h1` / `h2` with `<br>` plus an inline block (e.g. inline SVG): the line after the `<br>` slips down ≈ 1 line and overlaps the next paragraph | `emitInlineGroupsAround` splits inline groups at block boundaries; treating `<br>` as an ordinary inline element folds it into the next group → `group = [br, textNode]`. `range.setStartBefore(br)` makes the BCR cross into the previous line → record `h` doubles to ≈ 2 lines; meanwhile `styleHost = br` mistags the record as `'br#inline'`. | In `emitInlineGroupsAround`, strip leading / trailing `<br>` nodes from the group before computing the range / picking `styleHost` / walking inline runs. `trimmed = group` without the leading / trailing BRs goes into the range (fixes BCR), into `styleHost` (fixes the tag), and into `walkInline` (avoids spurious empty-line runs). **Tell**: a record with `tag === 'br#inline'` is always a bug. |

## Counter / Animation Synchronization

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| Big-number callouts show the wrong value in PPT (e.g. HTML final `42+ 20+ 7+ 100`, PPT shows `1+ 0+ 0+ 3`) | HTML uses a JS counter animation (generic `[data-target]` elements driven from 0 to `dataset.target` over ≈ 1500 ms by rAF). `measure` runs `evaluate(EXTRACT_JS)` right after `activate_js`, while the counters are still mid-animation. | After activation in the `measure.py` main loop, check `has_counter = document.querySelectorAll('[data-target]').length > 0`. If so, `wait_for_function` until every counter's `textContent` contains its `dataset.target`, timeout 2500 ms. On timeout, force `c.textContent = c.dataset.target` as a fallback. Pages without counters incur zero cost. |

## Self-Check Notes

| Symptom | Likely Cause | Inspect / Fix |
|---|---|---|
| Stage 5a / 5b skipped — no pptx renderer available | PowerPoint COM and LibreOffice are both unavailable | Install Office (Windows) or `apt install libreoffice` + `pip install pdf2image`. Without a renderer, Stage 5b audit cannot run either. See SKILL.md renderer-requirement section for the user-consent flow. |
| Stage 5a FULL-PIC warning | A `<p:pic>` covers ≥ ~98 % of the slide — likely a `deco_snapshot` double-layer bug | Inspect with `dump_records.py`; if confirmed, this is a skill-abstraction gap — apply an HTML workaround for the current deck and report the pattern to the maintainer. |
| Stage 5a LAYOUT warning | Two PPT text boxes overlap horizontally while HTML measurements show them apart | Inspect the slide's records via `dump_records.py`; usually a flex / grid gap got collapsed or font metrics differ enough to trigger. |
| Stage 5b audit compare 图大面积字体回退（serif / mono 出现在本该是 Bricolage / Inter / Space Grotesk 的位置），同时多页报 [HIGH] "标题与正文叠压" | **PowerPoint COM `slide.Export()` 不读 pptx 内嵌的 TTF 字体**——只用系统已装字体。回退字体（Times / Courier）比设计字体宽，标题强制多换行 → 挤进下方段落 → 假叠压。pptx 内嵌字体本身正确（slide.xml `typeface=` 对、`/ppt/fonts/fontN.fntdata` 都有、nameID 也对得上），不是 skill bug | **不要**先去诊断 OOXML / nameID / lessons-learned 的 GF Medium-as-Regular 条目——那些都不是这次的根因。直接告诉用户 "audit 渲染管线本身不读内嵌字体，需要把字体装到用户字体目录"，征得同意后用 `--install-user-fonts` 重跑。装完后 PowerPoint COM 通过系统字体路径找到这些字体，audit 渲染等于交付效果。SKILL.md "字体安装确认" 段要求**第一次 convert 之前**就 ask，从源头避免这一轮浪费 |

## What The Skill Already Covers

The converter has dedicated handling for the following features. If they look broken, debug rather than reporting them as unsupported:

- `background-image`, CSS gradients, `box-shadow`, `outline`, `::before` / `::after` decorations — via `deco_snapshot`.
- SVG, canvas, `<img>` — via screenshot picture records.
- `border-radius: 50%` / pill shapes — via oval geometry where applicable.
- Single-side borders — via line connectors.
- CJK font fallback — separate OOXML `a:latin` and `a:ea` slots.
- Text shadow — via OOXML `outerShdw` (hard / stacked shadows remain a boundary, see below).
- Non-bundled fonts — auto-resolved from Google Fonts on demand; CJK families are pulled via the variable-font direct-download path.

## HTML Anti-Patterns

Generic HTML patterns the source should avoid.

| Pattern | Why It Breaks PPT | HTML Rewrite |
|---|---|---|

## OOXML Limits

CSS / DOM patterns that OOXML or the PPT renderer cannot express precisely; the HTML source must take an alternate channel.

| CSS / DOM Pattern | What OOXML Lacks | Alternate HTML Channel |
|---|---|---|
| Colored emoji (COLR / CPAL fonts) | PowerPoint / WPS glyph rendering does not support color fonts; emoji characters lose color through the text channel | Replace emoji characters with Twemoji SVG `<img src="https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/svg/{cp}.svg">` and style with `.emoji-img { width: 1em; height: 1em; vertical-align: -0.125em; }`. The skill embeds the result via the `<img>` screenshot channel. |
| CSS `background-clip: text` + `linear-gradient` text gradient | OOXML text fill only supports solid color / pattern; there is no text-gradient-clip primitive | Replace gradient text with inline `<svg>` + `<text fill="url(#grad)">` + `<linearGradient>`. The skill embeds the SVG via the SVG screenshot channel. Give the `viewBox` some slack (e.g. 320 × 64), and use `text-anchor="middle"` + `x="50%"` for centering. |
| `backdrop-filter: blur` / complex `filter` | OOXML has no equivalent primitive | The skill already routes such containers through `deco_snapshot` pixel screenshots, so the visual is preserved (no HTML change needed). If the resulting flat look is unacceptable, substitute a solid background color in the HTML. |
| Multi-layer hard `text-shadow` (stacked zero-blur shadows) | OOXML supports a single `outerShdw`, not stacked hard shadows | Reduce to a single light shadow in HTML, or accept the simplification. Single soft shadows still convert well. |
| Tight CSS `line-height` (e.g. `line-height: .85`) | PPT text layout is consistently looser than the browser; small line-heights diverge most visibly | Single-line titles are usually fine. For multi-line titles either accept a slightly looser PPT layout or split the line in HTML so it stays single-line per `<p>`. |
| Chinese italic (`font-style: italic` on CJK characters) | CJK families generally lack a true italic; PPT renders upright or faux italic | If only Latin text needs italic, mark just those runs `italic`. |
| Inline flex / grid `gap` used between spans of one text leaf | The text leaf exports as one record; PPT receives ordinary spaces instead of CSS gap | Split the spans into separate block / inline-block elements so each becomes its own record, or accept slightly tighter spacing. |
| `<video>` element with content | Video frames are not extracted; only static placeholder geometry is emitted | Replace with a still image (`<img>`) in HTML. |
| WebGL / animated `<canvas>` | A static frame is captured; animation and interactivity are lost | Pre-render the desired frame and substitute an `<img>`, or accept the static capture. |
| Large rotated full-width ribbons | Both the geometry path and the element-screenshot AABB stress on extreme rotations; small visual differences are expected | Confirm the ribbon is not unintentionally hiding foreground content. Accept minor differences or restructure the decoration as a `deco_snapshot` of a smaller clipped container. |
| Font not available on Google Fonts (commercial, self-hosted, typo) | `font_resolver` cannot fetch it, so the family falls back to the viewer's system font | Verify the family name spelling. Fix the name or replace the font with a GF-available alternative. The `[font-resolve]` warning at convert time flags this. |

## Viewer-Side Boundaries (Not Conversion Bugs)

These are not converter bugs but expected behavior on the viewing side:

- **PowerPoint trust prompt for embedded fonts** — normal for PPTX files with embedded fonts; the user can trust the document if the source is trusted.
- **iOS / web PowerPoint ignores embedded fonts** — embedded fonts are not honored in those environments. Recheck in desktop PowerPoint before changing converter code.
- **Browser-only interactivity** (motion, scroll, touch) — PPT keeps slide pages, not browser interactions. Verify that the static slide state is captured correctly.

## Debug Commands

`<skill_dir>` is the skill install path (typically `~/.claude/skills/html-to-pptx/`).

```bash
python <skill_dir>/convert.py <input.html> --keep-screenshots
python <skill_dir>/scripts/dump_records.py <measurements.json> <slide_idx>
```

`--keep-screenshots` retains HTML reference PNGs + `measurements.json` + `preflight.json` next to the output `.pptx`. Stage 5b audit material lands in `<out>_audit/` when a PPT renderer (PowerPoint COM / LibreOffice) is available and `--no-visual-audit` is not set; otherwise 5b is skipped (see SKILL.md "渲染器要求").
