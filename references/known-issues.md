# Known Boundaries

Use this file after comparing the HTML reference screenshot with the PPT render. Do not treat an item here as a final answer until the observed symptom matches the listed condition.

## Diagnosis Rule

1. If the HTML reference screenshot already shows the problem, report it as a source HTML/layout issue.
2. If the HTML reference is correct but the PPT render differs, inspect the compare image and `lessons-learned.md`.
3. If the symptom matches a boundary below, explain the boundary and any practical workaround.
4. If it does not match, continue debugging the converter rather than assuming it is unsupported.

## Current Boundaries

| Area | When This Applies | User-Facing Explanation | Next Check |
|---|---|---|---|
| Tight CSS `line-height` | Multi-line titles use very small line-height, e.g. `line-height: .85` | PPT text layout may be looser than the browser. Single-line titles are usually unaffected. | Confirm HTML reference is tight and PPT is only looser, not missing content. |
| Chinese italic | Chinese characters use CSS `font-style: italic` | PPT may render CJK italic as upright or faux italic. | If only Latin text is italic, this should not apply. |
| Inline flex/grid gap | Multiple inline spans rely on CSS `gap` and are exported as one text leaf | Text spacing may be slightly tighter because PPT receives ordinary spaces. | Check whether records merged spans into one text record. |
| PowerPoint trust prompt | PowerPoint warns about embedded fonts | This is normal for PPTX files with embedded fonts. The user can trust the document if the source is trusted. | No converter change needed. |
| iOS / web PowerPoint fonts | PPT is opened in iOS or browser PowerPoint | Embedded fonts may be ignored, causing system font fallback. | Recheck in desktop PowerPoint before changing converter code. |
| Font not on Google Fonts | HTML uses a font that GF does not host (commercial / self-hosted / typo) | font_resolver fails to fetch; the deck may fall back to installed system fonts on the viewing machine. | Check the `[font-resolve]` warning at convert time and confirm the family spelling. |
| WebGL / animated canvas | Content is rendered in canvas/WebGL | The converter captures a static frame, not animation or interaction. | If the canvas is blank, check `lessons-learned.md` canvas entry. |
| Motion / scroll / touch interaction | Deck relies on browser-only interactivity | PPT keeps slide pages, not browser interactions. | Verify static slide state is captured correctly. |
| `<video>` | Important content is inside video elements | Video frames are not currently converted. | Consider replacing with a still image in source HTML. |
| Large rotated ribbons | Very large full-width rotated/skewed bands dominate a slide | Geometry and screenshot paths both have edge cases; expect visible differences. | Confirm it is not hiding foreground content unexpectedly. |
| Multi-layer hard `text-shadow` | CSS uses stacked zero-blur shadows for hard duplicate text | OOXML supports only a softer outer shadow, so exact hard-shadow effects may differ. | Single light text shadows should still convert well. |
| `backdrop-filter: blur` | Design relies on CSS backdrop blur | Background color is preserved, but blur itself has no direct OOXML equivalent. | Check whether visual difference is acceptable in compare image. |

## Covered Features

The converter has specific handling for these features. If they fail, debug rather than reporting them as unsupported:

- `background-image`, CSS gradients, `box-shadow`, `outline`, and `::before` / `::after` decorations via `deco_snapshot`.
- SVG and canvas via screenshot picture records.
- `border-radius: 50%` / pill shapes via oval geometry where applicable.
- Single-side borders via line connectors.
- CJK font fallback through separate OOXML `latin` and `ea` font slots.
- Text shadow through OOXML `outerShdw`, with the hard-shadow caveat above.

## Useful Checks

- Use `--keep-screenshots` to retain HTML reference PNGs + `measurements.json` next to the output.
- Stage 5b audit material lands in `<out>_audit/` automatically; open the `slide_NN_compare.png` files and follow `audit_prompt.md`.
- Inspect `measurements.json` when a slide is missing content; low record counts often point to hidden-state or adapter issues.
- Run `python scripts/dump_records.py <measurements.json> <slide_idx>` to list record kinds and positions.
