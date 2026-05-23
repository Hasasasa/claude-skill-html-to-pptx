"""dump_records.py — measurement.json 调试打印工具。

Usage:
    python dump_records.py <measurements.json> [slide_idx]

slide_idx 从 0 起，默认 1（第 2 张）。每条 record 一行：id / kind / tag.className / 几何 / 文本预览。
"""
import json
import sys


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    d = json.load(open(sys.argv[1], encoding='utf-8'))
    idx = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    s = d['slides'][idx]
    print(f"slide {idx+1}: theme={s['slide']['theme']}, records={len(s['records'])}")
    for r in s['records']:
        txt = r.get('text', '') or ''
        txt = txt.replace('\n', ' ')[:60]
        print(f"  [{r['id']:>2}] {r['kind']:6} <{r.get('tag','')}.{r.get('className','')}> "
              f"x={r['rect']['x']:.0f} y={r['rect']['y']:.0f} w={r['rect']['w']:.0f} h={r['rect']['h']:.0f} "
              f"text={txt!r}")


if __name__ == "__main__":
    main()
