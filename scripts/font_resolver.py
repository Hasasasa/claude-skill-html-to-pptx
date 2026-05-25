"""font_resolver.py — 按需自动解析并下载 HTML 用到的字体。

工作流：
1. collect_requested_fonts(meas) → 从 measurement 抽 {(family, weight, italic)} 集合
2. resolve_fonts(needed, bundled_aliases) → 对每个未内置家族查 Google Fonts CSS API
   - 命中：下载所有 weight/italic .ttf 到 CACHE_DIR，组织成 FONT_PLAN 条目
   - 未命中（GF 没有此家族）：进 unavailable 列表，标记到 _resolved.json 避免重 ping
3. register_in_font_plan(entries) → 把解析到的条目加进 embed_fonts.FONT_PLAN（运行时）

不命中（商用 / 自托管 / 笔误）的家族会回退到 first_font/cjk_font 的 alias 匹配。
"""
import json
import re
import urllib.error
import urllib.request
from pathlib import Path

from font_paths import CACHE_DIR

GF_CSS2_URL = "https://fonts.googleapis.com/css2"
# Firefox 7 UA：GF 此场景返回 .woff（每 variant 一条 @font-face）；woff fontTools 原生支持，
# 不需要 brotli。再老的 IE UA 会给 google 私有 /l/font?kit= 子集格式，无法解析；
# 更新 Firefox / Chrome UA 会回 .woff2 多子集（需 brotli）。Firefox 7 是甜点。
OLD_UA = "Mozilla/5.0 (Windows NT 6.1; rv:7.0) Gecko/20100101 Firefox/7.0"
RESOLVED_INDEX = CACHE_DIR / "_resolved.json"

# Variable-font 直链下载配方。GF CSS2 API 对 CJK 家族会返回大量 unicode-range
# 分片 @font-face（每片几百 KB），_parse_gf_css 拿不到完整字模。改用 google/fonts
# GitHub 仓库的 variable .ttf 一次性下载（~10 MB / family）→ fontTools instancer
# 抽 Regular/Bold 静态 ttf 落盘。Latin 字体走 GF CSS2 API 路径仍然没问题。
VARIABLE_RECIPES = {
    "noto serif sc": {
        "url": "https://raw.githubusercontent.com/google/fonts/main/ofl/notoserifsc/NotoSerifSC%5Bwght%5D.ttf",
        "style": "serif", "cjk": True, "axis": "wght",
        "slot_weights": {"regular": 400, "bold": 700},
        "aliases": ["source-han-serif-sc", "source han serif sc"],
    },
    "noto sans sc": {
        "url": "https://raw.githubusercontent.com/google/fonts/main/ofl/notosanssc/NotoSansSC%5Bwght%5D.ttf",
        "style": "sans", "cjk": True, "axis": "wght",
        "slot_weights": {"regular": 400, "bold": 700},
        "aliases": ["source-han-sans-sc", "source han sans sc"],
    },
}

# CSS 通用关键字，不当作字体家族解析
GENERICS = {
    "serif", "sans-serif", "monospace", "cursive", "fantasy",
    "system-ui", "ui-sans-serif", "ui-serif", "ui-monospace",
    "math", "emoji", "inherit", "initial", "unset", "revert",
}

_FONT_FACE_RE = re.compile(r"@font-face\s*\{([^}]+)\}", re.MULTILINE)


# === 工具 ===

def _strip_quotes(name: str) -> str:
    return name.strip().strip('"').strip("'")


def parse_font_family_stack(family_str: str) -> list[str]:
    """'Tektur, cursive' → ['Tektur', 'cursive']。"""
    if not family_str:
        return []
    return [_strip_quotes(p) for p in family_str.split(",") if p.strip()]


def normalize_weight(weight) -> int:
    """CSS weight → int。'bold' → 700, 'normal' → 400, '600' → 600。"""
    if isinstance(weight, (int, float)):
        return int(weight)
    s = str(weight or "").strip().lower()
    if s == "bold":
        return 700
    if s in ("normal", "", "regular"):
        return 400
    if s == "lighter":
        return 300
    if s == "bolder":
        return 700
    try:
        return int(s)
    except ValueError:
        return 400


def normalize_italic(style) -> bool:
    return str(style or "").strip().lower() in ("italic", "oblique")


# === 解析索引（持久化记录"已尝试过、GF 没有"的家族，避免每次重 ping）===

def _load_index() -> dict:
    if RESOLVED_INDEX.exists():
        try:
            return json.loads(RESOLVED_INDEX.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            pass
    return {}


def _save_index(idx: dict):
    RESOLVED_INDEX.write_text(json.dumps(idx, ensure_ascii=False, indent=2), encoding="utf-8")


# === 收集 ===

def collect_requested_fonts(meas: dict) -> dict[str, set[tuple[int, bool]]]:
    """返回 {family: {(weight, italic), ...}}。"""
    needed: dict[str, set[tuple[int, bool]]] = {}
    slides = meas.get("slides") if "slides" in meas else [meas]
    for s in slides:
        for rec in s.get("records", []):
            if rec.get("kind") != "text":
                continue
            style = rec.get("style") or {}
            for run in rec.get("runs", []) or []:
                fam_str = run.get("fontFamily") or style.get("fontFamily") or ""
                weight = run.get("fontWeight") or style.get("fontWeight") or 400
                italic_v = run.get("fontStyle") or style.get("fontStyle") or "normal"
                w = normalize_weight(weight)
                it = normalize_italic(italic_v)
                for fam in parse_font_family_stack(fam_str):
                    if fam.lower() in GENERICS:
                        continue
                    needed.setdefault(fam, set()).add((w, it))
    return needed


# === Google Fonts CSS API ===

def _fetch_gf_css(family: str, variants: set[tuple[int, bool]]) -> str | None:
    """family + variants 调 css2 API；404 / 400 返回 None。"""
    family_q = family.replace(" ", "+")
    pairs = sorted(variants, key=lambda x: (x[1], x[0]))
    has_italic = any(it for _, it in pairs)
    if has_italic:
        spec = ";".join(f"{1 if it else 0},{w}" for w, it in pairs)
        url = f"{GF_CSS2_URL}?family={family_q}:ital,wght@{spec}"
    else:
        if len(pairs) == 1 and pairs[0][0] == 400:
            url = f"{GF_CSS2_URL}?family={family_q}"
        else:
            spec = ";".join(str(w) for w, _ in pairs)
            url = f"{GF_CSS2_URL}?family={family_q}:wght@{spec}"
    req = urllib.request.Request(url, headers={"User-Agent": OLD_UA})
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        if e.code in (400, 404):
            return None
        return None
    except Exception:
        return None


def _parse_gf_css(css: str) -> list[dict]:
    """提取每条 @font-face 的 (weight, italic, url)。"""
    out = []
    for body in _FONT_FACE_RE.findall(css):
        weight = 400
        italic = False
        url = None
        for line in body.splitlines():
            line = line.strip().rstrip(";")
            if line.startswith("font-weight:"):
                try:
                    weight = int(line.split(":", 1)[1].strip())
                except ValueError:
                    pass
            elif line.startswith("font-style:"):
                italic = "italic" in line.split(":", 1)[1]
            elif line.startswith("src:"):
                m = re.search(r"url\(([^)]+)\)", line)
                if m:
                    url = m.group(1).strip().strip("'\"")
        # Firefox 7 UA 模式下 GF 给 .woff 直链（无后缀也可能，先收，后面下载时再校验）
        if url and (".woff" in url or ".ttf" in url):
            out.append({"weight": weight, "italic": italic, "url": url})
    return out


def _download_and_convert_to_ttf(url: str, dst: Path) -> int:
    """下载字体；如果是 woff/woff2，用 fontTools 转 ttf 落盘。
    返回最终落盘字节数。"""
    req = urllib.request.Request(url, headers={"User-Agent": OLD_UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        data = r.read()

    # 检测 magic：TTF 0x00010000 / OTF 'OTTO' / WOFF 'wOFF' / WOFF2 'wOF2'
    magic = data[:4]
    if magic in (b"\x00\x01\x00\x00", b"OTTO", b"true"):
        # 原生 TTF/OTF，直接落盘
        dst.write_bytes(data)
        return len(data)
    if magic in (b"wOFF", b"wOF2"):
        # WOFF/WOFF2 → fontTools 解压后另存为 ttf
        import io
        from fontTools.ttLib import TTFont
        font = TTFont(io.BytesIO(data))
        font.flavor = None  # 落盘时清掉 flavor → 纯 ttf
        font.save(str(dst))
        return dst.stat().st_size
    raise RuntimeError(f"未知字体格式 magic={magic!r} url={url[:80]}")


def _normalize_to_slot(path: Path, family: str, slot: str, source_weight: int | None = None) -> None:
    """改写 TTF 的 name 表 / OS/2，让 PowerPoint 认为它就是 family + slot 标准变体。

    GF 服务给 Space Grotesk wght@500 返回的文件 nameID=1='Space Grotesk Medium'，
    OOXML 里我们写 typeface='Space Grotesk' → PowerPoint 名称匹配失败，拒绝加载嵌入字体
    → 全部回退系统字体。这里强制把 name 表写成 (family, family+style) 让 PowerPoint 接受。
    """
    from fontTools.ttLib import TTFont
    style_map = {"regular": "Regular", "italic": "Italic", "bold": "Bold", "boldItalic": "Bold Italic"}
    style = style_map[slot]
    full = f"{family} {style}"
    psname = f"{family.replace(' ', '')}-{style.replace(' ', '')}"
    is_bold = slot in ("bold", "boldItalic")
    is_italic = slot in ("italic", "boldItalic")
    target_weight = 700 if is_bold else 400

    f = TTFont(str(path))
    name_table = f["name"]
    # 清掉旧的 1/2/3/4/6/16/17（避免 Windows 上"Space Grotesk Medium"残留触发其他匹配）
    name_table.names = [r for r in name_table.names if r.nameID not in (1, 2, 3, 4, 6, 16, 17)]
    # 只写 Windows English 一组，PowerPoint 足够认
    name_table.setName(family, 1, 3, 1, 0x0409)
    name_table.setName(style, 2, 3, 1, 0x0409)
    name_table.setName(f"html-to-pptx;{family};slot={slot};src={source_weight or target_weight}", 3, 3, 1, 0x0409)
    name_table.setName(full, 4, 3, 1, 0x0409)
    name_table.setName(psname, 6, 3, 1, 0x0409)
    # OS/2 weight 改成 slot 标准
    f["OS/2"].usWeightClass = target_weight
    # fsSelection flags
    fs = f["OS/2"].fsSelection
    fs = (fs | 0x01) if is_italic else (fs & ~0x01)
    fs = (fs | 0x20) if is_bold else (fs & ~0x20)
    fs = (fs | 0x40) if (not is_italic and not is_bold) else (fs & ~0x40)
    f["OS/2"].fsSelection = fs
    # head.macStyle 同步
    ms = 0
    if is_bold: ms |= 0x01
    if is_italic: ms |= 0x02
    f["head"].macStyle = ms
    f.save(str(path))


def _cached_font_matches(path: Path, family: str, weight: int, source_weight: int | None = None) -> bool:
    """校验 cache 里的 TTF 是不是真符合 (family, weight)。

    PowerPoint/WPS 用 name 表的 nameID=1 + OS/2.usWeightClass 来匹配嵌入字体。
    文件名对但 name/weight 不对时（如 Medium(500) 文件被命名为 Regular(400)），
    PowerPoint 加载会匹配失败、回退系统字体——校验通过才算命中，不通过则
    返回 False 让上游重下覆盖。
    """
    try:
        from fontTools.ttLib import TTFont
        f = TTFont(str(path))
        n1 = f["name"].getDebugName(1) or ""
        actual_weight = f["OS/2"].usWeightClass
        marker = f["name"].getDebugName(3) or ""
    except Exception:
        return False  # 文件破损也当 miss，强制重下
    # nameID=1 必须等于 family（去空格不区分大小写）
    if n1.replace(" ", "").lower() != family.replace(" ", "").lower():
        return False
    # weight 在 ±50 内算匹配（应对 Medium-vs-Regular 偏差等）
    if abs(actual_weight - weight) > 50:
        return False
    if source_weight is not None and marker.startswith("html-to-pptx;"):
        if f"src={source_weight}" not in marker:
            return False
    return True


# === 命名 / 分类 ===

def _safe_filename(family: str, weight: int, italic: bool) -> str:
    """Tektur 700 → 'Tektur-w700.ttf'；Caveat 400 italic → 'Caveat-w400Italic.ttf'。

    The source weight is part of the cache filename because each requested CSS
    weight is embedded as its own OOXML typeface, e.g. Syne 700 and Syne 800.
    This prevents a later download for one weight from masquerading as another
    after name-table normalization.
    """
    family_clean = re.sub(r"[^A-Za-z0-9]+", "", family) or "Font"
    style_part = "Italic" if italic else ""
    return f"{family_clean}-w{int(weight)}{style_part}.ttf"


def _slot_for(weight: int, italic: bool) -> str:
    if italic and weight >= 600:
        return "boldItalic"
    if italic:
        return "italic"
    if weight >= 600:
        return "bold"
    return "regular"


def _exact_typeface_name(family: str, weight: int, italic: bool) -> str:
    suffix = f"{int(weight)}"
    if italic:
        suffix += " Italic"
    return f"{family} {suffix}"


def _classify_style(family: str) -> tuple[str, bool]:
    """启发式分类。返回 (style, cjk)。
    cjk 命中：noto sc/tc/jp/kr、han、ming、song、kai 等关键字。
    style 命中：mono/code → mono；serif → serif；其他归 sans（含 display/handwriting）。
    """
    lower = family.lower()
    cjk_keys = ("sc", "tc", "jp", "kr", "cjk", "han", "ming", "song", "kai", "hei")
    is_cjk = (
        "noto" in lower and any(k in lower.split() for k in cjk_keys)
        or any(k in lower for k in ("han ", "cjk", "ming", "song", "kai"))
    )
    if "mono" in lower or "code" in lower or "courier" in lower:
        return ("mono", is_cjk)
    if "serif" in lower and "sans" not in lower:
        return ("serif", is_cjk)
    return ("sans", is_cjk)


# === 主入口 ===

def resolve_fonts(needed: dict[str, set[tuple[int, bool]]],
                  bundled_aliases: set[str]) -> dict:
    """对每个未内置家族尝试 Google Fonts 解析。

    返回 {'resolved': [...条目...], 'unavailable': [家族名], 'cached': [家族名]}。
    """
    idx = _load_index()
    resolved_entries = []
    unavailable = []
    cached_used = []

    for family, variants in needed.items():
        # 已被 FONT_PLAN / alias 覆盖
        if family.lower() in bundled_aliases:
            continue

        # 1) Variable 直链通路（CJK 等）：在 GF CSS API 之前优先匹配
        if family.lower() in VARIABLE_RECIPES:
            recipe = VARIABLE_RECIPES[family.lower()]
            entry, was_cached = _resolve_via_variable_recipe(family, recipe, variants)
            if entry is not None:
                resolved_entries.append(entry)
                if was_cached:
                    cached_used.append(family)
                idx[family.lower()] = "ok"
                continue
            # 失败兜底：仍尝试走 GF CSS API（小概率，比如直链 404 / 网断）

        record = idx.get(family.lower())

        # 之前 ping 过 GF 已知没有 → 直接跳到 unavailable
        if record == "not-in-google-fonts":
            unavailable.append(family)
            continue

        # 之前已解析过且 cache 文件齐全且 name/weight 校验通过 → 直接组装 entry
        if record == "ok":
            exact_entries = []
            all_valid = True
            style, cjk = _classify_style(family)
            for w, it in sorted(variants):
                typeface = _exact_typeface_name(family, w, it)
                fname = _safe_filename(family, w, it)
                p = CACHE_DIR / fname
                if not p.exists():
                    all_valid = False
                    break
                if not _cached_font_matches(p, typeface, 400, source_weight=w):
                    print(f"  [font-resolve] cache 校验失败 {p.name} → fall through 重新解析")
                    all_valid = False
                    break
                exact_entries.append(_make_exact_entry(family, w, it, style, cjk, p.name))
            if exact_entries and all_valid:
                resolved_entries.extend(exact_entries)
                cached_used.append(family)
                continue

        # 联网 GF
        css = _fetch_gf_css(family, variants)
        if not css:
            idx[family.lower()] = "not-in-google-fonts"
            unavailable.append(family)
            continue
        faces = _parse_gf_css(css)
        if not faces:
            idx[family.lower()] = "not-in-google-fonts"
            unavailable.append(family)
            continue

        exact_entries = []
        style, cjk = _classify_style(family)
        for face in faces:
            w, it, url = face["weight"], face["italic"], face["url"]
            typeface = _exact_typeface_name(family, w, it)
            fname = _safe_filename(family, w, it)
            dst = CACHE_DIR / fname
            need_download = (not dst.exists()) or (
                not _cached_font_matches(dst, typeface, 400, source_weight=w)
            )
            if need_download:
                try:
                    n = _download_and_convert_to_ttf(url, dst)
                    _normalize_to_slot(dst, typeface, "regular", source_weight=w)
                    print(f"  [font-resolve] {family} {w}{' italic' if it else ''} → {fname} ({n:,} B)")
                except Exception as e:
                    print(f"  [font-resolve] {family} {w}{' italic' if it else ''} 下载失败: {e}")
                    continue
            exact_entries.append(_make_exact_entry(family, w, it, style, cjk, fname))

        if not exact_entries:
            idx[family.lower()] = "not-in-google-fonts"
            unavailable.append(family)
            continue

        resolved_entries.extend(exact_entries)
        idx[family.lower()] = "ok"

    _save_index(idx)
    return {
        "resolved": resolved_entries,
        "unavailable": unavailable,
        "cached": cached_used,
    }


def _make_entry(family: str, style: str, cjk: bool, slots: dict,
                aliases: list[str] | None = None) -> dict:
    pitch = {"serif": "18", "sans": "34", "mono": "50"}.get(style, "18")
    return {
        "typeface": family,
        "charset": "-122" if cjk else "0",
        "pitchFamily": pitch,
        "cjk": cjk,
        "style": style,
        "slots": slots,
        "aliases": list(aliases or []),
    }


def _make_exact_entry(family: str, weight: int, italic: bool, style: str,
                      cjk: bool, fname: str, aliases: list[str] | None = None) -> dict:
    typeface = _exact_typeface_name(family, weight, italic)
    pitch = {"serif": "18", "sans": "34", "mono": "50"}.get(style, "18")
    return {
        "typeface": typeface,
        "charset": "-122" if cjk else "0",
        "pitchFamily": pitch,
        "cjk": cjk,
        "style": style,
        "slots": {"regular": fname},
        "aliases": list(aliases or []),
        "cssFamily": family,
        "sourceWeight": int(weight),
        "sourceItalic": bool(italic),
    }


def _http_fetch(url: str, out: Path) -> int:
    out.parent.mkdir(parents=True, exist_ok=True)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=180) as r:
        data = r.read()
    out.write_bytes(data)
    return len(data)


def _resolve_via_variable_recipe(family: str, recipe: dict,
                                 variants: set[tuple[int, bool]]) -> tuple[dict | None, bool]:
    """Variable-font 直链通路：下载 variable .ttf → instancer 抽 slot 静态 ttf。

    返回 (entry_or_none, all_cached)。all_cached=True 表示零网络（cache 全命中）。
    CJK 一律 instance regular + bold 两个 slot（变化源已经下了，多抽一个权重几乎无成本），
    避免下次只用到 bold 时又要重下 10MB variable。
    """
    from fontTools.ttLib import TTFont
    from fontTools.varLib.instancer import instantiateVariableFont

    slot_weights = recipe["slot_weights"]
    needed_slots = {"regular", "bold"} if recipe.get("cjk") else set()
    for w, italic in variants:
        slot = _slot_for(w, italic)
        # CJK 没有 italic 静态，降级到对应权重的正体
        if slot == "italic":
            slot = "regular"
        elif slot == "boldItalic":
            slot = "bold"
        if slot in slot_weights:
            needed_slots.add(slot)
    if not needed_slots:
        needed_slots = {"regular"}

    # 先看 cache：所有需要的 slot 都有合法 ttf → 零网络返回
    slots_out: dict[str, str] = {}
    all_cached = True
    for slot in needed_slots:
        target_w = slot_weights[slot]
        fname = _safe_filename(family, target_w, False)
        p = CACHE_DIR / fname
        if p.exists() and _cached_font_matches(p, family, target_w):
            slots_out[slot] = fname
        else:
            all_cached = False
            break

    if not all_cached:
        slots_out = {}
        # 下载 variable 源（一次）→ 缓存到 _tmp，instance 完再清掉
        tmp_dir = CACHE_DIR / "_tmp"
        tmp_dir.mkdir(exist_ok=True)
        var_path = tmp_dir / (family.replace(" ", "") + "-var.ttf")
        try:
            print(f"  [font-resolve] downloading {family} variable source (~10 MB)...")
            n = _http_fetch(recipe["url"], var_path)
            print(f"  [font-resolve] {family} variable {n:,} B")
            var_font = TTFont(str(var_path))
            for slot in needed_slots:
                target_w = slot_weights[slot]
                fname = _safe_filename(family, target_w, False)
                dst = CACHE_DIR / fname
                instance = instantiateVariableFont(
                    var_font, {recipe["axis"]: target_w}, inplace=False, optimize=True,
                )
                instance.save(str(dst))
                _normalize_to_slot(dst, family, slot)
                print(f"  [font-resolve] instanced {fname} (wght={target_w})")
                slots_out[slot] = fname
        except Exception as e:
            print(f"  [font-resolve] {family} 直链下载/抽样失败: {e}")
            return None, False
        finally:
            try:
                var_path.unlink(missing_ok=True)
            except OSError:
                pass

    entry = _make_entry(family, recipe["style"], recipe["cjk"], slots_out,
                        aliases=recipe.get("aliases"))
    return entry, all_cached


def register_in_font_plan(entries: list):
    """把解析到的条目加进 embed_fonts.FONT_PLAN（运行时变 list 共享）。"""
    from embed_fonts import FONT_PLAN
    existing = {e["typeface"].lower() for e in FONT_PLAN}
    for entry in entries:
        if entry["typeface"].lower() in existing:
            continue
        FONT_PLAN.append(entry)


def report_summary(report: dict):
    """终端打印一行总结，方便用户知道哪些自动拉到、哪些不行。"""
    resolved = report["resolved"]
    unavailable = report["unavailable"]
    cached = report["cached"]
    if resolved:
        new_names = [e["typeface"] for e in resolved if e["typeface"] not in cached]
        if new_names:
            print(f"[fonts] auto-resolved {len(new_names)} 个家族 from Google Fonts: {', '.join(new_names)}")
        if cached:
            print(f"[fonts] cache 复用 {len(cached)} 个家族: {', '.join(cached)}")
    if unavailable:
        print(f"[fonts] {len(unavailable)} 个家族 GF 没有，回退到内置 alias: {', '.join(unavailable)}")
