"""embed_fonts.py — 把 fonts/ 目录里的 TTF 嵌入到给定的 pptx 中。

支持按 measurement JSON 中所有页用到的字符做跨页 subset，避免 14MB 字库整包嵌入。

Usage:
    python embed_fonts.py <in.pptx> <measurement.json> <out.pptx>

OOXML 字体嵌入结构：
- ppt/fonts/fontN.fntdata           （原始 TTF 数据，扩展名 .fntdata）
- ppt/_rels/presentation.xml.rels   （每个独立 ttf 文件一条 Relationship）
- ppt/presentation.xml              （增加 <p:embeddedFontLst>，每 typeface 一个 entry）
- [Content_Types].xml               （增加 Default extension）
"""
import io
import json
import shutil
import sys
import zipfile
from pathlib import Path

from fontTools.subset import Subsetter
from fontTools.ttLib import TTFont
from lxml import etree

from font_paths import CACHE_DIR

# 字体配置权威源。assemble.py / preflight.py 都从这里派生映射逻辑。
# 字段：typeface（OOXML 字体名）/ charset / pitchFamily（OOXML <p:font> 必填）
#       / cjk / style（serif/sans/mono，用于 latin→CJK 配对）
#       / slots（{regular,bold,italic,boldItalic} → ttf 文件名）
#       / aliases（CSS 里映射到本 typeface 的别名，如 source-han 系列）
#
# 完全按需：启动时为空，font_resolver 在 convert 时按 HTML 实际用到的字体填充
# （Latin 走 Google Fonts CSS API，CJK 走 font_resolver.VARIABLE_RECIPES 直链）。
FONT_PLAN: list[dict] = []


# === 派生 helper，给 assemble.py / preflight.py 用 ===

def bundled_family_names_lower() -> set[str]:
    """所有可被 CSS font-family 命中的内置名（typeface + aliases），全小写。
    给 preflight.py 判定"是否非内置字体"用。"""
    s = set()
    for p in FONT_PLAN:
        s.add(p["typeface"].lower())
        for a in p.get("aliases", []):
            s.add(a.lower())
    return s


def cjk_typefaces() -> set[str]:
    """所有 cjk=True 的 typeface 名（保留大小写）。"""
    return {p["typeface"] for p in FONT_PLAN if p.get("cjk")}


def family_alias_map() -> dict[str, str]:
    """{CSS name (原大小写) → OOXML typeface}。
    包括 typeface 自己映射自己 + 所有 alias 映射到对应 typeface。
    给 assemble.first_font() 用。"""
    m = {}
    for p in FONT_PLAN:
        m[p["typeface"]] = p["typeface"]
        m[p["typeface"].lower()] = p["typeface"]
        for a in p.get("aliases", []):
            m[a] = p["typeface"]
            m[a.lower()] = p["typeface"]
    return m


def weighted_family_map() -> dict[tuple[str, int, bool], str]:
    """Map (CSS family/alias lower, source weight, italic) to an exact OOXML typeface.

    Used for families where the resolver embeds separate typefaces per source
    weight, e.g. CSS `Syne` 700 -> OOXML `Syne 700`, `Syne` 800 -> `Syne 800`.
    """
    m: dict[tuple[str, int, bool], str] = {}
    for p in FONT_PLAN:
        css_family = p.get("cssFamily")
        source_weight = p.get("sourceWeight")
        source_italic = bool(p.get("sourceItalic", False))
        if not css_family or source_weight is None:
            continue
        names = [css_family, *p.get("aliases", [])]
        for name in names:
            m[(name.lower(), int(source_weight), source_italic)] = p["typeface"]
    return m


def cjk_for_style(latin_style: str | None) -> str:
    """根据 latin 字体的风格（serif/sans/mono）选配对的 CJK 字体。
    没有 mono CJK 时 mono 也回退到 sans。"""
    target = latin_style or "sans"
    if target == "mono":
        target = "sans"  # 没有 mono CJK 字体
    for p in FONT_PLAN:
        if p.get("cjk") and p.get("style") == target:
            return p["typeface"]
    # 兜底：第一个 CJK 字体
    for p in FONT_PLAN:
        if p.get("cjk"):
            return p["typeface"]
    return "Noto Sans SC"


def style_of_typeface(typeface: str) -> str | None:
    """根据 typeface 名拿 style 字段。给 cjk_font() 用：知道 latin 字体风格才能匹配 CJK 风格。"""
    for p in FONT_PLAN:
        if p["typeface"] == typeface:
            return p.get("style")
    return None

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"


def chars_from_measurement(meas: dict) -> set[str]:
    """汇总 measurement（含 'slides' 列表）所有出现过的字符。"""
    chars: set[str] = set()
    slides = meas.get("slides") if "slides" in meas else [meas]
    for s in slides:
        for rec in s.get("records", []):
            for run in rec.get("runs", []) or []:
                t = run.get("text", "") or ""
                chars.update(t)
            txt = rec.get("text", "") or ""
            chars.update(txt)
            # text-transform: uppercase 在装配时会转大小写，两套都保留
            chars.update(txt.upper())
            chars.update(txt.lower())
    # 兜底：基本符号 + 全字母数字
    chars.update(" ·—–-,，。.:：;；()（）/0123456789"
                 "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
    return chars


def subset_ttf(path_in: Path, chars: set[str]) -> bytes:
    font = TTFont(str(path_in))
    sub = Subsetter()
    text = "".join(sorted(chars))
    sub.populate(text=text)
    sub.subset(font)
    buf = io.BytesIO()
    font.save(buf)
    return buf.getvalue()


def embed(in_pptx: Path, measurement, out_pptx: Path):
    """measurement 可以是 dict（in-process）或 Path / str（CLI）。"""
    if isinstance(measurement, (str, Path)):
        meas = json.loads(Path(measurement).read_text(encoding="utf-8"))
    else:
        meas = measurement
    chars = chars_from_measurement(meas)
    print(f"汇总 {len(chars)} 个唯一字符用于 subset")

    # 1) 收集所有独立的 ttf 文件，每个文件 subset 一次
    file_blobs: dict[str, bytes] = {}
    for plan in FONT_PLAN:
        for slot, fname in plan["slots"].items():
            if fname in file_blobs:
                continue
            ttf_path = CACHE_DIR / fname
            if not ttf_path.exists():
                raise FileNotFoundError(f"字体未缓存：{ttf_path}（font_resolver 应已下载，可能离线 / GF 失败）")
            blob = subset_ttf(ttf_path, chars)
            file_blobs[fname] = blob
            print(f"  subset {fname:<32} {ttf_path.stat().st_size:>11,} → {len(blob):>9,} B")

    # 2) 复制 pptx，进 zip 改 XML
    shutil.copyfile(in_pptx, out_pptx)
    with zipfile.ZipFile(out_pptx, "r") as zin:
        entries = {n: zin.read(n) for n in zin.namelist()}

    # 3) 把字体写入 ppt/fonts/，并给每个文件一个 rId
    file_to_rid: dict[str, str] = {}
    for i, (fname, blob) in enumerate(file_blobs.items(), start=1):
        part = f"ppt/fonts/font{i}.fntdata"
        entries[part] = blob
        file_to_rid[fname] = f"rIdFont{i}"

    # 4) ppt/_rels/presentation.xml.rels：每个独立文件一条 Relationship
    rels_path = "ppt/_rels/presentation.xml.rels"
    rels = etree.fromstring(entries[rels_path])
    # 清掉旧的 font Relationship（防重复运行）
    for old in rels.findall(f"{{{NS_RELS}}}Relationship"):
        if old.get("Type", "").endswith("/font"):
            rels.remove(old)
    for i, (fname, rid) in enumerate(file_to_rid.items(), start=1):
        rel = etree.SubElement(rels, f"{{{NS_RELS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font")
        rel.set("Target", f"fonts/font{i}.fntdata")
    entries[rels_path] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 5) ppt/presentation.xml：插入 <p:embeddedFontLst>
    pres_path = "ppt/presentation.xml"
    pres = etree.fromstring(entries[pres_path])
    for old in pres.findall(f"{{{NS_P}}}embeddedFontLst"):
        pres.remove(old)
    emb_list = etree.SubElement(pres, f"{{{NS_P}}}embeddedFontLst")
    for plan in FONT_PLAN:
        emb = etree.SubElement(emb_list, f"{{{NS_P}}}embeddedFont")
        font_el = etree.SubElement(emb, f"{{{NS_P}}}font")
        font_el.set("typeface", plan["typeface"])
        font_el.set("charset", plan["charset"])
        font_el.set("pitchFamily", plan["pitchFamily"])
        for slot in ("regular", "bold", "italic", "boldItalic"):
            if slot in plan["slots"]:
                fname = plan["slots"][slot]
                rid = file_to_rid[fname]
                slot_el = etree.SubElement(emb, f"{{{NS_P}}}{slot}")
                slot_el.set(f"{{{NS_R}}}id", rid)
    _reorder_pres_children(pres)
    entries[pres_path] = etree.tostring(pres, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 6) [Content_Types].xml：添加 fntdata Default
    ct_path = "[Content_Types].xml"
    ct = etree.fromstring(entries[ct_path])
    has_default = any(d.get("Extension") == "fntdata"
                      for d in ct.findall(f"{{{NS_CT}}}Default"))
    if not has_default:
        d = etree.Element(f"{{{NS_CT}}}Default")
        d.set("Extension", "fntdata")
        d.set("ContentType", "application/x-fontdata")
        ct.insert(0, d)
    entries[ct_path] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 7) 重写 zip
    with zipfile.ZipFile(out_pptx, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    print(f"saved {out_pptx} ({out_pptx.stat().st_size:,} B)")


def _reorder_pres_children(pres):
    order = [
        "sldMasterIdLst", "notesMasterIdLst", "handoutMasterIdLst",
        "sldIdLst", "sldSz", "notesSz", "smartTags",
        "embeddedFontLst", "custShowLst", "photoAlbum", "custDataLst",
        "kinsoku", "defaultTextStyle", "modifyVerifier", "extLst",
    ]
    rank = {name: i for i, name in enumerate(order)}
    children = list(pres)
    def keyfn(el):
        tag = etree.QName(el.tag).localname
        return rank.get(tag, len(order))
    children.sort(key=keyfn)
    for el in children:
        pres.remove(el)
    for el in children:
        pres.append(el)


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)
    embed(Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve(), Path(sys.argv[3]).resolve())
