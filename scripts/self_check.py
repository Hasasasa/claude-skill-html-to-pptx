"""self_check.py — Stage 5a 结构化自检（默认开启）

把生成的 pptx 渲染为 PNG（给 Stage 5b 视觉 audit 用），并扫 OOXML 结构告警
（FULL-PIC / LAYOUT / PREFLIGHT），不做像素 diff。

渲染器优先级（自动降级）：
1. PowerPoint COM（Windows + Office + pywin32）— 用户最终打开 pptx 的引擎
2. LibreOffice headless + pdf2image — 跨平台
3. 都不可用 → 跳过 PPT 截图，只产出结构化告警，不影响 pptx 产出

不会让 convert 失败 — 自检异常仅打印警告。
"""
import json
import re
import shutil
import subprocess
import sys
import tempfile
from contextlib import nullcontext
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


# 与 assemble.py 保持同一套换算（1920 px 视口 ↔ 12192000 EMU 16:9 幻灯片宽）。
# 旧版本用 9525 = 914400/96（96 dpi 标准）让 self_check 内部坐标变成 "1280 logical"
# 空间，再把 HTML measurement 也 scale 进同空间——能 round-trip 但读起来困惑。
# 现在直接对齐 assemble 的 6350 EMU/px：self_check 的所有几何就是 1920×1080 px 空间。
EMU_PER_PX = 12192000 / 1920  # = 6350
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def _try_powerpoint_com(pptx_path: Path, out_dir: Path, only_indices: set[int] | None = None):
    """返回 (count, error)。count = pres 总页数（含缓存命中页），不只是本次新渲染数。

    only_indices 给定时走增量：只对 (i ∈ only_indices) 或 (cache miss) 的页 Export。
    其它页保留 out_dir 里上轮的 slide_NN.png。
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return 0, "pywin32 未安装"
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            # 参数顺序: FileName, ReadOnly, Untitled, WithWindow
            pres = app.Presentations.Open(str(pptx_path.resolve()), True, False, True)
            try:
                total = 0
                for i, slide in enumerate(pres.Slides, start=1):
                    total += 1
                    out_png = out_dir / f"slide_{i:02d}.png"
                    must_render = (only_indices is None
                                   or i in only_indices
                                   or not out_png.exists())
                    if must_render:
                        slide.Export(str(out_png), "PNG", 1920, 1080)
                return total, None
            finally:
                pres.Close()
        finally:
            app.Quit()
    except Exception as e:
        return 0, f"PowerPoint COM 调用失败: {e}"


def _try_libreoffice(pptx_path: Path, out_dir: Path, only_indices: set[int] | None = None):
    """LibreOffice headless 把 pptx 转 PDF，再用 pdf2image 拆页。

    only_indices 给定时走增量：仍要做 pptx→PDF（无法绕过），但 PDF→PNG 只对
    (i ∈ only_indices) 或 (cache miss) 的页做（pdf2image 支持 first_page/last_page 选页）。
    其它页保留 out_dir 里上轮的 slide_NN.png。
    """
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return 0, "未找到 LibreOffice (soffice / libreoffice)"
    try:
        from pdf2image import convert_from_path, pdfinfo_from_path
    except ImportError:
        return 0, "pdf2image 未安装"
    try:
        with tempfile.TemporaryDirectory(prefix="h2p_lo_") as td:
            r = subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf",
                 "--outdir", td, str(pptx_path)],
                capture_output=True, timeout=90,
            )
            if r.returncode != 0:
                return 0, f"LibreOffice 转 PDF 失败: {r.stderr.decode(errors='replace')[:200]}"
            pdf = next(Path(td).glob("*.pdf"), None)
            if not pdf:
                return 0, "LibreOffice 未产出 PDF"
            if only_indices is None:
                pages = convert_from_path(str(pdf), size=(1920, 1080))
                for i, p in enumerate(pages, start=1):
                    p.save(out_dir / f"slide_{i:02d}.png")
                return len(pages), None
            # 增量：先用 pdfinfo 拿总页数，然后只渲列出的 + cache miss 的
            info = pdfinfo_from_path(str(pdf))
            total = int(info.get("Pages", 0))
            for i in range(1, total + 1):
                out_png = out_dir / f"slide_{i:02d}.png"
                must_render = i in only_indices or not out_png.exists()
                if not must_render:
                    continue
                rendered = convert_from_path(str(pdf), size=(1920, 1080),
                                             first_page=i, last_page=i)
                if rendered:
                    rendered[0].save(out_png)
            return total, None
    except Exception as e:
        return 0, f"LibreOffice 渲染异常: {e}"


def render_pptx_to_pngs(pptx_path: Path, out_dir: Path, only_indices: set[int] | None = None):
    """按优先级尝试渲染器。返回 (engine_label, count, errors_dict)。

    PowerPoint COM > LibreOffice。都不可用就跳过 PPT 截图——
    没有 PIL fallback，因为 PIL 估算渲染会误导 Stage 5b 的 VLM。

    only_indices 给定时走增量：未列出且 cache 命中的页跳过 export。
    返回 count = pres 总页数（含缓存页），engine 判定不变。
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    errors = {}

    count, err = _try_powerpoint_com(pptx_path, out_dir, only_indices=only_indices)
    if count > 0:
        return ("PowerPoint", count, errors)
    errors["powerpoint"] = err

    count, err = _try_libreoffice(pptx_path, out_dir, only_indices=only_indices)
    if count > 0:
        return ("LibreOffice", count, errors)
    errors["libreoffice"] = err

    return (None, 0, errors)


def _norm_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _emu_to_px(value: str | int) -> float:
    return int(value) / EMU_PER_PX


def _slide_num(name: str) -> int:
    return int(re.search(r"slide(\d+)\.xml$", name).group(1))


def _slide_xml_names(zip_file: ZipFile):
    names = [
        name for name in zip_file.namelist()
        if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
    ]
    return sorted(names, key=_slide_num)


def _presentation_size_px(zip_file: ZipFile):
    root = etree.fromstring(zip_file.read("ppt/presentation.xml"))
    size = root.find("p:sldSz", namespaces=NS)
    if size is None:
        return (1280, 720)
    return (_emu_to_px(size.get("cx")), _emu_to_px(size.get("cy")))


def _load_measurement_texts(measurements_path: Path, ppt_size):
    data = json.loads(measurements_path.read_text(encoding="utf-8"))
    by_slide = {}
    for slide_idx, slide in enumerate(data.get("slides", []), 1):
        meta = slide.get("slide", {})
        sx = ppt_size[0] / float(meta.get("width") or 1920)
        sy = ppt_size[1] / float(meta.get("height") or 1080)
        refs = []
        for rec in slide.get("records", []):
            if rec.get("kind") != "text":
                continue
            text = _norm_text("".join(run.get("text", "") for run in rec.get("runs", [])))
            if not text:
                continue
            r = rec.get("rect") or {}
            refs.append({
                "text": text,
                "x": float(r.get("x", 0)) * sx,
                "y": float(r.get("y", 0)) * sy,
                "w": float(r.get("w", 0)) * sx,
                "h": float(r.get("h", 0)) * sy,
            })
        by_slide[slide_idx] = refs
    return by_slide


def _vertical_overlap_ratio(a, b):
    overlap = min(a["y"] + a["h"], b["y"] + b["h"]) - max(a["y"], b["y"])
    if overlap <= 0:
        return 0
    return overlap / max(1, min(a["h"], b["h"]))


def _horizontal_overlap(a, b):
    return min(a["x"] + a["w"], b["x"] + b["w"]) - max(a["x"], b["x"])


def _attach_reference_rects(shapes, refs):
    remaining = list(refs)
    for shape in shapes:
        for idx, ref in enumerate(remaining):
            if shape["text"] == ref["text"]:
                shape["ref"] = ref
                del remaining[idx]
                break


FULL_SLIDE_PIC_CX = 12_000_000   # EMU；slide 16:9 默认 cx ≈ 12192000
FULL_SLIDE_PIC_CY = 6_800_000    #         cy ≈ 6858000


def full_slide_picture_warnings(pptx_path: Path):
    """扫每张 slide.xml 找 cx ≥ FULL_SLIDE_PIC_CX 且 cy ≥ FULL_SLIDE_PIC_CY 的 <p:pic>。

    命中即潜在 deco_snapshot 双层 bug 嫌疑（全屏 PNG + 文字另画一层）。
    详见 [[project-html-to-pptx-deco-snapshot-bug]]。
    """
    findings = []
    with ZipFile(pptx_path) as zf:
        for slide_name in _slide_xml_names(zf):
            slide_idx = _slide_num(slide_name)
            root = etree.fromstring(zf.read(slide_name))
            for pic in root.xpath(".//p:pic", namespaces=NS):
                ext = pic.find(".//a:ext", namespaces=NS)
                if ext is None:
                    continue
                try:
                    cx = int(ext.get("cx") or 0)
                    cy = int(ext.get("cy") or 0)
                except ValueError:
                    continue
                if cx >= FULL_SLIDE_PIC_CX and cy >= FULL_SLIDE_PIC_CY:
                    findings.append({
                        "idx": slide_idx,
                        "cx_emu": cx,
                        "cy_emu": cy,
                    })
                    break  # 一页一条即可
    return findings


def text_box_overlap_warnings(pptx_path: Path, measurements_path: Path | None,
                              min_overlap_px: float = 24,
                              min_vertical_ratio: float = 0.6,
                              min_reference_gap_px: float = 4):
    """Detect adjacent text boxes that overlap in PPTX but not in HTML measurements."""
    if measurements_path is None or not measurements_path.exists():
        return []

    warnings = []
    with ZipFile(pptx_path) as zip_file:
        ppt_size = _presentation_size_px(zip_file)
        refs_by_slide = _load_measurement_texts(measurements_path, ppt_size)

        for slide_name in _slide_xml_names(zip_file):
            slide_idx = _slide_num(slide_name)
            root = etree.fromstring(zip_file.read(slide_name))
            shapes = []
            for order, sp in enumerate(root.xpath(".//p:sp", namespaces=NS), 1):
                text = _norm_text("".join(sp.xpath(".//a:t/text()", namespaces=NS)))
                if not text:
                    continue
                xfrm = sp.find(".//a:xfrm", namespaces=NS)
                if xfrm is None:
                    continue
                off = xfrm.find("a:off", namespaces=NS)
                ext = xfrm.find("a:ext", namespaces=NS)
                if off is None or ext is None:
                    continue
                c_nv_pr = sp.find(".//p:cNvPr", namespaces=NS)
                shapes.append({
                    "order": order,
                    "id": c_nv_pr.get("id") if c_nv_pr is not None else str(order),
                    "text": text,
                    "x": _emu_to_px(off.get("x")),
                    "y": _emu_to_px(off.get("y")),
                    "w": _emu_to_px(ext.get("cx")),
                    "h": _emu_to_px(ext.get("cy")),
                })

            _attach_reference_rects(shapes, refs_by_slide.get(slide_idx, []))
            ordered = sorted(shapes, key=lambda item: (item["x"], item["y"]))
            for i, left in enumerate(ordered):
                for right in ordered[i + 1:]:
                    if right["x"] <= left["x"]:
                        continue
                    if _vertical_overlap_ratio(left, right) < min_vertical_ratio:
                        continue
                    if "ref" not in left or "ref" not in right:
                        continue
                    if _vertical_overlap_ratio(left["ref"], right["ref"]) < min_vertical_ratio:
                        continue

                    ref_left, ref_right = left["ref"], right["ref"]
                    if ref_right["x"] < ref_left["x"]:
                        ref_left, ref_right = ref_right, ref_left
                    if ref_right["x"] - (ref_left["x"] + ref_left["w"]) < min_reference_gap_px:
                        continue

                    overlap = _horizontal_overlap(left, right)
                    if overlap >= min_overlap_px:
                        warnings.append({
                            "idx": slide_idx,
                            "overlap_px": round(overlap, 1),
                            "left": left["text"][:48],
                            "right": right["text"][:48],
                        })
    return warnings


def self_check(pptx_path: Path, html_screenshots_dir: Path,
               measurements_path: Path | None = None,
               preflight_result: dict | None = None,
               ppt_screenshots_keep_dir: Path | None = None,
               only_indices: set[int] | None = None,
               verbose: bool = True) -> dict:
    """Stage 5a 结构化自检：渲染 pptx → PNG，扫 OOXML 结构告警，不做像素 diff。

    告警：FULL-PIC（全屏图嫌疑）/ LAYOUT（文本框重叠）/ PREFLIGHT（合并 preflight 高风险）。
    PPT 截图给 Stage 5b 视觉 audit 用（ppt_screenshots_keep_dir 不为 None 时持久保留）。
    """
    result = {"skipped": None, "engine": None, "pages": [],
              "warnings": [], "layout_warnings": [],
              "fullslide_pic_warnings": [], "preflight_warnings": [],
              "errors": {}}

    if not html_screenshots_dir.exists():
        result["skipped"] = "no-html-screenshots"
        if verbose:
            print("[self-check] 跳过：未找到 HTML 参考图目录")
        return result

    html_pngs = sorted(html_screenshots_dir.glob("slide_*.png"))
    if not html_pngs:
        result["skipped"] = "no-html-pngs"
        if verbose:
            print("[self-check] 跳过：HTML 参考图目录为空")
        return result

    try:
        result["layout_warnings"] = text_box_overlap_warnings(pptx_path, measurements_path)
    except Exception as e:
        result["errors"]["layout"] = f"text box overlap check failed: {e}"

    # 结构化检查：全屏 <p:pic> 嫌疑（deco_snapshot 双层 bug）
    try:
        result["fullslide_pic_warnings"] = full_slide_picture_warnings(pptx_path)
    except Exception as e:
        result["errors"]["fullslide_pic"] = f"fullslide pic scan failed: {e}"

    # Stage 1 风险合并：preflight 低置信度页强制纳入复核
    if preflight_result:
        for slide in preflight_result.get("slides", []):
            if slide.get("confidence") == "low":
                codes = [r["code"] for r in slide.get("risks", []) if r["severity"] == "high"]
                result["preflight_warnings"].append({
                    "idx": slide["index"],
                    "codes": codes,
                    "theme": slide.get("theme", ""),
                })

    # PPT 渲染：给 Stage 5b 视觉 audit 用。本函数不做像素 diff，视觉判断交 5b VLM
    if ppt_screenshots_keep_dir is not None:
        ppt_screenshots_keep_dir.mkdir(parents=True, exist_ok=True)
        _cm = nullcontext(str(ppt_screenshots_keep_dir.parent))
        ppt_dir_static = ppt_screenshots_keep_dir
    else:
        _cm = tempfile.TemporaryDirectory(prefix="h2p_check_")
        ppt_dir_static = None

    with _cm as td:
        ppt_dir = ppt_dir_static if ppt_dir_static is not None else Path(td) / "ppt_pngs"
        engine, count, errors = render_pptx_to_pngs(pptx_path, ppt_dir,
                                                    only_indices=only_indices)
        result["errors"].update(errors)
        result["only_indices"] = sorted(only_indices) if only_indices else None

        if engine is None:
            result["skipped"] = "no-renderer"
            if verbose:
                print()
                print("=" * 72)
                print("[self-check] ❗ 缺少 pptx 渲染器 — Stage 5b 视觉 audit 跑不了")
                print("=" * 72)
                print(f"  尝试结果: {errors}")
                print()
                print("  视觉 audit 是必需的（PPT 可能有 OOXML 转换层看不出的视觉 bug）。")
                print("  agent 必须在交付前 ask 用户选择以下之一：")
                print()
                print("  1) 装 LibreOffice（推荐，跨平台 2-3 分钟）：")
                print("     Windows : winget install LibreOffice.LibreOffice")
                print("     macOS   : brew install --cask libreoffice")
                print("     Linux   : sudo apt install libreoffice")
                print("     全平台还要：pip install pdf2image  （Windows 装 poppler 见 pdf2image README）")
                print("     装完重跑 convert.py")
                print()
                print("  2) 跳过 audit 直接交付（接受 PPT 可能有视觉 bug 的风险）：")
                print("     convert.py … --no-visual-audit")
                print()
                print("  3) 在已装 Office / LibreOffice 的另一台机器上重跑")
                print("=" * 72)
                if result["layout_warnings"]:
                    print(f"[self-check] !! {len(result['layout_warnings'])} 处文本框横向重叠（Stage 5a 结构化告警仍有效）")
            return result

        result["engine"] = engine

        # 每页一条 page_info（不含 diff_pct 字段），用来挂载结构化告警
        n = min(len(html_pngs), len(sorted(ppt_dir.glob("slide_*.png"))))
        for i in range(n):
            result["pages"].append({"idx": i + 1, "level": "OK"})

        # 合并三类结构化告警到 warnings 列表
        for w in result["layout_warnings"]:
            idx = w["idx"]
            page_info = next((p for p in result["pages"] if p["idx"] == idx), None)
            if page_info is None:
                page_info = {"idx": idx, "level": "LAYOUT"}
                result["pages"].append(page_info)
            page_info["level"] = "LAYOUT"
            page_info["layout_overlap_count"] = sum(
                1 for x in result["layout_warnings"] if x["idx"] == idx)
            if not any(p["idx"] == idx for p in result["warnings"]):
                result["warnings"].append(page_info)

        for w in result["fullslide_pic_warnings"]:
            idx = w["idx"]
            page_info = next((p for p in result["pages"] if p["idx"] == idx), None)
            if page_info is None:
                page_info = {"idx": idx, "level": "FULLSLIDE_PIC"}
                result["pages"].append(page_info)
            page_info["level"] = "FULLSLIDE_PIC"
            page_info["fullslide_pic"] = True
            if not any(p["idx"] == idx for p in result["warnings"]):
                result["warnings"].append(page_info)

        for w in result["preflight_warnings"]:
            idx = w["idx"]
            page_info = next((p for p in result["pages"] if p["idx"] == idx), None)
            if page_info is None:
                page_info = {"idx": idx, "level": "PREFLIGHT"}
                result["pages"].append(page_info)
            if page_info["level"] not in ("FULLSLIDE_PIC",):
                page_info["level"] = "PREFLIGHT"
            page_info["preflight_codes"] = w["codes"]
            if not any(p["idx"] == idx for p in result["warnings"]):
                result["warnings"].append(page_info)

    # 终端报告
    if verbose:
        mode_tag = f"  ·  增量重渲页 {sorted(only_indices)}" if only_indices else ""
        print(f"\n[self-check] 渲染器 {engine}  ·  {len(result['pages'])} 页{mode_tag}（无像素 diff 比较 — 视觉判断交 Stage 5b audit）")
        if result["warnings"]:
            n_layout = sum(1 for p in result["warnings"] if p["level"] == "LAYOUT")
            n_pic = sum(1 for p in result["warnings"] if p["level"] == "FULLSLIDE_PIC")
            n_pre = sum(1 for p in result["warnings"] if p["level"] == "PREFLIGHT")
            print(f"[self-check] !! {len(result['warnings'])} 页结构化告警 "
                  f"（布局重叠 LAYOUT={n_layout}，全屏图 FULLSLIDE_PIC={n_pic}，预扫 PREFLIGHT={n_pre}）")
            marker_map = {
                "LAYOUT": "[LAYOUT]    ",
                "FULLSLIDE_PIC": "[FULL-PIC]  ",
                "PREFLIGHT": "[PREFLIGHT] ",
            }
            for p in sorted(result["warnings"], key=lambda x: x["idx"]):
                marker = marker_map.get(p.get("level"), "[?]          ")
                extras = []
                if p.get("layout_overlap_count"):
                    extras.append(f"文本框重叠 {p['layout_overlap_count']} 处")
                if p.get("fullslide_pic"):
                    extras.append("含全屏 PNG（疑 deco_snapshot 双层）")
                if p.get("preflight_codes"):
                    extras.append(f"预扫风险 {','.join(p['preflight_codes'])}")
                extra = ("  " + " · ".join(extras)) if extras else ""
                print(f"           {marker}第 {p['idx']:02d} 页{extra}")
            print("           → 这些只是结构化提示。**最终视觉判断必须走 Stage 5b 视觉 audit**")
        else:
            print("[self-check] OK 无结构化告警 — 视觉判断仍由 Stage 5b audit 完成")

    return result


if __name__ == "__main__":
    # CLI 独立使用：python self_check.py <pptx> <html_screenshots_dir>
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    pptx = Path(sys.argv[1])
    html_dir = Path(sys.argv[2])
    self_check(pptx, html_dir, verbose=True)
