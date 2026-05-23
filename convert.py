"""convert.py — html-to-pptx 单入口 CLI。

Usage: python convert.py <input.html> [--out output.pptx] [--no-embed-fonts] ...

流水线：preflight → measure → assemble → embed_fonts → self_check → visual_audit
"""
import argparse
import shutil
import sys
import tempfile
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).parent
SCRIPTS = ROOT / "scripts"

sys.path.insert(0, str(SCRIPTS))

# 字体缓存目录（user-home 跨项目共享，不随 skill 删除而丢）
from font_paths import CACHE_DIR as FONT_CACHE_DIR
from text_utils import is_cjk_text


def _has_cjk_chars(meas: dict) -> bool:
    """measurement 里任何文字 run 包含 CJK 字符？CJK 范围定义见 text_utils.CJK_RE。"""
    for s in meas.get("slides") or []:
        for rec in s.get("records", []) or []:
            for run in rec.get("runs", []) or []:
                if is_cjk_text(run.get("text")):
                    return True
            if is_cjk_text(rec.get("text")):
                return True
    return False


def convert(html_path: Path, out_path: Path, keep_screenshots: bool, embed_fonts: bool,
            do_verify: bool = True,
            do_preflight: bool = True,
            do_visual_audit: bool = True,
            install_user_fonts: bool = False):
    html_path = html_path.resolve()
    out_path = out_path.resolve()
    if not html_path.exists():
        raise FileNotFoundError(html_path)

    # 同进程 import：避开 3 次 Python 启动开销
    from preflight import preflight
    from measure import measure
    from assemble import assemble, refresh_font_plan_caches
    from embed_fonts import embed
    from self_check import self_check

    # 自检需要 HTML 参考图。如用户未指定 --keep-screenshots，
    # 也内部开启，自检完随临时目录清掉，不污染输出目录
    measure_needs_screenshots = keep_screenshots or do_verify

    with tempfile.TemporaryDirectory(prefix="h2p_") as tmp:
        tmp_dir = Path(tmp)
        intermediate_pptx = tmp_dir / "no_fonts.pptx"

        # anchor_json 控制 measurements.json + HTML 参考图 + svg 资源的落盘位置
        if keep_screenshots:
            anchor_json = out_path.parent / f"{out_path.stem}_measurements.json"
            anchor_json.parent.mkdir(parents=True, exist_ok=True)
        else:
            anchor_json = tmp_dir / "measurements.json"

        # 0) preflight（Stage 1 输入识别 / 风险预扫）
        preflight_result = None
        if do_preflight:
            t0 = time.perf_counter()
            if keep_screenshots:
                preflight_out = out_path.parent / f"{out_path.stem}_preflight.json"
            else:
                preflight_out = tmp_dir / "preflight.json"
            try:
                preflight_result = preflight(html_path, preflight_out, verbose=True)
            except Exception as e:
                print(f"[preflight] 异常（忽略，继续转换）: {e}")
            print(f"[preflight] 耗时 {time.perf_counter()-t0:.2f}s")

        # 1) measure（结果通过 dict 在内存里传递；anchor 仅用于 svg / 截图资源定位）
        t0 = time.perf_counter()
        meas = measure(html_path, anchor_json,
                       no_screenshots=not measure_needs_screenshots, verbose=True)
        print(f"[measure]  {time.perf_counter()-t0:.2f}s")

        # 1.5) auto-resolve fonts（按需从 GF 拉，CJK 走 variable 直链）
        # FONT_PLAN 启动为空，所有字体都在这里按需解析。HTML 含 CJK 字符就强制种子
        # Noto Sans/Serif SC（即使 CSS 没显式声明 CJK family，cjk_font 配对也需要）。
        t0 = time.perf_counter()
        font_report = None
        try:
            from font_resolver import (collect_requested_fonts, resolve_fonts,
                                       register_in_font_plan, report_summary)
            from embed_fonts import bundled_family_names_lower
            needed = collect_requested_fonts(meas)
            if _has_cjk_chars(meas):
                # 兼顾 latin serif → 配 Noto Serif SC，latin sans/mono → 配 Noto Sans SC
                for fam in ("Noto Sans SC", "Noto Serif SC"):
                    needed.setdefault(fam, set()).update({(400, False), (700, False)})
            font_report = resolve_fonts(needed, bundled_family_names_lower())
            register_in_font_plan(font_report["resolved"])
            # FONT_PLAN 已被 resolver 填充，让 assemble 的 FONT_FALLBACKS / CJK 缓存看到新条目
            refresh_font_plan_caches()
            report_summary(font_report)
        except Exception as e:
            print(f"[fonts] auto-resolve 异常（忽略，未解析的字体会回退到 viewer 系统字体）: {e}")
            traceback.print_exc()
        print(f"[fonts]    {time.perf_counter()-t0:.2f}s")

        # 1.6) 装到用户字体目录（可选，让 WPS 能正确渲染）
        # 默认不装；--install-user-fonts 显式打开。理由：WPS 不读裸 TTF 嵌入字体
        # （只认 ECMA-376 obfuscated EOT），装到用户字体目录后 WPS / Word 一律
        # 把它当系统字体用。是改用户系统的行为，必须用户明确同意。
        if install_user_fonts and font_report is not None:
            t0 = time.perf_counter()
            try:
                from font_user_install import install_fonts, collect_ttfs_for_install
                ttfs = collect_ttfs_for_install(font_report, FONT_CACHE_DIR)
                if ttfs:
                    install_fonts(ttfs, verbose=True)
                else:
                    print("[font-install] 没有需要安装的字体")
            except Exception as e:
                print(f"[font-install] 异常（忽略，pptx 已嵌入字体，PowerPoint 可正常打开）: {e}")
            print(f"[font-install] {time.perf_counter()-t0:.2f}s")

        # 2) assemble（直接拿 dict）
        t0 = time.perf_counter()
        assemble(meas, intermediate_pptx)
        print(f"[assemble] {time.perf_counter()-t0:.2f}s")

        # 3) embed fonts（可选）
        t0 = time.perf_counter()
        if embed_fonts:
            embed(intermediate_pptx, meas, out_path)
            print(f"[embed]    {time.perf_counter()-t0:.2f}s")
        else:
            shutil.copy(intermediate_pptx, out_path)
            print(f"[embed]    跳过")

        # 4) 自检（默认开启）— 自检异常不影响 pptx 产出
        self_check_result = None
        html_screenshots_dir = anchor_json.with_suffix("").parent / (anchor_json.stem + "_screenshots")
        if do_verify:
            t0 = time.perf_counter()
            # 视觉 audit 需要 PPT 渲染结果落盘
            ppt_keep_dir = None
            if do_visual_audit:
                ppt_keep_dir = out_path.parent / f"{out_path.stem}_audit" / "_ppt_renders"
            try:
                self_check_result = self_check(
                    out_path, html_screenshots_dir,
                    measurements_path=anchor_json,
                    preflight_result=preflight_result,
                    ppt_screenshots_keep_dir=ppt_keep_dir,
                    verbose=True)
            except Exception as e:
                print(f"[self-check] 异常（忽略，不影响产出）: {e}")
            print(f"[self-check] 耗时 {time.perf_counter()-t0:.2f}s")

        # 5) 视觉 audit 物料（可选）— 给上游 agent 用 VLM 视觉判断
        if do_visual_audit and self_check_result and self_check_result.get("engine"):
            t0 = time.perf_counter()
            from visual_audit import build_audit_package
            audit_dir = out_path.parent / f"{out_path.stem}_audit"
            ppt_pngs_dir = audit_dir / "_ppt_renders"
            try:
                pkg = build_audit_package(
                    pptx_path=out_path,
                    html_screenshots_dir=html_screenshots_dir,
                    ppt_screenshots_dir=ppt_pngs_dir,
                    self_check_result=self_check_result,
                    preflight_result=preflight_result,
                    out_dir=audit_dir,
                )
                print(f"[audit]    {pkg['pages']} 页对比物料 → {pkg['out_dir']}")
                print(f"[audit]    给上游 agent 看: 读 {pkg['prompt']} 然后逐页看 slide_NN_compare.png")
            except Exception as e:
                print(f"[audit] 异常（忽略）: {e}")
            print(f"[audit]    耗时 {time.perf_counter()-t0:.2f}s")

    print(f"\n[done] {out_path} ({out_path.stat().st_size:,} B)")


def cleanup_artifacts(pptx_path: Path) -> list[Path]:
    """删 .pptx 旁边的 audit / measurement / preflight 中间物料，只留 .pptx 本身。

    给 agent 在「audit 工作流跑完、最终把 PPT 交给用户」那一刻调用。会清掉：
    - <stem>_audit/                       Stage 5b 视觉 audit 物料 + agent 写的 findings.md
    - <stem>_measurements.json            --keep-screenshots 时落盘的 measurement
    - <stem>_measurements_screenshots/    HTML 参考截图
    - <stem>_measurements_svg_assets/     抽出来的 SVG / canvas / deco PNG
    - <stem>_preflight.json               Stage 1 风险预扫报告
    """
    pptx_path = pptx_path.resolve()
    stem = pptx_path.stem
    parent = pptx_path.parent
    removed = []
    targets = [
        parent / f"{stem}_audit",
        parent / f"{stem}_measurements.json",
        parent / f"{stem}_measurements_screenshots",
        parent / f"{stem}_measurements_svg_assets",
        parent / f"{stem}_preflight.json",
    ]
    for t in targets:
        if t.is_dir():
            shutil.rmtree(t, ignore_errors=True)
            removed.append(t)
        elif t.is_file():
            try:
                t.unlink()
                removed.append(t)
            except OSError:
                pass
    return removed


def main():
    ap = argparse.ArgumentParser(prog="html-to-pptx", description=__doc__)
    ap.add_argument("input", help="输入 HTML 路径；加 --cleanup 时则为 .pptx 路径")
    ap.add_argument("--out", help="输出 .pptx 路径（默认同名 .pptx）", default=None)
    ap.add_argument("--cleanup", action="store_true",
                    help="不做转换。删掉 input(.pptx) 旁边的 audit/measurement/preflight "
                         "中间物料，只留 .pptx 本身。**agent 在最终交付前调用**——audit "
                         "工作流跑完、所有 finding 修完之后，把这些工作物清掉再交付给用户")
    ap.add_argument("--keep-screenshots", action="store_true",
                    help="保留每页 PNG 参考截图（输出到 <out>_measurements_screenshots/）")
    ap.add_argument("--no-embed-fonts", action="store_true",
                    help="跳过字体嵌入；pptx 在没装这些字体的机器上会替换字体")
    ap.add_argument("--no-verify", action="store_true",
                    help="跳过转换后的结构化自检（Stage 5a）。默认开启自检：扫 OOXML 结构找"
                         "全屏图嫌疑 / 文本框横向重叠 / 合并 preflight 风险。"
                         "**不再做像素 diff 相似度比较**——像素 diff 对局部 bug 不敏感且给假信心，"
                         "视觉判断完全交 Stage 5b audit")
    ap.add_argument("--no-preflight", action="store_true",
                    help="跳过转换前的风险预扫（Stage 1）；预扫会标出已知踩坑模式并把"
                         "高风险页强制纳入自检告警")
    ap.add_argument("--no-visual-audit", action="store_true",
                    help="跳过 Stage 5b 视觉审计物料产出。默认开启视觉 audit："
                         "产出每页 HTML|PPT 双栏对比图 + audit_index.json + audit_prompt.md，"
                         "**调用 skill 的 agent 必须**逐页对比识别问题、迭代修复，直到交付。"
                         "只在批量 / CI / 已知不需要 audit 时关闭")
    ap.add_argument("--install-user-fonts", action="store_true",
                    help="把解析到的非 CJK 字体安装到用户字体目录。"
                         "Windows → %%LOCALAPPDATA%%\\Microsoft\\Windows\\Fonts\\ + HKCU 注册；"
                         "macOS → ~/Library/Fonts/；"
                         "Linux → ~/.local/share/fonts/ + fc-cache。"
                         "WPS Office 不读 pptx 里裸 TTF 嵌入字体，装到系统后 WPS 才能正确渲染。"
                         "**SKILL.md 要求 agent 在调用前必须先 ask 用户**——这是改用户系统行为")
    args = ap.parse_args()

    if args.cleanup:
        pptx_path = Path(args.input).resolve()
        if not pptx_path.exists():
            print(f"[cleanup] 找不到 {pptx_path}")
            sys.exit(1)
        removed = cleanup_artifacts(pptx_path)
        if removed:
            print(f"[cleanup] 已删除 {len(removed)} 项中间物料:")
            for r in removed:
                print(f"  - {r}")
        else:
            print("[cleanup] 无中间物料可清理")
        return

    in_path = Path(args.input)
    out_path = Path(args.out) if args.out else in_path.with_suffix(".pptx")
    convert(in_path, out_path,
            keep_screenshots=args.keep_screenshots,
            embed_fonts=not args.no_embed_fonts,
            do_verify=not args.no_verify,
            do_preflight=not args.no_preflight,
            do_visual_audit=not args.no_visual_audit,
            install_user_fonts=args.install_user_fonts)


if __name__ == "__main__":
    main()
