"""font_user_install.py — 把解析到的字体装到用户字体目录。

为什么需要：WPS Office（含 macOS 版）不读 pptx 里裸 TTF 嵌入字体（只认 ECMA-376
obfuscated EOT），PowerPoint 接受裸 TTF 但 WPS 退回系统 fallback。
解决：把字体装到用户级字体目录（无需管理员），WPS / Word / 浏览器一律识别。

跨平台路径：
- Windows: `%LOCALAPPDATA%\\Microsoft\\Windows\\Fonts\\` + 写 HKCU 注册表
- macOS:   `~/Library/Fonts/`（系统自动扫，无需注册）
- Linux:   `~/.local/share/fonts/`（系统自动扫；改完跑 `fc-cache -f` 立即生效）

调用前必须先经用户同意（skill 不应擅自改用户系统）。convert.py 用 --install-user-fonts
flag 控制，SKILL.md 让上游 agent 在执行前 ask 用户。
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path


def user_font_dir() -> Path | None:
    """返回用户字体目录。不支持的平台返回 None。"""
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA")
        if not base:
            return None
        return Path(base) / "Microsoft" / "Windows" / "Fonts"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Fonts"
    if sys.platform.startswith("linux"):
        # XDG 优先；否则 ~/.local/share/fonts
        base = os.environ.get("XDG_DATA_HOME") or str(Path.home() / ".local" / "share")
        return Path(base) / "fonts"
    return None


def _full_name_from_ttf(ttf_path: Path) -> str:
    """从 TTF name 表读 nameID=4（Full font name），失败回退文件名。"""
    try:
        from fontTools.ttLib import TTFont
        f = TTFont(str(ttf_path))
        n4 = f["name"].getDebugName(4)
        if n4:
            return n4
    except Exception:
        pass
    return ttf_path.stem


def _register_windows(installed_paths: list[Path], verbose: bool):
    """Windows 把装好的字体写进 HKCU 注册表，不写 WPS / Office 找不到。"""
    try:
        import winreg
    except ImportError:
        if verbose:
            print("  [font-install] winreg 不可用，跳过注册（文件已落盘）")
        return
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
        0, winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE,
    )
    try:
        for dst in installed_paths:
            full_name = _full_name_from_ttf(dst)
            reg_name = f"{full_name} (TrueType)"
            try:
                winreg.SetValueEx(key, reg_name, 0, winreg.REG_SZ, str(dst))
            except Exception as e:
                if verbose:
                    print(f"  [font-install] 注册失败 {dst.name}: {e}")
    finally:
        try:
            key.Close()
        except Exception:
            pass


def _refresh_linux_fontconfig(verbose: bool):
    """Linux 调 fc-cache 让新装字体立即生效（mac/win 不需要，系统自动扫）。"""
    try:
        subprocess.run(["fc-cache", "-f"], check=False,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=30)
    except FileNotFoundError:
        if verbose:
            print("  [font-install] fc-cache 未安装；新字体可能需要重启会话才能识别")
    except Exception as e:
        if verbose:
            print(f"  [font-install] fc-cache 失败（忽略）: {e}")


def install_fonts(ttf_paths: list[Path], verbose: bool = True) -> dict:
    """把 TTF 拷到当前平台的用户字体目录，必要时写注册表 / 刷 fontconfig 缓存。

    Returns:
        {'installed': [path, ...], 'skipped_existing': [name], 'failed': [(name, err)]}
    """
    result = {"installed": [], "skipped_existing": [], "failed": []}
    dst_dir = user_font_dir()
    if dst_dir is None:
        if verbose:
            print(f"[font-install] 不支持的平台 sys.platform={sys.platform}，跳过")
        return result

    dst_dir.mkdir(parents=True, exist_ok=True)
    for src in ttf_paths:
        if not src.exists():
            result["failed"].append((src.name, "源文件不存在"))
            continue
        dst = dst_dir / src.name
        already_installed = dst.exists() and dst.stat().st_size == src.stat().st_size
        try:
            if not already_installed:
                shutil.copy2(src, dst)
                result["installed"].append(dst)
                if verbose:
                    print(f"  [font-install] 装: {src.name} → {dst}")
            else:
                result["skipped_existing"].append(src.name)
                if verbose:
                    print(f"  [font-install] 已存在: {src.name}")
        except Exception as e:
            result["failed"].append((src.name, str(e)))
            if verbose:
                print(f"  [font-install] 失败 {src.name}: {e}")

    # 平台后处理
    all_installed_or_existing = result["installed"] + [dst_dir / n for n in result["skipped_existing"]]
    if sys.platform == "win32" and all_installed_or_existing:
        _register_windows(all_installed_or_existing, verbose)
    elif sys.platform.startswith("linux") and result["installed"]:
        _refresh_linux_fontconfig(verbose)
    # macOS 不用做任何后处理，系统启动 fonts daemon 自动扫 ~/Library/Fonts/

    if verbose and (result["installed"] or result["skipped_existing"]):
        n_new = len(result["installed"])
        n_old = len(result["skipped_existing"])
        platform_hint = {
            "win32": "WPS / Word",
            "darwin": "WPS / Pages / Keynote",
            "linux": "LibreOffice / WPS",
        }.get(sys.platform, "")
        msg = f"[font-install] 完成: {n_new} 新装"
        if n_old:
            msg += f", {n_old} 已有"
        if result["failed"]:
            msg += f", {len(result['failed'])} 失败"
        if platform_hint:
            msg += f" — {platform_hint} 重启后生效"
        print(msg)
    return result


def collect_ttfs_for_install(report: dict, font_cache_dir: Path) -> list[Path]:
    """从 font_resolver 的 report 收集应该装的 TTF。

    只装 resolver 解析到的（auto-resolved + cached），不动 SC/TC 巨型 CJK
    （用户系统已有 SimSun/YaHei，WPS 能 fallback）。
    """
    ttfs = []
    seen = set()
    for entry in report.get("resolved", []):
        if entry.get("cjk"):
            continue
        for slot, fname in (entry.get("slots") or {}).items():
            p = font_cache_dir / fname
            if p.exists() and p.name not in seen:
                ttfs.append(p)
                seen.add(p.name)
    return ttfs
