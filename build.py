"""
打包脚本：
  Windows：dist\\ExcelLibrary\\ExcelLibrary.exe、便携版 zip、安装包（需 Inno Setup 6）
  macOS  ：dist/会计工作台.app、ExcelLibrary_Mac_<芯片>_<版本>.dmg
用法：python build.py   （Windows 上通常直接双击 打包成exe.bat；Mac 上运行 ./打包Mac版.sh）
"""
from __future__ import annotations

import os
import platform
import plistlib
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))
import core  # noqa: E402

VERSION = core.VERSION
APP_NAME = core.APP_NAME
IS_MAC = sys.platform == "darwin"
BUNDLE_ID = "com.excellibrary.accounting"


def common_args() -> list[str]:
    return [
        "desktop.py",
        "--name", "ExcelLibrary",
        "--noconfirm", "--clean", "--windowed",
        "--hidden-import", "openpyxl",
        "--hidden-import", "xlrd",
        "--hidden-import", "web",
        "--hidden-import", "chat",
        "--hidden-import", "excel_writer",
        "--hidden-import", "formula_eval",
        "--hidden-import", "winsys",
        "--collect-all", "webview",
        "--exclude-module", "tkinter",
        "--exclude-module", "matplotlib",
        "--exclude-module", "IPython",
        "--exclude-module", "pytest",
    ]


# ----------------------------------------------------------------------------- Windows
def build_windows() -> list[Path]:
    import PyInstaller.__main__

    ver = ROOT / "build" / "version_info.txt"
    ver.parent.mkdir(exist_ok=True)
    v = tuple(int(x) for x in VERSION.split(".")) + (0,)
    ver.write_text(f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={v}, prodvers={v}),
  kids=[StringFileInfo([StringTable('080404B0', [
    StringStruct('FileDescription', '{APP_NAME}'),
    StringStruct('ProductName', '{APP_NAME}'),
    StringStruct('FileVersion', '{VERSION}'),
    StringStruct('ProductVersion', '{VERSION}'),
    StringStruct('OriginalFilename', 'ExcelLibrary.exe')])]),
    VarFileInfo([VarStruct('Translation', [2052, 1200])])])
""", encoding="utf-8")
    PyInstaller.__main__.run(common_args() + [
        "--icon", "assets/app.ico",
        "--add-data", f"assets/app.ico{os.pathsep}assets",
        "--version-file", str(ver),
    ])
    exe_dir = ROOT / "dist" / "ExcelLibrary"
    shutil.copy2(ROOT / "使用说明.md", exe_dir / "使用说明.md")
    out = [Path(shutil.make_archive(str(ROOT / "dist" / f"ExcelLibrary_Portable_{VERSION}"), "zip",
                                    exe_dir.parent, exe_dir.name))]
    iscc = find_iscc()
    if iscc:
        subprocess.run([iscc, f"/DMyAppVersion={VERSION}", "installer.iss"], check=True)
        out.append(ROOT / "dist" / f"ExcelLibrary_Setup_{VERSION}.exe")
    else:
        print("\n[提示] 未找到 Inno Setup 6，跳过安装包。需要安装包请先安装：https://jrsoftware.org/isdl.php")
    print(f"\n程序：{exe_dir / 'ExcelLibrary.exe'}")
    return out


def find_iscc() -> str | None:
    cands = [shutil.which("iscc")]
    for env in ("ProgramFiles(x86)", "ProgramFiles", "LOCALAPPDATA"):
        base = os.environ.get(env)
        if base:
            cands += [str(Path(base) / "Inno Setup 6" / "ISCC.exe"), str(Path(base) / "Programs" / "Inno Setup 6" / "ISCC.exe")]
    return next((c for c in cands if c and Path(c).exists()), None)


# ----------------------------------------------------------------------------- macOS
def make_icns() -> Path:
    """由 app.ico 生成 Mac 图标 app.icns。"""
    from PIL import Image

    icns = ROOT / "build" / "app.icns"
    icns.parent.mkdir(exist_ok=True)
    ico = Image.open(ROOT / "assets" / "app.ico")
    sizes = sorted(ico.info.get("sizes") or [ico.size])
    ico.size = sizes[-1]
    img = ico.convert("RGBA").resize((1024, 1024), Image.LANCZOS)
    img.save(icns, format="ICNS")
    return icns


def build_mac() -> list[Path]:
    import PyInstaller.__main__

    icns = make_icns()
    PyInstaller.__main__.run(common_args() + [
        "--icon", str(icns),
        "--osx-bundle-identifier", BUNDLE_ID,
    ])
    built = ROOT / "dist" / "ExcelLibrary.app"
    app = ROOT / "dist" / f"{APP_NAME}.app"
    if app.exists():
        shutil.rmtree(app)
    built.rename(app)

    # Info.plist：中文显示名、版本、深色模式、权限说明
    plist_path = app / "Contents" / "Info.plist"
    with open(plist_path, "rb") as f:
        info = plistlib.load(f)
    info.update({
        "CFBundleName": APP_NAME,
        "CFBundleDisplayName": APP_NAME,
        "CFBundleShortVersionString": VERSION,
        "CFBundleVersion": VERSION,
        "LSMinimumSystemVersion": "11.0",
        "NSHighResolutionCapable": True,
        "LSApplicationCategoryType": "public.app-category.finance",
        "NSAppleEventsUsageDescription": "用于在访达中显示文件。",
        "NSDocumentsFolderUsageDescription": "用于在“文稿/会计工作台”中存放收件箱和归档的 Excel 文件。",
        "NSDesktopFolderUsageDescription": "用于读取你放在桌面上的 Excel 文件。",
        "NSDownloadsFolderUsageDescription": "用于读取你下载的 Excel 文件。",
        "NSNetworkVolumesUsageDescription": "用于读取共享盘上的 Excel 文件。",
        "CFBundleDevelopmentRegion": "zh_CN",
    })
    with open(plist_path, "wb") as f:
        plistlib.dump(info, f)
    shutil.copy2(ROOT / "使用说明.md", app / "Contents" / "Resources" / "使用说明.md")
    # 修改 Info.plist 后重新做 ad-hoc 签名（Apple 芯片要求程序必须有签名才能运行）
    subprocess.run(["codesign", "--force", "--deep", "--sign", "-", str(app)], check=True)
    subprocess.run(["codesign", "--verify", "--deep", "--strict", str(app)], check=True)

    arch = "AppleSilicon" if platform.machine() == "arm64" else "Intel"
    dmg = ROOT / "dist" / f"ExcelLibrary_Mac_{arch}_{VERSION}.dmg"
    stage = ROOT / "build" / "dmg"
    if stage.exists():
        shutil.rmtree(stage)
    stage.mkdir(parents=True)
    subprocess.run(["ditto", str(app), str(stage / app.name)], check=True)
    (stage / "Applications").symlink_to("/Applications")
    shutil.copy2(ROOT / "使用说明.md", stage / "使用说明.md")
    if dmg.exists():
        dmg.unlink()
    subprocess.run(["hdiutil", "create", "-volname", f"{APP_NAME} {VERSION}", "-srcfolder", str(stage),
                    "-ov", "-format", "UDZO", "-fs", "HFS+", str(dmg)], check=True)
    print(f"\n程序：{app}")
    return [dmg]


if __name__ == "__main__":
    outputs = build_mac() if IS_MAC else build_windows()
    print("\n================ 打包完成 ================")
    for o in outputs:
        print(o)
