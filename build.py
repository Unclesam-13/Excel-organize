"""
打包脚本（在 Windows 上运行）：生成
  dist\\ExcelLibrary\\ExcelLibrary.exe          绿色版程序
  dist\\ExcelLibrary_便携版.zip                 绿色版压缩包
  dist\\ExcelLibrary_Setup_x.y.z.exe            安装包（需已安装 Inno Setup 6）
用法：python build.py   （通常直接双击 打包成exe.bat）
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))
import core  # noqa: E402

VERSION = core.VERSION


def run_pyinstaller() -> Path:
    import PyInstaller.__main__

    args = [
        "desktop.py",
        "--name", "ExcelLibrary",
        "--noconfirm", "--clean", "--windowed",
        "--icon", "assets/app.ico",
        "--add-data", f"assets/app.ico{os.pathsep}assets",
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
    if os.name == "nt":
        ver = ROOT / "build" / "version_info.txt"
        ver.parent.mkdir(exist_ok=True)
        v = tuple(int(x) for x in VERSION.split(".")) + (0,)
        ver.write_text(f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={v}, prodvers={v}),
  kids=[StringFileInfo([StringTable('080404B0', [
    StringStruct('FileDescription', '会计工作台'),
    StringStruct('ProductName', '会计工作台'),
    StringStruct('FileVersion', '{VERSION}'),
    StringStruct('ProductVersion', '{VERSION}'),
    StringStruct('OriginalFilename', 'ExcelLibrary.exe')])]),
    VarFileInfo([VarStruct('Translation', [2052, 1200])])])
""", encoding="utf-8")
        args += ["--version-file", str(ver)]
    PyInstaller.__main__.run(args)
    exe_dir = ROOT / "dist" / "ExcelLibrary"
    shutil.copy2(ROOT / "使用说明.md", exe_dir / "使用说明.md")
    return exe_dir


def make_zip(exe_dir: Path) -> Path:
    out = ROOT / "dist" / f"ExcelLibrary_Portable_{VERSION}"
    return Path(shutil.make_archive(str(out), "zip", exe_dir.parent, exe_dir.name))


def find_iscc() -> str | None:
    cands = [shutil.which("iscc")]
    for env in ("ProgramFiles(x86)", "ProgramFiles", "LOCALAPPDATA"):
        base = os.environ.get(env)
        if base:
            cands += [str(Path(base) / "Inno Setup 6" / "ISCC.exe"), str(Path(base) / "Programs" / "Inno Setup 6" / "ISCC.exe")]
    return next((c for c in cands if c and Path(c).exists()), None)


def make_installer() -> Path | None:
    iscc = find_iscc()
    if not iscc:
        print("\n[提示] 未找到 Inno Setup 6，跳过安装包。需要安装包请先安装：https://jrsoftware.org/isdl.php")
        return None
    subprocess.run([iscc, f"/DMyAppVersion={VERSION}", "installer.iss"], check=True)
    return ROOT / "dist" / f"ExcelLibrary_Setup_{VERSION}.exe"


if __name__ == "__main__":
    exe_dir = run_pyinstaller()
    z = make_zip(exe_dir)
    setup = make_installer() if os.name == "nt" else None
    print("\n================ 打包完成 ================")
    print(f"程序：  {exe_dir / 'ExcelLibrary.exe'}")
    print(f"便携版：{z}")
    if setup:
        print(f"安装包：{setup}")
