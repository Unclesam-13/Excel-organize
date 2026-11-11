"""Windows 系统集成：开机自启（注册表 Run）、每日计划任务（schtasks）。非 Windows 下这些功能不可用。"""
from __future__ import annotations

import os
import subprocess
import sys

import core

IS_WIN = os.name == "nt"
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "ExcelLibrary"
TASK_NAME = "ExcelLibrary_DailyScan"
_NO_WINDOW = 0x08000000 if IS_WIN else 0


def launch_command(*args: str) -> str:
    """启动本程序的命令行（exe 或 pythonw + desktop.py）。"""
    if core.FROZEN:
        parts = [sys.executable]
    else:
        exe = sys.executable
        pyw = os.path.join(os.path.dirname(exe), "pythonw.exe")
        parts = [pyw if os.path.exists(pyw) else exe, str(core.APP_DIR / "desktop.py")]
    return " ".join(f'"{p}"' for p in parts) + ("" if not args else " " + " ".join(args))


# ---------------------------------------------------------------- 开机自启
def get_autostart() -> bool:
    if not IS_WIN:
        return False
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, RUN_NAME)
            return True
    except OSError:
        return False


def set_autostart(enabled: bool) -> None:
    if not IS_WIN:
        raise RuntimeError("仅 Windows 支持开机自启")
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
        if enabled:
            winreg.SetValueEx(k, RUN_NAME, 0, winreg.REG_SZ, launch_command("--background"))
        else:
            try:
                winreg.DeleteValue(k, RUN_NAME)
            except FileNotFoundError:
                pass


# ---------------------------------------------------------------- 每日计划任务
def _schtasks(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(["schtasks", *args], capture_output=True, text=True, creationflags=_NO_WINDOW)


def get_daily_task() -> str | None:
    """返回已设置的每日运行时间 HH:MM，未设置返回 None。"""
    if not IS_WIN:
        return None
    r = _schtasks("/query", "/tn", TASK_NAME, "/xml")
    if r.returncode != 0:
        return None
    import re
    m = re.search(r"<StartBoundary>[^<]*T(\d{2}:\d{2})", r.stdout)
    return m.group(1) if m else "已设置"


def set_daily_task(enabled: bool, at: str = "08:00") -> None:
    if not IS_WIN:
        raise RuntimeError("仅 Windows 支持计划任务")
    if enabled:
        r = _schtasks("/create", "/f", "/sc", "daily", "/st", at, "/tn", TASK_NAME, "/tr", launch_command("--scan"))
    else:
        r = _schtasks("/delete", "/f", "/tn", TASK_NAME)
        if r.returncode != 0 and get_daily_task() is None:
            return
    if r.returncode != 0:
        raise RuntimeError((r.stderr or r.stdout).strip() or "schtasks 执行失败")


def open_in_explorer(path, select: bool = False) -> None:
    path = str(path)
    if IS_WIN:
        if select:
            subprocess.Popen(["explorer", "/select,", path])
        else:
            os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", path])


def send_to_recycle_bin(path) -> bool:
    """把文件放进 Windows 回收站（可在系统回收站还原）。非 Windows 返回 False。"""
    if not IS_WIN:
        return False
    import ctypes
    from ctypes import wintypes

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [("hwnd", wintypes.HWND), ("wFunc", wintypes.UINT), ("pFrom", wintypes.LPCWSTR),
                    ("pTo", wintypes.LPCWSTR), ("fFlags", ctypes.c_ushort), ("fAnyOperationsAborted", wintypes.BOOL),
                    ("hNameMappings", ctypes.c_void_p), ("lpszProgressTitle", wintypes.LPCWSTR)]

    FO_DELETE, FOF_SILENT, FOF_NOCONFIRMATION, FOF_ALLOWUNDO, FOF_NOERRORUI = 3, 0x4, 0x10, 0x40, 0x400
    op = SHFILEOPSTRUCTW(None, FO_DELETE, os.path.abspath(str(path)) + "\0", None,
                         FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI, False, None, None)
    return ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op)) == 0 and not op.fAnyOperationsAborted
