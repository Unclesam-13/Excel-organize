"""
系统集成（Windows / macOS）：
- 开机自启：Windows 注册表 Run；macOS ~/Library/LaunchAgents（RunAtLoad）
- 每日定时整理：Windows 计划任务 schtasks；macOS launchd（StartCalendarInterval）
- 打开文件 / 在文件夹中显示；删除到系统回收站（废纸篓）
"""
from __future__ import annotations

import os
import plistlib
import shutil
import subprocess
import sys
import time
from pathlib import Path

import core

IS_WIN = os.name == "nt"
IS_MAC = sys.platform == "darwin"
SUPPORTED = IS_WIN or IS_MAC          # 支持开机自启/定时任务
CAN_OPEN_LOCAL = IS_WIN or IS_MAC     # 能在本机直接打开文件
RECYCLE_NAME = "Windows 回收站" if IS_WIN else ("废纸篓" if IS_MAC else "系统回收站")

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "ExcelLibrary"
TASK_NAME = "ExcelLibrary_DailyScan"
MAC_AUTOSTART = "com.excellibrary.autostart"
MAC_DAILY = "com.excellibrary.dailyscan"
_NO_WINDOW = 0x08000000 if IS_WIN else 0


def launch_args(*args: str) -> list[str]:
    """启动本程序的命令（打包后的程序，或 python + desktop.py）。"""
    if core.FROZEN:
        parts = [sys.executable]
    else:
        exe = sys.executable
        pyw = os.path.join(os.path.dirname(exe), "pythonw.exe")
        parts = [pyw if (IS_WIN and os.path.exists(pyw)) else exe, str(core.APP_DIR / "desktop.py")]
    return parts + list(args)


def launch_command(*args: str) -> str:
    return " ".join(f'"{p}"' for p in launch_args(*args))


# ---------------------------------------------------------------- macOS launchd
def _agent_path(label: str) -> Path:
    return Path.home() / "Library" / "LaunchAgents" / f"{label}.plist"


def _launchctl(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(["launchctl", *args], capture_output=True, text=True)


def _mac_write_agent(label: str, plist: dict, load: bool) -> None:
    p = _agent_path(label)
    p.parent.mkdir(parents=True, exist_ok=True)
    if p.exists():
        _launchctl("unload", str(p))
    with open(p, "wb") as f:
        plistlib.dump(plist, f)
    if load:
        r = _launchctl("load", "-w", str(p))
        if r.returncode != 0 and r.stderr.strip():
            core.log.warning("launchctl load %s：%s", label, r.stderr.strip())


def _mac_remove_agent(label: str) -> None:
    p = _agent_path(label)
    if p.exists():
        _launchctl("unload", str(p))
        p.unlink()


# ---------------------------------------------------------------- 开机自启
def get_autostart() -> bool:
    if IS_MAC:
        return _agent_path(MAC_AUTOSTART).exists()
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
    if IS_MAC:
        if enabled:
            # 只在下次登录时生效（RunAtLoad），不立即加载，避免再启动一个实例
            _mac_write_agent(MAC_AUTOSTART, {"Label": MAC_AUTOSTART, "ProgramArguments": launch_args("--background"),
                                             "RunAtLoad": True, "ProcessType": "Interactive"}, load=False)
        else:
            _mac_remove_agent(MAC_AUTOSTART)
        return
    if not IS_WIN:
        raise RuntimeError("当前系统不支持开机自启")
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
        if enabled:
            winreg.SetValueEx(k, RUN_NAME, 0, winreg.REG_SZ, launch_command("--background"))
        else:
            try:
                winreg.DeleteValue(k, RUN_NAME)
            except FileNotFoundError:
                pass


# ---------------------------------------------------------------- 每日定时整理
def _schtasks(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(["schtasks", *args], capture_output=True, text=True, creationflags=_NO_WINDOW)


def get_daily_task() -> str | None:
    """返回已设置的每日运行时间 HH:MM，未设置返回 None。"""
    if IS_MAC:
        p = _agent_path(MAC_DAILY)
        if not p.exists():
            return None
        try:
            with open(p, "rb") as f:
                sci = plistlib.load(f).get("StartCalendarInterval", {})
            return f"{int(sci.get('Hour', 8)):02d}:{int(sci.get('Minute', 0)):02d}"
        except Exception:
            return "已设置"
    if not IS_WIN:
        return None
    r = _schtasks("/query", "/tn", TASK_NAME, "/xml")
    if r.returncode != 0:
        return None
    import re
    m = re.search(r"<StartBoundary>[^<]*T(\d{2}:\d{2})", r.stdout)
    return m.group(1) if m else "已设置"


def set_daily_task(enabled: bool, at: str = "08:00") -> None:
    if IS_MAC:
        if enabled:
            hh, mm = (int(x) for x in at.split(":")[:2])
            _mac_write_agent(MAC_DAILY, {"Label": MAC_DAILY, "ProgramArguments": launch_args("--scan"),
                                         "StartCalendarInterval": {"Hour": hh, "Minute": mm},
                                         "StandardErrorPath": str(core.DATA_HOME / "logs" / "launchd_scan.log")},
                             load=True)
        else:
            _mac_remove_agent(MAC_DAILY)
        return
    if not IS_WIN:
        raise RuntimeError("当前系统不支持计划任务")
    if enabled:
        r = _schtasks("/create", "/f", "/sc", "daily", "/st", at, "/tn", TASK_NAME, "/tr", launch_command("--scan"))
    else:
        r = _schtasks("/delete", "/f", "/tn", TASK_NAME)
        if r.returncode != 0 and get_daily_task() is None:
            return
    if r.returncode != 0:
        raise RuntimeError((r.stderr or r.stdout).strip() or "schtasks 执行失败")


def refresh_registrations() -> None:
    """程序换了安装位置后，开机自启和定时任务里记录的程序路径会失效，启动时自动改成当前路径。"""
    if not (SUPPORTED and core.FROZEN):
        return
    exe = sys.executable
    if get_autostart():
        current = ""
        if IS_WIN:
            import winreg
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
                    current = winreg.QueryValueEx(k, RUN_NAME)[0]
            except OSError:
                pass
        else:
            try:
                with open(_agent_path(MAC_AUTOSTART), "rb") as f:
                    current = " ".join(plistlib.load(f).get("ProgramArguments", []))
            except Exception:
                pass
        if exe not in current:
            set_autostart(True)
    at = get_daily_task()
    if at and ":" in at:
        current = ""
        if IS_WIN:
            r = _schtasks("/query", "/tn", TASK_NAME, "/xml")
            current = r.stdout if r.returncode == 0 else ""
        else:
            try:
                with open(_agent_path(MAC_DAILY), "rb") as f:
                    current = " ".join(plistlib.load(f).get("ProgramArguments", []))
            except Exception:
                pass
        if exe not in current:
            set_daily_task(True, at)


# ---------------------------------------------------------------- 文件操作
def open_in_explorer(path, select: bool = False) -> None:
    path = str(path)
    if IS_WIN:
        if select:
            subprocess.Popen(["explorer", "/select,", path])
        else:
            os.startfile(path)  # type: ignore[attr-defined]
    elif IS_MAC:
        subprocess.Popen(["open", "-R", path] if select else ["open", path])
    else:
        subprocess.Popen(["xdg-open", os.path.dirname(path) if select else path])


def send_to_recycle_bin(path) -> bool:
    """把文件放进系统回收站（Windows 回收站 / macOS 废纸篓）。不支持时返回 False。"""
    if IS_MAC:
        trash = Path.home() / ".Trash"
        try:
            trash.mkdir(exist_ok=True)
            dest = trash / Path(path).name
            if dest.exists():
                dest = trash / f"{Path(path).stem} {time.strftime('%H.%M.%S')}{Path(path).suffix}"
            shutil.move(str(path), str(dest))
            return True
        except OSError as e:
            core.log.warning("移到废纸篓失败：%s", e)
            return False
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


def message_box(text: str, title: str = "") -> None:
    title = title or core.APP_NAME
    if IS_WIN:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, text, title, 0x40)
    elif IS_MAC:
        script = f'display dialog {_as_str(text)} with title {_as_str(title)} buttons {{"好"}} default button 1'
        subprocess.run(["osascript", "-e", script], capture_output=True)
    else:
        print(text)


def _as_str(s: str) -> str:
    return '"' + str(s).replace("\\", "\\\\").replace('"', '\\"') + '"'
