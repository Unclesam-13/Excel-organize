"""
会计工作台 - 桌面程序入口（打包为 ExcelLibrary.exe）

  ExcelLibrary.exe               打开程序窗口（已在后台运行时，直接打开窗口连接它）
  ExcelLibrary.exe --background  只在后台运行自动整理（开机自启用）
  ExcelLibrary.exe --scan        整理收件箱一次后退出（每日计划任务用）
"""
from __future__ import annotations

import os
import socket
import sys
import threading
import time
import urllib.request
import webbrowser

# 无控制台(--windowed)的 exe 中 stdout/stderr 为 None，部分库写日志会出错
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

import core  # noqa: E402


def _ping(url: str) -> bool:
    try:
        with urllib.request.urlopen(url + "/ping", timeout=1.5) as r:
            return r.read().decode() == "excel-library"
    except Exception:
        return False


def _port_free(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind((host, port))
            return True
        except OSError:
            return False


def _message_box(text: str, title: str = core.APP_NAME) -> None:
    import winsys
    winsys.message_box(text, title)


def start_server(host: str, port: int):
    from werkzeug.serving import make_server
    import web

    srv = make_server(host, port, web.app, threaded=True)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    threading.Thread(target=web.scanner_loop, daemon=True).start()
    return srv, web


def open_window(url: str) -> bool:
    """用独立窗口打开界面（需要 pywebview + Edge WebView2），失败返回 False。"""
    try:
        import webview
    except ImportError:
        return False
    try:
        try:
            webview.settings["ALLOW_DOWNLOADS"] = True
        except Exception:
            pass
        webview.create_window(core.APP_NAME, url, width=1280, height=840, min_size=(900, 600))
        webview.start()
        return True
    except Exception as e:
        core.log.error("无法打开程序窗口，改用浏览器：%s", e)
        return False


def main() -> None:
    args = set(sys.argv[1:])
    cfg = core.load_config()
    core.setup_logging(cfg)

    if "--scan" in args:
        res = core.scan_inbox(cfg)
        core.log.info("计划任务整理完成：新入库 %s，重复 %s，失败 %s", res["ok"], res["duplicate"], res["failed"])
        return

    host, port = cfg.get("web_host", "127.0.0.1"), int(cfg.get("web_port", 8765))
    url_ = f"http://{'127.0.0.1' if host == '0.0.0.0' else host}:{port}"
    if not _ping(url_):  # 只有第一个实例做迁移，避免和正在运行的程序同时搬文件
        rep = core.migrate_legacy_storage()
        if rep:
            core.log.info("已迁移到新的默认存储位置：搬移 %s 个文件", rep["moved"])
        try:
            import winsys
            winsys.refresh_registrations()
        except Exception as e:
            core.log.warning("更新开机自启/定时任务路径失败：%s", e)
        cfg = core.load_config()
    url = f"http://{'127.0.0.1' if host == '0.0.0.0' else host}:{port}"

    # 已有实例在运行：只打开窗口连接它
    if _ping(url):
        if "--background" not in args and not open_window(url):
            webbrowser.open(url)
        return

    if not _port_free(host, port):
        _message_box(f"端口 {port} 被其他程序占用。\n请在 {core.CONFIG_PATH} 中修改 web_port 后重试。")
        return

    srv, web = start_server(host, port)
    stop = threading.Event()
    web._shutdown_hook["fn"] = lambda: (srv.shutdown(), stop.set(), os._exit(0))
    core.log.info("%s %s 已启动：%s（数据目录 %s）", core.APP_NAME, core.VERSION, url, core.DATA_HOME)

    if "--background" in args:
        stop.wait()
        return

    if open_window(url):
        # 窗口关闭：如果开启了开机自启（说明用户希望常驻），继续在后台整理；否则退出
        import winsys
        if winsys.get_autostart():
            core.log.info("窗口已关闭，继续在后台运行")
            stop.wait()
        return

    # 没有 pywebview 时退回浏览器，进程常驻直到在“设置 > 退出程序”中退出
    webbrowser.open(url)
    try:
        while not stop.is_set():
            time.sleep(1)
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
