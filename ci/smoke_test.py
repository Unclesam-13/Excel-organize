"""CI 冒烟测试：直接运行打包好的 ExcelLibrary.exe，验证 --scan 整理流程和后台界面服务可用。"""
import json
import os
import sqlite3
import subprocess
import sys
import tempfile
import time
import urllib.parse
import urllib.request
from pathlib import Path

import pandas as pd

EXE = Path(sys.argv[1] if len(sys.argv) > 1 else "dist/ExcelLibrary/ExcelLibrary.exe").resolve()
home = Path(tempfile.mkdtemp(prefix="excellib_"))
(home / "inbox").mkdir()
(home / "config.json").write_text(json.dumps(
    {"inbox_dir": "inbox", "library_dir": "lib", "index_xlsx": "index.xlsx", "web_port": 8799},
    ensure_ascii=False), encoding="utf-8")
sample = home / "inbox" / "供应商报价_20261101.xlsx"
pd.DataFrame({"供应商": ["鑫源五金", "华通钢材"], "物料": ["螺栓", "钢板"], "报价": [0.35, 4200]}).to_excel(sample, index=False)
old = time.time() - 120
os.utime(sample, (old, old))
env = {**os.environ, "EXCEL_LIBRARY_HOME": str(home)}


def show_log():
    log = home / "logs" / "library.log"
    if log.exists():
        print(log.read_text(encoding="utf-8"))


# 1) --scan 模式
r = subprocess.run([str(EXE), "--scan"], env=env, timeout=300)
show_log()
con = sqlite3.connect(home / "data" / "library.db")
rows = con.execute("SELECT category, data_date, stored_path FROM files").fetchall()
print("入库记录:", rows)
assert len(rows) == 1, "扫描后应有 1 条记录"
assert rows[0][1] == "2026-11-01", "业务日期应取自文件名"
assert Path(rows[0][2]).exists(), "归档文件应存在"
assert (home / "index.xlsx").exists(), "应生成总目录"

# 2) --background 模式：界面服务可访问，然后退出
p = subprocess.Popen([str(EXE), "--background"], env=env)
base = "http://127.0.0.1:8799"
try:
    for _ in range(60):
        try:
            if urllib.request.urlopen(base + "/ping", timeout=2).read() == b"excel-library":
                break
        except Exception:
            time.sleep(1)
    else:
        raise AssertionError("后台服务未启动")
    for path in ("/", "/settings", "/?k=" + urllib.parse.quote("螺栓")):
        html = urllib.request.urlopen(base + path, timeout=10).read().decode("utf-8")
        assert "Excel 图书馆" in html, path
    assert "鑫源五金" in html or "供应商" in html, "关键词搜索应命中"
    urllib.request.urlopen(urllib.request.Request(base + "/quit", data=b"", method="POST"), timeout=5)
    p.wait(timeout=20)
finally:
    if p.poll() is None:
        p.kill()
        show_log()
print("SMOKE TEST PASSED")
