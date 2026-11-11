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
        assert "会计工作台" in html and "图书馆</h1>" not in html, path
    assert "鑫源五金" in html or "供应商" in html, "关键词搜索应命中"
    html = urllib.request.urlopen(base + "/chat?ids=1", timeout=10).read().decode("utf-8")
    assert "已选文件" in html and "#1" in html, "问答页应列出选中文件"
    req = urllib.request.Request(base + "/api/chat", data=json.dumps({"ids": [1], "question": "合计多少"}).encode(),
                                 headers={"Content-Type": "application/json"}, method="POST")
    resp = json.loads(urllib.request.urlopen(req, timeout=20).read().decode("utf-8"))
    assert "API Key" in resp.get("error", ""), f"未设置 Key 时应提示：{resp}"
    post = lambda path, body=b"": urllib.request.urlopen(urllib.request.Request(base + path, data=body, method="POST"), timeout=20)
    # 写入行 -> 档案行数更新 -> 撤销
    def post_json(path, obj):
        req = urllib.request.Request(base + path, data=json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                                     headers={"Content-Type": "application/json"}, method="POST")
        return json.loads(urllib.request.urlopen(req, timeout=30).read().decode("utf-8"))
    w = post_json("/api/write_rows", {"file_id": 1, "rows": [{"供应商": "宁波海通", "物料": "垫片", "报价": "1.2"}]})
    print("写入结果:", w)
    assert w.get("ok") and w["written"] == 1, w
    db = sqlite3.connect(home / "data" / "library.db")
    assert db.execute("SELECT total_rows FROM files WHERE id=1").fetchone()[0] == 3, "写入后应为 3 行"
    assert post_json("/api/undo_write", {"file_id": 1, "backup": w["backup"]}).get("ok")
    assert db.execute("SELECT total_rows FROM files WHERE id=1").fetchone()[0] == 2, "撤销后应恢复为 2 行"
    post("/delete", b"ids=1")
    assert "回收站 (1)" in urllib.request.urlopen(base + "/trash", timeout=10).read().decode("utf-8"), "删除后应在回收站"
    assert not Path(rows[0][2]).exists(), "删除后文件应移出原位置"
    post("/restore?one=1")
    assert Path(rows[0][2]).exists(), "恢复后文件应回到原位置"
    post("/delete", b"ids=1")
    post("/purge?one=1")  # 打包版在 Windows 上会放进系统回收站
    assert sqlite3.connect(home / "data" / "library.db").execute("SELECT COUNT(*) FROM files").fetchone()[0] == 0
    urllib.request.urlopen(urllib.request.Request(base + "/quit", data=b"", method="POST"), timeout=5)
    p.wait(timeout=20)
finally:
    if p.poll() is None:
        p.kill()
        show_log()
print("SMOKE TEST PASSED")
