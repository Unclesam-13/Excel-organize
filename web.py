"""
会计工作台 - 界面（Flask）。桌面版由 desktop.py 用独立窗口加载；也可单独运行：python web.py
"""
from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser
from pathlib import Path

from flask import Flask, abort, redirect, render_template_string, request, send_file, url_for

import chat as chat_mod
import core
import excel_writer
import winsys

core.setup_logging(core.load_config())
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024
_scan_lock = threading.Lock()
_status = {"last_scan": None, "last_result": None}
_shutdown_hook = {"fn": None}


def C() -> dict:
    """每次请求都重新读取配置，设置页修改后立即生效。"""
    return core.load_config()


def run_scan():
    with _scan_lock:
        res = core.scan_inbox(C())
        _status["last_scan"] = time.strftime("%H:%M:%S")
        _status["last_result"] = res
        return res


def scanner_loop():
    while True:
        cfg = C()
        if cfg.get("web_auto_scan", True):
            try:
                run_scan()
            except Exception:
                core.log.exception("后台扫描出错")
        time.sleep(max(10, int(cfg.get("scan_interval_seconds", 60))))


BASE = r"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>会计工作台</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
:root{--b:#1f4e78;--bg:#f5f7fa;--line:#e3e7ee;--mut:#6b7280}
*{box-sizing:border-box}body{margin:0;font:14px/1.6 "Microsoft YaHei UI","Microsoft YaHei","PingFang SC",sans-serif;background:var(--bg);color:#1f2937}
header{background:var(--b);color:#fff;padding:0 24px;display:flex;align-items:center;gap:6px;height:52px}
header h1{font-size:17px;margin:0 18px 0 0;white-space:nowrap}
header nav a{color:#dbe6f2;text-decoration:none;padding:6px 12px;border-radius:6px}header nav a:hover,header nav a.on{background:rgba(255,255,255,.15);color:#fff}
header .s{opacity:.8;font-size:12px;margin-left:auto;text-align:right}
.wrap{display:flex;gap:20px;padding:20px 24px;align-items:flex-start}
aside{width:200px;flex:none;background:#fff;border:1px solid var(--line);border-radius:8px;padding:12px 14px;position:sticky;top:12px}
aside h3{font-size:13px;color:var(--mut);margin:10px 0 4px}aside a{display:flex;justify-content:space-between;color:#1f2937;text-decoration:none;padding:3px 6px;border-radius:4px}
aside a:hover,aside a.on{background:#e8f0f8;color:var(--b)}
main{flex:1;min-width:0}
.card{background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px 16px;margin-bottom:14px}
.ask{display:flex;gap:8px}.ask input{flex:1;font-size:15px;padding:10px 12px;border:1px solid #c9d3df;border-radius:6px}
button,.btn{background:var(--b);color:#fff;border:0;border-radius:6px;padding:8px 14px;cursor:pointer;font-size:14px;text-decoration:none;display:inline-block;font-family:inherit}
button.ghost.del{color:#b42318;border-color:#f1c4c0}button.ghost.del:hover{background:#fef3f2}
.btn.ghost,button.ghost{background:#fff;color:var(--b);border:1px solid #c9d3df;padding:3px 9px;font-size:12px}
button.danger{background:#b42318}
form.flt{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-top:10px}
input,select,textarea{font-family:inherit;font-size:14px;padding:6px 8px;border:1px solid #c9d3df;border-radius:6px}
.ai{background:#eef6ee;border-color:#cfe5cf}.ai .q{color:var(--mut);font-size:12px}
table{width:100%;border-collapse:collapse;background:#fff}th,td{padding:9px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{background:#f0f3f7;font-weight:600;font-size:13px}td.d{white-space:nowrap}
.tag{display:inline-block;background:#eef2f7;border-radius:10px;padding:0 8px;margin:2px 3px 0 0;font-size:12px;color:#374151}
.cat{display:inline-block;background:#e8f0f8;color:var(--b);border-radius:4px;padding:0 6px;font-size:12px;white-space:nowrap}
.sum{color:#4b5563;font-size:13px}.why{color:#15803d;font-size:13px}.ops{white-space:nowrap}
.msg{background:#fff7e6;border:1px solid #ffd591;padding:8px 12px;border-radius:6px;margin-bottom:12px}
.ok{background:#ecfdf3;border-color:#abefc6}
.up{display:flex;gap:8px;align-items:center;font-size:13px;flex-wrap:wrap}
.pickbar{position:fixed;left:50%;transform:translateX(-50%);bottom:18px;background:#fff;border:1px solid #c9d3df;box-shadow:0 6px 24px rgba(0,0,0,.15);border-radius:10px;padding:10px 14px;display:flex;gap:12px;align-items:center;max-width:92vw;z-index:9}
.pickbar button{white-space:nowrap;flex:none}
.pickbar .names{color:var(--mut);display:inline-block;max-width:420px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;vertical-align:bottom}
input[type=checkbox]{width:16px;height:16px;cursor:pointer}
.set{max-width:860px;margin:20px auto;padding:0 24px}.set h2{font-size:16px;margin:4px 0 12px}
.row{display:grid;grid-template-columns:150px 1fr;gap:10px;align-items:center;margin:10px 0}.row label{color:#374151}
.row .hint{grid-column:2;color:var(--mut);font-size:12px;margin-top:-6px}
.row input[type=text],.row input[type=password],.row textarea{width:100%}
code{background:#f0f3f7;padding:1px 5px;border-radius:4px;font-size:12px}
</style></head><body>
<header><h1>📊 会计工作台</h1>
<nav><a href="{{ url_for('index') }}" class="{{ 'on' if page=='home' }}">文件库</a>
<a href="{{ url_for('chat_page') }}" class="{{ 'on' if page=='chat' }}">AI 助手</a>
<a href="{{ url_for('trash') }}" class="{{ 'on' if page=='trash' }}">回收站{% if trash_n %} ({{ trash_n }}){% endif %}</a>
<a href="{{ url_for('settings') }}" class="{{ 'on' if page=='settings' }}">设置</a>
<a href="{{ url_for('open_dir', which='inbox') }}">打开收件箱</a>
<a href="{{ url_for('open_dir', which='library') }}">打开图书馆</a></nav>
<span class="s">{{ '● DeepSeek 已连接' if ai_on else '○ 未设置 API Key' }}<br>上次自动整理 {{ status.last_scan or '—' }}</span></header>
{% block body %}{% endblock %}
</body></html>"""

HOME = BASE.replace("{% block body %}{% endblock %}", r"""
<div class="wrap">
<aside>
  <a href="{{ url_for('index') }}" class="{{ 'on' if not f.c and not f.m }}"><span>全部</span><span>{{ st.total }}</span></a>
  <h3>按类别</h3>
  {% for c,n in st.by_category %}<a href="{{ url_for('index', c=c) }}" class="{{ 'on' if f.c==c }}"><span>{{ c }}</span><span>{{ n }}</span></a>{% endfor %}
  <h3>按月份</h3>
  {% for m,n in st.by_month[:18] %}<a href="{{ url_for('index', m=m, c=f.c) }}" class="{{ 'on' if f.m==m }}"><span>{{ m }}</span><span>{{ n }}</span></a>{% endfor %}
</aside>
<main>
  {% if msg %}<div class="msg">{{ msg }}</div>{% endif %}
  {% if not ai_on and st.total == 0 %}<div class="msg">👋 欢迎使用！第一步：到 <a href="{{ url_for('settings') }}">设置</a> 填写 DeepSeek API Key；第二步：把 Excel 放进收件箱（顶部“打开收件箱”）或在页面底部上传。</div>{% endif %}
  <div class="card">
    <form class="ask" method="post" action="{{ url_for('ask') }}">
      <input name="q" value="{{ q or '' }}" placeholder="用一句话描述要找的表格，例如：上周华东区的销售日报 / 10月份和某供应商的报价" autofocus>
      <button>AI 找文件</button>
    </form>
    <form class="flt" method="get" action="{{ url_for('index') }}">
      <input name="k" value="{{ f.k or '' }}" placeholder="关键词（空格分隔）" size="22">
      <select name="c"><option value="">全部类别</option>{% for c,n in st.by_category %}<option {{ 'selected' if f.c==c }}>{{ c }}</option>{% endfor %}</select>
      <input type="date" name="from" value="{{ f.df or '' }}"> 至 <input type="date" name="to" value="{{ f.dt or '' }}">
      <button>筛选</button>
      <a class="btn ghost" href="{{ url_for('open_index') }}">打开总目录</a>
    </form>
  </div>
  {% if ai_res %}
  <div class="card ai"><div class="q">理解为：{{ ai_desc }}</div><div><b>AI：</b>{{ ai_res.answer }}</div></div>
  {% endif %}
  <table>
    <tr><th style="width:34px"><input type="checkbox" id="pickAll" title="全选本页"></th><th style="width:96px">日期</th><th style="width:70px">类别</th><th>标题 / 摘要</th><th style="width:200px">操作</th></tr>
    {% for r in rows %}
    <tr>
      <td><input type="checkbox" class="pick" value="{{ r.id }}" data-title="{{ r.title }}"></td>
      <td class="d">{{ r.data_date }}<br><span style="color:#9ca3af;font-size:12px">#{{ r.id }}</span></td>
      <td><span class="cat">{{ r.category }}</span></td>
      <td><b>{{ r.title }}</b>
        {% if r.reason %}<div class="why">★ {{ r.reason }}</div>{% endif %}
        <div class="sum">{{ r.summary }}</div>
        <div>{% for t in r.tags[:6] %}<span class="tag">{{ t }}</span>{% endfor %}
          <span style="color:#9ca3af;font-size:12px">{{ r.original_name }} · {{ r.total_rows }} 行</span></div>
      </td>
      <td class="ops">
        <a class="btn ghost" href="{{ url_for('chat_page', ids=r.id) }}" title="就这个文件的内容向 AI 提问">问 AI</a>
        <a class="btn ghost" href="{{ url_for('open_file', fid=r.id) }}">打开</a>
        <a class="btn ghost" href="{{ url_for('open_folder', fid=r.id) }}">文件夹</a>
        <a class="btn ghost" href="{{ url_for('download', fid=r.id) }}">下载</a>
        <form method="post" action="{{ url_for('delete_files') }}" style="display:inline" data-title="{{ r.title }}" onsubmit="return confirmDel(1, this.dataset.title)">
          <input type="hidden" name="ids" value="{{ r.id }}"><button class="ghost del" title="移到回收站">删除</button></form>
      </td>
    </tr>
    {% else %}<tr><td colspan="5" style="text-align:center;color:#9ca3af;padding:30px">没有找到文件</td></tr>{% endfor %}
  </table>
  <div class="card" style="margin-top:14px">
    <form class="up" method="post" action="{{ url_for('upload') }}" enctype="multipart/form-data">
      <b>上传新文件：</b><input type="file" name="files" multiple accept=".xlsx,.xlsm,.xls,.csv">
      <button>上传并整理</button>
    </form>
    <div class="up" style="margin-top:8px">
      <form method="post" action="{{ url_for('scan_now') }}"><button class="ghost">立即整理收件箱</button></form>
      <span style="color:var(--mut)">收件箱：<code>{{ inbox }}</code></span>
    </div>
  </div>
  <div style="height:70px"></div>
</main></div>
<div id="pickbar" class="pickbar" style="display:none">
  <span>已选 <b id="pickN">0</b> 个文件：<span id="pickNames" class="names"></span></span>
  <button onclick="goChat()">💬 向 AI 提问选中文件的内容</button>
  <button class="ghost del" onclick="deletePicked()">🗑 删除选中</button>
  <button class="ghost" onclick="clearPicks()">清空</button>
</div>
<form id="delForm" method="post" action="{{ url_for('delete_files') }}" style="display:none"><input type="hidden" name="ids" id="delIds"></form>
<script>
const KEY = 'excellib.picked';
let picked = {};
try { picked = JSON.parse(localStorage.getItem(KEY) || '{}') || {}; } catch (e) { picked = {}; }
function save() { try { localStorage.setItem(KEY, JSON.stringify(picked)); } catch (e) {} render(); }
function render() {
  const ids = Object.keys(picked);
  document.getElementById('pickbar').style.display = ids.length ? 'flex' : 'none';
  document.getElementById('pickN').textContent = ids.length;
  document.getElementById('pickNames').textContent = ids.map(i => picked[i]).join('、');
  document.querySelectorAll('.pick').forEach(cb => cb.checked = !!picked[cb.value]);
  const all = document.querySelectorAll('.pick');
  document.getElementById('pickAll').checked = all.length > 0 && [...all].every(cb => cb.checked);
}
document.querySelectorAll('.pick').forEach(cb => cb.addEventListener('change', () => {
  if (cb.checked) {
    if (Object.keys(picked).length >= 10) { cb.checked = false; alert('一次最多选择 10 个文件'); return; }
    picked[cb.value] = cb.dataset.title;
  } else delete picked[cb.value];
  save();
}));
document.getElementById('pickAll').addEventListener('change', e => {
  document.querySelectorAll('.pick').forEach(cb => {
    if (e.target.checked && Object.keys(picked).length < 10) picked[cb.value] = cb.dataset.title;
    if (!e.target.checked) delete picked[cb.value];
  });
  save();
});
function clearPicks() { picked = {}; save(); }
function confirmDel(n, title) {
  const ok = confirm((title ? '删除「' + title + '」？' : '删除选中的 ' + n + ' 个文件？') + '\n\n文件会移到“回收站”，可以随时恢复。');
  return ok;
}
function deletePicked() {
  const ids = Object.keys(picked); if (!ids.length) return;
  if (!confirmDel(ids.length)) return;
  document.getElementById('delIds').value = ids.join(',');
  picked = {}; save();
  document.getElementById('delForm').submit();
}
// 已被删除的文件不再保留在“已选”中
{% if deleted_ids %}{% for i in deleted_ids %}delete picked['{{ i }}'];{% endfor %}save();{% endif %}
function goChat() { location.href = '{{ url_for("chat_page") }}?ids=' + Object.keys(picked).join(','); }
render();
</script>""")

SETTINGS = BASE.replace("{% block body %}{% endblock %}", r"""
<div class="set">
  {% if msg %}<div class="msg {{ 'ok' if ok }}">{{ msg }}</div>{% endif %}
  <form method="post">
  <div class="card"><h2>DeepSeek</h2>
    <div class="row"><label>API Key</label><input type="password" name="api_key" value="{{ raw.deepseek.api_key }}" placeholder="sk-..."></div>
    <div class="row"><span></span><span class="hint">在 platform.deepseek.com 获取。不填也能用（按关键词规则归档）。</span></div>
    <div class="row"><label>模型</label><input type="text" name="model" value="{{ raw.deepseek.model }}"></div>
    <div class="row"><span></span><span class="hint">默认 deepseek-flash；需要更强理解可填 deepseek-v4-pro</span></div>
    <div class="row"><label>识图模型</label><input type="text" name="vision_model" value="{{ raw.vision_model }}" placeholder="留空 = 使用上面的模型"></div>
    <div class="row"><span></span><span class="hint">识别图片用的模型。deepseek-flash 支持图片；如果上面改成了不支持图片的模型，这里填 deepseek-flash。</span></div>
    <div class="row"><label>接口地址</label><input type="text" name="base_url" value="{{ raw.deepseek.base_url }}"></div>
    <div class="row"><span></span><span><button name="action" value="test" class="ghost">保存并测试连接</button></span></div>
  </div>
  <div class="card"><h2>文件夹</h2>
    <div class="row"><label>收件箱</label><input type="text" name="inbox_dir" value="{{ raw.inbox_dir }}"></div>
    <div class="row"><label>图书馆</label><input type="text" name="library_dir" value="{{ raw.library_dir }}"></div>
    <div class="row"><label>总目录文件</label><input type="text" name="index_xlsx" value="{{ raw.index_xlsx }}"></div>
    <div class="row"><span></span><span class="hint">可填共享盘路径，如 \\服务器\共享\收件箱。修改图书馆位置不会移动已归档的文件。</span></div>
    <div class="row"><label>处理后原文件</label><select name="after_ingest">
      <option value="move" {{ 'selected' if raw.after_ingest=='move' }}>移入图书馆（推荐）</option>
      <option value="copy" {{ 'selected' if raw.after_ingest=='copy' }}>复制，保留收件箱原件</option></select></div>
  </div>
  <div class="card"><h2>分类</h2>
    <div class="row"><label>类别与关键词</label><textarea name="categories" rows="10">{{ cats_text }}</textarea></div>
    <div class="row"><span></span><span class="hint">每行一个类别，格式：<code>类别: 关键词1, 关键词2</code>。关键词用于无 AI 时的匹配，也会提示给 AI。</span></div>
    <div class="row"><label>默认类别</label><input type="text" name="default_category" value="{{ raw.default_category }}"></div>
    <div class="row"><label></label><label><input type="checkbox" name="allow_new_category" {{ 'checked' if raw.allow_new_category }}> 允许 AI 自动新建类别</label></div>
  </div>
  <div class="card"><h2>自动运行</h2>
    <div class="row"><label></label><label><input type="checkbox" name="web_auto_scan" {{ 'checked' if raw.web_auto_scan }}> 程序运行时自动整理收件箱，每
      <input type="number" name="scan_interval_seconds" value="{{ raw.scan_interval_seconds }}" min="10" style="width:80px"> 秒检查一次</label></div>
    <div class="row"><label>回收站</label><label>删除的文件保留
      <input type="number" name="trash_keep_days" value="{{ raw.trash_keep_days }}" min="0" style="width:70px"> 天后自动彻底删除（0 = 永不自动删除）</label></div>
    <div class="row"><label></label><label><input type="checkbox" name="autostart" {{ 'checked' if autostart }} {{ 'disabled' if not is_win }}> 开机后在后台自动运行（不弹窗口）</label></div>
    <div class="row"><label></label><label><input type="checkbox" name="daily" {{ 'checked' if daily }} {{ 'disabled' if not is_win }}> 每天定时整理一次（即使程序没打开），时间
      <input type="time" name="daily_at" value="{{ daily if daily and ':' in daily else '08:00' }}" {{ 'disabled' if not is_win }}></label></div>
  </div>
  <div style="display:flex;gap:10px;justify-content:flex-end;margin-bottom:12px"><button name="action" value="save">保存设置</button></div>
  </form>
  <div class="card"><h2>关于</h2>
    <div>版本 {{ version }} · 数据与日志：<code>{{ data_home }}</code></div>
    <div class="up" style="margin-top:10px">
      <a class="btn ghost" href="{{ url_for('open_dir', which='data') }}">打开数据文件夹</a>
      <form method="post" action="{{ url_for('reanalyze_missing') }}"><button class="ghost">用 AI 重新归档未经 AI 处理的文件</button></form>
      <form method="post" action="{{ url_for('quit_app') }}" onsubmit="return confirm('退出后将停止自动整理，确定退出？')"><button class="danger">退出程序</button></form>
    </div>
  </div>
</div>""")


CHAT = BASE.replace("{% block body %}{% endblock %}", r"""
<style>
.chat{display:flex;gap:18px;padding:18px 24px;height:calc(100vh - 52px)}
.files{width:290px;flex:none;background:#fff;border:1px solid var(--line);border-radius:8px;padding:12px;overflow:auto}
.files h3{font-size:13px;color:var(--mut);margin:0 0 8px}
.fi{border:1px solid var(--line);border-radius:6px;padding:8px 10px;margin-bottom:8px;font-size:13px}
.fi b{display:block}.fi .m{color:var(--mut);font-size:12px}.fi .x{float:right;color:#9ca3af;text-decoration:none;font-size:15px;line-height:1}
.fi .x:hover{color:#b42318}
.conv{flex:1;display:flex;flex-direction:column;min-width:0;background:#fff;border:1px solid var(--line);border-radius:8px;position:relative}
#msgs{flex:1;overflow:auto;padding:16px 18px}
.m-u{margin:10px 0 10px auto;max-width:75%;background:#e8f0f8;border-radius:10px 10px 2px 10px;padding:8px 12px;white-space:pre-wrap;width:fit-content}
.m-u img{max-height:120px;max-width:180px;border-radius:6px;margin:4px 4px 0 0;border:1px solid #c9d3df;vertical-align:top}
.m-a{margin:10px 0;max-width:94%;background:#f8fafc;border:1px solid var(--line);border-radius:10px 10px 10px 2px;padding:10px 14px}
.m-a p{margin:4px 0}.m-a table.md{border-collapse:collapse;margin:8px 0;font-size:13px;width:auto}
.m-a table.md th,.m-a table.md td{border:1px solid #d5dce5;padding:4px 8px}.m-a table.md th{background:#eef2f7}
.m-a pre{background:#f0f3f7;padding:8px;border-radius:6px;overflow:auto;font-size:12px}
.m-a h4{margin:8px 0 4px}
.steps{font-size:12px;color:var(--mut);margin-top:6px}.steps summary{cursor:pointer}
.err{color:#b42318}
.prop{border:1px solid #9ec5e8;background:#f3f8fd;border-radius:8px;padding:10px 12px;margin:10px 0}
.prop .t{font-weight:600;color:var(--b)}.prop .r{color:var(--mut);font-size:12px;margin:2px 0 6px}
.prop .wrap{overflow:auto;max-height:340px}
.prop table{border-collapse:collapse;font-size:13px;background:#fff}
.prop th,.prop td{border:1px solid #c9d3df;padding:4px 8px;white-space:nowrap;min-width:60px}
.prop th{background:#e8f0f8;position:sticky;top:0}
.prop td[contenteditable]:focus{outline:2px solid #4a90d9;background:#fffbe6}
.prop td.empty{background:#fafafa}
.prop .acts{margin-top:8px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.prop .ok{color:#15803d}
.inp{display:flex;gap:8px;padding:12px;border-top:1px solid var(--line);align-items:flex-end}
.inp textarea{flex:1;resize:none;height:62px;font-size:14px}
.attach{background:#fff;color:var(--b);border:1px solid #c9d3df;font-size:18px;padding:6px 10px;height:40px}
#thumbs{display:flex;gap:8px;flex-wrap:wrap;padding:0 12px}
#thumbs:not(:empty){padding-top:10px}
.th{position:relative}.th img{height:64px;max-width:110px;border-radius:6px;border:1px solid #c9d3df;object-fit:cover}
.th a{position:absolute;top:-6px;right:-6px;background:#b42318;color:#fff;border-radius:50%;width:18px;height:18px;text-align:center;line-height:17px;font-size:12px;text-decoration:none}
.sugg{padding:8px 12px 0;display:flex;gap:6px;flex-wrap:wrap}
.sugg button{background:#f0f3f7;color:#374151;border:1px solid var(--line);font-size:12px;padding:4px 10px;border-radius:14px}
.hint0{color:var(--mut);text-align:center;margin-top:50px;line-height:1.9}
.drop{display:none;position:absolute;inset:0;background:rgba(31,78,120,.08);border:3px dashed #4a90d9;border-radius:8px;z-index:5;align-items:center;justify-content:center;font-size:18px;color:var(--b)}
.conv.dragging .drop{display:flex}
.dots:after{content:'…';animation:d 1s steps(3) infinite}@keyframes d{0%{content:'.'}50%{content:'..'}100%{content:'...'}}
</style>
<div class="chat">
  <div class="files">
    <h3>已选文件（{{ files|length }}）</h3>
    {% for f in files %}
    <div class="fi"><a class="x" href="#" onclick="removeFile({{ f.id }});return false" title="移除">×</a>
      <b>#{{ f.id }} {{ f.title }}</b>
      <div class="m">{{ f.data_date }} · {{ f.category }} · {{ f.total_rows }} 行</div>
      <div class="m">{% for s in f.sheets %}{{ s.name }}({{ s.rows }}行){% if not loop.last %}、{% endif %}{% endfor %}</div>
      <div style="margin-top:4px"><a class="btn ghost" href="{{ url_for('open_file', fid=f.id) }}">打开</a></div>
    </div>
    {% else %}
    <div class="m" style="color:var(--mut);font-size:13px;line-height:1.8">没有选中文件也可以使用：<br>· 发图片让 AI 识别，它会在资料库里找合适的表格登记<br>· 直接问资料库里的数据<br><br>想针对某几个表格提问，请到 <a href="{{ url_for('index') }}">文件库</a> 勾选（最多 10 个）。</div>
    {% endfor %}
    {% if files %}<a class="btn ghost" href="{{ url_for('index') }}">＋ 回文件库添加更多</a>{% endif %}
  </div>
  <div class="conv" id="conv">
    <div class="drop">松开鼠标，添加图片</div>
    <div id="msgs">
      {% if not ai_on %}<div class="hint0 err">还没有设置 DeepSeek API Key，请先到 <a href="{{ url_for('settings') }}">设置</a> 里填写。</div>
      {% else %}<div class="hint0" id="hello">可以直接问表格里的内容，比如汇总、对比、查找某条记录、找异常。<br>
        📷 也可以 <b>粘贴截图 / 拖入图片 / 点 📎 选择图片</b>（发票、收据、回单、报销单、表格截图等），<br>AI 会识别内容，整理好放到对应的表格里，你确认后才会写入。</div>{% endif %}
    </div>
    <div class="sugg">
      {% if files %}<button onclick="ask(this.textContent)">概括这几个表格的主要内容</button>
      <button onclick="ask(this.textContent)">各文件的关键数字分别是多少？做个对比表</button>
      <button onclick="ask(this.textContent)">有没有异常或需要注意的数据？</button>{% endif %}
      <button onclick="document.getElementById('pick').click()">📷 识别图片并登记到表格</button>
      <button onclick="newChat()" title="清空对话">🗑 新对话</button>
    </div>
    <div id="thumbs"></div>
    <div class="inp">
      <input type="file" id="pick" accept="image/*" multiple style="display:none">
      <button class="attach" title="添加图片（也可直接粘贴截图或拖入）" onclick="document.getElementById('pick').click()" {{ 'disabled' if not ai_on }}>📎</button>
      <textarea id="q" placeholder="输入问题，Enter 发送，Shift+Enter 换行；可 Ctrl+V 粘贴截图" {{ 'disabled' if not ai_on }}></textarea>
      <button id="send" onclick="ask()" {{ 'disabled' if not ai_on }}>发送</button>
    </div>
  </div>
</div>
<script>
const IDS = {{ ids|tojson }};
const KEY = 'excellib.picked';
const MAX_IMG = 10;
let history = [];
let busy = false;
let pending = [];   // 待发送图片（data URL）
if (!IDS.length && !sessionStorage.getItem('excellib.nofiles')) {
  try {
    const p = JSON.parse(localStorage.getItem(KEY) || '{}');
    if (Object.keys(p).length) location.replace('?ids=' + Object.keys(p).join(','));
  } catch (e) {}
}
function esc(s) { return String(s).replace(/[&<>"]/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c])); }
function inline(s) { return esc(s).replace(/`([^`]+)`/g, '<code>$1</code>').replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>'); }
function md(text) {
  const L = String(text).replace(/\r/g, '').split('\n'); const out = []; let i = 0;
  const isList = l => /^\s*([-*]|\d+\.)\s+/.test(l);
  while (i < L.length) {
    const l = L[i];
    if (l.startsWith('```')) { const b = []; i++; while (i < L.length && !L[i].startsWith('```')) b.push(L[i++]); i++; out.push('<pre>' + esc(b.join('\n')) + '</pre>'); continue; }
    if (/^\s*\|/.test(l)) {
      const rows = []; while (i < L.length && /^\s*\|/.test(L[i])) rows.push(L[i++]);
      const cells = r => r.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => c.trim());
      const body = rows.slice(1).filter(r => !/^\s*\|?\s*:?-{2,}/.test(r));
      out.push('<table class="md"><tr>' + cells(rows[0]).map(c => '<th>' + inline(c) + '</th>').join('') + '</tr>' +
        body.map(r => '<tr>' + cells(r).map(c => '<td>' + inline(c) + '</td>').join('') + '</tr>').join('') + '</table>');
      continue;
    }
    const h = l.match(/^#{1,4}\s+(.*)/); if (h) { out.push('<h4>' + inline(h[1]) + '</h4>'); i++; continue; }
    if (isList(l)) {
      const ord = /^\s*\d+\./.test(l); const it = [];
      while (i < L.length && isList(L[i])) it.push(L[i++].replace(/^\s*([-*]|\d+\.)\s+/, ''));
      out.push((ord ? '<ol>' : '<ul>') + it.map(x => '<li>' + inline(x) + '</li>').join('') + (ord ? '</ol>' : '</ul>')); continue;
    }
    if (!l.trim()) { i++; continue; }
    const para = [];
    while (i < L.length && L[i].trim() && !/^\s*(\||```|#{1,4}\s)/.test(L[i]) && !isList(L[i])) para.push(L[i++]);
    out.push('<p>' + para.map(inline).join('<br>') + '</p>');
  }
  return out.join('');
}
function add(cls, html) {
  const hello = document.getElementById('hello'); if (hello) hello.remove();
  const d = document.createElement('div'); d.className = cls; d.innerHTML = html;
  const box = document.getElementById('msgs'); box.appendChild(d); box.scrollTop = box.scrollHeight; return d;
}

// ---------- 图片：选择 / 粘贴 / 拖入，发送前在本机压缩 ----------
function shrink(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => {
      const img = new Image();
      img.onload = () => {
        const MAX = 2200; let w = img.width, h = img.height;
        const k = Math.min(1, MAX / Math.max(w, h)); w = Math.round(w * k); h = Math.round(h * k);
        const c = document.createElement('canvas'); c.width = w; c.height = h;
        const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, w, h); g.drawImage(img, 0, 0, w, h);
        resolve(c.toDataURL('image/jpeg', 0.9));
      };
      img.onerror = () => reject('无法读取图片：' + file.name);
      img.src = fr.result;
    };
    fr.onerror = () => reject('无法读取图片：' + file.name);
    fr.readAsDataURL(file);
  });
}
async function addImages(list) {
  for (const f of list) {
    if (!f.type || !f.type.startsWith('image/')) continue;
    if (pending.length >= MAX_IMG) { alert('一次最多 ' + MAX_IMG + ' 张图片'); break; }
    try { pending.push(await shrink(f)); } catch (e) { alert(e); }
  }
  drawThumbs();
}
function drawThumbs() {
  document.getElementById('thumbs').innerHTML = pending.map((u, i) =>
    '<div class="th"><img src="' + u + '"><a href="#" onclick="pending.splice(' + i + ',1);drawThumbs();return false">×</a></div>').join('');
}
document.getElementById('pick').addEventListener('change', e => { addImages([...e.target.files]); e.target.value = ''; });
document.addEventListener('paste', e => {
  const imgs = [...(e.clipboardData || {}).items || []].filter(it => it.type.startsWith('image/')).map(it => it.getAsFile());
  if (imgs.length) { e.preventDefault(); addImages(imgs); }
});
const conv = document.getElementById('conv');
['dragenter', 'dragover'].forEach(ev => conv.addEventListener(ev, e => { e.preventDefault(); conv.classList.add('dragging'); }));
['dragleave', 'drop'].forEach(ev => conv.addEventListener(ev, e => { e.preventDefault(); if (ev === 'drop' || e.target === conv.querySelector('.drop')) conv.classList.remove('dragging'); }));
conv.addEventListener('drop', e => addImages([...e.dataTransfer.files]));

// ---------- 写入预览卡片 ----------
const PROPS = {};
function propCard(p) {
  PROPS[p.id] = p;
  const target = p.file_id ? ('文件#' + p.file_id + ' ' + esc(p.file_title) + ' 「' + esc(p.sheet) + '」末尾')
                           : ('新建表格「' + esc(p.new_file_name) + '」（会自动整理进文件库）');
  const head = '<tr><th>#</th>' + p.columns.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>';
  const body = p.rows.map((r, i) => '<tr><td style="color:#9ca3af">' + (i + 1) + '</td>' + p.columns.map(c => {
    const v = r[c] === undefined || r[c] === null ? '' : r[c];
    return '<td contenteditable="true"' + (v === '' ? ' class="empty"' : '') + '>' + esc(v) + '</td>';
  }).join('') + '</tr>').join('');
  return '<div class="prop" id="p_' + p.id + '"><div class="t">📥 写入预览：' + target + '，' + p.rows.length + ' 行</div>' +
    (p.reason ? '<div class="r">' + esc(p.reason) + '</div>' : '') +
    '<div class="r">请核对，单元格可以直接点击修改；空白的单元格不写入，带公式的列（如价税合计）会沿用上一行的公式自动计算。写入前会自动备份，可撤销。</div>' +
    '<div class="wrap"><table>' + head + body + '</table></div>' +
    '<div class="acts"><button onclick="writeProp(\'' + p.id + '\')">✅ 写入 Excel</button>' +
    '<button class="ghost" onclick="dropProp(\'' + p.id + '\')">忽略</button><span class="st"></span></div></div>';
}
function readTable(id) {
  const p = PROPS[id]; const rows = [];
  document.querySelectorAll('#p_' + id + ' table tr').forEach((tr, i) => {
    if (i === 0) return;
    const tds = [...tr.querySelectorAll('td')].slice(1); const r = {};
    tds.forEach((td, j) => { const v = td.innerText.trim(); if (v !== '') r[p.columns[j]] = v; });
    if (Object.keys(r).length) rows.push(r);
  });
  return rows;
}
async function writeProp(id) {
  const p = PROPS[id]; const card = document.getElementById('p_' + id); const st = card.querySelector('.st');
  const rows = readTable(id);
  if (!rows.length) { st.innerHTML = '<span class="err">没有要写入的数据</span>'; return; }
  card.querySelectorAll('.acts button').forEach(b => b.disabled = true); st.innerHTML = '<span class="dots">正在写入</span>';
  try {
    const r = await fetch('{{ url_for("api_write_rows") }}', {method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({file_id: p.file_id, sheet: p.sheet, columns: p.columns, rows: rows, new_file_name: p.new_file_name})});
    const d = await r.json();
    if (d.error) { st.innerHTML = '<span class="err">' + esc(d.error) + '</span>'; card.querySelectorAll('.acts button').forEach(b => b.disabled = false); return; }
    card.querySelectorAll('td[contenteditable]').forEach(td => td.removeAttribute('contenteditable'));
    card.querySelector('.acts').innerHTML = '<span class="ok">✅ 已写入 ' + d.written + ' 行到「' + esc(d.title) + '」' +
      (d.first_row ? '（从第 ' + d.first_row + ' 行开始' + (d.before_total ? '，在合计行之前' : '') + '）' : '') + '</span>' +
      (d.file_id ? '<a class="btn ghost" href="/open/' + d.file_id + '">打开文件</a>' : '<span class="r">新表已放入收件箱，稍后自动整理</span>') +
      '<button class="ghost" onclick="undoWrite(\'' + id + '\')">撤销</button>';
    WRITES[id] = d;
    history.push({role: 'assistant', content: '（系统提示：用户已确认，' + rows.length + ' 行数据已写入「' + d.title + '」。）'});
  } catch (e) { st.innerHTML = '<span class="err">请求失败：' + esc(e) + '</span>'; card.querySelectorAll('.acts button').forEach(b => b.disabled = false); }
}
const WRITES = {};
async function undoWrite(id) {
  const d = WRITES[id];
  if (!confirm('撤销这次写入？' + (d.created ? '新建的表格会移到回收站。' : '文件会恢复到写入前的状态。'))) return;
  const card = document.getElementById('p_' + id);
  const r = await fetch('{{ url_for("api_undo_write") }}', {method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({file_id: d.file_id, backup: d.backup, created: !!d.created})});
  const res = await r.json();
  card.querySelector('.acts').innerHTML = res.error ? '<span class="err">' + esc(res.error) + '</span>' : '<span class="r">↩ 已撤销这次写入</span>';
  if (!res.error) history.push({role: 'assistant', content: '（系统提示：用户撤销了刚才写入「' + d.title + '」的数据。）'});
}
function dropProp(id) { document.getElementById('p_' + id).querySelector('.acts').innerHTML = '<span class="r">已忽略</span>'; }

// ---------- 发送 ----------
async function ask(text) {
  const q = (text || document.getElementById('q').value).trim();
  const imgs = pending.slice();
  if ((!q && !imgs.length) || busy) return;
  busy = true; document.getElementById('send').disabled = true; document.getElementById('q').value = '';
  pending = []; drawThumbs();
  add('m-u', esc(q || '（识别图片）') + (imgs.length ? '<div>' + imgs.map(u => '<img src="' + u + '">').join('') + '</div>' : ''));
  const a = add('m-a', '<span class="dots">' + (imgs.length ? '正在识别图片' : '正在读取表格并分析') + '</span>');
  try {
    const r = await fetch('{{ url_for("api_chat") }}', {method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ids: IDS, question: q, history: history, images: imgs})});
    const d = await r.json();
    if (d.error) { a.innerHTML = '<span class="err">出错了：' + esc(d.error) + '</span>'; }
    else {
      let html = md(d.answer);
      (d.proposals || []).forEach(p => html += propCard(p));
      if (d.steps && d.steps.length) html += '<details class="steps"><summary>🔍 处理了 ' + d.steps.length + ' 步</summary>' + d.steps.map(s => '<div>· ' + esc(s) + '</div>').join('') + '</details>';
      a.innerHTML = html;
      history.push({role: 'user', content: d.user_content || q}, {role: 'assistant', content: d.answer});
    }
  } catch (e) { a.innerHTML = '<span class="err">请求失败：' + esc(e) + '</span>'; }
  busy = false; document.getElementById('send').disabled = false; document.getElementById('q').focus();
  const box = document.getElementById('msgs'); box.scrollTop = box.scrollHeight;
}
function newChat() { history = []; document.getElementById('msgs').innerHTML = ''; }
function removeFile(id) {
  try { const p = JSON.parse(localStorage.getItem(KEY) || '{}'); delete p[id]; localStorage.setItem(KEY, JSON.stringify(p)); } catch (e) {}
  const rest = IDS.filter(x => x !== id);
  if (!rest.length) try { sessionStorage.setItem('excellib.nofiles', '1'); } catch (e) {}
  location.href = '?ids=' + rest.join(',');
}
const ta = document.getElementById('q');
if (ta) ta.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); ask(); } });
</script>""")


TRASH = BASE.replace("{% block body %}{% endblock %}", r"""
<div class="set" style="max-width:1100px">
  {% if msg %}<div class="msg">{{ msg }}</div>{% endif %}
  <div class="card" style="display:flex;align-items:center;gap:12px">
    <div style="flex:1"><b>回收站</b>：共 {{ rows|length }} 个文件。删除的文件在这里可以恢复；
      {% if keep_days %}超过 {{ keep_days }} 天会自动彻底删除。{% else %}不会自动清理。{% endif %}
      彻底删除的文件会进入 Windows 回收站。</div>
    {% if rows %}<form method="post" action="{{ url_for('purge') }}" onsubmit="return confirm('彻底删除回收站中的全部 {{ rows|length }} 个文件？此操作无法在本软件中恢复。')">
      <input type="hidden" name="all" value="1"><button class="danger">清空回收站</button></form>{% endif %}
  </div>
  <form method="post" id="tf">
  <table>
    <tr><th style="width:34px"><input type="checkbox" onclick="document.querySelectorAll('.tp').forEach(c=>c.checked=this.checked)"></th>
      <th style="width:150px">删除时间</th><th style="width:96px">业务日期</th><th style="width:70px">类别</th><th>标题</th><th style="width:170px">操作</th></tr>
    {% for r in rows %}
    <tr><td><input type="checkbox" class="tp" name="ids" value="{{ r.id }}"></td>
      <td class="d">{{ r.deleted_at.replace('T',' ') }}</td><td class="d">{{ r.data_date }}</td>
      <td><span class="cat">{{ r.category }}</span></td>
      <td><b>{{ r.title }}</b><div class="sum">{{ r.original_name }} · 原位置：{{ r.prev_path }}</div></td>
      <td class="ops">
        <button class="ghost" formaction="{{ url_for('restore', one=r.id) }}">恢复</button>
        <button class="ghost del" formaction="{{ url_for('purge', one=r.id) }}" data-title="{{ r.title }}" onclick="return confirm('彻底删除「' + this.dataset.title + '」？')">彻底删除</button>
      </td></tr>
    {% else %}<tr><td colspan="6" style="text-align:center;color:#9ca3af;padding:30px">回收站是空的</td></tr>{% endfor %}
  </table>
  {% if rows %}<div style="margin-top:12px;display:flex;gap:8px">
    <button class="ghost" formaction="{{ url_for('restore') }}">恢复选中</button>
    <button class="ghost del" formaction="{{ url_for('purge') }}" onclick="return confirm('彻底删除选中的文件？')">彻底删除选中</button>
  </div>{% endif %}
  </form>
</div>""")


def _form_ids() -> list[int]:
    raw = request.form.getlist("ids") + [request.args.get("one", "")]
    return [int(x) for part in raw for x in str(part).split(",") if x.strip().isdigit()]


@app.post("/delete")
def delete_files():
    ids = _form_ids()
    cfg = C()
    con = core.connect(cfg)
    try:
        with _scan_lock:
            n = core.trash_files(cfg, con, ids)
    except OSError as e:
        return redirect(url_for("index", msg=f"删除失败：{e}（文件可能正被 Excel 打开，请关闭后重试）"))
    finally:
        con.close()
    return redirect(url_for("index", msg=f"已将 {n} 个文件移到回收站，可在“回收站”中恢复。",
                            deleted=",".join(map(str, ids))))


@app.get("/trash")
def trash():
    cfg = C()
    con = core.connect(cfg)
    try:
        rows = core.list_trash(con)
    finally:
        con.close()
    return render_template_string(TRASH, rows=rows, msg=request.args.get("msg"),
                                  keep_days=int(cfg.get("trash_keep_days", 30) or 0), **_common("trash"))


@app.post("/restore")
def restore():
    cfg = C()
    con = core.connect(cfg)
    try:
        with _scan_lock:
            n = core.restore_files(cfg, con, _form_ids())
    except OSError as e:
        return redirect(url_for("trash", msg=f"恢复失败：{e}"))
    finally:
        con.close()
    return redirect(url_for("trash", msg=f"已恢复 {n} 个文件到原位置。"))


@app.post("/purge")
def purge():
    cfg = C()
    con = core.connect(cfg)
    try:
        with _scan_lock:
            n = core.purge_files(cfg, con, None if request.form.get("all") else _form_ids())
    finally:
        con.close()
    return redirect(url_for("trash", msg=f"已彻底删除 {n} 个文件（已放入 Windows 回收站）。"))


def _common(page: str) -> dict:
    cfg = C()
    con = core.connect(cfg)
    try:
        trash_n = con.execute("SELECT COUNT(*) FROM files WHERE deleted_at IS NOT NULL").fetchone()[0]
    finally:
        con.close()
    return {"page": page, "status": _status, "ai_on": core.DeepSeek(cfg).enabled, "cfg": cfg, "trash_n": trash_n}


def render_home(rows, q=None, ai_res=None, msg=None, **f):
    cfg = C()
    con = core.connect(cfg)
    try:
        st = core.stats(con)
    finally:
        con.close()
    ai_desc = ""
    if ai_res:
        fl = ai_res["filters"]
        parts = []
        if fl["date_from"] or fl["date_to"]:
            parts.append(f"日期 {fl['date_from'] or '…'} ~ {fl['date_to'] or '…'}")
        if fl["categories"]:
            parts.append("类别 " + "/".join(fl["categories"]))
        if fl["keywords"]:
            parts.append("关键词 " + "、".join(fl["keywords"]))
        ai_desc = "；".join(parts) or "不限条件"
    deleted_ids = [i for i in request.args.get("deleted", "").split(",") if i.isdigit()]
    return render_template_string(HOME, rows=rows, q=q, ai_res=ai_res, ai_desc=ai_desc, deleted_ids=deleted_ids,
                                  msg=msg or request.args.get("msg"), st=st, f=f, inbox=cfg["inbox_dir"],
                                  **_common("home"))


def _parse_ids(raw) -> list[int]:
    if isinstance(raw, str):
        raw = raw.split(",")
    out = []
    for x in raw or []:
        try:
            out.append(int(x))
        except (TypeError, ValueError):
            pass
    return list(dict.fromkeys(out))[:chat_mod.MAX_FILES]


@app.get("/chat")
def chat_page():
    ids = _parse_ids(request.args.get("ids", ""))
    con = core.connect(C())
    try:
        files = [core.row_to_dict(r) for i in ids
                 for r in [con.execute("SELECT * FROM files WHERE id=? AND deleted_at IS NULL", (i,)).fetchone()] if r]
    finally:
        con.close()
    return render_template_string(CHAT, files=files, ids=[f["id"] for f in files], msg=None, **_common("chat"))


@app.post("/api/chat")
def api_chat():
    data = request.get_json(silent=True) or {}
    question = str(data.get("question", "")).strip()
    if not question and not data.get("images"):
        return {"error": "问题不能为空"}
    cfg = C()
    con = core.connect(cfg)
    try:
        return chat_mod.ask(con, cfg, core.DeepSeek(cfg), _parse_ids(data.get("ids")), question,
                            history=data.get("history") or [], images=data.get("images") or [])
    except (RuntimeError, FileNotFoundError) as e:
        core.log.warning("AI 问答：%s", e)
        return {"error": str(e)}
    except Exception as e:
        core.log.exception("AI 问答出错")
        return {"error": f"{type(e).__name__}: {e}"}
    finally:
        con.close()


@app.post("/api/write_rows")
def api_write_rows():
    d = request.get_json(silent=True) or {}
    rows = [{str(k): v for k, v in r.items() if v not in (None, "")} for r in (d.get("rows") or []) if isinstance(r, dict)]
    rows = [r for r in rows if r]
    if not rows:
        return {"error": "没有要写入的数据"}
    cfg = C()
    con = core.connect(cfg)
    try:
        with _scan_lock:
            if d.get("file_id"):
                fid = int(d["file_id"])
                row = con.execute("SELECT stored_path, title FROM files WHERE id=? AND deleted_at IS NULL", (fid,)).fetchone()
                if not row:
                    return {"error": "目标文件不存在或已被删除"}
                res = excel_writer.append_rows(cfg, Path(row["stored_path"]), rows, d.get("sheet"))
                core.refresh_file(cfg, con, fid)
                core.export_index(cfg, con)
                return {"ok": True, "file_id": fid, "title": row["title"], **res}
            name = core._sanitize_name(d.get("new_file_name") or "AI新建表格", 80)
            if not name.lower().endswith(".xlsx"):
                name += ".xlsx"
            cfg["inbox_dir"].mkdir(parents=True, exist_ok=True)
            path = core._unique_path(cfg["inbox_dir"] / name)
            res = excel_writer.create_workbook(path, d.get("columns") or list(dict.fromkeys(k for r in rows for k in r)), rows)
            old = time.time() - 60
            os.utime(path, (old, old))
            scan = core.scan_inbox(cfg)
            item = next((i for i in scan.get("items", []) if i.get("name") == path.name and i.get("status") == "ok"), None)
            return {"ok": True, "created": True, "file_id": item["id"] if item else None, "title": path.name,
                    "written": res["written"], "first_row": None, "before_total": False}
    except excel_writer.WriteError as e:
        return {"error": str(e)}
    except Exception as e:
        core.log.exception("写入 Excel 出错")
        return {"error": f"{type(e).__name__}: {e}"}
    finally:
        con.close()


@app.post("/api/undo_write")
def api_undo_write():
    d = request.get_json(silent=True) or {}
    cfg = C()
    con = core.connect(cfg)
    try:
        with _scan_lock:
            fid = int(d.get("file_id") or 0)
            if d.get("created"):
                if fid:
                    core.trash_files(cfg, con, [fid])
                return {"ok": True}
            row = con.execute("SELECT stored_path FROM files WHERE id=?", (fid,)).fetchone()
            if not row:
                return {"error": "文件不存在"}
            excel_writer.restore_backup(cfg, str(d.get("backup") or ""), Path(row["stored_path"]))
            core.refresh_file(cfg, con, fid)
            core.export_index(cfg, con)
            core.log.info("已撤销写入 #%s（恢复备份 %s）", fid, d.get("backup"))
            return {"ok": True}
    except excel_writer.WriteError as e:
        return {"error": str(e)}
    except Exception as e:
        core.log.exception("撤销写入出错")
        return {"error": f"{type(e).__name__}: {e}"}
    finally:
        con.close()


@app.get("/ping")
def ping():
    return "excel-library"


@app.get("/")
def index():
    k, c, m = request.args.get("k", ""), request.args.get("c", ""), request.args.get("m", "")
    df, dt_ = request.args.get("from") or None, request.args.get("to") or None
    if m and not (df or dt_):
        df, dt_ = f"{m}-01", f"{m}-31"
    con = core.connect(C())
    try:
        rows = core.search(con, k.split(), [c] if c else [], df, dt_, limit=200)
    finally:
        con.close()
    return render_home(rows, k=k, c=c, m=m, df=request.args.get("from"), dt=request.args.get("to"))


@app.post("/ask")
def ask():
    q = request.form.get("q", "").strip()
    if not q:
        return redirect(url_for("index"))
    cfg = C()
    con = core.connect(cfg)
    try:
        res = core.ask(con, q, cfg, core.DeepSeek(cfg))
    finally:
        con.close()
    return render_home(res["results"], q=q, ai_res=res)


def _get_path(fid: int) -> Path:
    con = core.connect(C())
    try:
        row = con.execute("SELECT stored_path FROM files WHERE id=?", (fid,)).fetchone()
    finally:
        con.close()
    if not row or not Path(row[0]).exists():
        abort(404, "文件不存在或已被移动")
    return Path(row[0])


def _back(msg: str | None = None):
    ref = request.referrer or url_for("index")
    if msg:
        return redirect(url_for("index", msg=msg))
    return redirect(ref)


@app.get("/open/<int:fid>")
def open_file(fid):
    p = _get_path(fid)
    if not winsys.IS_WIN:
        return send_file(p, as_attachment=True)
    winsys.open_in_explorer(p)
    return _back()


@app.get("/folder/<int:fid>")
def open_folder(fid):
    winsys.open_in_explorer(_get_path(fid), select=True)
    return _back()


@app.get("/dir/<which>")
def open_dir(which):
    cfg = C()
    target = {"inbox": cfg["inbox_dir"], "library": cfg["library_dir"], "data": core.DATA_HOME}.get(which)
    if not target:
        abort(404)
    Path(target).mkdir(parents=True, exist_ok=True)
    winsys.open_in_explorer(target)
    return _back()


@app.get("/download/<int:fid>")
def download(fid):
    p = _get_path(fid)
    return send_file(p, as_attachment=True, download_name=p.name)


@app.get("/index-xlsx")
def open_index():
    p = core.export_index(C())
    if winsys.IS_WIN:
        winsys.open_in_explorer(p)
        return _back()
    return send_file(p, as_attachment=True, download_name=p.name)


@app.post("/upload")
def upload():
    cfg = C()
    saved = 0
    cfg["inbox_dir"].mkdir(parents=True, exist_ok=True)
    for f in request.files.getlist("files"):
        name = core._sanitize_name(Path(f.filename or "").name, 150)
        if Path(name).suffix.lower() in core.EXCEL_EXTS:
            dest = core._unique_path(cfg["inbox_dir"] / name)
            f.save(dest)
            os.utime(dest, (time.time() - 60, time.time() - 60))
            saved += 1
    if not saved:
        return redirect(url_for("index", msg="没有可识别的 Excel/CSV 文件"))
    res = run_scan()
    return redirect(url_for("index", msg=f"已上传 {saved} 个；新入库 {res['ok']}，重复 {res['duplicate']}，失败 {res['failed']}"))


@app.post("/scan")
def scan_now():
    res = run_scan()
    if res.get("busy"):
        return redirect(url_for("index", msg="另一个任务正在整理，请稍后刷新"))
    return redirect(url_for("index", msg=f"整理完成：新入库 {res['ok']}，重复 {res['duplicate']}，失败 {res['failed']}，"
                                         f"等待中 {res['waiting']}（文件被 Excel 打开时会等关闭后再处理）"))


# ------------------------------------------------------------------ 设置
def _cats_to_text(cats: dict) -> str:
    return "\n".join(f"{k}: {', '.join(v)}" for k, v in cats.items())


def _text_to_cats(text: str) -> dict:
    import re
    cats = {}
    for line in text.splitlines():
        if not line.strip():
            continue
        name, _, words = line.partition(":") if ":" in line else line.partition("：")
        name = core._sanitize_name(name.strip(), 20)
        if name:
            cats[name] = [w.strip() for w in re.split(r"[,，、;；]", words) if w.strip()]
    return cats


def render_settings(msg=None, ok=False):
    raw = core.load_raw_config()
    return render_template_string(SETTINGS, raw=raw, cats_text=_cats_to_text(raw["categories"]), msg=msg, ok=ok,
                                  autostart=winsys.get_autostart(), daily=winsys.get_daily_task(),
                                  is_win=winsys.IS_WIN, version=core.VERSION, data_home=core.DATA_HOME,
                                  **_common("settings"))


@app.get("/settings")
def settings():
    return render_settings(request.args.get("msg"))


@app.post("/settings")
def settings_save():
    fm = request.form
    raw = core.load_raw_config()
    raw["deepseek"]["api_key"] = fm.get("api_key", "").strip()
    raw["deepseek"]["model"] = fm.get("model", "").strip() or "deepseek-flash"
    raw["deepseek"]["base_url"] = fm.get("base_url", "").strip() or "https://api.deepseek.com"
    raw["vision_model"] = fm.get("vision_model", "").strip()
    for k in ("inbox_dir", "library_dir", "index_xlsx"):
        if fm.get(k, "").strip():
            raw[k] = fm[k].strip()
    raw["after_ingest"] = "copy" if fm.get("after_ingest") == "copy" else "move"
    cats = _text_to_cats(fm.get("categories", ""))
    if cats:
        raw["categories"] = cats
    raw["default_category"] = fm.get("default_category", "").strip() or "其他"
    if raw["default_category"] not in raw["categories"]:
        raw["categories"][raw["default_category"]] = []
    raw["allow_new_category"] = bool(fm.get("allow_new_category"))
    raw["web_auto_scan"] = bool(fm.get("web_auto_scan"))
    try:
        raw["scan_interval_seconds"] = max(10, int(fm.get("scan_interval_seconds") or 60))
    except ValueError:
        pass
    try:
        raw["trash_keep_days"] = max(0, int(fm.get("trash_keep_days") or 0))
    except ValueError:
        pass
    core.save_raw_config(raw)

    msgs, ok = ["设置已保存"], True
    if winsys.IS_WIN:
        try:
            if bool(fm.get("autostart")) != winsys.get_autostart():
                winsys.set_autostart(bool(fm.get("autostart")))
            want_daily, at = bool(fm.get("daily")), fm.get("daily_at") or "08:00"
            cur = winsys.get_daily_task()
            if want_daily and cur != at:
                winsys.set_daily_task(True, at)
            elif not want_daily and cur:
                winsys.set_daily_task(False)
        except Exception as e:
            msgs.append(f"自动运行设置失败：{e}")
            ok = False
    if fm.get("action") == "test":
        try:
            core.DeepSeek(C()).chat_json("只输出 JSON。", '返回 {"ok": true}', max_tokens=20, retries=1)
            msgs.append("DeepSeek 连接成功 ✓")
        except Exception as e:
            msgs.append(f"DeepSeek 连接失败：{e}")
            ok = False
    return render_settings("；".join(msgs), ok)


@app.post("/reanalyze-missing")
def reanalyze_missing():
    cfg = C()
    ai = core.DeepSeek(cfg)
    if not ai.enabled:
        return redirect(url_for("settings", msg="请先填写 DeepSeek API Key"))

    def job():
        with _scan_lock:
            con = core.connect(cfg)
            try:
                for (fid,) in con.execute("SELECT id FROM files WHERE ai_used=0 AND deleted_at IS NULL").fetchall():
                    try:
                        core.reanalyze(cfg, con, ai, fid)
                    except Exception as e:
                        core.log.error("#%s 重新归档失败：%s", fid, e)
                core.export_index(cfg, con)
            finally:
                con.close()

    threading.Thread(target=job, daemon=True).start()
    return redirect(url_for("index", msg="已在后台开始重新归档，稍后刷新查看"))


@app.post("/quit")
def quit_app():
    fn = _shutdown_hook["fn"]
    threading.Timer(0.5, fn if fn else (lambda: os._exit(0))).start()
    return "<p style='font-family:sans-serif;padding:40px'>程序已退出，可以关闭此窗口。</p>"


if __name__ == "__main__":
    cfg = C()
    threading.Thread(target=scanner_loop, daemon=True).start()
    url = f"http://{cfg.get('web_host', '127.0.0.1')}:{cfg.get('web_port', 8765)}"
    if "--no-browser" not in sys.argv:
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    core.log.info("网页已启动：%s", url)
    app.run(host=cfg.get("web_host", "127.0.0.1"), port=cfg.get("web_port", 8765), debug=False, threaded=True)
