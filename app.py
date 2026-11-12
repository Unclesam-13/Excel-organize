"""
会计工作台 - 命令行入口

  python app.py scan                         处理收件箱一次（适合每日计划任务）
  python app.py watch                        持续监控收件箱
  python app.py ask "上周华东区的销售日报"     AI 自然语言找文件（不带问题则进入对话模式）
  python app.py search -k 发票 -c 财务 --from 2026-11-01 --to 2026-11-10
  python app.py open 12                      用 Excel 打开编号 12 的文件
  python app.py export                       重新生成 文件库总目录.xlsx
  python app.py stats                        统计
  python app.py reanalyze --all              用 AI 重新归档已入库文件
  python app.py chat 3 5 8                   选中 #3 #5 #8 文件，向 AI 提问表格内容
  python app.py delete 3 5                   删除（移到回收站）；restore 3 恢复；trash 查看回收站；purge 3 / purge --all 彻底删除
  python app.py setkey                       设置 DeepSeek API Key
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

import core

if os.name == "nt":  # 让 Windows 控制台正常显示中文
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def open_path(path: str, select: bool = False) -> None:
    p = Path(path)
    if not p.exists():
        print(f"文件不存在：{p}")
        return
    import winsys
    winsys.open_in_explorer(p, select=select)


def print_rows(rows: list[dict], with_reason: bool = False) -> None:
    if not rows:
        print("（没有结果）")
        return
    for r in rows:
        print(f"\n#{r['id']:<5} {r['data_date']}  [{r['category']}]  {r['title']}")
        if with_reason and r.get("reason"):
            print(f"       ★ {r['reason']}")
        if r.get("summary"):
            print(f"       {r['summary'][:120]}")
        print(f"       {r['stored_path']}")


def cmd_scan(cfg, args):
    res = core.scan_inbox(cfg)
    print(f"\n完成：新入库 {res['ok']}，重复 {res['duplicate']}，失败 {res['failed']}，待处理 {res['waiting']}")


def cmd_watch(cfg, args):
    interval = args.interval or cfg.get("scan_interval_seconds", 60)
    core.log.info("开始监控 %s，每 %s 秒检查一次（Ctrl+C 退出）", cfg["inbox_dir"], interval)
    ai = core.DeepSeek(cfg)
    try:
        while True:
            try:
                core.scan_inbox(cfg, ai)
            except Exception:
                core.log.exception("扫描出错")
            time.sleep(interval)
    except KeyboardInterrupt:
        print("已停止")


def _ask_once(cfg, con, ai, q):
    res = core.ask(con, q, cfg, ai)
    f = res["filters"]
    cond = []
    if f["date_from"] or f["date_to"]:
        cond.append(f"日期 {f['date_from'] or '…'} ~ {f['date_to'] or '…'}")
    if f["categories"]:
        cond.append("类别 " + "/".join(f["categories"]))
    if f["keywords"]:
        cond.append("关键词 " + "、".join(f["keywords"]))
    print(f"\n理解为：{'；'.join(cond) or '不限条件'}  {f.get('explain') or ''}")
    print(f"AI：{res['answer']}")
    print_rows(res["results"], with_reason=True)
    return res["results"]


def cmd_ask(cfg, args):
    ai = core.DeepSeek(cfg)
    if not ai.enabled:
        print("提示：未设置 DeepSeek API Key，使用简易关键词解析（运行 app.py setkey 设置）")
    con = core.connect(cfg)
    q = " ".join(args.question).strip()
    if q:
        _ask_once(cfg, con, ai, q)
        return
    print("AI 找文件 —— 直接用中文描述你要找的表格；输入编号打开文件，'f 编号' 打开所在文件夹，q 退出")
    last: list[dict] = []
    while True:
        try:
            q = input("\n你要找什么> ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if not q:
            continue
        if q.lower() in ("q", "quit", "exit", "退出"):
            break
        if q.isdigit() or (q[:2].lower() == "f " and q[2:].strip().isdigit()):
            fid = int(q.split()[-1])
            row = con.execute("SELECT stored_path FROM files WHERE id=?", (fid,)).fetchone()
            if row:
                open_path(row[0], select=q.lower().startswith("f"))
            else:
                print("没有这个编号")
            continue
        last = _ask_once(cfg, con, ai, q)


def cmd_search(cfg, args):
    con = core.connect(cfg)
    rows = core.search(con, args.keyword or [], args.category or [], args.date_from, args.date_to,
                       match="any" if args.any else "all", limit=args.limit)
    print_rows(rows)
    print(f"\n共 {len(rows)} 条")


def cmd_open(cfg, args):
    con = core.connect(cfg)
    row = con.execute("SELECT stored_path FROM files WHERE id=?", (args.id,)).fetchone()
    if not row:
        print("没有这个编号")
        return
    open_path(row[0], select=args.folder)


def cmd_export(cfg, args):
    print("总目录已生成：", core.export_index(cfg))


def cmd_stats(cfg, args):
    s = core.stats(core.connect(cfg))
    print(f"共 {s['total']} 份文件\n\n按类别：")
    for c, n in s["by_category"]:
        print(f"  {c:<8} {n}")
    print("\n按月份：")
    for m, n in s["by_month"][:24]:
        print(f"  {m}  {n}")


def cmd_reanalyze(cfg, args):
    ai = core.DeepSeek(cfg)
    if not ai.enabled:
        print("请先设置 DeepSeek API Key（app.py setkey）")
        return
    con = core.connect(cfg)
    ids = [r[0] for r in con.execute(
        "SELECT id FROM files WHERE deleted_at IS NULL AND (ai_used=0 OR ?)", (int(args.all),))] \
        if (args.all or args.missing) else args.ids
    for fid in ids:
        try:
            core.reanalyze(cfg, con, ai, int(fid))
        except Exception as e:
            core.log.error("#%s 重新归档失败：%s", fid, e)
    core.export_index(cfg, con)


def cmd_chat(cfg, args):
    import chat
    ai = core.DeepSeek(cfg)
    con = core.connect(cfg)
    names = [f"#{r['id']} {r['title']}" for i in args.ids
             for r in [con.execute("SELECT id, title FROM files WHERE id=?", (i,)).fetchone()] if r]
    if not names:
        print("没有找到这些编号的文件（可先用 search / ask 查编号）")
        return
    print("已选：" + "、".join(names) + "\n直接输入问题，q 退出，new 开始新对话")
    history: list[dict] = []
    while True:
        try:
            q = input("\n问> ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if q.lower() in ("q", "quit", "exit", "退出"):
            break
        if q.lower() == "new":
            history = []
            continue
        if not q:
            continue
        try:
            res = chat.ask(con, cfg, ai, args.ids, q, history)
        except Exception as e:
            print(f"出错：{e}")
            continue
        for st in res["steps"]:
            print(f"  🔍 {st}")
        print("\n" + res["answer"])
        history += [{"role": "user", "content": q}, {"role": "assistant", "content": res["answer"]}]


def cmd_delete(cfg, args):
    print(f"已移到回收站 {core.trash_files(cfg, core.connect(cfg), args.ids)} 个")


def cmd_restore(cfg, args):
    print(f"已恢复 {core.restore_files(cfg, core.connect(cfg), args.ids)} 个")


def cmd_trash(cfg, args):
    rows = core.list_trash(core.connect(cfg))
    for r in rows:
        print(f"#{r['id']:<5} 删除于 {r['deleted_at']}  [{r['category']}] {r['title']}")
    print(f"\n回收站共 {len(rows)} 个")


def cmd_purge(cfg, args):
    if not args.all and not args.ids:
        print("请指定编号，或用 --all 清空回收站")
        return
    print(f"已彻底删除 {core.purge_files(cfg, core.connect(cfg), None if args.all else args.ids)} 个")


def cmd_setkey(cfg, args):
    key = args.key or input("请粘贴 DeepSeek API Key（sk- 开头）：").strip()
    if not key:
        print("未输入，已取消")
        return
    core.save_api_key(key)
    ok = core.DeepSeek(core.load_config())
    try:
        ok.chat_json("只输出 JSON。", '返回 {"ok": true}', max_tokens=20)
        print("已保存，并测试连接成功 ✓")
    except Exception as e:
        print(f"已保存，但测试调用失败：{e}")


def main(argv=None):
    p = argparse.ArgumentParser(description="会计工作台")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("scan", help="处理收件箱一次")
    w = sub.add_parser("watch", help="持续监控收件箱")
    w.add_argument("--interval", type=int)
    a = sub.add_parser("ask", help="AI 自然语言找文件")
    a.add_argument("question", nargs="*")
    s = sub.add_parser("search", help="按条件检索")
    s.add_argument("-k", "--keyword", action="append", help="关键词，可多次使用")
    s.add_argument("-c", "--category", action="append", help="类别，可多次使用")
    s.add_argument("--from", dest="date_from")
    s.add_argument("--to", dest="date_to")
    s.add_argument("--any", action="store_true", help="多个关键词任一命中即可（默认需全部命中）")
    s.add_argument("--limit", type=int, default=50)
    o = sub.add_parser("open", help="打开文件")
    o.add_argument("id", type=int)
    o.add_argument("-f", "--folder", action="store_true", help="打开所在文件夹")
    sub.add_parser("export", help="生成总目录 Excel")
    sub.add_parser("stats", help="统计")
    r = sub.add_parser("reanalyze", help="用 AI 重新归档")
    r.add_argument("ids", nargs="*", type=int)
    r.add_argument("--all", action="store_true", help="全部重新归档")
    r.add_argument("--missing", action="store_true", help="只处理当初没用上 AI 的文件")
    ch = sub.add_parser("chat", help="选中文件向 AI 提问内容")
    ch.add_argument("ids", nargs="+", type=int)
    for name, hlp in (("delete", "删除（移到回收站）"), ("restore", "从回收站恢复")):
        x = sub.add_parser(name, help=hlp)
        x.add_argument("ids", nargs="+", type=int)
    sub.add_parser("trash", help="查看回收站")
    pg = sub.add_parser("purge", help="彻底删除回收站中的文件")
    pg.add_argument("ids", nargs="*", type=int)
    pg.add_argument("--all", action="store_true")
    k = sub.add_parser("setkey", help="设置 DeepSeek API Key")
    k.add_argument("key", nargs="?")

    args = p.parse_args(argv)
    cfg = core.load_config()
    core.setup_logging(cfg)
    globals()[f"cmd_{args.cmd}"](cfg, args)


if __name__ == "__main__":
    main()
