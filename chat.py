"""
选中若干 Excel 后向 AI 提问。

做法：
- 把每个文件的结构（工作表、列名、类型、数值范围/合计）交给 DeepSeek；
  数据量小时附上完整数据，数据量大时只附前几行样例。
- 同时提供查询工具（筛选 / 分组汇总 / 排序 / 搜索 / 读取行），由本机 pandas 精确计算，
  AI 只负责决定"查什么"和组织答案，避免心算出错，也不需要把大表全部发给 AI。
- 工具只做固定的几种数据操作，不执行 AI 生成的代码。
"""
from __future__ import annotations

import datetime as dt
import inspect
import json
import threading
from collections import OrderedDict
from pathlib import Path

import pandas as pd

import core

MAX_FILES = 10
_cache: "OrderedDict[tuple, list]" = OrderedDict()
_cache_lock = threading.Lock()

AGG_FUNCS = {"sum": "合计", "mean": "平均", "count": "计数", "min": "最小", "max": "最大",
             "median": "中位数", "nunique": "去重数"}
NUMERIC_ONLY = {"sum", "mean", "median"}


# ----------------------------------------------------------------------------- 读表
def _dedupe(names) -> list[str]:
    out, seen = [], {}
    for i, n in enumerate(names):
        n = str(n).strip() or f"列{i + 1}"
        if n in seen:
            seen[n] += 1
            n = f"{n}_{seen[n]}"
        else:
            seen[n] = 0
        out.append(n)
    return out


def _infer_types(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    kinds = {}
    for c in df.columns:
        s = df[c].astype(str)
        ne = s[s != ""]
        if ne.empty:
            kinds[c] = "空"
            continue
        cleaned = ne.str.replace(r"[,¥￥$\s]", "", regex=True)
        has_leading_zero = cleaned.str.match(r"^0\d").mean() > 0.3  # 编号类(00123)保持文本
        num = pd.to_numeric(cleaned.str.rstrip("%"), errors="coerce")
        if not has_leading_zero and num.notna().mean() >= 0.9:
            df[c] = pd.to_numeric(s.str.replace(r"[,¥￥$\s]", "", regex=True).str.rstrip("%"), errors="coerce")
            kinds[c] = "数值(%)" if cleaned.str.endswith("%").mean() > 0.5 else "数值"
            continue
        sample = ne.head(200).tolist()
        if sum(1 for v in sample if core.parse_date(v)) / len(sample) >= 0.8:
            df[c] = s.map(lambda v: d.isoformat() if (d := core.parse_date(v)) else v)
            kinds[c] = "日期"
        else:
            kinds[c] = "文本"
    return df, kinds


def load_sheets(path: Path) -> list[dict]:
    key = (str(path), path.stat().st_mtime)
    with _cache_lock:
        if key in _cache:
            _cache.move_to_end(key)
            return _cache[key]
    sheets = []
    for name, raw in core.read_tables(path).items():
        df = core._clean(raw)
        if df.empty:
            continue
        h = core._guess_header(df)
        body = df.iloc[h + 1:].reset_index(drop=True)
        body.columns = _dedupe(df.iloc[h].tolist())
        body, kinds = _infer_types(body)
        sheets.append({"name": str(name), "df": body, "kinds": kinds,
                       "above_header": " / ".join(v for r in df.iloc[:h].values.tolist() for v in r if v)[:200]})
    with _cache_lock:
        _cache[key] = sheets
        while len(_cache) > 20:
            _cache.popitem(last=False)
    return sheets


def load_files(con, file_ids) -> dict[int, dict]:
    files = {}
    for fid in list(dict.fromkeys(int(i) for i in file_ids))[:MAX_FILES]:
        row = con.execute("SELECT * FROM files WHERE id=? AND deleted_at IS NULL", (fid,)).fetchone()
        if not row:
            continue
        meta = core.row_to_dict(row)
        p = Path(meta["stored_path"])
        if not p.exists():
            raise FileNotFoundError(f"文件#{fid} 不存在或已被移动：{p}")
        meta["sheets_data"] = load_sheets(p)
        files[fid] = meta
    return files


# ----------------------------------------------------------------------------- 给 AI 的文件说明
def _fmt(v) -> str:
    if isinstance(v, float):
        return f"{v:,.4f}".rstrip("0").rstrip(".")
    return f"{v:,}" if isinstance(v, int) else str(v)


def _csv(df: pd.DataFrame) -> str:
    return df.to_csv(index=False, float_format="%.10g")


def describe_files(files: dict, cfg: dict) -> tuple[str, bool]:
    full_budget = int(cfg.get("chat_full_data_chars", 40000))
    preview = int(cfg.get("chat_preview_rows", 30))
    total = sum(len(_csv(sh["df"])) for f in files.values() for sh in f["sheets_data"])
    full = total <= full_budget

    parts = []
    for fid, f in files.items():
        lines = [f"【文件#{fid}】{f['title']}｜原文件名：{f['original_name']}｜业务日期：{f['data_date']}｜类别：{f['category']}"]
        if f.get("summary"):
            lines.append(f"档案摘要：{f['summary']}")
        for sh in f["sheets_data"]:
            df, kinds = sh["df"], sh["kinds"]
            lines.append(f"  ◆ 工作表「{sh['name']}」：{len(df)} 行 × {len(df.columns)} 列"
                         + (f"（表头上方文字：{sh['above_header']}）" if sh["above_header"] else ""))
            for c in df.columns:
                k = kinds.get(c, "文本")
                if k.startswith("数值"):
                    s = df[c].dropna()
                    extra = f"最小 {_fmt(s.min())}，最大 {_fmt(s.max())}，合计 {_fmt(s.sum())}" if len(s) else "无数据"
                else:
                    vals = df[c].astype(str)
                    u = vals[vals != ""].unique()
                    extra = f"{len(u)} 种取值，例：{' / '.join(u[:5])}"[:120]
                lines.append(f"     - {c}［{k}］{extra}")
            data = df if full else df.head(preview)
            label = "完整数据" if full else f"前 {len(data)} 行（共 {len(df)} 行，其余需用工具查询）"
            lines.append(f"    {label}（CSV）：\n{_csv(data)}")
        parts.append("\n".join(lines))
    return "\n\n".join(parts), full


# ----------------------------------------------------------------------------- 查询工具
class ToolError(Exception):
    pass


def _get_file(files, file_id):
    try:
        f = files[int(file_id)]
    except (KeyError, ValueError, TypeError):
        raise ToolError(f"没有文件#{file_id}，可用：{list(files)}")
    return f


def _get_sheet(f, sheet):
    sheets = f["sheets_data"]
    if not sheets:
        raise ToolError("该文件没有数据")
    if not sheet:
        return sheets[0]
    for cond in (lambda s: s["name"] == sheet, lambda s: s["name"].lower() == str(sheet).lower(),
                 lambda s: str(sheet) in s["name"]):
        hit = [s for s in sheets if cond(s)]
        if hit:
            return hit[0]
    raise ToolError(f"没有工作表「{sheet}」，可用：{[s['name'] for s in sheets]}")


def _col(df: pd.DataFrame, name) -> str:
    name = str(name).strip()
    cols = [str(c) for c in df.columns]
    for cond in (lambda c: c == name, lambda c: c.strip().lower() == name.lower()):
        hit = [c for c in cols if cond(c)]
        if hit:
            return hit[0]
    hit = [c for c in cols if name in c or c in name]
    if len(hit) == 1:
        return hit[0]
    raise ToolError(f"没有列「{name}」，可用列：{cols}")


def _conv(v, kind):
    if kind.startswith("数值"):
        try:
            return float(str(v).replace(",", "").rstrip("%"))
        except ValueError:
            raise ToolError(f"“{v}”不是数字")
    if kind == "日期":
        d = core.parse_date(str(v))
        return d.isoformat() if d else str(v)
    return str(v).strip()


def _mask(s: pd.Series, op: str, value, kind: str) -> pd.Series:
    op = (op or "=").strip().lower()
    text = s.astype(str)
    if op in ("=", "==", "eq"):
        return (s == _conv(value, kind)) if kind.startswith("数值") else text.str.strip() == _conv(value, kind)
    if op in ("!=", "<>", "ne"):
        return ~_mask(s, "=", value, kind)
    if op in (">", ">=", "<", "<="):
        v = _conv(value, kind)
        t = s if kind.startswith("数值") else text
        return {">": t > v, ">=": t >= v, "<": t < v, "<=": t <= v}[op].fillna(False)
    if op == "between":
        if not isinstance(value, (list, tuple)) or len(value) != 2:
            raise ToolError("between 的 value 需要 [下限, 上限]")
        return _mask(s, ">=", value[0], kind) & _mask(s, "<=", value[1], kind)
    if op == "in":
        vals = value if isinstance(value, (list, tuple)) else [value]
        return s.isin([_conv(x, kind) for x in vals]) if kind.startswith("数值") \
            else text.str.strip().isin([_conv(x, kind) for x in vals])
    if op == "contains":
        return text.str.contains(str(value), case=False, regex=False, na=False)
    if op == "not_contains":
        return ~text.str.contains(str(value), case=False, regex=False, na=False)
    if op == "startswith":
        return text.str.startswith(str(value), na=False)
    if op == "is_empty":
        return s.isna() | (text.str.strip().isin(["", "nan"]))
    if op == "not_empty":
        return ~_mask(s, "is_empty", None, kind)
    raise ToolError(f"不支持的运算符 {op}")


def query_table(files, file_id, sheet=None, columns=None, filters=None, group_by=None,
                aggregations=None, sort_by=None, descending=False, limit=50):
    f = _get_file(files, file_id)
    sh = _get_sheet(f, sheet)
    df, kinds = sh["df"], sh["kinds"]
    mask = pd.Series(True, index=df.index)
    for flt in filters or []:
        c = _col(df, flt.get("column"))
        mask &= _mask(df[c], flt.get("op", "="), flt.get("value"), kinds.get(c, "文本"))
    d = df[mask]

    if aggregations:
        specs = {}
        for a in aggregations:
            func = str(a.get("func", "sum")).lower()
            if func not in AGG_FUNCS:
                raise ToolError(f"不支持的汇总函数 {func}，可用：{list(AGG_FUNCS)}")
            if func == "count" and not a.get("column"):
                specs["行数"] = (df.columns[0], "size")
                continue
            c = _col(df, a.get("column"))
            if func in NUMERIC_ONLY and not kinds.get(c, "").startswith("数值"):
                raise ToolError(f"列「{c}」不是数值列，不能{AGG_FUNCS[func]}")
            specs[f"{c}_{AGG_FUNCS[func]}"] = (c, func)
        if group_by:
            gb = [_col(df, g) for g in (group_by if isinstance(group_by, list) else [group_by])]
            out = d.groupby(gb, dropna=False).agg(**specs).reset_index()
        else:
            out = pd.DataFrame([{k: (len(d) if fn == "size" else d[c].agg(fn)) for k, (c, fn) in specs.items()}])
    elif group_by:
        gb = [_col(df, g) for g in (group_by if isinstance(group_by, list) else [group_by])]
        out = d.groupby(gb, dropna=False).size().reset_index(name="行数")
    else:
        out = d[[_col(df, c) for c in columns]] if columns else d

    if sort_by:
        sc = sort_by if sort_by in out.columns else _col(out, sort_by)
        out = out.sort_values(sc, ascending=not descending, kind="stable")
    limit = max(1, min(int(limit or 50), 200))
    return {"file_id": int(file_id), "sheet": sh["name"], "matched_rows": int(len(d)), "result_rows": int(len(out)),
            "truncated": len(out) > limit, "csv": _csv(out.head(limit))}


def search_text(files, keyword, file_ids=None, limit=30):
    kw = str(keyword).strip()
    if not kw:
        raise ToolError("keyword 不能为空")
    hits, total = [], 0
    for fid in (file_ids or list(files)):
        f = _get_file(files, fid)
        for sh in f["sheets_data"]:
            df = sh["df"]
            m = df.astype(str).apply(lambda col: col.str.contains(kw, case=False, regex=False, na=False)).any(axis=1)
            total += int(m.sum())
            for idx, row in df[m].head(max(0, int(limit) - len(hits))).iterrows():
                cells = "，".join(f"{c}={v}" for c, v in row.items() if str(v) not in ("", "nan"))
                hits.append(f"文件#{fid}「{sh['name']}」第{idx + 1}行：{cells}")
    return {"keyword": kw, "total_matches": total, "shown": len(hits), "rows": hits}


def get_rows(files, file_id, sheet=None, start=1, count=50):
    f = _get_file(files, file_id)
    sh = _get_sheet(f, sheet)
    start = max(1, int(start or 1))
    count = max(1, min(int(count or 50), 200))
    part = sh["df"].iloc[start - 1:start - 1 + count]
    return {"file_id": int(file_id), "sheet": sh["name"], "total_rows": int(len(sh["df"])),
            "from_row": start, "to_row": start + len(part) - 1, "csv": _csv(part)}


TOOLS = [
    {"type": "function", "function": {
        "name": "query_table",
        "description": "对某个文件的某张工作表做筛选、分组汇总、排序，返回精确结果(CSV)。所有计算类问题都用它。",
        "parameters": {"type": "object", "properties": {
            "file_id": {"type": "integer", "description": "文件编号，如 12"},
            "sheet": {"type": "string", "description": "工作表名；文件只有一张表时可省略"},
            "columns": {"type": "array", "items": {"type": "string"}, "description": "只返回这些列（不汇总时）"},
            "filters": {"type": "array", "description": "筛选条件，多个条件为“且”", "items": {
                "type": "object", "properties": {
                    "column": {"type": "string"},
                    "op": {"type": "string", "enum": ["=", "!=", ">", ">=", "<", "<=", "between", "in", "contains",
                                                     "not_contains", "startswith", "is_empty", "not_empty"]},
                    "value": {"description": "比较值；between 用 [下限, 上限]，in 用数组；日期写 YYYY-MM-DD，按月用 startswith 2026-11"}},
                "required": ["column", "op"]}},
            "group_by": {"type": "array", "items": {"type": "string"}, "description": "分组列"},
            "aggregations": {"type": "array", "description": "汇总，如 [{column:'金额', func:'sum'}]；统计行数用 {func:'count'}",
                             "items": {"type": "object", "properties": {
                                 "column": {"type": "string"},
                                 "func": {"type": "string", "enum": list(AGG_FUNCS)}}, "required": ["func"]}},
            "sort_by": {"type": "string", "description": "排序列（可用汇总后的列名，如 金额_合计）"},
            "descending": {"type": "boolean"},
            "limit": {"type": "integer", "description": "最多返回行数，默认 50，最大 200"}},
            "required": ["file_id"]}}},
    {"type": "function", "function": {
        "name": "search_text",
        "description": "在选中的文件里搜索包含某个词的行（客户名、产品、单号等），返回命中的行。",
        "parameters": {"type": "object", "properties": {
            "keyword": {"type": "string"},
            "file_ids": {"type": "array", "items": {"type": "integer"}, "description": "限定文件，省略则搜索全部"},
            "limit": {"type": "integer"}}, "required": ["keyword"]}}},
    {"type": "function", "function": {
        "name": "get_rows",
        "description": "按行号读取原始数据（从第 1 行数据开始计数，表头不算），每次最多 200 行。",
        "parameters": {"type": "object", "properties": {
            "file_id": {"type": "integer"}, "sheet": {"type": "string"},
            "start": {"type": "integer"}, "count": {"type": "integer"}}, "required": ["file_id"]}}},
]
DISPATCH = {"query_table": query_table, "search_text": search_text, "get_rows": get_rows}


def run_tool(files, name: str, args: dict) -> dict:
    fn = DISPATCH.get(name)
    if not fn:
        return {"error": f"未知工具 {name}"}
    allowed = set(inspect.signature(fn).parameters) - {"files"}
    try:
        return fn(files, **{k: v for k, v in (args or {}).items() if k in allowed})
    except ToolError as e:
        return {"error": str(e)}
    except Exception as e:  # 防御：任何意外错误都反馈给 AI 让它修正参数
        core.log.warning("工具 %s 出错：%s %s", name, e, args)
        return {"error": f"{type(e).__name__}: {e}"}


def describe_step(name: str, args: dict, result: dict) -> str:
    a = args or {}
    if name == "query_table":
        bits = [f"文件#{a.get('file_id')}" + (f"「{result.get('sheet') or a.get('sheet')}」" if result.get("sheet") or a.get("sheet") else "")]
        for flt in a.get("filters") or []:
            v = flt.get("value")
            bits.append(f"{flt.get('column')} {flt.get('op')} {json.dumps(v, ensure_ascii=False) if isinstance(v, list) else (v if v is not None else '')}".strip())
        if a.get("group_by"):
            bits.append("按 " + "、".join(a["group_by"] if isinstance(a["group_by"], list) else [a["group_by"]]) + " 分组")
        if a.get("aggregations"):
            bits.append("、".join(f"{x.get('column') or ''}{AGG_FUNCS.get(x.get('func'), x.get('func'))}" for x in a["aggregations"]))
        if a.get("sort_by"):
            bits.append(f"按 {a['sort_by']} {'降序' if a.get('descending') else '升序'}")
        desc = "查询 " + "；".join(bits)
    elif name == "search_text":
        desc = f"搜索“{a.get('keyword')}”"
    elif name == "get_rows":
        desc = f"读取 文件#{a.get('file_id')} 第 {a.get('start', 1)} 行起 {a.get('count', 50)} 行"
    else:
        desc = name
    if "error" in result:
        return f"{desc} → 出错：{result['error']}"
    if "result_rows" in result:
        return f"{desc} → 匹配 {result['matched_rows']} 行，结果 {result['result_rows']} 行"
    if "total_matches" in result:
        return f"{desc} → 命中 {result['total_matches']} 行"
    return f"{desc} → {result.get('to_row', 0) - result.get('from_row', 1) + 1} 行"


# ----------------------------------------------------------------------------- 对话
SYSTEM = """你是一名严谨的数据分析助手。用户选中了下面 {n} 个 Excel 文件，正在就其内容向你提问。今天是 {today}。

回答规则：
1. 只依据这些文件的数据回答；数据里没有的信息，直接说明“文件中没有”，不要编造或猜测。
2. 求和、平均、计数、最大最小、排名、筛选、分组、跨文件对比等计算，一律调用 query_table 工具获得精确结果，不要心算或估算。{partial_rule}
3. 表格里可能有“合计/小计/总计”行，汇总时要先用筛选排除，避免重复计算；数值列里的空值不参与计算。
4. 用中文回答，先给结论，再给依据；需要列明细或对比时使用 Markdown 表格。
5. 关键数字后注明出处，例如（文件#12「明细」）。
6. 工具返回 error 时，根据提示修正参数后重试。

=== 文件内容 ===
{files}"""


def ask(con, cfg: dict, ai: core.DeepSeek, file_ids, question: str, history=None, max_rounds: int = 8) -> dict:
    if not ai.enabled:
        raise RuntimeError("还没有设置 DeepSeek API Key，请先到“设置”里填写。")
    files = load_files(con, file_ids)
    if not files:
        raise RuntimeError("没有可用的文件，请先在文件库里勾选文件。")
    file_text, full = describe_files(files, cfg)
    system = SYSTEM.format(
        n=len(files), today=dt.date.today().isoformat(), files=file_text,
        partial_rule="" if full else "\n   注意：部分表格只给出了前几行样例，涉及具体记录、明细、统计时必须用工具查询全部数据。")

    msgs = [{"role": "system", "content": system}]
    for h in (history or [])[-12:]:
        if h.get("role") in ("user", "assistant") and str(h.get("content", "")).strip():
            msgs.append({"role": h["role"], "content": str(h["content"])[:8000]})
    msgs.append({"role": "user", "content": question})

    steps = []
    for _ in range(max_rounds):
        msg = ai.chat(msgs, tools=TOOLS)
        calls = msg.get("tool_calls") or []
        if not calls:
            return {"answer": (msg.get("content") or "").strip() or "（AI 没有返回内容）", "steps": steps,
                    "full_data": full, "files": list(files)}
        assistant = {"role": "assistant", "content": msg.get("content") or "", "tool_calls": calls}
        if msg.get("reasoning_content"):  # 思考模式下，同一轮工具调用需要回传思考内容
            assistant["reasoning_content"] = msg["reasoning_content"]
        msgs.append(assistant)
        for call in calls:
            fn = call.get("function") or {}
            try:
                args = json.loads(fn.get("arguments") or "{}")
            except json.JSONDecodeError:
                args, result = {}, {"error": "参数不是合法 JSON"}
            else:
                result = run_tool(files, fn.get("name", ""), args)
            steps.append(describe_step(fn.get("name", ""), args, result))
            msgs.append({"role": "tool", "tool_call_id": call.get("id", ""),
                         "content": json.dumps(result, ensure_ascii=False, default=str)[:30000]})

    msgs.append({"role": "user", "content": "查询次数已达上限，请根据以上已经得到的结果直接给出最终回答，并说明还缺哪些信息。"})
    msg = ai.chat(msgs, tools=TOOLS, tool_choice="none")
    return {"answer": (msg.get("content") or "").strip(), "steps": steps, "full_data": full, "files": list(files)}
