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
import uuid
from collections import OrderedDict
from pathlib import Path

import pandas as pd

import core
import excel_writer

MAX_FILES = 10
MAX_IMAGES = 10
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
        if sum(1 for v in sample if core.parse_date(v)) / len(sample) >= 0.7:
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


class FileSet(dict):
    """已加载的文件；AI 引用未选中的文件编号时（如通过 search_library 找到的）按需加载。"""

    def __init__(self, con, cfg, initial=None):
        super().__init__(initial or {})
        self.con, self.cfg = con, cfg
        self.proposals: list[dict] = []

    def __missing__(self, fid):
        loaded = load_files(self.con, [fid])
        if fid not in loaded:
            raise KeyError(fid)
        self[fid] = loaded[fid]
        return self[fid]


# ----------------------------------------------------------------------------- 给 AI 的文件说明
def _fmt(v) -> str:
    if isinstance(v, float):
        return f"{v:,.4f}".rstrip("0").rstrip(".")
    return f"{v:,}" if isinstance(v, int) else str(v)


def _csv(df: pd.DataFrame) -> str:
    return df.to_csv(index=False, float_format="%.10g")


def describe_files(files: dict, cfg: dict) -> tuple[str, bool]:
    if not files:
        return ("（用户没有选中文件。需要查数据或登记录入时，先用 search_library 在资料库里找合适的表格，"
                "再用 query_table / get_rows / propose_rows 操作。）"), True
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
        raise ToolError(f"没有文件#{file_id}（已选文件：{list(files)}；其他文件可先用 search_library 查编号）")
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


def search_library(files, keywords=None, category=None, limit=8):
    if isinstance(keywords, str):
        keywords = keywords.split()
    rows = core.search(files.con, keywords or [], [category] if category else [], match="any",
                       limit=max(1, min(int(limit or 8), 20)))
    return {"files": [{"file_id": r["id"], "title": r["title"], "category": r["category"], "data_date": r["data_date"],
                       "original_name": r["original_name"], "summary": (r["summary"] or "")[:120],
                       "sheets": [{"name": s_["name"], "rows": s_["rows"], "columns": s_["header"]} for s_ in r["sheets"]]}
                      for r in rows],
            "note": "可直接用这些 file_id 调用 query_table / get_rows / propose_rows"}


def propose_rows(files, rows, file_id=None, sheet=None, new_file_name=None, columns=None, reason=""):
    """生成"写入预览"，不直接写文件；用户在界面确认后才写入。"""
    if isinstance(rows, dict):
        rows = [rows]
    if not isinstance(rows, list) or not rows or not all(isinstance(r, dict) for r in rows):
        raise ToolError("rows 需要是对象数组，如 [{\"日期\": \"2026-11-10\", \"金额\": 100}]")
    rows = [{str(k).strip(): v for k, v in r.items() if str(k).strip()} for r in rows[:200]]
    keys = list(dict.fromkeys(k for r in rows for k in r))
    prop = {"id": uuid.uuid4().hex[:10], "reason": str(reason or "")[:300]}
    if file_id not in (None, "", 0):
        f = _get_file(files, file_id)
        path = Path(f["stored_path"])
        try:
            sheet_name, headers = excel_writer.sheet_headers(path, sheet)
        except excel_writer.WriteError as e:
            raise ToolError(str(e))
        mapping, unknown = excel_writer.map_columns(headers, keys)
        if unknown:
            raise ToolError(f"目标表没有这些列：{unknown}。可用列：{headers}。请改用这些列名，"
                            f"或把信息合并到相近的列（如“备注”/“摘要”）。")
        prop.update({"file_id": int(file_id), "file_title": f["title"], "sheet": sheet_name, "columns": headers,
                     "rows": [{mapping[k]: v for k, v in r.items()} for r in rows], "new_file_name": None})
        target = f"文件#{file_id}「{sheet_name}」"
    else:
        name = core._sanitize_name(new_file_name or "", 60)
        if not new_file_name or not name:
            raise ToolError("需要 file_id；如果资料库里确实没有合适的表格，请提供 new_file_name 和 columns 新建表格")
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        cols = [str(c).strip() for c in (columns or keys) if str(c).strip()]
        prop.update({"file_id": None, "file_title": name, "sheet": "Sheet1", "columns": cols,
                     "rows": [{c: r.get(c) for c in cols} for r in rows], "new_file_name": name})
        target = f"新建表格「{name}」"
    files.proposals.append(prop)
    return {"ok": True, "proposal_id": prop["id"], "target": target, "rows": len(rows),
            "note": "已生成写入预览，尚未写入。请在回答里简要说明写到哪里、写了什么，提醒用户核对后点击“写入 Excel”。"}


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
TOOLS += [
    {"type": "function", "function": {
        "name": "search_library",
        "description": "在整个资料库里按关键词/类别查找 Excel 文件（含工作表和列名）。用户没选文件、或需要找到应该登记到哪个表时使用。",
        "parameters": {"type": "object", "properties": {
            "keywords": {"type": "array", "items": {"type": "string"}, "description": "如 [\"发票\", \"登记\"]"},
            "category": {"type": "string"}, "limit": {"type": "integer"}}}}},
    {"type": "function", "function": {
        "name": "propose_rows",
        "description": "把要录入/登记的数据生成“写入预览”（追加到目标表末尾），由用户确认后写入。"
                       "列名必须用目标表的表头；金额写数字，日期写 YYYY-MM-DD。资料库里确实没有合适表格时，"
                       "用 new_file_name + columns 新建表格。",
        "parameters": {"type": "object", "properties": {
            "file_id": {"type": "integer", "description": "目标文件编号"},
            "sheet": {"type": "string", "description": "目标工作表"},
            "rows": {"type": "array", "items": {"type": "object"}, "description": "要追加的行，每行是 {列名: 值}"},
            "new_file_name": {"type": "string", "description": "新建表格时的文件名"},
            "columns": {"type": "array", "items": {"type": "string"}, "description": "新建表格时的列名"},
            "reason": {"type": "string", "description": "为什么写到这里"}},
            "required": ["rows"]}}},
]
DISPATCH = {"query_table": query_table, "search_text": search_text, "get_rows": get_rows,
            "search_library": search_library, "propose_rows": propose_rows}


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
    elif name == "search_library":
        kw = a.get("keywords") or []
        desc = "在资料库查找 " + ("、".join(kw) if isinstance(kw, list) else str(kw)) + (f"（{a['category']}）" if a.get("category") else "")
        if "files" in result:
            return f"{desc} → 找到 {len(result['files'])} 个文件"
    elif name == "propose_rows":
        desc = "生成写入预览"
        if "target" in result:
            return f"{desc} → {result['target']} {result['rows']} 行（待确认）"
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
SYSTEM = """你是会计工作台里的数据助手，做事严谨。用户选中了 {n} 个 Excel 文件（见下方“文件内容”），可能就文件内容提问，也可能发图片让你识别并登记到表格。今天是 {today}。

回答规则：
1. 只依据这些文件（以及用工具查到）的数据回答；数据里没有的信息，直接说明“文件中没有”，不要编造或猜测。
2. 求和、平均、计数、最大最小、排名、筛选、分组、跨文件对比等计算，一律调用 query_table 工具获得精确结果，不要心算或估算。{partial_rule}
3. 表格里可能有“合计/小计/总计”行，汇总时要先用筛选排除，避免重复计算；数值列里的空值不参与计算。
4. 用中文回答，先给结论，再给依据；需要列明细或对比时使用 Markdown 表格。
5. 关键数字后注明出处，例如（文件#12「明细」）。
6. 工具返回 error 时，根据提示修正参数后重试。
7. 用户附带图片时，消息里会给出图片识别结果。先把识别出的关键内容清楚地写在回答里，方便用户核对。
8. 用户要把图片或文字里的数据“录入/登记/记账/放到表格里”，或图片明显是要入账的单据（发票、收据、回单、报销单等）时：
   先确定目标表格：优先用已选文件，没有合适的就用 search_library 在资料库里找。对照目标表的列名整理数据，然后调用 propose_rows 生成写入预览。
   不要说“已写入”，实际写入要等用户在界面上确认。资料库里确实没有合适的表格时，才用 new_file_name 新建。
   识别结果里标了 [看不清] 的字段要在回答中提醒用户补充。

=== 文件内容 ===
{files}"""


VISION_PROMPT = """请仔细识别这 {n} 张图片，完整、准确地转写其中的内容，供后续核对和录入表格使用：
- 先说明每张图片是什么（如增值税发票、收据、银行回单、报销单、出入库单、表格截图、聊天截图等）
- 逐项列出关键字段：日期、单号/发票号码、开票方/收款方、购买方/付款方、项目名称、规格型号、数量、单价、金额、税率、税额、价税合计、备注等（图片里有什么写什么）
- 图片中的表格转成 Markdown 表格，逐行逐列照抄，不要遗漏行
- 金额写成阿拉伯数字（有大写金额时也照抄原文），日期写成 YYYY-MM-DD
- 看不清或被遮挡的内容标注 [看不清]，不要猜测
多张图片请用“图片1/图片2…”分开。用户的问题（仅供参考）：{q}"""


def recognize_images(cfg: dict, ai: core.DeepSeek, images: list[str], question: str) -> str:
    imgs = [u for u in images if isinstance(u, str) and u.startswith("data:image/")][:MAX_IMAGES]
    if not imgs:
        return ""
    content = [{"type": "text", "text": VISION_PROMPT.format(n=len(imgs), q=question or "（无）")}]
    content += [{"type": "image_url", "image_url": {"url": u, "detail": "original"}} for u in imgs]
    model = (cfg.get("vision_model") or "").strip() or None
    text = ""
    for attempt in (1, 2):
        try:
            # 识图不需要“思考模式”：开着时思考过程可能占满输出上限，导致识别结果为空
            msg = ai.chat([{"role": "user", "content": content}], max_tokens=8000 * attempt, model=model, thinking=False)
        except RuntimeError as e:
            if "image" in str(e).lower() or "400" in str(e):
                raise RuntimeError(f"当前模型不支持图片识别：{e}。请在“设置”里把模型改为 deepseek-flash（或在“识图模型”里填 deepseek-flash）。")
            raise
        text = (msg.get("content") or "").strip()
        if text:
            break
        core.log.warning("图片识别第 %d 次返回空内容（思考内容 %d 字）", attempt, len(msg.get("reasoning_content") or ""))
    if not text:
        raise RuntimeError("DeepSeek 没有返回图片识别结果。请稍后重试；如果一直这样，请在“设置”里确认模型为 deepseek-flash，"
                           "并检查图片是否清晰、是否为 JPG/PNG/GIF/WebP 格式。")
    return text


def ask(con, cfg: dict, ai: core.DeepSeek, file_ids, question: str, history=None, max_rounds: int = 10,
        images=None) -> dict:
    if not ai.enabled:
        raise RuntimeError("还没有设置 DeepSeek API Key，请先到“设置”里填写。")
    files = FileSet(con, cfg, load_files(con, file_ids))
    steps = []
    question = (question or "").strip()
    if images:
        n = len([u for u in images if isinstance(u, str) and u.startswith("data:image/")][:MAX_IMAGES])
        text = recognize_images(cfg, ai, images, question)
        steps.append(f"识别图片 {n} 张 → {len(text)} 字")
        question = (question or "请识别图片内容；如果是需要入账/登记的单据，帮我整理并放到对应的表格里。") + \
            f"\n\n【以下是用户附带的 {n} 张图片的识别结果】\n{text}"
    file_text, full = describe_files(files, cfg)
    system = SYSTEM.format(
        n=len(files), today=dt.date.today().isoformat(), files=file_text,
        partial_rule="" if full else "\n   注意：部分表格只给出了前几行样例，涉及具体记录、明细、统计时必须用工具查询全部数据。")

    msgs = [{"role": "system", "content": system}]
    for h in (history or [])[-12:]:
        if h.get("role") in ("user", "assistant") and str(h.get("content", "")).strip():
            m = {"role": h["role"], "content": str(h["content"])[:8000]}
            if h["role"] == "assistant" and ai.thinking:
                # 深度思考模式下带工具的多轮对话，必须回传之前每一轮的思考内容
                m["reasoning_content"] = str(h.get("reasoning_content") or "")
            msgs.append(m)
    msgs.append({"role": "user", "content": question})

    last_reasoning = {"text": ""}

    def finish(answer: str) -> dict:
        return {"answer": answer or "（AI 没有返回内容）", "steps": steps, "full_data": full, "files": list(files),
                "proposals": files.proposals, "user_content": question, "reasoning": last_reasoning["text"]}

    for _ in range(max_rounds):
        msg = ai.chat(msgs, tools=TOOLS)
        last_reasoning["text"] = msg.get("reasoning_content") or ""
        calls = msg.get("tool_calls") or []
        if not calls:
            return finish((msg.get("content") or "").strip())
        assistant = {"role": "assistant", "content": msg.get("content") or "", "tool_calls": calls}
        if ai.thinking or msg.get("reasoning_content"):  # 思考模式下，同一轮工具调用需要回传思考内容
            assistant["reasoning_content"] = msg.get("reasoning_content") or ""
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
    last_reasoning["text"] = msg.get("reasoning_content") or ""
    return finish((msg.get("content") or "").strip())
