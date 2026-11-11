"""
读取“没有缓存计算结果”的 xlsx（如被 openpyxl 写入过的文件）时，自行计算常见公式，
使 AI 查询、档案统计能读到 价税合计 等公式列的值。

只支持安全的子集：数字、本表单元格/区域引用、+ - * / ^、括号、SUM/AVERAGE/MIN/MAX/COUNT/ROUND/ABS。
其他公式（跨表引用、IF、VLOOKUP 等）返回空值。不使用任何不受限制的 eval。
"""
from __future__ import annotations

import ast
import datetime as dt
import operator
import re
import zipfile
from pathlib import Path

_EMPTY_FORMULA = re.compile(rb"</f>(?:<v\s*/>|<v></v>)?</c>")


def needs_eval(path: Path) -> bool:
    """工作表 XML 中存在“有公式但没有计算结果”的单元格。"""
    try:
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                    if _EMPTY_FORMULA.search(z.read(name)):
                        return True
    except (zipfile.BadZipFile, OSError, KeyError):
        return False
    return False


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def _num_to_col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


_RANGE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)")
_CELL = re.compile(r'(?<![A-Za-z_"\d])\$?([A-Z]{1,3})\$?(\d+)(?![\d(])')
_FUNCS = ("SUM", "AVERAGE", "MIN", "MAX", "COUNT", "ROUND", "ABS")
_BIN = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv,
        ast.Pow: operator.pow}
_UN = {ast.USub: operator.neg, ast.UAdd: operator.pos}


def _num(v):
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float)):
        return v
    if v is None or v == "":
        return 0
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        raise ValueError("非数值")


def _flat(args):
    for a in args:
        if isinstance(a, list):
            yield from a
        else:
            yield a


def _nums(args):
    return [v for v in _flat(args) if isinstance(v, (int, float)) and not isinstance(v, bool)]


_FN = {
    "SUM": lambda *a: sum(_nums(a)),
    "AVERAGE": lambda *a: (sum(_nums(a)) / len(_nums(a))) if _nums(a) else 0,
    "MIN": lambda *a: min(_nums(a)) if _nums(a) else 0,
    "MAX": lambda *a: max(_nums(a)) if _nums(a) else 0,
    "COUNT": lambda *a: len(_nums(a)),
    "ROUND": lambda x, n=0: round(_num(x), int(_num(n))),
    "ABS": lambda x: abs(_num(x)),
}


def evaluate(formula: str, get):
    """formula 不含开头的 '='；get(col_letters, row) 返回单元格值。无法计算时返回 None。"""
    expr = formula.strip()
    if not expr or "!" in expr or '"' in expr or len(expr) > 500:
        return None
    ranges: list[tuple] = []

    def sub_range(m):
        ranges.append((m[1], int(m[2]), m[3], int(m[4])))
        return f"__R{len(ranges) - 1}"

    expr = _RANGE.sub(sub_range, expr)
    cells: list[tuple] = []

    def sub_cell(m):
        cells.append((m[1], int(m[2])))
        return f"__C{len(cells) - 1}"

    expr = _CELL.sub(sub_cell, expr).replace("^", "**")
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError:
        return None

    def ev(node):
        if isinstance(node, ast.Expression):
            return ev(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return node.value
        if isinstance(node, ast.Name):
            if node.id.startswith("__C"):
                c, r = cells[int(node.id[3:])]
                return get(c, r)
            if node.id.startswith("__R"):
                c1, r1, c2, r2 = ranges[int(node.id[3:])]
                out = []
                for rr in range(min(r1, r2), max(r1, r2) + 1):
                    for cc in range(min(_col_to_num(c1), _col_to_num(c2)), max(_col_to_num(c1), _col_to_num(c2)) + 1):
                        out.append(get(_num_to_col(cc), rr))
                return out
            raise ValueError("未知名称")
        if isinstance(node, ast.BinOp) and type(node.op) in _BIN:
            return _BIN[type(node.op)](_num(ev(node.left)), _num(ev(node.right)))
        if isinstance(node, ast.UnaryOp) and type(node.op) in _UN:
            return _UN[type(node.op)](_num(ev(node.operand)))
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id.upper() in _FN \
                and not node.keywords:
            return _FN[node.func.id.upper()](*[ev(a) for a in node.args])
        raise ValueError("不支持的公式")

    try:
        v = ev(tree)
    except (ValueError, ZeroDivisionError, TypeError, OverflowError, RecursionError):
        return None
    if isinstance(v, float):
        v = round(v, 10)
        if v.is_integer() and abs(v) < 1e15:
            v = int(v)
    return v if isinstance(v, (int, float)) else None


def _to_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, float):
        v = round(v, 10)
        return str(int(v)) if v.is_integer() and abs(v) < 1e15 else str(v)
    return str(v)


def read_with_formulas(path: Path):
    """返回 {工作表名: DataFrame(全部为字符串、无表头)}，公式缺少结果时自行计算。"""
    import pandas as pd
    from openpyxl import load_workbook

    wb_f = load_workbook(path, data_only=False)
    wb_v = load_workbook(path, data_only=True)
    frames = {}
    for ws_f in wb_f.worksheets:
        ws_v = wb_v[ws_f.title]
        memo: dict = {}
        max_r, max_c = ws_f.max_row, ws_f.max_column

        def get(col, row, _wf=ws_f, _wv=ws_v, _memo=memo):
            key = (col, row)
            if key in _memo:
                return _memo[key]
            _memo[key] = None  # 循环引用保护
            if row < 1 or row > _wf.max_row:
                return None
            v = _wv[f"{col}{row}"].value
            f = _wf[f"{col}{row}"].value
            if v is None and isinstance(f, str) and f.startswith("="):
                v = evaluate(f[1:], get)
            _memo[key] = v
            return v

        rows = [[_to_str(get(_num_to_col(c), r)) for c in range(1, max_c + 1)] for r in range(1, max_r + 1)]
        frames[ws_f.title] = pd.DataFrame(rows, dtype=str) if rows else pd.DataFrame()
    wb_f.close()
    wb_v.close()
    return frames
