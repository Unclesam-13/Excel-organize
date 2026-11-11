"""
把行数据写入已有的 Excel / CSV（追加到表格末尾），或新建一个 Excel。

- 自动识别表头所在行（表头上方有标题也可以），按表头名称对应列
- 末尾有“合计/总计/小计”行时插到它前面，并把合计行里的 SUM 等公式范围扩展到新行
- 新行沿用上一行的格式（字体、边框、数字格式、对齐）
- 写入前自动备份到 图书馆\\_备份\\，可撤销
"""
from __future__ import annotations

import csv
import datetime as dt
import re
import shutil
from copy import copy
from pathlib import Path

import core

BACKUP_DIR = "_备份"
TOTAL_WORDS = ("合计", "总计", "小计", "共计", "Total", "TOTAL")


class WriteError(Exception):
    pass


# ----------------------------------------------------------------------------- 工具函数
def _norm(s) -> str:
    return re.sub(r"\s+", "", str(s if s is not None else "")).lower()


def map_columns(headers: list[str], keys) -> tuple[dict, list]:
    """把 AI 给的列名映射到真实表头：精确 > 忽略空格大小写 > 包含关系（唯一）。返回 (映射, 无法识别的列)。"""
    mapping, unknown = {}, []
    normed = {h: _norm(h) for h in headers if str(h).strip()}
    for k in keys:
        nk = _norm(k)
        hit = [h for h in normed if h == k] or [h for h, n in normed.items() if n == nk]
        if not hit:
            hit = [h for h, n in normed.items() if nk and (nk in n or n in nk)]
        if len(hit) == 1:
            mapping[k] = hit[0]
        else:
            unknown.append(k)
    return mapping, unknown


def _to_number(v):
    if isinstance(v, (int, float)):
        return v
    s = re.sub(r"[,¥￥$\s元]", "", str(v))
    if re.fullmatch(r"[-+]?\d+", s):
        return int(s)
    if re.fullmatch(r"[-+]?\d*\.\d+|[-+]?\d+\.\d*", s):
        return float(s)
    return None


def _convert(value, template_value, template_format: str | None):
    """参照模板单元格（上一行）的类型转换值。"""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    is_date_tpl = isinstance(template_value, (dt.date, dt.datetime)) or (
        template_format and any(x in template_format.lower() for x in ("yy", "mm-dd", "m/d", "年")))
    if is_date_tpl:
        d = core.parse_date(str(value))
        if d:
            return dt.datetime(d.year, d.month, d.day)
    if isinstance(template_value, (int, float)) and not isinstance(template_value, bool):
        n = _to_number(value)
        if n is not None:
            return n
    if isinstance(template_value, str):  # 上一行是文本（如发票号码），保持文本
        return str(value).strip()
    if isinstance(value, str):
        # 模板为空时：纯数字（且不是 0 开头的编号）转成数字，便于求和
        n = _to_number(value) if not re.match(r"^0\d", value.strip()) else None
        if n is not None and len(value.strip()) < 16:
            return n
        return value.strip()
    return value


# ----------------------------------------------------------------------------- 表头 / 位置
def _find_header(ws, max_scan: int = 15) -> tuple[int, dict]:
    best_row, best_score, best_map = None, -1, {}
    for r in range(1, min(ws.max_row, max_scan) + 1):
        vals = [(c.column, c.value) for c in ws[r] if c.value not in (None, "")]
        if len(vals) < 2:
            continue
        texty = sum(1 for _, v in vals if isinstance(v, str) and not core._looks_like_number(v) and not core.parse_date(v))
        score = len(vals) + 2 * texty
        if score > best_score:
            best_row, best_score = r, score
            best_map = {}
            for col, v in vals:
                name = str(v).strip()
                if name and name not in best_map:
                    best_map[name] = col
    if best_row is None:
        raise WriteError("没有找到表头行")
    return best_row, best_map


def _is_total_row(ws, r: int) -> bool:
    return any(isinstance(c.value, str) and any(w in c.value for w in TOTAL_WORDS) for c in ws[r])


def sheet_headers(path: Path, sheet: str | None = None) -> tuple[str, list[str]]:
    """返回 (工作表名, 表头列表)，供 AI 写入前校验列名。"""
    if path.suffix.lower() == ".csv":
        enc = _csv_encoding(path)
        with open(path, newline="", encoding=enc) as f:
            header = next(csv.reader(f), [])
        return "CSV", [h.strip() for h in header if h.strip()]
    if path.suffix.lower() == ".xls":
        raise WriteError("旧版 .xls 文件无法直接写入，请在 Excel 中另存为 .xlsx 后再试")
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = _pick_sheet(wb, sheet)
        rows = list(ws.iter_rows(min_row=1, max_row=15, values_only=True))
    finally:
        wb.close()
    best, best_score = [], -1
    for row in rows:
        vals = [str(v).strip() for v in row if v not in (None, "")]
        if len(vals) < 2:
            continue
        texty = sum(1 for v in vals if not core._looks_like_number(v) and not core.parse_date(v))
        if len(vals) + 2 * texty > best_score:
            best, best_score = vals, len(vals) + 2 * texty
    return ws.title, best


def _pick_sheet(wb, sheet):
    if not sheet:
        return wb.worksheets[0]
    for cond in (lambda n: n == sheet, lambda n: _norm(n) == _norm(sheet), lambda n: _norm(sheet) in _norm(n)):
        hit = [ws for ws in wb.worksheets if cond(ws.title)]
        if hit:
            return hit[0]
    raise WriteError(f"没有工作表「{sheet}」，可用：{[ws.title for ws in wb.worksheets]}")


def _csv_encoding(path: Path) -> str:
    raw = path.read_bytes()[:20000]
    for enc in ("utf-8-sig", "gb18030"):
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    return "utf-8-sig"


# ----------------------------------------------------------------------------- 备份 / 撤销
def backup(cfg: dict, path: Path) -> str:
    bdir = cfg["library_dir"] / BACKUP_DIR
    bdir.mkdir(parents=True, exist_ok=True)
    name = f"{dt.datetime.now():%Y%m%d_%H%M%S_%f}__{path.name}"
    shutil.copy2(path, bdir / name)
    return name


def restore_backup(cfg: dict, token: str, target: Path) -> None:
    bdir = (cfg["library_dir"] / BACKUP_DIR).resolve()
    src = (bdir / token).resolve()
    if src.parent != bdir or not src.exists():
        raise WriteError("备份不存在，无法撤销")
    try:
        shutil.copy2(src, target)
    except PermissionError:
        raise WriteError("文件正被 Excel 打开，请关闭后再撤销")


# ----------------------------------------------------------------------------- 写入
def append_rows(cfg: dict, path: Path, rows: list[dict], sheet: str | None = None) -> dict:
    """把 rows（{列名: 值}）追加到 path。返回 {sheet, written, first_row, backup}。"""
    if not rows:
        raise WriteError("没有要写入的数据")
    if not path.exists():
        raise WriteError(f"文件不存在：{path}")
    if core.is_locked(path):
        raise WriteError("文件正被 Excel 或其他程序打开，请先关闭再写入")
    ext = path.suffix.lower()
    if ext == ".xls":
        raise WriteError("旧版 .xls 文件无法直接写入，请在 Excel 中另存为 .xlsx 后再试")
    token = backup(cfg, path)
    try:
        res = _append_csv(path, rows) if ext == ".csv" else _append_xlsx(path, rows, sheet)
    except WriteError:
        raise
    except PermissionError:
        raise WriteError("文件正被 Excel 或其他程序打开，请先关闭再写入")
    res["backup"] = token
    core.log.info("已写入 %s 行到 %s「%s」(备份 %s)", res["written"], path.name, res["sheet"], token)
    return res


def _append_xlsx(path: Path, rows: list[dict], sheet: str | None) -> dict:
    from openpyxl import load_workbook

    keep_vba = path.suffix.lower() == ".xlsm"
    wb = load_workbook(path, keep_vba=keep_vba)
    ws = _pick_sheet(wb, sheet)
    header_row, header_map = _find_header(ws)
    keys = list(dict.fromkeys(k for r in rows for k in r))
    mapping, unknown = map_columns(list(header_map), keys)
    if unknown:
        raise WriteError(f"表中没有这些列：{unknown}；可用列：{list(header_map)}")

    cols = sorted(header_map.values())
    last = header_row
    for r in range(header_row + 1, ws.max_row + 1):
        if any(ws.cell(r, c).value not in (None, "") for c in cols):
            last = r
    insert_at = last + 1
    total_row = None
    if last > header_row and _is_total_row(ws, last):
        total_row, insert_at = last, last
        ws.insert_rows(insert_at, amount=len(rows))
        new_total = total_row + len(rows)
        _extend_total_formulas(ws, new_total, old_last_data=total_row - 1, shift=len(rows))

    tpl_row = insert_at - 1 if insert_at - 1 > header_row else None
    for i, row in enumerate(rows):
        r = insert_at + i
        for key, value in row.items():
            c = header_map[mapping[key]]
            tpl = ws.cell(tpl_row, c) if tpl_row else None
            cell = ws.cell(r, c)
            cell.value = _convert(value, tpl.value if tpl else None, tpl.number_format if tpl else None)
            if isinstance(cell.value, dt.datetime) and not (tpl and tpl.is_date):
                cell.number_format = "yyyy-mm-dd"
        if tpl_row:
            for c in range(1, ws.max_column + 1):
                src, dst = ws.cell(tpl_row, c), ws.cell(r, c)
                if src.has_style:
                    dst.font, dst.border, dst.fill = copy(src.font), copy(src.border), copy(src.fill)
                    dst.alignment, dst.protection = copy(src.alignment), copy(src.protection)
                    if not (isinstance(dst.value, dt.datetime) and not src.is_date):
                        dst.number_format = src.number_format
                # 上一行是公式而本行没填（如 =C5*D5）：沿用公式并调整行号
                if dst.value is None and isinstance(src.value, str) and src.value.startswith("=") \
                        and not any(w in src.value for w in TOTAL_WORDS):
                    dst.value = re.sub(rf"(?<![A-Z$])(\$?[A-Z]{{1,3}}){tpl_row}(?!\d)", rf"\g<1>{r}", src.value)
            if ws.row_dimensions[tpl_row].height:
                ws.row_dimensions[r].height = ws.row_dimensions[tpl_row].height
    try:
        wb.calculation.fullCalcOnLoad = True  # 用 Excel 打开时重新计算全部公式
    except Exception:
        pass
    wb.save(path)
    wb.close()
    return {"sheet": ws.title, "written": len(rows), "first_row": insert_at,
            "before_total": total_row is not None}


def _extend_total_formulas(ws, total_row: int, old_last_data: int, shift: int) -> None:
    """合计行下移后，把公式里以原最后一行数据结尾的区域（如 D3:D7）扩展到新的最后一行。"""
    pat = re.compile(r"(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?)(\d+)")
    for cell in ws[total_row]:
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cell.value = pat.sub(lambda m: m.group(1) + (str(int(m.group(2)) + shift)
                                                         if int(m.group(2)) == old_last_data else m.group(2)),
                                 cell.value)


def _append_csv(path: Path, rows: list[dict]) -> dict:
    enc = _csv_encoding(path)
    with open(path, newline="", encoding=enc) as f:
        header = next(csv.reader(f), [])
    if not header:
        raise WriteError("CSV 没有表头")
    keys = list(dict.fromkeys(k for r in rows for k in r))
    mapping, unknown = map_columns(header, keys)
    if unknown:
        raise WriteError(f"表中没有这些列：{unknown}；可用列：{header}")
    raw = path.read_bytes()
    needs_nl = bool(raw) and not raw.endswith(b"\n")
    with open(path, "a", newline="", encoding="utf-8" if enc == "utf-8-sig" else enc) as f:
        if needs_nl:
            f.write("\r\n")
        w = csv.writer(f)
        for row in rows:
            by_header = {mapping[k]: v for k, v in row.items()}
            w.writerow(["" if by_header.get(h) is None else by_header.get(h) for h in header])
    return {"sheet": "CSV", "written": len(rows), "first_row": None, "before_total": False}


def create_workbook(path: Path, columns: list[str], rows: list[dict], sheet: str = "Sheet1") -> dict:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    columns = [str(c).strip() for c in columns if str(c).strip()]
    if not columns:
        raise WriteError("新建表格需要列名")
    wb = Workbook()
    ws = wb.active
    ws.title = sheet[:31] or "Sheet1"
    ws.append(columns)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    for row in rows:
        vals = [_convert(row.get(col), None, None) for col in columns]
        ws.append(vals)
        for c in ws[ws.max_row]:
            if isinstance(c.value, dt.datetime):
                c.number_format = "yyyy-mm-dd"
    for i, col in enumerate(columns, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = max(10, min(40, len(col) * 2 + 4))
    ws.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return {"sheet": ws.title, "written": len(rows), "first_row": 2, "before_total": False}
