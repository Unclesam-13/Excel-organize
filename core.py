"""
Excel 图书馆 - 核心逻辑
收件箱里的 Excel -> 读取/画像 -> DeepSeek 归纳分类 -> 归档到 图书馆/类别/年-月/ -> 写入 SQLite -> 生成总目录
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import logging
import os
import re
import shutil
import sqlite3
import sys
import time
from logging.handlers import RotatingFileHandler
from pathlib import Path

import pandas as pd
import requests

APP_NAME = "Excel 图书馆"
VERSION = "1.3.0"
FROZEN = getattr(sys, "frozen", False)  # 是否为打包后的 exe
APP_DIR = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
EXCEL_EXTS = {".xlsx", ".xlsm", ".xls", ".csv"}
WEEKDAYS = "一二三四五六日"

log = logging.getLogger("excel_library")


def _documents_dir() -> Path:
    """Windows "文档" 文件夹（兼容被 OneDrive 重定向的情况）。"""
    if os.name == "nt":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders") as k:
                return Path(os.path.expandvars(winreg.QueryValueEx(k, "Personal")[0]))
        except OSError:
            pass
    return Path.home() / "Documents"


def _data_home() -> Path:
    """程序数据（配置/数据库/日志）位置：源码运行=程序目录；exe=%LOCALAPPDATA%\\ExcelLibrary。"""
    if os.environ.get("EXCEL_LIBRARY_HOME"):
        return Path(os.environ["EXCEL_LIBRARY_HOME"])
    if FROZEN:
        return Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "ExcelLibrary"
    return APP_DIR


DATA_HOME = _data_home()
CONFIG_PATH = DATA_HOME / "config.json"
BASE_DIR = DATA_HOME  # 配置里的相对路径以此为基准


def default_config() -> dict:
    user_root = (_documents_dir() / "Excel图书馆") if FROZEN else None
    return {
        "inbox_dir": str(user_root / "收件箱") if user_root else "收件箱",
        "library_dir": str(user_root / "图书馆") if user_root else "图书馆",
        "index_xlsx": str(user_root / "图书馆总目录.xlsx") if user_root else "图书馆总目录.xlsx",
        "db_path": "data/library.db",
        "log_dir": "logs",
        "after_ingest": "move",
        "scan_interval_seconds": 60,
        "settle_seconds": 5,
        "preview_rows": 15,
        "max_content_chars": 200000,
        "deepseek": {"api_key": "", "base_url": "https://api.deepseek.com", "model": "deepseek-flash",
                     "timeout": 120, "extra_body": {}},
        "default_category": "其他",
        "allow_new_category": False,
        "categories": {
            "财务": ["发票", "报销", "付款", "收款", "应收", "应付", "预算", "费用", "账", "利润", "成本", "税"],
            "销售": ["销售", "订单", "客户", "业绩", "成交", "报价单", "合同金额", "回款"],
            "采购": ["采购", "供应商", "进货", "询价", "比价", "到货"],
            "库存": ["库存", "入库", "出库", "盘点", "仓库", "物料", "SKU"],
            "人事": ["员工", "考勤", "工资", "薪资", "入职", "离职", "招聘", "绩效", "排班"],
            "项目": ["项目", "进度", "里程碑", "任务", "工时", "计划"],
            "生产": ["生产", "产量", "工单", "良率", "质检", "设备"],
            "其他": [],
        },
        "web_host": "127.0.0.1",
        "web_port": 8765,
        "web_auto_scan": True,
        "chat_full_data_chars": 40000,
        "chat_preview_rows": 30,
        "trash_keep_days": 30,
    }


def load_raw_config(path: Path | None = None) -> dict:
    """读取未解析路径的原始配置；不存在则生成默认配置，缺失的键用默认值补齐。"""
    path = Path(path or CONFIG_PATH)
    defaults = default_config()
    if not path.exists():
        save_raw_config(defaults, path)
        return defaults
    raw = json.loads(path.read_text(encoding="utf-8-sig"))
    for k, v in defaults.items():
        if k not in raw:
            raw[k] = v
        elif isinstance(v, dict) and k == "deepseek":
            for kk, vv in v.items():
                raw[k].setdefault(kk, vv)
    return raw


def save_raw_config(raw: dict, path: Path | None = None) -> None:
    path = Path(path or CONFIG_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def load_config(path: Path | None = None) -> dict:
    cfg = load_raw_config(path)
    for key in ("inbox_dir", "library_dir", "db_path", "index_xlsx", "log_dir"):
        p = Path(os.path.expandvars(str(cfg[key]))).expanduser()
        cfg[key] = p if p.is_absolute() else BASE_DIR / p
    ds = cfg["deepseek"]
    ds["api_key"] = os.environ.get("DEEPSEEK_API_KEY") or ds.get("api_key") or ""
    return cfg


def save_api_key(key: str, path: Path | None = None) -> None:
    raw = load_raw_config(path)
    raw["deepseek"]["api_key"] = key.strip()
    save_raw_config(raw, path)


def setup_logging(cfg: dict) -> None:
    if log.handlers:
        return
    cfg["log_dir"].mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    fh = RotatingFileHandler(cfg["log_dir"] / "library.log", maxBytes=5_000_000, backupCount=5, encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    if sys.stderr is not None:  # 无控制台的 exe 里 stderr 为 None
        sh = logging.StreamHandler()
        sh.setFormatter(fmt)
        log.addHandler(sh)
    log.setLevel(logging.INFO)


# ----------------------------------------------------------------------------- 数据库
SCHEMA = """
CREATE TABLE IF NOT EXISTS files (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    file_hash     TEXT UNIQUE,
    original_name TEXT,
    stored_path   TEXT,
    ingested_at   TEXT,
    data_date     TEXT,
    date_source   TEXT,
    date_min      TEXT,
    date_max      TEXT,
    category      TEXT,
    title         TEXT,
    summary       TEXT,
    tags          TEXT,
    keywords      TEXT,
    key_fields    TEXT,
    sheets        TEXT,
    total_rows    INTEGER,
    ai_used       INTEGER
);
CREATE INDEX IF NOT EXISTS idx_files_date ON files(data_date);
CREATE INDEX IF NOT EXISTS idx_files_cat  ON files(category);
CREATE TABLE IF NOT EXISTS contents (
    file_id INTEGER REFERENCES files(id) ON DELETE CASCADE,
    sheet   TEXT,
    text    TEXT
);
CREATE INDEX IF NOT EXISTS idx_contents_file ON contents(file_id);
CREATE TABLE IF NOT EXISTS seen (
    path  TEXT PRIMARY KEY,
    size  INTEGER,
    mtime REAL
);
"""


def connect(cfg: dict) -> sqlite3.Connection:
    cfg["db_path"].parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(cfg["db_path"], timeout=30)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA foreign_keys=ON")
    con.executescript(SCHEMA)
    cols = {r[1] for r in con.execute("PRAGMA table_info(files)")}
    for col in ("deleted_at", "prev_path"):  # v1.3 新增：回收站
        if col not in cols:
            con.execute(f"ALTER TABLE files ADD COLUMN {col} TEXT")
    con.commit()
    return con


def row_to_dict(row: sqlite3.Row) -> dict:
    d = dict(row)
    for k in ("tags", "keywords", "key_fields", "sheets"):
        try:
            d[k] = json.loads(d.get(k) or "[]")
        except (TypeError, json.JSONDecodeError):
            d[k] = []
    return d


# ----------------------------------------------------------------------------- 日期工具
_DATE_SEP_RE = re.compile(r"((?:19|20)\d{2})\s*[-/.年_]\s*(\d{1,2})\s*[-/.月_]\s*(\d{1,2})")
_DATE_COMPACT_RE = re.compile(r"(?<!\d)((?:19|20)\d{2})(\d{2})(\d{2})(?!\d)")


def _mkdate(y, m, d) -> dt.date | None:
    try:
        return dt.date(int(y), int(m), int(d))
    except ValueError:
        return None


def parse_date(text: str) -> dt.date | None:
    if not text:
        return None
    for rx in (_DATE_SEP_RE, _DATE_COMPACT_RE):
        for m in rx.finditer(str(text)):
            d = _mkdate(*m.groups())
            if d:
                return d
    return None


def _looks_like_number(v: str) -> bool:
    return bool(re.fullmatch(r"[-+]?[\d,]*\.?\d+%?", v.replace(" ", "")))


# ----------------------------------------------------------------------------- 读取 & 画像
def read_tables(path: Path) -> dict[str, pd.DataFrame]:
    if path.suffix.lower() == ".csv":
        last_err = None
        for enc in ("utf-8-sig", "gb18030"):
            try:
                return {"CSV": pd.read_csv(path, header=None, dtype=str, encoding=enc, keep_default_na=False)}
            except UnicodeDecodeError as e:
                last_err = e
        raise last_err  # type: ignore[misc]
    return pd.read_excel(path, sheet_name=None, header=None, dtype=str, keep_default_na=False)


def _clean(df: pd.DataFrame) -> pd.DataFrame:
    df = df.fillna("").astype(str)
    df = df.apply(lambda col: col.str.strip().str.replace(r"^(\d{4}-\d{2}-\d{2}) 00:00:00$", r"\1", regex=True))
    if df.empty:
        return df
    nonempty = df != ""
    return df.loc[nonempty.any(axis=1), nonempty.any(axis=0)].reset_index(drop=True)


def _guess_header(df: pd.DataFrame) -> int:
    best_i, best_score = 0, -1
    for i in range(min(10, len(df))):
        filled = [v for v in df.iloc[i] if v]
        if len(filled) < 2:
            continue
        texty = sum(1 for v in filled if not _looks_like_number(v) and not parse_date(v))
        score = len(filled) + 2 * texty
        if score > best_score:
            best_i, best_score = i, score
    return best_i


def profile_workbook(path: Path, preview_rows: int = 15, max_content_chars: int = 200_000) -> dict:
    """返回给 AI 看的概况文本 + 用于全文检索的内容文本 + 结构信息。"""
    tables = read_tables(path)
    sheets, profile_parts, contents = [], [], []
    all_dates: list[dt.date] = []
    total_rows = 0
    budget = max_content_chars

    for name, raw in tables.items():
        df = _clean(raw)
        if df.empty:
            sheets.append({"name": str(name), "rows": 0, "cols": 0, "header": []})
            continue
        h = _guess_header(df)
        header = [v for v in df.iloc[h].tolist()]
        body = df.iloc[h + 1:]
        total_rows += len(body)

        # 日期列
        sheet_dates = []
        for c in body.columns:
            vals = [v for v in body[c].tolist() if v]
            if not vals:
                continue
            sample = vals[:200]
            parsed = [parse_date(v) for v in sample]
            if sum(1 for p in parsed if p) / len(sample) >= 0.6:
                sheet_dates += [d for d in (parse_date(v) for v in vals) if d]
        all_dates += sheet_dates

        sheets.append({"name": str(name), "rows": int(len(body)), "cols": int(df.shape[1]),
                       "header": [x for x in header if x][:40]})

        if len(profile_parts) < 8:
            lines = [f"### 工作表「{name}」 {len(body)} 行 x {df.shape[1]} 列"]
            if h > 0:
                lines.append("表头上方: " + " / ".join(v for r in df.iloc[:h].values.tolist() for v in r if v)[:200])
            lines.append("表头: " + " | ".join(x for x in header if x)[:600])
            for r in body.head(preview_rows).values.tolist():
                lines.append(" | ".join(v[:30] for v in r)[:400])
            if sheet_dates:
                lines.append(f"日期范围: {min(sheet_dates)} ~ {max(sheet_dates)}")
            profile_parts.append("\n".join(lines))

        if budget > 0:
            text = "\n".join(" | ".join(v for v in r if v) for r in df.values.tolist())[:budget]
            budget -= len(text)
            contents.append((str(name), text))

    return {
        "sheets": sheets,
        "profile_text": "\n\n".join(profile_parts)[:9000],
        "contents": contents,
        "date_min": min(all_dates).isoformat() if all_dates else None,
        "date_max": max(all_dates).isoformat() if all_dates else None,
        "total_rows": total_rows,
    }


# ----------------------------------------------------------------------------- DeepSeek
class DeepSeek:
    def __init__(self, cfg: dict):
        ds = cfg.get("deepseek", {})
        self.api_key = ds.get("api_key", "")
        self.base_url = ds.get("base_url", "https://api.deepseek.com").rstrip("/")
        self.model = ds.get("model", "deepseek-flash")
        self.timeout = ds.get("timeout", 120)
        self.extra = ds.get("extra_body") or {}

    @property
    def enabled(self) -> bool:
        return bool(self.api_key)

    def _post(self, payload: dict, retries: int, parse=None):
        payload = {"model": self.model, "stream": False, **payload, **self.extra}
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        last_err: Exception | None = None
        for attempt in range(1, retries + 1):
            try:
                r = requests.post(f"{self.base_url}/chat/completions", json=payload, headers=headers, timeout=self.timeout)
                if r.status_code in (401, 402):
                    raise RuntimeError(f"DeepSeek 认证/余额错误 {r.status_code}: {r.text[:200]}")
                if r.status_code == 400:
                    raise RuntimeError(f"DeepSeek 请求被拒绝 400: {r.text[:300]}")
                r.raise_for_status()
                msg = r.json()["choices"][0]["message"]
                return parse(msg) if parse else msg
            except RuntimeError:
                raise
            except Exception as e:  # 网络、429、5xx、JSON 解析等 -> 重试
                last_err = e
                log.warning("DeepSeek 调用失败(第%d次): %s", attempt, e)
                if attempt < retries:
                    time.sleep(2 * attempt)
        raise RuntimeError(f"DeepSeek 多次调用失败: {last_err}")

    def chat_json(self, system: str, user: str, max_tokens: int = 2000, retries: int = 3) -> dict:
        def parse(msg):
            content = (msg.get("content") or "").strip()
            if not content:
                raise ValueError("DeepSeek 返回空内容")
            return json.loads(re.sub(r"^```(?:json)?\s*|\s*```$", "", content))

        return self._post({
            "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}],
            "response_format": {"type": "json_object"},
            "max_tokens": max_tokens,
            "temperature": 0.2,
        }, retries, parse)

    def chat(self, messages: list[dict], tools: list[dict] | None = None, max_tokens: int = 4000,
             retries: int = 2, tool_choice: str | None = None) -> dict:
        """普通多轮对话，可带函数调用工具；返回 assistant message（可能含 tool_calls）。"""
        payload = {"messages": messages, "max_tokens": max_tokens, "temperature": 0.3}
        if tools:
            payload["tools"] = tools
            if tool_choice:
                payload["tool_choice"] = tool_choice
        return self._post(payload, retries)


ANALYZE_SYSTEM = """你是一名严谨的资料管理员，负责给每天收到的 Excel 文件建立档案卡，方便以后检索。
阅读用户给出的文件概况，只输出 JSON，格式示例：
{{"title": "11月华东区销售日报", "category": "销售", "tags": ["华东区", "日报"], "summary": "……", "data_date": "2026-11-10", "key_fields": ["日期", "客户", "金额"], "keywords": ["销售额", "订单", "业绩"]}}
要求：
- title：20 字以内，说清这份表是什么
- category：{cat_rule}
- tags：3-6 个短标签（客户名、项目名、产品、部门、地区等实体）
- summary：60-150 字，说明内容、覆盖时间范围、关键数字或结论
- data_date：这份数据所属的业务日期 YYYY-MM-DD（优先文件名里的日期，其次表内日期），无法判断填 null
- key_fields：主要列名（最多 10 个）
- keywords：5-15 个以后可能用来搜到它的词，含同义词、简称"""


def _sanitize_name(name: str, maxlen: int = 80) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip(" .")
    return (name or "未命名")[:maxlen]


def _fallback_analyze(path: Path, prof: dict, cfg: dict) -> dict:
    """无 API Key 或 API 失败时的规则版归档。"""
    text = (path.stem + "\n" + prof["profile_text"]).lower()
    best, best_hits = cfg.get("default_category", "其他"), 0
    for cat, words in cfg["categories"].items():
        hits = text.count(cat.lower()) * 2 + sum(text.count(w.lower()) for w in words)
        if hits > best_hits:
            best, best_hits = cat, hits
    headers = [h for s in prof["sheets"] for h in s["header"]]
    rng = f"，日期 {prof['date_min']} ~ {prof['date_max']}" if prof["date_min"] else ""
    sheet_desc = "、".join(f"{s['name']}({s['rows']}行)" for s in prof["sheets"][:5])
    return {
        "title": path.stem[:40],
        "category": best,
        "tags": headers[:5],
        "summary": f"共 {len(prof['sheets'])} 个工作表：{sheet_desc}{rng}。主要列：{'、'.join(headers[:10])}。",
        "data_date": None,
        "key_fields": headers[:10],
        "keywords": headers[:15],
    }


def analyze(path: Path, prof: dict, cfg: dict, ai: DeepSeek) -> tuple[dict, bool]:
    cats = list(cfg["categories"].keys())
    if ai.enabled:
        cat_rule = (f"优先从 {cats} 中选一个；都不合适时可以新建一个 2-4 字的类别"
                    if cfg.get("allow_new_category") else f"必须从 {cats} 中选一个")
        user = (f"今天日期：{dt.date.today()}\n文件名：{path.name}\n"
                f"文件修改时间：{dt.datetime.fromtimestamp(path.stat().st_mtime):%Y-%m-%d %H:%M}\n\n{prof['profile_text']}")
        try:
            meta = ai.chat_json(ANALYZE_SYSTEM.format(cat_rule=cat_rule), user)
            return meta, True
        except Exception as e:
            log.error("AI 分析失败，改用规则归档：%s", e)
    return _fallback_analyze(path, prof, cfg), False


def _normalize_meta(meta: dict, cfg: dict) -> dict:
    cats = list(cfg["categories"].keys())
    cat = str(meta.get("category") or "").strip()
    if cat not in cats:
        cat = _sanitize_name(cat, 12) if (cat and cfg.get("allow_new_category")) else cfg.get("default_category", "其他")

    def as_list(v):
        if isinstance(v, str):
            v = re.split(r"[,，、;；\s]+", v)
        return [str(x).strip() for x in (v or []) if str(x).strip()][:20]

    return {
        "title": str(meta.get("title") or "").strip()[:60],
        "category": cat,
        "tags": as_list(meta.get("tags")),
        "summary": str(meta.get("summary") or "").strip()[:600],
        "data_date": meta.get("data_date"),
        "key_fields": as_list(meta.get("key_fields")),
        "keywords": as_list(meta.get("keywords")),
    }


# ----------------------------------------------------------------------------- 文件操作
def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def is_locked(path: Path) -> bool:
    """Excel 正在打开该文件时(Windows)无法以写方式打开，且旁边会有 ~$ 锁文件。"""
    if (path.parent / ("~$" + path.name[2:])).exists() or (path.parent / ("~$" + path.name)).exists():
        return True
    if not os.access(path, os.W_OK):  # 只读属性的文件，不算被占用
        return False
    try:
        with open(path, "r+b"):
            return False
    except (PermissionError, OSError):
        return True


def _unique_path(p: Path) -> Path:
    if not p.exists():
        return p
    i = 2
    while True:
        cand = p.with_name(f"{p.stem}_{i}{p.suffix}")
        if not cand.exists():
            return cand
        i += 1


def _move_or_copy(src: Path, dest: Path, mode: str) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest = _unique_path(dest)
    if mode == "copy":
        shutil.copy2(src, dest)
        return dest
    try:
        os.replace(src, dest)  # 同盘：原子移动
    except OSError:
        shutil.copy2(src, dest)  # 跨盘：复制后删除，删除失败则回滚
        try:
            src.unlink()
        except OSError:
            dest.unlink(missing_ok=True)
            raise
    return dest


def target_path(cfg: dict, category: str, data_date: str, original_name: str) -> Path:
    p = Path(original_name)
    stem = _sanitize_name(p.stem)
    if data_date and not parse_date(stem):
        stem = f"{data_date}_{stem}"
    return cfg["library_dir"] / _sanitize_name(category, 20) / data_date[:7] / f"{stem}{p.suffix.lower()}"


def _decide_date(path: Path, meta: dict, prof: dict) -> tuple[str, str]:
    d = parse_date(path.stem)
    if d:
        return d.isoformat(), "文件名"
    d = parse_date(str(meta.get("data_date") or ""))
    if d:
        return d.isoformat(), "AI判断"
    if prof.get("date_max"):
        return prof["date_max"], "表内日期"
    return dt.date.fromtimestamp(path.stat().st_mtime).isoformat(), "修改时间"


# ----------------------------------------------------------------------------- 入库
def ingest_file(path: Path, cfg: dict, con: sqlite3.Connection, ai: DeepSeek) -> dict:
    digest = file_sha256(path)
    dup = con.execute("SELECT id, stored_path, deleted_at FROM files WHERE file_hash=?", (digest,)).fetchone()
    if dup and dup["deleted_at"]:  # 同内容的文件在回收站里：彻底删除旧记录，按新文件重新入库
        purge_files(cfg, con, [dup["id"]])
        dup = None
    if dup:
        if cfg.get("after_ingest", "move") == "copy":
            log.info("重复文件(与 #%s 相同)，跳过：%s", dup["id"], path.name)
        else:
            dest = _move_or_copy(path, cfg["library_dir"] / "_重复文件" / path.name, "move")
            log.info("重复文件(与 #%s 相同)，已放到 %s", dup["id"], dest)
        return {"status": "duplicate", "id": dup["id"], "name": path.name}

    prof = profile_workbook(path, cfg.get("preview_rows", 15), cfg.get("max_content_chars", 200_000))
    raw_meta, ai_used = analyze(path, prof, cfg, ai)
    meta = _normalize_meta(raw_meta, cfg)
    data_date, date_source = _decide_date(path, raw_meta, prof)

    dest = _move_or_copy(path, target_path(cfg, meta["category"], data_date, path.name), cfg.get("after_ingest", "move"))
    try:
        with con:
            cur = con.execute(
                """INSERT INTO files(file_hash, original_name, stored_path, ingested_at, data_date, date_source,
                       date_min, date_max, category, title, summary, tags, keywords, key_fields, sheets, total_rows, ai_used)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (digest, path.name, str(dest), dt.datetime.now().isoformat(timespec="seconds"), data_date, date_source,
                 prof["date_min"], prof["date_max"], meta["category"], meta["title"] or path.stem, meta["summary"],
                 json.dumps(meta["tags"], ensure_ascii=False), json.dumps(meta["keywords"], ensure_ascii=False),
                 json.dumps(meta["key_fields"], ensure_ascii=False), json.dumps(prof["sheets"], ensure_ascii=False),
                 prof["total_rows"], int(ai_used)))
            fid = cur.lastrowid
            con.executemany("INSERT INTO contents(file_id, sheet, text) VALUES(?,?,?)",
                            [(fid, s, t) for s, t in prof["contents"]])
    except Exception:
        if cfg.get("after_ingest", "move") == "move" and not path.exists():
            os.replace(dest, path)  # 入库失败，把文件放回收件箱
        raise
    log.info("已入库 #%s [%s] %s  %s  -> %s", fid, meta["category"], data_date, meta["title"], dest)
    return {"status": "ok", "id": fid, "name": path.name, "category": meta["category"], "date": data_date,
            "title": meta["title"], "ai": ai_used}


class ScanLock:
    """跨进程锁：防止"每日任务"和"网页后台扫描"同时处理同一批文件。"""

    def __init__(self, cfg: dict, stale_seconds: int = 1800):
        self.path = cfg["db_path"].parent / "scan.lock"
        self.stale = stale_seconds
        self.acquired = False

    def __enter__(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.exists() and time.time() - self.path.stat().st_mtime > self.stale:
            self.path.unlink(missing_ok=True)
        try:
            fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            self.acquired = True
        except FileExistsError:
            self.acquired = False
        return self

    def __exit__(self, *exc):
        if self.acquired:
            self.path.unlink(missing_ok=True)


def _inbox_files(inbox: Path) -> list[Path]:
    out = []
    for p in inbox.rglob("*"):
        rel = p.relative_to(inbox).parts
        if (p.is_file() and p.suffix.lower() in EXCEL_EXTS and not p.name.startswith("~$")
                and not any(part.startswith("_") for part in rel[:-1])):
            out.append(p)
    return sorted(out, key=lambda x: x.stat().st_mtime)


def scan_inbox(cfg: dict, ai: DeepSeek | None = None) -> dict:
    ai = ai or DeepSeek(cfg)
    cfg["inbox_dir"].mkdir(parents=True, exist_ok=True)
    cfg["library_dir"].mkdir(parents=True, exist_ok=True)
    stats = {"ok": 0, "duplicate": 0, "failed": 0, "waiting": 0, "items": []}
    with ScanLock(cfg) as lock:
        if not lock.acquired:
            log.info("另一个进程正在处理收件箱，本次跳过")
            stats["busy"] = True
            return stats
        con = connect(cfg)
        copy_mode = cfg.get("after_ingest", "move") == "copy"
        try:
            for p in _inbox_files(cfg["inbox_dir"]):
                st = p.stat()
                if copy_mode:  # 复制模式下原件留在收件箱，处理过且未改动的直接跳过
                    s = con.execute("SELECT size, mtime FROM seen WHERE path=?", (str(p),)).fetchone()
                    if s and s["size"] == st.st_size and s["mtime"] == st.st_mtime:
                        continue
                if time.time() - st.st_mtime < cfg.get("settle_seconds", 5) or is_locked(p):
                    stats["waiting"] += 1
                    log.info("文件仍在写入或被 Excel 打开，稍后再处理：%s", p.name)
                    continue
                try:
                    res = ingest_file(p, cfg, con, ai)
                    if copy_mode:
                        with con:
                            con.execute("INSERT OR REPLACE INTO seen(path, size, mtime) VALUES(?,?,?)",
                                        (str(p), st.st_size, st.st_mtime))
                    stats[res["status"]] += 1
                    stats["items"].append(res)
                except Exception as e:
                    stats["failed"] += 1
                    log.exception("处理失败：%s", p.name)
                    try:
                        _move_or_copy(p, cfg["inbox_dir"] / "_处理失败" / p.name, "move")
                    except OSError:
                        pass
                    stats["items"].append({"status": "failed", "name": p.name, "error": str(e)})
            keep = int(cfg.get("trash_keep_days", 30) or 0)
            purged = purge_files(cfg, con, older_than_days=keep) if keep > 0 else 0
            if purged:
                log.info("回收站中超过 %s 天的 %s 个文件已彻底删除", keep, purged)
            if stats["ok"]:
                export_index(cfg, con)
        finally:
            con.close()
    return stats


def reanalyze(cfg: dict, con: sqlite3.Connection, ai: DeepSeek, file_id: int) -> dict:
    """对已入库文件重新跑 AI（比如后来才填了 API Key），类别变了会自动挪文件夹。"""
    row = con.execute("SELECT * FROM files WHERE id=?", (file_id,)).fetchone()
    if not row:
        raise KeyError(f"没有编号 {file_id}")
    path = Path(row["stored_path"])
    prof = profile_workbook(path, cfg.get("preview_rows", 15), cfg.get("max_content_chars", 200_000))
    raw, ai_used = analyze(path, prof, cfg, ai)
    meta = _normalize_meta(raw, cfg)
    d = parse_date(Path(row["original_name"]).stem) or parse_date(str(raw.get("data_date") or ""))
    data_date = d.isoformat() if d else row["data_date"]
    new_path = path
    if meta["category"] != row["category"] or data_date[:7] != (row["data_date"] or "")[:7]:
        new_path = _move_or_copy(path, target_path(cfg, meta["category"], data_date, row["original_name"]), "move")
    with con:
        con.execute("""UPDATE files SET category=?, title=?, summary=?, tags=?, keywords=?, key_fields=?,
                       data_date=?, stored_path=?, ai_used=? WHERE id=?""",
                    (meta["category"], meta["title"] or row["title"], meta["summary"],
                     json.dumps(meta["tags"], ensure_ascii=False), json.dumps(meta["keywords"], ensure_ascii=False),
                     json.dumps(meta["key_fields"], ensure_ascii=False), data_date, str(new_path), int(ai_used), file_id))
    log.info("已重新归档 #%s [%s] %s", file_id, meta["category"], meta["title"])
    return {"id": file_id, **meta, "data_date": data_date}


# ----------------------------------------------------------------------------- 删除 / 回收站
TRASH_DIR = "_回收站"


def trash_files(cfg: dict, con: sqlite3.Connection, ids) -> int:
    """删除：文件移到 图书馆\\_回收站\\，记录标记为已删除，可恢复。"""
    n = 0
    for fid in ids:
        row = con.execute("SELECT * FROM files WHERE id=? AND deleted_at IS NULL", (int(fid),)).fetchone()
        if not row:
            continue
        src = Path(row["stored_path"])
        dest = src
        if src.exists():
            dest = _move_or_copy(src, cfg["library_dir"] / TRASH_DIR / src.name, "move")
        with con:
            con.execute("UPDATE files SET deleted_at=?, prev_path=?, stored_path=? WHERE id=?",
                        (dt.datetime.now().isoformat(timespec="seconds"), str(src), str(dest), row["id"]))
        log.info("已删除(移到回收站) #%s %s", row["id"], row["title"])
        n += 1
    if n:
        export_index(cfg, con)
    return n


def restore_files(cfg: dict, con: sqlite3.Connection, ids) -> int:
    n = 0
    for fid in ids:
        row = con.execute("SELECT * FROM files WHERE id=? AND deleted_at IS NOT NULL", (int(fid),)).fetchone()
        if not row:
            continue
        src = Path(row["stored_path"])
        target = Path(row["prev_path"]) if row["prev_path"] else \
            target_path(cfg, row["category"], row["data_date"], row["original_name"])
        dest = _move_or_copy(src, target, "move") if src.exists() else src
        with con:
            con.execute("UPDATE files SET deleted_at=NULL, prev_path=NULL, stored_path=? WHERE id=?", (str(dest), row["id"]))
        log.info("已恢复 #%s %s", row["id"], row["title"])
        n += 1
    if n:
        export_index(cfg, con)
    return n


def _remove_file(path: Path) -> None:
    if not path.exists():
        return
    try:
        import winsys
        if winsys.send_to_recycle_bin(path):  # Windows：放进系统回收站，还能找回
            return
    except Exception as e:
        log.warning("放入系统回收站失败，直接删除：%s", e)
    path.unlink()


def purge_files(cfg: dict, con: sqlite3.Connection, ids=None, older_than_days: int | None = None) -> int:
    """彻底删除回收站里的文件：ids=None 表示全部（可配合 older_than_days）。"""
    sql, params = "SELECT id, stored_path, title FROM files WHERE deleted_at IS NOT NULL", []
    if ids is not None:
        ids = [int(i) for i in ids]
        if not ids:
            return 0
        sql += f" AND id IN ({','.join('?' * len(ids))})"
        params += ids
    if older_than_days:
        sql += " AND deleted_at < ?"
        params.append((dt.datetime.now() - dt.timedelta(days=older_than_days)).isoformat(timespec="seconds"))
    n = 0
    for row in con.execute(sql, params).fetchall():
        try:
            _remove_file(Path(row["stored_path"]))
        except OSError as e:
            log.error("删除文件失败 #%s：%s（文件可能正被打开）", row["id"], e)
            continue
        with con:
            con.execute("DELETE FROM contents WHERE file_id=?", (row["id"],))
            con.execute("DELETE FROM files WHERE id=?", (row["id"],))
        log.info("已彻底删除 #%s %s", row["id"], row["title"])
        n += 1
    return n


def list_trash(con: sqlite3.Connection) -> list[dict]:
    return [row_to_dict(r) for r in con.execute(
        "SELECT * FROM files WHERE deleted_at IS NOT NULL ORDER BY deleted_at DESC")]


# ----------------------------------------------------------------------------- 检索
META_COLS = ("title", "summary", "tags", "keywords", "key_fields", "original_name", "category")


def _like(kw: str) -> str:
    return "%" + kw.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%"


def search(con: sqlite3.Connection, keywords=(), categories=(), date_from=None, date_to=None,
           match: str = "all", limit: int = 100) -> list[dict]:
    keywords = [k.strip() for k in keywords if k and k.strip()]
    categories = [c for c in categories if c]
    where, params = ["deleted_at IS NULL"], []
    if categories:
        where.append(f"category IN ({','.join('?' * len(categories))})")
        params += categories
    if date_from:
        where.append("data_date >= ?")
        params.append(str(date_from))
    if date_to:
        where.append("data_date <= ?")
        params.append(str(date_to))
    clauses = []
    for kw in keywords:
        meta = " OR ".join(f"{c} LIKE ? ESCAPE '\\'" for c in META_COLS)
        clauses.append(f"(({meta}) OR id IN (SELECT file_id FROM contents WHERE text LIKE ? ESCAPE '\\'))")
        params += [_like(kw)] * (len(META_COLS) + 1)
    if clauses:
        where.append("(" + (" AND " if match == "all" else " OR ").join(clauses) + ")")
    sql = "SELECT * FROM files" + (" WHERE " + " AND ".join(where) if where else "")
    sql += " ORDER BY data_date DESC, id DESC LIMIT ?"
    params.append(limit if not keywords else max(limit * 5, 300))
    rows = [row_to_dict(r) for r in con.execute(sql, params)]

    if keywords:  # 元数据命中权重 3，仅内容命中权重 1
        for r in rows:
            blob = " ".join(str(r.get(c) or "") for c in META_COLS).lower()
            r["score"] = sum(3 if k.lower() in blob else 1 for k in keywords)
        rows.sort(key=lambda r: (r["score"], r["data_date"] or ""), reverse=True)
    return rows[:limit]


def stats(con: sqlite3.Connection) -> dict:
    return {
        "total": con.execute("SELECT COUNT(*) FROM files WHERE deleted_at IS NULL").fetchone()[0],
        "trash": con.execute("SELECT COUNT(*) FROM files WHERE deleted_at IS NOT NULL").fetchone()[0],
        "by_category": [tuple(r) for r in con.execute(
            "SELECT category, COUNT(*) FROM files WHERE deleted_at IS NULL GROUP BY category ORDER BY COUNT(*) DESC")],
        "by_month": [tuple(r) for r in con.execute(
            "SELECT substr(data_date,1,7) m, COUNT(*) FROM files WHERE deleted_at IS NULL GROUP BY m ORDER BY m DESC")],
    }


# ----------------------------------------------------------------------------- AI 找文件
PARSE_SYSTEM = """你负责把用户“找 Excel 文件”的自然语言需求转成检索条件，只输出 JSON：
{{"date_from": "YYYY-MM-DD 或 null", "date_to": "YYYY-MM-DD 或 null", "categories": [], "keywords": [], "explain": "一句话说明你的理解"}}
今天是 {today}（星期{wd}）。可选类别：{cats}
规则：
- 相对时间换算成具体日期：“上周”=上周一到上周日，“本月”=本月1日到今天，“最近”未说明默认近 30 天；没提时间就都填 null
- categories 只能从可选类别里选，用户没明确提到就留空数组（宁可不限）
- keywords：最能区分文件的实体词/主题词，每个 2-8 字，可附 1-2 个同义词；不要放日期、类别名，也不要放“表格、文件、Excel、数据、帮我找”这类泛词"""

RERANK_SYSTEM = """下面是 Excel 资料库里的候选文件档案（JSON 列表）。根据用户需求挑出真正相关的文件（最多 {n} 个），按相关度从高到低排序，只输出 JSON：
{{"results": [{{"id": 12, "reason": "一句话说明为什么相关"}}], "answer": "用一两句话直接回答用户（例如找到了几份、最相关的是哪份）"}}
如果都不相关，results 返回空数组，并在 answer 里说明。"""

_STOP = ["帮我", "请", "找一下", "找找", "查一下", "查查", "查找", "搜索", "找到", "找出", "找", "一下", "给我", "看看",
         "有没有", "相关", "关于", "有关", "的", "了", "吗", "呢", "那个", "那份", "这个", "这份", "所有", "全部",
         "表格", "表", "文件", "excel", "Excel", "数据", "资料", "记录", "一份", "哪些", "哪份", "什么", "里面", "中"]


def _rough_parse(question: str, cats: list[str]) -> dict:
    """无 API 时的简易解析：识别常见时间词、类别名，剩余片段当关键词。"""
    today = dt.date.today()
    q = question
    df = dt_ = None
    monday = today - dt.timedelta(days=today.weekday())
    rules = [
        ("前天", today - dt.timedelta(2), today - dt.timedelta(2)),
        ("昨天", today - dt.timedelta(1), today - dt.timedelta(1)),
        ("今天", today, today),
        ("上周", monday - dt.timedelta(7), monday - dt.timedelta(1)),
        ("上个星期", monday - dt.timedelta(7), monday - dt.timedelta(1)),
        ("本周", monday, today), ("这周", monday, today),
        ("上个月", (today.replace(day=1) - dt.timedelta(1)).replace(day=1), today.replace(day=1) - dt.timedelta(1)),
        ("上月", (today.replace(day=1) - dt.timedelta(1)).replace(day=1), today.replace(day=1) - dt.timedelta(1)),
        ("本月", today.replace(day=1), today), ("这个月", today.replace(day=1), today),
        ("今年", today.replace(month=1, day=1), today),
    ]
    for word, a, b in rules:
        if word in q:
            df, dt_ = a, b
            q = q.replace(word, " ")
            break
    m = re.search(r"最近\s*(\d+)\s*天", q)
    if m:
        df, dt_ = today - dt.timedelta(int(m.group(1))), today
        q = q.replace(m.group(0), " ")
    for m in list(_DATE_SEP_RE.finditer(q)):
        d = _mkdate(*m.groups())
        if d:
            df = df or d
            dt_ = d
            q = q.replace(m.group(0), " ")
    found_cats = [c for c in cats if c in q]
    for c in found_cats:
        q = q.replace(c, " ")
    for w in sorted(_STOP, key=len, reverse=True):
        q = q.replace(w, " ")
    kws = [k for k in re.split(r"[\s,，。.?？!！、;；:：\"'“”()（）]+", q) if len(k) >= 2]
    return {"date_from": df.isoformat() if df else None, "date_to": dt_.isoformat() if dt_ else None,
            "categories": found_cats, "keywords": kws, "explain": "（未配置 API Key，使用简易解析）"}


def ask(con: sqlite3.Connection, question: str, cfg: dict, ai: DeepSeek, top_n: int = 8) -> dict:
    cats = list(cfg["categories"].keys())
    known_cats = [r[0] for r in con.execute("SELECT DISTINCT category FROM files WHERE deleted_at IS NULL")]
    cats = cats + [c for c in known_cats if c not in cats]

    filters = None
    if ai.enabled:
        today = dt.date.today()
        try:
            filters = ai.chat_json(PARSE_SYSTEM.format(today=today, wd=WEEKDAYS[today.weekday()], cats=cats),
                                   question, max_tokens=600)
        except Exception as e:
            log.error("AI 解析问题失败：%s", e)
    if not filters:
        filters = _rough_parse(question, cats)

    d_from = parse_date(str(filters.get("date_from") or ""))
    d_to = parse_date(str(filters.get("date_to") or ""))
    f_cats = [c for c in (filters.get("categories") or []) if c in cats]
    kws = [str(k) for k in (filters.get("keywords") or []) if str(k).strip()]

    # 逐级放宽：全部条件 -> 去掉关键词 -> 去掉时间/类别只用关键词
    seen, cands = set(), []
    for args in ((kws, f_cats, d_from, d_to), ([], f_cats, d_from, d_to), (kws, [], None, None)):
        if not args[0] and not (args[1] or args[2] or args[3]):
            continue
        for r in search(con, args[0], args[1], args[2], args[3], match="any", limit=60):
            if r["id"] not in seen:
                seen.add(r["id"])
                cands.append(r)
        if len(cands) >= 5:
            break
    if not cands and not kws and not (f_cats or d_from or d_to):
        cands = search(con, limit=40)
    cands = cands[:40]

    answer = ""
    results = []
    if ai.enabled and cands:
        brief = [{"id": r["id"], "日期": r["data_date"], "类别": r["category"], "标题": r["title"],
                  "摘要": (r["summary"] or "")[:160], "标签": r["tags"][:6], "原文件名": r["original_name"]}
                 for r in cands]
        try:
            rr = ai.chat_json(RERANK_SYSTEM.format(n=top_n),
                              f"用户需求：{question}\n\n候选文件：\n{json.dumps(brief, ensure_ascii=False)}",
                              max_tokens=1500)
            by_id = {r["id"]: r for r in cands}
            for item in rr.get("results", [])[:top_n]:
                try:
                    rid = int(item.get("id"))
                except (TypeError, ValueError):
                    continue
                if rid in by_id:
                    results.append({**by_id[rid], "reason": item.get("reason", "")})
            answer = rr.get("answer", "")
        except Exception as e:
            log.error("AI 排序失败，使用关键词排序：%s", e)
    if not results and not answer:
        results = [{**r, "reason": "条件匹配"} for r in cands[:top_n]]
        answer = f"找到 {len(cands)} 份候选文件，显示前 {len(results)} 份。" if cands else "没有找到符合条件的文件。"

    return {"question": question,
            "filters": {"date_from": d_from.isoformat() if d_from else None,
                        "date_to": d_to.isoformat() if d_to else None,
                        "categories": f_cats, "keywords": kws, "explain": filters.get("explain", "")},
            "answer": answer, "results": results}


# ----------------------------------------------------------------------------- 总目录 Excel
def export_index(cfg: dict, con: sqlite3.Connection | None = None) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    own = con is None
    con = con or connect(cfg)
    try:
        rows = [row_to_dict(r) for r in con.execute(
            "SELECT * FROM files WHERE deleted_at IS NULL ORDER BY data_date DESC, id DESC")]
    finally:
        if own:
            con.close()

    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="1F4E78")
    head_font = Font(bold=True, color="FFFFFF")

    def write_sheet(ws, headers, data, widths):
        ws.append(headers)
        for c in ws[1]:
            c.fill, c.font = head_fill, head_font
            c.alignment = Alignment(vertical="center")
        for r in data:
            ws.append(r)
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        if data:
            ws.auto_filter.ref = ws.dimensions

    ws = wb.active
    ws.title = "总目录"
    headers = ["编号", "业务日期", "类别", "标题", "摘要", "标签", "数据行数", "原文件名", "打开", "入库时间", "日期来源"]
    data = [[r["id"], r["data_date"], r["category"], r["title"], r["summary"], "、".join(r["tags"]),
             r["total_rows"], r["original_name"], "打开", r["ingested_at"], r["date_source"]] for r in rows]
    write_sheet(ws, headers, data, [7, 12, 10, 30, 60, 28, 10, 30, 8, 19, 10])
    link_font = Font(color="0563C1", underline="single")
    for i, r in enumerate(rows, start=2):
        c = ws.cell(row=i, column=9)
        c.hyperlink = r["stored_path"]
        c.font = link_font
        ws.cell(row=i, column=5).alignment = Alignment(wrap_text=True, vertical="top")

    # 类别 x 月份 统计
    ws2 = wb.create_sheet("分类统计")
    months = sorted({(r["data_date"] or "")[:7] for r in rows}, reverse=True)
    cats = sorted({r["category"] for r in rows})
    grid = {(c, m): 0 for c in cats for m in months}
    for r in rows:
        grid[(r["category"], (r["data_date"] or "")[:7])] += 1
    data2 = [[c] + [grid[(c, m)] for m in months] + [sum(grid[(c, m)] for m in months)] for c in cats]
    write_sheet(ws2, ["类别"] + months + ["合计"], data2, [12] + [10] * (len(months) + 1))

    out = cfg["index_xlsx"]
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(out)
    except PermissionError:
        out = out.with_name(f"{out.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}{out.suffix}")
        wb.save(out)
        log.warning("总目录正被 Excel 打开，已另存为 %s", out.name)
    return out
