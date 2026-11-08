#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
学生档案管理系统 v7（首页精简 + 学生画像版）

在 v6.1 数据保真修复的基础上：
  1. 首页表格只显示学生的基本信息（学号、姓名、性别、电话号码、宿舍号、是否走读），
     数据列增多后首页不再臃肿；筛选、搜索、导入、导出功能不受影响。
  2. 双击首页任意学生行，弹出独立的「学生画像」窗口：
     顶部画像卡（姓名 + 关键标签），下方分区页展示详细情况；
     其中「学业画像」页把单科成绩归到所属考试批次之下（批次 → 科目两级），
     不再与往次考试并列平铺；其余各分类页可直接编辑并保存。
  3. 沿用 v6.1 的列类型分派与导入导出规则。

v6.1 修复要点（保留）：
  1. 旧版把「每一个单元格」都交给日期解析器，数字 95 会变成 1900-04-04、
     字符串 "95" 会变成 1995-11-08、"1" 会变成 2026-11-01，日期中文写法还会串年份；
     同时列名含「谈话 / 记录 / 时间」的文本列（如「谈话内容」）被误判为日期列，
     导出时被整列 coerce 成空值。
  2. 现在为每一列建立明确类型（text / number / date / id），按列分派转换，
     数据库统一存 TEXT，导出时按类型还原成 Excel 原生类型与显示格式。
  3. 任何解析失败都原样保留，绝不猜测、绝不清空。
  4. 导入支持「无学号但有姓名」的成绩表：按姓名匹配学号、剔除表尾统计行、
     跳过常量列，导入前自动备份数据库并给出校验摘要。
"""

import sys
import os
import re
import math
import shutil
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog
except ImportError:
    # 允许在没有图形界面的 Python 环境里只导入数据层（供批处理脚本复用）
    tk = None
    ttk = filedialog = messagebox = simpledialog = None
import pandas as pd
import sqlite3
from datetime import datetime, date, timedelta
from numbers import Number
import traceback

# 检查依赖  使用pip install pandas openpyxl
def check_dependencies():
    """检查并安装依赖"""
    missing = []
    
    try:
        import pandas  # noqa
    except ImportError:
        missing.append('pandas')
    
    try:
        import openpyxl  # noqa
    except ImportError:
        missing.append('openpyxl')
    
    if missing:
        msg = f"缺少以下依赖库：{', '.join(missing)}\n\n"
        msg += "请在命令行运行以下命令安装：\n"
        msg += f"pip install {' '.join(missing)}\n\n"
        msg += "或者使用国内镜像（更快）：\n"
        msg += f"pip install -i https://pypi.tuna.tsinghua.edu.cn/simple {' '.join(missing)}"
        if tk is not None:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("缺少依赖", msg)
        else:
            print(msg)
        return False
    
    return True

if not check_dependencies():
    sys.exit(1)


def center_window_relative(child, parent, offset_x=0, offset_y=0):
    """让子窗口基于父窗口位置进行居中/偏移显示"""
    if parent is None or child is None:
        return
    parent.update_idletasks()
    child.update_idletasks()
    
    px, py = parent.winfo_x(), parent.winfo_y()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    
    cw, ch = child.winfo_width(), child.winfo_height()
    if cw <= 1 or ch <= 1:
        cw = child.winfo_reqwidth()
        ch = child.winfo_reqheight()
    
    x = px + (pw - cw) // 2 + offset_x
    y = py + (ph - ch) // 2 + offset_y
    x = max(x, 0)
    y = max(y, 0)
    
    child.geometry(f"{cw}x{ch}+{x}+{y}")
    child.deiconify()


# ==========================================================================
# 数据层：与 tkinter 无关，GUI 与命令行脚本共用同一套规则
#
# 核心约定：
#   * 列类型 text / number / date / id 记录在 column_meta 表中
#   * SQLite 里所有业务列仍是 TEXT（便于动态增删列），类型只影响「怎么读进来、
#     怎么写出到 Excel」
#   * 解析失败一律保留原始文本，绝不 coerce 成空值
# ==========================================================================

DATE_KEYWORDS = ("日期", "生日", "出生", "入校", "毕业", "时间")
ID_KEYWORDS = ("学号", "身份证", "电话", "手机", "学籍号", "编号", "考号",
               "宿舍号", "邮编", "银行卡", "账号")
NUMBER_KEYWORDS = ("成绩", "分数", "总分", "得分", "名次", "语文", "数学", "英语",
                   "政治", "专业理论", "文化总分", "课时", "次数", "人数", "金额")
# 明确的文本列，优先于数值推断，防止「姓名」「谈话内容」等被误判成数值
TEXT_KEYWORDS = ("姓名", "性别", "地址", "内容", "简述", "备注", "说明", "是否",
                 "现居地", "民族", "政治面貌", "专业名称", "班级", "情况简述")
# Excel 序列号只在现代日期区间内才解释为日期（约 1954 年以后，序列号 ≥ 20000），
# 这样 95、456 这类分数/编号绝不会被转成 1900 年的日期
EXCEL_SERIAL_MIN = 20000
EXCEL_SERIAL_MAX = 80000

EXCEL_EPOCH = datetime(1899, 12, 30)
_ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_CN_DATE_RE = re.compile(r"^(\d{4})\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})\s*[日号]?$")
_DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
                 "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S")


def _is_missing(value):
    """安全判空，兼容 None / NaN / NaT / pd.NA。"""
    if value is None:
        return True
    try:
        flag = pd.isna(value)
        if isinstance(flag, bool):
            return flag
        if hasattr(flag, "item"):
            return bool(flag.item())
        return False
    except (TypeError, ValueError):
        return False


def quote_identifier(name):
    """SQLite 标识符转义，防止列名里的双引号破坏 SQL。"""
    return '"' + str(name).replace('"', '""') + '"'


def normalize_number(value):
    """数值保真：整数去掉 .0，浮点收敛精度，绝不解释为日期。"""
    if _is_missing(value):
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, Number):
        f = float(value)
        if not math.isfinite(f):
            return str(value)
        if f.is_integer():
            return str(int(f))
        return ("%.6f" % f).rstrip("0").rstrip(".")
    s = str(value).strip()
    if not s:
        return None
    m = re.fullmatch(r"([+-]?\d+)\.0+", s)
    if m:
        return m.group(1)
    if re.fullmatch(r"[+-]?\d+(\.\d+)?[eE][+-]?\d+", s):
        try:
            f = float(s)
            if f.is_integer():
                return str(int(f))
        except ValueError:
            pass
    return s


def normalize_text(value):
    """文本/编号保真：不做任何日期或数值解释。"""
    if _is_missing(value):
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, Number):
        return normalize_number(value)
    s = str(value).strip()
    return s if s else None


def parse_and_normalize_date(value):
    """严格日期解析：只认明确的日期形态，失败一律原样返回，绝不猜年份。"""
    if _is_missing(value):
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, Number):
        f = float(value)
        # 只有整数、且落在合理 Excel 序列号区间，才当作日期序列号
        if math.isfinite(f) and f.is_integer() and EXCEL_SERIAL_MIN <= f <= EXCEL_SERIAL_MAX:
            try:
                return (EXCEL_EPOCH + timedelta(days=int(f))).strftime("%Y-%m-%d")
            except (OverflowError, ValueError):
                pass
        return normalize_number(value)
    s = str(value).strip()
    if not s:
        return None
    m = _CN_DATE_RE.match(s)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).strftime("%Y-%m-%d")
        except ValueError:
            return s
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    if re.fullmatch(r"\d{8}", s):
        try:
            return datetime.strptime(s, "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            return s
    # "95"、"1" 这类纯数字，以及 "11月27日" 这类无年份写法：保持原样
    return s


def normalize_id(value):
    """学号等标识符：还原科学计数法，保留前导零。"""
    s = normalize_text(value)
    if s is None:
        return None
    if re.fullmatch(r"[+-]?\d+(\.\d+)?[eE][+-]?\d+", s):
        try:
            f = float(s)
            if f.is_integer():
                return str(int(f))
        except ValueError:
            pass
    return s


def normalize_cell_value(value, col_type="text"):
    """按列类型分派——这是修复的关键：不再对所有单元格统一猜日期。"""
    if col_type == "date":
        return parse_and_normalize_date(value)
    if col_type == "number":
        return normalize_number(value)
    return normalize_text(value)


def infer_column_type(col, values, known=None):
    """推断列类型；已登记的类型优先沿用，避免反复导入把日期列改成文本。"""
    if known:
        return known
    if col == "学号":
        return "id"
    name = str(col)
    # 文本关键词优先于编号关键词，避免「身份证地址」被当成编号列
    if any(k in name for k in TEXT_KEYWORDS):
        return "text"
    if any(k in name for k in DATE_KEYWORDS):
        return "date"
    if any(k in name for k in ID_KEYWORDS):
        return "id"
    if any(k in name for k in NUMBER_KEYWORDS):
        return "number"
    samples = [v for v in values if not _is_missing(v) and str(v).strip()][:200]
    if not samples:
        return "text"
    numeric = 0
    for v in samples:
        try:
            float(str(v).strip())
            numeric += 1
        except (TypeError, ValueError):
            pass
    if numeric / len(samples) < 0.8:
        return "text"
    # 全是 8 位以上的长数字，多半是编号/卡号，按文本保真，避免科学计数法
    long_ids = sum(1 for v in samples
                   if re.fullmatch(r"\d{8,}", str(v).strip()))
    return "id" if long_ids / len(samples) >= 0.8 else "number"


def ensure_meta_table(db_path, conn=None):
    own = conn is None
    c = conn or sqlite3.connect(db_path)
    try:
        c.execute("""
            CREATE TABLE IF NOT EXISTS column_meta (
                column_name TEXT PRIMARY KEY,
                value_type  TEXT NOT NULL DEFAULT 'text',
                updated_at  TEXT
            )
        """)
        if own:
            c.commit()
    finally:
        if own:
            c.close()


def get_column_types(db_path):
    conn = sqlite3.connect(db_path)
    try:
        return {r[0]: r[1] for r in conn.execute(
            "SELECT column_name, value_type FROM column_meta")}
    except sqlite3.OperationalError:
        return {}
    finally:
        conn.close()


def set_column_types(db_path, mapping, conn=None):
    if not mapping:
        return
    own = conn is None
    c = conn or sqlite3.connect(db_path)
    try:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        c.executemany("""
            INSERT INTO column_meta (column_name, value_type, updated_at) VALUES (?, ?, ?)
            ON CONFLICT(column_name) DO UPDATE SET
                value_type = excluded.value_type, updated_at = excluded.updated_at
        """, [(k, v, now) for k, v in mapping.items()])
        if own:
            c.commit()
    finally:
        if own:
            c.close()


def backfill_column_types(db_path, force=False):
    """为历史数据库补建列类型（启动时只补空缺，不覆盖人工设定）。"""
    conn = sqlite3.connect(db_path)
    try:
        ensure_meta_table(db_path, conn)
        cols = [r[1] for r in conn.execute("PRAGMA table_info(students)")]
        known = {} if force else {
            r[0]: r[1] for r in conn.execute(
                "SELECT column_name, value_type FROM column_meta")}
        result = {}
        for col in cols:
            vals = [r[0] for r in conn.execute(
                "SELECT %s FROM students" % quote_identifier(col))]
            result[col] = infer_column_type(col, vals, known.get(col))
        set_column_types(db_path, result, conn)
        conn.commit()
        return result
    finally:
        conn.close()


def backup_database(db_path):
    folder = os.path.join(os.path.dirname(db_path), "backups")
    os.makedirs(folder, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    dst = os.path.join(folder, "student_database_%s.db" % stamp)
    shutil.copy2(db_path, dst)
    return dst


def prepare_import_dataframe(df, db_path):
    """导入前清洗：列名、表尾统计行、按姓名补学号、跳过常量列。"""
    report = {"actions": [], "unmatched": [], "ambiguous": [], "dropped_rows": 0}

    cleaned, seen = [], {}
    for c in df.columns:
        name = re.sub(r"\s+", " ", str(c).strip().replace("\n", " ").replace("\r", " "))
        if not name:
            name = "未命名列"
        if name in seen:
            seen[name] += 1
            name = "%s_%d" % (name, seen[name])
        else:
            seen[name] = 0
        cleaned.append(name)
    df = df.copy()
    df.columns = cleaned

    n0 = len(df)
    df = df.dropna(how="all").copy()
    if n0 != len(df):
        report["actions"].append("删除全空行 %d 行" % (n0 - len(df)))

    # 有「序号」列时，仅保留正整数序号行，剔除平均分、及格人数、任课教师等表尾
    if "序号" in df.columns:
        def valid_seq(v):
            if _is_missing(v):
                return False
            try:
                f = float(v)
                return f.is_integer() and f > 0
            except (TypeError, ValueError):
                return False
        mask = df["序号"].apply(valid_seq)
        removed = int((~mask).sum())
        if removed:
            report["actions"].append("剔除表尾统计/说明行 %d 行" % removed)
        df = df[mask].copy()

    if "学号" not in df.columns:
        if "姓名" not in df.columns:
            raise ValueError("Excel 中既没有『学号』也没有『姓名』列，无法匹配学生。")
        conn = sqlite3.connect(db_path)
        try:
            rows = conn.execute("SELECT 学号, 姓名 FROM students").fetchall()
        finally:
            conn.close()
        name_map = {}
        for sid, name in rows:
            if name is not None:
                name_map.setdefault(str(name).strip(), []).append(str(sid))
        ids = []
        for n in df["姓名"]:
            key = "" if _is_missing(n) else str(n).strip()
            cand = name_map.get(key, [])
            if len(cand) == 1:
                ids.append(cand[0])
            else:
                ids.append(None)
                if len(cand) > 1 and key not in report["ambiguous"]:
                    report["ambiguous"].append(key)
                elif not cand and key and key not in report["unmatched"]:
                    report["unmatched"].append(key)
        df.insert(0, "学号", ids)
        before = len(df)
        df = df[df["学号"].notna()].copy()
        report["dropped_rows"] = before - len(df)
        report["actions"].append(
            "按姓名匹配库中已有学号：成功 %d 人，库中无此人 %d 人，重名 %d 人"
            % (len(df), len(report["unmatched"]), len(report["ambiguous"])))
    else:
        df["学号"] = df["学号"].apply(normalize_id)
        before = len(df)
        df = df[df["学号"].notna() & (df["学号"].astype(str).str.strip() != "")].copy()
        report["dropped_rows"] = before - len(df)

    drop = []
    if "序号" in df.columns:
        drop.append("序号")
    for c in df.columns:
        if c in ("学号", "姓名") or c in drop:
            continue
        vals = [str(v).strip() for v in df[c]
                if not _is_missing(v) and str(v).strip()]
        if len(vals) > 1 and len(set(vals)) == 1:
            drop.append(c)
    if drop:
        report["actions"].append("跳过无业务含义的列：" + "、".join(drop))
        df = df.drop(columns=drop)

    dup = df["学号"][df["学号"].duplicated()].tolist()
    if dup:
        report["actions"].append("同一学号重复出现，仅保留最后一条：" +
                                 "、".join(map(str, dup)))
        df = df.drop_duplicates(subset=["学号"], keep="last")
    return df.reset_index(drop=True), report


def import_excel_to_db(db_path, file_path, confirm=None, do_backup=True):
    """导入核心逻辑；confirm(title, message) -> bool，传 None 表示脚本模式自动同意。"""
    if os.path.splitext(file_path)[1].lower() == ".xls":
        raise ValueError("暂不支持旧版 .xls，请先在 Excel 中另存为 .xlsx。")
    df = pd.read_excel(file_path, sheet_name=0, engine="openpyxl", dtype=object)
    df, prep = prepare_import_dataframe(df, db_path)
    if df.empty:
        raise ValueError("没有可导入的有效数据行（学号/姓名无法与库中记录匹配）。")

    conn = sqlite3.connect(db_path)
    try:
        ensure_meta_table(db_path, conn)
        existing = [r[1] for r in conn.execute("PRAGMA table_info(students)")]
        known = {r[0]: r[1] for r in conn.execute(
            "SELECT column_name, value_type FROM column_meta")}
    finally:
        conn.close()

    types = {c: infer_column_type(c, df[c].tolist(), known.get(c)) for c in df.columns}
    new_cols = [c for c in df.columns if c not in existing]
    overlap = [c for c in df.columns if c in existing and c != "学号"]

    if confirm is not None:
        lines = ["文件：%s" % os.path.basename(file_path),
                 "有效数据：%d 人，%d 列" % (len(df), len(df.columns)), ""]
        lines += ["清洗动作："] + ["  · " + a for a in prep["actions"]]
        lines += ["", "列类型判定："]
        for c in df.columns:
            lines.append("  · %s [%s]" % (c, types[c]))
        if new_cols:
            lines += ["", "将新增列：" + "、".join(new_cols)]
        if overlap:
            lines += ["", "将覆盖已有列：" + "、".join(overlap)]
        if prep["unmatched"]:
            lines += ["", "库中查无此人（已跳过）：" + "、".join(prep["unmatched"])]
        if prep["ambiguous"]:
            lines += ["", "重名需人工处理（已跳过）：" + "、".join(prep["ambiguous"])]
        lines += ["", "导入前会自动备份数据库。是否继续？"]
        if not confirm("确认导入", "\n".join(lines)):
            return {"cancelled": True}

    backup_path = backup_database(db_path) if do_backup else None
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("BEGIN")
        now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for c in df.columns:
            if c not in existing:
                conn.execute("ALTER TABLE students ADD COLUMN %s TEXT" % quote_identifier(c))
                conn.execute(
                    "INSERT OR IGNORE INTO column_history (column_name, first_upload_time) "
                    "VALUES (?, ?)", (c, now_text))
        cols = list(df.columns)
        cols_sql = ", ".join(quote_identifier(c) for c in cols)
        placeholders = ", ".join("?" for _ in cols)
        update_cols = [c for c in cols if c != "学号"]
        set_clause = ", ".join("%s = ?" % quote_identifier(c) for c in update_cols)
        existing_ids = {r[0] for r in conn.execute("SELECT 学号 FROM students")}

        inserted = updated = 0
        for _, row in df.iterrows():
            sid = row["学号"]
            vals = [normalize_cell_value(row[c], types[c]) for c in cols]
            if sid in existing_ids:
                uvals = [normalize_cell_value(row[c], types[c]) for c in update_cols] + [sid]
                conn.execute("UPDATE students SET %s WHERE 学号 = ?" % set_clause, uvals)
                updated += 1
            else:
                conn.execute("INSERT INTO students (%s) VALUES (%s)" %
                             (cols_sql, placeholders), vals)
                inserted += 1
                existing_ids.add(sid)
        set_column_types(db_path, types, conn)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    return {"cancelled": False, "rows": len(df), "inserted": inserted,
            "updated": updated, "new_columns": new_cols, "overlap": overlap,
            "types": types, "prep": prep, "backup": backup_path}


def _export_date_value(v):
    if _is_missing(v):
        return None
    raw = str(v).strip()
    if not raw:
        return None
    parsed = parse_and_normalize_date(v)
    if isinstance(parsed, str) and _ISO_DATE_RE.match(parsed):
        try:
            return datetime.strptime(parsed, "%Y-%m-%d")
        except ValueError:
            pass
    return raw


def coerce_for_export(df, types):
    """按列类型把库里的 TEXT 还原成 Excel 应有的值类型。"""
    out = pd.DataFrame(index=df.index)
    for c in df.columns:
        t = types.get(c, "text")
        s = df[c]
        if t == "number":
            num = pd.to_numeric(s, errors="coerce")
            res = pd.Series([None] * len(s), index=s.index, dtype=object)
            for i in s.index:
                if pd.notna(num.loc[i]):
                    val = num.loc[i]
                    res.loc[i] = int(val) if float(val).is_integer() else float(val)
                elif not _is_missing(s.loc[i]):
                    res.loc[i] = normalize_text(s.loc[i])
            out[c] = res
        elif t == "date":
            out[c] = s.apply(_export_date_value)
        else:
            out[c] = s.apply(lambda v: None if _is_missing(v) else normalize_text(v))
    return out


def write_excel(df, file_path, types):
    """写出 Excel，并逐列设置单元格格式（文本 @ / 日期 yyyy-mm-dd / 数值 0.##）。"""
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="学生数据")
        ws = writer.sheets["学生数据"]
        ws.freeze_panes = "A2"
        for idx, col in enumerate(df.columns, start=1):
            fmt = {"id": "@", "date": "yyyy-mm-dd", "number": "0.##"}.get(
                types.get(col, "text"))
            if fmt == "yyyy-mm-dd":
                # 只对真正的日期单元格设日期格式；「11月22号」这类无年份文本保持原样
                for r in range(2, len(df) + 2):
                    cell = ws.cell(row=r, column=idx)
                    if isinstance(cell.value, (datetime, date)):
                        cell.number_format = fmt
            elif fmt:
                for r in range(2, len(df) + 2):
                    ws.cell(row=r, column=idx).number_format = fmt
            letter = ws.cell(row=1, column=idx).column_letter
            ws.column_dimensions[letter].width = max(10, min(30, len(str(col)) * 2 + 6))
        ws.auto_filter.ref = ws.dimensions


def _repair_one_value(value):
    """修复旧版污染：1900 年代日期还原成 Excel 序列号，整数去掉 .0 尾巴。"""
    s = str(value)
    if _ISO_DATE_RE.match(s) and s.startswith("190"):
        try:
            n = (datetime.strptime(s, "%Y-%m-%d") - EXCEL_EPOCH).days
            if 1 <= n <= 80000:
                return str(n)
        except ValueError:
            pass
    m = re.fullmatch(r"([+-]?\d+)\.0+", s.strip())
    if m:
        return m.group(1)
    return s


def repair_legacy_values(db_path):
    """扫描并修复旧版误转的数值，返回逐条变更明细。"""
    conn = sqlite3.connect(db_path)
    changes = []
    try:
        ensure_meta_table(db_path, conn)
        cols = [r[1] for r in conn.execute("PRAGMA table_info(students)")]
        for col in cols:
            if any(k in str(col) for k in DATE_KEYWORDS):
                continue
            rows = conn.execute(
                'SELECT 学号, 姓名, %s FROM students' % quote_identifier(col)).fetchall()
            for sid, name, val in rows:
                if val is None:
                    continue
                new = _repair_one_value(val)
                if new != str(val):
                    conn.execute("UPDATE students SET %s = ? WHERE 学号 = ?" %
                                 quote_identifier(col), (new, sid))
                    changes.append((sid, name, col, str(val), new))
        conn.commit()
    finally:
        conn.close()
    return changes


class StudentManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("25级汽修一班学生档案管理系统")
        self.root.geometry("1200x700")
        
        if getattr(sys, 'frozen', False):
            self.base_path = os.path.dirname(sys.executable)
        else:
            self.base_path = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.base_path)
        
        self.db_path = os.path.join(self.base_path, "student_database.db")
        self.init_database()
        
        # 调整此列表即可改变列的默认展示顺序（未列出的字段按字母顺序排列在后）
        self.preferred_order = [
            "学号",
            "姓名",
            "身份证号码",
            "出生日期",
            "电话号码",
            "母亲姓名",
            "母亲电话号码",
            "父亲姓名",
            "父亲电话号码",
            "最后一次谈话日期",
            "谈话内容",
            "红黄蓝情况简述",
            "中考名次",
            "中考成绩",
        ]
        
        # 首页只展示这些基本信息列，其余字段进入学生画像查看（v7）
        self.basic_display_columns = [
            "学号", "姓名", "性别", "电话号码", "宿舍号", "是否走读",
        ]

        # 日期列关键词已收紧（旧版含「谈话」「记录」「登记」，会把「谈话内容」等
        # 纯文本列误判成日期列，导出时整列被清空）。真正的类型以 column_meta 为准。
        self.date_keywords = list(DATE_KEYWORDS)
        self.id_keywords = list(ID_KEYWORDS)
        self.number_keywords = list(NUMBER_KEYWORDS)
        
        self.is_filtered = False
        self.current_filter_desc = ""
        self.column_category_map = {}
        self.column_categories = []
        self.all_columns = []

        self._setup_style()
        self.create_widgets()
        self.refresh_column_list()
        self.refresh_display()  # 启动时直接载入数据，不必再手动点「刷新显示」
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def on_closing(self):
        self.root.destroy()
    
    def _is_date_column(self, col_name):
        """判断是否为日期列：以 column_meta 登记的类型为准，其次按收紧后的关键词。"""
        if not col_name:
            return False
        col = str(col_name)
        known = get_column_types(self.db_path)
        if col in known:
            return known[col] == "date"
        return any(kw in col.lower() for kw in DATE_KEYWORDS)

    @staticmethod
    def parse_and_normalize_date(value):
        """兼容旧调用；实际逻辑已移到模块级严格解析器（不再猜年份 / 不再把数字当日期）。"""
        return parse_and_normalize_date(value)

    @staticmethod
    def normalize_cell_value(value, col_type="text"):
        """按列类型分派；未指定类型时按文本保真处理。"""
        return normalize_cell_value(value, col_type)

    def get_column_types(self):
        return get_column_types(self.db_path)
    
    def sort_columns(self, columns):
        """根据预设顺序对列名进行排序"""
        unique_columns = []
        for col in columns:
            if col not in unique_columns:
                unique_columns.append(col)
        order_map = {col: idx for idx, col in enumerate(self.preferred_order)}
        return sorted(unique_columns, key=lambda c: (order_map.get(c, len(order_map)), c.lower()))
    
    def _display_columns(self, all_columns):
        """首页展示列：只取基本信息列；库里缺这些列时退回全列，保证不白屏。"""
        ordered = self.sort_columns(all_columns)
        basic = [c for c in self.basic_display_columns if c in ordered]
        if "学号" in basic and "姓名" in basic and len(basic) >= 2:
            return basic
        return ordered

    def get_columns_grouped_by_category(self, columns_override=None):
        """按列大类分组（供详情页等功能使用）"""
        if columns_override is None:
            source_columns = [col for col in self.all_columns if col != "学号"]
        else:
            source_columns = [col for col in columns_override if col != "学号"]
        ordered_columns = self.sort_columns(source_columns)
        grouped = {}
        for col in ordered_columns:
            category = self.column_category_map.get(col, "未分类")
            grouped.setdefault(category, []).append(col)
        ordered_categories = sorted(grouped.keys(), key=lambda c: (c == "未分类", c))
        return [(cat, grouped[cat]) for cat in ordered_categories]
    
    def init_database(self):
        """初始化数据库"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS students (
                    学号 TEXT PRIMARY KEY
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS column_history (
                    column_name TEXT PRIMARY KEY,
                    first_upload_time TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS column_categories (
                    category_name TEXT PRIMARY KEY,
                    created_at TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS column_category_map (
                    column_name TEXT PRIMARY KEY,
                    category_name TEXT,
                    FOREIGN KEY (column_name) REFERENCES column_history(column_name),
                    FOREIGN KEY (category_name) REFERENCES column_categories(category_name)
                )
            ''')

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS column_meta (
                    column_name TEXT PRIMARY KEY,
                    value_type  TEXT NOT NULL DEFAULT 'text',
                    updated_at  TEXT
                )
            ''')

            conn.commit()
            conn.close()

            # 为历史数据补建列类型（只补空缺，不覆盖已有设定）
            backfill_column_types(self.db_path)
        except Exception as e:
            messagebox.showerror("数据库错误", f"初始化数据库失败：\n{str(e)}")
    
    def _setup_style(self):
        """全局 UI 风格：微软雅黑、clam 主题、统一配色。"""
        try:
            self.root.configure(bg="#F4F5FA")
        except tk.TclError:
            pass
        self.root.option_add("*Font", ("Microsoft YaHei", 10))

        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Treeview", rowheight=32, font=("Microsoft YaHei", 10),
                        background="#FFFFFF", fieldbackground="#FFFFFF",
                        borderwidth=0)
        style.configure("Treeview.Heading", font=("Microsoft YaHei", 10, "bold"),
                        background="#EEF0F7", foreground="#39406B",
                        padding=(8, 8), relief="flat", borderwidth=0)
        style.map("Treeview.Heading", background=[("active", "#E3E6F2")])
        style.map("Treeview",
                  background=[("selected", "#534AB7")],
                  foreground=[("selected", "#FFFFFF")])

        style.configure("TNotebook", borderwidth=0, background="#F4F5FA")
        style.configure("TNotebook.Tab", font=("Microsoft YaHei", 10),
                        padding=(16, 7))
        style.map("TNotebook.Tab",
                  background=[("selected", "#FFFFFF"), ("!selected", "#E8EAF3")],
                  foreground=[("selected", "#534AB7"), ("!selected", "#6B7280")])

        style.configure("TCombobox", font=("Microsoft YaHei", 10), padding=4)

    def _flat_button(self, master, text, command, bg, hover,
                     fg="white", padx=16, pady=6, font=("Microsoft YaHei", 10)):
        """扁平配色按钮，悬停变色、手型光标。"""
        btn = tk.Button(master, text=text, command=command, bg=bg, fg=fg,
                        font=font, activebackground=hover, activeforeground=fg,
                        relief="flat", bd=0, cursor="hand2", padx=padx, pady=pady)
        btn.bind("<Enter>", lambda e: btn.config(bg=hover))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg))
        return btn

    def create_widgets(self):
        """创建界面组件"""
        # ---- 顶部标题栏（白底卡片：左侧标题，右侧操作按钮）----
        top_frame = tk.Frame(self.root, bg="#FFFFFF", padx=16, pady=12)
        top_frame.pack(fill=tk.X)

        title_box = tk.Frame(top_frame, bg="#FFFFFF")
        title_box.pack(side=tk.LEFT)
        tk.Label(title_box, text="学生档案管理系统",
                 font=("Microsoft YaHei", 16, "bold"),
                 bg="#FFFFFF", fg="#30364B").pack(anchor="w")
        tk.Label(title_box, text="25级汽修一班 · 双击学生行查看学生画像",
                 font=("Microsoft YaHei", 9),
                 bg="#FFFFFF", fg="#8A90A6").pack(anchor="w")

        btn_box = tk.Frame(top_frame, bg="#FFFFFF")
        btn_box.pack(side=tk.RIGHT)
        self._flat_button(btn_box, "导入Excel", self.import_excel,
                          "#534AB7", "#6A5FD0").pack(side=tk.LEFT, padx=3)
        self._flat_button(btn_box, "导出Excel", self.export_excel,
                          "#2F80ED", "#4A97F2").pack(side=tk.LEFT, padx=3)
        self._flat_button(btn_box, "刷新显示", self.refresh_display,
                          "#1A7F4B", "#219663").pack(side=tk.LEFT, padx=3)
        self._flat_button(btn_box, "清空筛选", self.clear_filters,
                          "#6B7280", "#7D8694").pack(side=tk.LEFT, padx=3)
        self._flat_button(btn_box, "删除学生", self.delete_row,
                          "#E05252", "#E96A6A").pack(side=tk.LEFT, padx=3)
        self._flat_button(btn_box, "删除列", self.delete_column,
                          "#C2417A", "#D0558E").pack(side=tk.LEFT, padx=3)
        self._flat_button(btn_box, "列分类", self.open_column_category_manager,
                          "#8E24AA", "#A03BC2").pack(side=tk.LEFT, padx=3)

        # ---- 筛选区（卡片式）----
        filter_frame = tk.LabelFrame(self.root, text=" 筛选 ", pady=8, padx=10,
                                     font=("Microsoft YaHei", 10, "bold"),
                                     fg="#534AB7", bg="#FFFFFF", bd=0)
        filter_frame.pack(fill=tk.X, padx=14, pady=(12, 6))

        tk.Label(filter_frame, text="姓名搜索:", font=("Microsoft YaHei", 10),
                 bg="#FFFFFF", fg="#39406B").grid(row=0, column=0, padx=5, sticky=tk.W)
        self.name_entry = tk.Entry(filter_frame, font=("Microsoft YaHei", 10),
                                   width=20, relief="flat",
                                   highlightthickness=1,
                                   highlightbackground="#D5D9E8",
                                   highlightcolor="#534AB7")
        self.name_entry.grid(row=0, column=1, padx=5, ipady=3)
        self._flat_button(filter_frame, "搜索", self.search_by_name,
                          "#607D8B", "#72879B", padx=12, pady=3).grid(row=0, column=2, padx=5)

        tk.Label(filter_frame, text="列大类:", font=("Microsoft YaHei", 10),
                 bg="#FFFFFF", fg="#39406B").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.filter_category_combo = ttk.Combobox(filter_frame, width=18, state="readonly")
        self.filter_category_combo.grid(row=1, column=1, padx=5, pady=5)
        self.filter_category_combo.bind("<<ComboboxSelected>>", self.on_filter_category_change)
        self._flat_button(filter_frame, "全部列", self.reset_category_filter,
                          "#B0BEC5", "#C2CFD9", padx=12, pady=3).grid(row=1, column=2, padx=5)

        tk.Label(filter_frame, text="按列筛选:", font=("Microsoft YaHei", 10),
                 bg="#FFFFFF", fg="#39406B").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.filter_column_combo = ttk.Combobox(filter_frame, width=18, state="readonly")
        self.filter_column_combo.grid(row=2, column=1, padx=5, pady=5)
        self.filter_column_combo.bind("<<ComboboxSelected>>", self.load_filter_values)

        tk.Label(filter_frame, text="筛选值:", font=("Microsoft YaHei", 10),
                 bg="#FFFFFF", fg="#39406B").grid(row=2, column=3, padx=5, sticky=tk.W)
        self.filter_value_combo = ttk.Combobox(filter_frame, width=18, state="readonly")
        self.filter_value_combo.grid(row=2, column=4, padx=5, pady=5)
        self._flat_button(filter_frame, "应用筛选", self.apply_filter,
                          "#607D8B", "#72879B", padx=12, pady=3).grid(row=2, column=5, padx=5)

        # ---- 数据表格（卡片式容器）----
        display_frame = tk.Frame(self.root, bg="#FFFFFF", bd=0,
                                 highlightthickness=1,
                                 highlightbackground="#E3E6EE")
        display_frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=6)

        self.tree_scroll_x = tk.Scrollbar(display_frame, orient=tk.HORIZONTAL,
                                          troughcolor="#F4F5FA")
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree_scroll_y = tk.Scrollbar(display_frame, orient=tk.VERTICAL,
                                          troughcolor="#F4F5FA")
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(display_frame,
                                 yscrollcommand=self.tree_scroll_y.set,
                                 xscrollcommand=self.tree_scroll_x.set,
                                 selectmode="extended")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(1, 0), pady=(1, 0))
        self.tree.tag_configure("odd", background="#F8F9FC")
        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.bind("<Double-1>", self.on_tree_double_click)

        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        self.status_bar = tk.Label(self.root, text="就绪", bd=0,
                                   bg="#EEF0F7", fg="#39406B",
                                   font=("Microsoft YaHei", 9),
                                   anchor=tk.W, padx=12, pady=5)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def refresh_column_list(self):
        """刷新列名和分类"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall()]
            conn.close()
            
            if "学号" not in columns:
                columns.insert(0, "学号")
            
            columns = self.sort_columns(columns)
            self.all_columns = columns
            self.column_category_map = self.fetch_column_category_map()
            self.column_categories = self.get_all_categories()
            self.refresh_category_filter()
            
            return columns
        except Exception as e:
            messagebox.showerror("错误", f"刷新列名失败：{str(e)}")
            return []
    
    def fetch_column_category_map(self):
        mapping = {}
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT column_name, category_name FROM column_category_map")
            mapping = {row[0]: row[1] for row in cursor.fetchall()}
            conn.close()
        except Exception as e:
            print(f"加载列分类失败: {e}")
        return mapping
    
    def get_all_categories(self):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT category_name FROM column_categories ORDER BY category_name COLLATE NOCASE")
            rows = cursor.fetchall()
            conn.close()
            return [row[0] for row in rows]
        except Exception as e:
            messagebox.showerror("错误", f"加载列大类失败：{str(e)}")
            return []
    
    def refresh_category_filter(self):
        if not hasattr(self, 'filter_category_combo'):
            return
        options = ["全部"] + self.column_categories if self.column_categories else ["全部"]
        current = self.filter_category_combo.get()
        self.filter_category_combo['values'] = options
        if current in options:
            self.filter_category_combo.set(current)
        else:
            self.filter_category_combo.set("全部")
        self.update_filter_column_options()
    
    def update_filter_column_options(self):
        if not hasattr(self, 'filter_column_combo'):
            return
        selected_category = self.filter_category_combo.get()
        columns = [col for col in self.all_columns if col != "学号"]
        if selected_category and selected_category not in ("", "全部"):
            columns = [col for col in columns if self.column_category_map.get(col) == selected_category]
        self.filter_column_combo['values'] = columns
        if columns:
            if self.filter_column_combo.get() not in columns:
                self.filter_column_combo.current(0)
                self.load_filter_values()
        else:
            self.filter_column_combo.set('')
            self.filter_value_combo.set('')
            self.filter_value_combo['values'] = []
    
    def on_filter_category_change(self, event=None):
        self.update_filter_column_options()
    
    def reset_category_filter(self):
        if hasattr(self, 'filter_category_combo'):
            self.filter_category_combo.set("全部")
            self.update_filter_column_options()
    
    def import_excel(self):
        """导入 Excel：统一走数据层，含自动备份、清洗、类型判定与校验摘要。"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not file_path:
            return
        try:
            result = import_excel_to_db(
                self.db_path, file_path,
                confirm=lambda title, msg: messagebox.askyesno(title, msg))
            if result.get("cancelled"):
                return

            self.refresh_column_list()
            self.refresh_display()
            prep = result["prep"]
            lines = [
                "成功导入 %d 人：新增 %d 人，更新 %d 人。" %
                (result["rows"], result["inserted"], result["updated"]),
                "",
                "列类型：" + "、".join("%s[%s]" % (c, t)
                                    for c, t in result["types"].items()),
                "",
                "清洗记录：",
            ]
            lines += ["  · " + a for a in prep["actions"]]
            if prep["unmatched"]:
                lines.append("\n库中查无此人（已跳过）：" + "、".join(prep["unmatched"]))
            if prep["ambiguous"]:
                lines.append("\n重名需人工处理（已跳过）：" + "、".join(prep["ambiguous"]))
            if result["backup"]:
                lines.append("\n数据库已备份到：\n" + result["backup"])
            messagebox.showinfo("导入完成", "\n".join(lines))
            self.status_bar.config(
                text="导入完成: %d 人（新增 %d / 更新 %d）" %
                     (result["rows"], result["inserted"], result["updated"]))
        except Exception as e:
            messagebox.showerror("错误", f"导入失败：\n{str(e)}\n\n{traceback.format_exc()}")
    
    def delete_row(self):
        """删除选中的学生记录"""
        selected_items = self.tree.selection()
        
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要删除的学生记录")
            return
        
        student_ids = []
        for item in selected_items:
            values = self.tree.item(item)['values']
            if values:
                student_ids.append(str(values[0]))
        
        if not student_ids:
            return
        
        msg = f"确定要删除以下 {len(student_ids)} 个学生的记录吗？\n\n"
        msg += "学号: " + ", ".join(student_ids[:10])
        if len(student_ids) > 10:
            msg += f"\n...等共{len(student_ids)}个学生"
        msg += "\n\n此操作不可恢复！"
        
        if not messagebox.askyesno("确认删除", msg):
            return
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            for student_id in student_ids:
                cursor.execute("DELETE FROM students WHERE 学号 = ?", (student_id,))
            
            conn.commit()
            conn.close()
            
            self.refresh_display()
            
            messagebox.showinfo("成功", f"成功删除 {len(student_ids)} 条学生记录")
            self.status_bar.config(text=f"删除完成: {len(student_ids)} 条记录")
            
        except Exception as e:
            messagebox.showerror("错误", f"删除失败：{str(e)}")
    
    def delete_column(self):
        """删除指定的列"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall() if row[1] != "学号"]
            conn.close()
            
            columns = self.sort_columns(columns)
            
            if not columns:
                messagebox.showwarning("警告", "没有可删除的列（学号列不能删除）")
                return
            
            delete_window = tk.Toplevel(self.root)
            delete_window.withdraw()
            delete_window.title("选择要删除的列")
            delete_window.geometry("450x550")
            delete_window.transient(self.root)
            delete_window.grab_set()
            
            title_frame = tk.Frame(delete_window)
            title_frame.pack(fill=tk.X, pady=10)
            
            tk.Label(title_frame, text="请选择要删除的列：", 
                    font=("Arial", 11, "bold"), fg="#F44336").pack()
            
            tk.Label(title_frame, text="警告：删除列将永久删除该列的所有数据！", 
                    font=("Arial", 9), fg="red").pack(pady=5)
            
            middle_frame = tk.Frame(delete_window)
            middle_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            canvas = tk.Canvas(middle_frame, bg="white")
            scrollbar = tk.Scrollbar(middle_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg="white")
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            check_vars = {}
            
            for col in columns:
                var = tk.BooleanVar(value=False)
                check_vars[col] = var
                cb = tk.Checkbutton(scrollable_frame, text=col, variable=var, 
                                   font=("Arial", 10), bg="white")
                cb.pack(anchor=tk.W, padx=20, pady=3)
            
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            button_frame = tk.Frame(delete_window)
            button_frame.pack(side=tk.BOTTOM, pady=15)
            
            def do_delete():
                selected_columns = [col for col, var in check_vars.items() if var.get()]
                
                if not selected_columns:
                    messagebox.showwarning("警告", "请至少选择一列")
                    return
                
                confirm_msg = f"确定要删除以下列吗？\n\n"
                confirm_msg += ", ".join(selected_columns)
                confirm_msg += "\n\n此操作将永久删除这些列的所有数据，不可恢复！"
                
                if not messagebox.askyesno("最终确认", confirm_msg):
                    return
                
                try:
                    conn_in = sqlite3.connect(self.db_path)
                    cursor_in = conn_in.cursor()
                    
                    cursor_in.execute("PRAGMA table_info(students)")
                    all_columns = [row[1] for row in cursor_in.fetchall()]
                    keep_columns = [col for col in all_columns if col not in selected_columns]
                    
                    columns_def = ", ".join([f'"{col}" TEXT' for col in keep_columns])
                    cursor_in.execute(f'CREATE TABLE students_new ({columns_def}, PRIMARY KEY (学号))')
                    
                    columns_str = ", ".join([f'"{col}"' for col in keep_columns])
                    cursor_in.execute(f'INSERT INTO students_new SELECT {columns_str} FROM students')
                    
                    cursor_in.execute('DROP TABLE students')
                    cursor_in.execute('ALTER TABLE students_new RENAME TO students')
                    
                    for col in selected_columns:
                        cursor_in.execute('DELETE FROM column_history WHERE column_name = ?', (col,))
                        cursor_in.execute('DELETE FROM column_category_map WHERE column_name = ?', (col,))
                        cursor_in.execute('DELETE FROM column_meta WHERE column_name = ?', (col,))
                    
                    conn_in.commit()
                    conn_in.close()
                    
                    self.refresh_column_list()
                    self.refresh_display()
                    
                    messagebox.showinfo("成功", f"成功删除 {len(selected_columns)} 列")
                    delete_window.destroy()
                    
                except Exception as err:
                    messagebox.showerror("错误", f"删除列失败：\n{str(err)}")
            
            tk.Button(button_frame, text="确定删除", command=do_delete,
                     bg="#F44336", fg="white", font=("Arial", 11, "bold"), 
                     padx=30, pady=8).pack(side=tk.LEFT, padx=10)
            tk.Button(button_frame, text="取消", command=delete_window.destroy,
                     bg="#9E9E9E", fg="white", font=("Arial", 11), 
                     padx=30, pady=8).pack(side=tk.LEFT, padx=10)
            
            center_window_relative(delete_window, self.root)
            
        except Exception as e:
            messagebox.showerror("错误", f"删除列功能错误：\n{str(e)}")
    
    def export_excel(self):
        """导出Excel文件（已增强日期列自动格式化）"""
        try:
            columns = self.refresh_column_list()
            
            if not columns or columns == ["学号"]:
                messagebox.showwarning("警告", "没有可导出的数据")
                return
            
            export_type = "all"
            
            if self.is_filtered:
                msg = f"当前显示的是筛选后的数据（{self.current_filter_desc}）\n\n"
                msg += "请选择导出范围：\n\n"
                msg += "• 点击【是】：仅导出当前筛选的学生\n"
                msg += "• 点击【否】：导出所有学生数据"
                
                result = messagebox.askyesnocancel("选择导出范围", msg)
                if result is None:
                    return
                elif result:
                    export_type = "filtered"
                else:
                    export_type = "all"
            
            export_window = tk.Toplevel(self.root)
            export_window.withdraw()
            export_window.title("选择导出列")
            export_window.geometry("450x650")
            export_window.transient(self.root)
            export_window.grab_set()
            
            title_frame = tk.Frame(export_window)
            title_frame.pack(fill=tk.X, pady=10)
            
            title_text = "请选择要导出的列："
            if export_type == "filtered":
                title_text += f"\n（将导出筛选结果：{self.current_filter_desc}）"
            else:
                title_text += "\n（将导出所有学生数据）"
            
            tk.Label(title_frame, text=title_text, 
                    font=("Arial", 11, "bold"), fg="#2196F3").pack()
            
            middle_frame = tk.Frame(export_window)
            middle_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            canvas = tk.Canvas(middle_frame, bg="white")
            scrollbar = tk.Scrollbar(middle_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg="white")
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            check_vars = {}
            default_selected = {"学号", "姓名", "身份证号", "身份证号码"}

            var = tk.BooleanVar(value=True)
            check_vars["学号"] = var
            tk.Checkbutton(scrollable_frame, text="学号 ✓", variable=var,
                           font=("Arial", 10, "bold"), state="disabled",
                           fg="#4CAF50", bg="white").pack(anchor=tk.W, padx=20, pady=3)

            column_groups = self.get_columns_grouped_by_category(columns_override=columns)

            if not column_groups:
                column_groups = [("未分类", [col for col in columns if col != "学号"])]

            for category, cols in column_groups:
                if not cols:
                    continue

                group_container = tk.Frame(scrollable_frame, bg="white", bd=1, relief=tk.GROOVE)
                group_container.pack(fill=tk.X, padx=10, pady=6)

                header_frame = tk.Frame(group_container, bg="#ECEFF1")
                header_frame.pack(fill=tk.X)

                icon_label = tk.Label(header_frame, text="▼", font=("Arial", 10, "bold"), bg="#ECEFF1")
                icon_label.pack(side=tk.LEFT, padx=6, pady=6)

                title_label = tk.Label(header_frame, text=category, font=("Arial", 11, "bold"),
                                       bg="#ECEFF1", anchor="w")
                title_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2, pady=6)

                content_frame = tk.Frame(group_container, bg="white")
                content_frame.pack(fill=tk.X, padx=12, pady=6)

                is_expanded = tk.BooleanVar(value=True)

                def toggle_content(event=None, content=content_frame,
                                   icon=icon_label, state=is_expanded, container=group_container):
                    if state.get():
                        content.pack_forget()
                        icon.config(text="▶")
                        state.set(False)
                    else:
                        content.pack(fill=tk.X, padx=12, pady=6)
                        icon.config(text="▼")
                        state.set(True)

                toggle_btn = tk.Button(header_frame, text="▼", width=3, command=toggle_content)
                toggle_btn.pack(side=tk.RIGHT, padx=6, pady=4)

                header_frame.bind("<Button-1>", toggle_content)
                title_label.bind("<Button-1>", toggle_content)
                icon_label.bind("<Button-1>", toggle_content)

                for col in cols:
                    should_select = col in default_selected
                    var = tk.BooleanVar(value=should_select)
                    check_vars[col] = var
                    text = f"{col} (推荐)" if should_select else col
                    cb = tk.Checkbutton(content_frame, text=text, variable=var,
                                        font=("Arial", 10), bg="white")
                    if should_select:
                        cb.config(fg="#2196F3")
                    cb.pack(anchor=tk.W, fill=tk.X, padx=8, pady=3)
            
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            button_frame = tk.Frame(export_window)
            button_frame.pack(side=tk.BOTTOM, pady=15)
            
            def do_export():
                try:
                    selected_columns = [col for col, var in check_vars.items() if var.get()]
                    if not selected_columns:
                        messagebox.showwarning("警告", "请至少选择一列")
                        return

                    default_filename = "学生数据_筛选结果.xlsx" if export_type == "filtered" else "学生数据_全部.xlsx"
                    file_path = filedialog.asksaveasfilename(
                        title="保存Excel文件",
                        defaultextension=".xlsx",
                        initialfile=default_filename,
                        filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
                    )
                    if not file_path:
                        return

                    types = get_column_types(self.db_path)
                    columns_str = ", ".join(quote_identifier(c) for c in selected_columns)
                    conn = sqlite3.connect(self.db_path)
                    try:
                        if export_type == "filtered":
                            # 只取学号回库查询，避免 Treeview 把数值/日期统统转成字符串
                            tree_cols = list(self.tree['columns'])
                            id_idx = tree_cols.index("学号") if "学号" in tree_cols else 0
                            ids = []
                            for item in self.tree.get_children():
                                values = self.tree.item(item)['values']
                                if values and len(values) > id_idx:
                                    sid = str(values[id_idx]).strip()
                                    if sid:
                                        ids.append(sid)
                            if ids:
                                ph = ", ".join("?" for _ in ids)
                                df = pd.read_sql_query(
                                    "SELECT %s FROM students WHERE 学号 IN (%s)" %
                                    (columns_str, ph), conn, params=ids)
                            else:
                                df = pd.DataFrame(columns=selected_columns)
                        else:
                            df = pd.read_sql_query(
                                "SELECT %s FROM students" % columns_str, conn)
                    finally:
                        conn.close()

                    # 按列类型还原：数值是数值，日期是 yyyy-mm-dd，编号/身份证是文本
                    df = coerce_for_export(df, types)
                    write_excel(df, file_path, types)

                    export_desc = f"筛选结果：{self.current_filter_desc}" if export_type == "filtered" else "全部数据"
                    messagebox.showinfo("成功", f"成功导出 {len(df)} 条记录（{export_desc}）到：\n{file_path}")
                    export_window.destroy()
                    self.status_bar.config(text=f"导出完成: {len(df)} 条记录")
                except Exception as e:
                    messagebox.showerror("错误", f"导出失败：\n{str(e)}\n\n{traceback.format_exc()}")
            
            tk.Button(button_frame, text="确定导出", command=do_export,
                     bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), 
                     padx=40, pady=8).pack(side=tk.LEFT, padx=10)
            tk.Button(button_frame, text="取消", command=export_window.destroy,
                     bg="#F44336", fg="white", font=("Arial", 11), 
                     padx=40, pady=8).pack(side=tk.LEFT, padx=10)
            
            center_window_relative(export_window, self.root)
        
        except Exception as e:
            messagebox.showerror("错误", f"导出功能错误：\n{str(e)}")
    
    def refresh_display(self):
        """刷新显示所有数据"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall()]
            columns = self._display_columns(columns)
            
            if not columns:
                conn.close()
                self.status_bar.config(text="暂无数据")
                self.is_filtered = False
                return
            
            self.tree['columns'] = columns
            self.tree['show'] = 'headings'
            
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor=tk.W)
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            columns_str = ", ".join([f'"{col}"' for col in columns])
            cursor.execute(f"SELECT {columns_str} FROM students")
            rows = cursor.fetchall()
            
            for i, row in enumerate(rows):
                self.tree.insert('', tk.END, values=row,
                                 tags=("odd" if i % 2 else "even",))
            
            conn.close()
            self.status_bar.config(
                text=f"显示 {len(rows)} 名学生（基本信息）· 双击学生行查看学生画像")
            self.is_filtered = False
            self.current_filter_desc = ""
            
        except Exception as e:
            messagebox.showerror("错误", f"刷新显示失败：{str(e)}")
    
    def search_by_name(self):
        """按姓名搜索"""
        name = self.name_entry.get().strip()
        
        if not name:
            messagebox.showwarning("警告", "请输入姓名")
            return
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall()]
            columns = self._display_columns(columns)
            
            if "姓名" not in columns:
                conn.close()
                messagebox.showwarning("提示", "数据库中没有'姓名'列，无法按姓名搜索")
                return
            
            self.tree['columns'] = columns
            self.tree['show'] = 'headings'
            
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor=tk.W)
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            columns_str = ", ".join([f'"{col}"' for col in columns])
            cursor.execute(f'SELECT {columns_str} FROM students WHERE 姓名 LIKE ?', (f'%{name}%',))
            rows = cursor.fetchall()
            
            for i, row in enumerate(rows):
                self.tree.insert('', tk.END, values=row,
                                 tags=("odd" if i % 2 else "even",))
            
            conn.close()
            
            if len(rows) == 0:
                messagebox.showinfo("提示", f"未找到姓名包含 '{name}' 的学生")
                self.is_filtered = False
                self.current_filter_desc = ""
            else:
                self.status_bar.config(text=f"找到 {len(rows)} 条匹配记录")
                self.is_filtered = True
                self.current_filter_desc = f"姓名包含'{name}'"
            
        except Exception as e:
            messagebox.showerror("错误", f"搜索失败：{str(e)}")
    
    def load_filter_values(self, event=None):
        column = self.filter_column_combo.get()
        
        if not column:
            return
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(f'SELECT DISTINCT "{column}" FROM students WHERE "{column}" IS NOT NULL AND "{column}" != ""')
            values = sorted([row[0] for row in cursor.fetchall()])
            
            conn.close()
            
            self.filter_value_combo['values'] = values
            
            if values:
                self.filter_value_combo.current(0)
            
        except Exception as e:
            messagebox.showerror("错误", f"加载筛选值失败：{str(e)}")
    
    def apply_filter(self):
        column = self.filter_column_combo.get()
        value = self.filter_value_combo.get()
        
        if not column or not value:
            messagebox.showwarning("警告", "请选择筛选列和筛选值")
            return
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall()]
            columns = self._display_columns(columns)
            
            self.tree['columns'] = columns
            self.tree['show'] = 'headings'
            
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor=tk.W)
            
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            columns_str = ", ".join([f'"{col}"' for col in columns])
            cursor.execute(f'SELECT {columns_str} FROM students WHERE "{column}" = ?', (value,))
            rows = cursor.fetchall()
            
            for i, row in enumerate(rows):
                self.tree.insert('', tk.END, values=row,
                                 tags=("odd" if i % 2 else "even",))
            
            conn.close()
            
            self.status_bar.config(text=f"筛选结果: {len(rows)} 条记录 ({column} = {value})")
            self.is_filtered = True
            self.current_filter_desc = f"{column}={value}"
            
        except Exception as e:
            messagebox.showerror("错误", f"筛选失败：{str(e)}")
    
    def clear_filters(self):
        """清空筛选条件并显示所有数据"""
        self.name_entry.delete(0, tk.END)
        self.filter_column_combo.set('')
        self.filter_value_combo.set('')
        self.filter_value_combo['values'] = []
        if hasattr(self, 'filter_category_combo'):
            self.filter_category_combo.set("全部")
            self.update_filter_column_options()
        self.is_filtered = False
        self.current_filter_desc = ""
        self.refresh_display()
    
    def on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        values = self.tree.item(item_id)['values']
        if not values:
            return
        columns = self.tree['columns']
        if "学号" in columns:
            student_idx = columns.index("学号")
            student_id = values[student_idx]
        else:
            student_id = values[0]
        if not student_id:
            return
        StudentDetailWindow(self, str(student_id))
    
    def open_column_category_manager(self):
        self.refresh_column_list()
        ColumnCategoryManager(self)


class ColumnCategoryManager:
    def __init__(self, parent):
        self.parent = parent
        self.window = tk.Toplevel(parent.root)
        self.window.withdraw()
        self.window.title("列分类管理")
        self.window.geometry("700x600")
        self.window.transient(parent.root)
        self.window.grab_set()
        
        self.column_vars = {}
        self.column_combos = []
        
        self.build_ui()
        self.reload_categories()
        self.load_columns()
        center_window_relative(self.window, self.parent.root)
    
    def build_ui(self):
        main = tk.Frame(self.window, padx=10, pady=10)
        main.pack(fill=tk.BOTH, expand=True)
        
        category_frame = tk.LabelFrame(main, text="一级类目", padx=10, pady=10)
        category_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.category_list = tk.Listbox(category_frame, height=5)
        self.category_list.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        btn_frame = tk.Frame(category_frame)
        btn_frame.pack(side=tk.LEFT, padx=10)
        
        tk.Button(btn_frame, text="新增", width=8, command=self.add_category).pack(pady=2)
        tk.Button(btn_frame, text="重命名", width=8, command=self.rename_category).pack(pady=2)
        tk.Button(btn_frame, text="删除", width=8, command=self.delete_category).pack(pady=2)
        
        mapping_frame = tk.LabelFrame(main, text="列归属设置（学号默认不分类）", padx=10, pady=10)
        mapping_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(mapping_frame)
        scrollbar = tk.Scrollbar(mapping_frame, orient="vertical", command=canvas.yview)
        self.form_frame = tk.Frame(canvas)
        
        self.form_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.form_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        tk.Button(main, text="保存映射", bg="#4CAF50", fg="white", font=("Arial", 11),
                  command=self.save_mappings).pack(fill=tk.X, pady=(10, 0))
    
    def reload_categories(self):
        self.categories = self.parent.get_all_categories()
        self.category_list.delete(0, tk.END)
        for cat in self.categories:
            self.category_list.insert(tk.END, cat)
        self.update_combo_options()
    
    def update_combo_options(self):
        options = ["未分类"] + self.categories if self.categories else ["未分类"]
        for combo in self.column_combos:
            combo['values'] = options
            if combo.get() not in options:
                combo.set("未分类")
    
    def load_columns(self):
        for widget in self.form_frame.winfo_children():
            widget.destroy()
        self.column_vars.clear()
        self.column_combos.clear()
        
        columns = [col for col in self.parent.all_columns if col != "学号"]
        columns = self.parent.sort_columns(columns)
        for idx, col in enumerate(columns):
            var = tk.StringVar(value=self.parent.column_category_map.get(col, "未分类"))
            lbl = tk.Label(self.form_frame, text=col + ":", anchor="w")
            lbl.grid(row=idx, column=0, sticky="w", padx=5, pady=4)
            combo = ttk.Combobox(self.form_frame, textvariable=var, state="readonly", width=25)
            combo.grid(row=idx, column=1, sticky="ew", padx=5, pady=4)
            self.form_frame.grid_columnconfigure(1, weight=1)
            self.column_vars[col] = var
            self.column_combos.append(combo)
        
        self.update_combo_options()
    
    def add_category(self):
        name = simpledialog.askstring("新增类目", "请输入类目名称：", parent=self.window)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name in self.categories:
            messagebox.showwarning("提示", "该类目已存在", parent=self.window)
            return
        conn = sqlite3.connect(self.parent.db_path)
        cursor = conn.cursor()
        cursor.execute(
            'INSERT INTO column_categories (category_name, created_at) VALUES (?, ?)',
            (name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()
        conn.close()
        self.reload_categories()
        self.parent.refresh_column_list()
    
    def rename_category(self):
        selection = self.category_list.curselection()
        if not selection:
            messagebox.showwarning("提示", "请选择要重命名的类目", parent=self.window)
            return
        old_name = self.category_list.get(selection[0])
        new_name = simpledialog.askstring("重命名", "新的类目名称：", parent=self.window, initialvalue=old_name)
        if not new_name:
            return
        new_name = new_name.strip()
        if not new_name or new_name == old_name:
            return
        if new_name in self.categories:
            messagebox.showwarning("提示", "该名称已存在", parent=self.window)
            return
        conn = sqlite3.connect(self.parent.db_path)
        cursor = conn.cursor()
        cursor.execute('UPDATE column_categories SET category_name = ? WHERE category_name = ?', (new_name, old_name))
        cursor.execute('UPDATE column_category_map SET category_name = ? WHERE category_name = ?', (new_name, old_name))
        conn.commit()
        conn.close()
        self.reload_categories()
        self.parent.refresh_column_list()
    
    def delete_category(self):
        selection = self.category_list.curselection()
        if not selection:
            messagebox.showwarning("提示", "请选择要删除的类目", parent=self.window)
            return
        name = self.category_list.get(selection[0])
        if not messagebox.askyesno("确认", f"确定删除类目“{name}”并清除其下列的归属吗？", parent=self.window):
            return
        conn = sqlite3.connect(self.parent.db_path)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM column_categories WHERE category_name = ?', (name,))
        cursor.execute('DELETE FROM column_category_map WHERE column_name IN (SELECT column_name FROM column_category_map WHERE category_name = ?)', (name,))
        cursor.execute('DELETE FROM column_category_map WHERE category_name = ?', (name,))
        conn.commit()
        conn.close()
        self.reload_categories()
        self.parent.refresh_column_list()
    
    def save_mappings(self):
        try:
            conn = sqlite3.connect(self.parent.db_path)
            cursor = conn.cursor()
            for col, var in self.column_vars.items():
                category = var.get()
                if not category or category == "未分类":
                    cursor.execute('DELETE FROM column_category_map WHERE column_name = ?', (col,))
                else:
                    cursor.execute('''
                        INSERT INTO column_category_map (column_name, category_name)
                        VALUES (?, ?)
                        ON CONFLICT(column_name) DO UPDATE SET category_name = excluded.category_name
                    ''', (col, category))
            conn.commit()
            conn.close()
            messagebox.showinfo("成功", "列分类设置已保存", parent=self.window)
            self.parent.refresh_column_list()
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}", parent=self.window)


class StudentDetailWindow:
    """学生画像窗口：顶部画像卡 + 分区详情；单科成绩归入所属考试批次之下。"""

    # 「2510月考成绩 / 2610期中语文」→ 批次 2510月考 / 2610期中，子项 成绩 / 语文
    EXAM_RE = re.compile(r"^(\d{4})\s*(月考|期中|期末)(.+)$")
    # 分数直接归在考试名下的子项（视为该次考试的总分/名次）
    EXAM_TOTAL_SUBS = ("总分", "成绩", "名次")

    def __init__(self, parent, student_id):
        self.parent = parent
        self.student_id = student_id
        self.window = tk.Toplevel(parent.root)
        self.window.withdraw()
        self.window.title("学生画像")
        self.window.geometry("980x800")
        self.window.configure(bg="#F4F5FA")
        self.window.transient(parent.root)
        self.window.grab_set()

        self.entry_vars = {}

        if not self.load_student_data():
            self.window.destroy()
            return

        self.column_groups = self.parent.get_columns_grouped_by_category(self.columns)
        name = str(self.data.get("姓名") or self.student_id)
        self.window.title(f"学生画像 - {name}")
        self.build_profile()
        center_window_relative(self.window, self.parent.root)

    def load_student_data(self):
        try:
            conn = sqlite3.connect(self.parent.db_path)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(students)")
            columns = [row[1] for row in cursor.fetchall()]
            columns = self.parent.sort_columns(columns)
            self.columns = columns
            columns_str = ", ".join([f'"{col}"' for col in self.columns])
            cursor.execute(f'SELECT {columns_str} FROM students WHERE 学号 = ?', (self.student_id,))
            row = cursor.fetchone()
            conn.close()
            if not row:
                messagebox.showwarning("提示", "该学生记录不存在", parent=self.window)
                return False
            self.data = dict(zip(self.columns, row))
            return True
        except Exception as e:
            messagebox.showerror("错误", f"加载数据失败：{str(e)}", parent=self.window)
            return False

    # ------------------------------------------------------------------
    # 顶部画像卡
    # ------------------------------------------------------------------
    def _profile_chip_fields(self):
        return [("性别", "性别"), ("生日", "生日"), ("宿舍号", "宿舍号"),
                ("电话", "电话号码"), ("走读", "是否走读"),
                ("红黄蓝", "红黄蓝情况简述"), ("学生会", "是否学生会")]

    def build_profile_card(self, container):
        card = tk.Frame(container, bg="#FFFFFF", bd=0,
                        highlightthickness=1,
                        highlightbackground="#E3E6EE")
        card.pack(fill=tk.X)

        # 左侧紫色竖条作装饰
        tk.Frame(card, bg="#534AB7", width=6).pack(side=tk.LEFT, fill=tk.Y)

        inner = tk.Frame(card, bg="#FFFFFF", padx=16, pady=12)
        inner.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        name = str(self.data.get("姓名") or self.student_id)
        head = tk.Frame(inner, bg="#FFFFFF")
        head.pack(fill=tk.X)
        tk.Label(head, text=name, font=("Microsoft YaHei", 20, "bold"),
                 bg="#FFFFFF", fg="#30364B").pack(side=tk.LEFT)
        tk.Label(head, text="   学号 " + str(self.student_id),
                 font=("Microsoft YaHei", 11), bg="#FFFFFF",
                 fg="#8A90A6").pack(side=tk.LEFT, pady=(9, 0))

        chip_row = tk.Frame(inner, bg="#FFFFFF")
        chip_row.pack(fill=tk.X, pady=(10, 2))
        chip_colors = {"红黄蓝": ("#FDECEC", "#B4433F"),
                       "走读": ("#E7F3EC", "#1A7F4B")}
        for label, col in self._profile_chip_fields():
            value = self.data.get(col)
            text = "—" if value in (None, "") else str(value)
            bg, fg = chip_colors.get(label, ("#EEF0F7", "#39406B"))
            chip = tk.Label(chip_row, text=f" {label}  {text} ",
                            bg=bg, fg=fg, padx=10, pady=4)
            chip.pack(side=tk.LEFT, padx=(0, 8))

    # ------------------------------------------------------------------
    # 学业画像：考试批次 → 科目 两级
    # ------------------------------------------------------------------
    def _parse_exam_columns(self):
        """把列名解析成「考试批次 → 科目/名次」两级结构；单科只属于本次考试。"""
        exams = {}
        others = []   # 中考成绩/中考名次等非编号考试列
        for col in self.columns:
            if col == "学号":
                continue
            m = self.EXAM_RE.match(col)
            if m:
                batch = m.group(1) + m.group(2)
                sub = m.group(3).strip() or "总分"
                exams.setdefault(batch, []).append((sub, col))
            elif any(k in col for k in ("成绩", "名次", "得分", "总分")):
                others.append(col)
        # 同一批次内：总分/成绩在前、名次随后，单科按名称排序
        result = []
        for batch in sorted(exams.keys()):
            subs = exams[batch]
            subs.sort(key=lambda t: (t[0] not in self.EXAM_TOTAL_SUBS,
                                     self.EXAM_TOTAL_SUBS.index(t[0])
                                     if t[0] in self.EXAM_TOTAL_SUBS else 99,
                                     t[0]))
            result.append((batch, subs))
        return result, sorted(others)

    def _fmt_score(self, value):
        if value in (None, ""):
            return "—"
        s = str(value).strip()
        try:
            f = float(s)
            return str(int(f)) if f.is_integer() else ("%g" % f)
        except ValueError:
            return s

    def build_academic_tab(self, notebook):
        frame = tk.Frame(notebook)
        notebook.add(frame, text="学业画像")  # 先添加，即位于第一个标签位

        # ---- 左侧：考试批次明细树 ----
        left = tk.Frame(frame)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(6, 3), pady=6)

        tree = ttk.Treeview(left, columns=("item", "value"), show="headings",
                            height=18)
        tree.heading("item", text="考试 / 科目")
        tree.heading("value", text="分数")
        tree.column("item", width=170, anchor=tk.W)
        tree.column("value", width=60, anchor=tk.CENTER)
        ys = tk.Scrollbar(left, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=ys.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ys.pack(side=tk.RIGHT, fill=tk.Y)

        tree.tag_configure("total", font=("Microsoft YaHei", 10, "bold"), foreground="#1F2430")
        tree.tag_configure("batch", font=("Microsoft YaHei", 10, "bold"), foreground="#534AB7")
        tree.tag_configure("empty", foreground="#9AA0B5")

        exam_groups, other_cols = self._parse_exam_columns()
        if not exam_groups and not other_cols:
            tree.insert("", tk.END, values=("暂无成绩数据", "—"), tags=("empty",))

        def add_batch(batch_label, subs):
            batch_item = tree.insert("", tk.END, open=True,
                                     values=(batch_label, ""), tags=("batch",))
            for sub, col in subs:
                val = self._fmt_score(self.data.get(col))
                empty = val == "—"
                if sub in self.EXAM_TOTAL_SUBS:
                    # 总分/成绩/名次属于考试本身，不缩进、加粗
                    tag = ("empty",) if empty else ("total",)
                    name = sub
                else:
                    # 单科成绩：本次考试的下一层级，缩进展示
                    tag = ("empty",) if empty else ()
                    name = "  ├ " + sub
                tree.insert(batch_item, tk.END, values=(name, val), tags=tag)

        for batch, subs in exam_groups:
            add_batch(batch, subs)

        if other_cols:
            batch_item = tree.insert("", tk.END, open=True,
                                     values=("中考（入校成绩）", ""), tags=("batch",))
            for col in other_cols:
                val = self._fmt_score(self.data.get(col))
                tag = ("empty",) if val == "—" else ("total",)
                tree.insert(batch_item, tk.END, values=(col, val), tags=tag)

        # ---- 右侧：历史成绩 / 名次 / 单科趋势折线图 ----
        right = tk.Frame(frame)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(3, 6), pady=6)
        self._build_trend_charts(right)

    # ------------------------------------------------------------------
    # 折线图：历次考试总分、名次、单科成绩
    # ------------------------------------------------------------------
    @staticmethod
    def _to_number(value):
        if value in (None, ""):
            return None
        try:
            f = float(str(value).strip())
            return int(f) if f.is_integer() else f
        except (TypeError, ValueError):
            return None

    def _exam_series(self):
        """提取历次考试序列：批次、总分、名次、各单科分数（含中考作为最早一点）。"""
        batches, totals, ranks = [], [], []
        subjects = {}
        groups, other_cols = self._parse_exam_columns()

        if other_cols:
            # 中考并入序列最前端（中考成绩→总分，中考名次→名次）
            batches.append("中考")
            total = rank = None
            for col in other_cols:
                val = self._to_number(self.data.get(col))
                if "名次" in col:
                    rank = val
                elif "成绩" in col and total is None:
                    total = val
            totals.append(total)
            ranks.append(rank)

        for batch, subs in groups:
            batches.append(batch)
            total = rank = None
            for sub, col in subs:
                val = self._to_number(self.data.get(col))
                if sub == "名次":
                    rank = val
                elif sub in ("总分", "成绩") and total is None:
                    total = val
                else:
                    subjects.setdefault(sub, {})[batch] = val
            totals.append(total)
            ranks.append(rank)
        return batches, totals, ranks, subjects

    def _build_trend_charts(self, parent):
        try:
            import matplotlib
            matplotlib.use("Agg")
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        except ImportError:
            tk.Label(parent, text="未安装 matplotlib，暂无法绘制趋势图。\n"
                                  "安装：pip install matplotlib",
                     fg="#B45309", justify="left").pack(pady=30)
            return

        matplotlib.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
        matplotlib.rcParams["axes.unicode_minus"] = False

        batches, totals, ranks, subjects = self._exam_series()
        if not batches:
            tk.Label(parent, text="暂无成绩数据", fg="#757575").pack(pady=30)
            return

        # 单科选择器
        subj_names = sorted(subjects.keys())
        top = tk.Frame(parent)
        top.pack(fill=tk.X)
        tk.Label(top, text="单科成绩趋势：", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        subj_var = tk.StringVar(value=subj_names[0] if subj_names else "")
        combo = ttk.Combobox(top, textvariable=subj_var, width=12,
                             values=subj_names,
                             state="readonly" if subj_names else "disabled")
        combo.pack(side=tk.LEFT, padx=4)

        fig = Figure(figsize=(5.6, 6.4), dpi=100)
        ax1, ax2, ax3 = fig.subplots(3, 1)
        fig.subplots_adjust(left=0.10, right=0.96, top=0.95,
                            bottom=0.08, hspace=0.62)
        xs = list(range(len(batches)))

        def style_axis(ax):
            ax.grid(True, linestyle="--", alpha=0.35)
            ax.set_xticks(xs)
            ax.set_xticklabels(batches, fontsize=7, rotation=30, ha="right")
            ax.tick_params(labelsize=8)

        def annotate(ax, xsv, ysv, color="#534AB7"):
            for x, y in zip(xsv, ysv):
                ax.annotate(str(y), (x, y), textcoords="offset points",
                            xytext=(0, 7), fontsize=7, ha="center", color=color)

        def draw_line(ax, series, title, color, invert=False):
            pts = [(x, v) for x, v in zip(xs, series) if v is not None]
            if pts:
                xsv, ysv = [p[0] for p in pts], [p[1] for p in pts]
                ax.plot(xsv, ysv, marker="o", color=color, linewidth=1.8)
                annotate(ax, xsv, ysv, color)
            ax.set_title(title, fontsize=9)
            style_axis(ax)
            if invert:
                ax.invert_yaxis()  # 名次图：第 1 名在最上方

        # 图1 总分；图2 名次；图3 单科（随下拉框切换）
        draw_line(ax1, totals, "历次考试总分", "#534AB7")
        draw_line(ax2, ranks, "历次考试名次（第1名在最上）", "#B45309", invert=True)

        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self._chart_canvas = canvas  # 保留引用，便于切换单科时重绘

        def draw_subject(*_args):
            sub = subj_var.get()
            ax3.clear()
            series = subjects.get(sub, {})
            pts = [(x, series.get(b)) for x, b in zip(xs, batches)]
            pts = [(x, v) for x, v in pts if v is not None]
            if pts:
                xsv, ysv = [p[0] for p in pts], [p[1] for p in pts]
                ax3.plot(xsv, ysv, marker="s", color="#1A7F4B", linewidth=1.8)
                annotate(ax3, xsv, ysv, color="#1A7F4B")
            ax3.set_title("单科成绩趋势：%s" % (sub or "—"), fontsize=9)
            style_axis(ax3)
            canvas.draw()

        combo.bind("<<ComboboxSelected>>", draw_subject)
        if subj_names:
            draw_subject()
        else:
            ax3.set_title("暂无单科成绩", fontsize=9)
            style_axis(ax3)

    # ------------------------------------------------------------------
    # 其余分类页：可直接编辑（沿用 v6 逻辑）
    # ------------------------------------------------------------------
    def _create_scrollable_tab(self, notebook, title):
        frame = tk.Frame(notebook, bg="#FFFFFF")
        notebook.add(frame, text=title)
        canvas = tk.Canvas(frame, bg="#FFFFFF", highlightthickness=0)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg="#FFFFFF")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        inner.grid_columnconfigure(1, weight=1)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"),
                        add="+")
        return inner

    def _add_entry_widget(self, parent_frame, row_index, column_name, disabled=False):
        tk.Label(parent_frame, text=f"{column_name}：", anchor="w",
                 font=("Microsoft YaHei", 10), fg="#39406B").grid(
            row=row_index, column=0, sticky="w", padx=5, pady=4)
        value = '' if self.data.get(column_name) is None else str(self.data.get(column_name))
        var = tk.StringVar(value=value)
        entry = tk.Entry(parent_frame, textvariable=var, font=("Microsoft YaHei", 10),
                         relief="flat", highlightthickness=1,
                         highlightbackground="#D5D9E8", highlightcolor="#534AB7")
        if disabled:
            entry.config(state="disabled", disabledforeground="#9AA0B5")
        entry.grid(row=row_index, column=1, sticky="ew", padx=5, pady=4, ipady=3)
        self.entry_vars[column_name] = var

    def build_profile(self):
        container = tk.Frame(self.window, bg="#F4F5FA")
        container.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        self.build_profile_card(container)

        notebook = ttk.Notebook(container)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        self.build_academic_tab(notebook)

        if not self.column_groups:
            tab_inner = self._create_scrollable_tab(notebook, "详细信息")
            tk.Label(tab_inner, text="暂无其他字段", fg="#757575").grid(row=0, column=0, pady=10)
        else:
            for gi, (category, columns) in enumerate(self.column_groups):
                tab_inner = self._create_scrollable_tab(notebook, category)
                start = 0
                if gi == 0:
                    self._add_entry_widget(tab_inner, 0, "学号", disabled=True)
                    start = 1
                for idx, col in enumerate(columns, start=start):
                    self._add_entry_widget(tab_inner, idx, col)

        btn_frame = tk.Frame(self.window, bg="#F4F5FA", pady=10)
        btn_frame.pack(fill=tk.X)

        self.parent._flat_button(btn_frame, "保存修改", self.save_changes,
                                 "#534AB7", "#6A5FD0", padx=36, pady=6,
                                 font=("Microsoft YaHei", 11)).pack(side=tk.LEFT, expand=True, padx=10)
        self.parent._flat_button(btn_frame, "关闭", self.window.destroy,
                                 "#6B7280", "#7D8694", padx=36, pady=6,
                                 font=("Microsoft YaHei", 11)).pack(side=tk.LEFT, expand=True, padx=10)

    def save_changes(self):
        try:
            conn = sqlite3.connect(self.parent.db_path)
            cursor = conn.cursor()
            update_columns = [col for col in self.columns if col != "学号"]
            set_clause = ", ".join([f'"{col}" = ?' for col in update_columns])
            # 手工编辑同样按列类型规范化，避免详情页保存后类型又乱掉
            types = get_column_types(self.parent.db_path)
            values = []
            for col in update_columns:
                if col not in self.entry_vars:
                    values.append(None)
                    continue
                val = self.entry_vars[col].get().strip()
                values.append(normalize_cell_value(val, types.get(col, "text")) if val else None)
            values.append(self.student_id)
            cursor.execute(f'UPDATE students SET {set_clause} WHERE 学号 = ?', values)
            conn.commit()
            conn.close()
            messagebox.showinfo("成功", "学生信息已更新", parent=self.window)
            self.parent.refresh_display()
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}", parent=self.window)


def main():
    """主函数"""
    if tk is None:
        print("未检测到 tkinter，无法启动图形界面。")
        print("本文件的数据层（导入/导出/类型转换函数）仍可作为模块导入使用。")
        sys.exit(1)
    try:
        root = tk.Tk()
        app = StudentManagementSystem(root)
        root.mainloop()
    except Exception as e:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("错误", f"程序运行出错：\n{str(e)}\n\n{traceback.format_exc()}")
        except Exception:
            print(f"程序运行出错：{str(e)}")
            print(traceback.format_exc())
        sys.exit(1)

# 说明：原 v6 文件末尾带有 2026-10-10 到期的试用期检查，
# 到期后程序一启动就会弹窗退出，现已移除（备份文件里保留原样，可随时查对）。

if __name__ == "__main__":
    main()