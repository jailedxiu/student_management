#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
学生档案管理系统 v8（家校沟通滚动登记 + 学生情况智能生成）

在 v7 的基础上新增：
  1. 家校沟通改为独立台账（communication_log 表），按「日期 / 方式 / 对象 / 主题 /
     内容 / 家长反馈 / 后续跟进 / 记录人」滚动登记，和成绩一样只增不改历史记录；
     旧的「最近一次谈话时间 / 谈话内容」两列会在首次启动时自动迁入台账，且每次
     新增或删除记录后自动回写这两列，导出 Excel 时仍能看到最近一次谈话。
     · 学生画像窗口新增「家校沟通」页：看该生全部记录，逐条新增 / 修改 / 删除；
     · 首页新增「家校沟通」按钮：打开全班级台账，左侧选学生，右侧连续登记。
  2. 新增「学生情况报告」：把该生档案、历次成绩与沟通记录汇成上下文，一键生成
       一、学生信息概况  二、优势与短板  三、家校沟通注意点  四、管理注意点
     · 离线生成：内置规则引擎，不联网即可用，结果可解释；
     · 模型生成：接入任意 OpenAI 兼容接口（DeepSeek / 通义 / 智谱 / Kimi / OpenAI…），
       在「模型设置」里填地址、模型名和 Key 即可，Key 仅存本机 ai_config.json；
     · 「复制提示词」可把同一份上下文贴到 WorkBuddy、豆包等对话里直接用；
     · 生成结果自动存入 student_reports，可回看、导出 TXT；
     · 首页「批量报告」可按当前班级一次性生成全班情况报告并导出 Excel。

v7 保留的能力：
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
import json
import queue
import shutil
import threading
import urllib.request
import urllib.error
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


# ==========================================================================
# v8 数据层：家校沟通台账 / 学生情况报告
#
#   家校沟通做成独立明细表，和成绩一样「只追加、不改历史」：
#     communication_log(学号, 沟通日期, 沟通方式, 沟通对象, 沟通主题,
#                       沟通内容, 家长反馈, 后续跟进, 记录人, created_at)
#   学生情况报告存在 student_reports，由规则引擎或大模型生成。
#   两张表都按学号关联 students，不在学生表里新增列。
# ==========================================================================

COMM_METHODS = ("电话", "微信", "面谈", "家访", "家长会", "短信", "其他")
COMM_TOPICS = ("学业成绩", "课堂表现", "纪律考勤", "心理情绪", "生活起居",
               "安全提醒", "升学就业", "费用资助", "常规沟通", "其他")
COMM_OBJECTS = ("父亲", "母亲", "父母双方", "祖父母", "外祖父母", "其他监护人")

AI_CONFIG_NAME = "ai_config.json"
AI_PRESETS = [
    ("DeepSeek", "https://api.deepseek.com/v1", "deepseek-chat"),
    ("通义千问（阿里云）", "https://dashscope.aliyuncs.com/compatible-mode/v1", "qwen-plus"),
    ("智谱 GLM", "https://open.bigmodel.cn/api/paas/v4", "glm-4-flash"),
    ("月之暗面 Kimi", "https://api.moonshot.cn/v1", "moonshot-v1-8k"),
    ("OpenAI", "https://api.openai.com/v1", "gpt-4o-mini"),
    ("自定义", "", ""),
]

AI_SYSTEM_PROMPT = (
    "你是一名中职学校班主任工作助手，熟悉中职学生的学业管理、行为习惯养成与家校沟通。"
    "你只依据用户提供的学生档案数据作分析，缺失的数据就写「数据未录入」，绝不编造。"
    "你的输出要客观、具体、可执行，不对学生做定性评价，不使用「差生」「问题学生」等标签。"
)


def _now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def load_ai_config(base_path):
    """读取本机大模型接口配置（不存在时返回空配置）。"""
    cfg = {"preset": "", "base_url": "", "model": "", "api_key": ""}
    path = os.path.join(base_path, AI_CONFIG_NAME)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            cfg.update({k: (v or "") for k, v in data.items()})
    except Exception:
        pass
    return cfg


def save_ai_config(base_path, cfg):
    """写入本机配置；文件权限尽量收窄，避免误传到仓库。"""
    path = os.path.join(base_path, AI_CONFIG_NAME)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    try:
        os.chmod(path, 0o600)
    except Exception:
        pass
    return path


# ------------------------------ 建表与迁移 ------------------------------
def ensure_aux_tables(db_path):
    """建立 v8 新增的台账表（幂等，可反复调用）。"""
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute('''
            CREATE TABLE IF NOT EXISTS communication_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                学号 TEXT NOT NULL,
                沟通日期 TEXT,
                沟通方式 TEXT,
                沟通对象 TEXT,
                沟通主题 TEXT,
                沟通内容 TEXT,
                家长反馈 TEXT,
                后续跟进 TEXT,
                记录人 TEXT,
                created_at TEXT
            )
        ''')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_comm_student '
                    'ON communication_log(学号, 沟通日期)')
        cur.execute('''
            CREATE TABLE IF NOT EXISTS student_reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                学号 TEXT NOT NULL,
                生成时间 TEXT,
                来源 TEXT,
                内容 TEXT
            )
        ''')
        cur.execute('CREATE INDEX IF NOT EXISTS idx_report_student '
                    'ON student_reports(学号, 生成时间)')
        cur.execute('''
            CREATE TABLE IF NOT EXISTS app_meta (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        ''')
        conn.commit()
    finally:
        conn.close()


def migrate_legacy_talks(db_path):
    """把旧的「最近一次谈话时间 / 谈话内容」两列迁入台账，只做一次。"""
    conn = sqlite3.connect(db_path)
    migrated = 0
    try:
        cur = conn.cursor()
        done = cur.execute(
            "SELECT value FROM app_meta WHERE key = 'legacy_talk_migrated'").fetchone()
        if done:
            return 0
        cols = [r[1] for r in cur.execute("PRAGMA table_info(students)")]
        time_col = next((c for c in cols if "谈话" in c and ("时间" in c or "日期" in c)), None)
        text_col = next((c for c in cols if c != time_col and "谈话" in c), None)
        if text_col:
            tsel = f'"{time_col}"' if time_col else "''"
            rows = cur.execute(
                f'SELECT 学号, {tsel}, "{text_col}" FROM students').fetchall()
            for sid, when, content in rows:
                content = (content or "").strip()
                if not content:
                    continue
                dup = cur.execute(
                    "SELECT 1 FROM communication_log WHERE 学号 = ? AND 沟通内容 = ? LIMIT 1",
                    (sid, content)).fetchone()
                if dup:
                    continue
                cur.execute(
                    "INSERT INTO communication_log (学号, 沟通日期, 沟通方式, 沟通对象, "
                    "沟通主题, 沟通内容, 家长反馈, 后续跟进, 记录人, created_at) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?)",
                    (sid, _normalize_log_date(when), "面谈", "", "常规沟通",
                     content, "", "", "历史数据迁移", _now_str()))
                migrated += 1
        cur.execute("INSERT OR REPLACE INTO app_meta (key, value) VALUES (?, ?)",
                    ("legacy_talk_migrated", _now_str()))
        conn.commit()
    finally:
        conn.close()
    return migrated


def _normalize_log_date(value):
    """把各类日期写法统一成 YYYY-MM-DD；认不出就原样保留。"""
    if value in (None, ""):
        return ""
    s = str(value).strip()
    if _ISO_DATE_RE.match(s):
        return s
    n = _to_number(s)
    if n is not None and EXCEL_SERIAL_MIN <= n <= EXCEL_SERIAL_MAX:
        try:
            return (EXCEL_EPOCH + timedelta(days=int(n))).strftime("%Y-%m-%d")
        except Exception:
            return s
    m = _CN_DATE_RE.match(s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3))).strftime("%Y-%m-%d")
        except ValueError:
            return s
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s


# ------------------------------ 沟通记录增删改查 ------------------------------
def fetch_communication_logs(db_path, student_id=None, limit=None):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        sql = "SELECT * FROM communication_log"
        args = []
        if student_id:
            sql += " WHERE 学号 = ?"
            args.append(student_id)
        sql += " ORDER BY 沟通日期 DESC, id DESC"
        if limit:
            sql += f" LIMIT {int(limit)}"
        return [dict(r) for r in conn.execute(sql, args).fetchall()]
    finally:
        conn.close()


def add_communication_log(db_path, student_id, fields):
    """追加一条沟通记录；同时回写学生表的「最近一次谈话」两列，保持导出兼容。"""
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO communication_log (学号, 沟通日期, 沟通方式, 沟通对象, 沟通主题, "
            "沟通内容, 家长反馈, 后续跟进, 记录人, created_at) VALUES (?,?,?,?,?,?,?,?,?,?)",
            (student_id,
             _normalize_log_date(fields.get("date")),
             fields.get("method", ""),
             fields.get("target", ""),
             fields.get("topic", ""),
             fields.get("content", ""),
             fields.get("feedback", ""),
             fields.get("followup", ""),
             fields.get("recorder", ""),
             _now_str()))
        log_id = cur.lastrowid
        conn.commit()
    finally:
        conn.close()
    sync_latest_talk_to_student(db_path, student_id)
    return log_id


def update_communication_log(db_path, log_id, fields):
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            "UPDATE communication_log SET 沟通日期=?, 沟通方式=?, 沟通对象=?, 沟通主题=?, "
            "沟通内容=?, 家长反馈=?, 后续跟进=?, 记录人=? WHERE id=?",
            (_normalize_log_date(fields.get("date")),
             fields.get("method", ""), fields.get("target", ""),
             fields.get("topic", ""), fields.get("content", ""),
             fields.get("feedback", ""), fields.get("followup", ""),
             fields.get("recorder", ""), log_id))
        sid = cur.execute("SELECT 学号 FROM communication_log WHERE id=?",
                          (log_id,)).fetchone()
        conn.commit()
    finally:
        conn.close()
    if sid:
        sync_latest_talk_to_student(db_path, sid[0])


def delete_communication_log(db_path, log_id):
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        row = cur.execute("SELECT 学号 FROM communication_log WHERE id=?",
                          (log_id,)).fetchone()
        cur.execute("DELETE FROM communication_log WHERE id=?", (log_id,))
        conn.commit()
    finally:
        conn.close()
    if row:
        sync_latest_talk_to_student(db_path, row[0])


def sync_latest_talk_to_student(db_path, student_id):
    """把台账里最新一条记录回写到学生表的谈话两列（列不存在则跳过）。"""
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cols = [r[1] for r in cur.execute("PRAGMA table_info(students)")]
        time_col = next((c for c in cols if "谈话" in c and ("时间" in c or "日期" in c)), None)
        text_col = next((c for c in cols if c != time_col and "谈话" in c), None)
        if not time_col and not text_col:
            return
        row = cur.execute(
            "SELECT 沟通日期, 沟通内容 FROM communication_log WHERE 学号 = ? "
            "ORDER BY 沟通日期 DESC, id DESC LIMIT 1", (student_id,)).fetchone()
        sets, vals = [], []
        if time_col:
            sets.append(f'"{time_col}" = ?')
            vals.append((row[0] if row else None))
        if text_col:
            sets.append(f'"{text_col}" = ?')
            vals.append((row[1] if row else None))
        vals.append(student_id)
        cur.execute(f"UPDATE students SET {', '.join(sets)} WHERE 学号 = ?", vals)
        conn.commit()
    finally:
        conn.close()


# ------------------------------ 报告存取 ------------------------------
def save_report(db_path, student_id, content, source):
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO student_reports (学号, 生成时间, 来源, 内容) VALUES (?,?,?,?)",
            (student_id, _now_str(), source, content))
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def fetch_latest_report(db_path, student_id):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute(
            "SELECT * FROM student_reports WHERE 学号 = ? ORDER BY id DESC LIMIT 1",
            (student_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


# ------------------------------ 上下文汇总 ------------------------------
EXAM_COL_RE = re.compile(r"^(\d{4})\s*(月考|期中|期末)(.+)$")
EXAM_TOTAL_SUBS = ("总分", "成绩", "名次")


def _to_number(value):
    if value in (None, ""):
        return None
    try:
        f = float(str(value).strip())
        if math.isnan(f):
            return None
        return f
    except (TypeError, ValueError):
        return None


def _fmt_num(value):
    n = _to_number(value)
    if n is None:
        return None
    return int(n) if float(n).is_integer() else round(n, 1)


def _flag_true(value):
    """「是 / √ / 单亲 / 有」都算成立；空值和「否 / 无 / 0」算不成立。"""
    s = str(value).strip().lower()
    return s not in ("", "否", "无", "0", "false", "none", "nan", "—", "-")


def _txt(value):
    """把占位符（-、/、无）和 nan/None 统一当成空值。"""
    s = str(value if value is not None else "").strip()
    if s.lower() in ("-", "—", "/", "无", "空", "none", "nan", "null"):
        return ""
    return s


def _end_sentence(text):
    """去掉重复句号后再补一个句号，避免出现「。。」。"""
    s = (text or "").strip()
    s = s.rstrip("。；;，, ")
    return s + "。" if s else ""


def _age_from_birthday(birthday, id_card=""):
    text = str(birthday or "").strip()
    m = re.match(r"^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})", text)
    if not m and id_card:
        m2 = re.match(r"^\d{6}(\d{4})(\d{2})(\d{2})", str(id_card).strip())
        if m2:
            m = m2
    if not m:
        return None
    try:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        born = date(y, mo, d)
    except ValueError:
        return None
    today = date.today()
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))


def parse_exam_columns(columns):
    """把列名解析成「考试批次 → 科目」两级；与画像页同一套规则。"""
    exams, others = {}, []
    for col in columns:
        if col == "学号":
            continue
        m = EXAM_COL_RE.match(col)
        if m:
            batch = m.group(1) + m.group(2)
            sub = m.group(3).strip() or "总分"
            exams.setdefault(batch, []).append((sub, col))
        elif any(k in col for k in ("成绩", "名次", "得分", "总分")):
            others.append(col)

    def sort_key(item):
        sub = item[0]
        idx = EXAM_TOTAL_SUBS.index(sub) if sub in EXAM_TOTAL_SUBS else 99
        return (sub not in EXAM_TOTAL_SUBS, idx, sub)

    result = []
    for batch in sorted(exams.keys()):
        result.append((batch, sorted(exams[batch], key=sort_key)))
    return result, sorted(others)


def _rank_position(db_path, column, value):
    """同一批次里，本班有多少人的名次比该生靠前（名次数字越小越靠前）。"""
    n = _to_number(value)
    if n is None:
        return None
    conn = sqlite3.connect(db_path)
    try:
        vals = [_to_number(r[0]) for r in conn.execute(f'SELECT "{column}" FROM students')]
    finally:
        conn.close()
    vals = [v for v in vals if v is not None]
    if not vals:
        return None
    return sum(1 for v in vals if v < n) + 1, len(vals)


def collect_student_context(db_path, student_id):
    """汇总生成报告所需的全部数据（只读）。"""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        row = conn.execute("SELECT * FROM students WHERE 学号 = ?",
                           (student_id,)).fetchone()
        if row is None:
            return None
        data = {k: row[k] for k in row.keys()}
        class_size = conn.execute("SELECT COUNT(*) FROM students").fetchone()[0]
        raw_logs = conn.execute(
            "SELECT * FROM communication_log WHERE 学号 = ? "
            "ORDER BY 沟通日期 DESC, id DESC", (student_id,)).fetchall()
        logs = [dict(r) for r in raw_logs]
    finally:
        conn.close()

    columns = [c for c in data.keys() if c != "学号"]
    exam_groups, other_cols = parse_exam_columns(columns)

    exams = []
    for batch, subs in exam_groups:
        item = {"batch": batch, "subs": {}, "cols": {},
                "rank_pos": None, "rank_total": None}
        for sub, col in subs:
            item["subs"][sub] = data.get(col)
            item["cols"][sub] = col
        if "名次" in item["cols"]:
            pos = _rank_position(db_path, item["cols"]["名次"],
                                 data.get(item["cols"]["名次"]))
            if pos:
                item["rank_pos"], item["rank_total"] = pos
        exams.append(item)

    middle = {}
    for col in other_cols:
        if "中考" in col:
            key = "rank" if "名次" in col else "score"
            middle[key] = data.get(col)

    ctx = {
        "student_id": student_id,
        "name": _txt(data.get("姓名")),
        "gender": _txt(data.get("性别")),
        "birthday": _txt(data.get("生日")) or _txt(data.get("出生日期")),
        "age": _age_from_birthday(data.get("生日") or data.get("出生日期"),
                                  data.get("身份证号码")),
        "height": _txt(data.get("身高")),
        "dorm": _txt(data.get("宿舍号")),
        "day_student": data.get("是否走读"),
        "address": _txt(data.get("现居地")) or _txt(data.get("身份证地址")),
        "red_yellow_blue": _txt(data.get("红黄蓝情况简述")),
        "student_union": data.get("是否学生会"),
        "single_parent": data.get("单亲/孤儿"),
        "poverty": data.get("是否建档立卡"),
        "disabled": data.get("是否残疾"),
        "father_name": _txt(data.get("父亲姓名")),
        "father_phone": _txt(data.get("父亲电话号码")),
        "mother_name": _txt(data.get("母亲姓名")),
        "mother_phone": _txt(data.get("母亲电话号码")),
        "phone": _txt(data.get("电话号码")),
        "class_size": class_size,
        "exams": exams,
        "middle": middle,
        "logs": logs,
    }
    return ctx


def _subject_scores(exam):
    out = {}
    for sub, val in exam["subs"].items():
        if sub in EXAM_TOTAL_SUBS or "总分" in sub:
            continue
        n = _to_number(val)
        if n is not None:
            out[sub] = n
    return out


def _rank_series(ctx):
    series = []
    for e in ctx["exams"]:
        r = _to_number(e["subs"].get("名次"))
        if r is not None:
            series.append((e["batch"], r))
    return series


def _total_series(ctx):
    series = []
    for e in ctx["exams"]:
        t = _to_number(e["subs"].get("总分") or e["subs"].get("成绩"))
        if t is not None:
            series.append((e["batch"], t))
    return series


def _days_since(date_text):
    s = _normalize_log_date(date_text)
    try:
        d = datetime.strptime(s, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None
    return (date.today() - d).days


def _top_topic(logs):
    counter = {}
    for lg in logs:
        key = (lg.get("沟通主题") or "").strip()
        if key:
            counter[key] = counter.get(key, 0) + 1
    if not counter:
        return None
    return max(counter.items(), key=lambda kv: kv[1])


def _top_method(logs):
    counter = {}
    for lg in logs:
        key = (lg.get("沟通方式") or "").strip()
        if key:
            counter[key] = counter.get(key, 0) + 1
    if not counter:
        return None
    return max(counter.items(), key=lambda kv: kv[1])


# ------------------------------ 离线规则生成 ------------------------------
def generate_local_report(ctx):
    """不联网的规则引擎：按档案数据拼出四段式学生情况。"""
    name = ctx.get("name") or ctx["student_id"]
    logs = ctx.get("logs") or []
    ranks = _rank_series(ctx)
    totals = _total_series(ctx)
    last_exam = ctx["exams"][-1] if ctx["exams"] else None
    subjects = _subject_scores(last_exam) if last_exam else {}
    lines = []

    # ---------------- 一、信息概况 ----------------
    lines.append("一、学生信息概况")
    base = [f"{name}", ctx.get("gender") or "性别未录入"]
    if ctx.get("age"):
        base.append(f"{ctx['age']}岁")
    if ctx.get("dorm"):
        base.append(f"住{ctx['dorm']}宿舍")
    if _flag_true(ctx.get("day_student")):
        base.append("走读")
    lines.append("· " + "，".join(base) + "。")

    family = []
    if ctx.get("father_name") or ctx.get("father_phone"):
        family.append("父亲 " + (ctx.get("father_name") or "姓名未录")
                      + "（" + (ctx.get("father_phone") or "电话未录") + "）")
    if ctx.get("mother_name") or ctx.get("mother_phone"):
        family.append("母亲 " + (ctx.get("mother_name") or "姓名未录")
                      + "（" + (ctx.get("mother_phone") or "电话未录") + "）")
    tags = []
    if _flag_true(ctx.get("single_parent")):
        tags.append("单亲/孤儿家庭")
    if _flag_true(ctx.get("poverty")):
        tags.append("建档立卡")
    if _flag_true(ctx.get("disabled")):
        tags.append("残疾")
    fam_text = "· 家庭情况：" + ("；".join(family) if family else "父母信息未录入")
    if tags:
        fam_text += "；" + "、".join(tags)
    lines.append(fam_text + "。")

    exam_by_batch = {e["batch"]: e for e in ctx["exams"]}
    rank_map = dict(ranks)
    if totals:
        batch, total = totals[-1]
        seg = f"· 学业：最近一次考试（{batch}）总分 {_fmt_num(total)}"
        if batch in rank_map:
            seg += f"，名次 {_fmt_num(rank_map[batch])}"
        lines.append(seg + "。")
    elif last_exam:
        lines.append(f"· 学业：已录入 {last_exam['batch']} 成绩，但总分/名次为空，需核实是否缺考。")
    else:
        lines.append("· 学业：档案中暂无考试科目成绩，无法判断学业水平。")

    if ranks:
        lb, lr = ranks[-1]
        ex = exam_by_batch.get(lb, {})
        seg = f"· 排名：最近一次有名次的考试为 {lb}，名次 {_fmt_num(lr)}"
        if ex.get("rank_pos") and ex.get("rank_total"):
            seg += (f"，按名次在本班排第 {ex['rank_pos']} 位"
                    f"（{ex['rank_total']} 人有成绩）")
        lines.append(seg + "。")
    if last_exam and totals and last_exam["batch"] != totals[-1][0]:
        lines.append(f"· 说明：{last_exam['batch']} 已建好成绩列但为空，"
                     f"本次分析以 {totals[-1][0]} 为最近一次有效成绩。")

    if ctx.get("middle"):
        mid = ctx["middle"]
        seg = "· 入校基础："
        parts = []
        if mid.get("score") not in (None, ""):
            parts.append(f"中考成绩 {_fmt_num(mid['score'])}")
        if mid.get("rank") not in (None, ""):
            parts.append(f"中考名次 {_fmt_num(mid['rank'])}")
        lines.append(seg + ("，".join(parts) if parts else "未录入") + "。")

    identity = []
    if _flag_true(ctx.get("student_union")):
        identity.append("学生会/学生干部")
    ryb_clean = str(ctx.get("red_yellow_blue") or "").strip().rstrip("。；;，, ")
    if ryb_clean:
        identity.append(f"红黄蓝记录：{ryb_clean}")
    if identity:
        lines.append(_end_sentence("· 在校表现：" + "；".join(identity)))

    if logs:
        last = logs[0]
        topic = _top_topic(logs)
        seg = (f"· 家校沟通：已有 {len(logs)} 条记录，最近一次 "
               f"{last.get('沟通日期') or '日期未录'}"
               f"（{last.get('沟通方式') or '方式未录'}"
               f"·{last.get('沟通对象') or '对象未录'}）")
        if topic:
            seg += f"；沟通最多的是「{topic[0]}」"
        lines.append(seg + "。")
    else:
        lines.append("· 家校沟通：档案中暂无沟通记录，家庭情况与家长诉求均缺少一手信息。")

    # ---------------- 二、优势与短板 ----------------
    lines.append("")
    lines.append("二、优势与短板")
    strengths, gaps = [], []

    if len(ranks) >= 2:
        (pb, pr), (lb, lr) = ranks[-2], ranks[-1]
        delta = pr - lr
        if delta >= 3:
            strengths.append(f"名次从 {pb} 的第 {_fmt_num(pr)} 名提升到 {lb} 的第 "
                             f"{_fmt_num(lr)} 名，进步 {int(delta)} 位，上升趋势明确")
        elif delta <= -3:
            gaps.append(f"名次从 {pb} 的第 {_fmt_num(pr)} 名降到 {lb} 的第 "
                        f"{_fmt_num(lr)} 名，下滑 {int(-delta)} 位，需要尽快查明原因")
    if ranks:
        best_batch, best_rank = min(ranks, key=lambda x: x[1])
        strengths.append(f"历史最好名次为 {best_batch} 第 {_fmt_num(best_rank)} 名")
        last_ex = exam_by_batch.get(ranks[-1][0], {})
        pos, pos_total = last_ex.get("rank_pos"), last_ex.get("rank_total")
        if pos and pos_total and pos <= pos_total * 0.3:
            strengths.append(f"按名次在本班排第 {pos} 位（{pos_total} 人有成绩），"
                             f"属班内前 30%")
        if pos and pos_total and pos > pos_total * 0.7:
            gaps.append(f"按名次在本班排第 {pos} 位（{pos_total} 人有成绩），"
                        f"处于后 30%，基础薄弱面较大")
        if best_batch != ranks[-1][0] and ranks[-1][1] - best_rank >= 10:
            gaps.append(f"当前名次比历史最好的第 {_fmt_num(best_rank)} 名"
                        f"（{best_batch}）回落了 {int(ranks[-1][1] - best_rank)} 位")
    if subjects:
        top_sub, top_val = max(subjects.items(), key=lambda kv: kv[1])
        low_sub, low_val = min(subjects.items(), key=lambda kv: kv[1])
        strengths.append(f"{last_exam['batch']} 的 {top_sub} 得 {_fmt_num(top_val)} 分，"
                         f"是当前相对优势科目")
        if low_val <= 60 or (top_val and low_val <= top_val * 0.5):
            gaps.append(f"{low_sub} 仅 {_fmt_num(low_val)} 分，与 {top_sub} 差距明显，"
                        f"属当前最薄弱科目")
    missing_subs = []
    if last_exam:
        for sub, val in last_exam["subs"].items():
            if sub in EXAM_TOTAL_SUBS or "总分" in sub:
                continue
            if val in (None, ""):
                missing_subs.append(sub)
        if missing_subs:
            gaps.append(f"{last_exam['batch']} 的 {'、'.join(missing_subs)} 无成绩，"
                        f"需核实是缺考还是数据未录")
    if _flag_true(ctx.get("student_union")):
        strengths.append("担任学生会/学生干部，组织能力与自我管理有一定基础")
    if len(logs) >= 3:
        strengths.append(f"家长已参与 {len(logs)} 次沟通，家校配合有基础")
    elif not logs:
        gaps.append("家校沟通记录为 0，缺少对家庭结构与家长诉求的掌握")
    if not (ctx.get("father_phone") or ctx.get("mother_phone")) and not ctx.get("phone"):
        gaps.append("监护人联系电话缺失，突发情况难以第一时间联系")

    if not strengths:
        strengths.append("目前可量化的优势数据较少，建议补充课堂表现、技能实操等观察记录后再评估")
    if not gaps:
        gaps.append("暂无明显短板数据，建议关注单科均衡与学习习惯的持续跟踪")
    for s in strengths:
        lines.append("优势｜" + s + "。")
    for g in gaps:
        lines.append("短板｜" + g + "。")

    # ---------------- 三、家校沟通注意点 ----------------
    lines.append("")
    lines.append("三、家校沟通注意点")
    notes = []
    has_phone = bool(ctx.get("father_phone") or ctx.get("mother_phone") or ctx.get("phone"))
    if not has_phone:
        notes.append("先补齐监护人电话与称呼，再约定固定联系频率，避免每次沟通都从头确认基本信息")
    if len(logs) >= 1:
        days = _days_since(logs[0].get("沟通日期"))
        if days is not None and days > 30:
            notes.append(f"距上次沟通已 {days} 天，建议本月内安排一次电话或面谈，"
                         f"先问在家表现再通报在校情况")
        if len(logs) == 1:
            notes.append("目前只有 1 条记录，建议按「学业—表现—家庭」三段补记一次，"
                         "把家长诉求与约定事项写清楚，形成可对照的基线")
        method = _top_method(logs)
        if method and method[1] >= 2:
            notes.append(f"历史上用「{method[0]}」沟通最多、响应较顺，可优先沿用")
        empty_feedback = sum(1 for lg in logs if not str(lg.get("家长反馈") or "").strip())
        if len(logs) >= 2 and empty_feedback / len(logs) > 0.5:
            notes.append("多条记录没写家长反馈，登记时补记家长诉求与承诺事项，下次沟通可直接接着谈")
    else:
        notes.append("建议两周内完成首次家访或电话沟通，先了解家庭结构、在家作息与家长最关心的问题")
    if _flag_true(ctx.get("single_parent")):
        notes.append("涉及单亲/孤儿情况，沟通以稳定情绪和鼓励为主，避免标签化表述，注意保护隐私")
    if _flag_true(ctx.get("poverty")):
        notes.append("建档立卡家庭可同步介绍资助、免学费与勤工助学政策，减少家庭顾虑")
    if _flag_true(ctx.get("day_student")):
        notes.append("走读生需与家长明确上下学交通方式、到家时间与晚间手机管理，"
                     "把安全责任讲清楚")
    ryb_text = str(ctx.get("red_yellow_blue") or "")
    if any(k in ryb_text for k in ("转专业", "流生", "退学", "积极性", "厌学", "不想读")):
        notes.append("档案显示存在专业认同或学习动力的隐患，沟通时先了解学生对专业和出路的"
                     "真实想法，再谈成绩要求")
    if any(k in ryb_text for k in ("心理", "情绪", "抑郁", "焦虑", "自残")):
        notes.append("涉及心理与情绪情况，措辞以关心为主，必要时告知学校心理老师并同步家长")
    if any(k in ryb_text for k in ("在外地", "国外", "不在身边", "留守", "祖辈")):
        notes.append("父母不常在身边，日常联系优先找实际监护人，沟通时先说明孩子的具体进步，"
                     "避免让家长产生被指责的感觉")
    if len(ranks) >= 2 and ranks[-1][1] - ranks[-2][1] >= 3:
        notes.append("成绩下滑沟通时先听家长描述在家变化（作息、手机、交友、家庭事务），"
                     "再一起定一两条可执行的家庭配合动作")
    last_ex2 = exam_by_batch.get(ranks[-1][0], {}) if ranks else {}
    pos2, pos2_total = last_ex2.get("rank_pos"), last_ex2.get("rank_total")
    if pos2 and pos2_total and pos2 > pos2_total * 0.7:
        notes.append("与家长谈成绩先呈现具体进步点，再谈改进措施，避免单向通报分数引发对抗")
    if not notes:
        notes.append("保持每月一次常规沟通，学业与行为各谈一点，形成连续的成长记录")

    for n in notes:
        lines.append("· " + n + "。")

    # ---------------- 四、管理注意点 ----------------
    lines.append("")
    lines.append("四、管理注意点")
    mg = []
    ryb = str(ctx.get("red_yellow_blue") or "")
    if "红" in ryb:
        mg.append("红黄蓝记录含「红」色关注项，建议列为重点跟踪对象，明确跟踪责任人与每周反馈节点")
    elif "黄" in ryb:
        mg.append("红黄蓝记录含「黄」色提示，建议两周一次谈心，记录情绪与行为变化")
    elif "蓝" in ryb:
        mg.append("红黄蓝记录含「蓝」色提示，可结合日常表现做正向引导，避免过度关注造成压力")
    if any(k in ryb for k in ("转专业", "流生", "退学")):
        mg.append("把「留得住」放在成绩之前：安排技能展示、企业参观或师傅带徒等体验，"
                  "先建立专业认同")
    if "心脏" in ryb or "剧烈运动" in ryb:
        mg.append("档案提到身体限制（心脏/不宜剧烈运动），实训与体育活动前先与家长、"
                  "体育老师确认强度")
    if any(k in ryb for k in ("心理", "情绪", "抑郁", "焦虑")):
        mg.append("情绪风险列入日常观察，安排一名同学结对陪伴，出现持续低落及时转介心理老师")
    if _flag_true(ctx.get("student_union")):
        mg.append("学生干部事务较多，注意不占用晚自习与实训时间，成绩波动时及时减负")
    if ctx.get("dorm"):
        mg.append(f"住宿在{ctx['dorm']}，宿舍作息与室友关系可通过寝室长定期了解，"
                  f"发现熬夜、串寝及时介入")
    if _flag_true(ctx.get("day_student")):
        mg.append("走读生重点登记到校时间与迟到情况，与家长保持通学安全的信息互通")
    if subjects:
        low_sub, low_val = min(subjects.items(), key=lambda kv: kv[1])
        if low_val <= 60:
            mg.append(f"针对{low_sub}安排结对帮扶或每周固定辅导时段，连续跟踪 4 周看变化")
    if pos2 and pos2_total and pos2 <= pos2_total * 0.3:
        mg.append("学业居班内前列，管理重点可从督促转为给任务（技能竞赛、考证、带同伴），"
                  "避免只用分数评价")
    if missing_subs:
        mg.append(f"核实 {'、'.join(missing_subs)} 缺考原因，必要时安排补测，避免成绩档案断档")
    if len(ranks) >= 2 and ranks[-1][1] - ranks[-2][1] >= 3:
        mg.append("近期成绩下滑，先排查手机、作息、同伴关系三类常见诱因，再定干预措施")
    elif len(ranks) >= 2 and ranks[-2][1] - ranks[-1][1] >= 3:
        mg.append("名次有进步，及时用具体事实公开肯定，把进步归因到可控的学习行为上")
    if not mg:
        mg.append("保持常规观察：每周记录一次课堂参与和作业完成情况，出现连续两周异常再升级干预")
    for m in mg:
        lines.append("· " + m + "。")

    lines.append("")
    lines.append("（本报告由系统按已录入数据自动生成，事实性内容请核对后使用；"
                 "未录入的字段不会推测，建议补齐后再生成一次。）")
    return "\n".join(lines)


# ------------------------------ 提示词与大模型调用 ------------------------------
def _context_to_text(ctx):
    name = ctx.get("name") or ctx["student_id"]
    L = []
    L.append(f"姓名：{name}　学号：{ctx['student_id']}　性别：{ctx.get('gender') or '未录'}"
             f"　年龄：{ctx.get('age') or '未录'}　身高：{ctx.get('height') or '未录'}")
    L.append(f"住宿：{'走读' if _flag_true(ctx.get('day_student')) else ('住 ' + str(ctx['dorm']) + ' 宿舍' if ctx.get('dorm') else '未录')}"
             f"　现居地：{ctx.get('address') or '未录'}")
    L.append(f"学生会/学生干部：{ctx.get('student_union') or '未录'}"
             f"　红黄蓝情况：{ctx.get('red_yellow_blue') or '未录'}")
    fam = (f"父亲：{ctx.get('father_name') or '未录'}（{ctx.get('father_phone') or '电话未录'}）；"
           f"母亲：{ctx.get('mother_name') or '未录'}（{ctx.get('mother_phone') or '电话未录'}）；"
           f"单亲/孤儿：{ctx.get('single_parent') or '否'}；建档立卡：{ctx.get('poverty') or '否'}；"
           f"残疾：{ctx.get('disabled') or '否'}")
    L.append(f"家庭情况：{fam}")
    L.append(f"班级人数：{ctx.get('class_size') or '未录'}")
    if ctx.get("middle"):
        L.append(f"入校成绩：中考成绩 {ctx['middle'].get('score') or '未录'}，"
                 f"中考名次 {ctx['middle'].get('rank') or '未录'}")
    L.append("历次考试成绩（按时间先后）：")
    if not ctx["exams"]:
        L.append("　（无）")
    for e in ctx["exams"]:
        subs = []
        for sub, val in e["subs"].items():
            if val in (None, ""):
                continue
            subs.append(f"{sub} {_fmt_num(val) if _fmt_num(val) is not None else val}")
        L.append(f"　{e['batch']}：" + ("，".join(subs) if subs else "无成绩"))
    L.append("家校沟通记录（最近在前）：")
    if not ctx["logs"]:
        L.append("　（无记录）")
    for lg in ctx["logs"][:8]:
        content = str(lg.get("沟通内容") or "").strip().replace("\n", " ")
        if len(content) > 160:
            content = content[:160] + "…"
        feedback = str(lg.get("家长反馈") or "").strip().replace("\n", " ")
        if len(feedback) > 80:
            feedback = feedback[:80] + "…"
        L.append(f"　{lg.get('沟通日期') or '日期未录'}　{lg.get('沟通方式') or ''}"
                 f"　{lg.get('沟通对象') or ''}　{lg.get('沟通主题') or ''}：{content}"
                 + (f"（家长反馈：{feedback}）" if feedback else ""))
    return "\n".join(L)


def build_ai_prompt(ctx):
    """把上下文拼成可直接贴给任意大模型的提示词。"""
    return (
        "请依据下面这名中职学生的档案数据，写一份供班主任使用的情况分析。\n\n"
        "输出格式要求：\n"
        "1. 只输出四个部分，标题固定为「一、学生信息概况」「二、优势与短板」"
        "「三、家校沟通注意点」「四、管理注意点」，不要增加其他章节。\n"
        "2. 每部分 3–5 条，每条以「· 」开头，先讲清数据事实，再给一句可执行建议；"
        "「二、优势与短板」的每条用「优势｜」「短板｜」标明。\n"
        "3. 语言客观简洁，不评价学生品行，不使用「差生」「问题学生」等标签。\n"
        "4. 数据缺失就写「数据未录入」，并说明需要补充哪类信息，不要编造。\n"
        "5. 全文 600–900 字，中文。\n\n"
        "【学生档案数据】\n" + _context_to_text(ctx)
    )


def call_llm(cfg, prompt, timeout=120):
    """调用任意 OpenAI 兼容的 chat/completions 接口。"""
    base = (cfg.get("base_url") or "").strip().rstrip("/")
    if not base:
        raise ValueError("未填写接口地址")
    if base.endswith("/chat/completions"):
        url = base
    else:
        url = base + "/chat/completions"
    payload = {
        "model": (cfg.get("model") or "").strip(),
        "messages": [
            {"role": "system", "content": AI_SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.4,
        "stream": False,
    }
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {"Content-Type": "application/json",
               "Authorization": "Bearer " + (cfg.get("api_key") or "").strip()}
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        detail = ""
        try:
            detail = e.read().decode("utf-8", errors="replace")[:300]
        except Exception:
            pass
        hint = {401: "API Key 无效或已过期", 403: "没有该模型的访问权限",
                404: "接口地址或模型名不对", 429: "调用频率/额度受限",
                400: "请求被服务端拒绝（常见原因是模型名不对）"}.get(e.code, "")
        raise RuntimeError(f"HTTP {e.code} {hint}\n{detail}".strip())
    except urllib.error.URLError as e:
        raise RuntimeError(
            f"无法连接接口：{e.reason}\n若本机需要通过代理上网，请先在系统代理里放行，"
            f"或在「模型设置」里换一个可直连的地址。")
    except Exception as e:
        raise RuntimeError(f"请求失败：{e}")
    try:
        obj = json.loads(raw)
    except json.JSONDecodeError:
        raise RuntimeError("返回内容不是 JSON，检查接口地址是否为 OpenAI 兼容格式。\n"
                           + raw[:300])
    choices = obj.get("choices") or []
    if not choices:
        err = obj.get("error") or obj
        raise RuntimeError("接口未返回结果：" + json.dumps(err, ensure_ascii=False)[:300])
    msg = choices[0].get("message") or {}
    text = (msg.get("content") or "").strip()
    if not text:
        raise RuntimeError("接口返回了空内容，请检查模型名是否可用。")
    return text


# ------------------------------ 报告导出 ------------------------------
REPORT_SECTIONS = ("一、学生信息概况", "二、优势与短板", "三、家校沟通注意点", "四、管理注意点")


def split_report_sections(text):
    """把报告拆成四段，供 Excel 分列导出（缺段则留空）。"""
    result = {k: "" for k in REPORT_SECTIONS}
    current = None
    for line in (text or "").splitlines():
        stripped = line.strip()
        matched = None
        for key in REPORT_SECTIONS:
            if stripped.startswith(key):
                matched = key
                break
        if matched:
            current = matched
            rest = stripped[len(matched):].lstrip("：: 　")
            if rest:
                result[current] += rest + "\n"
            continue
        if current:
            result[current] += line + "\n"
    return {k: v.strip() for k, v in result.items()}


def export_communication_xlsx(path, logs, name_map=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "家校沟通记录"
    headers = ["学号", "姓名", "沟通日期", "沟通方式", "沟通对象", "沟通主题",
               "沟通内容", "家长反馈", "后续跟进", "记录人"]
    head_fill = PatternFill("solid", fgColor="EEF0F7")
    thin = Side(style="thin", color="D5D9E8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = head_fill
        cell.font = Font(name="微软雅黑", bold=True, size=10, color="39406B")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    name_map = name_map or {}
    for r, lg in enumerate(logs, 2):
        vals = [lg.get("学号", ""), name_map.get(lg.get("学号"), ""),
                lg.get("沟通日期", ""), lg.get("沟通方式", ""), lg.get("沟通对象", ""),
                lg.get("沟通主题", ""), lg.get("沟通内容", ""),
                lg.get("家长反馈", ""), lg.get("后续跟进", ""), lg.get("记录人", "")]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=c in (7, 8, 9))
            cell.border = border
    widths = [12, 10, 12, 10, 12, 12, 50, 40, 30, 12]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)
    return path


def export_reports_xlsx(path, items):
    """把多名学生的报告导出成一个工作簿：总览 + 全文。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "情况总览"
    headers = ["学号", "姓名", "性别", "最新考试", "总分", "名次",
               "沟通次数", "学生信息概况", "优势与短板", "家校沟通注意点", "管理注意点"]
    head_fill = PatternFill("solid", fgColor="EEF0F7")
    thin = Side(style="thin", color="D5D9E8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = head_fill
        cell.font = Font(name="微软雅黑", bold=True, size=10, color="39406B")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for r, item in enumerate(items, 2):
        ctx = item["ctx"]
        totals = _total_series(ctx)
        ranks = _rank_series(ctx)
        latest_batch = totals[-1][0] if totals else (ranks[-1][0] if ranks else "")
        latest_total = _fmt_num(totals[-1][1]) if totals else ""
        latest_rank = _fmt_num(ranks[-1][1]) if ranks else ""
        sec = split_report_sections(item["report"])
        vals = [ctx["student_id"], ctx.get("name", ""), ctx.get("gender", ""),
                latest_batch, latest_total, latest_rank, len(ctx.get("logs") or []),
                sec.get("一、学生信息概况", ""), sec.get("二、优势与短板", ""),
                sec.get("三、家校沟通注意点", ""), sec.get("四、管理注意点", "")]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=c >= 8,
                                       horizontal="center" if c in (1, 3, 4, 5, 6, 7) else "left")
            cell.border = border
        ws.row_dimensions[r].height = 96
    for idx, w in enumerate([12, 10, 8, 12, 8, 8, 9, 40, 40, 40, 36], 1):
        ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = w
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("报告全文")
    ws2.cell(row=1, column=1, value="学号").font = Font(name="微软雅黑", bold=True, size=10)
    ws2.cell(row=1, column=2, value="姓名").font = Font(name="微软雅黑", bold=True, size=10)
    ws2.cell(row=1, column=3, value="学生情况报告").font = Font(name="微软雅黑", bold=True, size=10)
    for r, item in enumerate(items, 2):
        ws2.cell(row=r, column=1, value=item["ctx"]["student_id"])
        ws2.cell(row=r, column=2, value=item["ctx"].get("name", ""))
        cell = ws2.cell(row=r, column=3, value=item["report"])
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws2.row_dimensions[r].height = 260
    ws2.column_dimensions["A"].width = 12
    ws2.column_dimensions["B"].width = 10
    ws2.column_dimensions["C"].width = 110
    wb.save(path)
    return path


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
        # 大模型接口配置（仅存本机 ai_config.json，不入库、不上传）
        self.ai_config = load_ai_config(self.base_path)
        
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

            # v8：家校沟通台账 / 学生情况报告两张表，并把旧的谈话两列迁入台账
            ensure_aux_tables(self.db_path)
            migrate_legacy_talks(self.db_path)
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
        self._flat_button(btn_box, "家校沟通", self.open_communication_window,
                          "#0F8B8D", "#14A0A2").pack(side=tk.LEFT, padx=3)
        self._flat_button(btn_box, "批量报告", self.open_batch_report,
                          "#B4433F", "#C95450").pack(side=tk.LEFT, padx=3)

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

    def open_communication_window(self):
        """家校沟通台账：全班学生，逐个滚动登记。"""
        CommunicationWindow(self)

    def open_batch_report(self):
        """批量生成全班学生情况报告。"""
        BatchReportDialog(self)


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


class CommunicationPanel(tk.Frame):
    """家校沟通登记面板。

    两种用法：
      · 嵌在学生画像里 —— 传入 student_id，只看这一个学生；
      · 放进家校沟通台账窗口 —— selectable=True，顶部先选学生再登记。
    """

    def __init__(self, master, app, student_id=None, selectable=False, show_export=True):
        super().__init__(master, bg="#FFFFFF")
        self.app = app
        self.fixed_id = student_id
        self.selectable = selectable
        self.show_export = show_export
        self.current_log_id = None
        self.student_map = {}
        self.date_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))
        self.method_var = tk.StringVar(value=COMM_METHODS[0])
        self.object_var = tk.StringVar(value="")
        self.topic_var = tk.StringVar(value=COMM_TOPICS[0])
        self.recorder_var = tk.StringVar(value="")
        self._build_ui()
        if selectable:
            self.load_students()
        else:
            self.reload()

    # ------------------------------ 界面 ------------------------------
    def _build_ui(self):
        top = tk.Frame(self, bg="#FFFFFF")
        top.pack(fill=tk.X, padx=10, pady=(8, 4))
        if self.selectable:
            tk.Label(top, text="学生：", bg="#FFFFFF", fg="#39406B",
                     font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
            self.student_combo = ttk.Combobox(top, width=26, state="readonly")
            self.student_combo.pack(side=tk.LEFT)
            self.student_combo.bind("<<ComboboxSelected>>", self.on_student_change)
        self.summary_label = tk.Label(top, text="", bg="#FFFFFF", fg="#8A90A6",
                                      font=("Microsoft YaHei", 9))
        self.summary_label.pack(side=tk.LEFT, padx=10)
        if self.show_export:
            self.app._flat_button(top, "导出记录", self.export_logs,
                                  "#2F80ED", "#4A97F2", padx=12, pady=3).pack(side=tk.RIGHT)

        mid = tk.Frame(self, bg="#FFFFFF")
        mid.pack(fill=tk.BOTH, expand=True, padx=10)
        cols = ("日期", "方式", "对象", "主题", "内容摘要")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=6)
        widths = (92, 62, 78, 84, 420)
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor=tk.W if col == "内容摘要" else tk.CENTER)
        sb = tk.Scrollbar(mid, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.tag_configure("odd", background="#F8F9FC")
        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        form = tk.Frame(self, bg="#FFFFFF")
        form.pack(fill=tk.X, padx=10, pady=(8, 0))
        form.grid_columnconfigure(9, weight=1)

        def lab(text, row, col, anchor="e"):
            tk.Label(form, text=text, bg="#FFFFFF", fg="#39406B",
                     font=("Microsoft YaHei", 9)).grid(
                row=row, column=col, sticky=anchor, padx=(0, 4), pady=3)

        def entry(var, row, col, width=12):
            e = tk.Entry(form, textvariable=var, width=width, font=("Microsoft YaHei", 10),
                         relief="flat", highlightthickness=1,
                         highlightbackground="#D5D9E8", highlightcolor="#534AB7")
            e.grid(row=row, column=col, sticky="w", pady=3, ipady=2)
            return e

        lab("日期", 0, 0)
        self.date_entry = entry(self.date_var, 0, 1, 12)
        lab("方式", 0, 2)
        self.method_combo = ttk.Combobox(form, textvariable=self.method_var,
                                         values=COMM_METHODS, width=8, state="readonly")
        self.method_combo.grid(row=0, column=3, sticky="w", pady=3)
        lab("对象", 0, 4)
        self.object_combo = ttk.Combobox(form, textvariable=self.object_var,
                                         values=COMM_OBJECTS, width=10)
        self.object_combo.grid(row=0, column=5, sticky="w", pady=3)
        lab("主题", 0, 6)
        self.topic_combo = ttk.Combobox(form, textvariable=self.topic_var,
                                        values=COMM_TOPICS, width=10, state="readonly")
        self.topic_combo.grid(row=0, column=7, sticky="w", pady=3)
        lab("记录人", 0, 8)
        entry(self.recorder_var, 0, 9, 10)

        lab("沟通内容", 1, 0, anchor="nw")
        self.content_text = tk.Text(form, height=3, wrap="word",
                                    font=("Microsoft YaHei", 10), relief="flat",
                                    highlightthickness=1, highlightbackground="#D5D9E8",
                                    highlightcolor="#534AB7")
        self.content_text.grid(row=1, column=1, columnspan=9, sticky="ew", pady=3, padx=(0, 4))

        lab("家长反馈", 2, 0, anchor="nw")
        self.feedback_text = tk.Text(form, height=2, wrap="word",
                                     font=("Microsoft YaHei", 10), relief="flat",
                                     highlightthickness=1, highlightbackground="#D5D9E8",
                                     highlightcolor="#534AB7")
        self.feedback_text.grid(row=2, column=1, columnspan=4, sticky="ew", pady=3, padx=(0, 6))
        lab("后续跟进", 2, 5, anchor="nw")
        self.followup_text = tk.Text(form, height=2, wrap="word",
                                     font=("Microsoft YaHei", 10), relief="flat",
                                     highlightthickness=1, highlightbackground="#D5D9E8",
                                     highlightcolor="#534AB7")
        self.followup_text.grid(row=2, column=6, columnspan=4, sticky="ew", pady=3)

        btns = tk.Frame(self, bg="#FFFFFF")
        btns.pack(fill=tk.X, padx=10, pady=(6, 8))
        self.app._flat_button(btns, "新增登记", self.add_log,
                              "#1A7F4B", "#219663", padx=18, pady=4).pack(side=tk.LEFT)
        self.app._flat_button(btns, "保存修改", self.update_log,
                              "#534AB7", "#6A5FD0", padx=18, pady=4).pack(side=tk.LEFT, padx=6)
        self.app._flat_button(btns, "删除选中", self.delete_log,
                              "#E05252", "#E96A6A", padx=18, pady=4).pack(side=tk.LEFT)
        self.app._flat_button(btns, "清空表单", self.clear_form,
                              "#6B7280", "#7D8694", padx=18, pady=4).pack(side=tk.LEFT, padx=6)
        self.hint = tk.Label(btns, text="选一条历史记录可载入修改", bg="#FFFFFF",
                             fg="#8A90A6", font=("Microsoft YaHei", 9))
        self.hint.pack(side=tk.LEFT, padx=10)

    # ------------------------------ 数据 ------------------------------
    def current_student_id(self):
        if self.fixed_id:
            return self.fixed_id
        return self.student_map.get(self.student_combo.get(), "")

    def load_students(self):
        conn = sqlite3.connect(self.app.db_path)
        try:
            rows = conn.execute(
                "SELECT 学号, 姓名 FROM students ORDER BY 学号").fetchall()
        finally:
            conn.close()
        self.student_map = {f"{name or '未命名'}（{sid}）": sid for sid, name in rows}
        values = list(self.student_map.keys())
        self.student_combo["values"] = values
        if values and not self.student_combo.get():
            self.student_combo.current(0)
            self.reload()

    def on_student_change(self, event=None):
        self.clear_form()
        self.reload()

    def reload(self):
        sid = self.current_student_id()
        for item in self.tree.get_children():
            self.tree.delete(item)
        if not sid:
            self.summary_label.config(text="请先选择学生")
            return
        logs = fetch_communication_logs(self.app.db_path, sid)
        for i, lg in enumerate(logs):
            content = str(lg.get("沟通内容") or "").replace("\n", " ")
            if len(content) > 40:
                content = content[:40] + "…"
            self.tree.insert("", tk.END, iid=str(lg["id"]),
                             values=(lg.get("沟通日期") or "", lg.get("沟通方式") or "",
                                     lg.get("沟通对象") or "", lg.get("沟通主题") or "",
                                     content),
                             tags=("odd" if i % 2 else "even",))
        if logs:
            self.summary_label.config(
                text=f"共 {len(logs)} 条记录 · 最近一次 {logs[0].get('沟通日期') or '日期未录'}")
        else:
            self.summary_label.config(text="暂无沟通记录，右侧表单填好后点「新增登记」")

    def get_form(self):
        return {
            "date": self.date_var.get().strip(),
            "method": self.method_var.get().strip(),
            "target": self.object_var.get().strip(),
            "topic": self.topic_var.get().strip(),
            "content": self.content_text.get("1.0", tk.END).strip(),
            "feedback": self.feedback_text.get("1.0", tk.END).strip(),
            "followup": self.followup_text.get("1.0", tk.END).strip(),
            "recorder": self.recorder_var.get().strip(),
        }

    def clear_form(self):
        self.current_log_id = None
        self.date_var.set(date.today().strftime("%Y-%m-%d"))
        self.method_var.set(COMM_METHODS[0])
        self.object_var.set("")
        self.topic_var.set(COMM_TOPICS[0])
        self.recorder_var.set(self._default_recorder())
        self.content_text.delete("1.0", tk.END)
        self.feedback_text.delete("1.0", tk.END)
        self.followup_text.delete("1.0", tk.END)
        self.hint.config(text="选一条历史记录可载入修改")
        for item in self.tree.selection():
            self.tree.selection_remove(item)

    def _default_recorder(self):
        if self.recorder_var.get().strip():
            return self.recorder_var.get().strip()
        return "班主任"

    def _validate(self, data):
        if not data["date"]:
            messagebox.showwarning("提示", "请填写沟通日期", parent=self.winfo_toplevel())
            return False
        if not data["content"]:
            messagebox.showwarning("提示", "请填写沟通内容", parent=self.winfo_toplevel())
            return False
        return True

    def add_log(self):
        sid = self.current_student_id()
        if not sid:
            messagebox.showwarning("提示", "请先选择学生", parent=self.winfo_toplevel())
            return
        data = self.get_form()
        if not self._validate(data):
            return
        try:
            add_communication_log(self.app.db_path, sid, data)
        except Exception as e:
            messagebox.showerror("错误", f"登记失败：{e}", parent=self.winfo_toplevel())
            return
        self.reload()
        self.clear_form()
        self.hint.config(text=f"已登记 {data['date']} 的沟通记录")
        if hasattr(self.app, "status_bar"):
            self.app.status_bar.config(text=f"已为 {sid} 新增一条家校沟通记录（{data['date']}）")

    def on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        log_id = int(sel[0])
        logs = {lg["id"]: lg for lg in
                fetch_communication_logs(self.app.db_path, self.current_student_id())}
        lg = logs.get(log_id)
        if not lg:
            return
        self.current_log_id = log_id
        self.date_var.set(lg.get("沟通日期") or "")
        self.method_var.set(lg.get("沟通方式") or COMM_METHODS[0])
        self.object_var.set(lg.get("沟通对象") or "")
        self.topic_var.set(lg.get("沟通主题") or COMM_TOPICS[0])
        self.recorder_var.set(lg.get("记录人") or "")
        self.content_text.delete("1.0", tk.END)
        self.content_text.insert("1.0", lg.get("沟通内容") or "")
        self.feedback_text.delete("1.0", tk.END)
        self.feedback_text.insert("1.0", lg.get("家长反馈") or "")
        self.followup_text.delete("1.0", tk.END)
        self.followup_text.insert("1.0", lg.get("后续跟进") or "")
        self.hint.config(text=f"正在修改 {lg.get('沟通日期') or ''} 的记录，改完点「保存修改」")

    def update_log(self):
        if not self.current_log_id:
            messagebox.showinfo("提示", "请先在上面的列表里选中一条记录",
                                parent=self.winfo_toplevel())
            return
        data = self.get_form()
        if not self._validate(data):
            return
        try:
            update_communication_log(self.app.db_path, self.current_log_id, data)
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{e}", parent=self.winfo_toplevel())
            return
        self.reload()
        self.clear_form()
        self.hint.config(text="已保存修改")

    def delete_log(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选中要删除的记录", parent=self.winfo_toplevel())
            return
        if not messagebox.askyesno("确认删除", f"确定删除选中的 {len(sel)} 条沟通记录？\n"
                                              f"删除后不可恢复。", parent=self.winfo_toplevel()):
            return
        for iid in sel:
            delete_communication_log(self.app.db_path, int(iid))
        self.reload()
        self.clear_form()
        self.hint.config(text="已删除选中记录")

    def export_logs(self):
        sid = self.current_student_id()
        logs = fetch_communication_logs(self.app.db_path, sid) if sid else []
        if not logs:
            messagebox.showinfo("提示", "没有可导出的记录", parent=self.winfo_toplevel())
            return
        conn = sqlite3.connect(self.app.db_path)
        try:
            name_map = {r[0]: r[1] for r in conn.execute("SELECT 学号, 姓名 FROM students")}
        finally:
            conn.close()
        name = name_map.get(sid, sid)
        path = filedialog.asksaveasfilename(
            title="导出家校沟通记录", defaultextension=".xlsx",
            initialfile=f"家校沟通记录_{name}_{date.today():%Y%m%d}.xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            export_communication_xlsx(path, logs, name_map)
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{e}", parent=self.winfo_toplevel())
            return
        messagebox.showinfo("完成", f"已导出 {len(logs)} 条记录：\n{path}",
                            parent=self.winfo_toplevel())


class CommunicationWindow:
    """家校沟通台账窗口：全班学生滚动登记。"""

    def __init__(self, app):
        self.app = app
        self.window = tk.Toplevel(app.root)
        self.window.withdraw()
        self.window.title("家校沟通台账")
        self.window.geometry("1120x780")
        self.window.configure(bg="#F4F5FA")
        self.window.transient(app.root)
        self.window.grab_set()

        header = tk.Frame(self.window, bg="#FFFFFF", padx=16, pady=12)
        header.pack(fill=tk.X)
        tk.Label(header, text="家校沟通台账", font=("Microsoft YaHei", 14, "bold"),
                 bg="#FFFFFF", fg="#30364B").pack(side=tk.LEFT)
        tk.Label(header, text="　选学生 → 填一条 → 新增登记；历史记录只增不改，"
                              "最近一次会自动同步到学生档案",
                 font=("Microsoft YaHei", 9), bg="#FFFFFF", fg="#8A90A6").pack(
            side=tk.LEFT, pady=(6, 0))

        body = tk.Frame(self.window, bg="#FFFFFF", highlightthickness=1,
                        highlightbackground="#E3E6EE")
        body.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        self.panel = CommunicationPanel(body, app, selectable=True, show_export=False)
        self.panel.pack(fill=tk.BOTH, expand=True)

        footer = tk.Frame(self.window, bg="#F4F5FA")
        footer.pack(fill=tk.X, padx=14, pady=(0, 12))
        app._flat_button(footer, "导出全班记录", self.export_all,
                         "#2F80ED", "#4A97F2", padx=18, pady=5).pack(side=tk.LEFT)
        app._flat_button(footer, "关闭", self.window.destroy,
                         "#6B7280", "#7D8694", padx=24, pady=5).pack(side=tk.RIGHT)

        center_window_relative(self.window, app.root)

    def export_all(self):
        logs = fetch_communication_logs(self.app.db_path)
        if not logs:
            messagebox.showinfo("提示", "台账里还没有记录", parent=self.window)
            return
        conn = sqlite3.connect(self.app.db_path)
        try:
            name_map = {r[0]: r[1] for r in conn.execute("SELECT 学号, 姓名 FROM students")}
        finally:
            conn.close()
        path = filedialog.asksaveasfilename(
            title="导出全班家校沟通记录", defaultextension=".xlsx",
            initialfile=f"家校沟通台账_全班_{date.today():%Y%m%d}.xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            export_communication_xlsx(path, logs, name_map)
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{e}", parent=self.window)
            return
        messagebox.showinfo("完成", f"已导出 {len(logs)} 条记录：\n{path}", parent=self.window)


class AIConfigDialog:
    """大模型接口配置（OpenAI 兼容），Key 仅保存在本机 ai_config.json。"""

    def __init__(self, parent_window, app, on_saved=None):
        self.app = app
        self.on_saved = on_saved
        cfg = dict(getattr(app, "ai_config", {}) or {})
        self.window = tk.Toplevel(app.root)
        self.window.withdraw()
        self.window.title("模型设置")
        self.window.geometry("580x380")
        self.window.configure(bg="#FFFFFF")
        self.window.transient(parent_window or app.root)
        self.window.grab_set()

        tk.Label(self.window, text="接入大模型（OpenAI 兼容接口）",
                 font=("Microsoft YaHei", 13, "bold"), bg="#FFFFFF", fg="#30364B").pack(
            anchor="w", padx=18, pady=(16, 4))
        tk.Label(self.window,
                 text="填好地址、模型名和 Key 即可用于「模型生成」。Key 只写在本机 "
                      "ai_config.json，不进数据库、不上传仓库。",
                 font=("Microsoft YaHei", 9), bg="#FFFFFF", fg="#8A90A6",
                 wraplength=540, justify="left").pack(anchor="w", padx=18)

        form = tk.Frame(self.window, bg="#FFFFFF")
        form.pack(fill=tk.X, padx=18, pady=12)
        form.grid_columnconfigure(1, weight=1)

        def lab(text, row):
            tk.Label(form, text=text, bg="#FFFFFF", fg="#39406B",
                     font=("Microsoft YaHei", 10)).grid(row=row, column=0, sticky="e",
                                                        padx=(0, 8), pady=5)

        def entry(text, row, show=None):
            e = tk.Entry(form, font=("Microsoft YaHei", 10), relief="flat",
                         highlightthickness=1, highlightbackground="#D5D9E8",
                         highlightcolor="#534AB7", show=show or "")
            e.insert(0, text or "")
            e.grid(row=row, column=1, sticky="ew", pady=5, ipady=3)
            return e

        self.preset_combo = ttk.Combobox(form, state="readonly",
                                         values=[p[0] for p in AI_PRESETS])
        preset_name = cfg.get("preset") or AI_PRESETS[0][0]
        if preset_name not in [p[0] for p in AI_PRESETS]:
            preset_name = AI_PRESETS[0][0]
        self.preset_combo.set(preset_name)
        lab("服务商", 0)
        self.preset_combo.grid(row=0, column=1, sticky="w", pady=5)
        self.preset_combo.bind("<<ComboboxSelected>>", self.apply_preset)

        lab("接口地址", 1)
        self.base_entry = entry(cfg.get("base_url") or "", 1)
        lab("模型名", 2)
        self.model_entry = entry(cfg.get("model") or "", 2)
        lab("API Key", 3)
        self.key_entry = entry(cfg.get("api_key") or "", 3, show="*")

        self.tip = tk.Label(self.window, text="", bg="#FFFFFF", fg="#8A90A6",
                            font=("Microsoft YaHei", 9), wraplength=540, justify="left")
        self.tip.pack(anchor="w", padx=18)

        btns = tk.Frame(self.window, bg="#FFFFFF")
        btns.pack(fill=tk.X, padx=18, pady=14)
        app._flat_button(btns, "保存", self.save, "#534AB7", "#6A5FD0",
                         padx=24, pady=5).pack(side=tk.RIGHT)
        app._flat_button(btns, "取消", self.window.destroy, "#6B7280", "#7D8694",
                         padx=20, pady=5).pack(side=tk.RIGHT, padx=8)
        app._flat_button(btns, "测试连接", self.test, "#2F80ED", "#4A97F2",
                         padx=20, pady=5).pack(side=tk.LEFT)

        if preset_name != "自定义":
            self.apply_preset()
            if cfg.get("base_url"):
                self.base_entry.delete(0, tk.END)
                self.base_entry.insert(0, cfg["base_url"])
            if cfg.get("model"):
                self.model_entry.delete(0, tk.END)
                self.model_entry.insert(0, cfg["model"])
        center_window_relative(self.window, app.root)

    def apply_preset(self, event=None):
        name = self.preset_combo.get()
        for pname, base, model in AI_PRESETS:
            if pname == name and base:
                self.base_entry.delete(0, tk.END)
                self.base_entry.insert(0, base)
                self.model_entry.delete(0, tk.END)
                self.model_entry.insert(0, model)
                self.tip.config(text=f"已填入 {name} 的默认地址与模型，Key 需自行申请。")
                return
        self.tip.config(text="自定义：地址形如 https://xxx/v1（不要带 /chat/completions）。")

    def _form_cfg(self):
        return {"preset": self.preset_combo.get(),
                "base_url": self.base_entry.get().strip(),
                "model": self.model_entry.get().strip(),
                "api_key": self.key_entry.get().strip()}

    def save(self):
        cfg = self._form_cfg()
        if not cfg["base_url"].startswith("http"):
            messagebox.showwarning("提示", "接口地址要以 http:// 或 https:// 开头",
                                   parent=self.window)
            return
        if not cfg["model"]:
            messagebox.showwarning("提示", "请填写模型名", parent=self.window)
            return
        try:
            path = save_ai_config(self.app.base_path, cfg)
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{e}", parent=self.window)
            return
        self.app.ai_config = cfg
        if self.on_saved:
            self.on_saved()
        messagebox.showinfo("已保存", f"配置已写入：\n{path}", parent=self.window)
        self.window.destroy()

    def test(self):
        cfg = self._form_cfg()
        if not (cfg["base_url"] and cfg["api_key"] and cfg["model"]):
            messagebox.showwarning("提示", "地址、模型名、Key 三项都要填",
                                   parent=self.window)
            return
        self.tip.config(text="正在测试连接…")
        self.window.update_idletasks()
        try:
            reply = call_llm(cfg, "请只回复四个字：连接正常", timeout=30)
            self.tip.config(text="连接正常：" + reply[:60].replace("\n", " "))
        except Exception as e:
            self.tip.config(text="连接失败：" + str(e)[:180])


class BatchReportDialog:
    """按班级批量生成学生情况报告。"""

    def __init__(self, app):
        self.app = app
        self.running = False
        self.cancel_flag = False
        self.items = []
        self._queue = queue.Queue()

        self.window = tk.Toplevel(app.root)
        self.window.withdraw()
        self.window.title("批量生成学生情况报告")
        self.window.geometry("780x600")
        self.window.configure(bg="#F4F5FA")
        self.window.transient(app.root)
        self.window.grab_set()

        head = tk.Frame(self.window, bg="#FFFFFF", padx=16, pady=12)
        head.pack(fill=tk.X)
        tk.Label(head, text="批量生成学生情况报告", font=("Microsoft YaHei", 14, "bold"),
                 bg="#FFFFFF", fg="#30364B").pack(anchor="w")
        tk.Label(head, text="每个学生一份「概况 / 优劣势 / 家校沟通注意点 / 管理注意点」，"
                            "生成后自动存入档案并导出 Excel。",
                 font=("Microsoft YaHei", 9), bg="#FFFFFF", fg="#8A90A6").pack(anchor="w")

        opts = tk.Frame(self.window, bg="#FFFFFF", padx=16, pady=10)
        opts.pack(fill=tk.X, padx=14, pady=(10, 6))
        self.mode_var = tk.StringVar(value="local")
        tk.Label(opts, text="生成方式：", bg="#FFFFFF", fg="#39406B",
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        tk.Radiobutton(opts, text="离线规则（快，不联网）", variable=self.mode_var,
                       value="local", bg="#FFFFFF", activebackground="#FFFFFF").pack(side=tk.LEFT)
        tk.Radiobutton(opts, text="大模型（逐人调用，需已配置）", variable=self.mode_var,
                       value="ai", bg="#FFFFFF", activebackground="#FFFFFF").pack(
            side=tk.LEFT, padx=10)
        self.export_var = tk.BooleanVar(value=True)
        tk.Checkbutton(opts, text="生成后导出 Excel", variable=self.export_var,
                       bg="#FFFFFF", activebackground="#FFFFFF").pack(side=tk.RIGHT)

        self.log_text = tk.Text(self.window, height=16, wrap="word",
                                font=("Microsoft YaHei", 9), relief="flat",
                                highlightthickness=1, highlightbackground="#E3E6EE",
                                bg="#FCFCFE", padx=10, pady=8)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=14)
        self.log_text.insert("1.0", "准备好后点「开始生成」。\n")
        self.log_text.config(state="disabled")

        self.progress = ttk.Progressbar(self.window, mode="determinate", maximum=100)
        self.progress.pack(fill=tk.X, padx=14, pady=(8, 0))

        footer = tk.Frame(self.window, bg="#F4F5FA")
        footer.pack(fill=tk.X, padx=14, pady=12)
        self.start_btn = app._flat_button(footer, "开始生成", self.start,
                                          "#1A7F4B", "#219663", padx=22, pady=5)
        self.start_btn.pack(side=tk.LEFT)
        self.stop_btn = app._flat_button(footer, "停止", self.stop,
                                         "#E05252", "#E96A6A", padx=18, pady=5)
        self.stop_btn.pack(side=tk.LEFT, padx=8)
        app._flat_button(footer, "关闭", self.window.destroy,
                         "#6B7280", "#7D8694", padx=22, pady=5).pack(side=tk.RIGHT)
        center_window_relative(self.window, app.root)

    def log(self, text):
        """线程安全：后台线程只投递消息，界面由主线程的 _poll 更新。"""
        self._queue.put(("log", text))

    def _append(self, text):
        try:
            self.log_text.config(state="normal")
            self.log_text.insert(tk.END, text + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state="disabled")
        except Exception:
            pass

    def _poll(self):
        if not self.window.winfo_exists():
            return
        try:
            while True:
                item = self._queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    self._append(item[1])
                elif kind == "progress":
                    self.progress.configure(value=item[1])
                elif kind == "done":
                    self._finish(*item[1:])
        except queue.Empty:
            pass
        except Exception:
            pass
        if self.running or not self._queue.empty():
            try:
                self.window.after(120, self._poll)
            except Exception:
                pass

    def stop(self):
        self.cancel_flag = True
        self.log("已请求停止，当前这名学生生成完就停。")

    def start(self):
        if self.running:
            return
        use_ai = self.mode_var.get() == "ai"
        cfg = self.app.ai_config or {}
        if use_ai and not (cfg.get("base_url") and cfg.get("api_key") and cfg.get("model")):
            messagebox.showwarning(
                "未配置模型",
                "还没填大模型接口。\n\n请先在学生画像的「情况报告」页点「模型设置」，"
                "或改用离线规则生成。", parent=self.window)
            return
        conn = sqlite3.connect(self.app.db_path)
        try:
            students = conn.execute("SELECT 学号, 姓名 FROM students ORDER BY 学号").fetchall()
        finally:
            conn.close()
        if not students:
            messagebox.showinfo("提示", "数据库里没有学生", parent=self.window)
            return
        if not messagebox.askyesno(
                "确认", f"将为 {len(students)} 名学生生成情况报告"
                        + ("（逐人调用大模型，可能较慢）" if use_ai else "（离线规则，几秒钟）")
                        + "，并存入档案。继续吗？", parent=self.window):
            return
        self.running = True
        self.cancel_flag = False
        self.items = []
        self._queue = queue.Queue()
        self.progress["maximum"] = len(students)
        self.progress["value"] = 0
        self.start_btn.config(state="disabled")
        do_export = bool(self.export_var.get())
        self.window.after(120, self._poll)
        threading.Thread(target=self._worker,
                         args=(students, use_ai, cfg, do_export),
                         daemon=True).start()

    def _worker(self, students, use_ai, cfg, do_export):
        total = len(students)
        ok = 0
        for idx, (sid, name) in enumerate(students, 1):
            if self.cancel_flag:
                self.log("已停止。")
                break
            try:
                ctx = collect_student_context(self.app.db_path, sid)
                if not ctx:
                    self.log(f"[{idx}/{total}] {name}：无档案数据，跳过")
                    continue
                if use_ai:
                    report = call_llm(cfg, build_ai_prompt(ctx), timeout=120)
                    source = "模型：" + (cfg.get("model") or "")
                else:
                    report = generate_local_report(ctx)
                    source = "离线规则"
                save_report(self.app.db_path, sid, report, source)
                self.items.append({"ctx": ctx, "report": report, "source": source})
                ok += 1
                self.log(f"[{idx}/{total}] {name}：" + ("模型生成" if use_ai else "离线生成")
                         + f" {len(report)} 字")
            except Exception as e:
                self.log(f"[{idx}/{total}] {name}：失败 - {str(e)[:120]}")
            self._queue.put(("progress", idx))
        path = None
        if do_export and self.items:
            try:
                desk = os.path.join(os.path.expanduser("~"), "Desktop")
                if not os.path.isdir(desk):
                    desk = self.app.base_path
                path = os.path.join(desk, f"班级学生情况报告_{datetime.now():%Y%m%d_%H%M}.xlsx")
                export_reports_xlsx(path, self.items)
            except Exception as e:
                self.log("导出 Excel 失败：" + str(e))
                path = None
        self._queue.put(("done", ok, total, path))

    def _finish(self, ok, total, path):
        self.running = False
        self.start_btn.config(state="normal")
        self._append(f"完成：成功 {ok} / 共 {total} 人。")
        if path:
            self._append("已导出：" + path)
        if hasattr(self.app, "status_bar"):
            self.app.status_bar.config(text=f"批量生成完成：{ok}/{total} 人")


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
        self.window.geometry("1040x820")
        self.window.configure(bg="#F4F5FA")
        self.window.transient(parent.root)
        self.window.grab_set()

        self.entry_vars = {}

        if not self.load_student_data():
            self.window.destroy()
            return

        self.column_groups = self.parent.get_columns_grouped_by_category(self.columns)
        # 「最近一次谈话」两列已由沟通台账接管：分类页不再重复展示，避免两处录入
        filtered = []
        for cat, cols in self.column_groups:
            kept = [c for c in cols if "谈话" not in c]
            if kept:
                filtered.append((cat, kept))
        self.column_groups = filtered
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
                ("走读", "是否走读"), ("红黄蓝", "红黄蓝情况简述"),
                ("学生会", "是否学生会")]

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

        # 家校沟通次数（来自沟通台账，不是学生表里的列）
        try:
            log_count = len(fetch_communication_logs(self.parent.db_path, self.student_id))
        except Exception:
            log_count = 0
        tk.Label(chip_row, text=f" 沟通  {log_count} 次 ",
                 bg="#E7F3EC" if log_count else "#FDECEC",
                 fg="#1A7F4B" if log_count else "#B4433F",
                 padx=10, pady=4).pack(side=tk.LEFT, padx=(0, 8))

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
    # ------------------------------------------------------------------
    # 家校沟通台账（v8）
    # ------------------------------------------------------------------
    def build_communication_tab(self, notebook):
        frame = tk.Frame(notebook, bg="#FFFFFF")
        notebook.add(frame, text="沟通记录")
        self.comm_frame = frame
        self.comm_panel = CommunicationPanel(frame, self.parent,
                                             student_id=self.student_id)
        self.comm_panel.pack(fill=tk.BOTH, expand=True)

    # ------------------------------------------------------------------
    # 学生情况报告（v8）
    # ------------------------------------------------------------------
    def build_report_tab(self, notebook):
        frame = tk.Frame(notebook, bg="#FFFFFF")
        notebook.add(frame, text="情况报告")
        self.report_frame = frame
        p = self.parent

        bar = tk.Frame(frame, bg="#FFFFFF")
        bar.pack(fill=tk.X, padx=10, pady=(10, 4))
        p._flat_button(bar, "模型生成", self.generate_report_ai,
                       "#534AB7", "#6A5FD0", padx=14, pady=4).pack(side=tk.LEFT)
        p._flat_button(bar, "离线生成", self.generate_report_local,
                       "#1A7F4B", "#219663", padx=14, pady=4).pack(side=tk.LEFT, padx=6)
        p._flat_button(bar, "复制提示词", self.copy_prompt,
                       "#2F80ED", "#4A97F2", padx=14, pady=4).pack(side=tk.LEFT)
        p._flat_button(bar, "模型设置", self.open_ai_config,
                       "#8E24AA", "#A03BC2", padx=14, pady=4).pack(side=tk.LEFT, padx=6)
        p._flat_button(bar, "导出TXT", self.export_report_txt,
                       "#6B7280", "#7D8694", padx=14, pady=4).pack(side=tk.RIGHT)
        p._flat_button(bar, "保存到档案", self.save_report_to_db,
                       "#0F8B8D", "#14A0A2", padx=14, pady=4).pack(side=tk.RIGHT, padx=6)

        self.report_status = tk.Label(frame, text="", bg="#FFFFFF", fg="#8A90A6",
                                      font=("Microsoft YaHei", 9), anchor="w")
        self.report_status.pack(fill=tk.X, padx=12)

        box = tk.Frame(frame, bg="#FFFFFF")
        box.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 10))
        self.report_text = tk.Text(box, wrap="word", font=("Microsoft YaHei", 10),
                                   relief="flat", highlightthickness=1,
                                   highlightbackground="#E3E6EE", bg="#FCFCFE",
                                   padx=12, pady=10, spacing1=2, spacing3=2)
        sb = tk.Scrollbar(box, orient=tk.VERTICAL, command=self.report_text.yview)
        self.report_text.configure(yscrollcommand=sb.set)
        self.report_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.load_latest_report()

    def _report_set(self, text):
        self.report_text.delete("1.0", tk.END)
        self.report_text.insert("1.0", text)

    def _report_get(self):
        return self.report_text.get("1.0", tk.END).strip()

    def _report_note(self, text):
        try:
            self.report_status.config(text=text)
        except Exception:
            pass

    def load_latest_report(self):
        rec = fetch_latest_report(self.parent.db_path, self.student_id)
        if rec:
            self._report_set(rec["内容"])
            self._report_note(f"已载入最近一次报告：{rec['生成时间']}（{rec['来源']}）"
                              f"· 可直接修改后点「保存到档案」")
        else:
            self._report_set(
                "尚未生成学生情况。\n\n"
                "· 「离线生成」：按档案数据由规则引擎直接出结果，不联网；\n"
                "· 「模型生成」：由大模型成文，需先在「模型设置」填接口地址、模型名和 Key；\n"
                "· 「复制提示词」：把同一份学生数据复制走，可贴到 WorkBuddy 等对话里生成。")
            self._report_note("暂无报告")

    def _context(self):
        return collect_student_context(self.parent.db_path, self.student_id)

    def generate_report_local(self):
        ctx = self._context()
        if not ctx:
            messagebox.showwarning("提示", "没有找到该学生的档案数据", parent=self.window)
            return
        content = generate_local_report(ctx)
        self._report_set(content)
        save_report(self.parent.db_path, self.student_id, content, "离线规则")
        self._report_note("已按档案数据离线生成并存入档案 · " + _now_str())

    def generate_report_ai(self):
        cfg = self.parent.ai_config or {}
        if not (cfg.get("base_url") and cfg.get("api_key") and cfg.get("model")):
            if messagebox.askyesno(
                    "尚未配置模型",
                    "还没有填写大模型接口地址和 Key。\n\n现在打开「模型设置」吗？\n"
                    "也可以先用「离线生成」，或「复制提示词」贴到 WorkBuddy 里生成。",
                    parent=self.window):
                self.open_ai_config()
            return
        ctx = self._context()
        if not ctx:
            messagebox.showwarning("提示", "没有找到该学生的档案数据", parent=self.window)
            return
        prompt = build_ai_prompt(ctx)
        self._report_set("正在调用 " + (cfg.get("model") or "模型") + " 生成，请稍候…\n\n"
                         "（首次调用可能要十几秒；如需中止可关闭窗口后再打开）")
        self._report_note("正在生成…")
        self._ai_queue = queue.Queue()
        self.window.after(120, self._poll_ai)
        threading.Thread(target=self._ai_worker, args=(dict(cfg), prompt),
                         daemon=True).start()

    def _ai_worker(self, cfg, prompt):
        """后台线程只做网络请求，界面更新全部交回主线程。"""
        try:
            text = call_llm(cfg, prompt)
            self._ai_queue.put(("ok", text, cfg))
        except Exception as e:
            self._ai_queue.put(("err", str(e), None))

    def _poll_ai(self):
        try:
            kind, payload, cfg = self._ai_queue.get_nowait()
        except queue.Empty:
            try:
                self.window.after(120, self._poll_ai)
            except Exception:
                pass
            return
        if kind == "ok":
            self._ai_done(payload, cfg)
        else:
            self._ai_failed(payload)

    def _ai_done(self, text, cfg):
        if not self.window.winfo_exists():
            return
        self._report_set(text)
        save_report(self.parent.db_path, self.student_id, text,
                    "模型：" + (cfg.get("model") or ""))
        self._report_note(f"已由 {cfg.get('model') or '模型'} 生成并存入档案 · {_now_str()}")

    def _ai_failed(self, msg):
        if not self.window.winfo_exists():
            return
        self._report_set("生成失败：\n" + msg +
                         "\n\n可以改用「离线生成」，或点「复制提示词」贴到 WorkBuddy 等对话里生成。")
        self._report_note("生成失败")

    def copy_prompt(self):
        ctx = self._context()
        if not ctx:
            return
        prompt = build_ai_prompt(ctx)
        try:
            self.window.clipboard_clear()
            self.window.clipboard_append(prompt)
            self._report_note("提示词已复制到剪贴板，可直接粘贴给 WorkBuddy / 豆包等对话")
        except Exception as e:
            messagebox.showerror("错误", f"复制失败：{e}", parent=self.window)

    def save_report_to_db(self):
        content = self._report_get()
        if not content:
            messagebox.showinfo("提示", "还没有内容可保存", parent=self.window)
            return
        save_report(self.parent.db_path, self.student_id, content, "手工编辑")
        self._report_note("已保存到档案 · " + _now_str())

    def export_report_txt(self):
        content = self._report_get()
        if not content:
            messagebox.showinfo("提示", "还没有内容可导出", parent=self.window)
            return
        name = str(self.data.get("姓名") or self.student_id)
        path = filedialog.asksaveasfilename(
            title="导出学生情况", defaultextension=".txt",
            initialfile=f"{name}_学生情况_{date.today():%Y%m%d}.txt",
            filetypes=[("文本文件", "*.txt")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig") as f:
                f.write(content)
        except Exception as e:
            messagebox.showerror("错误", f"写入失败：{e}", parent=self.window)
            return
        messagebox.showinfo("完成", f"已导出：\n{path}", parent=self.window)

    def open_ai_config(self):
        AIConfigDialog(self.window, self.parent,
                       on_saved=lambda: self._report_note("模型配置已更新"))

    def quick_generate(self):
        """底部「生成学生情况」：配了模型就走模型，没配就用离线规则。"""
        try:
            self.notebook.select(self.report_frame)
        except Exception:
            pass
        cfg = self.parent.ai_config or {}
        if cfg.get("base_url") and cfg.get("api_key") and cfg.get("model"):
            self.generate_report_ai()
        else:
            self.generate_report_local()

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
        self.notebook = notebook

        self.build_academic_tab(notebook)
        self.build_communication_tab(notebook)
        self.build_report_tab(notebook)

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

        self.parent._flat_button(btn_frame, "生成学生情况", self.quick_generate,
                                 "#0F8B8D", "#14A0A2", padx=26, pady=6,
                                 font=("Microsoft YaHei", 11)).pack(
            side=tk.LEFT, expand=True, padx=10)
        self.parent._flat_button(btn_frame, "保存修改", self.save_changes,
                                 "#534AB7", "#6A5FD0", padx=26, pady=6,
                                 font=("Microsoft YaHei", 11)).pack(
            side=tk.LEFT, expand=True, padx=10)
        self.parent._flat_button(btn_frame, "关闭", self.window.destroy,
                                 "#6B7280", "#7D8694", padx=26, pady=6,
                                 font=("Microsoft YaHei", 11)).pack(
            side=tk.LEFT, expand=True, padx=10)

    def save_changes(self):
        try:
            conn = sqlite3.connect(self.parent.db_path)
            cursor = conn.cursor()
            # 只更新这次真正展示出来的字段，没展示的列（如谈话两列）保持原值
            update_columns = [col for col in self.columns
                              if col != "学号" and col in self.entry_vars]
            if not update_columns:
                conn.close()
                messagebox.showinfo("提示", "没有可保存的字段", parent=self.window)
                return
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