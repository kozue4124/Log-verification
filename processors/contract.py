from typing import Optional, List, Tuple
"""雇用契約書（デフォルト勤務時間）読み込みプロセッサ"""
import pandas as pd
import pdfplumber
import chardet
import re
from datetime import time
from dateutil import parser as date_parser
from pathlib import Path


COLUMN_ALIASES = {
    "employee_id": ["社員番号", "社員ID", "従業員番号", "従業員ID", "ID", "id", "employee_id", "emp_id"],
    "employee_name": ["氏名", "名前", "社員名", "従業員名", "name", "employee_name"],
    "work_start": ["勤務開始", "始業時刻", "開始時刻", "work_start", "start_time", "出勤時刻", "始業", "所定始業"],
    "work_end": ["勤務終了", "終業時刻", "終了時刻", "work_end", "end_time", "退勤時刻", "終業", "所定終業"],
    "work_days": ["勤務曜日", "勤務日", "work_days", "曜日", "勤務形態"],
}

# 曜日マッピング
DAY_MAP = {
    "月": 0, "火": 1, "水": 2, "木": 3, "金": 4, "土": 5, "日": 6,
    "mon": 0, "tue": 1, "wed": 2, "thu": 3, "fri": 4, "sat": 5, "sun": 6,
    "monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3,
    "friday": 4, "saturday": 5, "sunday": 6,
}


def _parse_time(val) -> Optional[time]:
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass
    if isinstance(val, time):
        return val
    s = str(val).strip()
    if not s or s in ("nan", "None"):
        return None
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", s)
    if m:
        h, mi, sec = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        try:
            return time(h, mi, sec)
        except ValueError:
            return None
    try:
        return date_parser.parse(s).time()
    except Exception:
        return None


def _parse_work_days(val) -> Optional[List[int]]:
    """勤務曜日をweekday整数リストに変換（0=月〜6=日）"""
    if val is None:
        return None
    s = str(val).strip().lower()
    if not s or s in ("nan", "none"):
        return None

    # "月〜金" や "月-金" パターン
    range_match = re.search(r"([月火水木金土日])[〜~\-ー]([月火水木金土日])", s)
    if range_match:
        start_day = DAY_MAP.get(range_match.group(1))
        end_day = DAY_MAP.get(range_match.group(2))
        if start_day is not None and end_day is not None:
            if start_day <= end_day:
                return list(range(start_day, end_day + 1))

    # 個別の曜日を検索
    days = []
    for day_str, day_num in DAY_MAP.items():
        if day_str in s and day_num not in days:
            days.append(day_num)

    return sorted(days) if days else None


def _find_column(df: pd.DataFrame, field: str) -> Optional[str]:
    for alias in COLUMN_ALIASES.get(field, []):
        for col in df.columns:
            if str(col).strip() == alias or str(col).strip().lower() == alias.lower():
                return col
    for alias in COLUMN_ALIASES.get(field, []):
        for col in df.columns:
            if alias in str(col) or str(col) in alias:
                return col
    return None


def _df_to_contracts(df: pd.DataFrame) -> List[dict]:
    df = df.dropna(how="all").reset_index(drop=True)
    col_map = {}
    for field in COLUMN_ALIASES:
        col = _find_column(df, field)
        if col:
            col_map[field] = col

    contracts = []
    for _, row in df.iterrows():
        emp_id = str(row[col_map["employee_id"]]).strip() if "employee_id" in col_map else None
        emp_name = str(row[col_map["employee_name"]]).strip() if "employee_name" in col_map else None
        if emp_id in (None, "", "nan", "None") and emp_name in (None, "", "nan", "None"):
            continue

        work_days_raw = row[col_map["work_days"]] if "work_days" in col_map else None
        contracts.append({
            "employee_id": emp_id,
            "employee_name": emp_name,
            "work_start": _parse_time(row[col_map["work_start"]]) if "work_start" in col_map else None,
            "work_end": _parse_time(row[col_map["work_end"]]) if "work_end" in col_map else None,
            "work_days": _parse_work_days(work_days_raw),  # None=全日
        })
    return contracts


def _load_pdf_contracts(filepath: str) -> List[dict]:
    """PDFから雇用契約書情報を抽出"""
    contracts = []
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    header = [str(h).strip() if h else "" for h in table[0]]
                    df = pd.DataFrame(table[1:], columns=header)
                    contracts.extend(_df_to_contracts(df))

                if not contracts:
                    # テキストから抽出
                    text = page.extract_text() or ""
                    # 始業・終業パターン
                    start_match = re.search(r"始業[時刻]*[：:]\s*(\d{1,2}:\d{2})", text)
                    end_match = re.search(r"終業[時刻]*[：:]\s*(\d{1,2}:\d{2})", text)
                    name_match = re.search(r"氏名[：:]\s*([^\s\n]+)", text)
                    id_match = re.search(r"社員番号[：:]\s*([^\s\n]+)", text)

                    if start_match or end_match:
                        contracts.append({
                            "employee_id": id_match.group(1) if id_match else None,
                            "employee_name": name_match.group(1) if name_match else None,
                            "work_start": _parse_time(start_match.group(1)) if start_match else None,
                            "work_end": _parse_time(end_match.group(1)) if end_match else None,
                            "work_days": None,
                        })
    except Exception as e:
        raise ValueError(f"PDF雇用契約書の読み込みに失敗: {e}")
    return contracts


def load_contracts(filepath: str) -> List[dict]:
    """
    雇用契約書を読み込む
    各要素: {employee_id, employee_name, work_start, work_end, work_days}
    work_days: [0,1,2,3,4] = 月〜金、None = 全日
    """
    ext = Path(filepath).suffix.lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        df = pd.read_excel(filepath, dtype=str)
        return _df_to_contracts(df)
    elif ext == ".csv":
        with open(filepath, "rb") as f:
            raw = f.read()
        detected = chardet.detect(raw)
        encoding = detected.get("encoding") or "utf-8-sig"
        try:
            df = pd.read_csv(filepath, dtype=str, encoding=encoding)
        except Exception:
            df = pd.read_csv(filepath, dtype=str, encoding="cp932")
        return _df_to_contracts(df)
    elif ext == ".pdf":
        return _load_pdf_contracts(filepath)
    else:
        try:
            df = pd.read_excel(filepath, dtype=str)
            return _df_to_contracts(df)
        except Exception:
            try:
                df = pd.read_csv(filepath, dtype=str, encoding="utf-8-sig")
                return _df_to_contracts(df)
            except Exception as e:
                raise ValueError(f"雇用契約書の読み込みに失敗: {e}")
