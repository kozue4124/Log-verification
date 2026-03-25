from typing import Optional, List, Tuple
"""出勤記録読み込みプロセッサ（Excel / CSV / PDF）"""
import pandas as pd
import pdfplumber
import chardet
import re
import io
from datetime import datetime, time, date
from dateutil import parser as date_parser
from pathlib import Path


COLUMN_ALIASES = {
    "employee_id": ["社員番号", "社員ID", "従業員番号", "従業員ID", "ID", "id", "employee_id", "emp_id", "staff_id", "社員No", "No"],
    "employee_name": ["氏名", "名前", "社員名", "従業員名", "name", "employee_name", "full_name"],
    "date": ["日付", "出勤日", "date", "勤務日", "日", "Date", "年月日"],
    "clock_in": ["出勤時刻", "出勤", "始業", "開始時刻", "clock_in", "start", "出社時刻", "打刻（出）", "出勤打刻", "IN", "in"],
    "clock_out": ["退勤時刻", "退勤", "終業", "終了時刻", "clock_out", "end", "退社時刻", "打刻（退）", "退勤打刻", "OUT", "out"],
}


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


def _parse_time(val) -> Optional[time]:
    if pd.isna(val) if hasattr(pd, "isna") else val is None:
        return None
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, datetime):
        return val.time()
    s = str(val).strip()
    if not s or s in ("nan", "None", "-", "－", ""):
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


def _parse_date(val) -> Optional[date]:
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if hasattr(val, "date"):
        return val.date()
    s = str(val).strip()
    if not s or s in ("nan", "None"):
        return None
    try:
        return date_parser.parse(s, dayfirst=False).date()
    except Exception:
        return None


def _df_to_records(df: pd.DataFrame) -> List[dict]:
    df = df.dropna(how="all").reset_index(drop=True)
    col_map = {}
    for field in COLUMN_ALIASES:
        col = _find_column(df, field)
        if col:
            col_map[field] = col

    records = []
    for idx, row in df.iterrows():
        _EMPTY = (None, "", "nan", "None", "-", "－")
        emp_id_raw = str(row[col_map["employee_id"]]).strip() if "employee_id" in col_map else None
        emp_name_raw = str(row[col_map["employee_name"]]).strip() if "employee_name" in col_map else None
        emp_id = None if emp_id_raw in _EMPTY else emp_id_raw
        emp_name = None if emp_name_raw in _EMPTY else emp_name_raw
        if emp_id is None and emp_name is None:
            continue
        rec = {
            "employee_id": emp_id,
            "employee_name": emp_name,
            "date": _parse_date(row[col_map["date"]]) if "date" in col_map else None,
            "clock_in": _parse_time(row[col_map["clock_in"]]) if "clock_in" in col_map else None,
            "clock_out": _parse_time(row[col_map["clock_out"]]) if "clock_out" in col_map else None,
        }
        records.append(rec)
    return records


def _load_excel(filepath: str) -> List[dict]:
    try:
        df = pd.read_excel(filepath, dtype=str)
    except Exception as e:
        raise ValueError(f"Excelの読み込みに失敗: {e}")
    return _df_to_records(df)


def _load_csv(filepath: str) -> List[dict]:
    # エンコーディング自動検出
    with open(filepath, "rb") as f:
        raw = f.read()
    detected = chardet.detect(raw)
    encoding = detected.get("encoding") or "utf-8"
    # BOM付きUTF-8対応
    if encoding.lower() in ("utf-8-sig", "utf-8"):
        encoding = "utf-8-sig"
    try:
        df = pd.read_csv(filepath, dtype=str, encoding=encoding)
    except Exception:
        try:
            df = pd.read_csv(filepath, dtype=str, encoding="cp932")
        except Exception as e:
            raise ValueError(f"CSVの読み込みに失敗: {e}")
    return _df_to_records(df)


def _parse_kinmu_pdf(table: list, full_text: str) -> Optional[List[dict]]:
    """
    「月・日付・始業時刻・終業時刻」形式の出勤簿PDFをパース。
    月列が途中から省略され、氏名・年月がフッターに記載される形式に対応。
    """
    # ヘッダー行を特定（「月」「日付」「始業時刻」が含まれる行）
    header_idx = None
    for i, row in enumerate(table):
        cells = [str(c).strip() if c else "" for c in row]
        if "日付" in cells and ("始業時刻" in cells or "終業時刻" in cells):
            header_idx = i
            break
    if header_idx is None:
        return None

    header = [str(c).strip() if c else "" for c in table[header_idx]]

    # 必要列のインデックスを取得
    def col(names):
        for n in names:
            for i, h in enumerate(header):
                if n in h:
                    return i
        return None

    idx_month   = col(["月"])
    idx_day     = col(["日付"])
    idx_in      = col(["始業時刻"])
    idx_out     = col(["終業時刻"])
    idx_count   = col(["出勤日数"])

    if idx_day is None or idx_in is None or idx_out is None:
        return None

    # 氏名をフッターテキストから抽出（「氏名 〇〇 〇〇」パターン）
    emp_name = None
    m = re.search(r"氏名\s*([^\s　]+\s*[^\s　]+)", full_text)
    if m:
        emp_name = re.sub(r"\s+", "", m.group(1)).strip()

    # 年・月をフッターテキストから抽出（「2026年 3 月」など）
    year, month = None, None
    ym = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", full_text)
    if ym:
        year, month = int(ym.group(1)), int(ym.group(2))
    if year is None:
        from datetime import datetime
        year = datetime.today().year

    records = []
    current_month = month  # フッターから取得した主月をデフォルトに
    prev_month = None

    for row in table[header_idx + 1:]:
        cells = [str(c).strip() if c else "" for c in row]
        if len(cells) <= max(filter(None, [idx_day, idx_in, idx_out])):
            continue

        # 月列の更新（空白は前行の月を引き継ぐ）
        raw_month = cells[idx_month].strip() if idx_month is not None and idx_month < len(cells) else ""
        if re.match(r"^\d+$", raw_month):
            new_month = int(raw_month)
            if new_month != prev_month:
                current_month = new_month
                prev_month = new_month

        # 日付
        raw_day = cells[idx_day] if idx_day < len(cells) else ""
        if not re.match(r"^\d+$", raw_day):
            continue  # 合計行・ヘッダー行をスキップ
        day = int(raw_day)

        # 出勤日数が 0 → 休み・欠勤（時刻なしでも記録として追加）
        raw_count = cells[idx_count].strip() if idx_count is not None and idx_count < len(cells) else ""

        # 時刻
        raw_in  = cells[idx_in]  if idx_in  < len(cells) else ""
        raw_out = cells[idx_out] if idx_out < len(cells) else ""
        clock_in  = _parse_time(raw_in)
        clock_out = _parse_time(raw_out)

        # 年の推定（月が前月より小さくなったら翌年）
        rec_year = year
        if current_month and month and current_month < month - 1:
            rec_year = year + 1

        try:
            from datetime import date as date_cls
            d = date_cls(rec_year, current_month or month, day)
        except (ValueError, TypeError):
            continue

        records.append({
            "employee_id":   None,
            "employee_name": emp_name,
            "date":          d,
            "clock_in":      clock_in,
            "clock_out":     clock_out,
        })

    return records if records else None


def _load_pdf(filepath: str) -> List[dict]:
    """PDFからテーブルを抽出して出勤記録を読み込む"""
    records = []
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                full_text = page.extract_text() or ""
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    # まず専用パーサーを試みる（月+日付形式の出勤簿）
                    kinmu = _parse_kinmu_pdf(table, full_text)
                    if kinmu:
                        records.extend(kinmu)
                        continue

                    # 汎用パーサー（標準的な列名が揃っている場合）
                    header = [str(h).strip() if h else "" for h in table[0]]
                    df = pd.DataFrame(table[1:], columns=header)
                    page_records = _df_to_records(df)
                    records.extend(page_records)

    except Exception as e:
        raise ValueError(f"PDFの読み込みに失敗: {e}")

    if not records:
        records = _parse_pdf_text(filepath)

    return records


def _parse_pdf_text(filepath: str) -> List[dict]:
    """PDFのテキストから出勤記録を正規表現でパース（フォールバック）"""
    records = []
    try:
        with pdfplumber.open(filepath) as pdf:
            full_text = ""
            for page in pdf.pages:
                full_text += page.extract_text() or ""

        # 日付 + 時刻パターンを検索
        # 例: "2024/01/15  09:00  18:00" or "山田太郎 2024/01/15 09:00 18:00"
        lines = full_text.split("\n")
        for line in lines:
            line = line.strip()
            if not line:
                continue
            # 日付パターン検索
            date_match = re.search(r"(\d{4}[/\-年]\d{1,2}[/\-月]\d{1,2}日?)", line)
            times = re.findall(r"(\d{1,2}:\d{2})", line)
            if date_match and len(times) >= 2:
                try:
                    d = date_parser.parse(date_match.group(1)).date()
                    clock_in = _parse_time(times[0])
                    clock_out = _parse_time(times[1])
                    records.append({
                        "employee_id": None,
                        "employee_name": None,
                        "date": d,
                        "clock_in": clock_in,
                        "clock_out": clock_out,
                    })
                except Exception:
                    continue
    except Exception:
        pass
    return records


def load_attendance(filepath: str) -> List[dict]:
    """
    出勤記録を読み込む（Excel/CSV/PDF対応）
    各要素: {employee_id, employee_name, date, clock_in, clock_out}
    """
    ext = Path(filepath).suffix.lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        return _load_excel(filepath)
    elif ext == ".csv":
        return _load_csv(filepath)
    elif ext == ".pdf":
        return _load_pdf(filepath)
    else:
        # 拡張子不明の場合はExcelで試みてからCSV
        try:
            return _load_excel(filepath)
        except Exception:
            try:
                return _load_csv(filepath)
            except Exception as e:
                raise ValueError(f"ファイル形式が不明または読み込み不可: {e}")
