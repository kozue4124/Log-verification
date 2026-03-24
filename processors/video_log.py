from typing import Optional, List, Tuple
"""動画ログ読み込みプロセッサ（Excel形式）"""
import pandas as pd
from datetime import datetime, time, date, timedelta
from dateutil import parser as date_parser
import re


# 列名の候補（日本語・英語）
COLUMN_ALIASES = {
    "employee_id": [
        "社員番号", "社員ID", "従業員番号", "従業員ID", "ID", "id",
        "employee_id", "emp_id", "staff_id", "ログインID", "ユーザーID", "login_id",
    ],
    "employee_name": [
        "氏名", "名前", "社員名", "従業員名", "name", "employee_name",
        "full_name", "氏名（フリガナ）",
    ],
    # 日付単独列
    "date": ["日付", "視聴日", "date", "視聴日付", "日", "Date", "学習日"],
    # 開始時刻単独列
    "start_time": [
        "開始時刻", "視聴開始", "開始時間", "start_time", "start",
        "開始", "視聴開始時刻", "Start",
    ],
    # 終了時刻単独列
    "end_time": [
        "終了時刻", "視聴終了", "終了時間", "end_time", "end",
        "終了", "視聴終了時刻", "End",
    ],
    # 日時一体型列（学習開始日時 など）→ date + start_time を両方導出
    "start_datetime": [
        "学習開始日時", "視聴開始日時", "開始日時", "start_datetime",
        "datetime", "学習日時",
    ],
    # 所要時間（HH:MM:SS）→ end_time = start_time + duration
    "duration": [
        "所要時間", "視聴時間", "再生時間", "duration", "時間",
        "視聴時間（分）", "再生時間（秒）", "学習時間",
    ],
    "video_title": [
        "動画タイトル", "コンテンツ名", "動画名", "video_title", "title",
        "コース名", "研修名",
    ],
    "video_id": ["動画ID", "コンテンツID", "video_id", "content_id"],
    "progress": [
        "進捗率", "完了率", "progress", "視聴率", "達成率",
        "点数/達成率", "点数/達成率(数値のみ)", "合否",
    ],
}

_EMPTY = (None, "", "nan", "None", "NaN", "--", "-")


def _find_column(df: pd.DataFrame, field: str) -> Optional[str]:
    """列名の候補からDataFrameの列名を検索（完全一致優先→部分一致）"""
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
    """時刻文字列をtimeオブジェクトに変換"""
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass
    if val is None:
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, datetime):
        return val.time()
    s = str(val).strip()
    if not s or s in _EMPTY:
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


def _parse_datetime(val) -> Optional[datetime]:
    """日時文字列をdatetimeオブジェクトに変換"""
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass
    if val is None:
        return None
    if isinstance(val, datetime):
        return val
    s = str(val).strip()
    if not s or s in _EMPTY:
        return None
    try:
        return date_parser.parse(s)
    except Exception:
        return None


def _parse_date(val) -> Optional[date]:
    """日付文字列をdateオブジェクトに変換"""
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if hasattr(val, "date"):
        return val.date()
    s = str(val).strip()
    if not s or s in _EMPTY:
        return None
    try:
        return date_parser.parse(s).date()
    except Exception:
        return None


def _parse_duration(val) -> Optional[timedelta]:
    """所要時間文字列（HH:MM:SS）をtimedeltaに変換"""
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass
    if val is None:
        return None
    s = str(val).strip()
    if not s or s in _EMPTY:
        return None
    # HH:MM:SS
    m = re.match(r"^(\d+):(\d{2}):(\d{2})$", s)
    if m:
        return timedelta(
            hours=int(m.group(1)),
            minutes=int(m.group(2)),
            seconds=int(m.group(3)),
        )
    # MM:SS
    m2 = re.match(r"^(\d+):(\d{2})$", s)
    if m2:
        return timedelta(minutes=int(m2.group(1)), seconds=int(m2.group(2)))
    # 数値（分）
    try:
        return timedelta(minutes=float(s))
    except ValueError:
        return None


def load_video_log(filepath: str) -> List[dict]:
    """
    動画ログExcelを読み込みリスト形式で返す。

    対応フォーマット:
      A) 日付・開始時刻・終了時刻が別列
      B) 学習開始日時（日時一体） + 所要時間 → 終了時刻を計算

    各要素: {
        employee_id, employee_name, date, start_time, end_time,
        video_title, video_id, progress, raw_row_index
    }
    """
    try:
        df = pd.read_excel(filepath, dtype=str)
    except Exception as e:
        raise ValueError(f"動画ログの読み込みに失敗しました: {e}")

    df = df.dropna(how="all").reset_index(drop=True)

    col_map = {}
    for field in COLUMN_ALIASES:
        col = _find_column(df, field)
        if col:
            col_map[field] = col

    if not col_map:
        raise ValueError("動画ログの列が認識できません。列名を確認してください。")

    found_person = any(f in col_map for f in ("employee_id", "employee_name"))
    if not found_person:
        raise ValueError("社員番号または氏名の列が見つかりません。")

    records = []
    for idx, row in df.iterrows():
        # --- 氏名・ID ---
        emp_id_raw = str(row[col_map["employee_id"]]).strip() if "employee_id" in col_map else None
        emp_name_raw = str(row[col_map["employee_name"]]).strip() if "employee_name" in col_map else None
        emp_id = None if emp_id_raw in _EMPTY else emp_id_raw
        emp_name = None if emp_name_raw in _EMPTY else emp_name_raw

        if emp_id is None and emp_name is None:
            continue

        # --- 日付・開始時刻の解決 ---
        # パターンB: 学習開始日時（日時一体型）
        rec_date: Optional[date] = None
        start_time: Optional[time] = None
        end_time: Optional[time] = None

        if "start_datetime" in col_map:
            dt = _parse_datetime(row[col_map["start_datetime"]])
            if dt:
                rec_date = dt.date()
                start_time = dt.time()
        # パターンA: 日付・時刻が別列
        if rec_date is None and "date" in col_map:
            rec_date = _parse_date(row[col_map["date"]])
        if start_time is None and "start_time" in col_map:
            start_time = _parse_time(row[col_map["start_time"]])

        # 終了時刻: 終了時刻列があれば使用、なければ開始+所要時間で計算
        if "end_time" in col_map:
            end_time = _parse_time(row[col_map["end_time"]])

        if end_time is None and start_time and "duration" in col_map:
            dur = _parse_duration(row[col_map["duration"]])
            if dur:
                from datetime import datetime as dt_cls
                end_dt = dt_cls.combine(rec_date or date.today(), start_time) + dur
                end_time = end_dt.time()

        records.append({
            "employee_id": emp_id,
            "employee_name": emp_name,
            "date": rec_date,
            "start_time": start_time,
            "end_time": end_time,
            "duration": row[col_map["duration"]] if "duration" in col_map else None,
            "video_title": str(row[col_map["video_title"]]).strip() if "video_title" in col_map else "",
            "video_id": str(row[col_map["video_id"]]).strip() if "video_id" in col_map else "",
            "progress": str(row[col_map["progress"]]).strip() if "progress" in col_map else None,
            "raw_row_index": idx + 2,
        })

    return records
