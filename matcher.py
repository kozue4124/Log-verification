from typing import Optional, List, Tuple
"""動画ログと出勤記録の照合ロジック"""
from datetime import time, date, datetime, timedelta
from difflib import SequenceMatcher
import re


# 照合結果のステータス定義
STATUS_OK = "OK"                            # 勤務時間内
STATUS_OUTSIDE_HOURS = "勤務時間外"         # 勤務時間外に視聴
STATUS_NO_ATTENDANCE = "出勤記録なし"       # 該当日の出勤記録がない
STATUS_PERSON_MISMATCH = "人物不一致"       # 人物が一致しない
STATUS_CONTRACT_ONLY = "契約時間参照"       # タイムカードなし・契約時間で判定
STATUS_CONTRACT_OUTSIDE = "契約時間外"      # 契約時間外
STATUS_NO_TIME_INFO = "時刻情報なし"        # 動画ログに時刻情報がない


def _normalize_name(name: str) -> str:
    """氏名を正規化（スペース・全角半角・敬称削除）"""
    if not name:
        return ""
    name = str(name).strip()
    # 全角→半角
    name = name.translate(str.maketrans(
        "　ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ０１２３４５６７８９",
        " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    ))
    # スペース統一・削除
    name = re.sub(r"\s+", "", name)
    # 敬称削除
    for suffix in ["様", "さん", "殿", "氏"]:
        name = name.rstrip(suffix)
    return name.lower()


def _normalize_id(emp_id: str) -> str:
    """社員番号を正規化（前後ゼロ・スペース等）"""
    if not emp_id:
        return ""
    s = str(emp_id).strip()
    # 数値の場合は先頭ゼロ除去
    if s.isdigit():
        return str(int(s))
    return s.lower()


def _name_similarity(a: str, b: str) -> float:
    """氏名の類似度（0〜1）"""
    na, nb = _normalize_name(a), _normalize_name(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()


def _id_match(a: Optional[str], b: Optional[str]) -> bool:
    """社員番号が一致するか"""
    if not a or not b:
        return False
    return _normalize_id(a) == _normalize_id(b)


def _person_match(
    video_emp_id: Optional[str],
    video_emp_name: Optional[str],
    att_emp_id: Optional[str],
    att_emp_name: Optional[str],
    name_threshold: float = 0.85
) -> bool:
    """
    動画ログと出勤記録の人物照合。
    ・IDが一致し、かつ氏名が大きく乖離していなければ一致とみなす
    ・IDがない場合は氏名の類似度で判定
    """
    id_match_result = None
    if video_emp_id and att_emp_id:
        id_match_result = _id_match(video_emp_id, att_emp_id)
        if id_match_result:
            # IDが一致する場合でも、氏名が著しく異なれば不一致と判定
            if video_emp_name and att_emp_name:
                sim = _name_similarity(video_emp_name, att_emp_name)
                if sim < 0.5:  # 50%未満なら別人と判断
                    return False
            return True
        else:
            return False

    # IDがない場合は名前で判定
    if video_emp_name and att_emp_name:
        return _name_similarity(video_emp_name, att_emp_name) >= name_threshold

    return False


def _time_within(view_start: time, view_end: Optional[time],
                 work_start: time, work_end: time) -> bool:
    """視聴時刻が勤務時間内に含まれるか"""
    if view_end:
        # 視聴開始と終了の両方が勤務時間内
        return work_start <= view_start and view_end <= work_end
    else:
        # 開始時刻のみチェック
        return work_start <= view_start <= work_end


def build_lookup(attendance: List[dict]) -> dict:
    """
    出勤記録を {(employee_key, date): [records]} の辞書に変換
    employee_key = 社員番号 or 氏名（正規化済み）
    """
    lookup: dict[tuple, list] = {}
    for rec in attendance:
        emp_id = _normalize_id(rec.get("employee_id") or "")
        emp_name = _normalize_name(rec.get("employee_name") or "")
        d = rec.get("date")
        if not d:
            continue
        # IDと名前の両方でインデックス
        for key in filter(None, [emp_id, emp_name]):
            k = (key, d)
            lookup.setdefault(k, []).append(rec)
    return lookup


def build_contract_lookup(contracts: List[dict]) -> dict:
    """雇用契約書を {employee_key: contract} の辞書に変換"""
    lookup: dict[str, dict] = {}
    for c in contracts:
        emp_id = _normalize_id(c.get("employee_id") or "")
        emp_name = _normalize_name(c.get("employee_name") or "")
        for key in filter(None, [emp_id, emp_name]):
            lookup[key] = c
    return lookup


def _find_attendance(
    video: dict,
    att_lookup: dict,
    all_attendance: List[dict]
) -> Tuple[List[dict], bool]:
    """
    動画ログに対応する出勤記録を検索。
    Returns: (matched_records, person_mismatch_flag)

    person_mismatch_flag=True になるのは:
      ・社員番号で出勤記録が見つかったが、氏名が大きく異なる場合（なりすまし・誤登録）
      ・氏名で出勤記録が見つかったが、社員番号が大きく異なる場合
    """
    emp_id = _normalize_id(video.get("employee_id") or "")
    emp_name = _normalize_name(video.get("employee_name") or "")
    d = video.get("date")

    if not d:
        return [], False

    # 社員番号で検索
    id_matched = []
    if emp_id:
        id_matched = att_lookup.get((emp_id, d), [])

    # 氏名で検索
    name_matched = []
    if emp_name:
        name_matched = att_lookup.get((emp_name, d), [])

    # 両方で見つかった場合は統合（重複排除）
    candidates = list(id_matched)
    for r in name_matched:
        if r not in candidates:
            candidates.append(r)

    if not candidates:
        return [], False

    # 人物不一致チェック:
    # 社員番号で見つかった記録の氏名、または氏名で見つかった記録の社員番号が
    # 動画ログ側と大きく乖離している場合にフラグを立てる
    person_mismatch = False
    for ar in candidates:
        ar_id = _normalize_id(ar.get("employee_id") or "")
        ar_name = _normalize_name(ar.get("employee_name") or "")

        id_ok = (not emp_id or not ar_id) or (emp_id == ar_id)
        name_ok = (not emp_name or not ar_name) or (_name_similarity(
            video.get("employee_name") or "", ar.get("employee_name") or ""
        ) >= 0.7)

        if not id_ok or not name_ok:
            person_mismatch = True
            break

    return candidates, person_mismatch


def _find_contract(video: dict, contract_lookup: dict) -> Optional[dict]:
    emp_id = _normalize_id(video.get("employee_id") or "")
    emp_name = _normalize_name(video.get("employee_name") or "")
    for key in filter(None, [emp_id, emp_name]):
        c = contract_lookup.get(key)
        if c:
            return c
    return None


def match_records(
    video_logs: List[dict],
    attendance: List[dict],
    contracts: List[dict],
    name_threshold: float = 0.85,
) -> List[dict]:
    """
    動画ログと出勤記録を照合し、各動画ログに判定結果を付与して返す。

    Returns: 各要素に以下を追加した動画ログリスト
        - status: STATUS_* 定数
        - alerts: List[str] アラートメッセージ
        - matched_attendance: 対応する出勤記録（あれば）
        - matched_contract: 対応する雇用契約（あれば）
    """
    att_lookup = build_lookup(attendance)
    contract_lookup = build_contract_lookup(contracts)

    results = []
    for video in video_logs:
        result = dict(video)
        result["alerts"] = []
        result["status"] = STATUS_OK
        result["matched_attendance"] = None
        result["matched_contract"] = None

        view_start: Optional[time] = video.get("start_time")
        view_end: Optional[time] = video.get("end_time")
        d: Optional[date] = video.get("date")

        # 時刻情報がない場合
        if view_start is None and view_end is None:
            result["status"] = STATUS_NO_TIME_INFO
            results.append(result)
            continue

        # 出勤記録を検索
        att_records, _ = _find_attendance(video, att_lookup, attendance)

        if att_records:
            result["matched_attendance"] = att_records
            v_name_raw = video.get("employee_name") or ""
            ar_name_raw = att_records[0].get("employee_name") or ""

            # ---- ステップ１: 人物確認（名前照合を最優先）----
            same_person = (
                not v_name_raw
                or not ar_name_raw
                or _name_similarity(_normalize_name(v_name_raw), _normalize_name(ar_name_raw)) >= 0.6
            )

            if not same_person:
                # ステップ２: 別人物 → アラートのみ・時間照合はスキップ
                result["status"] = STATUS_PERSON_MISMATCH
                result["alerts"].append(
                    f"🚨 人物不一致: 動画ログ「{v_name_raw}」/ 出勤記録「{ar_name_raw}」"
                )
            else:
                # ステップ３: 同一人物 → 時間照合
                in_hours = False
                for ar in att_records:
                    clock_in = ar.get("clock_in")
                    clock_out = ar.get("clock_out")
                    if clock_in and clock_out and view_start:
                        if _time_within(view_start, view_end, clock_in, clock_out):
                            in_hours = True
                            break

                if not in_hours and view_start:
                    best_ar = att_records[0]
                    ci = best_ar.get("clock_in")
                    co = best_ar.get("clock_out")
                    result["status"] = STATUS_OUTSIDE_HOURS
                    result["alerts"].append(
                        f"⚠ 勤務時間外視聴: 視聴 {_fmt_time(view_start)}〜{_fmt_time(view_end)} / "
                        f"勤務 {_fmt_time(ci)}〜{_fmt_time(co)}"
                    )

        else:
            # 出勤記録なし → 雇用契約書のデフォルト時間を参照
            contract = _find_contract(video, contract_lookup)
            if contract:
                result["matched_contract"] = contract
                work_start = contract.get("work_start")
                work_end = contract.get("work_end")
                work_days = contract.get("work_days")

                # 曜日チェック
                day_ok = True
                if work_days is not None and d:
                    day_ok = d.weekday() in work_days

                if not day_ok:
                    result["status"] = STATUS_CONTRACT_OUTSIDE
                    result["alerts"].append(
                        f"⚠ 契約勤務日外: {d} ({_weekday_ja(d.weekday())}) は所定勤務日ではありません"
                        f"（所定: {_fmt_work_days(work_days)}）"
                    )
                elif work_start and work_end and view_start:
                    if _time_within(view_start, view_end, work_start, work_end):
                        result["status"] = STATUS_CONTRACT_ONLY
                        result["alerts"].append(
                            "ℹ タイムカードなし（雇用契約の所定勤務時間内と判断）"
                        )
                    else:
                        result["status"] = STATUS_CONTRACT_OUTSIDE
                        result["alerts"].append(
                            f"⚠ 契約時間外視聴: 視聴 {_fmt_time(view_start)}〜{_fmt_time(view_end)} / "
                            f"所定 {_fmt_time(work_start)}〜{_fmt_time(work_end)}"
                        )
                else:
                    result["status"] = STATUS_NO_ATTENDANCE
                    result["alerts"].append("ℹ タイムカードなし（契約時間情報も不完全）")
            else:
                result["status"] = STATUS_NO_ATTENDANCE
                result["alerts"].append(
                    f"⚠ 出勤記録なし: {d} の出勤記録が見つかりません"
                )

        results.append(result)

    return results


def _fmt_time(t: Optional[time]) -> str:
    if t is None:
        return "不明"
    return t.strftime("%H:%M")


def _weekday_ja(weekday: int) -> str:
    return ["月", "火", "水", "木", "金", "土", "日"][weekday]


def _fmt_work_days(work_days: Optional[List[int]]) -> str:
    if work_days is None:
        return "全日"
    return "・".join(_weekday_ja(d) for d in work_days)


def generate_summary(results: List[dict]) -> dict:
    """照合結果のサマリーを生成"""
    total = len(results)
    summary = {
        "total": total,
        "ok": sum(1 for r in results if r["status"] == STATUS_OK),
        "outside_hours": sum(1 for r in results if r["status"] == STATUS_OUTSIDE_HOURS),
        "no_attendance": sum(1 for r in results if r["status"] == STATUS_NO_ATTENDANCE),
        "person_mismatch": sum(1 for r in results if r["status"] == STATUS_PERSON_MISMATCH),
        "contract_only": sum(1 for r in results if r["status"] == STATUS_CONTRACT_ONLY),
        "contract_outside": sum(1 for r in results if r["status"] == STATUS_CONTRACT_OUTSIDE),
        "no_time_info": sum(1 for r in results if r["status"] == STATUS_NO_TIME_INFO),
        "alert_count": sum(1 for r in results if r.get("alerts")),
    }
    return summary
