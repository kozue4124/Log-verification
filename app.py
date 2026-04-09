"""動画ログ照合システム - Flask Webアプリケーション"""
import os
import tempfile
import traceback
import uuid
from datetime import datetime, date, time
from functools import wraps
from pathlib import Path
from flask import (Flask, render_template, request, send_file,
                   jsonify, redirect, url_for, session)
import io

# 生成済みExcelの一時キャッシュ {key: bytes}
_report_cache: dict[str, bytes] = {}

# 手書き勤怠変換CSVの一時キャッシュ {key: (emp_name, csv_bytes)}
_csv_cache: dict[str, tuple[str, bytes]] = {}

from processors.video_log import load_video_log
from processors.attendance import load_attendance
from processors.contract import load_contracts
from matcher import match_records, generate_summary
from reporter import generate_report

app = Flask(__name__)
# 本番では SECRET_KEY 環境変数を必ず設定すること
app.secret_key = os.environ.get("SECRET_KEY") or os.urandom(24)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv", ".pdf"}

# ---- 認証設定（環境変数から取得）----
APP_USERNAME = os.environ.get("APP_USERNAME", "admin")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")  # 未設定時はローカル開発用として警告のみ


def login_required(f):
    """ログイン必須デコレータ"""
    @wraps(f)
    def decorated(*args, **kwargs):
        # パスワード未設定 = ローカル開発環境とみなしてスキップ
        if not APP_PASSWORD:
            return f(*args, **kwargs)
        if not session.get("logged_in"):
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return decorated


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        return redirect(url_for("index"))
    error = None
    if request.method == "POST":
        if (request.form.get("username") == APP_USERNAME
                and request.form.get("password") == APP_PASSWORD):
            session["logged_in"] = True
            return redirect(request.args.get("next") or url_for("index"))
        error = "ユーザー名またはパスワードが正しくありません"
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
@login_required
def process():
    warnings = []

    # ---- 動画ログ（必須）----
    video_file = request.files.get("video_log")
    if not video_file or video_file.filename == "":
        return jsonify({"success": False, "error": "動画ログファイルが選択されていません"}), 400
    if not _allowed_file(video_file.filename):
        return jsonify({"success": False, "error": "動画ログはExcelファイルのみ対応しています"}), 400

    video_tmp = _save_upload(video_file)

    # ---- 出勤記録（複数可）----
    att_files   = request.files.getlist("attendance")
    att_names   = request.form.getlist("attendance_name")  # 氏名オーバーライド
    attendance_tmps = []
    for i, f in enumerate(att_files):
        if f and f.filename:
            if not _allowed_file(f.filename):
                warnings.append(f"出勤記録「{f.filename}」は非対応形式です")
                continue
            name_override = att_names[i].strip() if i < len(att_names) else ""
            attendance_tmps.append((f.filename, _save_upload(f), name_override))

    # ---- 雇用契約書（任意）----
    contract_tmps = []
    for f in request.files.getlist("contracts"):
        if f and f.filename:
            if not _allowed_file(f.filename):
                warnings.append(f"雇用契約書「{f.filename}」は非対応形式です")
                continue
            contract_tmps.append((f.filename, _save_upload(f)))

    try:
        # ---- データ読み込み ----
        try:
            video_logs = load_video_log(video_tmp)
        except Exception as e:
            return jsonify({"success": False, "error": f"動画ログの読み込みエラー: {e}"}), 400
        finally:
            _cleanup(video_tmp)

        attendance_all = []
        for fname, tmp_path, name_override in attendance_tmps:
            try:
                recs = load_attendance(tmp_path)
                # 氏名オーバーライドが指定されていれば全レコードに適用
                if name_override:
                    for r in recs:
                        r["employee_name"] = name_override
                attendance_all.extend(recs)
            except Exception as e:
                warnings.append(f"出勤記録「{fname}」の読み込みに失敗: {e}")
            finally:
                _cleanup(tmp_path)

        contracts_all = []
        for fname, tmp_path in contract_tmps:
            try:
                contracts_all.extend(load_contracts(tmp_path))
            except Exception as e:
                warnings.append(f"雇用契約書「{fname}」の読み込みに失敗: {e}")
            finally:
                _cleanup(tmp_path)

        if not video_logs:
            return jsonify({"success": False, "error": "動画ログにデータが見つかりませんでした"}), 400

        # ---- 照合・レポート ----
        results = match_records(video_logs, attendance_all, contracts_all)
        summary = generate_summary(results)
        report_bytes = generate_report(results)

        # Excelをキャッシュに保存してキーを発行
        dl_key = str(uuid.uuid4())
        _report_cache[dl_key] = report_bytes

        return jsonify({
            "success": True,
            "summary": summary,
            "results": _serialize_results(results),
            "download_key": dl_key,
            "warnings": warnings,
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({"success": False, "error": f"処理中にエラーが発生しました: {e}"}), 500


@app.route("/check_attendance_name", methods=["POST"])
@login_required
def check_attendance_name():
    """出勤記録ファイルから氏名が自動取得できるか確認する"""
    f = request.files.get("file")
    if not f or f.filename == "":
        return jsonify({"error": "ファイルが選択されていません"}), 400

    is_pdf = Path(f.filename).suffix.lower() == ".pdf"
    tmp_path = _save_upload(f)
    try:
        recs = load_attendance(tmp_path)
        detected_name = None
        for r in recs:
            name = (r.get("employee_name") or "").strip()
            if name:
                detected_name = name
                break

        # PDFは名前の誤検出が多いため「要確認」扱い
        # CSV/Excelで名前が取得できた場合のみ「確実」とする
        if is_pdf:
            return jsonify({
                "status": "verify",   # 黄色：要確認
                "name": detected_name or "",
            })
        elif detected_name:
            return jsonify({
                "status": "ok",       # 緑：自動取得成功
                "name": detected_name,
            })
        else:
            return jsonify({
                "status": "missing",  # オレンジ：取得できず・要入力
                "name": "",
            })
    except Exception as e:
        return jsonify({"status": "missing", "name": "", "error": str(e)}), 400
    finally:
        _cleanup(tmp_path)


@app.route("/preview", methods=["POST"])
@login_required
def preview():
    f = request.files.get("file")
    if not f or f.filename == "":
        return jsonify({"error": "ファイルが選択されていません"}), 400

    tmp_path = _save_upload(f)
    try:
        import pandas as pd
        ext = Path(f.filename).suffix.lower()
        if ext in (".xlsx", ".xls", ".xlsm"):
            df = pd.read_excel(tmp_path, dtype=str, nrows=5)
        elif ext == ".csv":
            import chardet
            with open(tmp_path, "rb") as raw_f:
                raw = raw_f.read()
            enc = chardet.detect(raw).get("encoding", "utf-8-sig")
            df = pd.read_csv(tmp_path, dtype=str, encoding=enc, nrows=5)
        else:
            return jsonify({"error": "PDFのプレビューは非対応です"}), 400
        return jsonify({
            "columns": list(df.columns),
            "rows": df.fillna("").values.tolist()[:3],
            "filename": f.filename,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    finally:
        _cleanup(tmp_path)


@app.route("/handwritten")
@login_required
def handwritten():
    """手書き勤怠簿 PDF → CSV 変換ページ"""
    api_configured = bool(os.environ.get("ANTHROPIC_API_KEY"))
    return render_template("handwritten.html", api_configured=api_configured)


@app.route("/handwritten/convert", methods=["POST"])
@login_required
def handwritten_convert():
    """手書き勤怠簿 PDF を Claude Vision で読み取り CSV を返す"""
    import csv as _csv
    from handwritten_attendance_to_csv import (
        pdf_to_images, extract_with_claude, pages_to_records, CSV_FIELDNAMES
    )

    pdf_file = request.files.get("pdf")
    if not pdf_file or pdf_file.filename == "":
        return jsonify({"success": False, "error": "PDFファイルが選択されていません"}), 400
    if Path(pdf_file.filename).suffix.lower() != ".pdf":
        return jsonify({"success": False, "error": "PDFファイルのみ対応しています"}), 400

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return jsonify({
            "success": False,
            "error": "ANTHROPIC_API_KEY が設定されていません。管理者にお問い合わせください。"
        }), 500

    tmp_path = _save_upload(pdf_file)
    try:
        images = pdf_to_images(tmp_path)
        if not images:
            return jsonify({"success": False, "error": "PDFから画像を取得できませんでした"}), 400

        pages_data = []
        for i, img_bytes in enumerate(images, 1):
            data = extract_with_claude(img_bytes, api_key)
            pages_data.append(data)

        emp_name, records = pages_to_records(pages_data)
        if not records:
            return jsonify({"success": False, "error": "勤務データが見つかりませんでした"}), 400

        # CSVをメモリ上に生成
        output = io.StringIO()
        writer = _csv.DictWriter(output, fieldnames=CSV_FIELDNAMES)
        writer.writeheader()
        for rec in records:
            writer.writerow({
                "氏名":     emp_name,
                "日付":     rec.get("date", ""),
                "出社時間": rec.get("clock_in", ""),
                "退社時間": rec.get("clock_out", ""),
                "就業時間": rec.get("work_hours", ""),
                "普通残業": rec.get("overtime", ""),
                "休出残業": rec.get("holiday_work", ""),
                "備考":     rec.get("note", ""),
            })

        dl_key = str(uuid.uuid4())
        _csv_cache[dl_key] = (emp_name, output.getvalue().encode("utf-8-sig"))

        return jsonify({
            "success": True,
            "employee_name": emp_name,
            "record_count": len(records),
            "records": records,
            "download_key": dl_key,
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({"success": False, "error": f"処理中にエラーが発生しました: {e}"}), 500
    finally:
        _cleanup(tmp_path)


@app.route("/handwritten/download/<key>")
@login_required
def handwritten_download(key):
    """変換済みCSVのダウンロード"""
    item = _csv_cache.pop(key, None)
    if not item:
        return "CSVが見つかりません（期限切れまたは無効なキーです）", 404
    emp_name, csv_bytes = item
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{emp_name}_勤怠データ_{timestamp}.csv"
    return send_file(
        io.BytesIO(csv_bytes),
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/download/<key>")
@login_required
def download(key):
    report_bytes = _report_cache.pop(key, None)
    if not report_bytes:
        return "レポートが見つかりません（期限切れまたは無効なキーです）", 404
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"動画ログ照合結果_{timestamp}.xlsx"
    return send_file(
        io.BytesIO(report_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


def _serialize_results(results: list) -> list:
    """date/time オブジェクトを文字列に変換してJSONシリアライズ可能にする"""
    out = []
    for r in results:
        item = {}
        for k, v in r.items():
            if k in ("matched_attendance", "matched_contract"):
                continue
            elif isinstance(v, datetime):
                item[k] = v.isoformat()
            elif isinstance(v, date):
                item[k] = v.strftime("%Y-%m-%d")
            elif isinstance(v, time):
                item[k] = v.strftime("%H:%M")
            else:
                item[k] = v
        out.append(item)
    return out


def _allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def _save_upload(file) -> str:
    suffix = Path(file.filename).suffix.lower()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    file.save(tmp.name)
    return tmp.name


def _cleanup(path: str):
    try:
        os.unlink(path)
    except Exception:
        pass


if __name__ == "__main__":
    if not APP_PASSWORD:
        print("⚠  APP_PASSWORD が未設定です。ローカル開発モードで起動します（認証なし）")
    port = int(os.environ.get("PORT", 5000))
    print(f"  http://localhost:{port}")
    app.run(debug=(os.environ.get("FLASK_ENV") == "development"), port=port)
