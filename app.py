"""動画ログ照合システム - Flask Webアプリケーション"""
import os
import tempfile
import traceback
from datetime import datetime
from functools import wraps
from pathlib import Path
from flask import (Flask, render_template, request, send_file,
                   jsonify, redirect, url_for, session)
import io

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
    attendance_tmps = []
    for f in request.files.getlist("attendance"):
        if f and f.filename:
            if not _allowed_file(f.filename):
                warnings.append(f"出勤記録「{f.filename}」は非対応形式です")
                continue
            attendance_tmps.append((f.filename, _save_upload(f)))

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
        for fname, tmp_path in attendance_tmps:
            try:
                attendance_all.extend(load_attendance(tmp_path))
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

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"動画ログ照合結果_{timestamp}.xlsx"

        response = send_file(
            io.BytesIO(report_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )
        response.headers["X-Summary-Total"] = str(summary["total"])
        response.headers["X-Summary-Alerts"] = str(summary["alert_count"])
        response.headers["X-Warnings"] = "|".join(warnings) if warnings else ""
        return response

    except Exception as e:
        traceback.print_exc()
        return jsonify({"success": False, "error": f"処理中にエラーが発生しました: {e}"}), 500


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
