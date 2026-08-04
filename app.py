"""出張精算 承認チェックシート — Web版 (Flask).

ブラウザから 出張精算CSV と 出勤簿xlsx をアップロードすると, チェックシート
(Excel + HTML) を生成してダウンロードできる。

同時アクセスでもファイルが混ざらないよう, リクエストごとに jobs/<id>/ の
作業フォルダを作り, 環境変数 CHECKSHEET_ROOT でそこを指すようにして
src/main.py を実行する (config._project_root 参照)。
マスタ (顧客・社員リスト等) はリポジトリ同梱のものをシンボリックリンクで共有する。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
import uuid

from flask import (Flask, abort, redirect, render_template, request,
                   send_from_directory, url_for)
from werkzeug.utils import secure_filename

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(BASE_DIR, "src")
MASTER_DIR = os.path.join(BASE_DIR, "出張精算データ一式")
JOBS_DIR = os.path.join(BASE_DIR, "jobs")

# 生成物を保持する時間。これを過ぎた作業フォルダはリクエストのたびに削除する。
JOB_TTL_SECONDS = 6 * 60 * 60
# 処理が終わらない場合に打ち切る秒数 (通常は15秒程度で終わる)
RUN_TIMEOUT_SECONDS = 300

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64MB


def _cleanup_old_jobs() -> None:
    """TTL を過ぎた作業フォルダを削除する."""
    if not os.path.isdir(JOBS_DIR):
        return
    cutoff = time.time() - JOB_TTL_SECONDS
    for name in os.listdir(JOBS_DIR):
        path = os.path.join(JOBS_DIR, name)
        try:
            if os.path.isdir(path) and os.path.getmtime(path) < cutoff:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass


def _prepare_workspace(job_id: str) -> str:
    """jobs/<id>/ に expenses・attendance・マスタを用意して返す."""
    work = os.path.join(JOBS_DIR, job_id)
    os.makedirs(os.path.join(work, "expenses"), exist_ok=True)
    os.makedirs(os.path.join(work, "attendance"), exist_ok=True)
    os.makedirs(os.path.join(work, "out"), exist_ok=True)

    # マスタは読み取り専用なのでコピーせずリンクする (Windows で失敗したらコピー)
    link = os.path.join(work, os.path.basename(MASTER_DIR))
    if not os.path.exists(link):
        try:
            os.symlink(MASTER_DIR, link)
        except (OSError, NotImplementedError):
            shutil.copytree(MASTER_DIR, link)
    return work


def _save_upload(file_storage, dest_dir: str) -> str:
    """アップロードされたファイルを保存する. 日本語名は secure_filename で
    消えてしまうため, 拡張子だけ検証して元の名前を活かす."""
    original = os.path.basename(file_storage.filename or "")
    name = original.replace("/", "_").replace("\\", "_").lstrip(".")
    if not name:
        name = secure_filename(original) or "uploaded"
    path = os.path.join(dest_dir, name)
    file_storage.save(path)
    return name


def _python_executable() -> str:
    """処理を実行する Python のパスを決める.

    WSGI サーバ (PythonAnywhere の uWSGI 等) の下では sys.executable が
    空になったり Web サーバ本体を指すことがあるため, その場合は
    同梱の仮想環境 → PATH 上の python3 の順にフォールバックする。
    """
    exe = sys.executable
    if exe and os.path.isfile(exe) and "python" in os.path.basename(exe).lower():
        return exe
    for candidate in (
        os.path.join(BASE_DIR, ".venv", "bin", "python"),
        os.path.join(BASE_DIR, ".venv", "Scripts", "python.exe"),
        os.path.join(BASE_DIR, "venv", "bin", "python"),
    ):
        if os.path.isfile(candidate):
            return candidate
    found = shutil.which("python3") or shutil.which("python")
    if found:
        return found
    raise RuntimeError("処理を実行する Python が見つかりませんでした。")


def _run_checksheet(work: str) -> subprocess.CompletedProcess:
    env = dict(os.environ, CHECKSHEET_ROOT=work, PYTHONIOENCODING="utf-8")
    return subprocess.run(
        [_python_executable(), os.path.join(SRC_DIR, "main.py"), "--no-pause"],
        capture_output=True, text=True, env=env,
        timeout=RUN_TIMEOUT_SECONDS,
    )


def _result_files(work: str) -> tuple[str, list[str]]:
    """最新の out/<stamp>/ とその中のファイル名一覧を返す."""
    out_root = os.path.join(work, "out")
    stamps = sorted(
        d for d in os.listdir(out_root)
        if os.path.isdir(os.path.join(out_root, d))
    ) if os.path.isdir(out_root) else []
    if not stamps:
        return "", []
    stamp = stamps[-1]
    files = sorted(os.listdir(os.path.join(out_root, stamp)))
    return stamp, files


@app.get("/")
def index():
    _cleanup_old_jobs()
    return render_template("index.html")


@app.post("/run")
def run():
    expense = request.files.get("expense")
    attendance = request.files.get("attendance")

    if not expense or not expense.filename:
        return render_template("index.html",
                               error="出張精算CSVを選択してください。"), 400
    if not attendance or not attendance.filename:
        return render_template("index.html",
                               error="出勤簿xlsxを選択してください。"), 400
    if not expense.filename.lower().endswith(".csv"):
        return render_template("index.html",
                               error="出張精算は CSV ファイルを選択してください。"), 400
    if not attendance.filename.lower().endswith((".xlsx", ".xls")):
        return render_template("index.html",
                               error="出勤簿は xlsx ファイルを選択してください。"), 400

    _cleanup_old_jobs()
    job_id = uuid.uuid4().hex[:12]
    work = _prepare_workspace(job_id)

    # ローダは 出張精算_*.csv / 出勤簿_*.xlsx というパターンで探すため,
    # アップロード名がそれに合わない場合は接頭辞を補う。
    exp_name = _save_upload(expense, os.path.join(work, "expenses"))
    if not exp_name.startswith("出張精算"):
        fixed = "出張精算_" + exp_name
        os.rename(os.path.join(work, "expenses", exp_name),
                  os.path.join(work, "expenses", fixed))
    att_name = _save_upload(attendance, os.path.join(work, "attendance"))
    if not att_name.startswith("出勤簿"):
        fixed = "出勤簿_" + att_name
        os.rename(os.path.join(work, "attendance", att_name),
                  os.path.join(work, "attendance", fixed))

    try:
        proc = _run_checksheet(work)
    except subprocess.TimeoutExpired:
        shutil.rmtree(work, ignore_errors=True)
        return render_template(
            "index.html",
            error=f"処理が {RUN_TIMEOUT_SECONDS} 秒以内に終わりませんでした。"
                  "ファイルの内容をご確認ください。"), 500

    log = _clean_log((proc.stdout or "") + (proc.stderr or ""), work)
    if proc.returncode != 0:
        return render_template("index.html",
                               error="処理中にエラーが発生しました。",
                               log=log), 500

    stamp, files = _result_files(work)
    if not files:
        return render_template("index.html",
                               error="結果ファイルが生成されませんでした。",
                               log=log), 500

    return render_template("index.html", log=log, job_id=job_id,
                           stamp=stamp, files=files,
                           summary=_parse_summary(log))


def _clean_log(log: str, work: str) -> str:
    """実行ログからサーバ内部のパスを取り除く.

    main.py は出力先の絶対パスを表示するが, Web版では作業フォルダの場所は
    利用者に関係が無いうえ, サーバの内部構造が見えてしまうため伏せる。
    """
    out = []
    for line in log.splitlines():
        line = line.replace(work + os.sep, "").replace(work, "")
        if line.startswith("出力フォルダ:"):
            continue
        out.append(line)
    return "\n".join(out).strip()


def _parse_summary(log: str) -> list[tuple[str, str]]:
    """実行ログ末尾の「=== 総合判定 内訳 ===」を (ラベル, 件数) に整形."""
    rows: list[tuple[str, str]] = []
    seen_header = False
    for line in log.splitlines():
        if "総合判定 内訳" in line:
            seen_header = True
            continue
        if seen_header:
            stripped = line.strip()
            if not stripped:
                break
            if ":" in stripped:
                label, _, count = stripped.partition(":")
                rows.append((label.strip(), count.strip()))
    return rows


@app.get("/download/<job_id>/<stamp>/<path:filename>")
def download(job_id: str, stamp: str, filename: str):
    # パストラバーサル防止: 各要素にセパレータを含めない
    for part in (job_id, stamp, filename):
        if "/" in part or "\\" in part or part.startswith("."):
            abort(400)
    directory = os.path.join(JOBS_DIR, job_id, "out", stamp)
    if not os.path.isdir(directory):
        abort(404)
    return send_from_directory(directory, filename, as_attachment=True)


MANUAL_DIR = os.path.join(BASE_DIR, "static", "manual")

# マニュアル本体 (static/manual/manual.html) の先頭に差し込むダウンロードバー。
# マニュアルのHTMLは pdef/manual/japan.html をそのままコピーしたものなので,
# ファイルには手を入れず, 表示するときにバーだけを足す。
_MANUAL_BAR = """
<div style="position:sticky;top:0;z-index:999;display:flex;flex-wrap:wrap;
            align-items:center;gap:10px;padding:12px 20px;margin:-40px -48px 28px;
            background:#14245a;color:#fff;
            font-family:'Hiragino Kaku Gothic ProN','Yu Gothic',Meiryo,sans-serif;">
  <span style="font-size:13px;font-weight:700;flex:1;min-width:180px;">操作マニュアル</span>
  <a href="/manual/download/pdf"
     style="padding:7px 16px;border-radius:6px;background:#fff;color:#14245a;
            text-decoration:none;font-size:12px;font-weight:700;">PDF をダウンロード</a>
  <a href="/manual/download/html"
     style="padding:7px 16px;border-radius:6px;border:1px solid rgba(255,255,255,.5);
            color:#fff;text-decoration:none;font-size:12px;font-weight:700;">HTML をダウンロード</a>
  <a href="/" style="color:rgba(255,255,255,.8);text-decoration:none;font-size:12px;">
     ← チェックシート作成へ戻る</a>
</div>
"""


@app.get("/manual")
def manual():
    """マニュアルを表示する (先頭にダウンロードバーを差し込む)."""
    path = os.path.join(MANUAL_DIR, "manual.html")
    if not os.path.isfile(path):
        abort(404)
    with open(path, encoding="utf-8") as f:
        html_text = f.read()
    marker = "<body>"
    if marker in html_text:
        html_text = html_text.replace(marker, marker + _MANUAL_BAR, 1)
    else:
        html_text = _MANUAL_BAR + html_text
    return html_text


@app.get("/manual/download/<kind>")
def manual_download(kind: str):
    names = {"pdf": "出張精算_承認チェックシート_操作マニュアル.pdf",
             "html": "出張精算_承認チェックシート_操作マニュアル.html"}
    files = {"pdf": "manual.pdf", "html": "manual.html"}
    if kind not in files:
        abort(404)
    if not os.path.isfile(os.path.join(MANUAL_DIR, files[kind])):
        abort(404)
    return send_from_directory(MANUAL_DIR, files[kind],
                               as_attachment=True,
                               download_name=names[kind])


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
