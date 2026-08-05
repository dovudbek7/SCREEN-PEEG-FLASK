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


_approver_cache: list[str] | None = None


def _approver_choices() -> list[str]:
    """承認者セレクタに出す 出張命令者 の一覧を 20期承認者名簿 から読む.

    名簿は起動後に変わらないため一度だけ読んで使い回す。
    読めなかった場合は空リストを返し, 画面では「全員」のみ選べるようにする。
    """
    global _approver_cache
    if _approver_cache is not None:
        return _approver_cache
    try:
        if SRC_DIR not in sys.path:
            sys.path.insert(0, SRC_DIR)
        from file_discovery import latest_matching
        from loaders.approver_loader import load_approver_rules
        path = latest_matching(MASTER_DIR, "*評価者・承認者一覧*.xlsx")
        rules = load_approver_rules(path, "20期")
        names = sorted({
            r.trip_approver_raw for r in rules
            if r.trip_approver_raw and "確認" not in r.trip_approver_raw
        })
        _approver_cache = names
    except Exception:  # noqa: BLE001  (名簿が読めなくても画面は出す)
        _approver_cache = []
    return _approver_cache


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
    空になったり Web サーバ本体を指すことがあるため, 順にフォールバックする:

      1. 環境変数 CHECKSHEET_PYTHON (明示指定)
      2. sys.executable
      3. sys.prefix (有効な仮想環境の場所。mkvirtualenv で ~/.virtualenvs/ に
         作った場合でも、プロジェクト直下の .venv でも、ここで拾える)
      4. プロジェクト直下の .venv / venv
      5. PATH 上の python3

    5番目まで来た場合はライブラリが入っていない可能性が高いが、
    何も返せないよりは実行を試みてエラー内容を見せたほうがよい。
    """
    env_exe = os.environ.get("CHECKSHEET_PYTHON")
    if env_exe and os.path.isfile(env_exe):
        return env_exe

    def _looks_like_python(path: str) -> bool:
        return bool(path) and os.path.isfile(path) \
            and "python" in os.path.basename(path).lower()

    if _looks_like_python(sys.executable):
        return sys.executable

    candidates = [
        os.path.join(sys.prefix, "bin", "python"),
        os.path.join(sys.prefix, "Scripts", "python.exe"),
        os.path.join(BASE_DIR, ".venv", "bin", "python"),
        os.path.join(BASE_DIR, ".venv", "Scripts", "python.exe"),
        os.path.join(BASE_DIR, "venv", "bin", "python"),
    ]
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate

    found = shutil.which("python3") or shutil.which("python")
    if found:
        return found
    raise RuntimeError("処理を実行する Python が見つかりませんでした。")


def _run_checksheet(work: str, approver: str = "") -> subprocess.CompletedProcess:
    env = dict(os.environ, CHECKSHEET_ROOT=work, PYTHONIOENCODING="utf-8")
    # --approver に空文字を渡すと絞り込みなし (全件対象) になる
    cmd = [_python_executable(), os.path.join(SRC_DIR, "main.py"),
           "--no-pause", "--approver", approver]
    return subprocess.run(
        cmd, capture_output=True, text=True, env=env,
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


def _render(**kwargs):
    """承認者セレクタの選択肢と現在の選択値を必ず渡してテンプレートを描画する."""
    kwargs.setdefault("approvers", _approver_choices())
    kwargs.setdefault("selected_approver", request.form.get("approver", ""))
    return render_template("index.html", **kwargs)


@app.get("/")
def index():
    _cleanup_old_jobs()
    return _render()


@app.post("/run")
def run():
    expense = request.files.get("expense")
    attendance = request.files.get("attendance")
    approver = (request.form.get("approver") or "").strip()

    # 未知の氏名が送られてきた場合は絞り込みなしに倒す
    if approver and approver not in _approver_choices():
        approver = ""

    if not expense or not expense.filename:
        return _render(error="出張精算CSVを選択してください。"), 400
    if not attendance or not attendance.filename:
        return _render(error="出勤簿xlsxを選択してください。"), 400
    if not expense.filename.lower().endswith(".csv"):
        return _render(error="出張精算は CSV ファイルを選択してください。"), 400
    if not attendance.filename.lower().endswith((".xlsx", ".xls")):
        return _render(error="出勤簿は xlsx ファイルを選択してください。"), 400

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
        proc = _run_checksheet(work, approver)
    except subprocess.TimeoutExpired:
        shutil.rmtree(work, ignore_errors=True)
        return _render(
            error=f"処理が {RUN_TIMEOUT_SECONDS} 秒以内に終わりませんでした。"
                  "ファイルの内容をご確認ください。"), 500

    log = _clean_log((proc.stdout or "") + (proc.stderr or ""), work)
    if proc.returncode != 0:
        return _render(error="処理中にエラーが発生しました。", log=log), 500

    stamp, files = _result_files(work)
    if not files:
        return _render(error="結果ファイルが生成されませんでした。", log=log), 500

    return _render(log=log, job_id=job_id, stamp=stamp, files=files,
                   summary=_parse_summary(log), used_approver=approver)


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

# マニュアル本体 (static/manual/manual.html) の先頭に差し込む操作バー。
# マニュアルのHTMLは pdef/manual/japan.html をそのままコピーしたものなので,
# ファイルには手を入れず, 表示するときにバーだけを足す。
#
# PDF は事前生成したファイルを配信するのではなく, ブラウザの印刷機能
# (window.print) を呼び出して「PDFに保存」してもらう方式にしている
# (2026-08-05)。表示中のHTMLからそのまま出力されるため, マニュアルを
# 更新するたびにPDFを作り直す必要がない。
_MANUAL_BAR = """
<style>
  #manual-bar {
    position: sticky; top: 0; z-index: 999;
    display: flex; flex-wrap: wrap; align-items: center; gap: 10px;
    padding: 12px 20px; margin: -40px -48px 28px;
    background: #14245a; color: #fff;
    font-family: 'Hiragino Kaku Gothic ProN','Yu Gothic',Meiryo,sans-serif;
  }
  #manual-bar .mb-title { font-size: 13px; font-weight: 700; flex: 1; min-width: 180px; }
  #manual-bar .mb-btn {
    padding: 7px 16px; border-radius: 6px; font-size: 12px; font-weight: 700;
    text-decoration: none; cursor: pointer; font-family: inherit;
    border: 1px solid rgba(255,255,255,.5); background: transparent; color: #fff;
  }
  #manual-bar .mb-btn.primary { background: #fff; color: #14245a; border-color: #fff; }
  #manual-bar .mb-back { color: rgba(255,255,255,.8); text-decoration: none; font-size: 12px; }
  /* 印刷 (PDF保存) のときは操作バーを出さない */
  @media print { #manual-bar { display: none !important; } }
</style>
<div id="manual-bar">
  <span class="mb-title">操作マニュアル</span>
  <button type="button" class="mb-btn primary" onclick="window.print()"
          title="印刷ダイアログが開きます。送信先で「PDFに保存」を選択してください。">PDFをダウンロード</button>
  <a class="mb-btn" href="/manual/download/html">HTMLをダウンロード</a>
  <a class="mb-back" href="/">← チェックシート作成へ戻る</a>
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
    """マニュアルのダウンロード。PDF はブラウザの印刷機能で出力するため
    ここでは HTML のみを配信する。"""
    if kind != "html":
        abort(404)
    if not os.path.isfile(os.path.join(MANUAL_DIR, "manual.html")):
        abort(404)
    return send_from_directory(
        MANUAL_DIR, "manual.html", as_attachment=True,
        download_name="出張精算_承認チェックシート_操作マニュアル.html")


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
