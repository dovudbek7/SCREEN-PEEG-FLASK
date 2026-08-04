"""PythonAnywhere 用 WSGI 設定ファイル (テンプレート).

PythonAnywhere の Web タブで作られる
    /var/www/<ユーザー名>_pythonanywhere_com_wsgi.py
の中身を、このファイルの内容で置き換えてください。
PROJECT_DIR のパスだけ、ご自身のユーザー名に合わせて書き換えます。

ローカル実行 (python app.py) や Render (Procfile) には影響しません。
"""
import os
import sys

# ── ここだけ書き換えてください ──────────────────────────────
PROJECT_DIR = "/home/YOURUSERNAME/SCREEN-PEEG-FLASK"
# ──────────────────────────────────────────────────────

if PROJECT_DIR not in sys.path:
    sys.path.insert(0, PROJECT_DIR)

# 作業フォルダを固定する。PythonAnywhere の web worker は
# カレントディレクトリが不定のため、明示的に移動しておく。
os.chdir(PROJECT_DIR)

from app import app as application  # noqa: E402  (sys.path 設定後に import する)
