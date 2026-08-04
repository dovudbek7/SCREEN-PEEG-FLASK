# 出張精算 承認チェックシート 自動生成 — Web版

ブラウザから 出張精算CSV と 出勤簿xlsx をアップロードすると、
チェックシート (Excel + HTML) を生成してダウンロードできる Flask アプリ。

CLI 版・実行ファイル版と判定ロジックは同一（`src/` は同じコードのコピー）。

---

## ローカルで動かす

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py
```

→ http://127.0.0.1:5000

ポートを変えたい場合は `PORT=8077 .venv/bin/python app.py`。

---

## PythonAnywhere へのデプロイ

### 1. コードを取得

Bash コンソールを開いて:

```bash
git clone https://github.com/dovudbek7/SCREEN-PEEG-FLASK.git
```

### 2. 仮想環境を作る

```bash
cd SCREEN-PEEG-FLASK
python3.10 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

### 3. Web アプリを作成

Web タブ → **Add a new web app** → **Manual configuration** → **Python 3.10**

### 4. 設定を入力

| 項目 | 値 |
|---|---|
| Source code | `/home/<ユーザー名>/SCREEN-PEEG-FLASK` |
| Working directory | `/home/<ユーザー名>/SCREEN-PEEG-FLASK` |
| Virtualenv | `/home/<ユーザー名>/SCREEN-PEEG-FLASK/.venv` |

### 5. WSGI ファイルを差し替え

Web タブの **WSGI configuration file** のリンクを開き、中身を
本リポジトリの `wsgi_pythonanywhere.py` の内容で置き換える。
`PROJECT_DIR` のユーザー名だけ自分のものに書き換えること。

### 6. Reload

Web タブの緑の **Reload** ボタンを押す。
`https://<ユーザー名>.pythonanywhere.com` で開けるようになる。

### コードを更新したとき

```bash
cd ~/SCREEN-PEEG-FLASK && git pull
```

そのあと Web タブで **Reload** を押す。

### 注意

- 無料プランの Web アプリは **3か月ごとに更新ボタンを押す**必要がある
  （Web タブに期限と更新ボタンが表示される）
- 無料プランは 1日あたりの CPU 秒に上限がある。
  1回の処理は CPU 約1.5秒なので、通常の利用では十分足りる

---

## Render へのデプロイ

`Procfile` を同梱しているため、GitHub リポジトリを連携するだけで動く。

- Build command: `pip install -r requirements.txt`
- Start command: Procfile が自動で使われる

無料プランは 15分アクセスが無いとスリープし、次回起動に約50秒かかる。

---

## 構成

```
app.py                    Flask 本体
templates/index.html      アップロード画面
static/manual/            マニュアル (HTML + PDF)
src/                      判定ロジック (CLI版と同一)
出張精算データ一式/         マスタ (社員・顧客・承認者名簿・手当)
jobs/                     実行時の作業フォルダ (自動生成・6時間で削除)
wsgi_pythonanywhere.py    PythonAnywhere 用 WSGI テンプレート
Procfile                  Render / Railway 用
```

### 同時アクセスについて

リクエストごとに `jobs/<id>/` を作り、環境変数 `CHECKSHEET_ROOT` で
そこを参照させて処理する（`src/config.py` の `_project_root()` を参照）。
複数人が同時にアップロードしてもファイルは混ざらない。

### 生成物の保持

`jobs/` 内の作業フォルダは **6時間で自動削除** される。
また PythonAnywhere / Render ともディスクは永続保証が無いため、
必要なファイルはその場でダウンロードすること。

---

## マニュアルを更新したとき

マニュアル本体は別リポジトリ (`mms-project/pdef/manual/japan.html`) が正。
更新したら以下でこちらへ反映する。

```bash
cp <mms-project>/pdef/manual/japan.html static/manual/manual.html

"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  --headless --disable-gpu --no-pdf-header-footer \
  --print-to-pdf="static/manual/manual.pdf" \
  "file://$(pwd)/static/manual/manual.html"
```

PDF は実行時に生成せず、あらかじめ作ったものを配信している
（サーバ側に PDF 生成ライブラリを入れずに済むため）。
