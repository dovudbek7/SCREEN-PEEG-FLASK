"""エントリポイント: 全データ読込 -> enrich -> 6観点判定 -> 7シートExcel出力.

実行: python3 src/main.py            (既定 config)
      python3 src/main.py --config conf.json
出力: out/出張精算_承認チェックシート_YYYYMMDD_HHMMSS.xlsx
"""
from __future__ import annotations

import argparse
import getpass
import os
import sys
from datetime import datetime

# src/ を import パスに追加 (フラット import 規約)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import load_config, Config
from normalize import norm
from loaders.expense_loader import load_expense_reports
from loaders.attendance_loader import load_attendance
from loaders.employee_loader import load_employees
from loaders.customer_loader import load_customers
from loaders.approver_loader import load_approver_rules
from checksheet import build_check_sheet
from excel_writer import write_excel
from html_writer_web import read_excel_and_write_html


def run(cfg: Config, stamp: str | None = None, with_uz: bool = False) -> str:
    import os as _os

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    import_log: list[dict] = []

    def _step(no, label, fn, path):
        """ローダを実行し取込ログ行を記録. 失敗は例外行を残して再送出."""
        print(f"[{no}/5] {label} 読込 ...", flush=True)
        try:
            data = fn()
            n = len(data)
            print(f"      -> {n} 件")
            import_log.append({
                "区分": label, "ファイル名": _os.path.basename(str(path)),
                "件数": n, "詳細": f"取込日時 {now}", "結果": "OK",
            })
            return data
        except Exception as e:  # noqa: BLE001
            import_log.append({
                "区分": label, "ファイル名": _os.path.basename(str(path)),
                "件数": 0, "詳細": f"{type(e).__name__}: {e}", "結果": "エラー",
            })
            raise

    reports = _step(1, "出張精算CSV", lambda: load_expense_reports(cfg.expense_csv_path, cfg), cfg.expense_csv_path)

    if cfg.target_approver_filter:
        target_norm = norm(cfg.target_approver_filter)
        before = len(reports)
        reports = [
            r for r in reports
            if any(a.slot == 1 and a.approver_name_norm == target_norm for a in r.approvers)
        ]
        print(f"      -> W列(承認実行者1名)『{cfg.target_approver_filter}』でフィルタ: {before}件 -> {len(reports)}件")
        import_log.append({
            "区分": "出張精算CSV(対象者フィルタ)",
            "ファイル名": _os.path.basename(str(cfg.expense_csv_path)),
            "件数": len(reports),
            "詳細": f"承認実行者1名='{cfg.target_approver_filter}' のみ対象 (フィルタ前 {before}件)",
            "結果": "OK",
        })

    attendance = _step(2, "勤怠(出勤簿)", lambda: load_attendance(cfg.attendance_paths),
                       " / ".join(_os.path.basename(p) for p in cfg.attendance_paths))
    employees = _step(3, "社員マスタ", lambda: load_employees(cfg.employee_master_path, cfg.master_password), cfg.employee_master_path)
    customers = _step(4, "顧客マスタ", lambda: load_customers(cfg.customer_master_path, cfg.master_password), cfg.customer_master_path)
    approvers = _step(5, "承認者名簿", lambda: load_approver_rules(cfg.approver_roster_path, cfg.approver_roster_sheet), cfg.approver_roster_path)

    print("[*] enrich + 判定 + 出力生成 ...", flush=True)
    sheet = build_check_sheet(reports, employees, customers, approvers, attendance, cfg,
                              import_log=import_log)

    stamp = stamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = os.path.join(cfg.output_dir, stamp)
    os.makedirs(out_dir, exist_ok=True)

    base_name = f"{cfg.output_prefix}_{stamp}"
    out_path  = os.path.join(out_dir, f"{base_name}.xlsx")
    html_ja   = os.path.join(out_dir, f"{base_name}_ja.html")

    write_excel(sheet, out_path, cfg)
    read_excel_and_write_html(out_path, html_ja, lang="ja")
    # ウズベク語版は社内確認用のため既定では出力しない (-uz 指定時のみ).
    if with_uz:
        html_uz = os.path.join(out_dir, f"{base_name}_uz.html")
        read_excel_and_write_html(out_path, html_uz, lang="uz")

    # サマリ統計
    # 01シートは日単位に展開され, 2日目以降の行は伝票単位の列が空欄になる。
    # 件数は伝票単位で数えたいので空欄行は除外する。
    overall = [r["総合判定"] for r in sheet["primary"] if r.get("総合判定")]
    from collections import Counter
    print("\n=== 総合判定 内訳 ===")
    for k, v in Counter(overall).most_common():
        print(f"  {k}: {v}")
    print(f"\n出力フォルダ: {out_dir}")
    return out_path


def _acquire_master_password(cfg: Config, max_attempts: int = 3) -> str:
    """社員マスタで実際に復号を試しながらパスワードを対話入力させる.

    誤ったパスワードは msoffcrypto の復号成功後, openpyxl のファイル読込段階で
    初めて失敗として顕在化するため, load_employees() を検証に使う。
    """
    for attempt in range(1, max_attempts + 1):
        pw = getpass.getpass("社員・顧客リストのパスワードを入力してください: ")
        try:
            load_employees(cfg.employee_master_path, pw)
            return pw
        except Exception:  # noqa: BLE001
            remaining = max_attempts - attempt
            if remaining > 0:
                print(f"[ERROR] パスワードが正しくありません。もう一度お試しください。(残り{remaining}回)")
            else:
                raise RuntimeError("パスワードの試行回数が上限に達しました。処理を中止します。")


def _resolve_master_password(cfg: Config, cli_password: str | None) -> str:
    """パスワードを解決する.

    優先順位: --password 指定 > config.py の既定値(復号できた場合) > 対話入力。
    既定値がマスタファイルの実際のパスワードと一致しない場合(=手動で変更された場合)は,
    自動的に対話入力にフォールバックする。
    """
    if cli_password:
        return cli_password
    if cfg.master_password:
        try:
            load_employees(cfg.employee_master_path, cfg.master_password)
            return cfg.master_password
        except Exception:  # noqa: BLE001
            print("[INFO] 既定のパスワードでは復号できませんでした。パスワードを入力してください。")
    return _acquire_master_password(cfg)


def _pause_before_exit(status: int) -> None:
    """実行ファイル(exe)をダブルクリックしたとき, 結果を読む前に
    コンソールが閉じてしまわないよう Enter を待つ.

    Windows のエクスプローラから起動するとプロセス終了と同時にウィンドウが
    閉じるため, これが無いと出力を確認できない。
    通常の python 実行時や CI (--no-pause / 対話端末でない場合) は待たない。
    """
    if not getattr(sys, "frozen", False):
        return
    try:
        if not sys.stdin.isatty():
            return
    except Exception:  # noqa: BLE001
        return
    print()
    if status == 0:
        print("✅ 完了しました。結果は out/ フォルダをご確認ください。")
    else:
        print("❌ エラーが発生しました。上記のメッセージをご確認ください。")
    try:
        input("Enterキーを押すと終了します... ")
    except (EOFError, KeyboardInterrupt):
        pass


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="出張精算 承認チェックシート生成")
    ap.add_argument("--config", default=None, help="JSON設定ファイル(任意)")
    ap.add_argument("--stamp", default=None, help="出力ファイル名のタイムスタンプ上書き")
    ap.add_argument("--password", default=None, help="社員・顧客リストの復号パスワード(省略時は対話入力)")
    ap.add_argument("-uz", "--uz", dest="uz", action="store_true",
                    help="ウズベク語版HTML(_uz.html)も出力する(既定は日本語版のみ)")
    ap.add_argument("--no-pause", action="store_true",
                    help="終了時に Enter 待ちをしない (CI/自動実行用)")
    args = ap.parse_args(argv)

    try:
        cfg = load_config(args.config)
    except FileNotFoundError as e:
        print(f"\n[ERROR] {e}", file=sys.stderr)
        return 1
    except Exception as e:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        print(f"\n[ERROR] {type(e).__name__}: {e}", file=sys.stderr)
        return 1

    try:
        cfg.master_password = _resolve_master_password(cfg, args.password)
    except Exception as e:  # noqa: BLE001
        print(f"\n[ERROR] {type(e).__name__}: {e}", file=sys.stderr)
        return 1

    try:
        run(cfg, args.stamp, with_uz=args.uz)
        return 0
    except Exception as e:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        print(f"\n[ERROR] {type(e).__name__}: {e}", file=sys.stderr)
        return 1


def _cli() -> int:
    status = main()
    if "--no-pause" not in sys.argv:
        _pause_before_exit(status)
    return status


if __name__ == "__main__":
    raise SystemExit(_cli())
