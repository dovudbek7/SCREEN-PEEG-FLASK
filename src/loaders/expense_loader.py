"""出張精算CSV ローダ (expense report loader).

楽々精算からエクスポートした出張精算CSV (cp932, 30列) を読み込み,
伝票No 単位で ExpenseReport / 明細単位で ExpenseLeg に整形する.

列はインデックスではなくヘッダ名の「コアトークン」を contains 照合して
解決する (防御的: 'ヘッダ情報:'/'明細情報:' 接頭辞や '(...)' 接尾辞があっても可).
"""
from __future__ import annotations

import csv
from collections import defaultdict

from models import (
    ExpenseReport,
    ExpenseLeg,
    ApprovalEntry,
    MOVEMENT_TRANSPORTS,
)
from normalize import (
    norm,
    norm_approver,
    parse_money,
    parse_jp_date,
    parse_time,
    validate_time_input,
)
from config import Config

# 承認スロット数 (承認実行者N名 / 承認日N, N=1..5)
_APPROVAL_SLOTS = 5


def _build_index(header: list[str]) -> dict[str, int]:
    """ヘッダ名 -> 列インデックスの辞書を構築 (生のヘッダ名キー)."""
    return {h: i for i, h in enumerate(header)}


def _find_col(header: list[str], token: str, exclude: str | None = None) -> int:
    """コアトークンを contains 照合して列インデックスを返す.

    exclude 指定時はそのトークンを含む列を除外する (例: '承認実行者' を
    検索する際に '承認日' 等の混同を避けるための保険).
    一致が無ければ ValueError.
    """
    for i, h in enumerate(header):
        if token in h and (exclude is None or exclude not in h):
            return i
    raise ValueError(f"列が見つかりません: token={token!r}")


def _find_col_optional(header: list[str], token: str) -> int:
    """コアトークンを contains 照合して列インデックスを返す (無ければ -1).

    通常経路到着時間のように現時点で顧客側に列自体が無い項目向け.
    列が追加された時点でトークンに一致し自動的に読み込まれる.
    """
    for i, h in enumerate(header):
        if token in h:
            return i
    return -1


def _cell(row: list[str], idx: int) -> str:
    """行から列値を安全に取得 (範囲外は空文字)."""
    return row[idx] if 0 <= idx < len(row) else ""


def _blank_to_none(s: str) -> str | None:
    """空文字を None に正規化 (手当CD等の TEXT 値はそのまま保持)."""
    if s is None:
        return None
    t = s.strip()
    return t if t else None


def load_expense_reports(path: str, cfg: Config) -> list[ExpenseReport]:
    """出張精算CSV を読み込み, 伝票No 単位の ExpenseReport リストを返す.

    path: 出張精算CSV のパス (cfg.expense_csv_path).
    cfg : 設定 (csv_encoding 等).
    戻り値は voucher_no 昇順でソート済み.
    """
    # --- CSV を cp932 で開いてヘッダと全データ行を読む ---
    with open(path, encoding=cfg.csv_encoding, newline="") as f:
        reader = csv.reader(f)
        header = next(reader)
        data_rows = list(reader)

    # --- 必要列をコアトークンの contains 照合で解決 (防御的) ---
    ix_voucher      = _find_col(header, "伝票No")
    ix_application  = _find_col(header, "申請No")
    ix_leg_count    = _find_col(header, "行数")
    ix_leg_no       = _find_col(header, "明細No")
    ix_leg_date     = _find_col(header, "明細日付")
    ix_time_start   = _find_col(header, "開始")
    ix_time_end     = _find_col(header, "終了")
    ix_origin       = _find_col(header, "出発地")
    ix_dest         = _find_col(header, "到着地")
    ix_transport    = _find_col(header, "交通機関")
    ix_amount       = _find_col(header, "金額")
    ix_receipt      = _find_col(header, "証票")
    ix_a1           = _find_col(header, "手当1CD")
    ix_a2           = _find_col(header, "手当2CD")
    ix_a3           = _find_col(header, "手当3CD")
    ix_remark       = _find_col(header, "フリー")
    ix_account      = _find_col(header, "勘定科目名")
    ix_subtotal     = _find_col(header, "小計")
    ix_total        = _find_col(header, "合計", exclude="手当")   # 手当計 と区別
    ix_allowance_total = _find_col(header, "手当計")
    ix_normal_arrival = _find_col_optional(header, "通常経路到着")  # 現時点で列自体が無い場合あり
    # 入力者名 (旧形式) / 申請者名 (新形式) の両対応
    try:
        ix_inputter = _find_col(header, "入力者名")
    except ValueError:
        ix_inputter = _find_col(header, "申請者名")
    # 申請者CD (新CSV のみ; 旧CSV にはない)
    try:
        ix_applicant_cd = _find_col(header, "申請者CD")
    except ValueError:
        ix_applicant_cd = -1

    # 承認スロット (承認実行者N名 / 承認日N) を 1..5 で解決
    ix_approver = []
    ix_approval_date = []
    for n in range(1, _APPROVAL_SLOTS + 1):
        ix_approver.append(_find_col(header, f"承認実行者{n}"))
        ix_approval_date.append(_find_col(header, f"承認日{n}"))

    # --- 明細行を ExpenseLeg に変換し, 伝票No でグループ化 ---
    legs_by_voucher: dict[str, list[ExpenseLeg]] = defaultdict(list)
    first_row_by_voucher: dict[str, list[str]] = {}

    for row in data_rows:
        voucher_no = _cell(row, ix_voucher).strip()
        if not voucher_no:
            continue  # 伝票No 欠落行はスキップ

        if voucher_no not in first_row_by_voucher:
            first_row_by_voucher[voucher_no] = row

        # 明細No. (パース不能は 0)
        try:
            leg_no = int(str(_cell(row, ix_leg_no)).strip() or 0)
        except ValueError:
            leg_no = 0

        transport = _cell(row, ix_transport).strip()
        origin_raw = _blank_to_none(_cell(row, ix_origin))
        dest_raw = _blank_to_none(_cell(row, ix_dest))
        receipt_label = _cell(row, ix_receipt).strip()

        # 移動レッグ: 交通機関が移動系 かつ 出発地・到着地が両方存在
        is_movement = (
            transport in MOVEMENT_TRANSPORTS
            and origin_raw is not None
            and dest_raw is not None
        )

        normal_arrival, normal_arrival_status = validate_time_input(
            _cell(row, ix_normal_arrival) if ix_normal_arrival >= 0 else ""
        )

        leg = ExpenseLeg(
            voucher_no=voucher_no,
            leg_no=leg_no,
            leg_date=parse_jp_date(_cell(row, ix_leg_date)),
            time_start=parse_time(_cell(row, ix_time_start)),
            time_end=parse_time(_cell(row, ix_time_end)),
            origin_raw=origin_raw,
            dest_raw=dest_raw,
            transport=transport,
            amount=parse_money(_cell(row, ix_amount)),
            subtotal=parse_money(_cell(row, ix_subtotal)),
            receipt_label=receipt_label,
            has_receipt=("あり" in receipt_label),
            # 手当CD は zero-pad TEXT のまま保持 (空は None)
            allowance_cd_perdiem=_blank_to_none(_cell(row, ix_a1)),
            allowance_cd_lodging=_blank_to_none(_cell(row, ix_a2)),
            allowance_cd_stay=_blank_to_none(_cell(row, ix_a3)),
            remark=_blank_to_none(_cell(row, ix_remark)),
            account_name=_cell(row, ix_account).strip(),
            is_movement_leg=is_movement,
            normal_route_arrival=normal_arrival,
            normal_route_arrival_status=normal_arrival_status,
        )
        legs_by_voucher[voucher_no].append(leg)

    # --- 伝票No 単位で ExpenseReport を構築 ---
    reports: list[ExpenseReport] = []
    for voucher_no, legs in legs_by_voucher.items():
        head = first_row_by_voucher[voucher_no]

        # 明細No 昇順にソート
        legs.sort(key=lambda lg: lg.leg_no)

        # 行数 (宣言上の明細数, パース不能は 0)
        try:
            declared = int(str(_cell(head, ix_leg_count)).strip() or 0)
        except ValueError:
            declared = 0

        # 承認エントリ: 非空の承認実行者N名 スロットのみ
        approvers: list[ApprovalEntry] = []
        for slot in range(1, _APPROVAL_SLOTS + 1):
            raw = _cell(head, ix_approver[slot - 1]).strip()
            if not raw:
                continue
            name_norm, is_confirm = norm_approver(raw)
            approvers.append(
                ApprovalEntry(
                    slot=slot,
                    approver_name_raw=raw,
                    approver_name_norm=name_norm,
                    approval_date=parse_jp_date(_cell(head, ix_approval_date[slot - 1])),
                    is_confirm_only=is_confirm,
                )
            )

        is_approved = len(approvers) > 0

        # 明細日付の最小・最大 (None は除外)
        leg_dates = [lg.leg_date for lg in legs if lg.leg_date is not None]
        date_min = min(leg_dates) if leg_dates else None
        date_max = max(leg_dates) if leg_dates else None

        inputter_raw = _cell(head, ix_inputter).strip()
        applicant_cd = _blank_to_none(_cell(head, ix_applicant_cd)) if ix_applicant_cd >= 0 else None

        reports.append(
            ExpenseReport(
                voucher_no=voucher_no,
                application_no=_blank_to_none(_cell(head, ix_application)),
                declared_leg_count=declared,
                actual_leg_count=len(legs),
                total_amount=parse_money(_cell(head, ix_total)),
                computed_total=sum(lg.amount for lg in legs),
                allowance_total_declared=parse_money(_cell(head, ix_allowance_total)),
                inputter_name_raw=inputter_raw,
                inputter_name_norm=norm(inputter_raw),
                legs=legs,
                approvers=approvers,
                is_approved=is_approved,
                approval_status="承認済" if is_approved else "未承認(PENDING)",
                date_min=date_min,
                date_max=date_max,
                source_file=path,
                applicant_cd=applicant_cd,
            )
        )

    # 伝票No 昇順でソートして返す
    reports.sort(key=lambda r: r.voucher_no)
    return reports
