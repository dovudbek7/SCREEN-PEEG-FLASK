"""承認ルート check (§6.6).

申請者 (入力者) を 20期承認者名簿 (aidx) で引いて想定 出張命令者 を求め,
伝票の実承認者集合に含まれているかを判定する.

判定方針:
- 未承認の伝票は NG.
- 申請者がマスタ/名簿に未登録なら判定不能 (UNMATCHED).
- 想定承認者が実承認者に居れば OK, 居なければ 要確認.
  - 想定が「確認のみ」承認者で, config 上 確認=承認 と見なさない場合は
    その旨を明示した 要確認 を返す.
"""
from __future__ import annotations

from models import (
    ExpenseReport, CheckResult,
    OK, NG, NEEDS_CHECK, UNMATCHED,
)
from config import Config


def check_approval_route(r: ExpenseReport, aidx: dict, eidx: dict, cfg: Config) -> CheckResult:
    """承認ルート判定 (§6.6).

    引数:
        r:    対象 出張精算レポート.
        aidx: 承認者ルール索引 norm(氏名) -> ApproverRule.
        eidx: 従業員索引 norm(氏名) -> Employee.
        cfg:  設定 (confirm_only_counts_as_approval を参照).
    """
    # --- 未承認は 要確認 (§3.4 承認ルートの語彙は OK/要確認 のみ; NG は出さない) ---
    if not r.is_approved:
        return CheckResult(
            status=NEEDS_CHECK,
            detail="未承認(PENDING): 承認実行者が未登録",
            evidence={"承認状態": r.approval_status, "実承認者": []},
            suggestion="承認手続きが未完了です。承認実行のうえ再提出してください。",
        )

    # --- 申請者(入力者)を名簿で引く ---
    # 別名解決後のマスタ氏名 (resolved_name_norm) を優先キーにする.
    # 入力者名の表記揺れ (例 '張学シン' vs マスタ '張学鑫') を吸収するため.
    lookup_key = r.resolved_name_norm or r.inputter_name_norm
    expected = aidx.get(lookup_key) or aidx.get(r.inputter_name_norm)
    if expected is None or r.employee_match_status == UNMATCHED:
        return CheckResult(
            status=UNMATCHED,
            detail="承認者ルート判定不可: 申請者がマスタ/名簿に未登録",
            evidence={
                "入力者名": r.inputter_name_raw,
                "正規化値": r.inputter_name_norm,
                "従業員突合": r.employee_match_status,
                "名簿登録": expected is not None,
            },
            suggestion="申請者を従業員マスタ/承認者名簿(20期)に登録のうえ再判定してください。",
        )

    # --- 想定 出張命令者 と 実承認者集合 を比較 ---
    expected_norm = expected.trip_approver_norm
    actual = {a.approver_name_norm for a in r.approvers}
    actual_raw = [a.approver_name_raw for a in r.approvers]
    matched = expected_norm in actual
    confirm_only = (
        expected.trip_approver_is_confirm_only
        and not cfg.confirm_only_counts_as_approval
    )

    # --- 確認のみ承認者は, 実承認者に含まれていても正式承認にならない ---
    # (名前一致を先に OK 判定すると確認のみ分岐が死ぬため, 確認のみを先に評価)
    if matched and confirm_only:
        return CheckResult(
            status=NEEDS_CHECK,
            detail=(
                f"確認のみ承認者: 想定={expected.trip_approver_raw}(確認) "
                f"実={actual_raw}"
            ),
            evidence={
                "想定承認者": expected.trip_approver_raw,
                "実承認者": actual_raw,
                "確認のみ": True,
            },
            suggestion=(
                f"想定承認者 {expected.trip_approver_raw} は確認のみのため正式承認に"
                "該当しません。正規の出張命令者による承認をご確認ください。"
            ),
        )

    if matched:
        return CheckResult(
            status=OK,
            detail=f"想定承認者 {expected.trip_approver_raw} が承認済",
            evidence={
                "想定承認者": expected.trip_approver_raw,
                "実承認者": actual_raw,
                "確認のみ": expected.trip_approver_is_confirm_only,
            },
            suggestion="",
        )

    # --- 不一致: 想定が確認のみ承認者でも実に居ない → 通常の不一致として扱う ---
    if confirm_only:
        return CheckResult(
            status=NEEDS_CHECK,
            detail=(
                f"確認のみ承認者: 想定={expected.trip_approver_raw}(確認) "
                f"実={actual_raw}"
            ),
            evidence={
                "想定承認者": expected.trip_approver_raw,
                "実承認者": actual_raw,
                "確認のみ": True,
            },
            suggestion=(
                f"想定承認者 {expected.trip_approver_raw} は確認のみのため正式承認に"
                "該当しません。正規の出張命令者による承認をご確認ください。"
            ),
        )

    return CheckResult(
        status=NEEDS_CHECK,
        detail=f"承認者不一致: 想定={expected.trip_approver_raw} 実={actual_raw}",
        evidence={
            "想定承認者": expected.trip_approver_raw,
            "実承認者": actual_raw,
            "確認のみ": expected.trip_approver_is_confirm_only,
        },
        suggestion=(
            f"想定の出張命令者 {expected.trip_approver_raw} による承認が確認できません。"
            "承認ルートをご確認ください。"
        ),
    )
