"""承認者ルール照合 (approver-rule matching).

20期 承認者一覧 (ApproverRule) を 氏名(正規化) でインデックス化し,
ある従業員に対する「出張命令者 (期待される出張承認者)」を引けるようにする.

照合キーは normalize.norm() で正規化した 氏名. ローダ側 (loaders.approver_loader)
が ApproverRule.employee_name_norm を既に norm() 済みで提供する前提だが,
ここでは念のため norm() を再適用してキーを統一する.
"""
from __future__ import annotations

from models import ApproverRule
from normalize import norm


def build_approver_index(rules: list[ApproverRule]) -> dict[str, ApproverRule]:
    """ApproverRule 群を norm(氏名) -> ApproverRule の辞書に索引化.

    同一 氏名 が重複する場合は後勝ち (last wins). 重複は判定の本質に
    影響しないため警告収集は行わない.
    """
    idx: dict[str, ApproverRule] = {}
    for rule in rules:
        key = norm(rule.employee_name_norm or rule.employee_name_raw)
        if not key:
            continue
        idx[key] = rule  # 重複は後勝ち
    return idx


def expected_trip_approver(
    employee_name_norm: str, idx: dict[str, ApproverRule]
) -> ApproverRule | None:
    """従業員 (正規化氏名) に対応する出張命令者ルールを返す. 無ければ None.

    引数 employee_name_norm が未正規化でも引けるよう norm() を通す.
    """
    if not employee_name_norm:
        return None
    return idx.get(norm(employee_name_norm))
