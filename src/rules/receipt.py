"""領収書(証票)チェック — §6.5.

明細ごとに「金額が発生しているのに領収書が無い」状態を検知する.
旅費規定(領収書要否閾値)が未提供のため, 高額判定は config 駆動.
規程が無い場合でも, 免除対象外で金額が発生している明細に領収書が無いのは
構造的に検知可能なため 要確認 を出す(規程未提供の旨を detail に明記).
"""
from __future__ import annotations

from models import (
    CheckResult,
    ExpenseReport,
    OK,
    NEEDS_CHECK,
    RULE_MISSING,
    worst_status,
)
from config import Config


def check_receipt(r: ExpenseReport, cfg: Config) -> CheckResult:
    """領収書の有無を明細単位で判定し, 総合(最悪)ステータスを返す.

    判定ロジック(§6.5):
      - 金額==0      → 対象外(OK 寄与).
      - 領収書あり   → OK.
      - 金額>0 かつ 領収書なし:
          * 交通機関が免除対象(運賃/IC 系) → 原則 OK.
              ただし規程あり かつ 金額>=閾値 → 要確認(高額は免除でも領収書要).
          * 免除対象外:
              - 規程あり: 金額>=閾値 → 要確認(高額・領収書なし), 未満 → OK.
              - 規程なし: 要確認(領収書要否閾値が規程未提供) を出す
                          (免除対象外・有償明細の領収書欠落は構造的に検知可能).
    """
    has_rule = cfg.has_receipt_rule()
    threshold = cfg.receipt_required_above
    exempt = set(cfg.receipt_exempt_transports or ())
    # 規程未提供時のフォールバック閾値
    high_value = getattr(cfg, "receipt_high_value_provisional", 10000)
    min_flag = getattr(cfg, "receipt_min_amount_to_flag", 0)

    statuses: list[str] = []
    no_receipt_paid_legs = 0      # 有償・領収書なしの明細数
    flagged_legs: list[dict] = []  # 要確認/規程未提供になった明細の根拠
    rule_missing_seen = False

    for leg in r.legs:
        amount = leg.amount or 0

        # 金額が無い明細は領収書要否の対象外
        if amount <= 0:
            statuses.append(OK)
            continue

        # 領収書がある → OK
        if leg.has_receipt:
            statuses.append(OK)
            continue

        # ここから: 金額>0 かつ 領収書なし
        no_receipt_paid_legs += 1
        is_exempt = leg.transport in exempt

        if is_exempt:
            # 運賃/IC 系は原則免除. ただし高額なら要確認.
            #   規程あり: 閾値超過で要確認. 規程なし: 暫定高額閾値で見逃し防止.
            limit = threshold if has_rule else high_value
            if amount >= limit:
                if not has_rule:
                    rule_missing_seen = True
                statuses.append(NEEDS_CHECK)
                flagged_legs.append({
                    "leg_no": leg.leg_no,
                    "transport": leg.transport,
                    "amount": amount,
                    "reason": (
                        "免除対象だが高額・領収書なし"
                        if has_rule else
                        "免除対象だが高額・領収書なし(暫定閾値; 規程未提供)"
                    ),
                })
            else:
                statuses.append(OK)
            continue

        # 免除対象外 かつ 有償 かつ 領収書なし
        if has_rule:
            if amount >= threshold:
                statuses.append(NEEDS_CHECK)
                flagged_legs.append({
                    "leg_no": leg.leg_no,
                    "transport": leg.transport,
                    "amount": amount,
                    "reason": "高額・領収書なし",
                })
            else:
                statuses.append(OK)
        else:
            # 規程未提供: 閾値で高額判定は不能だが, 領収書欠落自体は検知可能.
            #   ただし少額付随費(宿泊税/駐車代等)は過剰検知になるため min_flag 未満は除外.
            if amount < min_flag:
                statuses.append(OK)
            else:
                rule_missing_seen = True
                statuses.append(NEEDS_CHECK)
                flagged_legs.append({
                    "leg_no": leg.leg_no,
                    "transport": leg.transport,
                    "amount": amount,
                    "reason": "免除対象外・領収書なし(領収書要否閾値が規程未提供)",
                })

    status = worst_status(statuses)

    # フラグ種別 (文面を正確にするため flagged_legs の reason から判定)
    has_exempt_high = any("免除対象だが高額" in f["reason"] for f in flagged_legs)
    has_nonexempt = any("免除対象外" in f["reason"] for f in flagged_legs)

    # detail の組み立て (要確認になった明細数を報告; 免除でOKの明細は数えない)
    if status == OK:
        detail = "領収書: 問題なし"
    else:
        parts = []
        if rule_missing_seen:
            parts.append("領収書要否閾値が規程未提供(暫定判定)")
        parts.append(f"要確認の有償・領収書なし明細 {len(flagged_legs)} 件")
        ns = "/".join(f"明細{f['leg_no']}" for f in flagged_legs)
        if ns:
            parts.append(f"対象: {ns}")
        detail = "領収書: " + " / ".join(parts)

    evidence = {
        "no_receipt_paid_legs": no_receipt_paid_legs,
        "flagged_legs": flagged_legs,
        "has_receipt_rule": has_rule,
        "receipt_required_above": threshold,
        "leg_count": len(r.legs),
    }

    suggestion = ""
    if status == NEEDS_CHECK:
        kinds = []
        if has_exempt_high:
            kinds.append("免除交通機関だが高額")
        if has_nonexempt:
            kinds.append("免除対象外")
        kind_s = "・".join(kinds) if kinds else "高額"
        if rule_missing_seen:
            suggestion = (
                f"{kind_s}の有償明細に領収書がありません。旅費規定の領収書要否を"
                "確認の上、領収書の添付を依頼してください。"
            )
        else:
            suggestion = f"{kind_s}の明細に領収書がありません。領収書の添付を依頼してください。"

    return CheckResult(
        status=status,
        detail=detail,
        evidence=evidence,
        suggestion=suggestion,
    )
