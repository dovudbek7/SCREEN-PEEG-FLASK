"""二重申請チェック (duplicate-application detection, cross-report) — §6.4.

全レポート横断で 明細(レッグ) の署名を突き合わせ, 同一出張費の二重計上を検知する.

判定方針:
  - 強キー (strong key) = (社員番号 or 入力者名_norm, 明細日付, norm(出発地),
    norm(到着地), 交通機関, 金額). 異なる伝票 もしくは 同一伝票内の異なる明細No
    の2つ以上のレッグが同じ強キーを共有 → NG '二重申請疑い'.
  - 弱キー (weak key) = (社員番号 or 入力者名_norm, 明細日付, 到着地顧客番号).
    異なる伝票 にまたがって共有 → NEEDS_CHECK '同日同顧客の重複の可能性'.
  - 金額==0 の明細は対象外 (日当選択行・非金額行のノイズを除外).
  - 各レポートの CheckResult は配下レッグの重複判定の worst. 何も無ければ OK.

依存: 標準ライブラリのみ (collections). 名前正規化は normalize.norm に一元化.
"""
from __future__ import annotations

from collections import defaultdict

from models import (
    ExpenseReport,
    CheckResult,
    OK,
    NG,
    NEEDS_CHECK,
    worst_status,
)
from normalize import norm
from config import Config


def _actor_key(r: ExpenseReport) -> str:
    """申請者識別キー. 社員番号があれば優先, 無ければ入力者名の正規化値.

    社員番号は下流の従業員突合で付与される. 未突合でも入力者名で横断比較できる.
    """
    return r.employee_id or r.inputter_name_norm


def _strong_key(actor: str, leg) -> tuple:
    """強キー: 同一申請者・同日・同区間・同交通機関・同金額.

    金額==0 / 日付欠落 のレッグは呼び出し側で除外済みを前提とする.
    """
    return (
        actor,
        leg.leg_date,
        norm(leg.origin_raw),
        norm(leg.dest_raw),
        leg.transport,
        leg.amount,
    )


def _weak_dest(leg) -> str | None:
    """弱キーの到着識別子. 顧客番号があれば優先, 無ければ正規化到着地名.

    顧客マスタ未突合の到着地 (多数) を弱キーから漏らさないためのフォールバック.
    """
    if leg.dest_customer_no:
        return f"C:{leg.dest_customer_no}"
    if leg.dest_raw:
        d = norm(leg.dest_raw)
        return f"N:{d}" if d else None
    return None


def _weak_key(actor: str, leg) -> tuple:
    """弱キー: 同一申請者・同日・同到着(顧客番号 or 正規化地名)."""
    return (actor, leg.leg_date, _weak_dest(leg))


def _member(r: ExpenseReport, leg) -> dict:
    """重複グループのメンバ行 (シート5 で消費する辞書形)."""
    return {
        "voucher_no": r.voucher_no,
        "leg_no": leg.leg_no,
        "inputter": r.inputter_name_raw,
        "employee_id": r.employee_id,
        "date": leg.leg_date.isoformat() if leg.leg_date else "",
        "origin": leg.origin_raw or "",
        "dest": leg.dest_raw or "",
        "transport": leg.transport,
        "amount": leg.amount,
    }


def find_duplicate_groups(all_reports: list[ExpenseReport]) -> list[dict]:
    """二重申請の重複グループを返す (シート5 用).

    各グループ: {group_id, key, members:[{voucher_no, leg_no, ...}], severity}.
      - 強キー一致グループ severity=NG.
      - 弱キー一致グループ severity=NEEDS_CHECK (強で既に NG のものは弱から除外しない
        が, グループは別立て. シート5 は両方を一覧表示する).

    対象: 金額>0 かつ 明細日付ありのレッグのみ.
    """
    strong_buckets: dict[tuple, list[tuple[ExpenseReport, object]]] = defaultdict(list)
    weak_buckets: dict[tuple, list[tuple[ExpenseReport, object]]] = defaultdict(list)

    for r in all_reports:
        actor = _actor_key(r)
        for leg in r.legs:
            if leg.amount <= 0 or leg.leg_date is None:
                continue
            strong_buckets[_strong_key(actor, leg)].append((r, leg))
            if _weak_dest(leg) is not None:
                weak_buckets[_weak_key(actor, leg)].append((r, leg))

    groups: list[dict] = []
    gid = 0

    # --- 強キー: 異なる伝票 にまたがる同一内容レッグ2件以上で二重申請 ---
    # 同一伝票内の同区間反復 (往復・複数回移動) は正当な計上のため対象外.
    for key, items in strong_buckets.items():
        if len(items) < 2:
            continue
        vouchers = {r.voucher_no for r, leg in items}
        if len(vouchers) < 2:
            continue
        gid += 1
        groups.append({
            "group_id": f"D{gid:03d}",
            "key": key,
            "members": [_member(r, leg) for r, leg in items],
            "severity": NG,
        })

    # --- 弱キー: 異なる伝票 にまたがる同日同顧客のみ (同一伝票内は対象外) ---
    for key, items in weak_buckets.items():
        vouchers = {r.voucher_no for r, leg in items}
        if len(vouchers) < 2:
            continue
        gid += 1
        groups.append({
            "group_id": f"W{gid:03d}",
            "key": key,
            "members": [_member(r, leg) for r, leg in items],
            "severity": NEEDS_CHECK,
        })

    return groups


def check_duplicate(all_reports: list[ExpenseReport], cfg: Config) -> dict[str, CheckResult]:
    """全レポート横断の二重申請チェック. voucher_no -> CheckResult を返す.

    各レポートの判定は配下レッグが属する重複グループの worst:
      - 強キー重複に関与 → NG '二重申請疑い'.
      - 弱キー重複のみ → NEEDS_CHECK '同日同顧客の重複の可能性'.
      - いずれも無し → OK.
    evidence には関与した重複相手 (伝票No/明細No) を機械可読に格納する.
    """
    groups = find_duplicate_groups(all_reports)

    # voucher_no -> {status, 関与グループ情報} を集約
    findings: dict[str, dict] = {}

    for g in groups:
        sev = g["severity"]
        # このグループに属する (伝票No, 明細No) 集合
        member_ids = [(m["voucher_no"], m["leg_no"]) for m in g["members"]]
        for m in g["members"]:
            v = m["voucher_no"]
            f = findings.setdefault(v, {
                "statuses": [],
                "ng_groups": [],
                "needs_groups": [],
            })
            f["statuses"].append(sev)
            # 自分以外の相手レッグ
            others = [
                {"voucher_no": ov, "leg_no": ol}
                for (ov, ol) in member_ids
                if not (ov == v and ol == m["leg_no"])
            ]
            entry = {
                "group_id": g["group_id"],
                "leg_no": m["leg_no"],
                "counterparts": others,
            }
            if sev == NG:
                f["ng_groups"].append(entry)
            else:
                f["needs_groups"].append(entry)

    results: dict[str, CheckResult] = {}
    for r in all_reports:
        f = findings.get(r.voucher_no)
        if not f:
            results[r.voucher_no] = CheckResult(
                status=OK,
                detail="二重申請の疑いなし。",
                evidence={},
            )
            continue

        status = worst_status(f["statuses"])
        if status == NG:
            n = len(f["ng_groups"])
            detail = f"二重申請疑い: 強一致の重複明細が {n} 件。"
            suggestion = "同一日・同一区間・同額の明細が他伝票と重複していないか確認し、必要なら一方を取消。"
        elif status == NEEDS_CHECK:
            n = len(f["needs_groups"])
            detail = f"同日同顧客の重複の可能性: 要確認の重複明細が {n} 件。"
            suggestion = "同日・同一顧客先への複数申請の妥当性を確認。"
        else:  # 念のため (理論上ここには来ない)
            detail = "二重申請の疑いなし。"
            suggestion = ""

        results[r.voucher_no] = CheckResult(
            status=status,
            detail=detail,
            evidence={
                "ng": f["ng_groups"],
                "needs_check": f["needs_groups"],
            },
            suggestion=suggestion,
        )

    return results
