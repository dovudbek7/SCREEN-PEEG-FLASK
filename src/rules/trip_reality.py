"""出張実態チェック (trip-reality check). データ契約 §6.1.

出張精算レポートの「出張した日に実際に出社/移動の勤怠実態があるか」「休暇日に
出張を計上していないか」「到着地が顧客マスタと照合できるか」を判定する.

重要な前提 (既知の欠落): 勤怠データは 4月分(A) と 6月分(B) のみで, 出張月の
5月 を含まない. 対象日の月が勤怠データに存在しない場合は NG を出さず
ATT_MISSING(未確認(勤怠データ欠落)) で暫定評価し, 顧客照合と証票で補強する.

依存: 標準ライブラリのみ. 在席状態の派生は attendance_loader 側で完了済みで,
ここでは AttendanceLookup を通して (入力者名, 日付) を引くだけ.
"""
from __future__ import annotations

from models import (
    ExpenseReport,
    ExpenseLeg,
    AttendanceLookup,
    CheckResult,
    OK,
    NEEDS_CHECK,
    NG,
    UNMATCHED,
    ATT_MISSING,
    PRESENCE_OFFICE,
    PRESENCE_TELEWORK,
    PRESENCE_LEAVE,
    PRESENCE_HOLIDAY_WORK,
    MOVEMENT_TRANSPORTS,
)
from config import Config

# 出張実態を肯定する在席状態 (在席/休日出勤). これらは出社・移動実態とみなす.
_PRESENCE_CORROBORATE = {PRESENCE_OFFICE, PRESENCE_TELEWORK, PRESENCE_HOLIDAY_WORK}


def _is_dest_leg(leg: ExpenseLeg) -> bool:
    """到着地を持つレッグ (移動先=訪問先候補) か判定する."""
    return leg.dest_raw is not None or leg.is_movement_leg


def _is_customer_candidate_leg(leg: ExpenseLeg) -> bool:
    """顧客照合が実際に試行されたレッグか.

    移動レッグ (電車･ﾊﾞｽ等) の到着地は経由駅であり訪問先ではないため
    enrich_reports() 側で顧客照合自体をスキップし常に UNMATCHED を入れている
    (checksheet.py 参照). これを「全到着地が未突合」の判定対象に含めると,
    移動レッグを持つ (=ほぼ全ての) 伝票が常に要確認になってしまうため,
    ここでは移動レッグを除外し, 実際に照合が試みられた活動レッグのみを見る.
    """
    is_move = leg.is_movement_leg or leg.transport in MOVEMENT_TRANSPORTS
    if is_move:
        return False
    return leg.dest_raw is not None or leg.origin_raw is not None


def _unmatched_breakdown(legs: list[ExpenseLeg]) -> str:
    """未突合レッグを「未入力」と「該当取引先なし」に分けて明細No.付きで説明する.

    空欄(到着地・出発地とも文字列が無い; enrich_reports の照合対象
    target = dest_raw or origin_raw と同じ判定)と, 文字はあるがマスタに
    十分な一致度で当たらなかった場合を区別する (2026-07-10 客先質問対応).
    """
    blank_legs = [
        leg.leg_no for leg in legs
        if not (leg.dest_raw or "").strip() and not (leg.origin_raw or "").strip()
    ]
    notfound_legs = [
        leg.leg_no for leg in legs
        if (leg.dest_raw or "").strip() or (leg.origin_raw or "").strip()
    ]
    parts = []
    if blank_legs:
        parts.append(f"到着地未入力(明細{','.join(str(n) for n in blank_legs)})")
    if notfound_legs:
        parts.append(f"該当取引先なし(明細{','.join(str(n) for n in notfound_legs)})")
    return "; ".join(parts)


def _has_move_record(day) -> bool:
    """勤怠日次に移動時刻 (移動開始/終了, 移動2開始/終了) のいずれかがあるか."""
    return any((
        day.move_start is not None,
        day.move_end is not None,
        day.move2_start is not None,
        day.move2_end is not None,
    ))


def check_trip_reality(r: ExpenseReport, att: AttendanceLookup, cfg: Config) -> CheckResult:
    """出張実態を判定し CheckResult を返す (データ契約 §6.1).

    手順:
      1. 移動/到着レッグから出張日集合を収集.
      2. 出張日の月が勤怠データに無ければ ATT_MISSING で暫定評価
         (顧客照合と証票のみ). 全到着地が未突合なら detail に 要確認 理由を併記.
      3. 月が存在する日について (入力者名, 日付) を引き,
         休暇/欠勤 → NG(休暇日に出張計上), 在席/休日出勤/移動あり → OK 寄与,
         勤怠未入力(月はあるが行無し) → 要確認(勤怠未入力).
      4. 総合: 休暇衝突あれば NG; なければ 1日でも裏付けがあり全到着地が突合済なら OK;
         到着地が全て未突合なら 要確認; それ以外は ATT_MISSING/OK 準拠.
    """
    # 別名解決後のマスタ氏名を優先 (勤怠は社員マスタ名でキーされるため)
    name = r.resolved_name_norm or r.inputter_name_norm
    period = f"{r.date_min.isoformat() if r.date_min else '?'}〜" \
             f"{r.date_max.isoformat() if r.date_max else '?'}"

    # --- 1. 出張日集合 (移動/到着レッグの明細日付) ---
    dest_legs = [leg for leg in r.legs if _is_dest_leg(leg)]
    trip_dates = sorted({leg.leg_date for leg in dest_legs if leg.leg_date is not None})

    # --- 顧客照合状態 (移動レッグ=経由地を除いた, 実際に照合が試みられたレッグ) ---
    customer_legs = [leg for leg in r.legs if _is_customer_candidate_leg(leg)]
    matched_dest = [leg for leg in customer_legs if leg.dest_match_status != UNMATCHED]
    unmatched_legs = [leg for leg in customer_legs if leg.dest_match_status == UNMATCHED]
    all_dest_unmatched = bool(customer_legs) and not matched_dest

    evidence: dict = {
        "trip_dates": [d.isoformat() for d in trip_dates],
        "dest_leg_count": len(customer_legs),
        "matched_dest_count": len(matched_dest),
    }

    # 出張日が無い (移動・到着レッグが無い) 場合は評価対象外として要確認
    if not trip_dates:
        return CheckResult(
            status=NEEDS_CHECK,
            detail="移動・到着レッグが無く出張実態を判定できない。明細内容を確認。",
            evidence=evidence,
            suggestion=f"伝票No.{r.voucher_no}: 出張日(移動/到着明細)が特定できません。"
                       f"明細の到着地・交通機関をご確認ください。",
        )

    # --- 2. 月欠落判定 (5月ギャップ支配ケース) ---
    present_dates = [d for d in trip_dates if att.is_month_present(d)]
    missing_dates = [d for d in trip_dates if not att.is_month_present(d)]
    missing_months = sorted({(d.year, d.month) for d in missing_dates})
    evidence["missing_months"] = [f"{y:04d}-{m:02d}" for y, m in missing_months]

    if not present_dates:
        # どの出張日の月も勤怠に無い → 顧客照合と証票のみで暫定評価
        mm = ", ".join(f"{y:04d}-{m:02d}" for y, m in missing_months) or "対象月"
        detail = (
            f"対象月({mm})の出勤簿が未提供。顧客照合と証票のみで暫定評価。"
        )
        # 全到着地が未突合なら 要確認 理由を detail に併記 (NG は出さない)
        if all_dest_unmatched:
            detail += f" 到着地が全て顧客マスタ未突合のため要確認({_unmatched_breakdown(unmatched_legs)})。"
        elif matched_dest:
            detail += f" 到着地{len(matched_dest)}件は顧客マスタと突合済。"
        evidence["receipt_corroborated"] = any(leg.has_receipt for leg in dest_legs)
        return CheckResult(
            status=ATT_MISSING,
            detail=detail,
            evidence=evidence,
            suggestion=f"伝票No.{r.voucher_no}({period}): {mm}の出勤簿が未提供のため出張実態は"
                       f"暫定評価です。" + (
                           "到着地が顧客マスタと突合できないため、訪問先の妥当性を別途ご確認ください。"
                           if all_dest_unmatched else
                           "出勤簿提供後に最終確認をお願いします。"
                       ),
        )

    # --- 3. 月が存在する出張日を勤怠と照合 ---
    leave_dates: list = []          # 休暇/欠勤に出張計上
    leave_kinds: dict = {}          # 日付 -> 申請内容 (有休/欠勤 等)
    corroborated_dates: list = []   # 在席/休日出勤/移動ありで裏付け
    no_record_dates: list = []      # 月はあるが勤怠行が無い (勤怠未入力)

    for d in present_dates:
        day = att.get(name, d)
        if day is None:
            no_record_dates.append(d)
            continue
        # 休暇 (PRESENCE_LEAVE) または申請内容に 欠勤 → 休暇日に出張計上
        ac = day.application_content or ""
        if day.presence == PRESENCE_LEAVE or "欠勤" in ac:
            # 2026-08-07 客先指摘: 「休暇日」だと休日 (土日祝) と混同されるため,
            # どの申請で休みになっているのかが分かるよう申請内容も保持する.
            leave_dates.append(d)
            if ac:
                leave_kinds[d] = ac
        elif day.presence in _PRESENCE_CORROBORATE or _has_move_record(day):
            corroborated_dates.append(d)
        else:
            # 非労働日/不明など → 裏付けにも休暇にもならない
            no_record_dates.append(d)

    evidence["leave_kinds"] = {d.isoformat(): k for d, k in leave_kinds.items()}
    evidence["leave_dates"] = [d.isoformat() for d in leave_dates]
    evidence["corroborated_dates"] = [d.isoformat() for d in corroborated_dates]
    evidence["no_record_dates"] = [d.isoformat() for d in no_record_dates]

    # --- 4. 総合判定 ---
    # (a) 休暇日に出張計上 → NG
    if leave_dates:
        days_s = ", ".join(
            f"{d.isoformat()}({leave_kinds[d]})" if d in leave_kinds else d.isoformat()
            for d in leave_dates
        )
        return CheckResult(
            status=NG,
            detail=f"休暇取得日に出張を計上({days_s})。"
                   f"平日であっても勤怠に休暇/欠勤の申請があるため出張実態と整合しません。",
            evidence=evidence,
            suggestion=f"伝票No.{r.voucher_no}({period}): {days_s}は勤怠上 休暇/欠勤 の申請が"
                       f"入っています。出張実態と整合しないため、日付の誤りか休暇申請の"
                       f"取消をご確認ください。",
        )

    # (b) 裏付けあり かつ 全到着地が突合済 → OK
    if corroborated_dates and not all_dest_unmatched:
        return CheckResult(
            status=OK,
            detail=f"出張日{len(corroborated_dates)}件で出社/移動の勤怠実態を確認。"
                   f"到着地{len(matched_dest)}件は顧客マスタと突合済。",
            evidence=evidence,
        )

    # (c) 到着地が全て未突合 → 要確認
    if all_dest_unmatched:
        return CheckResult(
            status=NEEDS_CHECK,
            detail="到着地が全て顧客マスタ未突合のため出張先の妥当性が確認できない"
                   f"({_unmatched_breakdown(unmatched_legs)})。",
            evidence=evidence,
            suggestion=f"伝票No.{r.voucher_no}({period}): 到着地が顧客マスタと突合できません。"
                       f"訪問先(取引先)名の表記をご確認ください。",
        )

    # (d) 勤怠未入力の日が残る (月はあるが行無し) → 要確認
    if no_record_dates:
        days_s = ", ".join(d.isoformat() for d in no_record_dates)
        return CheckResult(
            status=NEEDS_CHECK,
            detail=f"勤怠未入力({days_s})のため出張実態を確認できない。",
            evidence=evidence,
            suggestion=f"伝票No.{r.voucher_no}({period}): {days_s}の勤怠が未入力です。"
                       f"出勤簿の入力状況をご確認ください。",
        )

    # (e) 裏付けあり (到着地突合は問わず) → OK
    if corroborated_dates:
        return CheckResult(
            status=OK,
            detail=f"出張日{len(corroborated_dates)}件で出社/移動の勤怠実態を確認。",
            evidence=evidence,
        )

    # (f) いずれにも該当しない → 暫定 (ATT_MISSING 準拠)
    return CheckResult(
        status=ATT_MISSING,
        detail="勤怠実態の裏付けが得られず暫定評価。",
        evidence=evidence,
        suggestion=f"伝票No.{r.voucher_no}({period}): 出張実態の裏付けが不足しています。"
                   f"出勤簿および訪問先をご確認ください。",
    )
