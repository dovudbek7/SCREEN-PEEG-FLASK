"""労務チェック (§6.2 labor / 勤怠整合).

出張レッグの移動時刻を勤怠 (AttendanceDay) と突き合わせ, 以下を判定する:
  - 移動日の勤怠が在席 (出社/テレワーク/休日出勤) と整合 → OK 寄与.
  - 非労働日/休暇に移動 (休日出勤の実態なし) → 要確認 寄与.
  - 当該月の勤怠が未提供 (5月ギャップ) → 未確認(勤怠データ欠落) 寄与.
  - 深夜帯 (cfg.late_night_*) の移動かつ深夜手当ヒント無し → 要確認 寄与.
  - 通常経路到着時間が入力ルール外、退勤時刻と不一致、または未入力 → 要確認 寄与
    (2026-07-10 客先指定: 未入力も確認が必要なため要確認とする).
  - 法定時間外/休日労働の上限判定は旅費規程・就業規則未提供のため 未確認(規程未提供).

総合は worst_status. すべてのレッグが裏付けされ問題が無ければ OK.
5月ギャップでは 未確認(勤怠データ欠落) が支配的になる.

依存: 標準ライブラリのみ. 名前正規化・時刻パースは normalize に一元化済み (上流付与済み).
"""
from __future__ import annotations

from datetime import time

from models import (
    ExpenseReport,
    AttendanceLookup,
    AttendanceDay,
    CheckResult,
    ExpenseLeg,
    MOVEMENT_TRANSPORTS,
    PRESENCE_OFFICE,
    PRESENCE_TELEWORK,
    PRESENCE_HOLIDAY_WORK,
    PRESENCE_LEAVE,
    PRESENCE_NONWORK,
    OK,
    NEEDS_CHECK,
    RULE_MISSING,
    ATT_MISSING,
    worst_status,
    effective_work_times,
)
from config import Config

# 在席 (移動を裏付ける) とみなす在席状態
_PRESENT = {PRESENCE_OFFICE, PRESENCE_TELEWORK, PRESENCE_HOLIDAY_WORK}
# 移動の裏付けとして疑義がある在席状態 (非労働日/休暇)
_NONWORK = {PRESENCE_NONWORK, PRESENCE_LEAVE}

# 備考・交通機関などに現れる深夜手当のヒント語
_LATE_NIGHT_HINTS = ("深夜", "夜勤", "残業", "宿泊", "前泊", "後泊")


def _has_late_night_allowance(leg: ExpenseLeg) -> bool:
    """レッグに深夜移動を許容するヒント (手当/備考) があるか."""
    # 手当CDのいずれかが付与されていれば日当系の裏付けありとみなす
    if any((leg.allowance_cd_perdiem, leg.allowance_cd_lodging, leg.allowance_cd_stay)):
        return True
    blob = " ".join(
        x for x in (leg.remark, leg.transport, leg.account_name) if x
    )
    return any(h in blob for h in _LATE_NIGHT_HINTS)


def _is_late_night(start: time | None, end: time | None, cfg: Config) -> bool:
    """移動開始が早朝閾値より前, または終了が深夜閾値より後か.

    終了側は『時刻 > HH:00』を意味するよう time 境界で比較する
    (時のみ比較だと 22:30 等の HH:xx を取りこぼすため).
    """
    if start is not None and start < time(cfg.late_night_start_before, 0):
        return True
    if end is not None and end > time(cfg.late_night_end_after, 0):
        return True
    return False


def _movement_legs(r: ExpenseReport) -> list[ExpenseLeg]:
    """労務照合対象 = 実際の移動レッグ (時刻付き).

    時刻を持つだけのレッグ (駐車代/ﾎﾃﾙ等の活動レッグは 00:00 プレースホルダ時刻を
    持つことがある) を全部含めると深夜判定が誤発火するため, 移動系交通機関
    (MOVEMENT_TRANSPORTS / is_movement_leg) に限定する.
    """
    out = []
    for leg in r.legs:
        if leg.leg_date is None:
            continue
        is_move = leg.is_movement_leg or leg.transport in MOVEMENT_TRANSPORTS
        if not is_move:
            continue
        if leg.time_start is None and leg.time_end is None:
            continue
        out.append(leg)
    return out


def check_labor(r: ExpenseReport, att: AttendanceLookup, cfg: Config) -> CheckResult:
    """出張レポートの労務 (勤怠整合) を判定し CheckResult を返す.

    手順 (§6.2):
      1. 時刻を持つレッグごとに勤怠を引き当て寄与ステータスを積む.
         - 当該月の勤怠未提供 → 未確認(勤怠データ欠落).
         - 勤怠日あり かつ 在席 (出社/TW/休出) → OK.
         - 勤怠日あり かつ 非労働日/休暇 (休出実態なし) → 要確認.
         - 勤怠日が無い (氏名はあるが該当日欠落) → 未確認(勤怠データ欠落).
      2. 深夜帯の移動 (cfg 閾値) かつ深夜手当ヒント無し → 要確認 を追加.
      3. その日の最終移動レッグの終了時刻と勤務終了時刻の差が cfg.labor_time_gap_minutes
         を超える → 要確認 を追加 (勤務終了は clock_out, 無ければ move_end/move2_end で
         代用). 往路/中間移動は対象外 (退勤時刻と比べても意味が無いため).
      4. 通常経路到着時間が入力ルール外 → 要確認. 入力有り かつ 退勤時刻と不一致
         → 要確認. 未入力 → 要確認 (未確認のため要確認扱い).
      5. 法定時間外/休日労働の上限は規程未提供 → 未確認(規程未提供) を控えめに付与.
      6. 総合は worst_status. 寄与が無ければ OK (照合対象レッグ無し).
    """
    # 別名解決後のマスタ氏名を優先 (勤怠は社員マスタ名でキーされるため)
    name = r.resolved_name_norm or r.inputter_name_norm
    legs = _movement_legs(r)

    # 日付ごとの最終移動レッグ (その日最も遅い time_end) を特定する.
    # 勤務終了時刻との乖離判定は「帰路」に相当するこのレッグのみに適用する —
    # 往路/中間移動 (午前の移動など) を夜の退勤時刻と比べると常に大差になり
    # 誤って要確認になってしまうため (2026-07-09 客先指摘: OKが一件も出ない).
    last_leg_by_date: dict = {}
    for leg in legs:
        if leg.leg_date is None or leg.time_end is None:
            continue
        cur = last_leg_by_date.get(leg.leg_date)
        if cur is None or leg.time_end > cur.time_end:
            last_leg_by_date[leg.leg_date] = leg

    contributions: list[str] = []
    reasons: list[str] = []
    late_night_legs: list[int] = []
    nonwork_legs: list[int] = []
    att_missing_legs: list[int] = []
    corroborated_legs: list[int] = []
    time_gap_legs: list[int] = []
    normal_arrival_invalid_legs: list[int] = []
    normal_arrival_mismatch_legs: list[int] = []
    normal_arrival_empty_legs: list[int] = []
    months_missing: set[str] = set()

    for leg in legs:
        d = leg.leg_date

        # --- 当該月の勤怠が未提供 (5月ギャップ) ---
        # 勤怠が無い月の移動は照合不能なので ATT_MISSING に留める.
        # 深夜判定 (要確認) はここでは付けない — 裏付けデータが無いのに
        # 「確定した要確認」にすると未確認(勤怠欠落)を誤って上書きしてしまうため.
        if not att.is_month_present(d):
            contributions.append(ATT_MISSING)
            att_missing_legs.append(leg.leg_no)
            months_missing.add(d.strftime("%Y-%m"))
            continue

        # --- 深夜移動判定 (当該月の勤怠が有る場合のみ評価) ---
        if _is_late_night(leg.time_start, leg.time_end, cfg) and not _has_late_night_allowance(leg):
            contributions.append(NEEDS_CHECK)
            late_night_legs.append(leg.leg_no)

        # --- 当該日の勤怠を引き当て ---
        day: AttendanceDay | None = att.get(name, d)
        if day is None:
            # 氏名・該当日の勤怠行が無い → 欠落
            contributions.append(ATT_MISSING)
            att_missing_legs.append(leg.leg_no)
            continue

        presence = day.presence
        if presence in _PRESENT:
            # 在席が移動を裏付け
            contributions.append(OK)
            corroborated_legs.append(leg.leg_no)
        elif presence in _NONWORK:
            # 非労働日/休暇に移動 — 休出実態が無ければ要確認
            contributions.append(NEEDS_CHECK)
            nonwork_legs.append(leg.leg_no)
        else:
            # 在席状態が不明 → 裏付け不可で欠落扱い
            contributions.append(ATT_MISSING)
            att_missing_legs.append(leg.leg_no)

        # --- 移動終了時刻と勤務終了時刻の乖離判定 (その日の最終移動のみ) ---
        if leg.time_end is not None and last_leg_by_date.get(d) is leg:
            _, work_end = effective_work_times(day)
            if work_end is not None:
                diff_min = abs(
                    (leg.time_end.hour * 60 + leg.time_end.minute)
                    - (work_end.hour * 60 + work_end.minute)
                )
                if diff_min > cfg.labor_time_gap_minutes:
                    contributions.append(NEEDS_CHECK)
                    time_gap_legs.append(leg.leg_no)

        # --- 通常経路到着時間と退勤時刻の一致判定 ---
        # 未入力 (empty) も要確認とする (2026-07-10 客先指定: 未入力は確認が必要).
        if leg.normal_route_arrival_status == "invalid":
            contributions.append(NEEDS_CHECK)
            normal_arrival_invalid_legs.append(leg.leg_no)
        elif leg.normal_route_arrival_status == "empty":
            contributions.append(NEEDS_CHECK)
            normal_arrival_empty_legs.append(leg.leg_no)
        elif leg.normal_route_arrival is not None and day.clock_out is not None:
            if (leg.normal_route_arrival.hour, leg.normal_route_arrival.minute) != (
                day.clock_out.hour, day.clock_out.minute
            ):
                contributions.append(NEEDS_CHECK)
                normal_arrival_mismatch_legs.append(leg.leg_no)

    # --- 法定時間外/休日労働の上限判定: 就業規則・規程未提供 ---
    # 深夜/休日移動の事実はあるが, 上限超過の是非は規程が無いと判定不能.
    if (late_night_legs or nonwork_legs):
        contributions.append(RULE_MISSING)

    overall = worst_status(contributions)

    # --- 理由文 (人間可読) の組立 ---
    if nonwork_legs:
        reasons.append(f"非労働日/休暇に移動 (明細{','.join(map(str, nonwork_legs))})")
    if late_night_legs:
        reasons.append(f"深夜移動 要確認 (明細{','.join(map(str, late_night_legs))})")
    if time_gap_legs:
        reasons.append(
            f"移動終了と勤務終了の差が{cfg.labor_time_gap_minutes}分超 (明細{','.join(map(str, time_gap_legs))})"
        )
    if normal_arrival_invalid_legs:
        reasons.append(
            f"通常経路到着時間が入力ルール外 (明細{','.join(map(str, normal_arrival_invalid_legs))})"
        )
    if normal_arrival_mismatch_legs:
        reasons.append(
            f"通常経路到着時間が退勤時刻と不一致 (明細{','.join(map(str, normal_arrival_mismatch_legs))})"
        )
    if normal_arrival_empty_legs:
        reasons.append(
            f"通常経路到着時間が未入力 (明細{','.join(map(str, normal_arrival_empty_legs))})"
        )
    if months_missing:
        reasons.append(f"勤怠データ欠落: {','.join(sorted(months_missing))} 月分未提供")
    elif att_missing_legs:
        reasons.append(f"勤怠日欠落 (明細{','.join(map(str, att_missing_legs))})")
    if not legs:
        reasons.append("時刻付き移動レッグなし — 労務照合対象外")
    elif overall == OK and corroborated_legs:
        reasons.append("全移動レッグが勤怠 (在席) と整合")

    detail = "; ".join(reasons) if reasons else "問題なし"

    evidence = {
        "対象レッグ数": len(legs),
        "整合レッグ": corroborated_legs,
        "非労働日レッグ": nonwork_legs,
        "深夜レッグ": late_night_legs,
        "時刻乖離レッグ": time_gap_legs,
        "通常経路到着不一致レッグ": normal_arrival_mismatch_legs,
        "通常経路到着入力エラーレッグ": normal_arrival_invalid_legs,
        "通常経路到着未入力レッグ": normal_arrival_empty_legs,
        "勤怠欠落レッグ": att_missing_legs,
        "欠落月": sorted(months_missing),
        "氏名突合": att.has_name(name),
    }

    # --- 差戻し候補 (OK 以外のとき) ---
    suggestion = ""
    if overall != OK:
        if nonwork_legs:
            suggestion = "非労働日/休暇日の移動について出張の必要性・実態をご確認ください。"
        elif late_night_legs:
            suggestion = "深夜帯の移動について時刻・深夜手当の要否をご確認ください。"
        elif time_gap_legs:
            suggestion = "移動終了時刻と勤務終了時刻の差が大きいため、実態をご確認ください。"
        elif normal_arrival_invalid_legs:
            suggestion = "通常経路到着時間の入力形式をご確認ください。"
        elif normal_arrival_mismatch_legs:
            suggestion = "通常経路到着時間と退勤時刻が一致しないため、実態をご確認ください。"
        elif normal_arrival_empty_legs:
            suggestion = "通常経路到着時間が未入力のため、実態をご確認ください。"
        elif months_missing or att_missing_legs:
            suggestion = "該当月の勤怠データが未提供のため労務照合は劣化しています。勤怠提供後に再確認してください。"
        else:
            suggestion = "労務 (勤怠整合) について手動でご確認ください。"

    return CheckResult(
        status=overall,
        detail=detail,
        evidence=evidence,
        suggestion=suggestion,
    )
