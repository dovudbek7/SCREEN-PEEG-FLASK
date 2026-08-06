"""チェックシート組立 (enrich + 6 判定 → spec §6 の 7 シート分の行 dict).

入力モデル群を受け取り, 従業員/地名照合で enrich し, 6 観点を判定し,
excel_writer.write_excel が期待する dict を返す.

出力 dict キー → spec §6 シート:
  primary       → 01_一次承認チェック (申請(伝票)1件1行: 総合判定/要確認項目.
                    2026-07-17客先依頼: '3.定時外の移動時間 勤務実態' を
                    '2.労務・健康管理の確認' の直後に列として組み込み済み — 独立シートではない)
  secondary     → 02_二次承認詳細     (明細(レッグ)単位: 金額/領収書/宿泊/日当/日程/距離)
  diff          → 03_差異一覧         (OK以外の観点のみ: 理由/確認先システム/対応案)
  reject        → 04_差戻し文面候補   (要確認/NG の観点ごとに差戻し文面)
  import_log    → 05_取込ログ         (取込件数/未承認件数照合/エラー)
  rules         → 06_判定ルール       (使用閾値/区分/規程提供状況)
  master_check  → 07_マスタ確認       (未登録/重複/不整合)
  banners       → 既知の前提・データ欠落バナー
"""
from __future__ import annotations

from collections import Counter
from datetime import date, datetime, time, timedelta

from models import (
    ExpenseReport, AttendanceLookup,
    MOVEMENT_TRANSPORTS,
    OK, NEEDS_CHECK, NG, UNMATCHED, MULTI, RULE_MISSING, ATT_MISSING,
    worst_status, to_axis_vocab, effective_work_times,
)
from loaders.allowance_master_loader import (load_allowance_master,
                                             load_allowance_per_day_codes)
from matching.employee_match import build_employee_index, resolve_employee
from matching.place_match import build_customer_index, match_place
from matching.approver_match import build_approver_index
from normalize import norm
from rules.trip_reality import check_trip_reality
from rules.labor import check_labor
from rules.amount import check_amount
from rules.duplicate import check_duplicate
from rules.receipt import check_receipt
from rules.approval_route import check_approval_route


def _d(d: date | None) -> str:
    return d.isoformat() if d else ""


# 観点 → 使用データ (§6.1 「確認先システムと確認観点を表示」)
# 2026-08-04 客先指摘: 1観点につき複数のデータを突合しているため, 単一の
# 「確認先システム」ではなく実際に使用しているデータをすべて列挙する.
_AXIS_SYSTEM = {
    "出張実態": "楽々精算 + 楽々勤怠 + 顧客リスト",
    "労務": "楽々精算 + 楽々勤怠",
    "金額規程": "楽々精算 + 旅費規定 + 社員リスト(役職) + 顧客リスト(都道府県)",
    "二重申請": "楽々精算 (全伝票横断) + 顧客リスト",
    "領収書": "楽々精算 (証票) + 旅費規定",
    "承認ルート": "楽々精算 + 20期承認者名簿 + 社員リスト",
}


def _grade_from_role(role: str | None) -> str | None:
    """出勤簿の役割区分 → 役職区分 (管理職/一般職) に変換."""
    if not role:
        return None
    if role == "社員":
        return "一般職"
    if "承認者" in role:
        return "管理職"
    return None


def _build_grade_lookup(attendance_days) -> dict[str, str]:
    """name_norm → grade の辞書を出勤簿 role から構築."""
    lookup: dict[str, str] = {}
    for d in attendance_days:
        if d.name_norm and d.role and d.name_norm not in lookup:
            grade = _grade_from_role(d.role)
            if grade:
                lookup[d.name_norm] = grade
    return lookup


def enrich_reports(reports, employees, customers, cfg,
                   grade_lookup: dict | None = None) -> None:
    """各レポートに従業員ID, 各レッグに照合顧客/距離を付与 (in place)."""
    eidx = build_employee_index(employees)
    cidx = build_customer_index(customers)
    emp_by_id = {e.employee_id: e for e in employees}
    role_override_lookup = {norm(k): v for k, v in (cfg.role_overrides or {}).items()}
    for r in reports:
        # ① 申請者CD による直接 ID ルックアップ (新CSV のみ)
        if r.applicant_cd and r.applicant_cd in emp_by_id:
            e = emp_by_id[r.applicant_cd]
            r.employee_id = e.employee_id
            r.employee_match_status = "突合"
            r.department = e.department
            r.email = e.email
            r.resolved_name_norm = e.name_norm
            r.grade = e.grade
        else:
            # ② 氏名ファジーマッチ (旧CSV / 申請者CD 未登録フォールバック)
            res = resolve_employee(r.inputter_name_raw, eidx, cfg)
            r.employee_id = res.employee_id
            r.employee_match_status = res.status
            if res.employee_id and res.employee_id in emp_by_id:
                e = emp_by_id[res.employee_id]
                r.department = e.department
                r.email = e.email
                r.resolved_name_norm = e.name_norm
                r.grade = e.grade
        # ③ 出勤簿 role からフォールバック
        if r.grade is None and grade_lookup:
            name_key = r.resolved_name_norm or r.inputter_name_norm
            r.grade = grade_lookup.get(name_key)
        # ④ config.role_overrides からフォールバック
        if r.grade is None and role_override_lookup:
            name_key = r.resolved_name_norm or r.inputter_name_norm
            r.grade = role_override_lookup.get(name_key)
        for leg in r.legs:
            # 車 + 賃借料 計上は 旅費交通費 へ統一 (2026-07-01 客先回答:
            # レンタカー・社用車とも区分不要, 旅費交通費に修正して問題ない).
            if leg.transport == "車" and leg.account_name and "賃借料" in leg.account_name:
                leg.account_name = leg.account_name.replace("賃借料", "旅費交通費")
            # 移動レッグ(電車･ﾊﾞｽ/車等)の到着地は経由地であり訪問先(顧客)ではない.
            # 地名→無関係社名の誤突合を避けるため顧客照合をスキップする.
            if leg.is_movement_leg or leg.transport in MOVEMENT_TRANSPORTS:
                leg.dest_match_status = UNMATCHED
                continue
            target = leg.dest_raw or leg.origin_raw
            pm = match_place(target, cidx, cfg)
            leg.dest_customer_no = pm.customer_no
            leg.dest_customer_name = pm.customer_name
            leg.dest_match_score = pm.score
            leg.dest_match_status = pm.status
            leg.dest_distance_band = pm.distance_band
            leg.dest_km_lower = pm.km_lower
            leg.dest_km_upper = pm.km_upper
            leg.dest_candidates = pm.candidates
            leg.dest_prefecture = pm.prefecture
    return eidx, cidx


def _allowance_flags(r: ExpenseReport) -> tuple[bool, bool, bool]:
    """伝票内に 日当/宿泊/滞在 手当コードを持つレッグがあるか."""
    per = any(leg.allowance_cd_perdiem for leg in r.legs)
    lodge = any(leg.allowance_cd_lodging for leg in r.legs)
    stay = any(leg.allowance_cd_stay for leg in r.legs)
    return per, lodge, stay


def _matched_bands(r: ExpenseReport) -> str:
    """突合できた訪問先レッグの距離区分一覧 (カンマ連結)."""
    bands = []
    for leg in r.legs:
        if leg.dest_match_status in ("突合", "別名突合") and leg.dest_distance_band:
            if leg.dest_distance_band not in bands:
                bands.append(leg.dest_distance_band)
    return ", ".join(bands)


def _fmt_hm(t) -> str:
    """time/datetime を HH:MM 表記に. None なら空文字."""
    return t.strftime("%H:%M") if t is not None else ""


def compute_allowances(r: ExpenseReport, perdiem_master: dict | None = None,
                       per_day_codes: set | None = None) -> dict:
    """伝票の 日当 / 滞在費補助 の実額と, 1日・1泊あたりの単価を返す.

    明細の金額欄には手当額そのものは入っていない (手当CDが付く明細の金額は
    交通費などの実費). 実際の支給額はヘッダの手当計にまとまっているため,
    日当は手当マスタの単価から積み上げ, 滞在費補助はその残差として求める。

    2026-08-06 客先指摘:「日当は1日で1700円、2日で3400円だが反映されていない」
    手当マスタの計算式入力フラグが1のコード (例 002「日当(連続)」) は
    手当金額が1日あたりの単価なので, 出張日数を掛ける。日数は明細日付の
    最初〜最後 (中日に明細が無くても数える) とする。

    返り値:
      perdiem      日当の合計額
      stay         滞在費補助の合計額 (手当計 − 日当)
      trip_days    出張日数 (1以上)
      nights       宿泊数 (0以上)
      perdiem_unit 1日あたりの日当 (上限判定用. 不明なら None)
      stay_unit    1泊あたりの滞在費補助 (上限判定用. 泊が無ければ None)
    """
    has_perdiem = any(leg.allowance_cd_perdiem for leg in r.legs)
    has_stay = any(leg.allowance_cd_stay for leg in r.legs)

    trip_days = 1
    nights = 0
    if r.date_min and r.date_max:
        trip_days = (r.date_max - r.date_min).days + 1
        nights = (r.date_max - r.date_min).days

    perdiem_from_master = None
    if perdiem_master and has_perdiem:
        per_day = per_day_codes or set()
        amounts = []
        for leg in r.legs:
            cd = leg.allowance_cd_perdiem
            if not cd:
                continue
            unit = perdiem_master.get(cd)
            if unit is None:
                amounts = None
                break
            amounts.append(unit * trip_days if cd in per_day else unit)
        if amounts:
            perdiem_from_master = sum(amounts)

    if perdiem_from_master is not None:
        perdiem = perdiem_from_master
        stay = (r.allowance_total_declared - perdiem_from_master) if has_stay else 0
    elif has_perdiem and not has_stay:
        # マスタが使えない場合のフォールバック: 単一区分ならヘッダ手当計をそのまま帰属
        # (2026-07-10 客先指摘: 日当計1,700に対し日当金額が0のままだった).
        perdiem = r.allowance_total_declared
        stay = 0
    elif has_stay and not has_perdiem:
        perdiem = 0
        stay = r.allowance_total_declared
    else:
        perdiem = sum(leg.amount for leg in r.legs if leg.allowance_cd_perdiem)
        stay = sum(leg.amount for leg in r.legs if leg.allowance_cd_stay)

    return {
        "perdiem": perdiem,
        "stay": stay,
        "trip_days": trip_days,
        "nights": nights,
        "perdiem_unit": (perdiem // trip_days) if (perdiem and trip_days) else None,
        "stay_unit": (stay // nights) if (stay and nights) else None,
    }


def _trip_detail_summary(r: ExpenseReport, att: AttendanceLookup,
                         perdiem_master: dict | None = None,
                         per_day_codes: set | None = None) -> dict:
    """01シート追加列 (§画面ビューワー拡張) 用の集計.

    出張先/取引先は移動レッグの照合結果から集約 (複数該当時は先頭を採用).
    取引先が1件も確定しない場合は 複数候補の先頭を表示用に採用する.
    移動開始/終了時間は時刻付き移動レッグの最早/最遅. 勤務開始/終了時間は
    移動開始日の勤怠(出勤簿)実績 (labor.py と同じ氏名キーで引き当て).
    整合(時間差)は移動終了時刻と勤怠退勤時刻の差 (分).
    通常経路到着時間は移動レッグの入力値をそのまま表示 (入力ルール外は
    「入力ルール外」と表示, 未入力は空欄). 退勤時刻との一致判定は labor.py 側で行う.
    """
    prefectures = [leg.dest_prefecture for leg in r.legs if leg.dest_prefecture]
    customers = [leg.dest_customer_name for leg in r.legs if leg.dest_customer_name]
    # 2026-07-30 客先依頼: 複数候補で取引先が確定できず空欄になる場合,
    # 候補一覧の先頭 (顧客マスタの並び順) を 01シートの取引先に表示する.
    # 距離区分は割れたままのため 02シートの照合状態/距離区分は 複数候補 を維持し,
    # ここでの採用は表示目的に限定する (金額判定には使わない).
    if not customers:
        customers = [
            leg.dest_candidates[0][0] for leg in r.legs
            if leg.dest_candidates and leg.dest_candidates[0][0]
        ]

    move_legs = [
        leg for leg in r.legs
        if (leg.is_movement_leg or leg.transport in MOVEMENT_TRANSPORTS)
        and (leg.time_start is not None or leg.time_end is not None)
    ]
    move_start = min((leg.time_start for leg in move_legs if leg.time_start), default=None)
    move_end = max((leg.time_end for leg in move_legs if leg.time_end), default=None)
    move_date = next((leg.leg_date for leg in move_legs if leg.leg_date), None)

    normal_arrival_invalid = any(
        leg.normal_route_arrival_status == "invalid" for leg in move_legs
    )
    normal_arrival = next(
        (leg.normal_route_arrival for leg in move_legs if leg.normal_route_arrival), None
    )

    name = r.resolved_name_norm or r.inputter_name_norm
    day = att.get(name, move_date) if move_date else None
    work_start, work_end = effective_work_times(day)

    # 整合(時間差) は 通常経路到着時間 と 勤務終了時間(退勤) の差とする
    # (2026-07-09 客先依頼: 移動終了時間との差から変更).
    diff_label = ""
    if normal_arrival is not None and work_end is not None:
        diff_min = abs((normal_arrival.hour * 60 + normal_arrival.minute)
                        - (work_end.hour * 60 + work_end.minute))
        diff_label = f"{diff_min}分"

    alw = compute_allowances(r, perdiem_master, per_day_codes)
    perdiem_amount = alw["perdiem"]
    stay_amount = alw["stay"]
    # 宿泊費(ホテル代)は手当2CD(allowance_cd_lodging)が実データで常に空のため使えず,
    # 宿泊泊数と同じ判定軸 (transport=ホテル) の明細金額合算で代用する.
    # 宿泊税・入湯税もホテル代に付随する費用のため合算に含める (2026-07-07 客先確認:
    # 分離集計だと明細合計(r.total_amount)との差額として金額不一致に見えていた).
    lodging_amount = sum(
        leg.amount for leg in r.legs
        if leg.transport in ("ﾎﾃﾙ", "ホテル", "宿泊税", "入湯税")
    )
    # 日当計は明細金額の合算ではなく, 手当計算システム側の実額 (ヘッダ手当計) を採用する.
    # 明細の金額欄が0のまま手当CDだけ立っているケースがあり (2026-07-08 客先指摘:
    # 日当計が元データ1,700に対しチェックシート上0), 明細合算だと取りこぼすため.
    perdiem_total = r.allowance_total_declared
    # 宿泊実態は手当CD(手当2/3CD の記帳ゆれ)ではなく実際のホテル計上日で数える.
    nights = len({
        leg.leg_date for leg in r.legs
        if leg.transport in ("ﾎﾃﾙ", "ホテル") and leg.leg_date
    })

    return {
        "出張先": prefectures[0] if prefectures else "",
        "取引先": customers[0] if customers else "",
        "移動開始時間": _fmt_hm(move_start),
        "通常経路到着時間": "入力ルール外" if normal_arrival_invalid else _fmt_hm(normal_arrival),
        "移動終了時間": _fmt_hm(move_end),
        "勤務開始時間": _fmt_hm(work_start),
        "勤務終了時間": _fmt_hm(work_end),
        "整合(時間差)": diff_label,
        "日当金額": perdiem_amount,
        "滞在費補助金額": stay_amount,
        "宿泊費（ホテル代）": lodging_amount,
        "日当計": perdiem_total,
        "宿泊泊数": nights,
    }


def _window_span(legs: list[ExpenseLeg]) -> tuple[time | None, time | None]:
    """レッグ群の最早開始・最遅終了 (時刻付きレッグの time_start/time_end の min/max)."""
    starts = [lg.time_start for lg in legs if lg.time_start is not None]
    ends = [lg.time_end for lg in legs if lg.time_end is not None]
    return (min(starts) if starts else None, max(ends) if ends else None)


_HOLIDAY_CALENDAR_TYPES = ("法定外", "法定内")  # 出勤簿カレンダー種別 (土曜/日曜等の休日)


def _split_by_worktime_boundary(day_legs: list[ExpenseLeg], work_start: time, work_end: time):
    """楽々精算側 (平日): 当日のレッグを 定時前/勤務時間/定時後 の3窓に分類する.

    移動レッグは 定時前(work_start より前の部分)/定時後(work_end 以降の部分) に
    属する区間だけを計上する — 境界を跨ぐレッグ (例: 08:40開始〜17:30終了) は,
    それぞれの窓の境界時刻でクリップする (客先指摘 2026-07-24: 開始時刻だけで
    判定し終了時刻をクリップしていなかったため, 勤務時間帯やその先まで範囲が
    はみ出して表示されていた不具合の修正). 勤務時間帯にちょうど収まる移動
    (境界時刻で終わる/始まるものを含む) はいずれの窓にも計上しない.
    作業･打合せ レッグは勤務時間群.
    """
    before_starts, before_ends = [], []
    after_starts, after_ends = [], []
    work = []
    for lg in day_legs:
        if lg.time_start is None and lg.time_end is None:
            continue
        is_move = lg.is_movement_leg or lg.transport in MOVEMENT_TRANSPORTS
        if is_move:
            s, e = lg.time_start, lg.time_end
            if s is not None and s < work_start:
                before_starts.append(s)
                before_ends.append(e if (e is not None and e < work_start) else work_start)
            elif (e is not None and e > work_end) or (e is None and s is not None and s >= work_end):
                after_starts.append(s if (s is not None and s > work_end) else work_end)
                after_ends.append(e if e is not None else s)
        elif lg.transport == "作業･打合せ":
            work.append(lg)
    before_span = (min(before_starts) if before_starts else None, max(before_ends) if before_ends else None)
    after_span = (min(after_starts) if after_starts else None, max(after_ends) if after_ends else None)
    return before_span, work, after_span


def _split_by_work_leg_anchor(day_legs: list[ExpenseLeg]):
    """[現在未使用 / 2026-07-27〜] 休日の別扱いを客先が再要望した場合に備えて残置.
    2026-07-27 の依頼で 平日・休日とも 9:00/17:30 境界に一本化したため, 現状は
    _build_off_hours_summary_for_report から呼ばれない (_is_holiday も同様).

    楽々精算側 (休日): 定時の概念が無いため, 作業･打合せレッグを基準に前後の移動を
    「顧客先への移動」「顧客先からの移動」としてそのまま反映する (客先確定ルール 2026-07-16).
    作業･打合せレッグが無い日は, 移動レッグをすべて「顧客先への移動」枠にそのまま反映する
    (行き/帰りを判別する基準が無いため).
    """
    move_legs = [
        lg for lg in day_legs
        if (lg.is_movement_leg or lg.transport in MOVEMENT_TRANSPORTS)
        and (lg.time_start is not None or lg.time_end is not None)
    ]
    work_legs = [lg for lg in day_legs if lg.transport == "作業･打合せ"]
    if not work_legs:
        return move_legs, [], []
    work_start = min((lg.time_start for lg in work_legs if lg.time_start is not None), default=None)
    work_end = max((lg.time_end for lg in work_legs if lg.time_end is not None), default=None)
    before, after = [], []
    for lg in move_legs:
        t = lg.time_end or lg.time_start
        if t is None:
            continue
        if work_start is not None and t <= work_start:
            before.append(lg)
        elif work_end is not None and t >= work_end:
            after.append(lg)
    return before, work_legs, after


def _classify_attendance_by_window(day, work_start: time, work_end: time):
    """楽々勤怠側 (平日): 移動開始/終了・移動2開始/終了 を 定時前/定時後 に分類する.

    移動レッグ側 (_split_by_worktime_boundary) と同じく, 境界を跨ぐ移動は
    それぞれの窓の境界時刻でクリップする (客先指摘 2026-07-24: 勤務時間帯に
    ちょうど収まる移動 (例: 15:45〜17:30) まで定時前/定時後として計上して
    いた不具合の修正). 勤務時間帯にちょうど収まる/境界時刻で終わる・始まる
    移動はいずれの窓にも計上しない.

    勤務時間側は clock_in/clock_out を呼び出し元 (_build_off_hours_summary_for_report) が直接使う —
    effective_work_times() の move_start/move_end フォールバックを介さない. 介すと
    打刻が無い日に 定時前/定時後 と同じ値を勤務時間列にも二重表示してしまうため.
    """
    if day is None:
        return (None, None), (None, None)
    before_starts, before_ends, after_starts, after_ends = [], [], [], []
    for s, e in ((day.move_start, day.move_end), (day.move2_start, day.move2_end)):
        if s is None:
            continue
        if s < work_start:
            before_starts.append(s)
            before_ends.append(e if (e is not None and e < work_start) else work_start)
        elif (e is not None and e > work_end) or (e is None and s >= work_end):
            after_starts.append(s if s > work_end else work_end)
            after_ends.append(e if e is not None else s)
    before = (min(before_starts) if before_starts else None, max(before_ends) if before_ends else None)
    after = (min(after_starts) if after_starts else None, max(after_ends) if after_ends else None)
    return before, after


# 片方のシステムにしかデータが無く, 突合そのものができない状態 (2026-08-06).
# 出勤簿の移動列は全体の約12%しか入力されておらず, 申請の誤りではなく
# 勤怠側の入力漏れであることが多い. 乖離の証拠ではないため 要確認 (黄) とは分け,
# 「未確認」で始めて Excel/HTML とも灰色表示にし, 総合判定にも反映しない.
# どちら側が欠けているかは差分列 (勤怠打刻なし / 精算計上なし / 作業計上なし) で示す.
UNVERIFIABLE = "未確認(突合不可)"


def _work_window_verdict(exp_pair: tuple, att_pair: tuple, ok_minutes: int) -> tuple[str, str]:
    """勤務時間の窓は「作業･打合せが勤怠の内側に収まっているか」で判定する.

    2026-08-06 客先指摘「要確認が本当にOKでないのか確認したい」への対応。
    従来は 作業･打合せ の開始/終了を 出退勤打刻 と直接比較していたが,
    作業は勤務時間の一部でしかないため乖離が出て当然だった
    (例: 09:00-15:00 に客先で作業し, その後移動して 17:30 退勤 → 2:30差).
    7月データでは 151件中 144件がこの誤検知だった。

    正しい問いは「申告した作業時間が勤怠の勤務時間に収まっているか」なので,
    包含関係で判定する。はみ出した場合のみ要確認とする。
    """
    es, ee = exp_pair
    as_, ae = att_pair
    if es is None and ee is None and as_ is None and ae is None:
        return "移動なし", OK
    if as_ is None and ae is None:
        return "勤怠打刻なし", UNVERIFIABLE
    if es is None and ee is None:
        return "作業計上なし", UNVERIFIABLE
    if es is None or ee is None or as_ is None or ae is None:
        return "データなし", UNVERIFIABLE

    def _m(t) -> int:
        return t.hour * 60 + t.minute

    # はみ出しの許容は他の窓と同じ 10分 (cfg.off_hours_diff_ok_minutes) に合わせる
    over_before = max(0, _m(as_) - _m(es))   # 出勤前に作業していた分
    over_after = max(0, _m(ee) - _m(ae))     # 退勤後に作業していた分
    if max(over_before, over_after) <= ok_minutes:
        return "勤務時間内", OK
    parts = []
    if over_before:
        parts.append(f"出勤前{over_before // 60}:{over_before % 60:02d}")
    if over_after:
        parts.append(f"退勤後{over_after // 60}:{over_after % 60:02d}")
    return "／".join(parts), NEEDS_CHECK


def _is_holiday(day) -> bool:
    """[現在未使用 / 2026-07-27〜] _split_by_work_leg_anchor と対で残置 (下記参照).
    出勤簿カレンダー種別から休日判定 (平日以外=法定外/法定内). 勤怠データが無い日は
    判定材料が無いため平日ロジック (定時9:00-17:30) にフォールバックする."""
    return day is not None and day.calendar_type in _HOLIDAY_CALENDAR_TYPES


def _off_hours_window_verdict(exp_pair: tuple, att_pair: tuple, ok_minutes: int) -> tuple[str, str]:
    """窓(定時前/勤務/定時後)1つ分の (差分ラベル, チェックverdict) を返す.

    開始・終了の両方を比較し, 悪い方 (差が大きい方) を採用する (客先指摘
    2026-07-24: 終了時刻だけ一致していれば開始時刻が数時間ずれていても
    OK扱いになっていた不具合の修正. 勤務チェック/定時後チェックも同一ロジック).
    客先確定ルール (2026-07-16): 差分が ok_minutes(既定10分) 以内なら OK,
    それを超えたら要確認 (NG無し).

    2026-08-06 客先依頼: 3日間の出張の中日のように, そもそも移動が発生しない日は
    定時前/定時後の欄が両システムとも空になる. これは乖離ではないため
    「移動なし」= OK とする. 一方だけにデータがある場合は突合できない乖離なので
    従来どおり「データなし」= 要確認 のままにする
    (データ欠落は確認対象であり, OK/要確認/NGとは別の第4状態ではない —
    客先モックアップ 2026-07-17 で確認: データなしの行のチェック欄も黄色い「要確認」)."""
    def _diff_min(a: time | None, b: time | None) -> int | None:
        if a is None or b is None:
            return None
        return abs((a.hour * 60 + a.minute) - (b.hour * 60 + b.minute))

    exp_empty = exp_pair[0] is None and exp_pair[1] is None
    att_empty = att_pair[0] is None and att_pair[1] is None
    if exp_empty and att_empty:
        return "移動なし", OK

    # 2026-08-06: 楽々勤怠に移動の打刻が無いケース. 出勤簿の移動列は
    # 全体の約12%しか入力されておらず, 申請の誤りではなく勤怠側の
    # 入力漏れであるため, 要確認ではなく「未確認」として区別する
    # (総合判定にも反映しない — 突合材料が無いだけで乖離の証拠ではない).
    if att_empty:
        return "勤怠打刻なし", UNVERIFIABLE
    if exp_empty:
        return "精算計上なし", UNVERIFIABLE

    d_start = _diff_min(exp_pair[0], att_pair[0])
    d_end = _diff_min(exp_pair[1], att_pair[1])
    if d_start is None or d_end is None:
        # 片側の開始または終了だけが欠けている場合も突合できない
        return "データなし", UNVERIFIABLE
    worse = max(d_start, d_end)
    if worse == 0:
        return "一致", OK
    label = f"{worse // 60}:{worse % 60:02d}差"
    return label, (OK if worse <= ok_minutes else NEEDS_CHECK)


def _build_off_hours_rows_for_report(r: ExpenseReport, att: AttendanceLookup, cfg) -> list[dict]:
    """伝票1件分の「定時外の移動時間 勤務実態」を <b>日単位</b>で返す (01シート '2.' グループ用).

    2026-07-31 客先依頼: 従来は伝票1件=1行だったため, 2泊3日の場合に
    「1日目の定時前」と「3日目の定時後」しか表示されず, 中日の移動が
    見えなくなっていた. 日ごとに1行を返すよう変更する.

    返り値は 明細日付 の昇順. 各要素に "日付" キー (ISO文字列) を含む.
    レッグに日付が1つも無い伝票は, 日付空欄の1行を返す (行が消えないようにするため).
    """
    work_start = cfg.off_hours_work_start
    work_end = cfg.off_hours_work_end
    ok_min = cfg.off_hours_diff_ok_minutes
    name = r.resolved_name_norm or r.inputter_name_norm

    legs_by_date: dict[date, list[ExpenseLeg]] = {}
    for lg in r.legs:
        if lg.leg_date is not None:
            legs_by_date.setdefault(lg.leg_date, []).append(lg)

    if not legs_by_date:
        return [_off_hours_row(None, (None, None), (None, None), (None, None),
                               (None, None), (None, None), (None, None), ok_min)]

    rows = []
    for d in sorted(legs_by_date):
        day_legs = legs_by_date[d]
        att_day = att.get(name, d)

        # 2026-07-27 客先依頼: 平日・休日を問わず 9:00/17:30 の定時境界でクリップする
        # (定時前=9:00まで / 勤務時間=9:00〜17:30 / 定時後=17:30以降 の3ルールを全日に統一).
        # 従来は休日 (法定内/法定外) のみ _split_by_work_leg_anchor で作業レッグを基準に
        # 前後判定していたが, 客先の時刻境界ルールに一本化した. 休日ロジック
        # (_is_holiday/_split_by_work_leg_anchor) は現在未使用だが, 客先が休日の
        # 別扱いを再要望した場合に備えて残置する (2026-07-16 確定ルール).
        exp_before, work_legs, exp_after = _split_by_worktime_boundary(
            day_legs, work_start, work_end)
        exp_work = _window_span(work_legs)

        if att_day is None:
            att_before = att_after = (None, None)
            att_work = (None, None)
        else:
            att_before, att_after = _classify_attendance_by_window(att_day, work_start, work_end)
            att_work = (
                att_day.clock_in.time() if att_day.clock_in is not None else None,
                att_day.clock_out.time() if att_day.clock_out is not None else None,
            )

        rows.append(_off_hours_row(d, exp_before, exp_work, exp_after,
                                   att_before, att_work, att_after, ok_min))
    return rows


def _off_hours_row(d: date | None, exp_before, exp_work, exp_after,
                   att_before, att_work, att_after, ok_min: int) -> dict:
    """1日分の 定時外の移動時間 行 dict を組み立てる."""
    att_work_start, att_work_end = att_work
    before_diff, before_chk = _off_hours_window_verdict(exp_before, att_before, ok_min)
    work_diff, work_chk = _work_window_verdict(exp_work, (att_work_start, att_work_end), ok_min)
    after_diff, after_chk = _off_hours_window_verdict(exp_after, att_after, ok_min)

    return {
        "日付": _d(d),
        "定時前_精算開始": _fmt_hm(exp_before[0]), "定時前_精算終了": _fmt_hm(exp_before[1]),
        "定時前_勤怠開始": _fmt_hm(att_before[0]), "定時前_勤怠終了": _fmt_hm(att_before[1]),
        "勤務_精算開始": _fmt_hm(exp_work[0]), "勤務_精算終了": _fmt_hm(exp_work[1]),
        "勤務_勤怠開始": _fmt_hm(att_work_start), "勤務_勤怠終了": _fmt_hm(att_work_end),
        "定時後_精算開始": _fmt_hm(exp_after[0]), "定時後_精算終了": _fmt_hm(exp_after[1]),
        "定時後_勤怠開始": _fmt_hm(att_after[0]), "定時後_勤怠終了": _fmt_hm(att_after[1]),
        "定時前差分": before_diff, "定時前チェック": before_chk,
        "勤務差分": work_diff, "勤務チェック": work_chk,
        "定時後差分": after_diff, "定時後チェック": after_chk,
    }


# 日単位に展開したときの列の扱い (2026-07-31 客先指摘: 全部空欄だとソートで
# 行がバラバラになり, どの伝票の行か分からなくなる).
#   REPEAT    … 識別用の列。2日目以降にも同じ値を繰り返す (ソートしても行が自分を名乗る)
#   FIRST_ONLY… 金額・判定など伝票単位の値。先頭日の行にのみ入れる
#               (繰り返すと合計金額を集計したとき日数分だけ多重計上されるため)
_PRIMARY_REPEAT_KEYS = (
    "No.", "伝票No.", "入力者名", "社員番号", "所属", "承認状態",
    "出張期間", "出張先", "取引先",
)
_PRIMARY_FIRST_ONLY_KEYS = (
    "出張実態",
    "日当金額", "滞在費補助金額", "宿泊費（ホテル代）", "日当計", "宿泊費・手当の整合",
    "合計金額", "金額確認", "二重申請確認", "証跡・領収書確認",
    "総合判定", "要確認項目",
)

_OFF_HOURS_WINDOWS = (
    ("定時前", "定時前差分", "定時前チェック"),
    ("勤務", "勤務差分", "勤務チェック"),
    ("定時後", "定時後差分", "定時後チェック"),
)


def _off_hours_axis_status(off_hours_rows: list[dict]) -> str:
    """定時外の移動時間ブロックを 総合判定へ反映するための1ステータスに畳む.

    2026-07-30 客先指摘: 勤務チェックが要確認なのに総合判定がOKになっていた
    (このブロックは表示専用で 判定6観点 に含まれていなかったため).

    差分が「データなし」の窓は判定材料が無いだけで乖離の証拠ではないため
    総合判定には反映しない (表示上は 2026-07-17 客先確認どおり黄色の要確認のまま).
    実際に時刻差が閾値を超えた窓のみを 要確認 として反映する.
    2026-07-31 の日単位化以降は全日分の worst を採る.
    """
    statuses = []
    for row in off_hours_rows:
        for _, diff_key, chk_key in _OFF_HOURS_WINDOWS:
            chk = row.get(chk_key, OK)
            # 突合材料が無いだけの状態は乖離の証拠ではないので総合判定に含めない
            if row.get(diff_key) == "データなし" or str(chk).startswith("未確認"):
                continue
            statuses.append(chk)
    return worst_status(statuses)


def _off_hours_reasons(off_hours_rows: list[dict]) -> list[str]:
    """03_差異一覧 用: 閾値超過した日と窓を「6/8 勤務: 1:30差」形式で列挙."""
    reasons = []
    for row in off_hours_rows:
        day = row.get("日付") or ""
        for label, diff_key, chk_key in _OFF_HOURS_WINDOWS:
            chk = row.get(chk_key, OK)
            if row.get(diff_key) == "データなし" or chk == OK or str(chk).startswith("未確認"):
                continue
            prefix = f"{day} " if day else ""
            reasons.append(f"{prefix}{label}: {row[diff_key]}")
    return reasons


def _add_minutes(t: time, minutes: int) -> time:
    """time に分を加算 (日跨ぎは丸めて捨てる, デモ用途のみなので簡易実装)."""
    dt = datetime.combine(date(2000, 1, 1), t) + timedelta(minutes=minutes)
    return dt.time()


def _apply_mock_normal_route_arrival(reports, att: AttendanceLookup) -> None:
    """通常経路到着時間の疑似データを埋め込む (デモ用, cfg.mock_normal_route_arrival=True 時のみ).

    伝票No のハッシュで2パターン (退勤と同じ/異なる) を均等に割り振り, 可能な限り
    全行に値を入れる (空欄パターンは廃止). 勤怠データが無い行のみ結果的に空欄のまま
    残る (デモ上の意図的な空欄ではなく, 実データ欠落によるもの).
    実データ提供後は cfg.mock_normal_route_arrival=False に戻せば消える (副作用なし).
    """
    for r in reports:
        name = r.resolved_name_norm or r.inputter_name_norm
        pattern = sum(ord(c) for c in r.voucher_no) % 2
        for leg in r.legs:
            is_move = leg.is_movement_leg or leg.transport in MOVEMENT_TRANSPORTS
            if not is_move or leg.time_end is None or leg.leg_date is None:
                continue
            day = att.get(name, leg.leg_date)
            if day is None or day.clock_out is None:
                continue
            clock_out_t = day.clock_out.time() if hasattr(day.clock_out, "time") else day.clock_out
            if pattern == 0:
                leg.normal_route_arrival = clock_out_t          # 退勤と同じ
            else:
                leg.normal_route_arrival = _add_minutes(clock_out_t, 30)  # 退勤と異なる
            leg.normal_route_arrival_status = "ok"


def build_check_sheet(reports, employees, customers, approver_rules,
                      attendance_days, cfg, import_log=None) -> dict:
    grade_lookup = _build_grade_lookup(attendance_days)
    eidx, cidx = enrich_reports(reports, employees, customers, cfg, grade_lookup)
    aidx = build_approver_index(approver_rules)
    att = AttendanceLookup(attendance_days)

    perdiem_master = None
    per_day_codes: set[str] = set()
    if getattr(cfg, "allowance_master_path", None):
        master = load_allowance_master(cfg.allowance_master_path)
        perdiem_master = master.get("手当1")
        per_day_codes = load_allowance_per_day_codes(cfg.allowance_master_path)

    if getattr(cfg, "mock_normal_route_arrival", False):
        _apply_mock_normal_route_arrival(reports, att)

    dup_results = check_duplicate(reports, cfg)        # voucher -> CheckResult

    primary_rows, secondary_rows = [], []
    diff_rows, reject_rows = [], []

    for no, r in enumerate(reports, start=1):
        cr_trip = check_trip_reality(r, att, cfg)
        cr_labor = check_labor(r, att, cfg)
        cr_amt = check_amount(r, cfg, compute_allowances(r, perdiem_master, per_day_codes))
        cr_dup = dup_results.get(r.voucher_no)
        cr_rcpt = check_receipt(r, cfg)
        cr_appr = check_approval_route(r, aidx, eidx, cfg)

        # 観点ごとの (raw CheckResult, spec語彙ステータス)
        axes = [
            ("出張実態", cr_trip),
            ("労務", cr_labor),
            ("金額規程", cr_amt),
            ("二重申請", cr_dup),
            ("領収書", cr_rcpt),
            ("承認ルート", cr_appr),
        ]
        spec_status = {ax: to_axis_vocab(ax, (cr.status if cr else OK)) for ax, cr in axes}
        # 01_一次承認 の総合判定・要確認項目からは 承認ルート を除外する
        # (二次承認者向けの観点のため; 03_差異一覧/04_差戻し文面候補には引き続き含める).
        primary_axes = [(ax, cr) for ax, cr in axes if ax != "承認ルート"]

        # 定時外の移動時間 (01シート '2.' ブロック) も総合判定に反映する
        # (2026-07-30 客先指摘). データなしの窓は除外 — 詳細は _off_hours_axis_status.
        off_hours_rows = _build_off_hours_rows_for_report(r, att, cfg)
        off_hours_status = _off_hours_axis_status(off_hours_rows)

        overall = worst_status(
            [spec_status[ax] for ax, _ in primary_axes] + [off_hours_status]
        )

        # 要確認項目 (OK 以外の観点名, 01表示分のみ)
        flagged_axes = [ax for ax, _ in primary_axes if spec_status[ax] != OK]
        if off_hours_status != OK:
            flagged_axes.append("定時外移動時間")
        # 差戻し候補 (OK 以外の観点の suggestion を連結; 全観点分を漏らさない)
        suggestions = []
        for ax, cr in axes:
            if cr and spec_status[ax] != OK and cr.suggestion:
                suggestions.append(cr.suggestion)

        period = f"{_d(r.date_min)}〜{_d(r.date_max)}"
        detail = _trip_detail_summary(r, att, perdiem_master, per_day_codes)
        nights = detail.pop("宿泊泊数")
        lodging_label = ""
        if nights:
            lodging_label = f"{nights}泊・" + ("整合" if spec_status["金額規程"] == OK else "要確認")

        # --- 01_一次承認チェック ---
        # 列順: No./伝票No./入力者名/社員番号/所属/承認状態 →
        # 1.出張実態の確認 → 2.労務・健康管理の確認(定時外の移動時間 勤務実態) →
        # 3.出張費・宿泊費上限確認 → 4.全体チェック → 詳細
        # (承認ルート/差戻し候補は 01 表示からは除外. 2026-07-17 客先依頼で
        #  定時外の移動時間 勤務実態を独立シートから 2.の直後に組み込み.
        #  2026-07-23 客先依頼で 通常経路到着時間/整合/労務実態 の3列を削除)
        # 2026-07-31 客先依頼: 定時外の移動時間を日単位で表示するため, 1伝票を
        # 「明細日付の数」だけの行に展開する. 伝票単位の値 (金額/総合判定/取引先等) は
        # 先頭日の行にのみ入れ, 2日目以降は空欄にする — 合計金額を集計したときに
        # 日数分だけ多重計上されるのを防ぐため.
        repeat_cells = {
            "No.": no,
            "伝票No.": r.voucher_no,
            "入力者名": r.inputter_name_raw,
            "社員番号": r.employee_id or "(未突合)",
            "所属": r.department or "",
            "承認状態": r.approval_status,
            "出張期間": period,
            "出張先": detail["出張先"],
            "取引先": detail["取引先"],
        }
        first_only_cells = {
            "出張実態": spec_status["出張実態"],
            "日当金額": detail["日当金額"],
            "滞在費補助金額": detail["滞在費補助金額"],
            "宿泊費（ホテル代）": detail["宿泊費（ホテル代）"],
            "日当計": detail["日当計"],
            "宿泊費・手当の整合": lodging_label,
            "合計金額": r.total_amount,
            "金額確認": spec_status["金額規程"],
            "二重申請確認": spec_status["二重申請"],
            "証跡・領収書確認": spec_status["領収書"],
            "総合判定": overall,
            "要確認項目": "・".join(flagged_axes) if flagged_axes else "",
        }
        total_days = len(off_hours_rows)
        for day_no, oh_row in enumerate(off_hours_rows, start=1):
            day_label = f"{day_no}日目"
            if total_days > 1 and day_no == total_days:
                day_label += "(最終日)"
            primary_rows.append({
                **repeat_cells,
                **(first_only_cells if day_no == 1
                   else {k: "" for k in _PRIMARY_FIRST_ONLY_KEYS}),
                **oh_row,
                "日次": day_label,
                # 帯の塗り分け用 (先頭 '_' の内部キーは出力列に含まれない)
                "_group": r.voucher_no,
            })

        # --- 03_差異一覧 (OK以外の観点) ---
        for ax, cr in axes:
            if cr and spec_status[ax] != OK:
                diff_rows.append({
                    "伝票No.": r.voucher_no,
                    "入力者名": r.inputter_name_raw,
                    "観点": ax,
                    "判定": spec_status[ax],
                    "判定理由": cr.detail,
                    "確認先システム": _AXIS_SYSTEM.get(ax, ""),
                    "対応案": cr.suggestion,
                })

        # 定時外の移動時間 は判定6観点ではないが, 総合判定に反映する以上
        # 差異一覧にも理由を残す (2026-07-30 客先指摘への対応).
        if off_hours_status != OK:
            # データなし の窓は総合判定に反映していないため理由からも除く
            # (閾値以内で OK の窓も同様). 日単位化以降は日付付きで列挙する.
            windows = _off_hours_reasons(off_hours_rows)
            diff_rows.append({
                "伝票No.": r.voucher_no,
                "入力者名": r.inputter_name_raw,
                "観点": "定時外移動時間",
                "判定": off_hours_status,
                "判定理由": "楽々精算と楽々勤怠の時刻差が"
                            f"{cfg.off_hours_diff_ok_minutes}分を超過 (" + " / ".join(windows) + ")",
                "確認先システム": "楽々精算 + 楽々勤怠 (移動/勤務時刻)",
                "対応案": "申請の移動時刻と出勤簿の打刻を照合し, 相違があれば修正を依頼",
            })

        # --- 04_差戻し文面候補 (要確認/NG/未突合 かつ suggestion あり) ---
        for ax, cr in axes:
            if cr and spec_status[ax] in (NEEDS_CHECK, NG, UNMATCHED) and cr.suggestion:
                reject_rows.append({
                    "伝票No.": r.voucher_no,
                    "入力者名": r.inputter_name_raw,
                    "宛先(メール)": r.email or "(未登録)",
                    "理由区分": ax,
                    "判定": spec_status[ax],
                    "差戻し文面候補": cr.suggestion,
                })

        # --- 02_二次承認詳細 (明細レッグ単位) ---
        per, lodge, stay = _allowance_flags(r)
        for leg in r.legs:
            cand = ""
            if leg.dest_candidates:
                cand = " / ".join(f"{n}({b})" for n, b in leg.dest_candidates)
            secondary_rows.append({
                "伝票No.": r.voucher_no,
                "入力者名": r.inputter_name_raw,
                "明細No.": leg.leg_no,
                "明細日付": _d(leg.leg_date),
                "開始": leg.time_start.strftime("%H:%M") if leg.time_start else "",
                "終了": leg.time_end.strftime("%H:%M") if leg.time_end else "",
                "出発地": leg.origin_raw or "",
                "到着地": leg.dest_raw or "",
                "交通機関": leg.transport,
                "金額": leg.amount,
                "証票": leg.receipt_label,
                "日当CD": leg.allowance_cd_perdiem or "",
                "宿泊CD": leg.allowance_cd_lodging or "",
                "滞在CD": leg.allowance_cd_stay or "",
                "勘定科目名": leg.account_name,
                "照合顧客名": leg.dest_customer_name or "",
                "距離区分": leg.dest_distance_band or "",
                "照合状態": leg.dest_match_status,
                "複数候補": cand,
            })

    # 02_二次承認詳細: 日付単位で見やすくするため明細日付順に並び替える
    # (客先要望 2026-07-13). 同日内は伝票No.で安定ソート.
    secondary_rows.sort(key=lambda row: (row["明細日付"] or "9999-99-99", row["伝票No."]))

    # --- 07_マスタ確認 ---
    master_rows = _build_master_check(reports, employees, approver_rules, aidx, cfg)

    # --- 05_取込ログ (取込明細 + 件数照合) ---
    pending = sum(1 for r in reports if not r.is_approved)
    log_rows = list(import_log or [])
    log_rows.append({
        "区分": "件数照合", "ファイル名": "(集計)",
        "件数": len(reports),
        "詳細": f"対象件数(伝票)={len(reports)} / 未承認件数={pending} / 承認済={len(reports)-pending}",
        "結果": "OK",
    })

    # --- 06_判定ルール ---
    rule_rows = _build_rule_rows(cfg)

    # --- 既知欠落バナー ---
    # 2026-08-06 客先指摘: 表の上のバナーは場所を取るだけなので出さない.
    # 唯一の内容だった「勤怠データ対象月」は 05_取込ログ の詳細欄へ移した
    # (どの月の勤怠を読んだかは 未確認(勤怠データ欠落) の原因確認に必要なため).
    banners = list(cfg.known_gaps)
    if attendance_days and import_log:
        months = sorted({(d.work_date.year, d.work_date.month)
                         for d in attendance_days if d.work_date})
        label = "対象月 " + ", ".join(f"{y}-{m:02d}" for y, m in months)
        for row in import_log:
            if "勤怠" in str(row.get("区分", "")):
                row["詳細"] = f'{row.get("詳細", "")} / {label}'.strip(" /")
                break

    return {
        "primary": primary_rows,
        "secondary": secondary_rows,
        "diff": diff_rows,
        "reject": reject_rows,
        "import_log": log_rows,
        "rules": rule_rows,
        "master_check": master_rows,
        "banners": banners,
    }


def _build_master_check(reports, employees, approver_rules, aidx, cfg=None) -> list[dict]:
    """マスタ品質問題 (未登録/重複/不整合) を 07_マスタ確認 行に整形.

    cfg.master_check_exclude_names に指定された氏名は出力しない
    (2026-07-30 客先依頼: 対応不要と確認済みの氏名を除外).
    """
    rows: list[dict] = []
    emp_names = {e.name_norm for e in employees if e.name_norm}
    roster_names = {norm(r.employee_name_norm or r.employee_name_raw)
                    for r in approver_rules}
    roster_names.discard("")

    # (a) 申請者が従業員マスタ/20期名簿に未登録
    seen_applicant = set()
    for r in reports:
        key = r.inputter_name_norm
        if key in seen_applicant:
            continue
        seen_applicant.add(key)
        rkey = r.resolved_name_norm or key
        if r.employee_match_status == "未突合":
            rows.append({
                "種別": "社員マスタ未突合", "対象": r.inputter_name_raw,
                "詳細": "入力者名が社員マスタと突合できない", "対応": "氏名表記を社員マスタと突合・修正",
            })
        elif aidx.get(rkey) is None and aidx.get(key) is None:
            rows.append({
                "種別": "承認者名簿 未登録", "対象": r.inputter_name_raw,
                "詳細": "申請者が20期承認者名簿に未登録 (出張命令者を判定不可)",
                "対応": "20期名簿に申請者を登録",
            })

    # (b) 名簿にあるが社員マスタに無い氏名 (不整合)
    for nm in sorted(roster_names - emp_names):
        rows.append({
            "種別": "名簿/マスタ不整合", "対象": nm,
            "詳細": "20期名簿に存在するが社員マスタに無い氏名",
            "対応": "社員マスタ/名簿の氏名表記を突合・統一",
        })

    # (c) 社員マスタ内の氏名重複
    name_counts = Counter(e.name_norm for e in employees if e.name_norm)
    for nm, cnt in name_counts.items():
        if cnt > 1:
            rows.append({
                "種別": "社員マスタ重複", "対象": nm,
                "詳細": f"同一氏名が {cnt} 件 (氏名のみでは一意化不可)",
                "対応": "社員番号で区別 / 重複登録を確認",
            })

    excluded = {norm(n) for n in (getattr(cfg, "master_check_exclude_names", None) or ())}
    if excluded:
        rows = [row for row in rows if norm(row["対象"]) not in excluded]

    return rows


def _build_rule_rows(cfg) -> list[dict]:
    """06_判定ルール: 使用した閾値・区分・規程提供状況を行に整形."""
    al = cfg.amount_limits
    rows = [
        {"項目": "氏名ファジー閾値", "値": cfg.fuzzy_name_threshold, "備考": "difflib ratio (社員突合)"},
        {"項目": "地名照合閾値", "値": cfg.place_match_threshold, "備考": "rapidfuzz partial_ratio (顧客突合)"},
        {"項目": "深夜発 閾値", "値": f"< {cfg.late_night_start_before:02d}:00", "備考": "移動開始がこれ以前"},
        {"項目": "深夜着 閾値", "値": f"> {cfg.late_night_end_after:02d}:00", "備考": "移動終了がこれ以降"},
        {"項目": "労務時刻乖離 閾値", "値": f"{cfg.labor_time_gap_minutes}分", "備考": "移動終了-勤務終了の差がこれ超で要確認"},
        {"項目": "定時外移動時間 定時", "値": f"{cfg.off_hours_work_start.strftime('%H:%M')}〜{cfg.off_hours_work_end.strftime('%H:%M')}",
         "備考": "01シート'3.': 平日の定時前/勤務時間/定時後の境界"},
        {"項目": "定時外移動時間 差分閾値", "値": f"{cfg.off_hours_diff_ok_minutes}分",
         "備考": "01シート'3.': 差分がこれ以下=OK, 超えたら要確認 (客先確定 2026-07-16)"},
        {"項目": "領収書 高額暫定閾値", "値": cfg.receipt_high_value_provisional, "備考": "規程未提供時の高額判定"},
        {"項目": "領収書 検知下限", "値": cfg.receipt_min_amount_to_flag, "備考": "これ未満の少額は要確認にしない"},
        {"項目": "金額規程(上限)提供", "値": "あり" if cfg.has_amount_rules() else "未提供",
         "備考": "未提供時は金額上限を要確認(NGは出さない)"},
        {"項目": "領収書要否規程提供", "値": "あり" if cfg.has_receipt_rule() else "未提供",
         "備考": "未提供時は暫定閾値で判定"},
        {"項目": "確認のみ=承認 扱い", "値": "する" if cfg.confirm_only_counts_as_approval else "しない",
         "備考": "(確認)承認者を正式承認と見なすか"},
        {"項目": "出張日当_一般職", "値": al.get("出張日当", {}).get("一般職", "未提供"), "備考": "円/日"},
        {"項目": "出張日当_管理職", "値": al.get("出張日当", {}).get("管理職", "未提供"), "備考": "円/日"},
        {"項目": "滞在補助費_一般職", "値": al.get("滞在補助費", {}).get("一般職", "未提供"), "備考": "円/日"},
        {"項目": "滞在補助費_管理職", "値": al.get("滞在補助費", {}).get("管理職", "未提供"), "備考": "円/日"},
        {"項目": "出張加算日当_一般職", "値": al.get("出張加算日当", {}).get("一般職", "未提供"), "備考": "円/日"},
        {"項目": "出張加算日当_主任以上", "値": al.get("出張加算日当", {}).get("主任以上", "未提供"), "備考": "円/日"},
        {"項目": "ホテル代_東京23区_管理職", "値": al.get("ホテル代", {}).get("東京23区", {}).get("管理職", "未提供"), "備考": "円"},
        {"項目": "ホテル代_東京23区_一般職", "値": al.get("ホテル代", {}).get("東京23区", {}).get("一般職", "未提供"), "備考": "円"},
        {"項目": "ホテル代_その他_管理職", "値": al.get("ホテル代", {}).get("その他", {}).get("管理職", "未提供"), "備考": "円"},
        {"項目": "ホテル代_その他_一般職", "値": al.get("ホテル代", {}).get("その他", {}).get("一般職", "未提供"), "備考": "円"},
    ]
    for g in cfg.known_gaps:
        rows.append({"項目": "既知の前提/欠落", "値": "", "備考": g})
    return rows
