"""データモデル (canonical internal record shapes) と判定ステータス語彙.

データ契約 §1 に基づく. すべてのローダ・照合・ルール・出力はこれらの型を共有する.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time

# ---------------------------------------------------------------------------
# 判定ステータス語彙 (status vocabulary)
# ---------------------------------------------------------------------------
OK = "OK"
NEEDS_CHECK = "要確認"
NG = "NG"
UNMATCHED = "未突合"
MULTI = "複数候補"                # 同点候補が複数で1件に確定できない (§6.1 赤)
RULE_MISSING = "未確認(規程未提供)"
ATT_MISSING = "未確認(勤怠データ欠落)"

# 総合判定の悪さ順 (大きいほど悪い). build_summary で worst を取る.
# 複数候補 は §6.1 で NG と同じ赤. 承認者が必ず確認すべきため 要確認 と同等の重さ.
SEVERITY = {
    OK: 0,
    RULE_MISSING: 1,
    ATT_MISSING: 1,
    UNMATCHED: 2,
    MULTI: 3,
    NEEDS_CHECK: 3,
    NG: 4,
}

# ---------------------------------------------------------------------------
# 観点別の spec 許容ステータス語彙 (§3.4). 出力の各観点列はこの集合に丸める.
# 内部はより細かい語彙 (ATT_MISSING 等) を保持し, 表示層で spec 語彙へ写像する.
# ---------------------------------------------------------------------------
AXIS_VOCAB = {
    "出張実態": {OK, NEEDS_CHECK, UNMATCHED},
    "労務": {OK, NEEDS_CHECK},
    "金額規程": {OK, NEEDS_CHECK, NG},
    "二重申請": {OK, NEEDS_CHECK, NG},
    "領収書": {OK, NEEDS_CHECK, NG},
    "承認ルート": {OK, NEEDS_CHECK, UNMATCHED},
}


def to_axis_vocab(axis: str, status: str) -> str:
    """内部ステータスを観点の spec 許容語彙 (§3.4) に写像する.

    許容集合に既にあればそのまま. それ以外は意味の近い許容値へ落とす:
      - ATT_MISSING/RULE_MISSING/MULTI → 要確認 (確認が必要)
      - NG が観点で非許容 (出張実態/承認ルート) → 要確認
      - UNMATCHED が観点で非許容 → 要確認
    """
    allowed = AXIS_VOCAB.get(axis)
    if allowed is None or status in allowed:
        return status
    # 勤怠欠落/規程未提供 は『判定不能(データ/規程未提供)』の meta 状態.
    # 要確認に潰すと全件要確認になり一次判断のトリアージが効かなくなるため,
    # 灰色の 未確認(...) としてそのまま残す (banner/取込ログ/判定ルールで明示).
    if status in (ATT_MISSING, RULE_MISSING):
        return status
    # 複数候補 は実在の『要確認(赤)』なので保持
    if status == MULTI:
        return status
    # NG が観点で非許容 (出張実態/承認ルート) → 要確認
    if status == NG and NG not in allowed:
        return NEEDS_CHECK
    # 未突合 が観点で非許容 → 要確認
    if status == UNMATCHED and UNMATCHED not in allowed:
        return NEEDS_CHECK
    return NEEDS_CHECK if NEEDS_CHECK in allowed else OK


def worst_status(statuses) -> str:
    """複数ステータスから最悪 (最大 severity) を返す. 空なら OK."""
    items = [s for s in statuses if s]
    if not items:
        return OK
    return max(items, key=lambda s: SEVERITY.get(s, 0))


# ---------------------------------------------------------------------------
# チェック結果 (per-check result)
# ---------------------------------------------------------------------------
@dataclass
class CheckResult:
    status: str                      # 上記語彙のいずれか
    detail: str = ""                 # 人間可読の判定理由
    evidence: dict = field(default_factory=dict)  # 機械可読の根拠
    suggestion: str = ""             # 差戻し文面候補 / 推奨対応


# ---------------------------------------------------------------------------
# 照合結果 (resolution results)
# ---------------------------------------------------------------------------
@dataclass
class ResolveResult:
    employee_id: str | None
    status: str                      # 突合 | 別名突合 | 未突合
    matched_name: str | None = None
    score: float | None = None


@dataclass
class PlaceMatch:
    customer_no: str | None
    customer_name: str | None
    score: float | None
    status: str                      # 突合 | 別名突合 | 未突合 | 複数候補
    distance_band: str | None = None
    km_lower: int | None = None
    km_upper: int | None = None
    candidates: list | None = None   # 複数候補時: [(取引先名, 距離区分), ...]
    prefecture: str | None = None    # 顧客都道府県 (ホテル代地域判定用)


# ---------------------------------------------------------------------------
# (a) 出張精算レポート (one per 伝票No)
# ---------------------------------------------------------------------------
@dataclass
class ApprovalEntry:
    slot: int                        # 1..5
    approver_name_raw: str
    approver_name_norm: str
    approval_date: date | None = None
    is_confirm_only: bool = False
    approver_employee_id: str | None = None


@dataclass
class ExpenseLeg:
    voucher_no: str                  # FK -> ExpenseReport.voucher_no
    leg_no: int                      # 明細No.
    leg_date: date | None            # 明細日付
    time_start: time | None
    time_end: time | None
    origin_raw: str | None           # 出発地
    dest_raw: str | None             # 到着地
    transport: str                   # 交通機関
    amount: int                      # 金額
    subtotal: int                    # 小計
    receipt_label: str               # 証票(領収証) ラベル
    has_receipt: bool
    allowance_cd_perdiem: str | None  # 手当1CD (日当)  TEXT
    allowance_cd_lodging: str | None  # 手当2CD (宿泊料) TEXT
    allowance_cd_stay: str | None     # 手当3CD (滞在費補助) TEXT
    remark: str | None               # フリー1(備考)
    account_name: str                # 勘定科目名
    normal_route_arrival: time | None = None       # 通常経路到着時間 (パース済み)
    normal_route_arrival_status: str = "empty"     # 'empty' | 'ok' | 'invalid'
    # --- 下流で付与 (enriched) ---
    dest_customer_no: str | None = None
    dest_customer_name: str | None = None
    dest_match_score: float | None = None
    dest_match_status: str = UNMATCHED
    dest_distance_band: str | None = None
    dest_km_lower: int | None = None
    dest_km_upper: int | None = None
    dest_candidates: list | None = None  # 複数候補時の候補一覧
    dest_prefecture: str | None = None   # 顧客都道府県 (ホテル代地域判定用)
    is_movement_leg: bool = False


# 移動レッグとみなす交通機関 (mode), 非移動 (activity)
# 飛行機/ﾀｸｼｰ(同乗) が漏れていると labor.py の勤怠照合・checksheet.py の移動時間集計から
# 当該レッグが丸ごと除外されるため (2026-07-07 金額不一致調査で発覚), 明示的に含める.
MOVEMENT_TRANSPORTS = {"電車･ﾊﾞｽ", "車", "車(同乗)", "徒歩", "ﾀｸｼｰ", "ﾀｸｼｰ(同乗)", "ｶﾞｿﾘﾝ代", "飛行機"}
ACTIVITY_TRANSPORTS = {"作業･打合せ", "ﾎﾃﾙ", "宿泊税", "入湯税", "駐車代"}


@dataclass
class ExpenseReport:
    voucher_no: str                  # 伝票No. (PK)
    application_no: str | None
    declared_leg_count: int          # 行数
    actual_leg_count: int
    total_amount: int                # 合計 (header)
    computed_total: int              # sum(leg.amount)
    allowance_total_declared: int    # 手当計 (header). 明細の金額欄が0でも
                                      # ここに手当計算システム側の実額が入ることがある.
    inputter_name_raw: str           # 入力者名 or 申請者名
    inputter_name_norm: str
    legs: list[ExpenseLeg]
    approvers: list[ApprovalEntry]
    is_approved: bool
    approval_status: str             # 承認済 | 未承認(PENDING)
    date_min: date | None
    date_max: date | None
    source_file: str
    # --- 下流で付与 ---
    applicant_cd: str | None = None  # 申請者CD (新CSV のみ; IDルックアップに使用)
    employee_id: str | None = None
    employee_match_status: str = UNMATCHED  # 突合 | 別名突合 | 未突合
    department: str | None = None
    email: str | None = None
    resolved_name_norm: str | None = None   # 突合した社員マスタ氏名の正規化形 (別名解決後)
    grade: str | None = None                # 役職区分 (employee masterから引き継ぎ)


# ---------------------------------------------------------------------------
# (c) 勤怠日次 (one per employee, date)
# ---------------------------------------------------------------------------
PRESENCE_OFFICE = "在席(出社)"
PRESENCE_TELEWORK = "在席(テレワーク)"
PRESENCE_LEAVE = "休暇"
PRESENCE_NONWORK = "非労働日"
PRESENCE_HOLIDAY_WORK = "休日出勤"
PRESENCE_UNKNOWN = "不明"

LEAVE_CONTENTS = {
    "有休", "AM有休", "PM有休", "代休", "育児休業", "リフレッシュ休暇",
    "特休/その他", "特休", "欠勤", "産前産後休業", "介護休業",
}


@dataclass
class AttendanceDay:
    emp_id_raw: str                  # 社員番号 (file 内のみ信頼)
    name_raw: str                    # 氏名
    name_norm: str                   # join key
    department: str | None
    role: str | None
    work_date: date | None           # 日付
    weekday: str | None              # 曜日
    calendar_type: str | None        # カレンダー
    application_content: str | None  # 申請内容
    clock_in: datetime | None
    clock_out: datetime | None
    telework_in: datetime | None
    telework_out: datetime | None
    move_start: time | None
    move_end: time | None
    move2_start: time | None
    move2_end: time | None
    holiday_work_minutes: bool        # 休出時間系のいずれか filled
    remark: str | None
    presence: str                    # 派生
    source_file: str
    sheet_name: str


def effective_work_times(day: "AttendanceDay | None"):
    """勤務開始/終了 時刻を取得. clock_in/clock_out が無い日は
    move_start/move_end (無ければ move2_start/move2_end) で代用する.

    非労働日に出張した場合は通常の出退勤打刻が無く, 移動時刻のみが
    勤怠側に記録されるため (楽々精算からの自動連携). 複数区間があれば
    最早の開始・最遅の終了を採用する.
    """
    if day is None:
        return None, None
    if day.clock_in is not None:
        start = day.clock_in
    else:
        starts = [t for t in (day.move_start, day.move2_start) if t is not None]
        start = min(starts) if starts else None
    if day.clock_out is not None:
        end = day.clock_out
    else:
        ends = [t for t in (day.move_end, day.move2_end) if t is not None]
        end = max(ends) if ends else None
    return start, end


# ---------------------------------------------------------------------------
# (d) 従業員マスタ
# ---------------------------------------------------------------------------
@dataclass
class Employee:
    employee_id: str                 # 社員番号 (PK)
    name_raw: str                    # 氏名
    name_norm: str                   # canonical join key
    email: str | None
    department: str | None           # 所属部署
    jurisdiction: str | None         # 所属管轄 東/西
    grade: str | None = None         # 役職区分 (管理職/一般職/主任以上 等)


# ---------------------------------------------------------------------------
# (e) 顧客マスタ
# ---------------------------------------------------------------------------
@dataclass
class Customer:
    customer_no: str                 # 取引先番号 (PK, string)
    name_raw: str                    # 取引先名
    name_norm_base: str              # 正規化ベース
    site_token: str | None           # 拠点トークン
    distance_band: str               # 距離区分 raw
    km_lower: int                    # 下限
    km_upper: int | None             # 上限 (None=open)
    prefecture: str | None           # 都道府県
    region: str | None               # 東西


# ---------------------------------------------------------------------------
# (f) 承認者ルール (20期 roster)
# ---------------------------------------------------------------------------
@dataclass
class ApproverRule:
    employee_name_raw: str           # 氏名
    employee_name_norm: str          # join key
    trip_approver_raw: str           # 出張命令者
    trip_approver_norm: str
    trip_approver_is_confirm_only: bool
    attendance_approver_raw: str | None  # 勤怠管理 承認者
    attendance_approver_norm: str | None
    department: str | None
    section: str | None
    team: str | None
    dept_code: str | None


# ---------------------------------------------------------------------------
# 勤怠ルックアップ (attendance lookup with May-gap awareness)
# ---------------------------------------------------------------------------
class AttendanceLookup:
    """(name_norm, date) -> AttendanceDay. 月欠落の検知を担う.

    data_months: 勤怠データに存在する (year, month) 集合.
    対象日が data_months に無ければ get() は None かつ is_month_missing()=True.
    """

    def __init__(self, days: list[AttendanceDay]):
        self._index: dict[tuple[str, date], AttendanceDay] = {}
        self.data_months: set[tuple[int, int]] = set()
        self.names: set[str] = set()
        for d in days:
            if d.work_date is None:
                continue
            self.data_months.add((d.work_date.year, d.work_date.month))
            self.names.add(d.name_norm)
            self._index[(d.name_norm, d.work_date)] = d

    def get(self, name_norm: str, d: date) -> AttendanceDay | None:
        return self._index.get((name_norm, d))

    def is_month_present(self, d: date) -> bool:
        return (d.year, d.month) in self.data_months

    def has_name(self, name_norm: str) -> bool:
        return name_norm in self.names
