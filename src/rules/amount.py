"""金額規程チェック (§6.3 amount-rule check).

出張精算レポート (ExpenseReport) に対し, 金額面の妥当性を2系統で判定する.

(1) データ整合性 (data-integrity, 常時実行・規程不要):
    - 合計(申告) と 明細金額合計(計算) の一致
    - 行数(宣言) と 実明細数の一致
    - 各明細の 金額>=0 かつ 小計==金額

(2) 規程ベース上限判定 (rule-based, 旅費規定 必須):
    手当CD / account_name ごとに上限を引き, 役職・地域を考慮して判定.
    役職 (grade) 不明時は一般職上限以下→OK, 管理職上限以下→要確認, 超過→NG.
    旅費規定 未提供時は RULE_MISSING を返す.

(3) 勘定科目分類チェック:
    委託サービス費で交通・宿泊系の明細が計上されている場合は要確認.
    なお 車(transport) + 賃借料(account_name) は enrich_reports() (checksheet.py)
    で読込直後に旅費交通費へ自動置換されるため, 本モジュールに到達する時点で
    賃借料表記は残らない (2026-07-01 客先回答: レンタカー・社用車とも区分不要).
"""
from __future__ import annotations

# 東京23区の区名リスト (都道府県フィールドに区名が含まれる場合に 東京23区 と判定)
_TOKYO_23_KU: frozenset[str] = frozenset([
    "千代田区", "中央区", "港区", "新宿区", "文京区",
    "台東区", "墨田区", "江東区", "品川区", "目黒区",
    "大田区", "世田谷区", "渋谷区", "中野区", "杉並区",
    "豊島区", "北区", "荒川区", "板橋区", "練馬区",
    "足立区", "葛飾区", "江戸川区",
])


def _is_tokyo_23ku(prefecture: str | None) -> bool:
    """都道府県フィールドが東京23区内かを判定する.

    prefecture に "東京" が含まれ、かつ23区のいずれかの区名が含まれる場合のみ True.
    "東京都八王子市" など多摩地区は False になる.
    prefecture が区名のみ ("港区") でも True を返す.
    """
    if not prefecture:
        return False
    pref = prefecture.strip()
    if "東京" in pref:
        return any(ku in pref for ku in _TOKYO_23_KU)
    return any(ku in pref for ku in _TOKYO_23_KU)


from models import (
    ExpenseReport,
    CheckResult,
    OK,
    NG,
    NEEDS_CHECK,
    RULE_MISSING,
    worst_status,
)
from config import Config


# ---------------------------------------------------------------------------
# データ整合性チェック
# ---------------------------------------------------------------------------

def _check_integrity(r: ExpenseReport) -> tuple[str, list[str], dict]:
    reasons: list[str] = []
    evidence: dict = {}
    statuses: list[str] = []

    gap = r.total_amount - r.computed_total
    has_allowance = any(
        (leg.allowance_cd_perdiem or leg.allowance_cd_lodging or leg.allowance_cd_stay)
        for leg in r.legs
    )
    if r.computed_total > r.total_amount:
        statuses.append(NG)
        reasons.append(
            f"明細金額合計が申告合計を超過: 計算{r.computed_total} > 申告{r.total_amount}"
        )
        evidence["合計超過"] = {
            "申告合計": r.total_amount,
            "計算合計": r.computed_total,
            "超過額": -gap,
        }
    elif gap > 0 and not has_allowance:
        statuses.append(NEEDS_CHECK)
        reasons.append(
            f"申告合計と明細合計に差額{gap}があるが手当コードなし: 要確認"
        )
        evidence["差額要確認"] = {
            "申告合計": r.total_amount,
            "計算合計": r.computed_total,
            "差額": gap,
        }
    elif gap > 0:
        evidence["手当相当差額"] = {
            "申告合計": r.total_amount,
            "明細金額合計": r.computed_total,
            "手当相当差額": gap,
        }

    if r.actual_leg_count != r.declared_leg_count:
        statuses.append(NEEDS_CHECK)
        reasons.append(
            f"行数不一致: 宣言{r.declared_leg_count} != 実{r.actual_leg_count}"
        )
        evidence["行数不一致"] = {
            "宣言行数": r.declared_leg_count,
            "実明細数": r.actual_leg_count,
        }

    bad_legs: list[dict] = []
    for leg in r.legs:
        if leg.amount < 0 or leg.subtotal != leg.amount:
            bad_legs.append({"明細No": leg.leg_no, "金額": leg.amount, "小計": leg.subtotal})
    if bad_legs:
        statuses.append(NG)
        reasons.append(f"明細金額異常 ({len(bad_legs)}件): 金額<0 または 小計!=金額")
        evidence["明細金額異常"] = bad_legs

    return worst_status(statuses), reasons, evidence


# ---------------------------------------------------------------------------
# 役職別上限ヘルパー
# ---------------------------------------------------------------------------

def _grade_caps(table: dict, grade: str | None) -> tuple[int | None, int | None]:
    """(exact_cap, None) if grade known; (一般職_cap, higher_cap) if grade unknown."""
    if not isinstance(table, dict):
        return None, None
    if grade:
        if grade in table:
            return table[grade], None
        # 主任以上エイリアス
        if any(t in grade for t in ("主任", "リーダー", "係長", "班長")):
            cap = table.get("主任以上")
            if cap is not None:
                return cap, None
        # 管理職エイリアス
        if any(t in grade for t in ("部長", "課長", "副参事", "参事", "取締役", "執行役")):
            cap = table.get("管理職")
            if cap is not None:
                return cap, None
        # フォールバック: 一般職
        cap = table.get("一般職")
        return cap, None
    # 役職不明: 一般職と最上位の差で判定
    cap_gen = table.get("一般職")
    cap_high = table.get("管理職") or table.get("主任以上")
    return cap_gen, cap_high


def _check_grade_allowance(
    leg_no: int, kind: str, amount: int,
    table: dict, grade: str | None,
    overs: list, needs_check: list,
) -> None:
    cap_exact, cap_high = _grade_caps(table, grade)

    if grade:
        # 役職確定: 単一上限で判定
        if cap_exact is None:
            needs_check.append({
                "明細No": leg_no, "種別": kind, "金額": amount,
                "理由": f"役職({grade})の上限が規程に未定義",
            })
        elif amount > cap_exact:
            overs.append({"明細No": leg_no, "種別": kind, "金額": amount,
                          "上限": cap_exact, "役職": grade})
    else:
        # 役職不明: 二段階判定
        if cap_exact is None:
            needs_check.append({"明細No": leg_no, "種別": kind, "金額": amount,
                                 "理由": "役職不明かつ規程未定義"})
            return
        if amount <= cap_exact:
            return  # 一般職上限以内 → OK
        if cap_high is not None and amount <= cap_high:
            needs_check.append({
                "明細No": leg_no, "種別": kind, "金額": amount,
                "一般職上限": cap_exact, "上位役職上限": cap_high,
                "理由": "役職確認が必要 (一般職超過・管理職上限以内)",
            })
        else:
            overs.append({
                "明細No": leg_no, "種別": kind, "金額": amount,
                "一般職上限": cap_exact,
                **({"上位役職上限": cap_high} if cap_high else {}),
                "理由": "全役職上限超過",
            })


# ---------------------------------------------------------------------------
# ホテル代チェック
# ---------------------------------------------------------------------------

def _check_hotel(
    leg_no: int, amount: int, prefecture: str | None, grade: str | None,
    hotel_limits: dict, overs: list, needs_check: list,
) -> None:
    region = "東京23区" if _is_tokyo_23ku(prefecture) else "その他"
    rtable = hotel_limits.get(region, hotel_limits.get("その他", {}))
    _check_grade_allowance(leg_no, f"ホテル代({region})", amount, rtable, grade, overs, needs_check)


# ---------------------------------------------------------------------------
# 長距離運転手当チェック
# ---------------------------------------------------------------------------

def _check_long_distance(leg_no: int, amount: int, km_lower: int | None,
                          has_lodging: bool, limits: dict,
                          overs: list, needs_check: list) -> None:
    ld = limits.get("長距離運転手当", {})
    if not ld:
        return
    trip_type = "宿泊時" if has_lodging else "日帰り"
    lower = ld.get(f"{trip_type}_下限")
    upper = ld.get(f"{trip_type}_上限")
    if lower is None or upper is None:
        return
    # 長距離加算
    add_km = ld.get("長距離加算_km", 300)
    add_amt = ld.get("長距離加算額", 1000)
    if km_lower is not None and km_lower >= add_km:
        upper = upper + add_amt
    if amount < lower:
        needs_check.append({
            "明細No": leg_no, "種別": "長距離運転手当", "金額": amount,
            "下限": lower, "上限": upper, "旅行種別": trip_type,
            "理由": "規程下限未満",
        })
    elif amount > upper:
        overs.append({
            "明細No": leg_no, "種別": "長距離運転手当", "金額": amount,
            "上限": upper, "旅行種別": trip_type,
        })


# ---------------------------------------------------------------------------
# 規程ベース上限判定
# ---------------------------------------------------------------------------

def _check_rules(r: ExpenseReport, cfg: Config,
                 allowances: dict | None = None) -> tuple[str, list[str], dict]:
    if not cfg.has_amount_rules():
        return (
            RULE_MISSING,
            ["旅費規定 未提供のため金額上限判定不可"],
            {"規程": "未提供"},
        )

    limits = cfg.amount_limits
    grade: str | None = getattr(r, "grade", None)

    overs: list[dict] = []       # → NG
    needs_check: list[dict] = [] # → NEEDS_CHECK

    # report に宿泊実態があるか (長距離運転手当の日帰り/宿泊判定)
    has_lodging_cd = any(leg.allowance_cd_lodging for leg in r.legs)

    for leg in r.legs:
        # (1) 出張日当・(3) 滞在補助費 は明細の金額欄では判定しない。
        #     手当CDが付く明細の金額は交通費などの実費であり手当額ではないため
        #     (例: 日当CDが付いた電車代 3,520円 を日当上限 1,700円と比較して
        #     誤って上限超過にしていた。2026-08-06 修正)。
        #     実額はヘッダの手当計から算出し, 下の (8) で1日/1泊あたりに
        #     換算して判定する。

        # (2) ホテル代は 手当2CD ではなく宿泊日ごとに判定する (下の (7) を参照).
        #     実データで 手当2CD が常に空のため, ここでの判定は行わない
        #     (2026-08-06: 二重に上限判定しないよう明細ループからは外した).

        # (4) 出張加算日当 (account_name 識別)
        if leg.account_name and "出張加算日当" in leg.account_name:
            table = limits.get("出張加算日当", {})
            _check_grade_allowance(
                leg.leg_no, "出張加算日当", leg.amount,
                table, grade, overs, needs_check,
            )

        # (5) 長距離運転手当 (account_name 識別)
        if leg.account_name and "長距離運転手当" in leg.account_name:
            _check_long_distance(
                leg.leg_no, leg.amount, leg.dest_km_lower,
                has_lodging_cd, limits, overs, needs_check,
            )

        # (6) 委託サービス費 + 交通・宿泊系 分類チェック
        # 実データは半角表記 (ﾀｸｼｰ/ﾎﾃﾙ等) が主のため半角側を主に列挙し,
        # 表記ゆれ (全角/別名) にも備えて全角側・駐車代・ガソリン代・宿泊税・入湯税も含める.
        _TRAVEL_TRANSPORTS = {
            "電車･ﾊﾞｽ", "電車・バス", "ﾎﾃﾙ", "ホテル", "車", "車(同乗)",
            "ﾀｸｼｰ", "ﾀｸｼｰ(同乗)", "タクシー", "飛行機", "新幹線", "バス",
            "駐車代", "ｶﾞｿﾘﾝ代", "ガソリン代", "宿泊税", "入湯税",
        }
        if (leg.account_name and "委託サービス費" in leg.account_name
                and (leg.transport in _TRAVEL_TRANSPORTS
                     or leg.allowance_cd_perdiem
                     or leg.allowance_cd_lodging
                     or leg.allowance_cd_stay)):
            needs_check.append({
                "明細No": leg.leg_no, "種別": "勘定科目分類",
                "金額": leg.amount,
                "理由": "交通・宿泊費が委託サービス費で申請されています。旅費交通費への変更要確認",
            })

    # (7) ホテル代の上限 — 1泊あたりに換算して判定する
    #
    # 2026-08-06 客先質問「ホテル代上限は宿泊日数に上限を計算しているか？」への対応。
    # 従来は 手当2CD (allowance_cd_lodging) を条件にしていたが, 実データでは
    # この列が常に空のため上限判定が一度も動いていなかった。
    #
    # 実データではホテル代が「滞在全体の合計」として初日の明細1件に計上される
    # (例: 2泊で 18,900円 を1行で計上)。そのため日付でまとめるのではなく,
    # 伝票のホテル関連明細を合算し, 泊数で割った額を1泊分として上限と比較する。
    # 宿泊税・入湯税を含めるのは 2026-07-07 客先確認による。
    hotel_limits = limits.get("ホテル代", {})
    hotel_legs = [lg for lg in r.legs
                  if lg.transport in ("ﾎﾃﾙ", "ホテル", "宿泊税", "入湯税")]
    if hotel_limits and hotel_legs:
        hotel_total = sum(lg.amount or 0 for lg in hotel_legs)
        nights = 1
        if r.date_min and r.date_max:
            nights = max(1, (r.date_max - r.date_min).days)
        per_night = hotel_total // nights
        prefecture = next((lg.dest_prefecture for lg in r.legs if lg.dest_prefecture), None)
        leg_no = next((lg.leg_no for lg in hotel_legs
                       if lg.transport in ("ﾎﾃﾙ", "ホテル")), hotel_legs[0].leg_no)
        if per_night > 0:
            _check_hotel(leg_no, per_night, prefecture, grade,
                         hotel_limits, overs, needs_check)

    # (8) 出張日当・滞在補助費の上限 — 1日 / 1泊あたりに換算して判定する
    #
    # 実額は明細ではなくヘッダの手当計にまとまっているため, 呼び出し側
    # (checksheet.compute_allowances) が算出した単価を受け取って判定する。
    # allowances が渡されない場合 (単体テスト等) はこの判定を行わない。
    if allowances:
        unit = allowances.get("perdiem_unit")
        if unit:
            _check_grade_allowance(0, "出張日当(1日あたり)", unit,
                                   limits.get("出張日当", {}), grade,
                                   overs, needs_check)
        unit = allowances.get("stay_unit")
        if unit:
            _check_grade_allowance(0, "滞在補助費(1泊あたり)", unit,
                                   limits.get("滞在補助費", {}), grade,
                                   overs, needs_check)

    statuses: list[str] = []
    reasons: list[str] = []
    evidence: dict = {}

    if overs:
        statuses.append(NG)
        reasons.append(f"金額上限超過 ({len(overs)}件)")
        evidence["金額上限超過"] = overs
    if needs_check:
        statuses.append(NEEDS_CHECK)
        reasons.append(f"要確認 ({len(needs_check)}件)")
        evidence["要確認"] = needs_check
    if not overs and not needs_check:
        statuses.append(OK)

    return worst_status(statuses), reasons, evidence


# ---------------------------------------------------------------------------
# 公開エントリポイント
# ---------------------------------------------------------------------------

def check_amount(r: ExpenseReport, cfg: Config,
                 allowances: dict | None = None) -> CheckResult:
    """金額規程チェック (§6.3). データ整合性 + 規程上限判定の総合."""
    integ_status, integ_reasons, integ_ev = _check_integrity(r)
    rule_status, rule_reasons, rule_ev = _check_rules(r, cfg, allowances)

    overall = worst_status([integ_status, rule_status])

    reasons = integ_reasons + rule_reasons
    detail = " / ".join(reasons) if reasons else "金額整合性: 問題なし"

    evidence: dict = {}
    evidence.update(integ_ev)
    evidence.update(rule_ev)
    evidence["整合性判定"] = integ_status
    evidence["規程判定"] = rule_status

    suggestion = ""
    if overall in (NG, NEEDS_CHECK):
        parts: list[str] = []
        if "合計超過" in integ_ev:
            parts.append("明細金額の合計が申告合計を超過しています。金額を確認・修正してください")
        if "差額要確認" in integ_ev:
            parts.append("申告合計と明細合計の差額(手当コードなし)の内訳を確認してください")
        if "行数不一致" in integ_ev:
            parts.append("行数(宣言)と明細件数の差異を確認してください")
        if "明細金額異常" in integ_ev:
            parts.append("金額がマイナス/小計不一致の明細を是正してください")
        if "金額上限超過" in rule_ev:
            parts.append("旅費規定の上限を超過した明細の妥当性を確認してください")
        if "要確認" in rule_ev:
            items = rule_ev["要確認"]
            kinds = list({i.get("種別", "") for i in items if i.get("種別")})
            if kinds:
                parts.append(f"{'・'.join(kinds)}について確認が必要です")
        suggestion = "。".join(parts) + "。" if parts else ""

    return CheckResult(
        status=overall,
        detail=detail,
        evidence=evidence,
        suggestion=suggestion,
    )
