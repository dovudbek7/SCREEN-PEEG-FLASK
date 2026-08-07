"""設定 (paths, password, thresholds, 規程デフォルト).

旅費規定ファイルが未提供のため, 金額系の閾値はすべて config 駆動の
プレースホルダ. 実値が無い限り判定は NG ではなく 未確認(規程未提供) を出す.
"""
from __future__ import annotations

import os
import sys
from dataclasses import dataclass, field
from datetime import time

from file_discovery import latest_matching, all_matching, latest_matching_recursive


def _project_root() -> str:
    """データフォルダ (expenses/ 等) を探す基準ディレクトリを返す.

    優先順位:
      1. 環境変数 CHECKSHEET_ROOT (Webアプリなど, リクエストごとに作業フォルダを
         分けたい場合に使う)
      2. PyInstaller で単一実行ファイル化した場合は実行ファイル自身の場所
         (__file__ は展開先の一時フォルダ sys._MEIPASS を指すため使えない)
      3. 通常実行時は このファイルの2つ上 (= リポジトリ直下)
    """
    env = os.environ.get("CHECKSHEET_ROOT")
    if env:
        return os.path.abspath(env)
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# プロジェクトルート
ROOT       = _project_root()
EXPENSES   = os.path.join(ROOT, "expenses")     # 運用担当者が出張精算CSVを配置するフォルダ (ルート直下)
ATTENDANCE = os.path.join(ROOT, "attendance")   # 運用担当者が出勤簿xlsxを配置するフォルダ (ルート直下)
DATA       = os.path.join(ROOT, "出張精算データ一式")
LIST       = os.path.join(DATA, "顧客・社員リスト")


@dataclass
class Config:
    # --- 入力パス (フォルダ内を自動検索, ファイル名のハードコードなし) ---
    # 2026-08-04 客先依頼: 「アップロードしたデータ毎にチェックシートを作成する」
    # 運用に変更. 週次で楽々精算・楽々勤怠が出力されるため, 全ファイルを積み上げると
    # 毎回チェック対象が増え続けてしまう. そのため 精算・勤怠とも
    # 「フォルダ内の最新1ファイルのみ」を採用する入れ替え式にする。
    # (旧: 勤怠のみ all_matching でフォルダ内全ファイルを合算していた)
    expense_csv_path: str = field(default_factory=lambda: latest_matching(EXPENSES, "出張精算_*.csv"))
    attendance_paths: list = field(
        default_factory=lambda: [latest_matching(ATTENDANCE, "出勤簿_*.xlsx")]
    )
    employee_master_path: str = field(default_factory=lambda: latest_matching(LIST, "社員リスト_*.xlsx"))
    customer_master_path: str = field(default_factory=lambda: latest_matching(LIST, "顧客リスト_*.xlsx"))
    approver_roster_path: str = field(default_factory=lambda: latest_matching(DATA, "*評価者・承認者一覧*.xlsx"))
    approver_roster_sheet: str = "20期"
    # 手当CDの正式な金額マスタ (2026-07-10 客先提供). 日当CDから公式の手当金額を
    # 直接引けるため, 明細金額/ヘッダ手当計に依存しない集計が可能になる。
    # 旧データセットには存在しないため任意 (無ければ従来ロジックにフォールバック)。
    allowance_master_path: str | None = field(
        default_factory=lambda: latest_matching_recursive(DATA, "手当マスタ_*.csv")
    )

    # --- デモ用モックデータ ---
    # 通常経路到着時間は実データに列が無いため, 動作デモ用に疑似データを
    # 埋め込むフラグ (2026-07-08 客先依頼). True の間は本番データではなく
    # デモ表示専用とみなすこと. 実列が提供され次第 False に戻す.
    # 2026-07-16: 本番提供までデフォルトTrueにして毎回自動表示 (客先確認用, 開発者依頼).
    # 本番データを渡す直前に False に戻すこと.
    mock_normal_route_arrival: bool = True

    # --- 対象者フィルタ ---
    # 出張精算データの W列 (ヘッダ情報:承認実行者1名) がこの氏名の伝票のみを
    # チェックシートに反映する (2026-07-07 客先指定). None なら全件対象.
    target_approver_filter: str | None = "岡田　高明"

    # --- 出力 ---
    output_dir: str = os.path.join(ROOT, "out")
    output_prefix: str = "出張精算_承認チェックシート"

    # --- 読込パラメータ ---
    # 社員/顧客マスタの復号パスワード (既知の既定値、フォールバック用)。
    # main.py はまずこの既定値で復号を試み, 失敗した場合 (=マスタ側のパスワードが
    # 変更された場合) のみ対話入力(getpass)にフォールバックする。
    master_password: str | None = "peeg0608"
    csv_encoding: str = "cp932"

    # --- 照合閾値 ---
    place_match_threshold: int = 82          # rapidfuzz partial_ratio (0-100)
    fuzzy_name_threshold: float = 1.0         # difflib ratio (0-1)
    allow_adjacent_month_attendance: bool = False  # 隣接月勤怠での暫定照合

    # --- 労務閾値 ---
    late_night_start_before: int = 5          # 時刻 < 05:00 を深夜発とみなす
    late_night_end_after: int = 22            # 時刻 > 22:00 を深夜着とみなす
    labor_time_gap_minutes: int = 60          # 移動終了-勤務終了の差 > 60分 で要確認

    # --- 定時外の移動時間 勤務実態 (08シート) ---
    # 客先確定ルール (2026-07-16):
    #   平日: 定時=9:00〜17:30. 定時前=何時から9時まで移動したか, 定時後=勤務後何時まで移動したか.
    #   休日: 定時の概念なし. 移動・作業時間をそのまま反映 (顧客先への移動/作業・打合せ/顧客先からの移動).
    #   差分チェック: 10分以内=OK, 11分以上=要確認 (NG無し).
    off_hours_work_start: time = time(9, 0)   # 定時前/勤務時間 の境界 (平日のみ)
    off_hours_work_end: time = time(17, 30)   # 勤務時間/定時後 の境界 (平日のみ)
    off_hours_diff_ok_minutes: int = 10       # 差分がこれ以下なら OK, 超えたら要確認

    # --- 金額/規程 (J-4-1 国内出張旅費規定 2025-10-01施行 より) ---
    amount_limits: dict = field(default_factory=lambda: {
        # 手当1CD (日当)
        "出張日当": {"一般職": 1700, "管理職": 3000},
        # 手当3CD (滞在費補助)
        "滞在補助費": {"一般職": 3500, "管理職": 5000},
        # 2026-08-07 客先提供の別表: 食事代を先方/会社が負担した場合は別単価.
        # 手当マスタの 滞在費補助(食事代先方/会社負担) = 手当3CD 013/016/019.
        "滞在補助費(食事代先方/会社負担)": {"一般職": 1000, "管理職": 3000},
        # 手当マスタの 滞在費補助(特例) = 手当3CD 022〜027 (固定 2,000円)
        "滞在補助費(特例)": {"一般職": 2000, "管理職": 2000},
        # account_name で識別: 出張加算日当
        "出張加算日当": {"一般職": 800, "主任以上": 1100},
        # account_name で識別: 長距離運転手当 (range check)
        "長距離運転手当": {
            "日帰り_下限": 2500, "日帰り_上限": 3500,
            "宿泊時_下限": 1500, "宿泊時_上限": 2500,
            "長距離加算_km": 300, "長距離加算額": 1000,
        },
        # 手当2CD (宿泊料) — 地域×役職
        "ホテル代": {
            "東京23区": {"管理職": 15000, "一般職": 13500},
            "その他":   {"管理職": 11000, "一般職": 9500},
        },
    })
    receipt_required_above: int | None = 1000    # 1,000円以上の非免除明細は領収書必須
    receipt_exempt_transports: tuple = ("電車･ﾊﾞｽ",)  # IC/運賃系は領収書免除候補
    # 規程未提供時のフォールバック (旅費規定 入手後は receipt_required_above が優先):
    #   - high_value_provisional: 免除交通機関でもこの額以上・領収書なしは要確認
    #     (例: 新幹線相当の高額交通費の見逃し防止)
    #   - min_amount_to_flag: 非免除でもこの額未満は要確認にしない
    #     (宿泊税/駐車代等の少額付随費による過剰検知を抑制)
    receipt_high_value_provisional: int = 10000
    receipt_min_amount_to_flag: int = 1000
    confirm_only_counts_as_approval: bool = False

    # --- 07_マスタ確認 の除外対象 ---
    # 客先で対応不要と確認済みの氏名は 07_マスタ確認 に出さない
    # (2026-07-30 客先依頼: 名古屋さんは反映不要). 氏名は normalize.norm 後に比較する。
    master_check_exclude_names: tuple = ("名古屋健司",)

    # --- 役職オーバーライド (組織図から手動抽出; 社員マスタに役職列がない場合に使用) ---
    # 値: "管理職" | "一般職"
    role_overrides: dict = field(default_factory=lambda: {
        # ── 管理職 (部長・副部長・課長・課長代理) ──
        "西 三照":      "管理職",   # 代表取締役社長
        "高橋 昭太":    "管理職",   # 管理部長 / 管理課長
        "岡田 高明":    "管理職",   # 技術部長
        "河本 実":      "管理職",   # 技術部副部長
        "浜内 邦嘉":    "管理職",   # ソリューションサポート部長
        "岡部 信一":    "管理職",   # SS部副部長 / 課長
        "茅野 義洋":    "管理職",   # 技術1課長
        "中田 雅史":    "管理職",   # 技術2課長
        "杉原 竜彦":    "管理職",   # 管理課長代理
        # ── 一般職 (係長・係長代理は2026-07-02付で一般職へ変更) ──
        "内田 修平":    "一般職",   # 技術2課1G係長
        "鈴木 景大郎":  "一般職",   # 技術1課1G係長
        "志村 一磨":    "一般職",   # 技術1課2G係長
        "鈴木 和行":    "一般職",   # 技術2課2G係長
        "藤倉 亮":      "一般職",   # 管理課1係長
        "小野 智紀":    "一般職",   # 管理課2係長
        "張 学鑫":      "一般職",   # 技術2課2G係長代理
        "高橋 直樹":    "一般職",   # 技術2課1G係長代理
        "長橋 正輝":    "一般職",   # 技術1課1G係長代理
        "松沢 響":      "一般職",
        "藤岡 拓己":    "一般職",
        "磯 優樹":      "一般職",
        "品川 祐太朗":  "一般職",
        "山本 翔也":    "一般職",
        "清水 雄太":    "一般職",
        "武田 幸大":    "一般職",
        "石川 直樹":    "一般職",
        "俣野 寛太":    "一般職",
        "前田 逸人":    "一般職",
        "岩岬 一尋":    "一般職",
        "影山 知紀":    "一般職",
        "水分 香織":    "一般職",
        "坂東 和哉":    "一般職",
        "藤本 宏治":    "一般職",
        "山中 里紗":    "一般職",
        "石原 直樹":    "一般職",
        "井口 大昌":    "一般職",
        "上野 勇輝":    "一般職",
        "西澤 裕貴":    "一般職",
        "小幡 裕亮":    "一般職",
        "清水 俊貴":    "一般職",
        "黒田 洋平":    "一般職",
        "大塲 智徳":    "一般職",
        "山本 一成":    "一般職",   # 部付(嘱託)
        "伊藤 孝弘":    "一般職",   # 嘱託
        "勝又 亮":      "一般職",
        "木島 早希":    "一般職",
        "福永 康平":    "一般職",
    })

    # --- 名前エイリアス (カタカナ/漢字ゆれの手動辞書) ---
    name_aliases: dict = field(default_factory=lambda: {"張学シン": "張学鑫"})

    # --- 地名エイリアス (カタカナ↔ラテン社名ゆれ; 誤突合防止) ---
    # 例: CSV 'シムテック' は顧客マスタ 'SIMMTECH GRAPHICS' のカタカナ表記.
    # 部分文字列置換で適用するため 'シムテック中大塩' も SIMMTECH へ寄せられる.
    place_aliases: dict = field(default_factory=lambda: {
        "シムテック": "SIMMTECH GRAPHICS",
        "OKI": "OKIサーキットテクノロジー",
        "新光電気": "新光電気工業",
        "新光": "新光電気工業",
        "大昌": "大昌電子",
        "ケイツー": "ケイツープリント",
        # 2026-07-31 客先指摘: 'DMS' はマスタ側で 'ディーエムエス' とカナ表記のため
        # 部分一致でも突合できず未突合になっていた.
        "東芝DMS": "東芝ディーエムエス",
    })

    # --- 既知の欠落 (常に出力にバナー表示) ---
    known_gaps: list = field(default_factory=lambda: [])

    def has_amount_rules(self) -> bool:
        """金額規程の実値が設定されているか."""
        return bool(self.amount_limits)

    def has_receipt_rule(self) -> bool:
        return self.receipt_required_above is not None


def load_config(path: str | None = None) -> Config:
    """設定をロード. path 指定時は JSON で上書き (将来拡張用)."""
    cfg = Config()
    if path and os.path.exists(path):
        import json
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        for k, v in data.items():
            if hasattr(cfg, k):
                setattr(cfg, k, v)
    return cfg
