"""正規化・パース関数 (normalization & parsing helpers).

すべての名前比較・地名照合・金額/日付パースの単一の真実の源.
依存: 標準ライブラリのみ (rapidfuzz は matching 層でのみ使用).
"""
from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime, time

# ---------------------------------------------------------------------------
# 名前正規化 (name normalization) — 全名前比較はこれを通す
# ---------------------------------------------------------------------------




def norm(s) -> str:
    """NFKC 正規化 + 全空白除去. 名前・キーの正準形.

    例: '土田　光樹' -> '土田光樹', '松沢 響' -> '松沢響'.
    """
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    # NFKC は全角空白 U+3000 を半角空白に変換するので両方除去
    return s.replace(" ", "").replace("　", "").strip()


# 確認印などの接尾辞を除去して承認者名を正準化
_CONFIRM_SUFFIX = re.compile(r"[（(]\s*確認\s*[）)]\s*$")


def norm_approver(s) -> tuple[str, bool]:
    """承認者名を正規化. (正規化名, 確認のみフラグ) を返す.

    '高橋昭太（確認）' -> ('高橋昭太', True)
    """
    if s is None:
        return "", False
    raw = unicodedata.normalize("NFKC", str(s))
    # 確認のみ判定は『(確認)』接尾辞に限定する.
    # ('確認' 部分一致は氏名等を誤検知するため使わない — 過剰検知防止)
    is_confirm = bool(_CONFIRM_SUFFIX.search(raw))
    raw = _CONFIRM_SUFFIX.sub("", raw)
    return norm(raw), is_confirm


# ---------------------------------------------------------------------------
# 地名正規化 (place-name normalization) — 顧客マスタ照合用
# ---------------------------------------------------------------------------

_BRACKET = re.compile(r"[\[\(（【].*?[\]\)）】]")
_CORP_SUFFIXES = (
    "株式会社", "(株)", "（株）", "有限会社", "合同会社", "合資会社",
    "公益財団法人", "一般財団法人", "独立行政法人", "国立大学法人", "株式",
)


def normalize_place(s) -> tuple[str, str | None]:
    """地名/取引先名を (ベース, 拠点トークン) に正規化.

    'リンクステック株式会社[下館工場]' -> ('リンクステック', '下館工場')
    'リンクステック石岡' -> ('リンクステック石岡', None)  (CSV 側は括弧無し)
    """
    if s is None:
        return "", None
    s = unicodedata.normalize("NFKC", str(s)).strip()
    # 拠点トークン抽出 (最初の括弧の中身)
    m = _BRACKET.search(s)
    site_token = m.group(0)[1:-1] if m else None
    if site_token is not None:
        site_token = site_token.strip() or None
    # 括弧内容を除去
    base = _BRACKET.sub("", s)
    for suf in _CORP_SUFFIXES:
        base = base.replace(suf, "")
    base = base.replace(" ", "").replace("　", "").strip().lower()
    return base, site_token


# ---------------------------------------------------------------------------
# 距離区分パース (distance band parsing)
# ---------------------------------------------------------------------------

_BAND_RE = re.compile(r"^(\d+)\s*~\s*(\d*)$")


def parse_band(b) -> tuple[int, int | None]:
    """'150~200 km' -> (150, 200);  '500~ km' -> (500, None).

    パース不能時は (0, None) ではなく ValueError を投げず (-1, None) を返し
    呼び出し側で未知バンドとして扱えるようにする.
    """
    if b is None:
        return -1, None
    s = (
        str(b)
        .replace("km", "")
        .replace("ｋｍ", "")
        .replace("～", "~")
        .replace("〜", "~")
        .replace("－", "~")
        .strip()
    )
    m = _BAND_RE.match(s)
    if not m:
        return -1, None
    lo = int(m.group(1))
    hi = int(m.group(2)) if m.group(2) else None
    return lo, hi


# ---------------------------------------------------------------------------
# 金額パース (money parsing)
# ---------------------------------------------------------------------------

def parse_money(s) -> int:
    """カンマ/空白/円記号を除去して int に. 空/不能は 0."""
    if s is None or s == "":
        return 0
    if isinstance(s, (int, float)):
        return int(s)
    t = unicodedata.normalize("NFKC", str(s)).strip()
    t = t.replace(",", "").replace(" ", "").replace("¥", "").replace("円", "")
    if t in ("", "-"):
        return 0
    try:
        return int(float(t))
    except ValueError:
        return 0


# ---------------------------------------------------------------------------
# 日付・時刻パース (date / time parsing)
# ---------------------------------------------------------------------------

_DATE_SEPS = re.compile(r"[/\-．\.]")


def parse_jp_date(s) -> date | None:
    """'YYYY/MM/DD' / 'YYYY-MM-DD' / Excel datetime を date に."""
    if s is None or s == "":
        return None
    if isinstance(s, datetime):
        return s.date()
    if isinstance(s, date):
        return s
    t = unicodedata.normalize("NFKC", str(s)).strip()
    if not t:
        return None
    parts = _DATE_SEPS.split(t.split()[0])
    try:
        if len(parts) >= 3:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            return date(y, m, d)
    except (ValueError, IndexError):
        return None
    return None


def normalize_time_str(s) -> str:
    """時刻文字列を正規化: 全角->半角(NFKC) + 区切り文字(';'含む)を':'に統一 + 空白除去."""
    if s is None:
        return ""
    t = unicodedata.normalize("NFKC", str(s))
    t = t.replace(";", ":").replace(" ", "").replace("　", "")
    return t.strip()


def parse_time(s) -> time | None:
    """'HH:MM:SS' / 'HH:MM' / 'HH;MM' / 全角混在 / datetime を time に."""
    if s is None or s == "":
        return None
    if isinstance(s, datetime):
        return s.time()
    if isinstance(s, time):
        return s
    t = normalize_time_str(s)
    if not t:
        return None
    parts = t.split(":")
    try:
        h = int(parts[0])
        mi = int(parts[1]) if len(parts) > 1 else 0
        se = int(parts[2]) if len(parts) > 2 else 0
        # 24:00 等の翌日表記を黙って 00:00 に丸めると順序が壊れるため None で表面化.
        if not (0 <= h < 24 and 0 <= mi < 60 and 0 <= se < 60):
            return None
        return time(h, mi, se)
    except (ValueError, IndexError):
        return None


def validate_time_input(s) -> tuple[time | None, str]:
    """時刻入力を検証し (パース結果, ステータス) を返す.

    ステータス: 'empty' (未入力), 'ok' (正常パース), 'invalid' (入力ルール外 = 数字以外/範囲外).
    'invalid' は None を返す点は 'empty' と同じだが, 呼び出し側が両者を区別して
    (例: 取込ログに警告行を残す) 検知できるようにするためのもの.
    """
    if s is None or s == "":
        return None, "empty"
    if isinstance(s, (datetime, time)):
        return parse_time(s), "ok"
    t = normalize_time_str(s)
    if not t:
        return None, "empty"
    parsed = parse_time(t)
    return (parsed, "ok") if parsed is not None else (None, "invalid")


def parse_excel_dt(v) -> datetime | None:
    """Excel セル値 (datetime / 'YYYY-MM-DD HH:MM' 文字列) を datetime に."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    t = unicodedata.normalize("NFKC", str(v)).strip()
    if not t:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M"):
        try:
            return datetime.strptime(t, fmt)
        except ValueError:
            continue
    d = parse_jp_date(t)
    return datetime(d.year, d.month, d.day) if d else None
