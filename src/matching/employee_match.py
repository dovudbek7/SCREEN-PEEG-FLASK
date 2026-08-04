"""従業員照合 (employee resolution).

入力者名 (入力者名) を従業員マスタの 氏名 と突合し 社員番号 を引き当てる.
照合キーは norm(氏名). 完全一致 → 別名辞書 → difflib ファジー の順に試行する.

依存: 標準ライブラリのみ (difflib). 名前正規化は normalize.norm に一元化.
"""
from __future__ import annotations

import difflib

from models import Employee, ResolveResult
from normalize import norm
from config import Config


def build_employee_index(emps: list[Employee]) -> dict[str, Employee]:
    """従業員リストから norm(氏名) -> Employee の索引を構築.

    氏名衝突時は先勝ち (マスタ先頭行を優先). norm が空のものは除外.
    """
    idx: dict[str, Employee] = {}
    for e in emps:
        key = e.name_norm or norm(e.name_raw)
        if not key:
            continue
        idx.setdefault(key, e)
    return idx


def resolve_employee(name: str, idx: dict[str, Employee], cfg: Config) -> ResolveResult:
    """入力者名を従業員マスタへ突合し ResolveResult を返す.

    手順:
      1. key = norm(name). 元名が cfg.name_aliases にあれば別名へ差し替え.
      2. 裸の整数名はレガシー勤怠ID (マスタIDではない) → 未突合.
      3. 完全一致 → '突合' (score=1.0).
      4. difflib ファジー: 最良比率 >= 閾値 → '別名突合', それ未満 → '未突合'.
    """
    # --- 別名適用 (正規化キーで辞書を引く) ---
    # 入力者名は全角空白を含むことがある ('張　学シン') ため,
    # 別名辞書も正規化してから突合する (空白差で別名が外れるのを防ぐ).
    key = norm(name)
    alias_norm = {norm(k): norm(v) for k, v in cfg.name_aliases.items()}
    if key in alias_norm:
        key = alias_norm[key]

    # --- 裸の整数名はレガシー勤怠ID. マスタへ整数結合はしない ---
    if key.isdigit():
        return ResolveResult(None, "未突合")

    # --- 完全一致 ---
    e = idx.get(key)
    if e is not None:
        return ResolveResult(e.employee_id, "突合", matched_name=e.name_raw, score=1.0)

    # --- ファジー (difflib SequenceMatcher 比率) ---
    best_key: str | None = None
    best_score = 0.0
    sm = difflib.SequenceMatcher()
    sm.set_seq2(key)
    for cand in idx.keys():
        sm.set_seq1(cand)
        r = sm.ratio()
        if r > best_score:
            best_score = r
            best_key = cand

    if best_key is not None and best_score >= cfg.fuzzy_name_threshold:
        e = idx[best_key]
        return ResolveResult(
            e.employee_id, "別名突合", matched_name=e.name_raw, score=best_score
        )

    return ResolveResult(None, "未突合", score=best_score)
