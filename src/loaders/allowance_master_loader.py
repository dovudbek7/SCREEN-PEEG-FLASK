"""手当マスタ (手当CD -> 公式金額) ローダ.

楽々精算の手当マスタCSVを読み込み, 手当種別(手当1/手当2/手当3/手当4)ごとに
手当コード -> 手当金額 の辞書を返す。手当金額が0の行は「計算式入力フラグ」等で
自動計算される項目 (例: 滞在費補助の多くのコード) であり, 固定金額としては
使えないため呼び出し側で判別すること。
"""
from __future__ import annotations

import csv


def load_allowance_master(path: str) -> dict[str, dict[str, int]]:
    """手当種別 (例 '手当1') -> {手当コード: 手当金額} の辞書を返す."""
    result: dict[str, dict[str, int]] = {}
    with open(path, encoding="cp932", newline="") as f:
        for row in csv.DictReader(f):
            kind = row["手当種別(名称)"].strip()
            code = row["手当コード"].strip()
            try:
                amount = int(row["手当金額"])
            except (ValueError, KeyError):
                continue
            result.setdefault(kind, {})[code] = amount
    return result


def load_allowance_per_day_codes(path: str) -> set[str]:
    """1日あたりの単価として扱う手当コードの集合を返す.

    手当マスタの「計算式入力フラグ」が 1 のコードは, 手当金額が
    1日あたりの単価であり, 出張日数を掛けた額が実際の支給額になる
    (例: 手当1 の 002「日当(連続)」= 1,700円/日 → 3日間で 5,100円)。
    フラグが 0 のコードは固定額。

    2026-08-06 客先指摘:「日当は1日で1700円、2日で3400円だが反映されていない」
    への対応。従来は単価をそのまま日当金額として表示していた。
    """
    codes: set[str] = set()
    with open(path, encoding="cp932", newline="") as f:
        for row in csv.DictReader(f):
            if (row.get("計算式入力フラグ") or "").strip() == "1":
                codes.add(row["手当コード"].strip())
    return codes
