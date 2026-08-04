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
