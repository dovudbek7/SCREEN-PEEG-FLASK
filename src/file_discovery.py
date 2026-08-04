"""フォルダ内のファイルをパターンで検索するユーティリティ.

運用担当者が expenses/ attendance/ 等のフォルダにファイルを配置するだけで
config.py がファイル名を意識せずに済むようにするための共通ヘルパー.
"""
from __future__ import annotations

import glob
import os
import re


def _sort_key(path: str) -> tuple[str, float]:
    """ファイル名末尾の数字列(タイムスタンプ)を優先して比較するキー.

    楽々精算/楽々勤怠のエクスポートファイル名は '..._YYYYMMDD_HHMMSS' 等の
    数字列を含み, 文字列としてソートしても時系列順になる。ファイルシステムの
    更新日時(mtime)はコピー/チェックアウト操作で意図せず変わることがあるため,
    数字列が取れる場合はそちらを優先し, 取れない場合のみ mtime にフォールバックする。
    """
    name = os.path.basename(path)
    digits = "".join(re.findall(r"\d+", name))
    return (digits, os.path.getmtime(path))


def find_matching(folder: str, pattern: str) -> list[str]:
    """folder 内で pattern (glob) に一致するファイルを古い→新しい順で返す."""
    if not os.path.isdir(folder):
        return []
    paths = [p for p in glob.glob(os.path.join(folder, pattern)) if os.path.isfile(p)]
    paths.sort(key=_sort_key)
    return paths


def latest_matching(folder: str, pattern: str) -> str:
    """folder 内で pattern に一致する最新の1件を返す. 無ければ例外."""
    matches = find_matching(folder, pattern)
    if not matches:
        raise FileNotFoundError(
            f"「{folder}」に「{pattern}」に一致するファイルが見つかりません。"
            f"必要なファイルをこのフォルダに配置してください。"
        )
    return matches[-1]


def latest_matching_recursive(folder: str, pattern: str) -> str | None:
    """folder 以下 (サブフォルダ含む) で pattern に一致する最新の1件を返す.

    手当マスタ等, 日付付きサブフォルダに格納される参照データ用. 見つからない場合は
    例外にせず None を返す (旧データセットに存在しないケースがあるため任意項目扱い).
    """
    if not os.path.isdir(folder):
        return None
    paths = [p for p in glob.glob(os.path.join(folder, "**", pattern), recursive=True) if os.path.isfile(p)]
    if not paths:
        return None
    paths.sort(key=_sort_key)
    return paths[-1]


def all_matching(folder: str, pattern: str) -> list[str]:
    """folder 内で pattern に一致する全ファイルを古い順で返す. 無ければ例外."""
    matches = find_matching(folder, pattern)
    if not matches:
        raise FileNotFoundError(
            f"「{folder}」に「{pattern}」に一致するファイルが見つかりません。"
            f"必要なファイルをこのフォルダに配置してください。"
        )
    return matches
