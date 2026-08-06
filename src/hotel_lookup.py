"""ホテル名から東京23区内/23区外を判定するルックアップ (楽天トラベル キーワード検索API).

領収書からOCRで取得したホテル名を検索し, 住所 (address1/address2) に
東京23区のいずれかの区名が含まれるかで判定する. 23区の一覧は固定・既知のため
外部データの一括ダウンロードは不要 (2026-07-08 検討結果).
"""
from __future__ import annotations

import urllib.parse
import urllib.request
import urllib.error
import json

from secrets_local import RAKUTEN_APPLICATION_ID, RAKUTEN_ACCESS_KEY

_API_URL = "https://openapi.rakuten.co.jp/engine/api/Travel/KeywordHotelSearch/20170426"

TOKYO_23_WARDS = (
    "千代田区", "中央区", "港区", "新宿区", "文京区", "台東区", "墨田区", "江東区",
    "品川区", "目黒区", "大田区", "世田谷区", "渋谷区", "中野区", "杉並区", "豊島区",
    "北区", "荒川区", "板橋区", "練馬区", "足立区", "葛飾区", "江戸川区",
)


def search_hotel(keyword: str) -> list[dict]:
    """キーワードでホテルを検索し, ヒットした施設情報のリストを返す."""
    params = {
        "applicationId": RAKUTEN_APPLICATION_ID,
        "accessKey": RAKUTEN_ACCESS_KEY,
        "keyword": keyword,
        "format": "json",
        "responseType": "large",
        "hits": 10,
    }
    url = _API_URL + "?" + urllib.parse.urlencode(params)
    try:
        with urllib.request.urlopen(url, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return []  # ヒット無し
        raise
    if "error" in data:
        raise RuntimeError(f"Rakuten API error: {data.get('error')}: {data.get('error_description')}")
    hotels = []
    for item in data.get("hotels", []):
        basic = item["hotel"][0]["hotelBasicInfo"]
        hotels.append(basic)
    return hotels


def classify_ward(hotel_name: str) -> tuple[str | None, str | None]:
    """ホテル名から (区名 or None, 判定元の住所) を返す.

    複数候補がヒットした場合は先頭 (関連度が最も高い) を採用する.
    ヒット無し, または住所に23区名が含まれない場合は (None, address).
    """
    hotels = search_hotel(hotel_name)
    if not hotels:
        return None, None
    basic = hotels[0]
    address = f"{basic.get('address1', '')}{basic.get('address2', '')}"
    for ward in TOKYO_23_WARDS:
        if ward in address:
            return ward, address
    return None, address


if __name__ == "__main__":
    import sys
    name = sys.argv[1] if len(sys.argv) > 1 else "ベルズイン土浦"
    ward, address = classify_ward(name)
    print(f"検索語: {name}")
    print(f"住所: {address}")
    print(f"東京23区: {ward or '該当なし (23区外 or 東京都外)'}")
