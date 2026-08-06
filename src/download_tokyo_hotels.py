"""東京23区内の全ホテル一覧をダウンロードし, ローカルCSVに保存する (一括取得スクリプト).

楽天トラベル SimpleHotelSearch API を smallClassCode=tokyo (23区内) の
detailClassCode (A〜I) ごとにページング取得する. QPS=1 で申請したため,
1リクエストごとに1秒スリープする. 出力先は git 管理外 (.gitignore 参照).
"""
from __future__ import annotations

import csv
import time
import urllib.parse
import urllib.request
import urllib.error
import json

from secrets_local import RAKUTEN_APPLICATION_ID, RAKUTEN_ACCESS_KEY

_API_URL = "https://openapi.rakuten.co.jp/engine/api/Travel/SimpleHotelSearch/20170426"
_HITS_PER_PAGE = 30
_OUT_PATH = "/private/tmp/claude-501/-Users-dovudbek-Documents-projects-mms-project/9ebd43c1-6209-4a4a-ad96-79f72477be01/scratchpad/tokyo23_hotels_cache.csv"

_DETAIL_CODES = {
    "A": "東京駅・銀座・秋葉原・東陽町・葛西",
    "B": "新橋・汐留・浜松町・お台場",
    "C": "赤坂・六本木・霞ヶ関・永田町",
    "D": "渋谷・恵比寿・目黒・二子玉川",
    "E": "品川・大井町・蒲田・羽田空港",
    "F": "新宿・中野・荻窪・四谷",
    "G": "池袋・赤羽・巣鴨・大塚",
    "H": "東京ドーム・飯田橋・御茶ノ水",
    "I": "上野・浅草・錦糸町・新小岩・北千住",
}


def _fetch_page(detail_code: str, page: int, retries: int = 4) -> dict:
    params = {
        "applicationId": RAKUTEN_APPLICATION_ID,
        "accessKey": RAKUTEN_ACCESS_KEY,
        "format": "json",
        "largeClassCode": "japan",
        "middleClassCode": "tokyo",
        "smallClassCode": "tokyo",
        "detailClassCode": detail_code,
        "hits": _HITS_PER_PAGE,
        "page": page,
        "responseType": "large",
    }
    url = _API_URL + "?" + urllib.parse.urlencode(params)
    last_err = None
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(url, timeout=15) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except (urllib.error.URLError, ConnectionError, TimeoutError) as e:
            last_err = e
            wait = 2 ** attempt  # 1,2,4,8秒 のバックオフ
            print(f"  [warn] {detail_code} p{page} 通信エラー ({e}), {wait}秒後リトライ ({attempt+1}/{retries})")
            time.sleep(wait)
    raise last_err


_FIELDNAMES = ["hotelNo", "hotelName", "address1", "address2", "postalCode", "areaCode", "areaName"]


def main() -> None:
    seen_hotel_no = set()
    total_saved = 0

    with open(_OUT_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=_FIELDNAMES)
        writer.writeheader()

        for code, area_name in _DETAIL_CODES.items():
            try:
                data = _fetch_page(code, 1)
            except Exception as e:  # noqa: BLE001
                print(f"[{code}] 取得失敗, スキップ: {e}")
                continue
            time.sleep(1)
            paging = data.get("pagingInfo", {})
            total_pages = paging.get("pageCount", 1)
            record_count = paging.get("recordCount", 0)
            print(f"[{code}] {area_name}: {record_count}件 ({total_pages}ページ)")

            page = 1
            while True:
                if page > 1:
                    try:
                        data = _fetch_page(code, page)
                    except Exception as e:  # noqa: BLE001
                        print(f"  [{code}] p{page} 取得失敗, この先スキップ: {e}")
                        break
                    time.sleep(1)

                new_rows = 0
                for item in data.get("hotels", []):
                    b = item["hotel"][0]["hotelBasicInfo"]
                    hotel_no = b.get("hotelNo")
                    if hotel_no in seen_hotel_no:
                        continue  # 複数エリアに跨る重複を除外
                    seen_hotel_no.add(hotel_no)
                    writer.writerow({
                        "hotelNo": hotel_no,
                        "hotelName": b.get("hotelName", ""),
                        "address1": b.get("address1", ""),
                        "address2": b.get("address2", ""),
                        "postalCode": b.get("postalCode", ""),
                        "areaCode": code,
                        "areaName": area_name,
                    })
                    new_rows += 1
                total_saved += new_rows
                f.flush()  # 途中終了しても保存済み分は失われないようにする

                if page >= total_pages:
                    break
                page += 1

    print(f"\n合計 {total_saved} 件のホテルを {_OUT_PATH} に保存しました。")


if __name__ == "__main__":
    main()
