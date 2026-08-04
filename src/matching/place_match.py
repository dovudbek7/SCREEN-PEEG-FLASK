"""到着地(地名) → 顧客マスタ 照合 (place → customer matching).

CSV の到着地テキストを顧客マスタの取引先名へ突合し, 距離区分(km)を得る.
照合は normalize_place() の (ベース, 拠点トークン) を起点に, まず完全一致,
次に rapidfuzz の partial_ratio による別名(あいまい)突合を試みる.

誤突合対策 (重要):
  - partial_ratio は短いクエリが長い候補名の部分文字列だと 100 を返すため,
    '東京'→'竹田東京プロセスサービス' のような地名→無関係社名の誤突合が起きる.
    → 長さ比ガード(クエリが候補名の一部に過ぎない場合は別名突合を認めない).
  - 同一ベースで距離区分が異なる候補が複数あり拠点トークンで一意化できない場合は
    1件に確定せず status='複数候補' を返す(§6.1 赤; 03_差異一覧で人が確認).
  - 拠点/残差トークン(例 'メイコー石巻' の '石巻')は tie-break だけでなく
    候補 SELECTION に使い, スコアが低くても該当拠点を拾えるようにする.

なお移動レッグ(電車･ﾊﾞｽ/車 等)の到着地は経由地であり訪問先(顧客)ではないため,
呼び出し側(checksheet.enrich)で移動レッグは place 照合をスキップする.

依存: rapidfuzz (この層でのみ使用), 標準ライブラリ.
"""
from __future__ import annotations

from collections import defaultdict

from rapidfuzz import fuzz, process

from models import Customer, PlaceMatch, UNMATCHED, MULTI
from normalize import normalize_place
from config import Config

# 突合ステータス文字列 (models の語彙には無いマッチ種別ラベル)
_EXACT = "突合"
_ALIAS = "別名突合"

# 別名突合で「同点近傍」とみなすスコア差.
_TIE_MARGIN = 3.0
# 長さ比ガード: クエリベースが候補ベースのこの割合未満なら部分文字列誤突合とみなす.
_MIN_LEN_RATIO = 0.6
# 残差/拠点トークンによる広域 SELECTION で候補に含める最低 partial_ratio.
_WIDE_FLOOR = 55.0


def _clean(s: str | None) -> str:
    return (s or "").replace(" ", "").replace("　", "")


class CustomerIndex:
    """顧客集合と照合用の補助構造を保持する索引."""

    def __init__(self, custs: list[Customer]):
        self.customers: list[Customer] = list(custs)
        self.by_base: dict[str, list[Customer]] = defaultdict(list)
        for c in self.customers:
            if c.name_norm_base:
                self.by_base[c.name_norm_base].append(c)
        self.bases: list[str] = list(self.by_base.keys())


def build_customer_index(custs: list[Customer]) -> CustomerIndex:
    """Customer 一覧から照合索引を構築."""
    return CustomerIndex(custs)


def _has_token(c: Customer, tokens: list[str]) -> bool:
    """候補の 拠点トークン / 取引先名 / 都道府県 にいずれかのトークンが部分一致するか."""
    if not tokens:
        return False
    fields = [_clean(c.site_token), _clean(c.name_raw), _clean(c.prefecture)]
    for tok in tokens:
        t = _clean(tok)
        if not t:
            continue
        for f in fields:
            if f and (t in f or f in t):
                return True
    return False


def _bands(cands: list[Customer]) -> set:
    return {c.distance_band for c in cands}


def _hit(c: Customer, score: float, status: str) -> PlaceMatch:
    """Customer から距離区分付き PlaceMatch を生成."""
    return PlaceMatch(
        customer_no=c.customer_no,
        customer_name=c.name_raw,
        score=score,
        status=status,
        distance_band=c.distance_band,
        km_lower=c.km_lower,
        km_upper=c.km_upper,
        prefecture=c.prefecture,
    )


def _multi(cands: list[Customer], score: float) -> PlaceMatch:
    """複数候補(同点・別band)を 1件に確定せず返す."""
    seen, cand_list = set(), []
    for c in cands:
        key = (c.name_raw, c.distance_band)
        if key in seen:
            continue
        seen.add(key)
        cand_list.append((c.name_raw, c.distance_band))
    return PlaceMatch(
        customer_no=None, customer_name=None, score=score, status=MULTI,
        distance_band=None, km_lower=None, km_upper=None, candidates=cand_list,
    )


def _residual_tokens(query_base: str, cand_bases: list[str]) -> list[str]:
    """クエリベースから候補ベースを差し引いた残差を識別トークンとして抽出.

    'メイコー石巻' に対し候補ベース 'メイコー' があれば残差 '石巻' を返す.
    """
    out: list[str] = []
    for cb in cand_bases:
        if cb and cb != query_base and cb in query_base:
            resid = query_base.replace(cb, "")
            if resid and resid not in out:
                out.append(resid)
    return out


def _choose(cands: list[Customer], id_tokens: list[str], query_base: str,
            score: float, status: str, *, guard: bool) -> PlaceMatch:
    """候補集合から 1件を選ぶ. 一意化できなければ 複数候補/未突合 を返す.

    guard=True (あいまい突合) のときは長さ比ガードを適用し, 部分文字列だけの
    過大スコア(地名→無関係社名)を別名突合として確定しない.
    """
    # dedupe (取引先番号で)
    uniq: list[Customer] = []
    seen = set()
    for c in cands:
        if c.customer_no in seen:
            continue
        seen.add(c.customer_no)
        uniq.append(c)
    if not uniq:
        return PlaceMatch(None, None, score, UNMATCHED)

    # (1) 拠点/残差トークンで一意化できるか (最優先)
    token_hits = [c for c in uniq if _has_token(c, id_tokens)] if id_tokens else []
    if len(token_hits) == 1:
        return _hit(token_hits[0], score, status)
    if len(token_hits) > 1:
        if len(_bands(token_hits)) == 1:
            return _hit(token_hits[0], score, status)
        return _multi(token_hits, score)

    # (2) 部分文字列誤突合ガード (あいまい突合のみ):
    # クエリが全候補名の一部に過ぎず識別トークンも当たらない場合は確定しない
    # (例 '東京'/'大阪' が社名に含まれるだけの無関係社名への誤突合を防ぐ).
    if guard and all(
        len(query_base) < _MIN_LEN_RATIO * max(len(c.name_norm_base or ""), 1)
        for c in uniq
    ):
        return PlaceMatch(None, None, score, UNMATCHED)

    # (3) 単一候補 → 確定
    if len(uniq) == 1:
        return _hit(uniq[0], score, status)
    # (4) 距離区分が全候補で同一なら確定 (band 一意)
    if len(_bands(uniq)) == 1:
        return _hit(uniq[0], score, status)
    # (5) band が割れる & 識別不能 → 複数候補 (確定しない)
    return _multi(uniq, score)


def match_place(text: str, cidx: CustomerIndex, cfg: Config) -> PlaceMatch:
    """到着地テキストを顧客へ突合し PlaceMatch を返す."""
    if text is None or not str(text).strip():
        return PlaceMatch(None, None, None, UNMATCHED)

    base, site_token = normalize_place(text)
    if not base:
        return PlaceMatch(None, None, None, UNMATCHED)

    # --- 地名エイリアス適用 (カタカナ↔ラテン社名ゆれ) ---
    # 長いキー優先・1回適用で打ち切り (例 '新光電気'/'新光' の両方が登録されている場合,
    # 短い'新光'を後から再適用すると '新光電気工業' が二重展開されて壊れるのを防ぐ).
    for k, v in sorted(cfg.place_aliases.items(), key=lambda kv: -len(kv[0])):
        k_base = normalize_place(k)[0]
        if k_base and k_base in base:
            base = base.replace(k_base, normalize_place(v)[0])
            break

    id_tokens: list[str] = []
    if site_token:
        id_tokens.append(site_token)

    # --- (1) ベース完全一致 ---
    exact = cidx.by_base.get(base)
    if exact:
        return _choose(exact, id_tokens, base, 100.0, _EXACT, guard=False)

    if not cidx.bases:
        return PlaceMatch(None, None, None, UNMATCHED)

    # --- (2) あいまい突合 (partial_ratio) ---
    near = process.extract(base, cidx.bases, scorer=fuzz.partial_ratio, limit=15)
    if not near:
        return PlaceMatch(None, None, None, UNMATCHED)
    best_base, best_score, _ = near[0]
    if best_score < cfg.place_match_threshold:
        return PlaceMatch(None, None, best_score, UNMATCHED)

    # 残差トークン(例 '石巻')を識別トークンに追加
    id_tokens += _residual_tokens(base, [b for b, _, _ in near])

    # 候補プール: 同点近傍 + 残差/拠点トークンが当たる広域候補
    pool: list[Customer] = []
    for b, s, _ in near:
        cands_b = cidx.by_base.get(b, [])
        if s >= best_score - _TIE_MARGIN:
            pool += cands_b
        elif s >= _WIDE_FLOOR and id_tokens:
            # スコアは低いが残差トークンで該当拠点を拾える候補 (B2: メイコー石巻)
            pool += [c for c in cands_b if _has_token(c, id_tokens)]
    if not pool:
        pool = list(cidx.by_base.get(best_base, []))

    return _choose(pool, id_tokens, base, best_score, _ALIAS, guard=True)
