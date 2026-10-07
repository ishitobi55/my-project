"""メルカリで評価件数が一定以上の出品者を抽出するスクリプト。

キーワード検索の結果に出てくる出品者のプロフィールを1件ずつ開き、
評価件数が閾値以上の出品者を CSV に書き出す。

使い方:
    python mercari_sellers.py "トレカ" --pages 3 --min-ratings 10000

サーバーへの負荷を避けるため、ページ移動ごとに数秒の待機を入れている。
"""

import argparse
import csv
import json
import random
import re
import time
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

SEARCH_URL = "https://jp.mercari.com/search?keyword={kw}&status=on_sale&page_token=v1:{page}"
PROFILE_URL = "https://jp.mercari.com/user/profile/{uid}"


def polite_sleep(min_s, max_s):
    time.sleep(random.uniform(min_s, max_s))


def collect_seller_ids(page, keyword, pages, min_s, max_s):
    """検索結果ページの API レスポンスから出品者 ID を集める。"""
    seller_ids = []
    seen = set()
    for n in range(pages):
        found = []
        url = SEARCH_URL.format(kw=quote(keyword), page=n)
        print(f"[検索] {n + 1}/{pages} ページ目: {url}")
        try:
            # 検索 API は読み込み完了後に遅れて呼ばれることがあるので、応答そのものを待つ
            with page.expect_response(lambda r: "entities:search" in r.url, timeout=60000) as info:
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
            data = info.value.json()
            for item in data.get("items", []):
                sid = item.get("sellerId")
                # メルカリShops の商品は通常のユーザープロフィールを持たないので除外
                if sid and not item.get("shopName"):
                    found.append(str(sid))
        except Exception as e:
            print(f"  検索結果を取得できませんでした: {e}")

        new = [s for s in found if s not in seen]
        seen.update(new)
        seller_ids.extend(new)
        print(f"  新規出品者 {len(new)} 人（累計 {len(seller_ids)} 人）")
        if not found:
            print("  商品が見つからないため検索を終了します")
            break
        polite_sleep(min_s, max_s)
    return seller_ids


def parse_profile_json(data):
    d = data.get("data", data)
    count = d.get("num_ratings")
    if count is None and isinstance(d.get("ratings"), dict):
        count = sum(v for v in d["ratings"].values() if isinstance(v, int))
    return d.get("name"), count


def parse_profile_dom(page):
    """API が取れなかった場合の予備: ページ本文から評価件数らしき数字を拾う。"""
    name = None
    try:
        name = page.locator("h1").first.inner_text(timeout=3000).strip()
    except Exception:
        pass
    text = page.inner_text("body")
    m = re.search(r"([\d,]+)\s*件?の?評価", text) or re.search(r"★[^\n]*?\(?([\d,]{1,9})\)?", text)
    count = int(m.group(1).replace(",", "")) if m else None
    return name, count


def fetch_profile(page, uid):
    data = None
    try:
        with page.expect_response(lambda r: "users/get_profile" in r.url, timeout=30000) as info:
            page.goto(PROFILE_URL.format(uid=uid), wait_until="domcontentloaded", timeout=60000)
        data = info.value.json()
    except Exception:
        pass

    if data:
        name, count = parse_profile_json(data)
        if count is not None:
            return name, count
    return parse_profile_dom(page)


def main():
    ap = argparse.ArgumentParser(description="メルカリの高評価件数出品者を抽出")
    ap.add_argument("keyword", help="検索キーワード（例: トレカ）")
    ap.add_argument("--pages", type=int, default=3, help="検索結果を何ページ分見るか（既定 3）")
    ap.add_argument("--min-ratings", type=int, default=10000, help="評価件数の下限（既定 10000）")
    ap.add_argument("--out", default="sellers.csv", help="出力 CSV（既定 sellers.csv）")
    ap.add_argument("--cache", default="profile_cache.json", help="確認済み出品者のキャッシュ")
    ap.add_argument("--min-wait", type=float, default=3.0, help="アクセス間隔の最小秒数")
    ap.add_argument("--max-wait", type=float, default=6.0, help="アクセス間隔の最大秒数")
    ap.add_argument("--browser", default=None, help="使う Chromium の実行ファイル（通常は指定不要）")
    ap.add_argument("--headless", action="store_true", help="ブラウザ画面を表示せずに実行")
    args = ap.parse_args()

    cache_path = Path(args.cache)
    cache = json.loads(cache_path.read_text(encoding="utf-8")) if cache_path.exists() else {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless, executable_path=args.browser)
        page = browser.new_context(locale="ja-JP").new_page()

        seller_ids = collect_seller_ids(page, args.keyword, args.pages, args.min_wait, args.max_wait)

        for i, uid in enumerate(seller_ids, 1):
            if uid in cache:
                continue
            name, count = fetch_profile(page, uid)
            cache[uid] = {"name": name, "num_ratings": count}
            cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
            mark = "★" if count is not None and count >= args.min_ratings else " "
            print(f"[{i}/{len(seller_ids)}] {mark} {name} 評価 {count}")
            polite_sleep(args.min_wait, args.max_wait)

        browser.close()

    hits = [
        (uid, v["name"], v["num_ratings"])
        for uid, v in cache.items()
        if uid in set(seller_ids) and v["num_ratings"] is not None and v["num_ratings"] >= args.min_ratings
    ]
    hits.sort(key=lambda r: r[2], reverse=True)
    with open(args.out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["出品者ID", "名前", "評価件数", "プロフィールURL"])
        for uid, name, count in hits:
            w.writerow([uid, name, count, PROFILE_URL.format(uid=uid)])

    unknown = sum(1 for uid in seller_ids if cache.get(uid, {}).get("num_ratings") is None)
    print(f"\n完了: {len(seller_ids)} 人中 {len(hits)} 人が評価 {args.min_ratings} 件以上 → {args.out}")
    if unknown:
        print(f"注意: {unknown} 人は評価件数を読み取れませんでした（{args.cache} を確認してください）")


if __name__ == "__main__":
    main()
