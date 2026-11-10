import json
import os
import random
import time
import urllib.parse
import urllib.request
import uuid
from pathlib import Path

BASE_API = "https://ttvip.sj88ai.com/final/api"
OUT_DIR = Path(__file__).resolve().parent
POOL_FILE = OUT_DIR / "kalodata_pool.json"
STATE_FILE = OUT_DIR / ".kalodata_sync_state.json"

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0",
]


def api_get(path: str) -> dict | list:
    req_id = f"catalogeasy-{uuid.uuid4()}"
    url = f"{BASE_API}{path}"
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "application/json",
            "Referer": "https://ttvip.sj88ai.com/final/catalog_easy.html",
            "X-Request-ID": req_id,
        },
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def extract_date_range(kalodata_url: str) -> str:
    if not kalodata_url:
        return ""
    try:
        parsed = urllib.parse.urlparse(kalodata_url)
        qs = urllib.parse.parse_qs(parsed.query)
        dr = qs.get("dateRange", [""])[0]
        arr = json.loads(dr)
        if isinstance(arr, list) and len(arr) == 2:
            return f"{arr[0]} ถึง {arr[1]}"
    except Exception:
        pass
    return ""


def slim_product(p: dict) -> dict | None:
    sku = str(p.get("tiktok_product_id") or "").strip()
    name = str(p.get("name") or "").strip()
    image = str(p.get("image") or "").strip()
    if not sku or not name or not image:
        return None

    rev = float(p.get("revenue") or 0)
    v_rev = float(p.get("video_rev") or 0)
    l_rev = float(p.get("live_rev") or 0)
    c_rev = float(p.get("card_rev") or 0)
    sales = float(p.get("sales") or 0)
    creators = float(p.get("creators") or 0)
    conv = float(p.get("conv") or 0)
    comm = float(p.get("comm") or 0)
    growth = float(p.get("gmv_growth") or 0)
    ai_score = float(p.get("ai_score") or 0)
    price = p.get("price")
    avg_price = float(p.get("avg_price") or price or 0)
    price_raw = str(p.get("price_raw") or price or avg_price or "")

    return {
        "id": str(p.get("id") or ""),
        "sku": sku,
        "name": name,
        "image": image,
        "cat": str(p.get("cat_top") or ""),
        "cat_full": str(p.get("cat_full") or ""),
        "price": round(float(price), 2) if price is not None else round(avg_price, 2),
        "price_raw": price_raw,
        "avg_price": round(avg_price, 2),
        "rev": round(rev, 2),
        "v_rev": round(v_rev, 2),
        "l_rev": round(l_rev, 2),
        "c_rev": round(c_rev, 2),
        "sales": int(round(sales)),
        "creators": int(round(creators)),
        "conv": round(conv, 2),
        "comm": round(comm, 2),
        "growth": round(growth, 1),
        "ai": int(round(ai_score)),
        "score": float(p.get("score") or 0),
        "vdo": int(p.get("vdo_friendly") or 0),
        "hcomm": int(p.get("high_comm") or 0),
        "fship": int(p.get("free_ship") or 0),
        "link": f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH",
    }


def main(force: bool = False):
    print("🔍 กำลังเช็กสถานะฐานข้อมูล Kalodata ต้นทาง (Stealth Heartbeat)...")
    stats = api_get("/products/stats")
    total_products = stats.get("total", 0)
    total_revenue = stats.get("total_revenue", 0)

    if STATE_FILE.exists() and POOL_FILE.exists() and not force:
        try:
            prev = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            if prev.get("total_revenue") == total_revenue and prev.get("total") == total_products:
                print(f"✅ ข้อมูลเป็นชุดล่าสุดอยู่แล้ว (สินค้า {total_products:,} รายการ) ไม่ต้องดึงซ้ำ")
                return
        except Exception:
            pass

    cats_raw = api_get("/products/categories")
    categories = [c for c in cats_raw if c.get("count", 0) >= 5000]

    by_sku: dict[str, dict] = {}
    sample_date_range = ""

    # 1. ดึงภาพรวมทุกหมวด (Overall Sorts)
    global_queries = [
        "vdo=1&sort=revenue&limit=100",
        "vdo=1&sort=ai&limit=100",
        "vdo=1&sort=creators&rev_min=100000&limit=100",
        "vdo=1&hcomm=1&sort=revenue&limit=100",
        "vdo=1&sort=growth&rev_min=100000&limit=100",
        "vdo=1&sort=conv&rev_min=100000&limit=100",
        "sort=revenue&limit=80",
    ]
    for q in global_queries:
        try:
            res = api_get(f"/products/search?{q}")
            for p in res.get("products", []):
                if not sample_date_range and p.get("kalodata"):
                    sample_date_range = extract_date_range(p["kalodata"])
                sp = slim_product(p)
                if sp and sp["v_rev"] >= 15000:
                    by_sku[sp["sku"]] = sp
            time.sleep(random.uniform(0.4, 0.8))
        except Exception as e:
            print(f"[!] ข้าม query {q}: {e}")

    # 2. ดึงแยกตามหมวดหมู่หลัก เพื่อให้เวลาเลือกหมวดหมู่ในเว็บมีครบทั้ง 3 สูตร
    for c in categories[:22]:
        cat_name = c["name"]
        cat_enc = urllib.parse.quote(cat_name)
        for q in [
            f"cat={cat_enc}&vdo=1&sort=revenue&limit=45",
            f"cat={cat_enc}&vdo=1&sort=ai&limit=35",
            f"cat={cat_enc}&vdo=1&hcomm=1&sort=revenue&limit=30",
        ]:
            try:
                res = api_get(f"/products/search?{q}")
                for p in res.get("products", []):
                    sp = slim_product(p)
                    if sp and sp["v_rev"] >= 10000:
                        by_sku[sp["sku"]] = sp
                time.sleep(random.uniform(0.25, 0.55))
            except Exception as e:
                print(f"[!] ข้ามหมวด {cat_name}: {e}")

    products_list = list(by_sku.values())
    cat_summary = []
    cat_counts = {}
    for item in products_list:
        cat_counts[item["cat"]] = cat_counts.get(item["cat"], 0) + 1
    for c in cats_raw:
        if c["name"] in cat_counts:
            cat_summary.append({
                "name": c["name"],
                "pool_count": cat_counts[c["name"]],
                "total_db_count": c["count"],
            })

    payload = {
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "kalodata_date_range": sample_date_range or "30 วันล่าสุด",
        "source_total_products": total_products,
        "source_total_revenue": total_revenue,
        "pool_size": len(products_list),
        "categories": cat_summary,
        "products": products_list,
    }

    POOL_FILE.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    STATE_FILE.write_text(
        json.dumps({"total": total_products, "total_revenue": total_revenue, "updated_at": payload["updated_at"]}, indent=2),
        encoding="utf-8",
    )
    print(f"🎉 บันทึกคลังสินค้าขายดี Kalodata สำเร็จ: {len(products_list):,} รายการ ({POOL_FILE.stat().st_size / 1024:.1f} KB)")

    if "--auto-push" in os.sys.argv:
        import subprocess
        try:
            subprocess.run(["git", "-C", str(OUT_DIR), "add", "kalodata_pool.json"], check=False)
            subprocess.run(["git", "-C", str(OUT_DIR), "commit", "-m", f"chore: auto-sync Kalodata pool ({payload['updated_at']})"], check=False)
            subprocess.run(["git", "-C", str(OUT_DIR), "push"], check=False)
            print("🚀 อัปเดตข้อมูล Kalodata ขึ้น GitHub Pages เรียบร้อยแล้ว!")
        except Exception as e:
            print(f"[!] ไม่สามารถ git push อัตโนมัติได้: {e}")


if __name__ == "__main__":
    import sys
    main(force="--force" in sys.argv)
