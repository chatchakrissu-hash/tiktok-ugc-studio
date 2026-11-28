"""Local planner backend and static server for TikTok Monthly Content Planner."""

from __future__ import annotations

import base64
import concurrent.futures
import csv
import datetime as dt
import http.server
import io
import json
import math
import mimetypes
import os
import random
import re
import socketserver
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request
import uuid
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image

BASE_DIR = Path(__file__).resolve().parent
POOL_FILE = BASE_DIR / "kalodata_pool.json"
PLAN_FILE = BASE_DIR / "monthly_plan.json"
CHANNEL_CATALOG_FILE = BASE_DIR / "my_channel_catalog.json"
PLAN_ASSETS_DIR = BASE_DIR / "plan_assets"
PLAN_ASSETS_DIR.mkdir(parents=True, exist_ok=True)

KALODATA_API_BASE = "https://ttvip.sj88ai.com/final/api"
DEFAULT_PORT = 8765
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)

# Exact 16 styles synchronized 1:1 with index.html #styleChips
STYLES_16: list[dict[str, Any]] = [
    {"num": 1, "role": "ugc", "chip_match": "1. Real-Use Demo", "short": "แนว 1. Real-Use Demo (ใช้งานจริงให้เห็นผล)", "full": "แนว 1. Real-Use Demo (ใช้งานจริงให้เห็นผลชัด)"},
    {"num": 2, "role": "ugc", "chip_match": "2. Problem -> Solution", "short": "แนว 2. Problem → Solution (ปัญหา → ตัวช่วย)", "full": "แนว 2. Problem → Solution (ขยี้ปัญหา → จบด้วยตัวช่วย)"},
    {"num": 3, "role": "ugc", "chip_match": "3. Hands-Only Proof", "short": "แนว 3. Hands-Only Proof (เน้นมือสาธิต)", "full": "แนว 3. Hands-Only Proof (เห็นแต่มือสาธิตสินค้าชัดๆ)"},
    {"num": 4, "role": "ugc", "chip_match": "4. POV Experience", "short": "แนว 4. POV Experience (มุมมองผู้ใช้จริง)", "full": "แนว 4. POV Experience (มุมมองสายตาคนใช้จริง)"},
    {"num": 5, "role": "ugc", "chip_match": "5. Curiosity Gap", "short": "แนว 5. Curiosity Gap (ชวนสงสัยแล้วเฉลย)", "full": "แนว 5. Curiosity Gap (ชวนสงสัย ต้องดูจนจบ)"},
    {"num": 6, "role": "ugc", "chip_match": "6. Mistake & Discovery", "short": "แนว 6. Mistake & Discovery (พลาดแล้วเจอวิธีใหม่)", "full": "แนว 6. Mistake & Discovery (เพิ่งรู้ว่ามีวิธีที่ง่ายกว่านี้)"},
    {"num": 7, "role": "ugc", "chip_match": "7. Hidden Feature", "short": "แนว 7. Hidden Feature (เจาะดีเทลที่ไม่ค่อยรู้)", "full": "แนว 7. Hidden Feature (เจาะดีเทลลับ/วัสดุคุ้มเกินราคา)"},
    {"num": 8, "role": "ugc", "chip_match": "8. Before -> After Ease", "short": "แนว 8. Before → After Ease (ก่อนใช้ → หลังใช้ง่ายขึ้น)", "full": "แนว 8. Before → After Ease (ก่อนใช้ vs หลังใช้คล่องขึ้น)"},
    {"num": 9, "role": "ugc", "chip_match": "9. Expectation vs Reality", "short": "แนว 9. Expectation vs Reality (คาดหวังเทียบใช้จริง)", "full": "แนว 9. Expectation vs Reality (พิสูจน์ของจริงตรงปก)"},
    {"num": 10, "role": "ugc", "chip_match": "10. Product Hero", "short": "แนว 10. Product Hero / Aesthetic (โชว์สินค้าแต่งบ้าน)", "full": "แนว 10. Product Hero / Aesthetic (มินิมอล แต่งบ้านคาเฟ่)"},
    {"num": 11, "role": "ugc", "chip_match": "11. Banter", "short": "แนว 11. เพื่อนป้ายยาเพื่อน (เล่าสนุกเป็นกันเอง)", "full": "แนว 11. เพื่อนป้ายยาเพื่อน (เล่าสนุก เป็นกันเอง)"},
    {"num": 12, "role": "ugc", "chip_match": "12. Social Proof", "short": "แนว 12. Social Proof / ตอบกลับรีวิวจริง", "full": "แนว 12. Social Proof / ตอบกลับรีวิวลูกค้า"},
    {"num": 13, "role": "hardsell", "chip_match": "13. โกดัง/โรงงาน", "short": "แนว 13. โกดัง/โรงงานปล่อยของเอง (ฮาร์ดเซล)", "full": "แนว 13. โกดัง/โรงงานปล่อยของเอง (ฮาร์ดเซล ราคาหน้าโรงงาน)"},
    {"num": 14, "role": "hardsell", "chip_match": "14. พ่อค้าแม่ค้าหน้าร้าน", "short": "แนว 14. พ่อค้าแม่ค้าหน้าร้าน/ตลาด (ฮาร์ดเซล)", "full": "แนว 14. พ่อค้าแม่ค้าหน้าร้าน/ตลาดค้าส่ง (ฮาร์ดเซล ขายมันส์)"},
    {"num": 15, "role": "hardsell", "chip_match": "15. พนักงานขายในห้าง", "short": "แนว 15. พนักงานในห้าง/โชว์รูม (โปรห้างลดหนัก)", "full": "แนว 15. พนักงานขายในห้าง/โชว์รูม (ฮาร์ดเซล โปรห้างลดหนัก)"},
    {"num": 16, "role": "hardsell", "chip_match": "16. คนงานแพ็กของหน้าคลัง", "short": "แนว 16. คนงานแพ็กของหน้าคลัง (โชว์ของจริงก่อนส่ง)", "full": "แนว 16. คนงานแพ็กของหน้าคลัง (โชว์ของจริงก่อนส่ง + โปรราคาคลัง)"},
]
STYLE_BY_NUM: dict[int, dict[str, Any]] = {s["num"]: s for s in STYLES_16}
UGC_STYLE_NUMS = [s["num"] for s in STYLES_16 if s["role"] == "ugc"]
HARDSELL_STYLE_NUMS = [s["num"] for s in STYLES_16 if s["role"] == "hardsell"]

_CATALOG_LOCK = threading.Lock()
_POOL_CACHE: list[dict[str, Any]] = []
_POOL_BY_SKU: dict[str, dict[str, Any]] = {}
_POOL_BY_NORM_NAME: dict[str, dict[str, Any]] = {}
_REMOTE_QUERY_CACHE: dict[str, list[dict[str, Any]]] = {}


def normalize_thai_text(text: str) -> str:
    """Fix decomposed Sara Am and duplicate Thai combining marks."""
    s = str(text or "")
    s = s.replace("\u0E4D\u0E32", "\u0E33")
    s = re.sub(r"([\u0E31\u0E34-\u0E3A\u0E47-\u0E4E])\1+", r"\1", s)
    return s.strip()


def clean_product_query(name: str) -> str:
    """Strip noisy promo brackets, symbols, and prices for better DB matching."""
    s = normalize_thai_text(name)
    s = re.sub(r"\[[^\]]{1,40}\]|【[^】]{1,40}】|\([^\)]{1,35}\)", " ", s)
    s = re.sub(r"[■★☆🔥💥✨❤️✅🎁⭐🌟🚨]+", " ", s)
    s = re.sub(r"[-+]\s*\d+\s*%|฿\s*[\d,.]+", " ", s)
    s = s.replace('"', " ").replace("'", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def norm_token_key(text: str) -> str:
    """Produce a normalized comparison key for deduplication and fuzzy matching."""
    s = clean_product_query(text).lower()
    s = re.sub(r"[^0-9a-z\u0E00-\u0E7F\s]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def tokenize_name(text: str) -> set[str]:
    """Split product name into Thai/English/numeric tokens for similarity scoring."""
    norm = norm_token_key(text)
    raw_tokens = norm.split()
    tokens: set[str] = set()
    for t in raw_tokens:
        if len(t) >= 2:
            tokens.add(t)
        if len(t) >= 6:
            # Add character 3-grams for unspaced Thai phrases
            for i in range(len(t) - 2):
                tokens.add(t[i : i + 3])
    return tokens


def name_similarity(a: str, b: str) -> float:
    """Compute Jaccard + prefix similarity between two product names (0.0 .. 1.0)."""
    ka, kb = norm_token_key(a), norm_token_key(b)
    if not ka or not kb:
        return 0.0
    if ka == kb:
        return 1.0
    if ka in kb or kb in ka:
        shorter = min(len(ka), len(kb))
        longer = max(len(ka), len(kb))
        if shorter >= 10:
            return max(0.75, shorter / longer)
    ta, tb = tokenize_name(ka), tokenize_name(kb)
    if not ta or not tb:
        return 0.0
    inter = len(ta & tb)
    union = len(ta | tb)
    return inter / union if union else 0.0


ANGLE_CODES = ("A1", "A2", "A3")
ANGLE_TAG_RE = re.compile(r"\s*\[(?:มุม\s*)?(A[1-3])\]", re.IGNORECASE)


def extract_angle_tag(text: Any) -> str:
    """Read the marketing angle tag `[A1]` / `[A2]` / `[A3]` from a filename or name (empty when absent)."""
    m = ANGLE_TAG_RE.search(str(text or ""))
    return m.group(1).upper() if m else ""


def build_safe_video_filename(sku: str, name: str, angle_code: str = "") -> str:
    """Match index.html buildSafeVideoFilename(): `<19-digit-sku>__<clean_name>[ [A1]]`."""
    clean_sku = extract_clean_sku(sku)
    clean = normalize_thai_text(ANGLE_TAG_RE.sub(" ", str(name or "")))
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', " ", clean)
    clean = re.sub(r"\s+", " ", clean).strip()
    if len(clean) > 105:
        clean = clean[:105].strip()
    clean = clean or "คลิปสินค้า_TikTok"
    code = str(angle_code or "").strip().upper()
    if code in ANGLE_CODES:
        clean = f"{clean} [{code}]"
    return f"{clean_sku}__{clean}" if clean_sku else clean


def extract_clean_sku(raw: Any) -> str:
    """Extract a 15-22 digit TikTok Product SKU cleanly, rejecting scientific floats."""
    if raw is None:
        return ""
    s = str(raw).strip()
    if not s or s.lower() in {"nan", "none", "null", "-", "n/a"}:
        return ""
    # Reject spreadsheet float corruption like 1.7295962E+18
    if re.search(r"\d\.\d+\s*[eE][+-]?\d+", s):
        return ""
    # Prefer /product/<digits> URL pattern first
    m_url = re.search(r"/product/(\d{15,22})", s)
    if m_url:
        return m_url.group(1)
    # Strip trailing ".0" from numeric string formatting only if 15-22 digits remain intact
    s_no_dot = re.sub(r"\.0+$", "", s.lstrip("'").strip())
    m = re.search(r"(?<!\d)(\d{15,22})(?!\d)", s_no_dot)
    return m.group(1) if m else ""


def is_likely_float_corrupted_sku(raw: Any, sku: str) -> bool:
    """Detect whether an Excel export converted a 19-digit ID to float64 (e.g. 4+ trailing zeros or E+18)."""
    s = str(raw or "").strip()
    if re.search(r"\d\.\d+\s*[eE][+-]?\d+", s):
        return True
    if sku and len(sku) >= 18 and sku.endswith("0000"):
        return True
    return False


def unwrap_tiktok_link_or_sku(raw_input: str) -> dict[str, str]:
    """Resolve full/short TikTok URLs or plain SKUs into {sku, resolved_url, og_title, og_image}."""
    s = str(raw_input or "").strip()
    sku = extract_clean_sku(s)
    res = {"sku": sku, "resolved_url": s if s.startswith("http") else "", "og_title": "", "og_image": ""}
    if not s.startswith("http"):
        if sku:
            res["resolved_url"] = f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH"
        return res

    # If URL already contains the 19-digit SKU, normalize link and optionally skip heavy fetch unless needed
    if sku and "shop.tiktok.com/view/product/" in s:
        res["resolved_url"] = f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH"

    try:
        req = urllib.request.Request(
            s,
            headers={
                "User-Agent": USER_AGENT,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "th-TH,th;q=0.9,en;q=0.8",
            },
        )
        with urllib.request.urlopen(req, timeout=12) as resp:
            final_url = resp.geturl() or s
            res["resolved_url"] = final_url
            if not res["sku"]:
                res["sku"] = extract_clean_sku(final_url)
            raw_html = resp.read(262144).decode("utf-8", errors="ignore")
            if not res["sku"]:
                for pat in (
                    r'"product_id"\s*:\s*"(\d{15,22})"',
                    r'"productId"\s*:\s*"(\d{15,22})"',
                    r'/product/(\d{15,22})',
                    r'"id"\s*:\s*"(17\d{15,20})"',
                ):
                    m = re.search(pat, raw_html)
                    if m:
                        res["sku"] = m.group(1)
                        break
            m_title = re.search(r'<meta[^>]+property=["\']og:title["\'][^>]+content=["\']([^"\']+)["\']', raw_html, re.I)
            if m_title:
                res["og_title"] = m_title.group(1).strip()
            m_img = re.search(r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)["\']', raw_html, re.I)
            if m_img:
                res["og_image"] = m_img.group(1).strip()
    except Exception:
        pass

    if res["sku"] and not res["resolved_url"].startswith("https://shop.tiktok.com/view/product/"):
        res["resolved_url"] = f"https://shop.tiktok.com/view/product/{res['sku']}?region=TH&locale=th-TH"
    return res


def load_channel_catalog() -> dict[str, dict[str, Any]]:
    """Read persistent channel SKU/image dictionary from disk."""
    with _CATALOG_LOCK:
        if not CHANNEL_CATALOG_FILE.exists():
            return {}
        try:
            data = json.loads(CHANNEL_CATALOG_FILE.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}


def save_channel_catalog_entries(entries: list[dict[str, Any]]) -> None:
    """Persist resolved SKU -> {sku, name, image, cat, price, comm, link} mappings."""
    clean_entries = [e for e in entries if extract_clean_sku(e.get("sku"))]
    if not clean_entries:
        return
    with _CATALOG_LOCK:
        cur: dict[str, dict[str, Any]] = {}
        if CHANNEL_CATALOG_FILE.exists():
            try:
                loaded = json.loads(CHANNEL_CATALOG_FILE.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    cur = loaded
            except Exception:
                cur = {}
        now_str = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for e in clean_entries:
            sku = extract_clean_sku(e.get("sku"))
            if not sku:
                continue
            prev = cur.get(sku, {})
            cur[sku] = {
                "sku": sku,
                "name": str(e.get("name") or prev.get("name") or "").strip(),
                "image": str(e.get("image") or prev.get("image") or "").strip(),
                "cat": str(e.get("cat") or prev.get("cat") or "").strip(),
                "price": float(e.get("price") or prev.get("price") or 0),
                "comm": float(e.get("comm") or prev.get("comm") or 0),
                "link": str(e.get("link") or prev.get("link") or f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH"),
                "updated_at": now_str,
            }
        CHANNEL_CATALOG_FILE.write_text(json.dumps(cur, ensure_ascii=False, indent=2), encoding="utf-8")


def ensure_pool_loaded() -> list[dict[str, Any]]:
    """Load local kalodata_pool.json into memory indices once."""
    global _POOL_CACHE, _POOL_BY_SKU, _POOL_BY_NORM_NAME
    if _POOL_CACHE:
        return _POOL_CACHE
    if not POOL_FILE.exists():
        return []
    try:
        data = json.loads(POOL_FILE.read_text(encoding="utf-8"))
        prods = data.get("products", []) if isinstance(data, dict) else []
        _POOL_CACHE = prods
        _POOL_BY_SKU = {str(p.get("sku")): p for p in prods if p.get("sku")}
        _POOL_BY_NORM_NAME = {}
        for p in prods:
            nk = norm_token_key(p.get("name", ""))
            if nk and nk not in _POOL_BY_NORM_NAME:
                _POOL_BY_NORM_NAME[nk] = p
    except Exception:
        _POOL_CACHE = []
    return _POOL_CACHE


def remote_search_kalodata(query: str, limit: int = 8) -> list[dict[str, Any]]:
    """Query ttvip.sj88ai.com Kalodata search API with simple in-memory caching."""
    q = clean_product_query(query)[:48].strip()
    if len(q) < 2:
        return []
    if q in _REMOTE_QUERY_CACHE:
        return _REMOTE_QUERY_CACHE[q]

    url = f"{KALODATA_API_BASE}/products/search?q={urllib.parse.quote(q)}&limit={limit}"
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "application/json",
            "Referer": "https://ttvip.sj88ai.com/final/catalog_easy.html",
            "X-Request-ID": f"planner-{uuid.uuid4()}",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=18) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
            products = payload.get("products", []) if isinstance(payload, dict) else []
            _REMOTE_QUERY_CACHE[q] = products
            return products
    except Exception:
        return []


def progressive_search_queries(name: str) -> list[str]:
    """Generate up to 2 concise search queries (specific prefix, then core keywords)."""
    cleaned = clean_product_query(name)
    if not cleaned:
        return []
    words = cleaned.split()
    queries: list[str] = []
    if len(cleaned) <= 32:
        queries.append(cleaned)
    else:
        queries.append(cleaned[:32].rsplit(" ", 1)[0] or cleaned[:32])
    thai_words = [w for w in words if re.search(r"[\u0E00-\u0E7F]{2,}", w)]
    if len(thai_words) >= 2:
        q2 = " ".join(thai_words[:2])[:28].strip()
        if q2 and q2 not in queries:
            queries.append(q2)
    return queries[:2]


def _normalize_remote_product(p: dict[str, Any]) -> dict[str, Any]:
    """Convert raw Kalodata API item into studio-compatible metadata dict."""
    sku = extract_clean_sku(p.get("tiktok_product_id") or p.get("sku"))
    return {
        "sku": sku,
        "name": str(p.get("name") or "").strip(),
        "image": str(p.get("image") or "").strip(),
        "cat": str(p.get("cat_top") or p.get("cat") or "").strip(),
        "price": float(p.get("price") or p.get("avg_price") or 0),
        "comm": float(p.get("comm") or 0),
        "v_rev": float(p.get("video_rev") or p.get("v_rev") or 0),
        "rev": float(p.get("revenue") or p.get("rev") or 0),
        "creators": int(float(p.get("creators") or 0)),
        "conv": float(p.get("conv") or 0),
        "growth": float(p.get("gmv_growth") or p.get("growth") or 0),
        "ai": int(float(p.get("ai_score") or p.get("ai") or 0)),
        "link": f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH" if sku else "",
    }


def resolve_product_metadata(
    item: dict[str, Any],
    channel_catalog: dict[str, dict[str, Any]] | None = None,
    use_remote_api: bool = True,
) -> dict[str, Any]:
    """Resolve missing SKU, image, category, and price using local cache + Kalodata 2.98M DB."""
    ensure_pool_loaded()
    cat_map = channel_catalog if channel_catalog is not None else load_channel_catalog()

    res = dict(item)
    raw_sku = extract_clean_sku(res.get("sku"))
    raw_name = normalize_thai_text(res.get("name", ""))
    res["name"] = raw_name
    res["sku"] = raw_sku
    res.setdefault("image", "")
    res.setdefault("cat", "")
    res.setdefault("price", 0.0)
    res.setdefault("comm", 0.0)
    res.setdefault("candidates", [])

    # Step 1: Exact SKU lookup in channel catalog & local kalodata pool
    if raw_sku:
        res["link"] = res.get("link") or f"https://shop.tiktok.com/view/product/{raw_sku}?region=TH&locale=th-TH"
        if raw_sku in cat_map:
            saved = cat_map[raw_sku]
            res["image"] = res["image"] or saved.get("image", "")
            res["cat"] = res["cat"] or saved.get("cat", "")
            res["price"] = res["price"] or float(saved.get("price") or 0)
            res["comm"] = res["comm"] or float(saved.get("comm") or 0)
            res["name"] = res["name"] or saved.get("name", "")
        if raw_sku in _POOL_BY_SKU:
            kp = _POOL_BY_SKU[raw_sku]
            res["image"] = res["image"] or kp.get("image", "")
            res["cat"] = res["cat"] or kp.get("cat", "")
            res["price"] = res["price"] or float(kp.get("price") or kp.get("avg_price") or 0)
            res["comm"] = res["comm"] or float(kp.get("comm") or 0)
            res["name"] = res["name"] or kp.get("name", "")
        res["match_status"] = "exact_sku"
        if res["image"]:
            return res

    # Step 2: Fast local name matching (channel catalog + local pool) before touching remote API
    best_local: tuple[float, dict[str, Any]] | None = None
    if raw_name:
        for cand in list(cat_map.values()) + _POOL_CACHE:
            sim = name_similarity(raw_name, cand.get("name", ""))
            if sim >= 0.62 and (best_local is None or sim > best_local[0]):
                best_local = (sim, cand)
        if best_local:
            sim, cand = best_local
            cand_sku = extract_clean_sku(cand.get("sku"))
            if not raw_sku and sim >= 0.65 and cand_sku:
                res["sku"] = cand_sku
                res["link"] = f"https://shop.tiktok.com/view/product/{cand_sku}?region=TH&locale=th-TH"
                res["match_status"] = "auto_matched"
            res["image"] = res["image"] or cand.get("image", "")
            res["cat"] = res["cat"] or cand.get("cat", "")
            res["price"] = res["price"] or float(cand.get("price") or 0)
            res["comm"] = res["comm"] or float(cand.get("comm") or 0)
            if res["sku"] and res["image"]:
                return res

    # Step 3: Live lookup against 2.98M Kalodata DB by cleaned name
    if use_remote_api and raw_name:
        scored_candidates: list[tuple[float, dict[str, Any]]] = []
        seen_c_skus: set[str] = set()
        queries = progressive_search_queries(raw_name)
        # If item already has valid SKU and only needs a thumbnail, 1 query is enough
        if raw_sku and len(queries) > 1:
            queries = queries[:1]
        for q in queries:
            hits = remote_search_kalodata(q, limit=8)
            for raw_h in hits:
                h = _normalize_remote_product(raw_h)
                if not h["sku"] or h["sku"] in seen_c_skus:
                    continue
                seen_c_skus.add(h["sku"])
                sim = 1.0 if (raw_sku and h["sku"] == raw_sku) else name_similarity(raw_name, h["name"])
                scored_candidates.append((sim, h))
            if hits or any(sc >= 0.72 for sc, _ in scored_candidates):
                break
        scored_candidates.sort(key=lambda x: x[0], reverse=True)
        res["candidates"] = [
            {
                "sku": h["sku"],
                "name": h["name"],
                "image": h["image"],
                "cat": h["cat"],
                "price": h["price"],
                "comm": h["comm"],
                "score": round(sc, 2),
            }
            for sc, h in scored_candidates[:4]
        ]
        if scored_candidates:
            best_sc, best_h = scored_candidates[0]
            if raw_sku:
                # Preserve the exact SKU from Excel; only borrow image/cat when confidently similar or exact
                if best_h["sku"] == raw_sku or best_sc >= 0.55:
                    res["image"] = res["image"] or best_h["image"]
                    res["cat"] = res["cat"] or best_h["cat"]
                    res["price"] = res["price"] or best_h["price"]
                    res["comm"] = res["comm"] or best_h["comm"]
            else:
                if best_sc >= 0.48:
                    res["sku"] = best_h["sku"]
                    res["link"] = best_h["link"]
                    res["image"] = res["image"] or best_h["image"]
                    res["cat"] = res["cat"] or best_h["cat"]
                    res["price"] = res["price"] or best_h["price"]
                    res["comm"] = res["comm"] or best_h["comm"]
                    res["match_status"] = "auto_matched"

    if not res.get("sku"):
        res["match_status"] = "missing_sku"
    elif not res.get("match_status"):
        res["match_status"] = "exact_sku"
    return res


def parse_number(val: Any) -> float:
    """Parse Thai/English currency and count strings ('฿1,250.5', '1.2K', '-') into float."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return 0.0 if math.isnan(val) else float(val)
    s = str(val).strip()
    if not s or s in {"-", "--", "N/A", "nan", "None"}:
        return 0.0
    mult = 1.0
    if re.search(r"[kK]\s*$", s):
        mult = 1000.0
        s = re.sub(r"[kK]\s*$", "", s)
    elif re.search(r"[mM]\s*$", s):
        mult = 1_000_000.0
        s = re.sub(r"[mM]\s*$", "", s)
    cleaned = re.sub(r"[^0-9.\-]+", "", s.replace(",", ""))
    if not cleaned or cleaned in {"-", "."}:
        return 0.0
    try:
        return float(cleaned) * mult
    except ValueError:
        return 0.0


def _cell_to_str(val: Any) -> str:
    """Convert Excel cell value without turning 19-digit IDs into scientific floats if intact."""
    if val is None:
        return ""
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if math.isnan(val):
            return ""
        if val.is_integer() and val >= 1e14:
            return f"{int(val)}"
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val).strip()


def _detect_column_map(headers: list[str]) -> tuple[dict[str, int], int]:
    """Identify header columns for sku, link, name, gmv, orders, refunds, status, comm, price, shop."""
    col_map: dict[str, int] = {}
    score = 0
    norm_headers = [re.sub(r"\s+", " ", str(h or "").strip().lower()) for h in headers]

    rules: list[tuple[str, tuple[str, ...], tuple[str, ...]]] = [
        (
            "sku",
            ("รหัสสินค้า", "product id", "product_id", "ไอดีสินค้า", "id สินค้า", "item id", "item_id", "รหัสไอเทม", "sku"),
            ("id ของ sku", "sku id", "seller sku", "รหัสร้านค้า", "shop id", "รหัสเนื้อหา", "content id", "video", "วิดีโอ", "creator", "หมายเลขคำสั่งซื้อ", "order id"),
        ),
        ("link", ("product link", "product url", "ลิงก์สินค้า", "url สินค้า", "link", "url"), ("video",)),
        (
            "name",
            ("ชื่อสินค้า", "product name", "product_name", "product title", "item name", "ข้อมูลสินค้า", "ชื่อตะกร้า", "product", "สินค้า", "title"),
            ("id", "รหัส", "ลิงก์", "link", "ขายได้", "คืนเงิน", "ร้านค้า", "shop", "ประเภท", "สถานะ"),
        ),
        (
            "gmv",
            ("affiliate gmv", "video gmv", "gross merchandise", "gmv", "ยอดขายรวม", "มูลค่าสินค้า", "ยอดขายจากวิดีโอ", "ยอดขาย", "ฐานค่าคอมมิชชั่นตามจริง", "ฐานค่าคอมมิชชั่นโดยประมาณ", "revenue", "sales amount"),
            ("อัตรา", "rate", "%", "เปรียบเทียบ"),
        ),
        (
            "orders",
            ("สินค้าที่ขายได้", "จำนวนสินค้าที่ขายได้", "จำนวนที่ขายได้", "items sold", "units sold", "affiliate orders", "ยอดสั่งซื้อ", "ออเดอร์", "จำนวนชิ้น", "orders", "quantity", "sold"),
            ("หมายเลข", "ประเภท", "สถานะ", "วันที่", "คืนเงิน", "refund", "rate", "อัตรา", "%", "id", "รหัส"),
        ),
        (
            "refunds",
            ("สินค้าที่มีการคืนเงิน", "จำนวนคืนเงิน", "refunded items", "refunded units", "refunds"),
            ("rate", "อัตรา", "%"),
        ),
        (
            "status",
            ("สถานะการชำระคำสั่งซื้อ", "สถานะคำสั่งซื้อ", "order status", "settlement status"),
            (),
        ),
        (
            "comm",
            ("ยอดรายได้รวมสุดท้าย", "ค่าคอมมิชชั่นโดยประมาณ", "estimated commission", "est. commission", "ค่าคอมมิชชั่นมาตรฐาน", "commission", "ค่าคอมมิชชั่น", "คอมมิชชั่น", "ค่าคอม"),
            ("ประเภท", "ฐาน", "base", "rate", "อัตรา", "%", "วันที่", "โฆษณาร้านค้าโดยประมาณ", "มาตรฐานโดยประมาณ"),
        ),
        (
            "est_comm",
            ("ค่าคอมมิชชั่นมาตรฐานโดยประมาณ", "ค่าคอมมิชชั่นโฆษณาร้านค้าโดยประมาณ", "estimated standard commission"),
            ("อัตรา", "rate", "%"),
        ),
        ("price", ("avg price", "price", "ราคาเฉลี่ย", "ราคา"), ("gmv", "รวม", "total")),
        ("shop", ("ชื่อร้านค้า", "ร้านค้า", "shop name", "seller name"), ("รหัส", "id", "ประเภท", "โฆษณา", "ค่าคอม")),
        ("content_id", ("รหัสเนื้อหา", "content id", "content_id", "video id", "รหัสวิดีโอ"), ()),
        ("order_id", ("หมายเลขคำสั่งซื้อ", "order id", "order_id", "รหัสคำสั่งซื้อ"), ()),
        ("order_date", ("วันที่สั่งซื้อ", "เวลาที่สั่งซื้อ", "created time", "order date", "order time"), ("ชำระเงิน", "paid", "settlement")),
    ]

    used_indices: set[int] = set()
    for field, keywords, excludes in rules:
        best_idx = -1
        best_rank = -1
        for idx, h in enumerate(norm_headers):
            if not h or idx in used_indices:
                continue
            if any(ex in h for ex in excludes):
                continue
            for rank_i, kw in enumerate(keywords):
                if h == kw:
                    cand_rank = 100 - rank_i
                    if cand_rank > best_rank:
                        best_rank = cand_rank
                        best_idx = idx
                elif kw in h and (len(kw) >= 3 or h == kw):
                    cand_rank = 50 - rank_i
                    if cand_rank > best_rank:
                        best_rank = cand_rank
                        best_idx = idx
        if best_idx >= 0:
            col_map[field] = best_idx
            used_indices.add(best_idx)
            score += 2 if field in {"name", "sku", "gmv", "orders"} else 1

    return col_map, score


def _read_rows_from_file(filename: str, raw_bytes: bytes) -> list[list[str]]:
    """Read raw rows as strings from .xlsx, .xlsm, .csv, or .tsv without float SKU truncation or A1 dimension bugs."""
    ext = Path(filename).suffix.lower()
    if ext in {".xlsx", ".xlsm"}:
        # TikTok Affiliate exports often set <dimension ref="A1"/>, causing openpyxl read_only=True
        # to stop at cell A1! pandas.read_excel(..., dtype=str) reads all rows fast and preserves 19-digit IDs.
        try:
            sheets = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None, header=None, dtype=str, keep_default_na=False)
            best_rows: list[list[str]] = []
            for df in sheets.values():
                rows = [[str(c).strip() for c in row] for row in df.itertuples(index=False, name=None)]
                non_empty = [r for r in rows if any(c for c in r)]
                if len(non_empty) > len(best_rows):
                    best_rows = non_empty
            if best_rows:
                return best_rows
        except Exception:
            pass

        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=False, data_only=True)
        best_rows = []
        for ws in wb.worksheets:
            rows = [[_cell_to_str(c) for c in row] for row in ws.iter_rows(values_only=True)]
            non_empty = [r for r in rows if any(c.strip() for c in r)]
            if len(non_empty) > len(best_rows):
                best_rows = non_empty
        wb.close()
        return best_rows

    if ext in {".csv", ".tsv", ".txt"}:
        decoded = ""
        for enc in ("utf-8-sig", "utf-8", "utf-16", "cp874", "tis-620", "latin1"):
            try:
                decoded = raw_bytes.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if not decoded:
            decoded = raw_bytes.decode("utf-8", errors="ignore")
        delim = "\t" if ext == ".tsv" or decoded.count("\t") > decoded.count(",") else ","
        reader = csv.reader(io.StringIO(decoded), delimiter=delim)
        return [[str(c).strip() for c in r] for r in reader if any(str(c).strip() for c in r)]

    if ext == ".xls":
        raise ValueError(f"ไฟล์ {filename} เป็นรูปแบบ .xls รุ่นเก่า กรุณาเปิดใน Excel แล้ว Save As เป็น .xlsx หรือ .csv ครับ")
    raise ValueError(f"ยังไม่รองรับนามสกุลไฟล์ {ext} (รองรับ .xlsx และ .csv ครับ)")


def _parse_date_iso(raw_val: Any) -> str:
    """Parse DD/MM/YYYY or YYYY-MM-DD (including Thai Buddhist year) into ISO YYYY-MM-DD."""
    s = str(raw_val or "").strip()
    if not s:
        return ""
    m_ymd = re.search(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m_ymd:
        y, mo, d = int(m_ymd.group(1)), int(m_ymd.group(2)), int(m_ymd.group(3))
        if y > 2400:
            y -= 543
        try:
            return dt.date(y, mo, d).isoformat()
        except ValueError:
            return ""
    m_dmy = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m_dmy:
        d, mo, y = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
        if y > 2400:
            y -= 543
        try:
            return dt.date(y, mo, d).isoformat()
        except ValueError:
            return ""
    return ""


def _decode_tiktok_content_pub_date(content_id: Any) -> str:
    """Decode publish date (YYYY-MM-DD, local time) from 64-bit TikTok Snowflake content_id (int >> 32)."""
    ts = _decode_content_ts(content_id)
    return dt.datetime.fromtimestamp(ts).strftime("%Y-%m-%d") if ts else ""


def _decode_content_ts(content_id: Any) -> int:
    """Decode the Unix timestamp (seconds) embedded in a TikTok Snowflake content_id; 0 when invalid."""
    s = re.sub(r"\D+", "", str(content_id or "").strip())
    if not (17 <= len(s) <= 20):
        return 0
    try:
        ts = int(s) >> 32
    except ValueError:
        return 0
    return ts if 1_500_000_000 <= ts <= 2_100_000_000 else 0


def parse_sales_files(files_payload: list[dict[str, str]]) -> dict[str, Any]:
    """Parse uploaded TikTok monthly Excel/CSV files and aggregate unique products cleanly."""
    aggregated: dict[str, dict[str, Any]] = {}
    warnings: list[str] = []
    parsed_files: list[str] = []
    min_order_date = ""
    max_order_date = ""

    for fobj in files_payload:
        fname = str(fobj.get("filename") or "sales.xlsx")
        b64 = str(fobj.get("content_base64") or "")
        if "," in b64 and b64.startswith("data:"):
            b64 = b64.split(",", 1)[1]
        raw_bytes = base64.b64decode(b64)
        rows = _read_rows_from_file(fname, raw_bytes)
        if len(rows) < 2:
            warnings.append(f"ไฟล์ {fname} ไม่มีข้อมูลสินค้า")
            continue

        best_header_idx = 0
        best_map: dict[str, int] = {}
        best_score = -1
        for r_idx in range(min(15, len(rows))):
            cmap, sc = _detect_column_map(rows[r_idx])
            if sc > best_score:
                best_score = sc
                best_map = cmap
                best_header_idx = r_idx

        if "name" not in best_map and "sku" not in best_map:
            warnings.append(f"ไม่พบคอลัมน์ชื่อสินค้าหรือรหัสสินค้าในไฟล์ {fname}")
            continue

        parsed_files.append(fname)
        corrupted_count = 0

        for row in rows[best_header_idx + 1 :]:
            def get_col(key: str) -> str:
                i = best_map.get(key, -1)
                return row[i].strip() if 0 <= i < len(row) else ""

            status_val = get_col("status").lower()
            if status_val and any(bad in status_val for bad in ("ยกเลิก", "ไม่มีสิทธิ์", "cancelled", "canceled", "ineligible")):
                continue

            o_date = _parse_date_iso(get_col("order_date"))
            if o_date:
                if not min_order_date or o_date < min_order_date:
                    min_order_date = o_date
                if not max_order_date or o_date > max_order_date:
                    max_order_date = o_date

            raw_sku_cell = get_col("sku")
            raw_link_cell = get_col("link")
            name = normalize_thai_text(get_col("name"))
            sku = extract_clean_sku(raw_sku_cell) or extract_clean_sku(raw_link_cell) or extract_clean_sku(name)

            if is_likely_float_corrupted_sku(raw_sku_cell, sku):
                corrupted_count += 1
                sku = ""

            if not name and not sku:
                continue

            gmv = parse_number(get_col("gmv"))
            sold = int(round(parse_number(get_col("orders")))) if "orders" in best_map else (1 if gmv > 0 else 0)
            refunded = int(round(parse_number(get_col("refunds")))) if "refunds" in best_map else 0
            net_orders = max(0, sold - refunded)
            fully_refunded = sold > 0 and net_orders == 0
            gmv_eff = 0.0 if fully_refunded else gmv
            comm_raw = parse_number(get_col("comm")) or parse_number(get_col("est_comm"))
            comm_eff = 0.0 if fully_refunded else comm_raw
            price = parse_number(get_col("price"))
            shop_name = normalize_thai_text(get_col("shop"))
            content_id = get_col("content_id")
            order_id = get_col("order_id")
            if price <= 0 and gmv_eff > 0 and net_orders > 0:
                price = round(gmv_eff / net_orders, 2)

            agg_key = sku if sku else norm_token_key(name)
            if not agg_key:
                continue

            if agg_key not in aggregated:
                cid_stats: dict[str, dict[str, Any]] = {}
                if content_id and (net_orders > 0 or gmv_eff > 0):
                    cid_stats[content_id] = {
                        "pub_date": _decode_tiktok_content_pub_date(content_id),
                        "comm": comm_eff,
                        "gmv": gmv_eff,
                        "orders": net_orders,
                    }
                aggregated[agg_key] = {
                    "sku": sku,
                    "name": name or f"สินค้า TikTok ({sku})",
                    "shop": shop_name,
                    "link": raw_link_cell if raw_link_cell.startswith("http") else (
                        f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH" if sku else ""
                    ),
                    "gmv": round(gmv_eff, 2),
                    "orders": net_orders,
                    "refunded": refunded,
                    "comm": round(comm_eff, 2),
                    "price": round(price, 2),
                    "_cid_stats": cid_stats,
                    "_order_ids": {order_id} if order_id else set(),
                }
            else:
                cur = aggregated[agg_key]
                cur["gmv"] = round(cur["gmv"] + gmv_eff, 2)
                cur["orders"] += net_orders
                cur["refunded"] = cur.get("refunded", 0) + refunded
                cur["comm"] = round(cur["comm"] + comm_eff, 2)
                if content_id and (net_orders > 0 or gmv_eff > 0):
                    cid_map = cur["_cid_stats"]
                    if content_id not in cid_map:
                        cid_map[content_id] = {
                            "pub_date": _decode_tiktok_content_pub_date(content_id),
                            "comm": comm_eff,
                            "gmv": gmv_eff,
                            "orders": net_orders,
                        }
                    else:
                        st = cid_map[content_id]
                        st["comm"] += comm_eff
                        st["gmv"] += gmv_eff
                        st["orders"] += net_orders
                if order_id:
                    cur["_order_ids"].add(order_id)
                if len(name) > len(cur["name"]):
                    cur["name"] = name
                if shop_name and not cur.get("shop"):
                    cur["shop"] = shop_name
                if price > 0 and cur["price"] <= 0:
                    cur["price"] = round(price, 2)

        if corrupted_count > 0:
            warnings.append(
                f"ไฟล์ {fname} มีรหัสสินค้า {corrupted_count} รายการที่ถูก Excel ปัดเศษ (เช่น E+18) — ระบบกำลังค้นรหัส 19 หลักจริงจากชื่อสินค้าให้แทนครับ"
            )

    products_out: list[dict[str, Any]] = []
    for prod in aggregated.values():
        cid_stats = prod.pop("_cid_stats", {})
        oids = prod.pop("_order_ids", set())
        orders = int(prod.get("orders") or 0)
        gmv = float(prod.get("gmv") or 0)
        comm = float(prod.get("comm") or 0)
        videos_count = len(cid_stats) if cid_stats else (1 if (orders > 0 or gmv > 0) else 0)
        distinct_orders = len(oids) if oids else orders
        if min_order_date:
            evergreen_cids = [c for c, st in cid_stats.items() if st["pub_date"] and st["pub_date"] < min_order_date]
            new_cids = [c for c, st in cid_stats.items() if not st["pub_date"] or st["pub_date"] >= min_order_date]
        else:
            evergreen_cids = []
            new_cids = list(cid_stats.keys())
        evergreen_comm = round(sum(cid_stats[c]["comm"] for c in evergreen_cids), 2)
        prod["videos_count"] = videos_count
        prod["new_videos_count"] = len(new_cids) if cid_stats else videos_count
        prod["evergreen_videos_count"] = len(evergreen_cids)
        prod["evergreen_comm"] = evergreen_comm
        prod["new_comm"] = round(max(0.0, comm - evergreen_comm), 2)
        prod["converting_new_clips"] = 0
        prod["distinct_orders"] = distinct_orders
        prod["comm_pct_eff"] = round((comm / gmv) * 100.0, 1) if gmv > 0 else 0.0
        prod["comm_per_video"] = round(comm / max(1, videos_count), 2) if (orders > 0 or gmv > 0) else 0.0
        prod["comm_per_order"] = round(comm / max(1, distinct_orders), 2) if (orders > 0 or gmv > 0) else 0.0
        prod["posted_clips"] = 0
        prod["bot_logged_clips"] = 0
        prod["hit_rate_pct"] = 0.0
        prod["real_comm_per_posted"] = 0.0
        prod["_cid_pub_comm"] = {
            c: {
                "pub_date": st["pub_date"],
                "comm": round(st["comm"], 2),
                "orders": int(st.get("orders") or 0),
                "pub_ts": _decode_content_ts(c),
            }
            for c, st in cid_stats.items()
        }
        products_out.append(prod)

    return {
        "products": products_out,
        "parsed_files": parsed_files,
        "warnings": warnings,
        "min_order_date": min_order_date,
        "max_order_date": max_order_date,
    }


def enrich_with_bot_posted_stats(
    items: list[dict[str, Any]],
    min_date: str = "",
    max_date: str = "",
    hist_path: Path | None = None,
    log_path: Path | None = None,
) -> list[dict[str, Any]]:
    """Cross-reference Desktop bot posted history strictly against new-in-window clips (excluding old evergreen clips)."""
    if not min_date and not max_date:
        for it in items:
            it.pop("_cid_pub_comm", None)
        return items
    hist_file = hist_path or (Path.home() / "Desktop" / "posted_clip_history.json")
    log_file = log_path or (Path.home() / "Desktop" / "ประวัติการทำงานออโต้_TikTok.log")

    matured_cutoff = max_date
    if max_date:
        try:
            matured_cutoff = (dt.date.fromisoformat(max_date) - dt.timedelta(days=2)).isoformat()
        except ValueError:
            matured_cutoff = max_date

    sku_post_counts: dict[str, int] = {}
    sku_matured_counts: dict[str, int] = {}
    sku_meta: dict[str, dict[str, str]] = {}
    name_entries: list[str] = []
    matured_name_entries: list[str] = []
    all_log_dates: list[str] = []

    if hist_file.is_file():
        try:
            data = json.loads(hist_file.read_text(encoding="utf-8"))
            if isinstance(data, list):
                for row in data:
                    d_str = str(row.get("date") or "")[:10]
                    if d_str:
                        all_log_dates.append(d_str)
                    if min_date and d_str and d_str < min_date:
                        continue
                    if max_date and d_str and d_str > max_date:
                        continue
                    is_matured = (not matured_cutoff) or (not d_str) or (d_str <= matured_cutoff)
                    sku = extract_clean_sku(row.get("sku") or row.get("filename"))
                    nm = normalize_thai_text(row.get("name") or "")
                    if sku:
                        sku_post_counts[sku] = sku_post_counts.get(sku, 0) + 1
                        if is_matured:
                            sku_matured_counts[sku] = sku_matured_counts.get(sku, 0) + 1
                        if nm and sku not in sku_meta:
                            sku_meta[sku] = {"name": nm}
                    elif nm:
                        name_entries.append(nm)
                        if is_matured:
                            matured_name_entries.append(nm)
        except Exception:
            pass

    if log_file.is_file():
        try:
            for line in log_file.read_text(encoding="utf-8", errors="ignore").splitlines():
                m = re.match(r"^\[(\d{4}-\d{2}-\d{2})\s+[^\]]+\]\s+(?:โพสต์|ตั้งเวลา[^:]*|บันทึกร่าง)[^:]*สำเร็จ[^:]*:\s*(.+)$", line)
                if not m:
                    continue
                d_str, body = m.group(1), m.group(2).split("|", 1)[0].strip()
                all_log_dates.append(d_str)
                if min_date and d_str < min_date:
                    continue
                if max_date and d_str > max_date:
                    continue
                is_matured = (not matured_cutoff) or (d_str <= matured_cutoff)
                sku = extract_clean_sku(body)
                if sku:
                    if not hist_file.is_file():
                        sku_post_counts[sku] = sku_post_counts.get(sku, 0) + 1
                        if is_matured:
                            sku_matured_counts[sku] = sku_matured_counts.get(sku, 0) + 1
                else:
                    clean_b = re.sub(r"\.mp4$", "", body, flags=re.I)
                    clean_b = re.sub(r"^ตอบ(?:คอมเมนต์|comment)\s*", "", clean_b, flags=re.I).strip()
                    if clean_b:
                        name_entries.append(clean_b)
                        if is_matured:
                            matured_name_entries.append(clean_b)
        except Exception:
            pass

    first_log_date = min(all_log_dates) if all_log_dates else ""
    if min_date and first_log_date:
        match_start = max(min_date, first_log_date)
    else:
        match_start = min_date or first_log_date
    item_skus_present: set[str] = set()

    for it in items:
        sku = str(it.get("sku") or "")
        if sku:
            item_skus_present.add(sku)
        cid_pub_comm = it.pop("_cid_pub_comm", {})

        if cid_pub_comm:
            win_cids = [
                c for c, st in cid_pub_comm.items()
                if (not st["pub_date"]) or (
                    (not match_start or st["pub_date"] >= match_start)
                    and (not max_date or st["pub_date"] <= max_date)
                )
            ]
            conv_new = len(win_cids)
            win_new_comm = sum(float(cid_pub_comm[c]["comm"]) for c in win_cids)
        else:
            conv_new = int(it.get("new_videos_count") if it.get("new_videos_count") is not None else (it.get("videos_count") or 0))
            win_new_comm = float(it.get("new_comm") if it.get("new_comm") is not None else (it.get("comm") or 0))

        logged = sku_post_counts.get(sku, 0) if sku else 0
        matured_logged = sku_matured_counts.get(sku, 0) if sku else 0
        if logged == 0 and name_entries and it.get("name"):
            it_name = str(it["name"])
            logged = sum(1 for ne in name_entries if name_similarity(it_name, ne) >= 0.72)
            matured_logged = sum(1 for ne in matured_name_entries if name_similarity(it_name, ne) >= 0.72)

        it["bot_logged_clips"] = logged
        it["converting_new_clips"] = conv_new
        window_aligned = (
            bool(cid_pub_comm)
            or bool(first_log_date and min_date and min_date >= first_log_date)
            or (logged > conv_new)
        )
        if window_aligned and (matured_logged > 0 or (logged > 0 and conv_new > 0)):
            base_posted = matured_logged if matured_logged >= conv_new else logged
            eff_posted = max(base_posted, conv_new, 1)
            it["posted_clips"] = eff_posted
            it["hit_rate_pct"] = round(min(100.0, (conv_new / eff_posted) * 100.0), 0)
            it["real_comm_per_posted"] = round(win_new_comm / eff_posted, 2)
        else:
            it["posted_clips"] = 0
            it["hit_rate_pct"] = 0.0
            it["real_comm_per_posted"] = 0.0

    # Inject SKUs with >= 3 matured bot posts (>48h before max_date) in window and 0 orders in Excel
    for sku, matured_cnt in sku_matured_counts.items():
        if matured_cnt >= 3 and sku not in item_skus_present:
            meta = sku_meta.get(sku, {})
            items.append({
                "sku": sku,
                "name": meta.get("name") or f"สินค้าที่บอทโพสต์แล้วขายไม่ได้ ({sku})",
                "shop": "",
                "link": f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH",
                "gmv": 0.0,
                "orders": 0,
                "refunded": 0,
                "comm": 0.0,
                "price": 0.0,
                "videos_count": 0,
                "new_videos_count": 0,
                "evergreen_videos_count": 0,
                "converting_new_clips": 0,
                "distinct_orders": 0,
                "comm_pct_eff": 0.0,
                "comm_per_video": 0.0,
                "comm_per_order": 0.0,
                "posted_clips": matured_cnt,
                "bot_logged_clips": sku_post_counts.get(sku, matured_cnt),
                "hit_rate_pct": 0.0,
                "real_comm_per_posted": 0.0,
            })

    return items


HOME_CORE_CATS = {
    "อุปกรณ์เครื่องใช้ภายในบ้าน",
    "สิ่งทอและเครื่องตกแต่งที่ใช้วัสดุผิวนุ่ม",
    "เครื่องครัว",
    "เครื่องใช้ในบ้าน",
    "เฟอร์นิเจอร์",
    "การปรับปรุงบ้าน",
    "อุปกรณ์สัตว์เลี้ยง",
}
CROSS_IMPULSE_CATS = {
    "ความงามและของใช้ส่วนตัว",
    "อาหารและเครื่องดื่ม",
    "สุขภาพ",
    "เครื่องมือและเครื่องโลหะ",
    "ยานยนต์และจักรยานยนต์",
}


def _discover_gemini_api_key(explicit_key: str = "") -> str:
    """Resolve Gemini API key dynamically from request, environment, or local desktop pipeline config."""
    k = str(explicit_key or "").strip()
    if k:
        return k
    env_k = os.environ.get("GEMINI_API_KEY", "").strip()
    if env_k:
        return env_k
    desktop_bot = Path.home() / "Desktop" / "tiktok_auto_pipeline.py"
    if desktop_bot.is_file():
        try:
            txt = desktop_bot.read_text(encoding="utf-8", errors="ignore")
            m = re.search(r"GEMINI_API_KEY\s*=\s*['\"]([^'\"]+)['\"]", txt)
            if m:
                return m.group(1).strip()
        except Exception:
            pass
    return ""


def _build_new_star_candidate_pool(
    winners: list[dict[str, Any]],
    dropped: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Filter Kalodata pool into high-commission impulse video candidates (core vs cross-category)."""
    pool = ensure_pool_loaded()
    if not pool:
        return [], []

    used_skus = {p.get("sku") for p in winners + dropped if p.get("sku")}
    used_names = [norm_token_key(p.get("name", "")) for p in winners + dropped if p.get("name")]

    winner_cats = {str(w.get("cat") or "").strip() for w in winners if w.get("cat")}
    core_cats = HOME_CORE_CATS | winner_cats

    core_scored: list[tuple[float, dict[str, Any]]] = []
    cross_scored: list[tuple[float, dict[str, Any]]] = []

    for p in pool:
        sku = str(p.get("sku") or "")
        if not sku or sku in used_skus:
            continue
        comm = float(p.get("comm") or 0)
        price = float(p.get("price") or p.get("avg_price") or 0)
        v_rev = max(0.0, float(p.get("v_rev") or 0))
        rev = max(1.0, float(p.get("rev") or v_rev or 1))
        v_share = v_rev / rev

        # Strict quality gate: eliminate 0%/tiny-comm traps & unsuited big-ticket electronics
        if comm < 5.0 or price < 25.0 or price > 1290.0 or v_rev < 20000.0 or v_share < 0.40:
            continue

        nk = norm_token_key(p.get("name", ""))
        if any(nk == un or (len(nk) >= 12 and (nk in un or un in nk)) for un in used_names):
            continue

        cat = str(p.get("cat") or "").strip()
        if cat not in core_cats and cat not in CROSS_IMPULSE_CATS:
            continue

        creators = max(1.0, float(p.get("creators") or 1))
        conv = max(1.0, float(p.get("conv") or 5))
        growth = float(p.get("growth") or 0)
        ai_score = max(40.0, float(p.get("ai") or 65))

        comm_baht = price * (comm / 100.0)
        blue_ocean = (v_rev / math.pow(creators + 6.0, 1.15)) * (0.65 + conv / 45.0)
        growth_boost = 1.0 + min(max(growth, 0.0), 400.0) / 120.0
        comm_boost = 0.8 + min(comm_baht, 120.0) / 50.0
        score = blue_ocean * growth_boost * comm_boost * (0.5 + 0.5 * min(1.0, v_share)) * (ai_score / 100.0)

        entry = {
            "sku": sku,
            "name": p["name"],
            "image": p.get("image", ""),
            "cat": cat,
            "price": price,
            "comm": comm,
            "comm_baht": round(comm_baht, 1),
            "v_rev": v_rev,
            "creators": int(p.get("creators") or 0),
            "conv": conv,
            "growth": growth,
            "link": p.get("link") or f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH",
            "tier": "new_star",
            "starred": True,
            "segment": "core" if cat in core_cats else "cross",
            "reason": (
                f"🌟 ตรงหมวดบ้าน/ของใช้ (คอม {comm:.0f}% ~฿{comm_baht:,.0f}/ชิ้น)"
                if cat in core_cats
                else f"🚀 ข้ามหมวดคนซื้อกลุ่มเดียวกัน (คอม {comm:.0f}% ~฿{comm_baht:,.0f}/ชิ้น)"
            ),
        }
        if cat in core_cats:
            core_scored.append((score, entry))
        else:
            cross_scored.append((score, entry))

    core_scored.sort(key=lambda x: x[0], reverse=True)
    cross_scored.sort(key=lambda x: x[0], reverse=True)
    return [x[1] for x in core_scored], [x[1] for x in cross_scored]


def select_new_stars_from_pool(
    winners: list[dict[str, Any]],
    dropped: list[dict[str, Any]],
    desired_count: int = 6,
    strategy: str = "balanced",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Pick diversified, high-commission Blue Ocean products matching core_only / balanced / expand strategy."""
    if desired_count <= 0:
        return [], []
    core_list, cross_list = _build_new_star_candidate_pool(winners, dropped)
    if not core_list and not cross_list:
        return [], []

    def pick_diverse(source_list: list[dict[str, Any]], count: int, used_prefixes: dict[str, int], cat_counts: dict[str, int], picked_skus: set[str]) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        for item in source_list:
            if len(out) >= count:
                break
            if item["sku"] in picked_skus:
                continue
            prefix = " ".join(clean_product_query(item["name"]).lower().split()[:2])
            if used_prefixes.get(prefix, 0) >= 1:
                continue
            if cat_counts.get(item["cat"], 0) >= 2:
                continue
            used_prefixes[prefix] = used_prefixes.get(prefix, 0) + 1
            cat_counts[item["cat"]] = cat_counts.get(item["cat"], 0) + 1
            picked_skus.add(item["sku"])
            out.append(item)
        return out

    used_prefixes: dict[str, int] = {}
    cat_counts: dict[str, int] = {}
    picked_skus: set[str] = set()

    if strategy == "core_only":
        n_core, n_cross = desired_count, 0
    elif strategy == "expand":
        n_cross = max(1, round(desired_count * 0.65))
        n_core = max(0, desired_count - n_cross)
    else:  # balanced: ~65% core home-living + ~35% cross-category same-buyer impulse
        n_cross = max(1, desired_count // 3) if desired_count >= 3 else 0
        n_core = max(0, desired_count - n_cross)

    picked = pick_diverse(core_list, n_core, used_prefixes, cat_counts, picked_skus)
    picked.extend(pick_diverse(cross_list, n_cross, used_prefixes, cat_counts, picked_skus))
    if len(picked) < desired_count:
        picked.extend(pick_diverse(core_list + cross_list, desired_count - len(picked), used_prefixes, cat_counts, picked_skus))

    # Build backup list for swaps and AI selection (keep diversity across core and cross)
    backups: list[dict[str, Any]] = []
    max_backup = 40
    i_c, i_x = 0, 0
    while len(backups) < max_backup and (i_c < len(core_list) or i_x < len(cross_list)):
        if i_c < len(core_list):
            if core_list[i_c]["sku"] not in picked_skus:
                backups.append(core_list[i_c])
            i_c += 1
        if len(backups) < max_backup and i_c < len(core_list):
            if core_list[i_c]["sku"] not in picked_skus:
                backups.append(core_list[i_c])
            i_c += 1
        if len(backups) < max_backup and i_x < len(cross_list):
            if cross_list[i_x]["sku"] not in picked_skus:
                backups.append(cross_list[i_x])
            i_x += 1

    return picked, backups


def grade_channel_products_smart(items: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Classify channel products dynamically using real commission share (Pareto 85/93.5%), traps, and Hidden Gems."""
    has_any_sales = any((float(x.get("gmv") or 0) > 0 or int(x.get("orders") or 0) > 0) for x in items)
    if not has_any_sales:
        winners: list[dict[str, Any]] = []
        for it in items:
            it["tier"] = "winner"
            it["grade"] = "A"
            it["starred"] = True
            it["reason"] = "รายการสินค้าจากช่อง (ไม่มีคอลัมน์ยอดขาย — จัดเข้ากลุ่มทำคลิป)"
            winners.append(it)
        return winners, []

    total_comm = sum(float(x.get("comm") or 0) for x in items)
    use_comm = total_comm > 0.0

    def perf_val(x: dict[str, Any]) -> float:
        return float(x.get("comm") or 0) if use_comm else float(x.get("gmv") or 0)

    positive = [x for x in items if (int(x.get("orders") or 0) > 0 or float(x.get("gmv") or 0) > 0)]
    zero_items = [x for x in items if x not in positive]

    positive.sort(key=lambda x: (perf_val(x), float(x.get("gmv") or 0), int(x.get("orders") or 0)), reverse=True)
    total_metric = max(1.0, sum(perf_val(x) for x in positive))

    winners = []
    dropped = []

    for it in zero_items:
        p_clips = int(it.get("posted_clips") or 0)
        it["tier"] = "dropped"
        it["grade"] = "ZERO"
        it["starred"] = False
        it["reason"] = (
            f"❌ โพสต์จริง {p_clips} คลิป แต่ขายได้ 0 ชิ้น — ปัดตกอัตโนมัติ"
            if p_clips > 0
            else "❌ ยอดขาย ฿0 / 0 ชิ้น — ปัดตกอัตโนมัติ"
        )
        dropped.append(it)

    cum_metric = 0.0
    for rank_idx, it in enumerate(positive):
        gmv = float(it.get("gmv") or 0)
        comm = float(it.get("comm") or 0)
        orders = int(it.get("orders") or 0)
        d_orders = int(it.get("distinct_orders") or orders)
        vids = int(it.get("videos_count") or 1)
        new_vids = int(it.get("new_videos_count") if it.get("new_videos_count") is not None else vids)
        old_vids = int(it.get("evergreen_videos_count") or 0)
        conv_new = int(it.get("converting_new_clips") or 0)
        posted_c = int(it.get("posted_clips") or 0)
        hit_pct = float(it.get("hit_rate_pct") or 0)
        rcpp = float(it.get("real_comm_per_posted") or 0)
        cpv = float(it.get("comm_per_video") or 0)
        cpo = float(it.get("comm_per_order") or 0)
        cpct = float(it.get("comm_pct_eff") or 0)

        m_val = perf_val(it)
        cum_before = cum_metric / total_metric
        cum_metric += m_val

        # Volume trap or Low Hit-Rate waste check
        is_volume_trap = use_comm and orders >= 6 and cpo < 2.2 and (comm / total_metric) < 0.004
        is_low_hit_waste = posted_c >= 3 and hit_pct < 20.0 and rcpp < 8.0 and cum_before >= 0.85

        # Anti-Fluke Hidden Gem: MUST either prove >= 4 distinct orders (repeatable buyer demand)
        # OR have verified bot post count <= 3 with >= 50% hit rate! Never tag 1-3 order flukes as A_GEM.
        is_verified_gem = (
            (posted_c > 0 and posted_c <= 3 and hit_pct >= 50.0 and comm >= 60.0)
            or (
                d_orders >= 4
                and vids <= 4
                and (
                    (use_comm and (comm >= 110.0 or (comm >= 70.0 and cpct >= 8.0)))
                    or (not use_comm and gmv >= total_metric * 0.015)
                )
            )
        )

        if posted_c > 0:
            clip_summary = (
                f"โพสต์ใหม่ {posted_c} ➡️ ขายติด {conv_new} คลิป ({hit_pct:.0f}%) · เฉลี่ย ฿{rcpp:,.1f}/คลิปใหม่"
                + (f" + คลิปเก่ากินยาว {old_vids} คลิป" if old_vids > 0 else "")
            )
        else:
            if old_vids > 0 and new_vids > 0:
                vid_txt = f"ขายออก {vids} คลิป (ใหม่ {new_vids} · เก่ากินยาว {old_vids})"
            elif old_vids > 0:
                vid_txt = f"ขายออกจากคลิปเก่ากินยาว {old_vids} คลิป"
            else:
                vid_txt = f"ขายออก {vids} คลิป"
            clip_summary = f"{d_orders} ออเดอร์ · {vid_txt} (฿{cpv:,.0f}/คลิป)"

        if (cum_before < 0.85 or rank_idx < 3) and not (use_comm and comm <= 0 and gmv < 1000):
            it["tier"] = "winner"
            it["grade"] = "S"
            it["starred"] = True
            it["reason"] = (
                f"👑 เกรด S ตัวแบกช่อง (คอม ฿{comm:,.0f} · {cpct}% · {clip_summary})"
                if use_comm
                else f"👑 เกรด S ตัวแบกช่อง (GMV ฿{gmv:,.0f} · {orders} ชิ้น)"
            )
            winners.append(it)
        elif is_low_hit_waste:
            it["tier"] = "dropped"
            it["grade"] = "LOW_HIT"
            it["starred"] = False
            extra_old = f" (ส่วนคลิปเก่าค้างช่อง {old_vids} คลิปปล่อยกินยาวต่อได้ ไม่ต้องเปลืองคิวถ่ายใหม่)" if old_vids > 0 else ""
            it["reason"] = f"⚠️ ปัดตกคิวถ่ายใหม่: บอทโพสต์ใหม่ {posted_c} คลิป ขายติดแค่ {conv_new} คลิป (ติด {hit_pct:.0f}% · เฉลี่ย ฿{rcpp:.1f}/คลิปใหม่){extra_old}"
            dropped.append(it)
        elif is_volume_trap:
            it["tier"] = "dropped"
            it["grade"] = "TRAP"
            it["starred"] = False
            it["reason"] = f"⚠️ ปัดตก (เหนื่อยฟรีคอมเตี้ย): ขายตั้ง {orders} ชิ้น แต่ได้คอมแค่ ฿{comm:,.1f} (เฉลี่ย ฿{cpo:.1f}/ชิ้น)"
            dropped.append(it)
        elif is_verified_gem:
            it["tier"] = "winner"
            it["grade"] = "A_GEM"
            it["starred"] = True
            it["reason"] = (
                f"💎 ช้างเผือกมีหลักฐานซ้ำ: ทำยอดซ้ำได้ถึง {d_orders} ออเดอร์ จากขายออก {vids} คลิป (คอม ฿{comm:,.0f} · {cpct}%)"
                if use_comm
                else f"💎 ช้างเผือกน่าปั้น: {d_orders} ออเดอร์จากขายออก {vids} คลิป (฿{gmv:,.0f})"
            )
            winners.append(it)
        elif cum_before < 0.935 and (d_orders >= 3 or comm >= 40.0 or not use_comm):
            it["tier"] = "winner"
            it["grade"] = "A"
            it["starred"] = True
            it["reason"] = (
                f"⭐ เกรด A ทำเงิน ({clip_summary} · คอม ฿{comm:,.0f} · {cpct}%)"
                if use_comm
                else f"⭐ เกรด A ทำเงิน (GMV ฿{gmv:,.0f} · {orders} ชิ้น)"
            )
            winners.append(it)
        else:
            it["tier"] = "dropped"
            it["grade"] = "C"
            it["starred"] = False
            if d_orders <= 2:
                it["reason"] = f"🗑️ ปัดตก (กันยอดฟลุค): มีแค่ {d_orders} ออเดอร์ (ขายออก {vids} คลิป · คอม ฿{comm:,.1f}) เสี่ยงลงหลายคลิปแต่เพิ่งขายได้ 1–2 ชิ้น"
            else:
                it["reason"] = (
                    f"🗑️ หางแถว (~6% ท้ายตาราง): ทั้งเดือนได้คอม ฿{comm:,.1f} ({d_orders} ออเดอร์ · ขายออก {vids} คลิป)"
                    if use_comm
                    else f"🗑️ หางแถวท้ายตาราง: GMV ฿{gmv:,.0f} ({orders} ชิ้น)"
                )
            dropped.append(it)

    if not winners and dropped:
        promoted = dropped[: min(10, len(dropped))]
        dropped = dropped[len(promoted) :]
        for it in promoted:
            it["tier"] = "winner"
            it["grade"] = "A"
            it["starred"] = True
            winners.append(it)

    return winners, dropped


def run_ai_portfolio_advisor(
    winners: list[dict[str, Any]],
    dropped: list[dict[str, Any]],
    picked_stars: list[dict[str, Any]],
    backup_stars: list[dict[str, Any]],
    strategy: str = "balanced",
    desired_count: int = 6,
    explicit_key: str = "",
) -> dict[str, Any]:
    """Use Gemini AI (blunt, data-first) to review borderline items and curate best New Star products."""
    api_key = _discover_gemini_api_key(explicit_key)
    if not api_key:
        return {"ai_used": False, "ai_model": "", "ai_critique": [], "winners": winners, "dropped": dropped, "new_stars": picked_stars}

    try:
        from google import genai
        from google.genai import types
    except Exception:
        return {"ai_used": False, "ai_model": "", "ai_critique": [], "winners": winners, "dropped": dropped, "new_stars": picked_stars}

    all_channel = winners + dropped
    compact_channel = [
        {
            "sku": p["sku"],
            "tier_auto": p["tier"],
            "grade": p.get("grade", ""),
            "gmv": round(float(p.get("gmv") or 0), 0),
            "comm": round(float(p.get("comm") or 0), 1),
            "comm_pct": float(p.get("comm_pct_eff") or 0),
            "orders": int(p.get("orders") or 0),
            "distinct_orders": int(p.get("distinct_orders") or p.get("orders") or 0),
            "converting_videos": int(p.get("videos_count") or 0),
            "new_converting_videos": int(p.get("new_videos_count") or 0),
            "evergreen_old_videos": int(p.get("evergreen_videos_count") or 0),
            "posted_clips": int(p.get("posted_clips") or 0),
            "converting_new_from_posted": int(p.get("converting_new_clips") or 0),
            "comm_per_converting_vid": float(p.get("comm_per_video") or 0),
            "name": str(p.get("name") or "")[:55],
        }
        for p in all_channel[:45]
        if p.get("sku")
    ]

    candidate_pool = (picked_stars + backup_stars)[:28]
    compact_cands = [
        {
            "sku": c["sku"],
            "segment": c.get("segment", "core"),
            "cat": c.get("cat", ""),
            "price": round(float(c.get("price") or 0), 0),
            "comm_pct": float(c.get("comm") or 0),
            "comm_baht": float(c.get("comm_baht") or 0),
            "v_rev": round(float(c.get("v_rev") or 0), 0),
            "creators": int(c.get("creators") or 0),
            "conv": float(c.get("conv") or 0),
            "growth": float(c.get("growth") or 0),
            "name": str(c.get("name") or "")[:60],
        }
        for c in candidate_pool
    ]

    strat_desc = {
        "balanced": f"สมดุล: เลือกทั้งหมด {desired_count} ตัว โดยเน้นหมวดบ้าน/ของใช้ใกล้เคียงประมาณ {max(1, desired_count - max(1, desired_count // 3))} ตัว และอีก {max(1, desired_count // 3) if desired_count >= 3 else 0} ตัวเป็นสินค้าข้ามหมวดที่คนดูกลุ่มเดียวกันซื้อง่าย คอมสูง",
        "core_only": f"เน้นหมวดเดิม: เลือก {desired_count} ตัวจาก segment='core' เป็นหลัก",
        "expand": f"ขยายหมวด: เลือก {desired_count} ตัว โดยผสม segment='cross' ที่คอมคุ้มคู่แข่งน้อยเป็นหลัก",
    }.get(strategy, f"เลือก {desired_count} ตัวที่ดีที่สุด")

    prompt = (
        "คุณคือที่ปรึกษากลยุทธ์ TikTok Shop Affiliate สายข้อมูลจริง (ห้ามอวย ห้ามพูดเอาใจ ให้วิเคราะห์ตรงไปตรงมาตามตัวเลขเท่านั้น):\n"
        "⚠️ กฎตีความข้อมูลสำคัญมาก: `converting_videos` คือจำนวนคลิปที่ 'ขายออกแล้วอย่างน้อย 1 ชิ้น' เท่านั้น (ไม่รวมคลิปที่โพสต์แล้วขายไม่ได้ 0 บาท, ยกเว้นเมื่อ `posted_clips > 0` นั่นคือจำนวนที่บอทโพสต์จริง)\n"
        "- ระบบถอดวันโพสต์จริงจาก Snowflake ID แล้วแยก `new_converting_videos` (คลิปใหม่ที่โพสต์ในรอบเดือนนี้แล้วขายออก) ออกจาก `evergreen_old_videos` (คลิปเก่าจากเดือนก่อนๆ ที่ยังกินยาวขายออกอยู่) เพื่อไม่ให้คลิปเก่ามาหลอกสัดส่วนคลิปใหม่ (`converting_new_from_posted`)\n"
        "- ห้ามหลงอวยสินค้าที่ `distinct_orders` แค่ 1–3 ออเดอร์จาก `converting_videos = 1` ว่าเป็นช้างเผือกเด็ดขาด เพราะอาจลงไป 10 คลิปแล้วฟลุคขายได้คลิปเดียว\n"
        "- ช้างเผือกที่แท้จริงต้องมีออเดอร์ซ้ำ (`distinct_orders >= 4`) และค่าคอมรวม/ต่อคลิปเด่นชัด (เช่น พรม 3D 34 ออเดอร์จาก 2 คลิปที่ขายออก)\n"
        f"1. ข้อมูลสินค้าของช่องตอนนี้ ({len(compact_channel)} รายการ): {json.dumps(compact_channel, ensure_ascii=False)}\n"
        f"2. รายชื่อสินค้าใหม่ตัวเต็งจาก Kalodata ({len(compact_cands)} รายการ): {json.dumps(compact_cands, ensure_ascii=False)}\n"
        f"3. กลยุทธ์สินค้าใหม่ที่ต้องการ: {strat_desc}\n\n"
        "คำสั่ง (ตอบเป็น JSON เท่านั้น ห้ามมีข้อความอื่น):\n"
        "- `ai_critique`: อาร์เรย์ข้อความสั้นๆ 3 ข้อ วิจารณ์พอร์ตช่องตรงๆ เช่น ตัวไหน GMV สูงแต่ค่าคอม%ต่ำมากเป็นความเสี่ยง, ตัวไหนเป็นช้างเผือกต่อคลิป, และทำไมถึงเลือกสินค้าใหม่ชุดนี้\n"
        "- `promote_skus`: เสนอไม่เกิน 3 ตัวที่ระบบคัดเป็น dropped แต่อาจคุ้มค่าแก่การดันต่อ พร้อมเหตุผลสั้นๆ: [{\"sku\":\"...\",\"reason\":\"...\"}]\n"
        "- `drop_skus`: เสนอไม่เกิน 3 ตัวที่อยู่ใน winners แต่คอมต่อชิ้น/ประสิทธิภาพแย่กว่าตัวอื่นจริง: [{\"sku\":\"...\",\"reason\":\"...\"}]\n"
        f"- `selected_new_stars`: คัด SKU จากรายการข้อ 2 มาให้ครบ {desired_count} ตัวไม่ซ้ำกัน ตามกลยุทธ์ที่สั่ง พร้อมเหตุผลตรงๆ: [{{\"sku\":\"...\",\"reason\":\"...\"}}]"
    )

    client = genai.Client(api_key=api_key)
    for model_name in ("gemini-2.5-flash", "gemini-3-flash-preview", "gemini-2.5-flash-lite"):
        try:
            resp = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.25,
                    response_mime_type="application/json",
                ),
            )
            parsed = json.loads(resp.text or "{}")
            critique = [str(x).strip() for x in (parsed.get("ai_critique") or []) if str(x).strip()][:4]

            # Apply bounded promote/drop overrides cleanly
            winners_by_sku = {p["sku"]: p for p in winners if p.get("sku")}
            dropped_by_sku = {p["sku"]: p for p in dropped if p.get("sku")}

            for promo in (parsed.get("promote_skus") or [])[:3]:
                psku = str(promo.get("sku") or "").strip()
                if psku in dropped_by_sku:
                    obj = dropped_by_sku[psku]
                    d_ord = int(obj.get("distinct_orders") or obj.get("orders") or 0)
                    p_clips = int(obj.get("posted_clips") or 0)
                    h_pct = float(obj.get("hit_rate_pct") or 0)
                    # Strictly enforce anti-fluke repeatability rule even on AI promotions
                    if not (d_ord >= 4 or (p_clips > 0 and p_clips <= 3 and h_pct >= 50.0)):
                        continue
                    dropped_by_sku.pop(psku)
                    obj["tier"] = "winner"
                    obj["grade"] = "A_GEM"
                    obj["starred"] = True
                    obj["reason"] = f"🤖 AI ดึงกลับ: {promo.get('reason') or 'น่าลองดันต่อ'}"
                    winners.append(obj)
                    dropped = [d for d in dropped if d.get("sku") != psku]

            for dr in (parsed.get("drop_skus") or [])[:3]:
                dsku = str(dr.get("sku") or "").strip()
                if dsku in winners_by_sku and len(winners) > 5:
                    obj = winners_by_sku.pop(dsku)
                    # Do not allow AI to drop top-3 earners
                    if obj.get("grade") == "S" and obj in winners[:5]:
                        continue
                    obj["tier"] = "dropped"
                    obj["grade"] = "AI_DROP"
                    obj["starred"] = False
                    obj["reason"] = f"🤖 AI ปัดตก: {dr.get('reason') or 'ประสิทธิภาพไม่คุ้มโควตา'}"
                    dropped.insert(0, obj)
                    winners = [w for w in winners if w.get("sku") != dsku]

            # Map selected_new_stars cleanly from vetted candidate_pool
            cand_map = {c["sku"]: c for c in candidate_pool}
            ai_stars: list[dict[str, Any]] = []
            seen_s: set[str] = set()
            for sel in parsed.get("selected_new_stars") or []:
                ssku = str(sel.get("sku") or "").strip()
                if ssku in cand_map and ssku not in seen_s and len(ai_stars) < desired_count:
                    item = dict(cand_map[ssku])
                    if sel.get("reason"):
                        item["reason"] = f"🤖 AI คัด: {sel['reason']}"
                    ai_stars.append(item)
                    seen_s.add(ssku)

            for fallback_item in picked_stars + backup_stars:
                if len(ai_stars) >= desired_count:
                    break
                if fallback_item["sku"] not in seen_s:
                    ai_stars.append(fallback_item)
                    seen_s.add(fallback_item["sku"])

            return {
                "ai_used": True,
                "ai_model": model_name,
                "ai_critique": critique,
                "winners": winners,
                "dropped": dropped,
                "new_stars": ai_stars,
            }
        except Exception:
            continue

    return {"ai_used": False, "ai_model": "", "ai_critique": [], "winners": winners, "dropped": dropped, "new_stars": picked_stars}


# ======================================================================================
# Marketing Angles Engine (skill: ~/.gemini/config/skills/tiktok-shop-marketing)
# Each product -> visual proof score + #1 buyer objection + 3 film-ready angles A1/A2/A3
# ======================================================================================
ANGLE_CACHE_FILE = BASE_DIR / "product_angles.json"
VISUAL_PROOF_MIN_SCORE = 3
_ANGLE_LOCK = threading.Lock()
_ANGLE_MODELS = ("gemini-2.5-flash", "gemini-3-flash-preview", "gemini-2.5-flash-lite")

_FORBIDDEN_CLAIM_REPLACEMENTS: list[tuple[str, str]] = [
    ("หายปวดหลัง", "นอนสบายขึ้น"),
    ("ไม่ปวดหลัง", "นอนนุ่มสบาย"),
    ("ลดอาการปวดหลัง", "รองรับสรีระได้ดี"),
    ("ลดปวดหลัง", "นอนนุ่มสบาย"),
    ("แก้ปวดหลัง", "นอนสบายขึ้น"),
    ("คนปวดหลัง", "คนชอบนอนนุ่ม"),
    ("ปวดหลัง", "ที่นอนแข็ง"),
    ("ปวดเอว", "นอนไม่สบาย"),
    ("ปวดคอ", "นอนไม่สบาย"),
    ("ออฟฟิศซินโดรม", "นั่งนานเมื่อยล้า"),
    ("ลดอาการปวด", "สัมผัสนุ่มสบาย"),
    ("อาการปวด", "ความไม่สบายตัว"),
    ("แก้ภูมิแพ้", "ลดฝุ่นสะสม"),
    ("ภูมิแพ้", "ฝุ่นสะสม"),
    ("ฆ่าเชื้อแบคทีเรีย", "ดูแลความสะอาด"),
    ("ฆ่าเชื้อโรค", "ทำความสะอาด"),
    ("ฆ่าเชื้อ", "ทำความสะอาด"),
    ("รักษาโรค", "ดูแลตัวเอง"),
    ("หายขาด", "ดีขึ้นมาก"),
    ("ชีวิตเปลี่ยน", "สบายขึ้นทันตา"),
    ("เปลี่ยนชีวิต", "สบายขึ้นทันตา"),
    ("ดีที่สุด", "ตัวตึง"),
    ("ดูแลสุขภาพการนอน", "อยากนอนสบายขึ้น"),
    ("สุขภาพการนอน", "การนอนที่สบาย"),
    ("ประโยชน์ต่อสุขภาพ", "ความสบายตอนใช้งาน"),
    ("ลงทุนเพื่อสุขภาพ", "ลงทุนเพื่อความสบาย"),
    ("ตะกร้าส้ม", "ตะกร้าเหลือง"),
    ("Shopee", "TikTok Shop"),
    ("Lazada", "TikTok Shop"),
]
_THAI_HUNDREDS = {1: "ร้อย", 2: "สองร้อย", 3: "สามร้อย", 4: "สี่ร้อย", 5: "ห้าร้อย", 6: "หกร้อย", 7: "เจ็ดร้อย", 8: "แปดร้อย", 9: "เก้าร้อย", 10: "พัน"}
# Dead openers or overused clichés that made recent clips sound repetitive
_WEAK_HOOK_RE = re.compile(r"^\s*(?:เบื่อไหม|ใครเคย|ใครที่|บ้านไหน|สวัสดี|วันนี้มารีวิว|วันนี้จะมารีวิว|อยากให้|หา.{0,30}อยู่)|ใช่ไหม\s*[?!]*\s*$|จะรอดหรือร่วง")


def sanitize_marketing_text(text: Any, limit: int = 220) -> str:
    """Strip TikTok Shop TH forbidden claims (medical/pain/germ/100%), cliché phrases, and competitor apps from AI marketing copy."""
    s = normalize_thai_text(str(text or ""))
    for bad, good in _FORBIDDEN_CLAIM_REPLACEMENTS:
        s = s.replace(bad, good)
    s = re.sub(r"\.{0,3}\s*จะรอดหรือร่วง\s*(?:มาดู)?\s*[?!]*", " พิสูจน์ของจริงสดๆ!", s)
    s = re.sub(r"(?:ราคาแค่|ราคา|แค่|งบ)?หลักสิบ(?:บาท)?", "งบไม่ถึงร้อย", s)
    s = re.sub(r"99\.9\s*%|100\s*%", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:limit].strip()


def price_cushion_phrase(price: Any) -> str:
    """Safe price ceiling wording (never the exact price): 79 -> งบไม่ถึงร้อย, 259 -> สามร้อยมีทอน."""
    try:
        p = float(price or 0)
    except (TypeError, ValueError):
        p = 0.0
    if p <= 0:
        return "ราคาโปรคุ้มมาก"
    if p < 100:
        return "งบไม่ถึงร้อย"
    n = int(math.floor(p / 100.0)) + 1
    return f"{_THAI_HUNDREDS[n]}มีทอน" if n in _THAI_HUNDREDS else "ราคาโปรคุ้มมาก"


def short_product_label(name: Any, max_len: int = 26) -> str:
    """Short spoken product label from a long shop title (cut on word boundaries)."""
    s = clean_product_query(str(name or ""))
    s = re.sub(r"[^\u0E00-\u0E7Fa-zA-Z0-9 ]+", " ", s)
    words = s.split()
    out = ""
    for w in words:
        cand = f"{out} {w}".strip()
        if len(cand) > max_len:
            break
        out = cand
    return out or (words[0][:max_len] if words else "สินค้านี้")


def _angle_key(prod: dict[str, Any]) -> str:
    return extract_clean_sku(prod.get("sku")) or norm_token_key(prod.get("name", ""))


def load_angle_cache() -> dict[str, dict[str, Any]]:
    with _ANGLE_LOCK:
        if not ANGLE_CACHE_FILE.exists():
            return {}
        try:
            data = json.loads(ANGLE_CACHE_FILE.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}


def save_angle_cache(updates: dict[str, dict[str, Any]]) -> None:
    if not updates:
        return
    with _ANGLE_LOCK:
        cur: dict[str, Any] = {}
        if ANGLE_CACHE_FILE.exists():
            try:
                loaded = json.loads(ANGLE_CACHE_FILE.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    cur = loaded
            except Exception:
                cur = {}
        cur.update(updates)
        ANGLE_CACHE_FILE.write_text(json.dumps(cur, ensure_ascii=False, indent=1), encoding="utf-8")


def heuristic_product_angles(prod: dict[str, Any]) -> dict[str, Any]:
    """Deterministic fallback angles (used when AI is off/unavailable). Never gates a product (score 3)."""
    short = short_product_label(prod.get("name"))
    cushion = price_cushion_phrase(prod.get("price"))
    seed = sum(ord(ch) for ch in (str(prod.get("sku") or "") + short))
    a1_hooks = [
        f"พิสูจน์{short}ของจริงคาตา! {cushion}ตรงปกแค่ไหนดู!",
        f"แกะเช็กเนื้อจริง! {short} {cushion} งานจริงเป็นยังไงมาดู!",
        f"เทสต์{short}ให้ดูสดๆ ไม่ตัดต่อ! {cushion}คุ้มจริงไหมดูเอง!",
        f"ซูมดูความหนา{short}ชัดๆ! {cushion}ได้งานระดับนี้เลย!",
        f"ลองกดลองดึง{short}หน้ากล้องสดๆ! {cushion}เอาอยู่ไหมดู!",
        f"ของจริงมาแล้ว! เช็กสเปก{short} {cushion} ให้เห็นกันจะๆ!",
    ]
    a2_hooks = [
        f"บุกคลังตัดสต็อก! {short} {cushion}",
        f"ยกเซ็ตหน้าคลัง! {short} {cushion} คุ้มกว่าที่คิด",
        f"รอบนี้ปล่อยราคาโรงงาน! {short} {cushion} ดูขนาดจริง!",
        f"งบเท่านี้ได้ครบขนาดนี้เลย! {short} {cushion} เช็กในตะกร้า!",
    ]
    a3_hooks = [
        f"ปัญหาเดิมทุกวันจบได้ด้วย{short} ดูนี่!",
        f"เปลี่ยนมุมบ้านให้จบในชิ้นเดียว! ลองวาง{short}ให้ดูสดๆ!",
        f"ของใช้ติดบ้านที่หยิบใช้ทุกวัน! ดูจังหวะใช้งาน{short}จริง!",
        f"เทียบบ้านก่อนใช้กับหลังวาง{short} ความต่างเห็นชัดทันที!",
    ]
    angles = [
        {
            "code": "A1",
            "label": "สายขี้ระแวง กลัวของจกตา",
            "persona": "คนที่เคยโดนของไม่ตรงปก ต้องเห็นการทดสอบจริงก่อนถึงจะยอมกดตะกร้า",
            "buying_behavior": "ปัดทิ้งทันทีถ้าเปิดมาด้วยคำชม เชื่อภาพทดสอบหนักๆ คาตามากกว่าคำพูด และชอบอ่านคอมเมนต์ก่อนซื้อ",
            "objection": f"กลัวว่า{short}ของจริงจะไม่ตรงปก คุณภาพไม่ถึง",
            "hook_th": a1_hooks[seed % len(a1_hooks)],
            "visual_proof": "เปิดเฟรมแรกด้วยการทดสอบที่หนักที่สุดของสินค้า (เท/กด/ดึง/ใช้งานจริง) ให้เห็นผลภายใน 2 วินาที",
            "trigger_words": ["พิสูจน์สด", "ของจริง", "ตรงปก"],
            "style_nums": [3, 9, 1],
            "hardsell_style": 16,
        },
        {
            "code": "A2",
            "label": "สายล่าโปรคุ้ม เห็นราคาแล้วมือลั่น",
            "persona": "คนชอบของคุ้ม ไม่ได้ตั้งใจหาซื้อแต่แพ้ความคุ้มค่าและภาพสต็อกแน่นๆ",
            "buying_behavior": "ตัดสินใจเร็วเมื่อเห็นจำนวนชิ้นหรือความคุ้มชัดๆ แพ้คำว่าตัดสต็อก ราคาโรงงาน และกดตะกร้าทันทีถ้ารู้สึกคุ้ม",
            "objection": "กลัวซื้อแพงกว่าที่อื่น หรือได้ของน้อยกว่าที่คิด",
            "hook_th": a2_hooks[seed % len(a2_hooks)],
            "visual_proof": "โชว์ปริมาณหรือจำนวนชิ้นจริงเต็มเฟรม เทียบขนาดกับมือหรือของรอบตัว ให้เห็นความคุ้มทันที",
            "trigger_words": ["ตัดสต็อก", "ราคาโรงงาน", "คุ้ม"],
            "style_nums": [5, 11, 8],
            "hardsell_style": 13,
        },
        {
            "code": "A3",
            "label": "สายแก้ปัญหาในบ้านทุกวัน",
            "persona": "คนที่รำคาญปัญหาเดิมในชีวิตประจำวัน อยากได้ตัวช่วยที่จบปัญหาได้จริงในบ้านตัวเอง",
            "buying_behavior": "หยุดดูเมื่อเห็นปัญหาที่ตัวเองเจออยู่ทุกวัน เชื่อคนใช้จริงในบ้านจริง และซื้อเมื่อเห็นก่อน-หลังชัดเจน",
            "objection": "ไม่แน่ใจว่าจะช่วยแก้ปัญหาในบ้านตัวเองได้จริง หรือใช้ยุ่งยาก",
            "hook_th": a3_hooks[seed % len(a3_hooks)],
            "visual_proof": "เปิดด้วยภาพปัญหาจริงในบ้าน แล้วตัดเห็นผลหลังใช้สินค้าภายในไม่กี่วินาที",
            "trigger_words": ["จบปัญหา", "ใช้ง่าย", "ในบ้าน"],
            "style_nums": [2, 1, 6],
            "hardsell_style": 14,
        },
    ]
    return {
        "visual_proof_score": 3,
        "visual_proof_note": "ประเมินด้วยสูตรสำรอง (ยังไม่ได้ให้ AI วิเคราะห์สินค้าตัวนี้)",
        "core_objection": angles[0]["objection"],
        "angles": angles,
        "source": "heuristic",
        "model": "",
        "updated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _validate_angle_entry(raw: Any, prod: dict[str, Any]) -> dict[str, Any] | None:
    """Normalize an AI/cached angle entry: exactly 3 angles, valid style numbers, sanitized claims."""
    if not isinstance(raw, dict):
        return None
    angles_in = raw.get("angles") if isinstance(raw.get("angles"), list) else []
    if len(angles_in) < 3:
        return None
    base = heuristic_product_angles(prod)
    out_angles: list[dict[str, Any]] = []
    for i, code in enumerate(ANGLE_CODES):
        a = angles_in[i] if isinstance(angles_in[i], dict) else {}
        fb = base["angles"][i]
        styles: list[int] = []
        for v in a.get("style_nums") or []:
            try:
                n = int(v)
            except (TypeError, ValueError):
                continue
            if n in UGC_STYLE_NUMS and n not in styles:
                styles.append(n)
        for n in fb["style_nums"]:
            if len(styles) >= 3:
                break
            if n not in styles:
                styles.append(n)
        try:
            hs = int(a.get("hardsell_style"))
        except (TypeError, ValueError):
            hs = fb["hardsell_style"]
        if hs not in HARDSELL_STYLE_NUMS:
            hs = fb["hardsell_style"]
        triggers = [sanitize_marketing_text(t, 28) for t in (a.get("trigger_words") or []) if str(t).strip()][:4]
        raw_hook_text = normalize_thai_text(str(a.get("hook_th") or ""))
        hook = sanitize_marketing_text(raw_hook_text, 140)
        if not hook or _WEAK_HOOK_RE.search(raw_hook_text) or _WEAK_HOOK_RE.search(hook):
            hook = fb["hook_th"]
        out_angles.append({
            "code": code,
            "label": sanitize_marketing_text(a.get("label") or fb["label"], 70),
            "persona": sanitize_marketing_text(a.get("persona") or fb["persona"], 220),
            "buying_behavior": sanitize_marketing_text(a.get("buying_behavior") or fb["buying_behavior"], 260),
            "objection": sanitize_marketing_text(a.get("objection") or fb["objection"], 200),
            "hook_th": hook,
            "visual_proof": sanitize_marketing_text(a.get("visual_proof") or fb["visual_proof"], 220),
            "trigger_words": [t for t in triggers if t] or fb["trigger_words"],
            "style_nums": styles[:3],
            "hardsell_style": hs,
        })
    try:
        vps = int(round(float(raw.get("visual_proof_score"))))
    except (TypeError, ValueError):
        vps = 3
    return {
        "visual_proof_score": max(1, min(5, vps)),
        "visual_proof_note": sanitize_marketing_text(raw.get("visual_proof_note") or "", 220),
        "core_objection": sanitize_marketing_text(raw.get("core_objection") or out_angles[0]["objection"], 200),
        "angles": out_angles,
        "source": str(raw.get("source") or "ai"),
        "model": str(raw.get("model") or ""),
        "updated_at": str(raw.get("updated_at") or dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    }


def _build_angle_prompt(compact_products: list[dict[str, Any]]) -> str:
    ugc_list = ", ".join(f"{s['num']}={s['short']}" for s in STYLES_16 if s["role"] == "ugc")
    hard_list = ", ".join(f"{s['num']}={s['short']}" for s in STYLES_16 if s["role"] == "hardsell")
    return (
        "คุณคือ TikTok Shop Thailand Marketing Strategist (สกิล tiktok-shop-marketing) ให้ช่อง Affiliate ที่ทำคลิป AI แนวตั้ง 10 วินาที "
        "ขายของใช้ในบ้านให้คนไทย วิเคราะห์ตรงไปตรงมา ห้ามอวย ห้ามแต่งข้อมูลที่ไม่รู้\n"
        "วิเคราะห์สินค้าทุกตัวตาม 5 ชั้น:\n"
        "1) visual_proof_score (1-5): ถ้าทำคลิป 10 วินาที จะเห็นผลพิสูจน์ด้วยตาภายใน 2 วินาทีแรกได้ชัดแค่ไหน "
        "(5 = เห็นชัดมาก เช่น เทน้ำแล้วซึมหาย, 3 = พอทำได้ถ้าคิดช็อตดีๆ, 1 = ดูด้วยตาไม่ออก ต้องอธิบายยาว) พร้อม visual_proof_note สั้นๆ ว่าทำไม\n"
        "2) core_objection: เหตุผลอันดับ 1 ที่คนไทยไถฟีดแล้วไม่กดซื้อสินค้านี้ (กลัวจกตา บาง เล็กกว่าปก ลื่น พังไว ใช้ยาก แพงไป ฯลฯ)\n"
        "3) angles: 3 มุมขาย A1 A2 A3 สำหรับกลุ่มลูกค้าที่ต่างกันจริง (ห้ามกว้างแบบ 'ผู้หญิง 20-50 ปี') แต่ละมุมมีฟิลด์:\n"
        "   label (ชื่อกลุ่มสั้นๆ), persona (ใครซื้อ ซื้อเพื่ออะไร), buying_behavior (หยุดนิ้วเพราะอะไร เชื่อคนพูดแบบไหน ตัดสินใจซื้อยังไง), "
        "objection (สิ่งที่กลุ่มนี้กลัว), hook_th (ประโยคเปิด 1.5 วิแรก ภาษาพูดไทย ห้ามขึ้นต้นด้วย เบื่อไหม/ใครเคย/สวัสดี/วันนี้มารีวิว), "
        "visual_proof (ช็อตเปิด 2 วิแรกที่ต้องถ่ายให้เห็น), trigger_words (2-4 คำกระตุ้น), "
        f"style_nums (เลือก 3 แนวจาก 1-12 ที่เข้ากับมุมนี้: {ugc_list}), hardsell_style (1 แนวจาก 13-16: {hard_list})\n"
        "กฎเหล็ก: ห้ามเคลมการแพทย์/อาการปวด/ภูมิแพ้/ฆ่าเชื้อ/99.9%/100% และห้ามวางมุมเป็นเรื่องสุขภาพ (เช่น 'ดูแลสุขภาพการนอน') ให้พูดเรื่องสัมผัส/ความสบาย/ความคุ้มแทน, ห้ามพูดราคาเป๊ะหรือคำว่าหลักสิบ "
        "ให้ใช้เพดานราคา (งบไม่ถึงร้อย / สองร้อยมีทอน / สามร้อยมีทอน), ห้ามพูดถึงแอปอื่น\n"
        "กฎฮุค (สำคัญมาก): hook_th ต้องเป็นประโยคบอกเล่า/ท้าทาย/ช็อตราคาที่เห็นภาพทันที ห้ามเป็นคำถามลอยๆ ลงท้าย 'ใช่ไหม' "
        "และห้ามขึ้นต้นด้วย 'บ้านไหน...', 'ใครที่...', 'หา...อยู่', 'อยากให้...' "
        "A1 ต้องเป็นมุม 'พิสูจน์ของจริงด้วยตาตั้งแต่วินาทีแรก (Visual Stress-Test / เช็กงานจริงตรงปก)' เสมอ "
        "แต่ห้ามใช้วลีซ้ำซากอย่าง 'จะรอดหรือร่วง' เด็ดขาด! ให้คิด hook_th แยกตามวิธีทดสอบจริงของสินค้าแต่ละชิ้นให้ไม่ซ้ำแพตเทิร์นกัน "
        "ตัวอย่างฮุคที่หลากหลายและขายได้จริง: 'เทน้ำสาดให้ดูสดๆ ซับไวแค่ไหนดูเอง!', 'บุกคลังตัดสต็อก! ยกเซ็ตนี้แบงก์ร้อยมีทอน', "
        "'กดลงไปสุดมือ...ดูความเด้งคืนตัวชัดๆ!', 'แกะเทียบขนาดจริงคาตา สองร้อยมีทอนได้ไซส์นี้เลย!'\n"
        "ตอบเป็น JSON เท่านั้น: {\"products\":[{\"key\":\"...\",\"visual_proof_score\":4,\"visual_proof_note\":\"...\","
        "\"core_objection\":\"...\",\"angles\":[{...A1...},{...A2...},{...A3...}]}]}\n"
        f"สินค้า ({len(compact_products)} รายการ): {json.dumps(compact_products, ensure_ascii=False)}"
    )


def _ai_generate_angles(batch: list[tuple[str, dict[str, Any]]], api_key: str) -> tuple[dict[str, dict[str, Any]], str]:
    """Ask Gemini for angles of up to ~10 products in one JSON call; returns ({key: raw_entry}, model_used)."""
    try:
        from google import genai
        from google.genai import types
    except Exception:
        return {}, ""
    compact = [
        {
            "key": k,
            "name": str(p.get("name") or "")[:100],
            "cat": str(p.get("cat") or ""),
            "price_band": price_cushion_phrase(p.get("price")),
            "tier": str(p.get("tier") or ""),
        }
        for k, p in batch
    ]
    prompt = _build_angle_prompt(compact)
    try:
        client = genai.Client(api_key=api_key, http_options=types.HttpOptions(timeout=120_000))
    except Exception:
        try:
            client = genai.Client(api_key=api_key)
        except Exception:
            return {}, ""
    for model_name in _ANGLE_MODELS:
        try:
            resp = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(temperature=0.4, response_mime_type="application/json"),
            )
            parsed = json.loads(resp.text or "{}")
            items = parsed.get("products") if isinstance(parsed, dict) else parsed
            out: dict[str, dict[str, Any]] = {}
            for it in items or []:
                if isinstance(it, dict) and str(it.get("key") or ""):
                    out[str(it["key"])] = it
            if out:
                return out, model_name
        except Exception:
            continue
    return {}, ""


def build_product_angles(
    products: list[dict[str, Any]],
    use_ai: bool = True,
    explicit_key: str = "",
    force: bool = False,
) -> dict[str, dict[str, Any]]:
    """Attach `marketing` (visual proof + objection + 3 angles) to every product. Cached in product_angles.json."""
    cache = load_angle_cache()
    results: dict[str, dict[str, Any]] = {}
    need_ai: list[tuple[str, dict[str, Any]]] = []
    for prod in products:
        k = _angle_key(prod)
        if not k:
            continue
        entry: dict[str, Any] | None = None
        if not force and isinstance(prod.get("marketing"), dict):
            cand = _validate_angle_entry(prod["marketing"], prod)
            if cand and (cand["source"] == "ai" or not use_ai):
                entry = cand
        if entry is None and not force and k in cache:
            cand = _validate_angle_entry(cache[k], prod)
            if cand and (cand["source"] == "ai" or not use_ai):
                entry = cand
        if entry is None:
            if use_ai:
                need_ai.append((k, prod))
                continue
            entry = heuristic_product_angles(prod)
        results[k] = entry
        prod["marketing"] = entry

    if need_ai:
        api_key = _discover_gemini_api_key(explicit_key)
        ai_out: dict[str, dict[str, Any]] = {}
        model_by_key: dict[str, str] = {}
        if api_key:
            chunks = [need_ai[i : i + 10] for i in range(0, len(need_ai), 10)]
            with concurrent.futures.ThreadPoolExecutor(max_workers=min(3, len(chunks))) as ex:
                for out, model_used in ex.map(lambda ch: _ai_generate_angles(ch, api_key), chunks):
                    for kk, vv in out.items():
                        ai_out[kk] = vv
                        model_by_key[kk] = model_used
        new_cache: dict[str, dict[str, Any]] = {}
        for k, prod in need_ai:
            entry = None
            if k in ai_out:
                raw = dict(ai_out[k])
                raw["source"] = "ai"
                raw["model"] = model_by_key.get(k, "")
                raw["updated_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                entry = _validate_angle_entry(raw, prod)
                if entry:
                    new_cache[k] = entry
            if entry is None:
                entry = heuristic_product_angles(prod)
            results[k] = entry
            prod["marketing"] = entry
        save_angle_cache(new_cache)
    return results


def gate_new_stars_by_visual_proof(
    picked: list[dict[str, Any]],
    backups: list[dict[str, Any]],
    desired_count: int,
    use_ai: bool = True,
    explicit_key: str = "",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    """Replace new-star candidates whose visual proof score < 3 with backups that pass (keeps count when possible)."""
    notes: list[str] = []
    if not picked or desired_count <= 0:
        return picked, backups, notes
    build_product_angles(picked, use_ai=use_ai, explicit_key=explicit_key)
    kept = [p for p in picked if int((p.get("marketing") or {}).get("visual_proof_score") or 3) >= VISUAL_PROOF_MIN_SCORE]
    rejected = [p for p in picked if p not in kept]
    for r in rejected:
        m = r.get("marketing") or {}
        notes.append(
            f"🧠 ตัด '{str(r.get('name') or '')[:42]}' ออกจากสินค้าใหม่: ภาพพิสูจน์ {m.get('visual_proof_score')}/5 — {m.get('visual_proof_note') or 'คลิป 10 วิดูด้วยตาไม่ออก'}"
        )
    used = {p.get("sku") for p in picked}
    pool = [b for b in backups if b.get("sku") not in used]
    idx = 0
    while len(kept) < desired_count and idx < len(pool) and idx < 15:
        batch = pool[idx : idx + max(1, desired_count - len(kept))]
        idx += len(batch)
        build_product_angles(batch, use_ai=use_ai, explicit_key=explicit_key)
        for b in batch:
            score = int((b.get("marketing") or {}).get("visual_proof_score") or 3)
            if score >= VISUAL_PROOF_MIN_SCORE and len(kept) < desired_count:
                b["reason"] = f"{b.get('reason') or ''} · 🧠 ผ่านเกณฑ์ภาพพิสูจน์ {score}/5".strip(" ·")
                kept.append(b)
    if len(kept) < desired_count and rejected:
        # Not enough passing backups: keep the strongest rejected ones but flag them honestly
        for r in sorted(rejected, key=lambda x: -int((x.get("marketing") or {}).get("visual_proof_score") or 0)):
            if len(kept) >= desired_count:
                break
            r["reason"] = f"{r.get('reason') or ''} · ⚠️ ภาพพิสูจน์อ่อน ({(r.get('marketing') or {}).get('visual_proof_score')}/5) ต้องคิดช็อตเปรียบเทียบให้ดี".strip(" ·")
            kept.append(r)
            notes.append(f"⚠️ ไม่มีตัวสำรองที่ผ่านเกณฑ์พอ จึงคง '{str(r.get('name') or '')[:42]}' ไว้แบบมีคำเตือน")
    kept_skus = {p.get("sku") for p in kept}
    return kept, [b for b in backups if b.get("sku") not in kept_skus], notes


def get_angle(marketing: dict[str, Any] | None, code: str) -> dict[str, Any]:
    for a in (marketing or {}).get("angles") or []:
        if a.get("code") == code:
            return a
    angles = (marketing or {}).get("angles") or []
    return angles[0] if angles else {}


def angle_sequence_code(marketing: dict[str, Any] | None, best_angle: str, occ_idx: int) -> str:
    """Rotate A1->A2->A3 per product; when a best angle is proven by sales it takes 4 of every 6 slots."""
    codes = [a.get("code") for a in (marketing or {}).get("angles") or [] if a.get("code") in ANGLE_CODES] or list(ANGLE_CODES)
    if best_angle in codes and len(codes) >= 2:
        others = [c for c in codes if c != best_angle]
        pattern = [best_angle, others[0], best_angle, best_angle, others[-1], best_angle]
        return pattern[occ_idx % len(pattern)]
    return codes[occ_idx % len(codes)]


def preferred_styles_for(prod: dict[str, Any], code: str, role: str) -> list[int]:
    a = get_angle(prod.get("marketing"), code)
    if role == "hardsell":
        hs = a.get("hardsell_style")
        return [hs] if hs in HARDSELL_STYLE_NUMS else []
    return [n for n in (a.get("style_nums") or []) if n in UGC_STYLE_NUMS]


def angle_slot_fields(prod: dict[str, Any], code: str) -> dict[str, Any]:
    """Fields copied into a schedule slot so the clip maker gets the exact persona/objection/hook to film."""
    m = prod.get("marketing") or {}
    a = get_angle(m, code)
    if not a:
        return {"angle_code": ""}
    brief = (
        f"🎯 มุม {code} {a.get('label', '')} · ลูกค้า: {a.get('persona', '')} · "
        f"ข้อกังวลที่ต้องพิสูจน์: {a.get('objection', '')} · เปิดคลิป 2 วิแรก: {a.get('visual_proof', '')} · "
        f"ฮุค: \"{a.get('hook_th', '')}\""
    )
    return {
        "angle_code": code,
        "angle_label": a.get("label", ""),
        "angle_persona": a.get("persona", ""),
        "angle_behavior": a.get("buying_behavior", ""),
        "angle_objection": a.get("objection", ""),
        "angle_hook": a.get("hook_th", ""),
        "angle_visual": a.get("visual_proof", ""),
        "angle_triggers": list(a.get("trigger_words") or []),
        "angle_brief": brief,
        "angle_source": m.get("source", ""),
        "visual_proof_score": m.get("visual_proof_score", 3),
    }


def _parse_bot_row_times(row: dict[str, Any]) -> list[float]:
    """Upload time + scheduled publish time of a bot post row (epoch seconds)."""
    times: list[float] = []
    base_dt: dt.datetime | None = None
    ts_str = str(row.get("ts") or row.get("date") or "").strip()
    try:
        base_dt = dt.datetime.fromisoformat(ts_str[:19]) if len(ts_str) >= 16 else (
            dt.datetime.fromisoformat(ts_str[:10] + "T12:00:00") if len(ts_str) >= 10 else None
        )
    except ValueError:
        base_dt = None
    if base_dt:
        times.append(base_dt.timestamp())
    sched = str(row.get("scheduled_at") or "").strip()
    if sched:
        try:
            times.append(dt.datetime.fromisoformat(sched[:19]).timestamp())
        except ValueError:
            pass
    elif base_dt:
        m = re.search(r"\((\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{2})\)", str(row.get("mode") or ""))
        if m:
            d_, mo_, hh_, mm_ = (int(x) for x in m.groups())
            year = base_dt.year + (1 if mo_ < base_dt.month - 6 else 0)
            try:
                times.append(dt.datetime(year, mo_, d_, hh_, mm_).timestamp())
            except ValueError:
                pass
    return times


def load_bot_post_rows(hist_path: Path | None = None) -> list[dict[str, Any]]:
    """Read Desktop/posted_clip_history.json written by tiktok_auto_pipeline.py (with angle tags when present)."""
    hist_file = hist_path or (Path.home() / "Desktop" / "posted_clip_history.json")
    if not hist_file.is_file():
        return []
    try:
        data = json.loads(hist_file.read_text(encoding="utf-8"))
    except Exception:
        return []
    rows: list[dict[str, Any]] = []
    for row in data if isinstance(data, list) else []:
        if not isinstance(row, dict):
            continue
        sku = extract_clean_sku(row.get("sku") or row.get("filename"))
        angle = str(row.get("angle") or "").upper() or extract_angle_tag(row.get("filename"))
        rows.append({
            "sku": sku,
            "angle": angle if angle in ANGLE_CODES else "",
            "date": str(row.get("date") or row.get("ts") or "")[:10],
            "times": _parse_bot_row_times(row),
            "name": normalize_thai_text(row.get("name") or ""),
        })
    return rows


def attribute_angle_sales(
    items: list[dict[str, Any]],
    hist_path: Path | None = None,
    min_date: str = "",
    max_date: str = "",
    max_gap_hours: float = 36.0,
) -> list[dict[str, Any]]:
    """Match each converting content_id to the nearest same-SKU bot post -> orders/comm per angle A1/A2/A3.

    Must run BEFORE enrich_with_bot_posted_stats (which consumes `_cid_pub_comm`).
    Unmatched orders are reported as `unattributed` (never guessed).
    """
    rows = load_bot_post_rows(hist_path)
    by_sku: dict[str, list[dict[str, Any]]] = {}
    for r in rows:
        if not r["sku"]:
            continue
        if min_date and r["date"] and r["date"] < min_date:
            continue
        if max_date and r["date"] and r["date"] > max_date:
            continue
        by_sku.setdefault(r["sku"], []).append(r)

    for it in items:
        sku = extract_clean_sku(it.get("sku"))
        posts = by_sku.get(sku, []) if sku else []
        stats: dict[str, dict[str, Any]] = {c: {"posted": 0, "converting_clips": 0, "orders": 0, "comm": 0.0} for c in ANGLE_CODES}
        stats["untagged"] = {"posted": 0, "converting_clips": 0, "orders": 0, "comm": 0.0}
        stats["unattributed"] = {"posted": 0, "converting_clips": 0, "orders": 0, "comm": 0.0}
        for p in posts:
            stats[p["angle"] or "untagged"]["posted"] += 1

        cids = {
            c: st for c, st in (it.get("_cid_pub_comm") or {}).items()
            if not (min_date and st.get("pub_date") and st["pub_date"] < min_date)
        }
        pairs: list[tuple[float, str, int]] = []
        for cid, st in cids.items():
            cts = int(st.get("pub_ts") or _decode_content_ts(cid))
            if not cts:
                continue
            for pi, p in enumerate(posts):
                if not p["times"]:
                    continue
                gap_h = min(abs(cts - t) for t in p["times"]) / 3600.0
                if gap_h <= max_gap_hours:
                    pairs.append((gap_h, cid, pi))
        pairs.sort()
        used_c: set[str] = set()
        used_p: set[int] = set()
        assigned: dict[str, str] = {}
        for _gap, cid, pi in pairs:
            if cid in used_c or pi in used_p:
                continue
            used_c.add(cid)
            used_p.add(pi)
            assigned[cid] = posts[pi]["angle"] or "untagged"
        for cid, st in cids.items():
            bucket = assigned.get(cid, "unattributed")
            stats[bucket]["converting_clips"] += 1
            stats[bucket]["orders"] += int(st.get("orders") or 0)
            stats[bucket]["comm"] = round(stats[bucket]["comm"] + float(st.get("comm") or 0), 2)

        best = ""
        best_key = (0.0, 0)
        for c in ANGLE_CODES:
            s = stats[c]
            if s["orders"] > 0 and (s["comm"], s["orders"]) > best_key:
                best, best_key = c, (s["comm"], s["orders"])
        it["angle_stats"] = stats
        it["best_angle"] = best
    return items


def analyze_sales_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """Parse sales file(s) or pasted lines, grade via Real Commission + AI Advisor, and pick New Stars."""
    files = payload.get("files") or []
    text_lines = str(payload.get("text_lines") or "").strip()
    new_stars_count = int(payload.get("new_stars_count") if payload.get("new_stars_count") is not None else 6)
    new_star_strategy = str(payload.get("new_star_strategy") or "balanced").strip()
    use_ai = bool(payload.get("use_ai") if payload.get("use_ai") is not None else True)
    explicit_key = str(payload.get("gemini_api_key") or "").strip()

    warnings: list[str] = []
    parsed_files: list[str] = []
    raw_items: list[dict[str, Any]] = []

    min_order_date = ""
    max_order_date = ""

    if files:
        parsed = parse_sales_files(files)
        raw_items.extend(parsed["products"])
        warnings.extend(parsed["warnings"])
        parsed_files.extend(parsed["parsed_files"])
        min_order_date = parsed.get("min_order_date", "")
        max_order_date = parsed.get("max_order_date", "")
    elif text_lines:
        for line in text_lines.splitlines():
            line = line.strip()
            if not line:
                continue
            unwrapped = unwrap_tiktok_link_or_sku(line)
            name = line
            if line.startswith("http"):
                name = unwrapped.get("og_title") or (f"สินค้า TikTok ({unwrapped['sku']})" if unwrapped.get("sku") else line)
            raw_items.append({
                "sku": unwrapped.get("sku", ""),
                "name": name,
                "link": unwrapped.get("resolved_url", ""),
                "image": unwrapped.get("og_image", ""),
                "gmv": 500.0,
                "orders": 5,
                "comm": 50.0,
                "videos_count": 1,
                "distinct_orders": 5,
                "comm_pct_eff": 10.0,
                "comm_per_video": 50.0,
                "comm_per_order": 10.0,
                "price": 100.0,
            })

    if not raw_items:
        return {
            "ok": False,
            "error": "ไม่พบรายการสินค้าในไฟล์ที่อัปโหลด กรุณาตรวจสอบว่าไฟล์มีคอลัมน์ชื่อสินค้า/รหัสสินค้าและยอดขายครับ",
            "warnings": warnings,
        }

    # Fast local-only SKU/thumbnail pass + enrich with real Desktop bot posted history
    cat_map = load_channel_catalog()
    # Angle attribution must read `_cid_pub_comm` before enrich_with_bot_posted_stats() consumes it
    attribute_angle_sales(raw_items, min_date=min_order_date, max_date=max_order_date)
    raw_items = enrich_with_bot_posted_stats(raw_items, min_order_date, max_order_date)
    local_resolved = [resolve_product_metadata(item, cat_map, use_remote_api=False) for item in raw_items]

    # Dynamic portfolio grading (Real Commission 85/93.5% Pareto + Volume Traps + Hidden Gems)
    winners, dropped = grade_channel_products_smart(local_resolved)

    # Remote Kalodata DB lookup ONLY for winners missing SKU or top winners missing thumbnail
    need_remote_indices: list[int] = []
    img_remote_budget = 15
    for idx, item in enumerate(winners):
        if not item.get("sku"):
            need_remote_indices.append(idx)
        elif not item.get("image") and img_remote_budget > 0:
            need_remote_indices.append(idx)
            img_remote_budget -= 1

    if need_remote_indices:
        with concurrent.futures.ThreadPoolExecutor(max_workers=6) as executor:
            fut_to_idx = {
                executor.submit(resolve_product_metadata, winners[idx], cat_map, True): idx
                for idx in need_remote_indices
            }
            for fut in concurrent.futures.as_completed(fut_to_idx):
                idx = fut_to_idx[fut]
                try:
                    winners[idx] = fut.result()
                except Exception:
                    pass

    save_channel_catalog_entries(winners)

    picked_stars, backup_stars = select_new_stars_from_pool(
        winners, dropped, desired_count=new_stars_count, strategy=new_star_strategy
    )

    ai_result = {
        "ai_used": False,
        "ai_model": "",
        "ai_critique": [],
        "winners": winners,
        "dropped": dropped,
        "new_stars": picked_stars,
    }
    if use_ai:
        ai_result = run_ai_portfolio_advisor(
            winners, dropped, picked_stars, backup_stars,
            strategy=new_star_strategy, desired_count=new_stars_count, explicit_key=explicit_key
        )
        winners = ai_result["winners"]
        dropped = ai_result["dropped"]
        picked_stars = ai_result["new_stars"]
        used_star_skus = {s["sku"] for s in picked_stars}
        backup_stars = [b for b in backup_stars if b["sku"] not in used_star_skus]

    # Marketing skill pass: gate new stars on 2-second visual proof, then 3 angles for every product kept
    picked_stars, backup_stars, marketing_notes = gate_new_stars_by_visual_proof(
        picked_stars, backup_stars, new_stars_count, use_ai=use_ai, explicit_key=explicit_key
    )
    build_product_angles(winners, use_ai=use_ai, explicit_key=explicit_key)
    angle_sources = [str((p.get("marketing") or {}).get("source") or "") for p in winners + picked_stars]

    save_channel_catalog_entries(picked_stars)
    missing_sku_items = [x for x in winners if not extract_clean_sku(x.get("sku"))]

    total_gmv = round(sum(float(x.get("gmv") or 0) for x in winners + dropped), 2)
    total_comm = round(sum(float(x.get("comm") or 0) for x in winners + dropped), 2)
    gems_count = sum(1 for x in winners if x.get("grade") == "A_GEM")

    return {
        "ok": True,
        "parsed_files": parsed_files,
        "warnings": warnings,
        "winners": winners,
        "dropped": dropped,
        "new_stars": picked_stars,
        "backup_new_stars": backup_stars,
        "missing_sku_count": len(missing_sku_items),
        "ai_used": ai_result["ai_used"],
        "ai_model": ai_result["ai_model"],
        "ai_critique": ai_result["ai_critique"],
        "marketing_notes": marketing_notes,
        "angles_ai_count": sum(1 for s in angle_sources if s == "ai"),
        "angles_heuristic_count": sum(1 for s in angle_sources if s == "heuristic"),
        "portfolio_stats": {
            "total_gmv": total_gmv,
            "total_comm": total_comm,
            "gems_count": gems_count,
            "winners_count": len(winners),
            "dropped_count": len(dropped),
        },
    }


def _build_date_list(start_str: str, end_str: str) -> list[dt.date]:
    """Generate inclusive date list from YYYY-MM-DD bounds."""
    s = dt.date.fromisoformat(start_str)
    e = dt.date.fromisoformat(end_str)
    if e < s:
        s, e = e, s
    days: list[dt.date] = []
    cur = s
    while cur <= e and len(days) < 90:
        days.append(cur)
        cur += dt.timedelta(days=1)
    return days


def _format_thai_date(d: dt.date) -> str:
    """Format YYYY-MM-DD into DD/MM/25xx Thai Buddhist year like user's PDF."""
    return f"{d.day:02d}/{d.month:02d}/{d.year + 543}"


def _slot_role_for_index(slot_num: int, clips_per_day: int) -> str:
    """Map daily OMNI slot 1..K to ugc / reply_fallback (OMNI 5) / hardsell (OMNI 6-7)."""
    if clips_per_day >= 5 and slot_num == 5:
        return "reply_fallback"
    if clips_per_day >= 6 and slot_num >= clips_per_day - 1:
        return "hardsell"
    return "ugc"


def recompute_plan_counts(plan: dict[str, Any]) -> dict[str, Any]:
    """Recalculate main_slots, fallback_slots, total_slots, and summary invariants cleanly."""
    days = plan.get("days") or []
    catalog = plan.get("catalog") or []
    cat_by_sku = {str(p.get("sku")): p for p in catalog if p.get("sku")}

    main_counts: dict[str, int] = {}
    fallback_counts: dict[str, int] = {}
    dates_by_sku: dict[str, list[str]] = {}
    total_clips = 0
    omni5_violations: list[dict[str, Any]] = []

    for day in days:
        d_thai = day.get("date_thai") or day.get("date")
        for sl in day.get("slots") or []:
            total_clips += 1
            sku = str(sl.get("sku") or "")
            role = sl.get("slot_role") or "ugc"
            # Sync catalog entry if slot introduces a SKU not yet listed
            if sku and sku not in cat_by_sku:
                new_entry = {
                    "code": sl.get("code") or "F01",
                    "sku": sku,
                    "name": sl.get("name", ""),
                    "image": sl.get("image", ""),
                    "cat": sl.get("cat", ""),
                    "price": float(sl.get("price") or 0),
                    "comm": float(sl.get("comm") or 0),
                    "gmv": 0.0,
                    "orders": 0,
                    "tier": sl.get("tier") or "winner",
                    "starred": True,
                    "link": sl.get("link") or f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH",
                    "safe_filename": build_safe_video_filename(sku, sl.get("name", "")),
                    "reason": sl.get("reason", ""),
                }
                catalog.append(new_entry)
                cat_by_sku[sku] = new_entry

            if role == "reply_fallback":
                fallback_counts[sku] = fallback_counts.get(sku, 0) + 1
            else:
                main_counts[sku] = main_counts.get(sku, 0) + 1
            dates_by_sku.setdefault(sku, [])
            label = f"{d_thai}*" if role == "reply_fallback" else str(d_thai)
            if label not in dates_by_sku[sku]:
                dates_by_sku[sku].append(label)

    # Check strict OMNI 5 invariant & build product stats
    for day in days:
        for sl in day.get("slots") or []:
            if sl.get("slot_role") == "reply_fallback":
                sku = str(sl.get("sku") or "")
                m_cnt = main_counts.get(sku, 0)
                if m_cnt < 2:
                    omni5_violations.append({"date": day.get("date"), "slot_num": sl.get("slot_num"), "sku": sku, "main_slots": m_cnt})

    for prod in catalog:
        sku = str(prod.get("sku") or "")
        ms = main_counts.get(sku, 0)
        fs = fallback_counts.get(sku, 0)
        prod["main_slots"] = ms
        prod["fallback_slots"] = fs
        prod["total_slots"] = ms + fs
        prod["scheduled_dates"] = dates_by_sku.get(sku, [])
        prod["safe_filename"] = build_safe_video_filename(sku, prod.get("name", ""))

    plan["catalog"] = catalog
    plan["summary"] = {
        "total_days": len(days),
        "total_clips": total_clips,
        "winners_count": sum(1 for p in catalog if p.get("tier") == "winner"),
        "new_stars_count": sum(1 for p in catalog if p.get("tier") == "new_star"),
        "free_samples_count": sum(1 for p in catalog if p.get("tier") == "free_sample"),
        "dropped_count": len(plan.get("dropped_products") or []),
        "omni5_invariant_ok": len(omni5_violations) == 0,
        "omni5_violations": omni5_violations,
        "missing_sku_in_slots": sum(
            1 for day in days for sl in (day.get("slots") or []) if not extract_clean_sku(sl.get("sku"))
        ),
    }
    return plan


def generate_monthly_schedule(payload: dict[str, Any]) -> dict[str, Any]:
    """Generate the full 7-Omni x 16-Style monthly schedule respecting the OMNI 5 >= 2 invariant."""
    start_str = str(payload.get("start_date") or "").strip()
    end_str = str(payload.get("end_date") or "").strip()
    if not start_str or not end_str:
        today = dt.date.today()
        first_next = (today.replace(day=1) + dt.timedelta(days=32)).replace(day=1)
        last_next = (first_next + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
        start_str = start_str or first_next.isoformat()
        end_str = end_str or last_next.isoformat()

    dates = _build_date_list(start_str, end_str)
    num_days = len(dates)
    clips_per_day = max(3, min(10, int(payload.get("clips_per_day") or 7)))
    new_star_main_quota = max(1, min(10, int(payload.get("new_star_main_quota") or 5)))

    winners_in = [dict(x) for x in (payload.get("winners") or []) if x.get("name") or x.get("sku")]
    new_stars_in = [dict(x) for x in (payload.get("new_stars") or []) if x.get("name") or x.get("sku")]
    free_in = [dict(x) for x in (payload.get("free_samples") or []) if x.get("name") or x.get("sku")]
    dropped_in = [dict(x) for x in (payload.get("dropped_products") or payload.get("dropped") or [])]

    winners_in.sort(
        key=lambda x: (float(x.get("comm") or 0), float(x.get("gmv") or 0), int(x.get("orders") or 0)),
        reverse=True,
    )

    # Deduplicate by SKU (or normalized name when SKU missing) keeping highest tier priority
    active_list: list[dict[str, Any]] = []
    seen_keys: set[str] = set()
    p_idx, f_idx, n_idx = 0, 0, 0

    for item in free_in:
        key = extract_clean_sku(item.get("sku")) or norm_token_key(item.get("name", ""))
        if not key or key in seen_keys:
            continue
        seen_keys.add(key)
        f_idx += 1
        item["tier"] = "free_sample"
        item["starred"] = True
        item["code"] = f"F{f_idx:02d}"
        active_list.append(item)

    for item in winners_in:
        key = extract_clean_sku(item.get("sku")) or norm_token_key(item.get("name", ""))
        if not key or key in seen_keys:
            continue
        seen_keys.add(key)
        p_idx += 1
        item["tier"] = "winner"
        item["starred"] = True
        item["code"] = f"P{p_idx:02d}"
        active_list.append(item)

    for item in new_stars_in:
        key = extract_clean_sku(item.get("sku")) or norm_token_key(item.get("name", ""))
        if not key or key in seen_keys:
            continue
        seen_keys.add(key)
        n_idx += 1
        item["tier"] = "new_star"
        item["starred"] = True
        item["code"] = f"N{n_idx:02d}"
        active_list.append(item)

    if not active_list:
        return {"ok": False, "error": "ไม่มีสินค้าในแผนที่จะจัดตาราง กรุณาเลือกอย่างน้อย 1 รายการครับ"}

    # Marketing angles (A1/A2/A3) for every product: reuse analysis result / cache, AI only when asked
    build_product_angles(
        active_list,
        use_ai=bool(payload.get("use_ai") or False),
        explicit_key=str(payload.get("gemini_api_key") or ""),
    )

    has_omni5 = clips_per_day >= 5
    slot_roles_pattern = [_slot_role_for_index(s_num, clips_per_day) for s_num in range(1, clips_per_day + 1)]
    main_slot_nums = [i + 1 for i, r in enumerate(slot_roles_pattern) if r != "reply_fallback"]
    main_per_day = len(main_slot_nums)
    total_main_slots = num_days * main_per_day

    # Step A: Allocate guaranteed main-slot quotas
    quotas: dict[str, int] = {}

    def prod_key(prod: dict[str, Any]) -> str:
        return str(prod.get("sku") or prod.get("code") or norm_token_key(prod.get("name", "")))

    remaining = total_main_slots
    # Baseline >= 1 main slot per active product (up to capacity)
    for prod in active_list:
        k = prod_key(prod)
        base = 1 if remaining > 0 else 0
        quotas[k] = base
        remaining -= base

    # Give free samples their target (default 3) first
    for prod in active_list:
        if prod["tier"] == "free_sample" and remaining > 0:
            k = prod_key(prod)
            want = max(1, min(num_days, int(prod.get("target_main_slots") or 3)))
            add = min(remaining, max(0, want - quotas[k]))
            quotas[k] += add
            remaining -= add

    # Give new_stars their target test quota (default 5 main slots -> ~6 total with OMNI 5, matching PDF's N01..N06)
    for prod in active_list:
        if prod["tier"] == "new_star" and remaining > 0:
            k = prod_key(prod)
            custom_q = prod.get("target_main_slots")
            want = int(custom_q) if custom_q else min(new_star_main_quota, max(2, total_main_slots // max(1, len(active_list) + 2)))
            want = max(2 if total_main_slots >= len(active_list) * 2 else 1, min(num_days, want))
            add = min(remaining, max(0, want - quotas[k]))
            quotas[k] += add
            remaining -= add

    # Distribute remaining main slots among Winners weighted by sqrt(sales/orders)
    winner_prods = [p for p in active_list if p["tier"] == "winner"]
    if not winner_prods:
        winner_prods = [p for p in active_list if p["tier"] != "free_sample"] or active_list

    # Honor explicit user-edited target_main_slots on winners first if given
    for prod in winner_prods:
        if prod.get("target_main_slots") and remaining > 0:
            k = prod_key(prod)
            want = max(1, min(num_days, int(prod["target_main_slots"])))
            add = min(remaining, max(0, want - quotas[k]))
            quotas[k] += add
            remaining -= add

    if remaining > 0 and winner_prods:
        # Ensure top winners and Hidden Gems (A_GEM) get at least 2-3 main slots when capacity allows
        if total_main_slots >= len(active_list) + 2:
            for prod in winner_prods:
                k = prod_key(prod)
                min_floor = 3 if (prod.get("grade") == "A_GEM" and num_days >= 10) else 2
                if quotas[k] < min_floor and remaining > 0:
                    add = min(remaining, min_floor - quotas[k])
                    quotas[k] += add
                    remaining -= add

        max_per_prod = num_days if len(active_list) >= main_per_day else num_days * 2
        weights = [
            math.sqrt(max(float(p.get("comm") or 0) * 25.0, float(p.get("gmv") or 0), float(p.get("orders") or 0) * 150.0, 80.0))
            for p in winner_prods
        ]
        while remaining > 0:
            eligible = [
                (idx, p, w)
                for idx, (p, w) in enumerate(zip(winner_prods, weights))
                if quotas[prod_key(p)] < max_per_prod
            ]
            if not eligible:
                max_per_prod += 1
                continue
            _, best_p, _ = min(
                eligible,
                key=lambda item: (quotas[prod_key(item[1])] / max(1e-6, item[2]), item[0]),
            )
            quotas[prod_key(best_p)] += 1
            remaining -= 1

    # Step B: Spread main slots evenly across days (minimizing collisions on same day)
    day_main_buckets: list[list[dict[str, Any]]] = [[] for _ in range(num_days)]
    placed_counts: dict[str, int] = {prod_key(p): 0 for p in active_list}

    # Sort descending by quota so high-frequency winners claim smooth strides first,
    # then interleave across the month cleanly
    prods_for_placement = sorted(active_list, key=lambda p: quotas[prod_key(p)], reverse=True)
    for p in prods_for_placement:
        k = prod_key(p)
        q = quotas[k]
        if q <= 0:
            continue
        stride = num_days / float(q)
        offset = (hash(k) % 7) * 0.13
        used_days_for_p: set[int] = set()
        for occ in range(q):
            ideal = min(num_days - 1, int(round((occ + 0.35 + offset) * stride)) % num_days)
            # Candidate search prioritizing: open capacity AND SKU not yet placed on that day
            best_d = -1
            for radius in range(num_days):
                for sign in (0,) if radius == 0 else (-1, 1):
                    d_idx = (ideal + sign * radius) % num_days
                    if len(day_main_buckets[d_idx]) < main_per_day and d_idx not in used_days_for_p:
                        best_d = d_idx
                        break
                if best_d >= 0:
                    break
            if best_d < 0:
                # Fallback if fewer distinct products than main_per_day
                best_d = min(range(num_days), key=lambda d_i: (len(day_main_buckets[d_i]), sum(1 for x in day_main_buckets[d_i] if prod_key(x) == k)))
            day_main_buckets[best_d].append(p)
            used_days_for_p.add(best_d)
            placed_counts[k] += 1

    # Arrange within each day's main slots: top winners tend toward OMNI 1..2, varied items fill OMNI 3..4 & 6..7
    prod_lookup = {prod_key(p): p for p in active_list}
    day_main_assigned: list[dict[int, dict[str, Any]]] = []
    for d_idx in range(num_days):
        bucket = day_main_buckets[d_idx]
        # Sort winners before new_stars, alternating day rotation on hardsell tail so different items hit OMNI 6-7
        bucket_sorted = sorted(
            bucket,
            key=lambda p: (
                0 if p["tier"] == "free_sample" else (1 if p["tier"] == "winner" else 2),
                -placed_counts[prod_key(p)],
                ((hash(prod_key(p)) + d_idx) % 13),
            ),
        )
        mapping: dict[int, dict[str, Any]] = {}
        for slot_pos, s_num in enumerate(main_slot_nums):
            mapping[s_num] = bucket_sorted[slot_pos]
        day_main_assigned.append(mapping)

    # Step C: Populate OMNI 5 (reply_fallback) STRICTLY from products with guaranteed_main_slots >= 2
    omni5_pool = [
        p
        for p in active_list
        if p["tier"] in ("winner", "new_star") and placed_counts[prod_key(p)] >= 2
    ]
    if not omni5_pool:
        # Only possible on tiny 1-day plans; pick highest main_slots non-free-sample
        omni5_pool = sorted(
            [p for p in active_list if p["tier"] != "free_sample"],
            key=lambda p: placed_counts[prod_key(p)],
            reverse=True,
        ) or active_list

    omni5_assigned: list[dict[str, Any] | None] = [None] * num_days
    omni5_use_counts: dict[str, int] = {prod_key(p): 0 for p in omni5_pool}
    last_omni5_key = ""

    if has_omni5 and omni5_pool:
        for d_idx in range(num_days):
            main_keys_today = {prod_key(p) for p in day_main_assigned[d_idx].values()}
            # Prefer eligible product NOT appearing in main slots today, least used in OMNI 5 so far, not same as yesterday
            ranked = sorted(
                omni5_pool,
                key=lambda p: (
                    1 if prod_key(p) in main_keys_today else 0,
                    1 if prod_key(p) == last_omni5_key else 0,
                    omni5_use_counts[prod_key(p)] / max(1.0, float(placed_counts[prod_key(p)])),
                    omni5_use_counts[prod_key(p)],
                    -placed_counts[prod_key(p)],
                ),
            )
            chosen = ranked[0]
            omni5_assigned[d_idx] = chosen
            ck = prod_key(chosen)
            omni5_use_counts[ck] = omni5_use_counts.get(ck, 0) + 1
            last_omni5_key = ck

    # Step D: Assign non-repeating Styles 1..16 (1..12 for UGC/OMNI 5, 13..16 for OMNI 6..7 Hard-Sell)
    used_styles_by_prod: dict[str, dict[int, int]] = {}
    last_style_by_prod: dict[str, int] = {}

    def pick_style(
        prod: dict[str, Any],
        role: str,
        used_styles_today: set[int],
        slot_offset: int,
        preferred: list[int] | None = None,
    ) -> dict[str, Any]:
        k = prod_key(prod)
        counts = used_styles_by_prod.setdefault(k, {})
        prev_st = last_style_by_prod.get(k, -1)
        candidates = list(HARDSELL_STYLE_NUMS if role == "hardsell" else UGC_STYLE_NUMS)
        # Rotate baseline candidate list per slot offset for natural day-to-day variety
        shift = slot_offset % len(candidates)
        candidates = candidates[shift:] + candidates[:shift]
        pref = [p for p in (preferred or []) if p in candidates]
        ranked = sorted(
            candidates,
            key=lambda st_num: (
                1 if st_num in used_styles_today else 0,
                1 if st_num == prev_st else 0,
                0 if st_num in pref else 1,
                counts.get(st_num, 0),
                pref.index(st_num) if st_num in pref else 99,
            ),
        )
        chosen_num = ranked[0]
        counts[chosen_num] = counts.get(chosen_num, 0) + 1
        last_style_by_prod[k] = chosen_num
        used_styles_today.add(chosen_num)
        return STYLE_BY_NUM[chosen_num]

    angle_occ: dict[str, int] = {}
    days_out: list[dict[str, Any]] = []
    for d_idx, cur_date in enumerate(dates):
        used_styles_today: set[int] = set()
        slots_out: list[dict[str, Any]] = []
        for s_num in range(1, clips_per_day + 1):
            role = slot_roles_pattern[s_num - 1]
            prod = omni5_assigned[d_idx] if role == "reply_fallback" else day_main_assigned[d_idx][s_num]
            assert prod is not None
            pk = prod_key(prod)
            occ = angle_occ.get(pk, 0)
            angle_occ[pk] = occ + 1
            angle_code = angle_sequence_code(prod.get("marketing"), str(prod.get("best_angle") or ""), occ)
            st = pick_style(prod, role, used_styles_today, d_idx + s_num, preferred_styles_for(prod, angle_code, role))
            sku = extract_clean_sku(prod.get("sku"))
            name = str(prod.get("name") or "")
            urgent_badge = ""
            if prod["tier"] == "free_sample":
                urgent_badge = "🎁 ตัวอย่างฟรีรีวิวกลางเดือน"
            slot_obj = {
                "slot_num": s_num,
                "slot_label": f"OMNI {s_num}",
                "slot_role": role,
                "code": prod["code"],
                "sku": sku,
                "name": name,
                "image": str(prod.get("image") or ""),
                "cat": str(prod.get("cat") or ""),
                "price": float(prod.get("price") or 0),
                "comm": float(prod.get("comm") or 0),
                "tier": prod["tier"],
                "starred": bool(prod.get("starred", True)),
                "style_num": st["num"],
                "style_name": st["short"],
                "style_chip_match": st["chip_match"],
                "fallback_note": "ตอบคอมเมนต์แทนได้: ถ้ามีเม้นต์ให้ทำสินค้าตามคำถาม · ถ้าไม่มีให้ทำสินค้าสำรองนี้ (การันตีมีคิวหลักอยู่แล้ว)"
                if role == "reply_fallback"
                else "",
                "urgent_badge": urgent_badge,
                "deadline": str(prod.get("deadline") or ""),
                "brief_note": str(prod.get("brief_note") or ""),
                "link": prod.get("link") or (f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH" if sku else ""),
                "safe_filename": build_safe_video_filename(sku, name, angle_code),
            }
            slot_obj.update(angle_slot_fields(prod, angle_code))
            slots_out.append(slot_obj)

        days_out.append({
            "date": cur_date.isoformat(),
            "date_thai": _format_thai_date(cur_date),
            "slots": slots_out,
        })

    catalog_out: list[dict[str, Any]] = []
    for p in active_list:
        sku = extract_clean_sku(p.get("sku"))
        catalog_out.append({
            "code": p["code"],
            "sku": sku,
            "name": str(p.get("name") or ""),
            "image": str(p.get("image") or ""),
            "cat": str(p.get("cat") or ""),
            "price": float(p.get("price") or 0),
            "comm": float(p.get("comm") or 0),
            "gmv": float(p.get("gmv") or 0),
            "orders": int(p.get("orders") or 0),
            "v_rev": float(p.get("v_rev") or 0),
            "growth": float(p.get("growth") or 0),
            "tier": p["tier"],
            "starred": bool(p.get("starred", True)),
            "deadline": str(p.get("deadline") or ""),
            "brief_note": str(p.get("brief_note") or ""),
            "reason": str(p.get("reason") or ""),
            "link": p.get("link") or (f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH" if sku else ""),
            "safe_filename": build_safe_video_filename(sku, str(p.get("name") or "")),
            "marketing": p.get("marketing") or {},
            "best_angle": str(p.get("best_angle") or ""),
            "angle_stats": p.get("angle_stats") or {},
            "grade": str(p.get("grade") or ""),
        })

    month_label = str(payload.get("month_label") or f"{_format_thai_date(dates[0])} – {_format_thai_date(dates[-1])}")
    title = str(payload.get("title") or f"แผนลงคลิป {month_label} | วันละ {clips_per_day} Omni (16 แนว)")

    plan = {
        "ok": True,
        "plan_id": f"plan_{dates[0].isoformat()}_{int(time.time())}",
        "title": title,
        "month_label": month_label,
        "start_date": dates[0].isoformat(),
        "end_date": dates[-1].isoformat(),
        "clips_per_day": clips_per_day,
        "updated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "catalog": catalog_out,
        "dropped_products": dropped_in,
        "days": days_out,
    }
    return recompute_plan_counts(plan)


def save_uploaded_sample_image(sku: str, raw_image_b64: str) -> str:
    """Normalize uploaded image to compressed JPEG inside ./plan_assets/<sku>.jpg."""
    b64 = raw_image_b64.split(",", 1)[1] if "," in raw_image_b64 else raw_image_b64
    raw = base64.b64decode(b64)
    img = Image.open(io.BytesIO(raw))
    if img.mode not in ("RGB", "L"):
        bg = Image.new("RGB", img.size, (255, 255, 255))
        if img.mode == "RGBA":
            bg.paste(img, mask=img.split()[3])
        else:
            bg.paste(img.convert("RGB"))
        img = bg
    elif img.mode == "L":
        img = img.convert("RGB")
    max_dim = 1200
    w, h = img.size
    if max(w, h) > max_dim:
        if w >= h:
            img = img.resize((max_dim, max(1, int(round(h * max_dim / w)))), Image.Resampling.LANCZOS)
        else:
            img = img.resize((max(1, int(round(w * max_dim / h))), max_dim), Image.Resampling.LANCZOS)
    safe_sku = re.sub(r"[^0-9a-zA-Z_-]+", "", sku) or f"sample_{int(time.time())}"
    out_path = PLAN_ASSETS_DIR / f"{safe_sku}.jpg"
    img.save(out_path, format="JPEG", quality=85, optimize=True)
    return f"./plan_assets/{out_path.name}"


def inject_free_sample_into_plan(payload: dict[str, Any]) -> dict[str, Any]:
    """Inject urgent Free Sample clips (default 3) into guaranteed Main Slots without touching OMNI 5."""
    plan = payload.get("plan")
    if not plan or not plan.get("days"):
        if PLAN_FILE.exists():
            plan = json.loads(PLAN_FILE.read_text(encoding="utf-8"))
        else:
            return {"ok": False, "error": "ยังไม่มีตารางรายเดือนให้แทรกครับ กรุณาสร้างตารางในแท็บที่ 1 ก่อนครับ"}

    days = plan.get("days") or []
    if not days:
        return {"ok": False, "error": "ตารางปัจจุบันยังว่างอยู่ กรุณาสร้างตารางรายเดือนในแท็บที่ 1 ก่อนครับ"}

    sku_or_link = str(payload.get("sku_or_link") or payload.get("sku") or "").strip()
    if not sku_or_link:
        return {"ok": False, "error": "กรุณาวางลิงก์สินค้า TikTok หรือรหัสสินค้า 19 หลักก่อนครับ"}

    unwrapped = unwrap_tiktok_link_or_sku(sku_or_link)
    sku = unwrapped.get("sku") or extract_clean_sku(sku_or_link)
    if not sku:
        return {"ok": False, "error": "ไม่พบเลขรหัสสินค้า 19 หลักจากลิงก์ที่วาง กรุณาก๊อปปี้ลิงก์จากหน้าสินค้า TikTok หรือวางรหัส 19 หลักโดยตรงครับ"}

    raw_name = normalize_thai_text(payload.get("name") or unwrapped.get("og_title") or "")
    resolved = resolve_product_metadata({"sku": sku, "name": raw_name, "link": unwrapped.get("resolved_url", "")}, use_remote_api=True)
    final_name = normalize_thai_text(payload.get("name") or resolved.get("name") or unwrapped.get("og_title") or f"สินค้ารีวิวฟรี ({sku})")
    final_image = str(payload.get("image") or resolved.get("image") or unwrapped.get("og_image") or "").strip()

    images_b64 = payload.get("images_base64") or []
    if isinstance(images_b64, list) and images_b64:
        final_image = save_uploaded_sample_image(sku, str(images_b64[0]))

    clips_count = max(1, min(10, int(payload.get("clips_count") or 3)))
    plan_start = str(days[0].get("date") or "")
    plan_end = str(days[-1].get("date") or "")
    start_date = str(payload.get("start_date") or "").strip() or max(dt.date.today().isoformat(), plan_start)
    if start_date < plan_start:
        start_date = plan_start
    deadline_date = str(payload.get("deadline_date") or "").strip()
    if not deadline_date:
        deadline_date = min(plan_end, (dt.date.fromisoformat(start_date) + dt.timedelta(days=7)).isoformat())
    brief_note = normalize_thai_text(payload.get("brief_note") or "สินค้ารีวิวฟรีกลางเดือน — ใช้เลข SKU ร้านที่ส่งตัวอย่างให้ตรงตัว 100%")

    save_channel_catalog_entries([{
        "sku": sku,
        "name": final_name,
        "image": final_image,
        "cat": resolved.get("cat", ""),
        "price": resolved.get("price", 0),
        "comm": resolved.get("comm", 0),
        "link": unwrapped.get("resolved_url") or f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH",
    }])

    # Select candidate day indices cleanly inside [start_date, deadline_date]
    in_window_indices = [
        idx for idx, d in enumerate(days) if start_date <= str(d.get("date") or "") <= deadline_date
    ]
    if len(in_window_indices) < clips_count:
        extra_indices = [idx for idx, d in enumerate(days) if str(d.get("date") or "") > deadline_date]
        in_window_indices.extend(extra_indices[: max(0, clips_count - len(in_window_indices))])
    if len(in_window_indices) < clips_count:
        # Wrap backward within plan if plan ends sooner than requested count
        prev_indices = [idx for idx in range(len(days)) if idx not in in_window_indices]
        in_window_indices.extend(prev_indices[: max(0, clips_count - len(in_window_indices))])

    # Pick evenly spaced indices starting with Day 0 of window so Quest 1 finishes fast
    chosen_indices: list[int] = []
    if len(in_window_indices) >= clips_count:
        if clips_count == 1:
            chosen_indices = [in_window_indices[0]]
        else:
            step = (len(in_window_indices) - 1) / float(clips_count - 1)
            for c_i in range(clips_count):
                pick_pos = min(len(in_window_indices) - 1, int(round(c_i * step)))
                cand_d = in_window_indices[pick_pos]
                while cand_d in chosen_indices and pick_pos + 1 < len(in_window_indices):
                    pick_pos += 1
                    cand_d = in_window_indices[pick_pos]
                if cand_d not in chosen_indices:
                    chosen_indices.append(cand_d)
    else:
        chosen_indices = in_window_indices[:]

    # Compute current guaranteed main_slot counts before swapping
    plan = recompute_plan_counts(plan)
    cat_by_sku = {str(p.get("sku")): p for p in plan.get("catalog") or [] if p.get("sku")}
    current_main_counts = {sku_k: int(p.get("main_slots") or 0) for sku_k, p in cat_by_sku.items()}
    omni5_skus = {
        str(sl.get("sku") or "")
        for d in days
        for sl in (d.get("slots") or [])
        if sl.get("slot_role") == "reply_fallback"
    }

    existing_f_codes = [p.get("code", "") for p in plan.get("catalog") or [] if str(p.get("code", "")).startswith("F")]
    if sku in cat_by_sku and cat_by_sku[sku].get("code", "").startswith("F"):
        f_code = cat_by_sku[sku]["code"]
    else:
        f_code = f"F{len(existing_f_codes) + 1:02d}"

    try:
        dl_dt = dt.date.fromisoformat(deadline_date)
        dl_short = f"{dl_dt.day}/{dl_dt.month}"
    except ValueError:
        dl_short = deadline_date

    swapped_events: list[dict[str, Any]] = []
    safe_fn = build_safe_video_filename(sku, final_name)
    link_url = f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH"
    fs_prod = {
        "sku": sku,
        "name": final_name,
        "cat": resolved.get("cat", ""),
        "price": float(resolved.get("price") or 0),
        "tier": "free_sample",
    }
    build_product_angles(
        [fs_prod],
        use_ai=bool(payload.get("use_ai") or False),
        explicit_key=str(payload.get("gemini_api_key") or ""),
    )

    # Style plan for the 3 Free Sample clips:
    # Clip 1 -> Style 1 Real-Use Demo (UGC), Clip 2 -> Style 2 Problem->Solution (UGC), Clip 3 -> Style 13 Warehouse Hard-Sell (if OMNI 6/7) or Style 11 Banter (UGC)
    preferred_roles = ["ugc", "ugc", "hardsell"]
    ugc_fallback_styles = [1, 2, 11, 4, 8, 9]
    hardsell_styles = [13, 14, 15, 16]

    for clip_idx, day_idx in enumerate(chosen_indices):
        day = days[day_idx]
        slots = day.get("slots") or []
        want_role = preferred_roles[clip_idx] if clip_idx < len(preferred_roles) else "ugc"

        # Strict rules: NEVER touch OMNI 5 (reply_fallback) and NEVER replace an existing free_sample slot
        eligible_slots = [
            sl for sl in slots
            if sl.get("slot_role") != "reply_fallback" and sl.get("tier") != "free_sample"
        ]
        if not eligible_slots:
            continue

        def swap_score(sl: dict[str, Any]) -> tuple[int, int, int]:
            cur_sku = str(sl.get("sku") or "")
            cur_cnt = current_main_counts.get(cur_sku, 1)
            # Protect OMNI 5 invariant: don't push an OMNI 5 item below 2 main_slots if avoidable
            min_floor = 2 if cur_sku in omni5_skus else 1
            safe_surplus = 1 if (cur_cnt - 1) >= min_floor else 0
            role_match = 1 if sl.get("slot_role") == want_role else 0
            return (safe_surplus, cur_cnt, role_match)

        eligible_slots.sort(key=swap_score, reverse=True)
        target_slot = eligible_slots[0]
        replaced_sku = str(target_slot.get("sku") or "")
        replaced_name = str(target_slot.get("name") or "")
        replaced_code = str(target_slot.get("code") or "")

        if replaced_sku in current_main_counts:
            current_main_counts[replaced_sku] = max(0, current_main_counts[replaced_sku] - 1)

        if target_slot.get("slot_role") == "hardsell":
            st_num = hardsell_styles[clip_idx % len(hardsell_styles)]
        else:
            st_num = ugc_fallback_styles[clip_idx % len(ugc_fallback_styles)]
        st_obj = STYLE_BY_NUM[st_num]

        urgent_txt = (
            f"🚨 ตัวอย่างฟรีด่วน คลิป {clip_idx + 1}/{clips_count} (ส่งเควสต์ก่อน {dl_short} — ห้ามข้าม)"
            if clip_idx == 0
            else f"🎁 ตัวอย่างฟรี คลิป {clip_idx + 1}/{clips_count} (ส่งก่อน {dl_short})"
        )

        target_slot.update({
            "code": f_code,
            "sku": sku,
            "name": final_name,
            "image": final_image,
            "cat": resolved.get("cat", ""),
            "price": float(resolved.get("price") or 0),
            "comm": float(resolved.get("comm") or 0),
            "tier": "free_sample",
            "starred": True,
            "style_num": st_obj["num"],
            "style_name": st_obj["short"],
            "style_chip_match": st_obj["chip_match"],
            "fallback_note": "",
            "urgent_badge": urgent_txt,
            "deadline": deadline_date,
            "brief_note": brief_note,
            "link": link_url,
            "safe_filename": build_safe_video_filename(sku, final_name, ANGLE_CODES[clip_idx % len(ANGLE_CODES)]),
            "reason": f"แทรกคิวตัวอย่างฟรีแทน {replaced_code}",
        })
        target_slot.update(angle_slot_fields(fs_prod, ANGLE_CODES[clip_idx % len(ANGLE_CODES)]))
        swapped_events.append({
            "clip_index": clip_idx + 1,
            "date": day.get("date"),
            "date_thai": day.get("date_thai"),
            "slot_label": target_slot.get("slot_label"),
            "style_name": st_obj["short"],
            "replaced_code": replaced_code,
            "replaced_name": replaced_name[:60],
        })

    # Ensure Free Sample exists cleanly in catalog
    cat_list = plan.get("catalog") or []
    existing_entry = next((c for c in cat_list if str(c.get("sku")) == sku), None)
    entry_payload = {
        "code": f_code,
        "sku": sku,
        "name": final_name,
        "image": final_image,
        "cat": resolved.get("cat", ""),
        "price": float(resolved.get("price") or 0),
        "comm": float(resolved.get("comm") or 0),
        "gmv": 0.0,
        "orders": 0,
        "v_rev": 0.0,
        "growth": 0.0,
        "tier": "free_sample",
        "starred": True,
        "deadline": deadline_date,
        "brief_note": brief_note,
        "reason": f"สินค้ารีวิวฟรีกลางเดือน (แทรก {len(swapped_events)} คลิป ส่งก่อน {dl_short})",
        "link": link_url,
        "safe_filename": safe_fn,
        "marketing": fs_prod.get("marketing") or {},
    }
    if existing_entry:
        existing_entry.update(entry_payload)
    else:
        cat_list.insert(0, entry_payload)
    plan["catalog"] = cat_list
    plan["updated_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    plan = recompute_plan_counts(plan)
    plan["ok"] = True
    plan["swapped_events"] = swapped_events
    return plan


def _pick_style_simple(role: str, preferred: list[int], used_today: set[int]) -> dict[str, Any]:
    """Pick a style for a single re-assigned slot: angle preference first, avoid styles already used that day."""
    cands = list(HARDSELL_STYLE_NUMS if role == "hardsell" else UGC_STYLE_NUMS)
    order = [p for p in preferred if p in cands] + [c for c in cands if c not in preferred]
    for c in order:
        if c not in used_today:
            return STYLE_BY_NUM[c]
    return STYLE_BY_NUM[order[0]]


def _apply_product_to_slot(sl: dict[str, Any], prod: dict[str, Any], angle_code: str, style: dict[str, Any], reason: str = "") -> None:
    """Overwrite a schedule slot with a product + angle + style (keeps slot_num/slot_role/fallback_note)."""
    sku = extract_clean_sku(prod.get("sku"))
    name = str(prod.get("name") or "")
    sl.update({
        "code": prod.get("code") or sl.get("code"),
        "sku": sku,
        "name": name,
        "image": str(prod.get("image") or ""),
        "cat": str(prod.get("cat") or ""),
        "price": float(prod.get("price") or 0),
        "comm": float(prod.get("comm") or 0),
        "tier": prod.get("tier") or "winner",
        "starred": True,
        "style_num": style["num"],
        "style_name": style["short"],
        "style_chip_match": style["chip_match"],
        "urgent_badge": "",
        "deadline": "",
        "brief_note": "",
        "link": prod.get("link") or (f"https://shop.tiktok.com/view/product/{sku}?region=TH&locale=th-TH" if sku else ""),
        "safe_filename": build_safe_video_filename(sku, name, angle_code),
        "reason": reason or sl.get("reason", ""),
    })
    sl.update(angle_slot_fields(prod, angle_code))


def _reangle_slot(sl: dict[str, Any], prod: dict[str, Any], angle_code: str, used_today: set[int]) -> None:
    """Switch an existing slot (same product) to another angle and an angle-matching style."""
    role = sl.get("slot_role") or "ugc"
    style = _pick_style_simple(role, preferred_styles_for(prod, angle_code, role), used_today)
    sl.update({
        "style_num": style["num"],
        "style_name": style["short"],
        "style_chip_match": style["chip_match"],
        "safe_filename": build_safe_video_filename(sl.get("sku", ""), sl.get("name", ""), angle_code),
    })
    sl.update(angle_slot_fields(prod, angle_code))


def _repair_future_omni5(plan: dict[str, Any], after_date: str) -> int:
    """Re-assign future OMNI 5 slots whose product no longer has >= 2 main slots (keeps the OMNI 5 invariant)."""
    repaired = 0
    catalog = plan.get("catalog") or []
    for v in (plan.get("summary") or {}).get("omni5_violations") or []:
        if str(v.get("date") or "") <= after_date:
            continue
        day = next((d for d in plan.get("days") or [] if d.get("date") == v.get("date")), None)
        if not day:
            continue
        sl = next((s for s in day.get("slots") or [] if s.get("slot_num") == v.get("slot_num")), None)
        if not sl:
            continue
        main_today = {s.get("sku") for s in day.get("slots") or [] if s.get("slot_role") != "reply_fallback"}
        cands = [
            p for p in catalog
            if int(p.get("main_slots") or 0) >= 2
            and p.get("tier") in ("winner", "new_star")
            and p.get("review_status") != "dropped_weekly"
            and p.get("sku")
        ]
        cands.sort(key=lambda p: (1 if p.get("sku") in main_today else 0, -int(p.get("main_slots") or 0)))
        if not cands:
            continue
        best = cands[0]
        if not best.get("marketing"):
            build_product_angles([best], use_ai=False)
        used_today = {s.get("style_num") for s in day.get("slots") or [] if s is not sl}
        code = angle_sequence_code(best.get("marketing"), str(best.get("best_angle") or ""), int(best.get("fallback_slots") or 0))
        _apply_product_to_slot(sl, best, code, _pick_style_simple("ugc", preferred_styles_for(best, code, "ugc"), used_today), reason="สำรอง OMNI 5 หลังรีวิวราย 7 วัน")
        repaired += 1
    return repaired


def weekly_review_plan(payload: dict[str, Any]) -> dict[str, Any]:
    """Sunday 5-minute review: re-plan ONLY future slots from the latest sales file + bot post log.

    Rules (agreed with the owner):
      - >= `min_posts_to_drop` (4) bot posts older than `min_age_days` (3) AND 0 orders from new clips -> drop,
        replace its future slots with a fresh visual-proof-gated new star.
      - orders attributed to an angle -> lock that SKU's future slots onto its best angle (4 of 6 slots).
      - 1-2 posts -> keep testing. Free samples are never dropped.
    """
    plan = payload.get("plan")
    if not plan or not plan.get("days"):
        if PLAN_FILE.exists():
            plan = json.loads(PLAN_FILE.read_text(encoding="utf-8"))
        else:
            return {"ok": False, "error": "ยังไม่มีตารางรายเดือนให้รีวิว กรุณาสร้างตารางในแท็บที่ 1 ก่อนครับ"}
    plan = json.loads(json.dumps(plan, ensure_ascii=False))
    days = plan.get("days") or []
    if not days:
        return {"ok": False, "error": "ตารางปัจจุบันยังว่างอยู่ กรุณาสร้างตารางรายเดือนก่อนครับ"}
    files = payload.get("files") or []
    if not files:
        return {"ok": False, "error": "กรุณาแนบไฟล์ยอดขายล่าสุด (affiliate_orders_*.xlsx) ก่อนกดรีวิวครับ"}

    review_date = str(payload.get("review_date") or "").strip() or dt.date.today().isoformat()
    min_posts = max(2, int(payload.get("min_posts_to_drop") or 4))
    min_age = max(1, int(payload.get("min_age_days") or 3))
    use_ai = bool(payload.get("use_ai") or False)
    explicit_key = str(payload.get("gemini_api_key") or "")
    hist_path = Path(str(payload["_hist_path"])) if payload.get("_hist_path") else None
    plan_start = str(plan.get("start_date") or days[0].get("date") or "")

    parsed = parse_sales_files(files)
    items = parsed["products"]
    attribute_angle_sales(items, hist_path=hist_path, min_date=plan_start, max_date=review_date)

    sales_by_sku: dict[str, dict[str, Any]] = {}
    for it in items:
        sku = extract_clean_sku(it.get("sku"))
        cids = it.pop("_cid_pub_comm", None) or {}
        if not sku:
            continue
        if cids:
            new_cids = [st for st in cids.values() if not st.get("pub_date") or st["pub_date"] >= plan_start]
            new_orders = sum(int(st.get("orders") or 0) for st in new_cids)
            new_comm = sum(float(st.get("comm") or 0) for st in new_cids)
        else:
            new_orders = int(it.get("orders") or 0)
            new_comm = float(it.get("comm") or 0)
        sales_by_sku[sku] = {
            "orders_total": int(it.get("orders") or 0),
            "comm_total": round(float(it.get("comm") or 0), 2),
            "new_orders": new_orders,
            "new_comm": round(new_comm, 2),
            "angle_stats": it.get("angle_stats") or {},
            "best_angle": str(it.get("best_angle") or ""),
        }

    matured_cutoff = (dt.date.fromisoformat(review_date) - dt.timedelta(days=min_age)).isoformat()
    posted_by_sku: dict[str, int] = {}
    matured_by_sku: dict[str, int] = {}
    for r in load_bot_post_rows(hist_path):
        if not r["sku"] or not r["date"] or r["date"] < plan_start or r["date"] > review_date:
            continue
        posted_by_sku[r["sku"]] = posted_by_sku.get(r["sku"], 0) + 1
        if r["date"] <= matured_cutoff:
            matured_by_sku[r["sku"]] = matured_by_sku.get(r["sku"], 0) + 1

    future_idx = [i for i, d in enumerate(days) if str(d.get("date") or "") > review_date]
    future_slots_by_sku: dict[str, list[tuple[int, dict[str, Any]]]] = {}
    for i in future_idx:
        for sl in days[i].get("slots") or []:
            future_slots_by_sku.setdefault(str(sl.get("sku") or ""), []).append((i, sl))

    catalog = plan.get("catalog") or []
    cat_by_sku = {extract_clean_sku(p.get("sku")): p for p in catalog if extract_clean_sku(p.get("sku"))}
    build_product_angles(list(cat_by_sku.values()), use_ai=use_ai, explicit_key=explicit_key)

    decisions: dict[str, list[dict[str, Any]]] = {"drop": [], "scale": [], "keep_testing": [], "watching": [], "protected": []}
    for sku, prod in cat_by_sku.items():
        if prod.get("review_status") == "dropped_weekly":
            continue
        s = sales_by_sku.get(sku, {})
        posted = posted_by_sku.get(sku, 0)
        matured = matured_by_sku.get(sku, 0)
        new_orders = int(s.get("new_orders") or 0)
        row = {
            "sku": sku,
            "code": prod.get("code", ""),
            "name": str(prod.get("name") or "")[:80],
            "tier": prod.get("tier", ""),
            "posted": posted,
            "matured": matured,
            "new_orders": new_orders,
            "new_comm": float(s.get("new_comm") or 0),
            "orders_total": int(s.get("orders_total") or 0),
            "future_slots": len(future_slots_by_sku.get(sku, [])),
            "best_angle": str(s.get("best_angle") or ""),
            "angle_stats": s.get("angle_stats") or {},
        }
        if s.get("angle_stats"):
            prod["angle_stats"] = s["angle_stats"]
        if row["best_angle"]:
            prod["best_angle"] = row["best_angle"]
        if prod.get("tier") == "free_sample":
            row["reason"] = "🎁 สินค้ารีวิวฟรี (ข้อตกลงกับร้าน) — ไม่ตัดออกแม้ยอดยังไม่มา"
            decisions["protected"].append(row)
        elif matured >= min_posts and new_orders == 0:
            row["reason"] = f"❌ บอทโพสต์คลิปใหม่แล้ว {posted} คลิป (เกิน {min_age} วันแล้ว {matured} คลิป) แต่คลิปใหม่ขายได้ 0 ชิ้น"
            decisions["drop"].append(row)
        elif new_orders > 0:
            angle_txt = f" · มุมที่ขายออก: {row['best_angle']}" if row["best_angle"] else " · ยังแยกมุมไม่ได้ (คลิปเก่าไม่มีป้ายมุม)"
            row["reason"] = f"✅ คลิปใหม่ขายออก {new_orders} ชิ้น (คอม ฿{row['new_comm']:,.0f}){angle_txt}"
            decisions["scale"].append(row)
        elif posted <= 2:
            row["reason"] = f"⏳ เพิ่งโพสต์ {posted} คลิป — ยังเร็วเกินไปที่จะตัดสิน ปล่อยเทสต์ต่อ"
            decisions["keep_testing"].append(row)
        else:
            row["reason"] = f"👀 โพสต์ {posted} คลิป (ครบ {min_age} วันแล้ว {matured} คลิป) ยังไม่มียอด — เฝ้าดูอีกสัปดาห์"
            decisions["watching"].append(row)

    # Replace future slots of dropped products with fresh, visual-proof-gated new stars
    drop_with_future = [r for r in decisions["drop"] if r["future_slots"] > 0]
    gate_notes: list[str] = []
    replacements: list[dict[str, Any]] = []
    if drop_with_future:
        winners_like = [p for p in catalog if p.get("tier") in ("winner", "new_star") and p.get("review_status") != "dropped_weekly"]
        prev_dropped = list(plan.get("dropped_products") or []) + [cat_by_sku[r["sku"]] for r in decisions["drop"]]
        picked, backups = select_new_stars_from_pool(
            winners_like,
            prev_dropped,
            desired_count=len(drop_with_future),
            strategy=str(payload.get("new_star_strategy") or "balanced"),
        )
        plan_skus = set(cat_by_sku.keys())
        picked = [p for p in picked if p.get("sku") not in plan_skus]
        backups = [b for b in backups if b.get("sku") not in plan_skus]
        picked, backups, gate_notes = gate_new_stars_by_visual_proof(picked, backups, len(drop_with_future), use_ai=use_ai, explicit_key=explicit_key)
        replacements = picked

    n_nums = [int(str(p.get("code"))[1:]) for p in catalog if str(p.get("code") or "").startswith("N") and str(p.get("code"))[1:].isdigit()]
    next_n = (max(n_nums) if n_nums else 0) + 1
    scale_sorted = sorted(
        [cat_by_sku[r["sku"]] for r in decisions["scale"] if r["sku"] in cat_by_sku],
        key=lambda p: -float((sales_by_sku.get(extract_clean_sku(p.get("sku")), {}) or {}).get("new_comm") or 0),
    )
    replaced_slots = 0
    for i, r in enumerate(drop_with_future):
        old_prod = cat_by_sku[r["sku"]]
        if i < len(replacements):
            target = dict(replacements[i])
            target["code"] = f"N{next_n:02d}"
            next_n += 1
            target["tier"] = "new_star"
            target["starred"] = True
            target["reason"] = f"🔁 เสียบแทน {old_prod.get('code', '')} หลังรีวิวยอดขาย {review_date} · {target.get('reason') or ''}".strip(" ·")
            catalog.append({
                "code": target["code"],
                "sku": extract_clean_sku(target.get("sku")),
                "name": str(target.get("name") or ""),
                "image": str(target.get("image") or ""),
                "cat": str(target.get("cat") or ""),
                "price": float(target.get("price") or 0),
                "comm": float(target.get("comm") or 0),
                "gmv": 0.0,
                "orders": 0,
                "v_rev": float(target.get("v_rev") or 0),
                "growth": float(target.get("growth") or 0),
                "tier": "new_star",
                "starred": True,
                "reason": target["reason"],
                "link": target.get("link") or f"https://shop.tiktok.com/view/product/{extract_clean_sku(target.get('sku'))}?region=TH&locale=th-TH",
                "safe_filename": build_safe_video_filename(target.get("sku", ""), str(target.get("name") or "")),
                "marketing": target.get("marketing") or {},
                "best_angle": "",
                "angle_stats": {},
            })
            r["replaced_by"] = {"sku": extract_clean_sku(target.get("sku")), "code": target["code"], "name": str(target.get("name") or "")[:80], "fallback": False}
        elif scale_sorted:
            target = scale_sorted[i % len(scale_sorted)]
            r["replaced_by"] = {"sku": extract_clean_sku(target.get("sku")), "code": target.get("code", ""), "name": str(target.get("name") or "")[:80], "fallback": True}
        else:
            r["replaced_by"] = None
            continue
        occ = 0
        for d_i, sl in future_slots_by_sku.get(r["sku"], []):
            role = sl.get("slot_role") or "ugc"
            code = angle_sequence_code(target.get("marketing"), str(target.get("best_angle") or ""), occ)
            occ += 1
            used_today = {s2.get("style_num") for s2 in days[d_i].get("slots") or [] if s2 is not sl}
            style = _pick_style_simple(role, preferred_styles_for(target, code, role), used_today)
            _apply_product_to_slot(sl, target, code, style, reason=f"แทน {old_prod.get('code', '')} (รีวิว {review_date})")
            replaced_slots += 1
        old_prod["review_status"] = "dropped_weekly"
        old_prod["reason"] = f"{r['reason']} (ตัดออกจากคิวหลัง {review_date})"
        plan.setdefault("dropped_products", []).append({
            "code": old_prod.get("code", ""),
            "sku": r["sku"],
            "name": old_prod.get("name", ""),
            "tier": "dropped",
            "reason": old_prod["reason"],
        })

    # Lock proven angles onto future slots of products that are selling
    for r in decisions["drop"]:
        if r["future_slots"] == 0 and r["sku"] in cat_by_sku:
            cat_by_sku[r["sku"]]["review_status"] = "dropped_weekly"
            r["replaced_by"] = None
    relocked = 0
    for r in decisions["scale"]:
        best = r.get("best_angle") or ""
        prod = cat_by_sku.get(r["sku"])
        if not best or not prod:
            continue
        occ = 0
        for d_i, sl in future_slots_by_sku.get(r["sku"], []):
            if extract_clean_sku(sl.get("sku")) != r["sku"]:
                continue
            code = angle_sequence_code(prod.get("marketing"), best, occ)
            occ += 1
            if sl.get("angle_code") != code:
                used_today = {s2.get("style_num") for s2 in days[d_i].get("slots") or [] if s2 is not sl}
                _reangle_slot(sl, prod, code, used_today)
                relocked += 1

    # Older plans (made before angles existed): give every remaining future slot an angle
    backfilled = 0
    occ_by_sku: dict[str, int] = {}
    for i in future_idx:
        for sl in days[i].get("slots") or []:
            sku = extract_clean_sku(sl.get("sku"))
            if sl.get("angle_code") or not sku:
                continue
            prod = cat_by_sku.get(sku) or {"sku": sku, "name": sl.get("name", ""), "price": sl.get("price", 0), "cat": sl.get("cat", "")}
            if not prod.get("marketing"):
                build_product_angles([prod], use_ai=False)
            occ = occ_by_sku.get(sku, 0)
            occ_by_sku[sku] = occ + 1
            code = angle_sequence_code(prod.get("marketing"), str(prod.get("best_angle") or ""), occ)
            used_today = {s2.get("style_num") for s2 in days[i].get("slots") or [] if s2 is not sl}
            _reangle_slot(sl, prod, code, used_today)
            backfilled += 1

    plan["catalog"] = catalog
    plan = recompute_plan_counts(plan)
    omni5_repaired = _repair_future_omni5(plan, review_date)
    if omni5_repaired:
        plan = recompute_plan_counts(plan)

    report = {
        "review_date": review_date,
        "plan_window": f"{plan_start} → {plan.get('end_date', '')}",
        "sales_window": f"{parsed.get('min_order_date', '')} → {parsed.get('max_order_date', '')}",
        "rules": {"min_posts_to_drop": min_posts, "min_age_days": min_age},
        "dropped": decisions["drop"],
        "scale": decisions["scale"],
        "keep_testing": decisions["keep_testing"],
        "watching": decisions["watching"],
        "protected": decisions["protected"],
        "replaced_slots": replaced_slots,
        "relocked_slots": relocked,
        "backfilled_slots": backfilled,
        "omni5_repaired": omni5_repaired,
        "future_days": len(future_idx),
        "marketing_notes": gate_notes,
        "warnings": parsed.get("warnings", []),
    }
    reviews = list(plan.get("weekly_reviews") or [])
    reviews.append({
        "review_date": review_date,
        "dropped": [r["code"] for r in decisions["drop"]],
        "scale": [r["code"] for r in decisions["scale"]],
        "replaced_slots": replaced_slots,
        "relocked_slots": relocked,
    })
    plan["weekly_reviews"] = reviews[-10:]
    plan["updated_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    plan["ok"] = True
    return {"ok": True, "plan": plan, "report": report}


def export_plan_excel_bytes(plan: dict[str, Any]) -> bytes:
    """Export the monthly plan to a styled .xlsx workbook preserving 19-digit SKUs as strings."""
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "ตารางรายวัน_7Omni"

    clips_per_day = max(1, int(plan.get("clips_per_day") or 7))
    headers1 = ["วันที่", "วัน/เดือน/ปี (ISO)"] + [f"OMNI {i}" for i in range(1, clips_per_day + 1)]
    ws1.append(headers1)

    header_fill = PatternFill("solid", fgColor="1E293B")
    header_font = Font(name="Kanit", size=11, bold=True, color="FFFFFF")
    omni5_fill = PatternFill("solid", fgColor="FEF3C7")
    urgent_fill = PatternFill("solid", fgColor="FEE2E2")
    hardsell_fill = PatternFill("solid", fgColor="ECFEFF")
    normal_fill = PatternFill("solid", fgColor="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    for col_idx in range(1, len(headers1) + 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_i, day in enumerate(plan.get("days") or [], start=2):
        ws1.cell(row=r_i, column=1, value=str(day.get("date_thai") or "")).alignment = Alignment(horizontal="center", vertical="top")
        ws1.cell(row=r_i, column=2, value=str(day.get("date") or "")).alignment = Alignment(horizontal="center", vertical="top")
        for s_idx, sl in enumerate(day.get("slots") or [], start=3):
            sku = str(sl.get("sku") or "")
            lines = [
                f"{sl.get('code', '')} | {sl.get('name', '')}",
                f"SKU: {sku}" if sku else "⚠️ ยังไม่มี SKU",
                f"{sl.get('style_name', '')}",
            ]
            if sl.get("angle_code"):
                lines.append(f"🎯 {sl.get('angle_code')} {sl.get('angle_label', '')}")
                if sl.get("angle_hook"):
                    lines.append(f"ฮุค: {sl.get('angle_hook')}")
            if sl.get("slot_role") == "reply_fallback":
                lines.append("💬 ตอบคอมเมนต์แทนได้")
            elif sl.get("urgent_badge"):
                lines.append(sl["urgent_badge"])
            cell = ws1.cell(row=r_i, column=s_idx, value="\n".join(lines))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin_border
            if sl.get("tier") == "free_sample":
                cell.fill = urgent_fill
            elif sl.get("slot_role") == "reply_fallback":
                cell.fill = omni5_fill
            elif sl.get("slot_role") == "hardsell":
                cell.fill = hardsell_fill
            else:
                cell.fill = normal_fill
        ws1.row_dimensions[r_i].height = 104

    ws1.column_dimensions["A"].width = 14
    ws1.column_dimensions["B"].width = 14
    for col_idx in range(3, len(headers1) + 1):
        ws1.column_dimensions[get_column_letter(col_idx)].width = 34

    # Sheet 2: Product Summary & 19-digit SKUs (stored explicitly as text strings)
    ws2 = wb.create_sheet("คลังสินค้า_SKU19หลัก")
    headers2 = ["รหัส", "กลุ่ม", "SKU 19 หลัก", "ชื่อสินค้า", "คิวหลัก", "สำรอง OMNI5", "รวมทั้งเดือน", "ชื่อไฟล์อัปลง Drive (.mp4)", "ลิงก์สินค้า TikTok", "หมายเหตุ", "มุมที่ขายดี", "ภาพพิสูจน์ (1-5)", "ข้อกังวลหลักของลูกค้า"]
    ws2.append(headers2)
    for col_idx in range(1, len(headers2) + 1):
        cell = ws2.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    tier_labels = {
        "winner": "⭐ ขายดีประจำช่อง",
        "new_star": "🌟 สินค้าใหม่น่าลอง",
        "free_sample": "🎁 ตัวอย่างฟรีกลางเดือน",
        "dropped": "❌ ปัดตกเดือนนี้",
    }
    row_cursor = 2
    for prod in (plan.get("catalog") or []) + (plan.get("dropped_products") or []):
        sku_str = str(prod.get("sku") or "")
        tier = str(prod.get("tier") or "winner")
        ws2.cell(row=row_cursor, column=1, value=str(prod.get("code") or "-"))
        ws2.cell(row=row_cursor, column=2, value=tier_labels.get(tier, tier))
        sku_cell = ws2.cell(row=row_cursor, column=3, value=sku_str)
        sku_cell.number_format = "@"  # Preserve 19 digits strictly as text
        ws2.cell(row=row_cursor, column=4, value=str(prod.get("name") or ""))
        ws2.cell(row=row_cursor, column=5, value=int(prod.get("main_slots") or 0))
        ws2.cell(row=row_cursor, column=6, value=int(prod.get("fallback_slots") or 0))
        ws2.cell(row=row_cursor, column=7, value=int(prod.get("total_slots") or 0))
        ws2.cell(row=row_cursor, column=8, value=build_safe_video_filename(sku_str, str(prod.get("name") or "")))
        ws2.cell(row=row_cursor, column=9, value=str(prod.get("link") or ""))
        ws2.cell(row=row_cursor, column=10, value=str(prod.get("reason") or ""))
        mk = prod.get("marketing") or {}
        ws2.cell(row=row_cursor, column=11, value=str(prod.get("best_angle") or "-"))
        ws2.cell(row=row_cursor, column=12, value=int(mk.get("visual_proof_score") or 0) or "-")
        ws2.cell(row=row_cursor, column=13, value=str(mk.get("core_objection") or ""))
        row_cursor += 1

    widths2 = [8, 22, 22, 46, 10, 13, 13, 52, 42, 36, 12, 14, 40]
    for idx, w in enumerate(widths2, start=1):
        ws2.column_dimensions[get_column_letter(idx)].width = w

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def publish_plan_to_github(plan: dict[str, Any]) -> dict[str, Any]:
    """Write monthly_plan.json cleanly and push plan + plan_assets to GitHub Pages."""
    plan = recompute_plan_counts(plan)
    missing = int(plan.get("summary", {}).get("missing_sku_in_slots") or 0)
    if missing > 0 and not plan.get("_allow_missing_sku"):
        return {
            "ok": False,
            "error": f"ยังมีช่องตารางที่ขาดรหัสสินค้า 19 หลักอยู่ {missing} ช่อง กรุณากรอกเลข 19 หลักหรือวางลิงก์สินค้าให้ครบก่อนอัปเดตขึ้นเว็บครับ",
        }

    plan_clean = {k: v for k, v in plan.items() if not k.startswith("_") and k != "ok"}
    plan_clean["updated_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    PLAN_FILE.write_text(json.dumps(plan_clean, ensure_ascii=False, indent=2), encoding="utf-8")

    # Stage ONLY monthly_plan.json and plan_assets (leave any unrelated working tree changes untouched unless staged)
    add_res = subprocess.run(["git", "-C", str(BASE_DIR), "add", "monthly_plan.json", "plan_assets"], capture_output=True, text=True)
    if add_res.returncode != 0:
        return {"ok": False, "error": f"git add ไม่สำเร็จ: {add_res.stderr}"}

    diff_check = subprocess.run(["git", "-C", str(BASE_DIR), "diff", "--cached", "--quiet", "--", "monthly_plan.json", "plan_assets"])
    if diff_check.returncode != 0:
        commit_msg = f"chore: update monthly plan ({plan_clean.get('updated_at', '')})"
        cm_res = subprocess.run(
            [
                "git",
                "-c",
                "user.name=Chatchakris",
                "-c",
                "user.email=chatchakrissu@gmail.com",
                "-C",
                str(BASE_DIR),
                "commit",
                "--only",
                "-m",
                commit_msg,
                "--",
                "monthly_plan.json",
                "plan_assets",
            ],
            capture_output=True,
            text=True,
        )
        if cm_res.returncode != 0:
            return {"ok": False, "error": f"git commit ไม่สำเร็จ: {cm_res.stderr or cm_res.stdout}"}

    push_res = subprocess.run(["git", "-C", str(BASE_DIR), "push"], capture_output=True, text=True)
    if push_res.returncode != 0:
        return {
            "ok": False,
            "error": f"บันทึกไฟล์ monthly_plan.json ในเครื่องแล้ว แต่ git push ขึ้น GitHub ไม่ผ่าน: {push_res.stderr or push_res.stdout}",
        }

    return {
        "ok": True,
        "updated_at": plan_clean["updated_at"],
        "live_url": "https://chatchakrissu-hash.github.io/tiktok-ugc-studio/",
        "summary": plan_clean.get("summary", {}),
    }


class PlannerHTTPHandler(http.server.BaseHTTPRequestHandler):
    """Serve static assets from tiktok-ugc-studio/ and JSON API endpoints for planner.html."""

    server_version = "TikTokPlannerHTTP/1.0"

    def log_message(self, format: str, *args: Any) -> None:
        # Keep console logs concise without noisy static asset spam
        if args and isinstance(args[0], str) and "/api/" in args[0]:
            sys.stderr.write(f"[Planner] {args[0]}\n")

    def _send_cors_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _send_json(self, obj: Any, status: int = 200) -> None:
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self._send_cors_headers()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _read_json_body(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        return json.loads(raw.decode("utf-8"))

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self._send_cors_headers()
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/api/ping":
            self._send_json({"ok": True, "service": "tiktok-planner", "port": self.server.server_address[1]})
            return
        if path == "/api/styles":
            self._send_json({"ok": True, "styles": STYLES_16})
            return
        if path == "/api/plan":
            plan = json.loads(PLAN_FILE.read_text(encoding="utf-8")) if PLAN_FILE.exists() else {}
            self._send_json({"ok": True, "plan": plan, "channel_catalog_size": len(load_channel_catalog())})
            return
        if path == "/api/pool-suggestions":
            winners = []
            if PLAN_FILE.exists():
                cur = json.loads(PLAN_FILE.read_text(encoding="utf-8"))
                winners = [c for c in (cur.get("catalog") or []) if c.get("tier") == "winner"]
            stars, backups = select_new_stars_from_pool(winners, [], desired_count=12)
            self._send_json({"ok": True, "new_stars": stars, "backups": backups})
            return

        # Serve local files safely within BASE_DIR
        rel = urllib.parse.unquote(path.lstrip("/")) or "planner.html"
        target = (BASE_DIR / rel).resolve()
        if not str(target).startswith(str(BASE_DIR.resolve())) or not target.is_file():
            self.send_error(404, "Not Found")
            return
        mime, _ = mimetypes.guess_type(str(target))
        content = target.read_bytes()
        self.send_response(200)
        self._send_cors_headers()
        self.send_header("Content-Type", f"{mime or 'application/octet-stream'}; charset=utf-8" if str(mime or "").startswith("text/") or target.suffix in {".json", ".html", ".js", ".css"} else (mime or "application/octet-stream"))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        try:
            body = self._read_json_body()
            if path == "/api/analyze-sales":
                self._send_json(analyze_sales_payload(body))
                return
            if path == "/api/resolve-item":
                raw = str(body.get("sku_or_link") or body.get("query") or "").strip()
                unwrapped = unwrap_tiktok_link_or_sku(raw)
                item = {
                    "sku": unwrapped.get("sku", ""),
                    "name": normalize_thai_text(body.get("name") or unwrapped.get("og_title") or ("" if raw.startswith("http") else raw)),
                    "image": unwrapped.get("og_image", ""),
                    "link": unwrapped.get("resolved_url", ""),
                }
                resolved = resolve_product_metadata(item, use_remote_api=True)
                if resolved.get("sku"):
                    save_channel_catalog_entries([resolved])
                self._send_json({"ok": True, "item": resolved})
                return
            if path == "/api/generate-schedule":
                self._send_json(generate_monthly_schedule(body))
                return
            if path == "/api/recompute-plan":
                self._send_json({"ok": True, **recompute_plan_counts(body)})
                return
            if path == "/api/inject-free-sample":
                self._send_json(inject_free_sample_into_plan(body))
                return
            if path == "/api/weekly-review":
                self._send_json(weekly_review_plan(body))
                return
            if path == "/api/product-angles":
                prods = [dict(p) for p in (body.get("products") or []) if isinstance(p, dict)]
                build_product_angles(
                    prods,
                    use_ai=bool(body.get("use_ai") if body.get("use_ai") is not None else True),
                    explicit_key=str(body.get("gemini_api_key") or ""),
                    force=bool(body.get("force") or False),
                )
                self._send_json({"ok": True, "products": prods})
                return
            if path == "/api/save-draft":
                plan = recompute_plan_counts(body.get("plan") or body)
                plan_clean = {k: v for k, v in plan.items() if not k.startswith("_") and k != "ok"}
                plan_clean["updated_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                PLAN_FILE.write_text(json.dumps(plan_clean, ensure_ascii=False, indent=2), encoding="utf-8")
                self._send_json({"ok": True, "updated_at": plan_clean["updated_at"], "summary": plan_clean["summary"]})
                return
            if path == "/api/publish":
                plan = body.get("plan") or body
                self._send_json(publish_plan_to_github(plan))
                return
            if path == "/api/export-excel":
                plan = recompute_plan_counts(body.get("plan") or body)
                xlsx_bytes = export_plan_excel_bytes(plan)
                self.send_response(200)
                self._send_cors_headers()
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", 'attachment; filename="TikTok_Monthly_Schedule.xlsx"')
                self.send_header("Content-Length", str(len(xlsx_bytes)))
                self.end_headers()
                self.wfile.write(xlsx_bytes)
                return
            self._send_json({"ok": False, "error": f"Unknown endpoint: {path}"}, status=404)
        except Exception as exc:
            self._send_json({"ok": False, "error": str(exc)}, status=500)


class ReusableThreadingServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def create_server(host: str = "127.0.0.1", port: int = DEFAULT_PORT) -> ReusableThreadingServer:
    """Create configured HTTP server instance for CLI or automated verification."""
    ensure_pool_loaded()
    return ReusableThreadingServer((host, port), PlannerHTTPHandler)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="TikTok Monthly Content Planner Server")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    args = parser.parse_args()

    srv = create_server(args.host, args.port)
    print(f"🚀 TikTok Monthly Planner เปิดทำงานที่ http://{args.host}:{args.port}/planner.html")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n👋 ปิดระบบจัดตารางรายเดือนแล้ว")
        srv.shutdown()
