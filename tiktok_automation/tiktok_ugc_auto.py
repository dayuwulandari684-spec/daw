#!/usr/bin/env python3
"""
TikTok UGC Automation
Scrape produk best-selling TikTok Shop → buat konten UGC → auto upload TikTok

Cara pakai:
  pip install -r requirements.txt
  playwright install chromium
  python tiktok_ugc_auto.py
"""

import os
import io
import re
import sys
import gc
import time
import json
import random
import shutil
import subprocess
import textwrap
import urllib.request
import urllib.parse

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)

# =============================================================================
# KONFIGURASI
# =============================================================================

KEYWORD           = "skincare viral"   # kata kunci produk di TikTok Shop
MAX_PRODUCTS      = 2                  # jumlah produk yang diproses per run
SHUFFLE           = True               # acak urutan produk hasil scraping
SEARCH_SCROLL     = 3                  # berapa kali scroll saat scraping

# Kalau scraping sering crash/force quit, set True → pakai data di PRODUK_MANUAL
SKIP_SCRAPING     = True

# Isi produk sendiri di sini (aktif kalau SKIP_SCRAPING = True)
PRODUK_MANUAL = [
    {
        "product_id" : "manual_001",
        "name"       : "Serum Vitamin C Glowing 30ml",
        "price"      : 89_000,
        "rating"     : 4.8,
        "sold_count" : 15_420,
        "image_urls" : [],
        "image_path" : "",   # ← path foto produk lokal, contoh: r"C:\foto\serum.jpg"
        "category"   : "skincare",
        "url"        : "",
        "description": "Cerahkan kulit dalam 7 hari. BPOM certified.",
    },
    {
        "product_id" : "manual_002",
        "name"       : "Sunscreen SPF50 PA++++ 50ml",
        "price"      : 65_000,
        "rating"     : 4.7,
        "sold_count" : 8_930,
        "image_urls" : [],
        "image_path" : "",   # ← path foto produk lokal, contoh: r"C:\foto\sunscreen.jpg"
        "category"   : "skincare",
        "url"        : "",
        "description": "Ringan, tidak lengket, cocok kulit berminyak.",
    },
]

OUTPUT_FOLDER     = "output"           # folder hasil video
FRAMES_FOLDER     = "output/frames"   # folder sementara slide PNG

# Video
VIDEO_W           = 1080
VIDEO_H           = 1920
FPS               = 30
SLIDE_DURATION    = 4                  # detik per slide

# Voiceover
CONTENT_LANG      = "id"              # "id" = Bahasa Indonesia, "en" = English
TTS_TLD           = "co.id"           # domain Google TTS (co.id = aksen Indonesia)

# TikTok upload (via Playwright + cookies)
TIKTOK_COOKIE_FILE = "tiktok_cookies.txt"   # export dari browser pakai ekstensi
AUTO_UPLOAD        = True
HEADLESS_SCRAPE    = True             # True = scraping tanpa tampilkan browser
HEADLESS_UPLOAD    = False            # False = tampilkan browser saat upload


CAPTION_TEMPLATE = (
    "{hook}\n\n"
    "✨ {nama}\n"
    "💰 Rp{harga}\n"
    "⭐ {rating}/5\n\n"
    "🛒 Link ada di keranjang di bawah!\n\n"
    "{hashtags}"
)

BASE_HASHTAGS = [
    "#TikTokShop", "#FYP", "#Viral",
    "#Rekomendasi", "#ProdukViral", "#TikTokMadeMeBuyIt",
]

# =============================================================================
# SETUP FOLDER
# =============================================================================

for _d in [OUTPUT_FOLDER, FRAMES_FOLDER, "output/images"]:
    os.makedirs(_d, exist_ok=True)

# =============================================================================
# CEK DEPENDENCY
# =============================================================================

try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter
    import numpy as np
    PIL_OK = True
except ImportError:
    PIL_OK = False
    print("[WARN] Pillow/numpy tidak ada. Jalankan: pip install Pillow numpy")

try:
    from gtts import gTTS
    GTTS_OK = True
except ImportError:
    GTTS_OK = False
    print("[WARN] gTTS tidak ada. Jalankan: pip install gTTS")

try:
    from playwright.sync_api import sync_playwright
    PW_OK = True
except ImportError:
    PW_OK = False
    print("[WARN] Playwright tidak ada. Jalankan: pip install playwright && playwright install chromium")

# =============================================================================
# UTILITAS UMUM
# =============================================================================

PALETTES = {
    "skincare"  : [(255, 182, 193), (219, 112, 147)],
    "kecantikan": [(255, 160, 122), (220,  20,  60)],
    "fashion"   : [(147, 112, 219), ( 75,   0, 130)],
    "makanan"   : [(255, 165,   0), (255,  69,   0)],
    "gadget"    : [( 30, 144, 255), (  0,   0, 139)],
    "default"   : [(255,  99,  71), (220,  20,  60)],
}


def get_palette(keyword: str) -> list:
    kw = keyword.lower()
    for key in PALETTES:
        if key in kw:
            return PALETTES[key]
    return PALETTES["default"]


def run_ffmpeg(args: list, label: str = "ffmpeg") -> bool:
    """Jalankan ffmpeg. Return True jika sukses."""
    if shutil.which("ffmpeg") is None:
        print("  [ERROR] ffmpeg tidak ditemukan. Download: https://ffmpeg.org/download.html")
        return False
    try:
        subprocess.run(
            ["ffmpeg", "-y"] + args,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
        )
        return True
    except subprocess.CalledProcessError as e:
        err = e.stderr.decode(errors="ignore")[-400:]
        print(f"  [ERROR] ffmpeg ({label}): ...{err}")
        return False


def _get_font(size: int):
    """Cari font bold di sistem (Windows & Linux)."""
    candidates = [
        "C:/Windows/Fonts/ariblk.ttf",
        "C:/Windows/Fonts/arialbd.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
        "/usr/share/fonts/truetype/ubuntu/Ubuntu-B.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


def _wrap_text(text: str, font, max_w: int, draw: "ImageDraw") -> list:
    """Bagi teks menjadi baris yang muat dalam max_w."""
    words, lines, cur = text.split(), [], ""
    for word in words:
        test = f"{cur} {word}".strip()
        bb   = draw.textbbox((0, 0), test, font=font)
        if bb[2] - bb[0] <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines


def _gradient_bg(c1: tuple, c2: tuple) -> "Image.Image":
    a = np.array(c1, dtype="float32")
    b = np.array(c2, dtype="float32")
    t = np.linspace(0, 1, VIDEO_H, dtype="float32")[:, None]
    data = (a * (1 - t) + b * t).astype("uint8")
    data = np.broadcast_to(data[:, None, :], (VIDEO_H, VIDEO_W, 3)).copy()
    return Image.fromarray(data)


def _shadow_text(draw, x: int, y: int, text: str, font, fill="white") -> None:
    """Teks dengan shadow 8-arah supaya terbaca di atas background apapun."""
    for ox, oy in [(-4,-4),(4,-4),(-4,4),(4,4),(0,-5),(0,5),(-5,0),(5,0)]:
        draw.text((x + ox, y + oy), text, font=font, fill=(0, 0, 0, 160))
    draw.text((x, y), text, font=font, fill=fill)


def _center_shadow(draw, y: int, text: str, font,
                   fill="white", max_w: int = None) -> int:
    """Gambar teks shadow di tengah. Return Y setelah teks."""
    mw    = max_w or (VIDEO_W - 100)
    lines = _wrap_text(text, font, mw, draw)
    for line in lines:
        bb = draw.textbbox((0, 0), line, font=font)
        x  = (VIDEO_W - (bb[2] - bb[0])) // 2
        _shadow_text(draw, x, y, line, font, fill)
        y += (bb[3] - bb[1]) + 18
    return y


def _parse_price(raw) -> float:
    if isinstance(raw, (int, float)):
        v = float(raw)
        return v / 100 if v > 1_000_000 else v
    s = re.sub(r"[^\d.,]", "", str(raw)).replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


# =============================================================================
# SCRAPING PRODUK TIKTOK SHOP
# =============================================================================

def scrape_products(keyword: str, max_n: int = MAX_PRODUCTS) -> list:
    """
    Scraping produk best-selling TikTok Shop.
    Strategi:
      1. Buka TikTok search shop dengan Playwright
      2. Intercept JSON API response yang berisi data produk
      3. Fallback: parse DOM
      4. Fallback: data demo
    Return: list of dict produk.
    """
    print(f"Scraping TikTok Shop: '{keyword}' ...")

    if SKIP_SCRAPING:
        print("  SKIP_SCRAPING = True, pakai PRODUK_MANUAL.\n")
        return PRODUK_MANUAL[:max_n]

    if not PW_OK:
        print("  Playwright tidak tersedia, pakai data demo.\n")
        return _demo_products(keyword)[:max_n]

    captured = []

    def _on_response(resp):
        url = resp.url
        if any(k in url for k in ["product", "goods", "item_list", "search"]):
            try:
                body = resp.json()
                captured.append(body)
            except Exception:
                pass

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=HEADLESS_SCRAPE,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        ctx = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            ),
            locale="id-ID",
            viewport={"width": 1280, "height": 800},
        )
        page = ctx.new_page()
        page.on("response", _on_response)

        search_url = (
            "https://www.tiktok.com/search?q="
            + urllib.parse.quote(keyword)
            + "&type=shop"
        )

        try:
            print(f"  Membuka: {search_url}")
            page.goto(search_url, wait_until="domcontentloaded", timeout=30_000)
            page.wait_for_timeout(4000)

            for _ in range(SEARCH_SCROLL):
                page.evaluate("window.scrollBy(0, 700)")
                page.wait_for_timeout(1500)

        except Exception as e:
            print(f"  [WARN] Navigasi: {e}")

        # Parse dari API yang ter-intercept
        products = _parse_captured_responses(captured, keyword)

        # Fallback DOM scrape
        if not products:
            print("  Coba scrape DOM ...")
            products = _scrape_dom_products(page, keyword)

        browser.close()

    if not products:
        print("  Scraping tidak dapat data, pakai demo ...")
        products = _demo_products(keyword)

    if SHUFFLE:
        random.shuffle(products)

    print(f"  {len(products)} produk ditemukan.\n")
    return products[:max_n]


def _parse_captured_responses(captured: list, keyword: str) -> list:
    """Parse JSON yang ter-intercept dari network request."""
    products = []
    seen_ids = set()

    for payload in captured:
        if not isinstance(payload, dict):
            continue
        # TikTok bisa wrap data di berbagai level
        items = (
            payload.get("data", {}).get("itemList")
            or payload.get("data", {}).get("products")
            or payload.get("data", {}).get("goods")
            or payload.get("data", {}).get("items")
            or payload.get("itemList")
            or payload.get("products")
            or []
        )
        if not items:
            continue

        for item in items:
            try:
                name = (
                    item.get("title")
                    or item.get("name")
                    or item.get("desc")
                    or ""
                ).strip()
                if not name:
                    continue

                pid = str(
                    item.get("itemId")
                    or item.get("productId")
                    or item.get("id")
                    or ""
                )
                if pid in seen_ids:
                    continue
                seen_ids.add(pid)

                price = _parse_price(
                    item.get("price")
                    or item.get("salePrice")
                    or item.get("minPrice")
                    or 0
                )

                # Kumpulkan URL gambar
                imgs = []
                for field in ["images", "imgUrls", "imageUrls"]:
                    val = item.get(field, [])
                    if isinstance(val, list):
                        imgs += [
                            v if isinstance(v, str) else v.get("url", "")
                            for v in val
                        ]
                cover = item.get("cover") or item.get("thumbnail") or ""
                if cover:
                    imgs.insert(0, cover)

                products.append({
                    "product_id" : pid or f"scrape_{len(products)}",
                    "name"       : name,
                    "price"      : price,
                    "rating"     : float(item.get("rating") or 4.5),
                    "sold_count" : int(
                        item.get("soldCount")
                        or item.get("salesVolume")
                        or item.get("sales")
                        or 0
                    ),
                    "image_urls" : [u for u in imgs if u],
                    "category"   : keyword,
                    "url"        : f"https://www.tiktok.com/shop/product/{pid}",
                    "description": item.get("description") or "",
                })
            except Exception:
                pass

    return products


def _scrape_dom_products(page, keyword: str) -> list:
    """Fallback: scrape elemen DOM produk."""
    products = []
    selectors = [
        "[data-e2e='shop-product-card']",
        "[class*='ProductCard']",
        "[class*='product-card']",
        ".shop-search-item",
    ]
    cards = []
    for sel in selectors:
        try:
            cards = page.query_selector_all(sel)
            if cards:
                break
        except Exception:
            pass

    for card in cards[:MAX_PRODUCTS * 3]:
        try:
            name_el  = card.query_selector(
                "[class*='title'], [class*='name'], h3, h2"
            )
            price_el = card.query_selector("[class*='price']")
            img_el   = card.query_selector("img")
            link_el  = card.query_selector("a")

            name  = (name_el.inner_text() if name_el else "").strip()
            if not name:
                continue
            price_txt = price_el.inner_text() if price_el else "0"
            img       = img_el.get_attribute("src") if img_el else ""
            href      = link_el.get_attribute("href") if link_el else ""

            m = re.search(r"/product/(\w+)", href or "")
            pid = m.group(1) if m else f"dom_{len(products)}"

            products.append({
                "product_id" : pid,
                "name"       : name,
                "price"      : _parse_price(price_txt),
                "rating"     : 4.5,
                "sold_count" : 0,
                "image_urls" : [img] if img else [],
                "category"   : keyword,
                "url"        : (href if href.startswith("http")
                                else f"https://www.tiktok.com{href}"),
                "description": "",
            })
        except Exception:
            pass
    return products


def _demo_products(keyword: str) -> list:
    """Data demo saat scraping tidak bisa jalan (offline / diblokir)."""
    return [
        {
            "product_id" : "demo_001",
            "name"       : "Serum Vitamin C Glowing Anti-Kusam 30ml",
            "price"      : 89_000,
            "rating"     : 4.8,
            "sold_count" : 15_420,
            "image_urls" : [],
            "category"   : keyword,
            "url"        : "https://www.tiktok.com/shop/product/demo_001",
            "description": "Cerahkan kulit dalam 7 hari. BPOM certified.",
        },
        {
            "product_id" : "demo_002",
            "name"       : "Sunscreen SPF50 PA++++ Ringan Tidak Lengket 50ml",
            "price"      : 65_000,
            "rating"     : 4.7,
            "sold_count" : 8_930,
            "image_urls" : [],
            "category"   : keyword,
            "url"        : "https://www.tiktok.com/shop/product/demo_002",
            "description": "Proteksi UV terbaik, cocok kulit berminyak.",
        },
        {
            "product_id" : "demo_003",
            "name"       : "Paket Skincare Acne Care Set 3in1",
            "price"      : 145_000,
            "rating"     : 4.9,
            "sold_count" : 22_100,
            "image_urls" : [],
            "category"   : keyword,
            "url"        : "https://www.tiktok.com/shop/product/demo_003",
            "description": "Cleanser + toner + moisturizer khusus jerawat.",
        },
    ]


# =============================================================================
# DOWNLOAD GAMBAR PRODUK
# =============================================================================

def _download_first_image(img_urls: list, dest: str) -> str:
    """Coba download dari list URL sampai satu berhasil."""
    for url in img_urls:
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) "
                              "Chrome/124.0.0.0 Safari/537.36",
                "Referer": "https://www.bing.com/",
            })
            with urllib.request.urlopen(req, timeout=8) as r:
                data = r.read()
            if len(data) < 2000:   # terlalu kecil, bukan gambar
                continue
            Image.open(io.BytesIO(data)).convert("RGB").save(dest, "JPEG")
            return dest
        except Exception:
            continue
    return ""


def _auto_search_image(product_name: str, pid: str) -> str:
    """Cari foto produk dari Tokopedia pakai Playwright. Return path lokal atau ''."""
    if not PW_OK:
        print("  [SKIP] Playwright tidak tersedia untuk cari foto.")
        return ""

    img_dir = os.path.join("output", "images", pid)
    os.makedirs(img_dir, exist_ok=True)
    dest = os.path.join(img_dir, "auto_img.jpg")
    if os.path.exists(dest):
        print(f"  Foto (cache): {dest}")
        return dest

    print(f"  Cari foto: '{product_name}' ...")
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            ctx  = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                           "AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/124.0.0.0 Safari/537.36",
                locale="id-ID",
                extra_http_headers={"Accept-Language": "id-ID,id;q=0.9"},
            )
            page = ctx.new_page()
            q = urllib.parse.quote_plus(product_name)
            page.goto(
                f"https://www.tokopedia.com/search?q={q}&st=product",
                timeout=30_000,
                wait_until="domcontentloaded",
            )
            page.wait_for_timeout(3500)
            # scroll sedikit agar lazy-load trigger
            page.evaluate("window.scrollBy(0, 600)")
            page.wait_for_timeout(1500)

            img_urls = page.evaluate("""
                () => Array.from(document.querySelectorAll('img'))
                     .map(i => i.src || i.currentSrc || i.getAttribute('data-src') || '')
                     .filter(s => s.startsWith('http') && s.length > 40 &&
                             !s.includes('icon') && !s.includes('logo') &&
                             !s.includes('svg') && !s.includes('gif') &&
                             (s.includes('.jpg') || s.includes('.jpeg') ||
                              s.includes('.png')  || s.includes('.webp')))
            """)
            browser.close()

        print(f"  Tokopedia: {len(img_urls)} gambar ditemukan")
        result = _download_first_image(img_urls[:15], dest)
        if result:
            print(f"  Foto OK: {dest}")
            return result

    except Exception as e:
        print(f"  [WARN] Foto gagal: {e}")

    print("  [WARN] Tidak ada foto berhasil didownload.")
    return ""


def download_product_images(product: dict) -> list:
    """Download/load gambar produk. Return list path PNG/JPG."""
    pid     = product["product_id"]
    img_dir = os.path.join("output", "images", pid)
    os.makedirs(img_dir, exist_ok=True)

    paths = []

    # Prioritas 1: image_path lokal (foto dari PC user)
    local = product.get("image_path", "").strip()
    if local and os.path.exists(local):
        print(f"  Gambar lokal: {local}")
        return [local]

    # Prioritas 2: download dari image_urls (hasil scraping)
    for i, url in enumerate(product.get("image_urls", [])[:3]):
        if not url:
            continue
        dest = os.path.join(img_dir, f"img_{i}.jpg")
        if os.path.exists(dest):
            paths.append(dest)
            continue
        try:
            req = urllib.request.Request(
                url, headers={"User-Agent": "Mozilla/5.0"}
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = resp.read()
            img = Image.open(io.BytesIO(data)).convert("RGB")
            img.save(dest, "JPEG")
            paths.append(dest)
            print(f"  Gambar {i+1}: {dest}")
        except Exception as e:
            print(f"  [WARN] Gambar {i+1} gagal: {e}")

    # Prioritas 3: cari otomatis di DuckDuckGo Images
    if not paths:
        auto = _auto_search_image(product["name"], pid)
        if auto:
            paths.append(auto)

    return paths


# =============================================================================
# GENERATE SCRIPT UGC
# =============================================================================

_HOOKS = {
    "skincare" : [
        "POV: nemu skincare yang actually works! 🔥",
        "Jujur review — ini yang bikin kulit gue glowing!",
        "STOP beli skincare mahal dulu, dengerin ini dulu!",
        "Skin goal tercapai gara-gara ini, serius!",
    ],
    "kecantikan": [
        "Makeup budget-an tapi hasilnya setara artis? Ada nih!",
        "Ini dia yang bikin ribuan orang kalap belanja!",
    ],
    "fashion": [
        "Outfit kece nggak harus mahal, ini buktinya!",
        "Tampil stylish under 100 ribu? Bisa banget!",
    ],
    "makanan": [
        "Snack ini bikin nagih parah! Jangan bilang aku nggak ngingetin!",
        "Kalau belum pernah coba ini, rugi banget!",
    ],
    "gadget": [
        "Gadget murah tapi kualitas premium? Ada!",
        "Ini alasan kenapa semua orang lagi rebutan beli ini!",
    ],
    "default": [
        "Nggak nyangka produk ini bisa sekece ini!",
        "Honest review: ini yang lagi paling banyak dibeli!",
        "Teman-teman, aku nemu sesuatu yang wajib kalian tau!",
    ],
}

_PROBLEMS = {
    "skincare": [
        "Udah coba banyak skincare tapi kulit tetep kusam? Aku juga dulu gitu.",
        "Nyari skincare yang cocok itu susah banget kan? Apalagi yang nggak bikin bruntusan.",
    ],
    "default": [
        "Pernah ragu beli produk ini? Aku udah coba duluan buat kamu!",
        "Udah capek nyari yang bagus tapi tetep kecewa? Nah, ini jawabannya.",
    ],
}

_CTAS = [
    "Tap keranjang di bawah sekarang sebelum kehabisan! ⬇️",
    "Link produknya ada di keranjang di bawah video ini! 🛒",
    "Buruan order, stok terbatas! Klik keranjang di bawah! 🔥",
    "Udah ribuan yang order, giliran kamu! Tap keranjang di bawah!",
]


def generate_script(product: dict) -> dict:
    """Buat script UGC + caption dari data produk."""
    name      = product["name"]
    price     = product["price"]
    rating    = product["rating"]
    sold      = product["sold_count"]
    cat       = product.get("category", "default").lower()

    # Pilih template sesuai kategori
    cat_key = next((k for k in _HOOKS if k in cat), "default")
    hook    = random.choice(_HOOKS[cat_key])
    problem = random.choice(_PROBLEMS.get(cat_key, _PROBLEMS["default"]))
    cta     = random.choice(_CTAS)

    solution = (
        f"Ini dia, {name}! "
        f"Harganya cuma Rp{price:,.0f} tapi kualitasnya luar biasa. "
        f"Rating {rating} bintang, sudah {sold:,} lebih yang order dan puas!"
    )
    if product.get("description"):
        solution += f" {product['description']}"

    voiceover = f"{hook} {problem} {solution} {cta}"

    # Hashtag
    kw_tags = [
        "#" + w.capitalize()
        for w in cat.split()
        if len(w) > 2
    ]
    hashtags = BASE_HASHTAGS + kw_tags
    ht_str   = " ".join(hashtags)

    caption = CAPTION_TEMPLATE.format(
        hook     = hook,
        nama     = name,
        harga    = f"{price:,.0f}",
        rating   = rating,
        hashtags = ht_str,
    )

    return {
        "hook"      : hook,
        "problem"   : problem,
        "solution"  : solution,
        "cta"       : cta,
        "voiceover" : voiceover,
        "caption"   : caption,
        "hashtags"  : hashtags,
    }


# =============================================================================
# GENERATE VOICEOVER (gTTS)
# =============================================================================

def generate_voiceover(text: str, filename: str) -> str:
    """
    Text-to-Speech gratis pakai gTTS.
    Return: path MP3 atau "" jika gagal.
    """
    if not GTTS_OK:
        print("  [SKIP] gTTS tidak tersedia.")
        return ""

    out = os.path.join(OUTPUT_FOLDER, f"{filename}_voice.mp3")
    print("  Generating voiceover ...")
    try:
        gTTS(text=text, lang=CONTENT_LANG, tld=TTS_TLD, slow=False).save(out)
        print(f"  Voiceover: {out}")
        return out
    except Exception as e:
        print(f"  [ERROR] TTS: {e}")
        return ""


# =============================================================================
# BUAT SLIDE PNG (PIL)
# =============================================================================

def _slide_hook(hook: str, emoji: str, palette: list) -> "Image.Image":
    """Slide 1: hook dramatis — gradient gelap + emoji besar + teks bold."""
    c1, c2 = palette
    # Gelap-kan warna palette untuk kesan dramatis
    dark1 = tuple(max(0, v - 80) for v in c1)
    dark2 = tuple(max(0, v - 60) for v in c2)
    img  = _gradient_bg(dark1, dark2)
    draw = ImageDraw.Draw(img)

    # Garis aksen atas
    draw.rectangle([0, 0, VIDEO_W, 12], fill=c1)

    # Emoji besar
    ef = _get_font(200)
    eb = draw.textbbox((0, 0), emoji, font=ef)
    draw.text(((VIDEO_W - (eb[2] - eb[0])) // 2, 220), emoji, font=ef,
              fill=(255, 255, 255))

    # Hook text
    font = _get_font(74)
    y = _center_shadow(draw, 640, hook, font, fill="white")

    # Tag "UGC Review" di bawah
    tag_font = _get_font(40)
    tag = "✨ UGC Review ✨"
    tb  = draw.textbbox((0, 0), tag, font=tag_font)
    draw.rounded_rectangle(
        [(VIDEO_W - (tb[2] - tb[0]) - 60) // 2, y + 60,
         (VIDEO_W + (tb[2] - tb[0]) + 60) // 2, y + 130],
        radius=35, fill=(255, 255, 255, 40),
    )
    _center_shadow(draw, y + 70, tag, tag_font, fill=(255, 255, 200))

    # Garis aksen bawah
    draw.rectangle([0, VIDEO_H - 12, VIDEO_W, VIDEO_H], fill=c2)
    return img


def _slide_product_hero(product: dict, img_path: str,
                        palette: list) -> "Image.Image":
    """Slide 2: foto produk + info — RGB only, tanpa blur/RGBA."""
    # Canvas RGB — tidak ada RGBA sama sekali
    canvas = _gradient_bg(*palette)
    draw   = ImageDraw.Draw(canvas)

    img_bottom = 80
    if img_path and os.path.exists(img_path):
        try:
            prod = Image.open(img_path).convert("RGB")
            prod.thumbnail((int(VIDEO_W * 0.78), int(VIDEO_H * 0.46)), Image.LANCZOS)
            pad = 10
            # Bingkai putih: gambar rectangle dulu, paste foto di atasnya
            x = (VIDEO_W - prod.width) // 2
            y = 70
            draw.rectangle([x - pad, y - pad,
                            x + prod.width + pad, y + prod.height + pad],
                           fill=(255, 255, 255))
            canvas.paste(prod, (x, y))
            img_bottom = y + prod.height + pad + 24
            prod.close(); del prod
        except Exception:
            pass

    # Info card: rectangle gelap (RGB — tidak perlu alpha)
    card_top = max(img_bottom, int(VIDEO_H * 0.56))
    draw.rectangle([0, card_top, VIDEO_W, VIDEO_H], fill=(12, 12, 24))

    fn  = _get_font(46)
    fp  = _get_font(66)
    fsm = _get_font(36)

    y = card_top + 30
    for line in _wrap_text(product["name"], fn, VIDEO_W - 60, draw)[:2]:
        bb = draw.textbbox((0, 0), line, font=fn)
        draw.text(((VIDEO_W - (bb[2] - bb[0])) // 2, y), line, font=fn, fill="white")
        y += (bb[3] - bb[1]) + 10

    y += 12
    price_txt = f"Rp{product['price']:,.0f}"
    pb = draw.textbbox((0, 0), price_txt, font=fp)
    draw.text(((VIDEO_W - (pb[2] - pb[0])) // 2, y), price_txt, font=fp, fill="#FF6235")
    y += (pb[3] - pb[1]) + 14

    stars = "★" * int(product["rating"]) + "☆" * (5 - int(product["rating"]))
    info  = f"{stars} {product['rating']}  |  {product['sold_count']:,}+ terjual"
    ib = draw.textbbox((0, 0), info, font=fsm)
    draw.text(((VIDEO_W - (ib[2] - ib[0])) // 2, y), info, font=fsm, fill="#FFD700")

    return canvas


def _slide_benefits(product: dict, script: dict) -> "Image.Image":
    """Slide 3: daftar keunggulan produk."""
    img  = _gradient_bg((14, 20, 48), (28, 45, 90))
    draw = ImageDraw.Draw(img)

    # Header
    fh  = _get_font(60)
    fi  = _get_font(50)
    fsm = _get_font(40)

    y = _center_shadow(draw, 160, "Kenapa wajib coba? 🤔", fh, fill="white")

    # Garis pemisah
    draw.rectangle([80, y + 20, VIDEO_W - 80, y + 26],
                   fill=(255, 255, 255, 80))
    y += 70

    # Daftar benefit
    items = []
    if product.get("description"):
        items.append(f"  {product['description'][:65]}")
    items.append(f"  Rating bintang {product['rating']}/5")
    items.append(f"  {product['sold_count']:,}+ pembeli sudah puas")
    items.append(f"  Harga terjangkau Rp{product['price']:,.0f}")

    icons = ["✅", "⭐", "👥", "💰"]
    for icon, item in zip(icons, items[:4]):
        full = icon + item
        lines = _wrap_text(full, fi, VIDEO_W - 100, draw)
        for j, line in enumerate(lines[:2]):
            _shadow_text(draw, 60, y, line, fi, fill="white")
            bb = draw.textbbox((0, 0), line, font=fi)
            y += (bb[3] - bb[1]) + 12
        y += 28

    # Footer
    _center_shadow(draw, VIDEO_H - 200, "Swipe untuk lihat lebih →", fsm,
                   fill=(200, 200, 255))
    return img


def _slide_cta(product: dict, cta: str) -> "Image.Image":
    """Slide 4: CTA dengan harga + tombol bold."""
    img  = _gradient_bg((200, 30, 10), (140, 10, 50))
    draw = ImageDraw.Draw(img)

    # Ikon keranjang
    ik = _get_font(220)
    ib = draw.textbbox((0, 0), "🛒", font=ik)
    draw.text(((VIDEO_W - (ib[2] - ib[0])) // 2, 130),
              "🛒", font=ik, fill="#FFD700")

    # Harga di atas CTA
    fp = _get_font(82)
    y  = 520
    pb = draw.textbbox((0, 0), f"Rp{product['price']:,.0f}", font=fp)
    draw.text(((VIDEO_W - (pb[2] - pb[0])) // 2, y),
              f"Rp{product['price']:,.0f}", font=fp, fill="#FFD700")
    y += (pb[3] - pb[1]) + 30

    # Teks CTA
    fc = _get_font(64)
    y  = _center_shadow(draw, y, cta, fc, fill="white")

    # Tombol
    btn_font = _get_font(52)
    btn_txt  = "TAP KERANJANG DI BAWAH ⬇"
    bb       = draw.textbbox((0, 0), btn_txt, font=btn_font)
    bw       = bb[2] - bb[0]
    btn_x    = (VIDEO_W - bw - 80) // 2
    btn_y    = y + 60
    draw.rounded_rectangle([btn_x, btn_y, btn_x + bw + 80, btn_y + 130],
                            radius=40, fill="#FFD700")
    draw.text((btn_x + 40, btn_y + 32), btn_txt, font=btn_font, fill="#1A0A00")

    # Urgency text
    fu = _get_font(42)
    _center_shadow(draw, btn_y + 170, "⚡ Stok terbatas — jangan sampai kehabisan!",
                   fu, fill=(255, 220, 180))
    return img




# =============================================================================
# BUAT VIDEO (PIL + ffmpeg)
# =============================================================================

def create_video(product: dict, script: dict,
                 voice_path: str, img_paths: list) -> str:
    """
    Buat video TikTok 1080×1920 dari slide PIL + voiceover.
    Return: path MP4 final atau "" jika gagal.
    """
    if not PIL_OK:
        print("  [SKIP] Pillow tidak tersedia, skip buat video.")
        return ""

    pid      = product["product_id"]
    palette  = get_palette(product.get("category", ""))
    slides_d = os.path.join(FRAMES_FOLDER, pid)
    os.makedirs(slides_d, exist_ok=True)

    # ── Render slide PNG ──────────────────────────────────────────────────────
    slide_defs = []
    img_path   = img_paths[0] if img_paths else ""

    # 1. Hook — dramatis
    s = _slide_hook(script["hook"], "🔥", palette)
    p = os.path.join(slides_d, "01_hook.png")
    s.save(p); s.close(); del s; gc.collect()
    slide_defs.append((p, SLIDE_DURATION))

    # 2. Produk hero — foto besar + info card
    s = _slide_product_hero(product, img_path, palette)
    p = os.path.join(slides_d, "02_product.png")
    s.save(p); s.close(); del s; gc.collect()
    slide_defs.append((p, SLIDE_DURATION + 1))

    # 3. Benefits — daftar keunggulan
    s = _slide_benefits(product, script)
    p = os.path.join(slides_d, "03_benefits.png")
    s.save(p); s.close(); del s; gc.collect()
    slide_defs.append((p, SLIDE_DURATION))

    # 5. CTA — tombol beli
    s = _slide_cta(product, script["cta"])
    p = os.path.join(slides_d, "04_cta.png")
    s.save(p); s.close(); del s; gc.collect()
    slide_defs.append((p, SLIDE_DURATION))

    print(f"  {len(slide_defs)} slide dibuat.")

    # ── Tiap slide → clip MP4 dengan ffmpeg ──────────────────────────────────
    clip_paths = []
    for i, (slide_png, dur) in enumerate(slide_defs):
        clip_out = os.path.join(slides_d, f"clip_{i:02d}.mp4")
        ok = run_ffmpeg([
            "-loop", "1", "-t", str(dur),
            "-i", slide_png,
            "-vf", f"scale={VIDEO_W}:{VIDEO_H},setsar=1",
            "-r", str(FPS),
            "-c:v", "libx264", "-preset", "ultrafast", "-crf", "28",
            "-pix_fmt", "yuv420p",
            clip_out,
        ], f"slide {i+1}")
        if ok:
            clip_paths.append(clip_out)

    if not clip_paths:
        print("  [ERROR] Tidak ada clip yang berhasil dibuat.")
        return ""

    # ── Concat semua clip ─────────────────────────────────────────────────────
    list_file = os.path.join(slides_d, "clips.txt")
    with open(list_file, "w") as f:
        for cp in clip_paths:
            f.write(f"file '{os.path.abspath(cp)}'\n")

    merged = os.path.join(OUTPUT_FOLDER, f"{pid}_merged.mp4")
    ok = run_ffmpeg([
        "-f", "concat", "-safe", "0",
        "-i", list_file,
        "-c", "copy",
        merged,
    ], "concat")

    if not ok or not os.path.exists(merged):
        print("  [ERROR] Concat clip gagal.")
        return ""

    # ── Tambahkan voiceover ───────────────────────────────────────────────────
    final = os.path.join(OUTPUT_FOLDER, f"{pid}_final.mp4")

    if voice_path and os.path.exists(voice_path):
        ok = run_ffmpeg([
            "-i", merged,
            "-i", voice_path,
            "-map", "0:v:0",
            "-map", "1:a:0",
            "-c:v", "copy",
            "-c:a", "aac", "-b:a", "128k",
            "-shortest",
            final,
        ], "add audio")
    else:
        shutil.copy2(merged, final)
        ok = True

    if ok and os.path.exists(final):
        size_mb = os.path.getsize(final) / 1e6
        print(f"  Video final: {final} ({size_mb:.1f} MB)")
        return final

    print("  [ERROR] File final tidak terbentuk.")
    return ""


# =============================================================================
# LOAD COOKIES (Netscape format)
# =============================================================================

def _load_netscape_cookies(path: str) -> list:
    """Parse cookies.txt format Netscape → list dict untuk Playwright."""
    cookies = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split("\t")
            if len(parts) < 7:
                continue
            domain, _, cpath, secure, expires, name, value = parts[:7]
            try:
                exp_i = int(expires)
            except ValueError:
                exp_i = -1
            cookies.append({
                "name"   : name,
                "value"  : value,
                "domain" : domain,
                "path"   : cpath,
                "secure" : secure.upper() == "TRUE",
                "expires": exp_i if exp_i > 0 else -1,
            })
    return cookies


# =============================================================================
# UPLOAD KE TIKTOK (Playwright + cookies)
# =============================================================================

def upload_to_tiktok(video_path: str, caption: str) -> bool:
    """
    Upload video ke TikTok Studio via Playwright + file cookies Netscape.
    Cookie didapat dengan login TikTok di Chrome → pakai ekstensi
    'Get cookies.txt LOCALLY' → simpan sebagai tiktok_cookies.txt.

    Return: True jika berhasil.
    """
    if not AUTO_UPLOAD:
        return False
    if not video_path or not os.path.exists(video_path):
        print("  [SKIP] File video tidak ditemukan.")
        return False
    if not os.path.exists(TIKTOK_COOKIE_FILE):
        print(f"  [SKIP] {TIKTOK_COOKIE_FILE} tidak ditemukan.")
        print("  Cara dapat cookies:")
        print("  1. Install ekstensi Chrome 'Get cookies.txt LOCALLY'")
        print("  2. Login ke tiktok.com")
        print("  3. Klik ekstensi → Export → simpan sebagai tiktok_cookies.txt")
        return False
    if not PW_OK:
        print("  [SKIP] Playwright tidak tersedia.")
        return False

    video_abs = os.path.abspath(video_path)
    print(f"  Upload TikTok: {os.path.basename(video_path)}")
    print(f"  Caption: {caption[:80]}...")

    try:
        cookies = _load_netscape_cookies(TIKTOK_COOKIE_FILE)

        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=HEADLESS_UPLOAD,
                args=["--no-sandbox"],
            )
            ctx  = browser.new_context()
            ctx.add_cookies(cookies)
            page = ctx.new_page()

            print("  Membuka TikTok Studio Upload ...")
            page.goto(
                "https://www.tiktok.com/tiktokstudio/upload",
                wait_until="domcontentloaded",
            )
            page.wait_for_timeout(3000)

            # Pilih file video
            print("  Memilih file ...")
            file_input = page.locator('input[type="file"]').first
            file_input.set_input_files(video_abs)

            print("  Menunggu video diproses TikTok (~15 detik) ...")
            page.wait_for_timeout(15_000)

            # Isi caption
            print("  Mengisi caption ...")
            try:
                cap_box = page.locator('div[contenteditable="true"]').first
                cap_box.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Delete")
                page.keyboard.type(caption[:2200], delay=15)
            except Exception as e:
                print(f"  (Caption auto-fill gagal: {e})")

            # Tutup popup tutorial / notifikasi
            for popup_sel in [
                'button:has-text("Mengerti")', 'button:has-text("Got it")',
                'button:has-text("OK")',        'button:has-text("Selesai")',
                'div[role="button"]:has-text("Mengerti")',
            ]:
                try:
                    btn = page.locator(popup_sel).first
                    if btn.count() > 0 and btn.is_visible():
                        btn.click(timeout=2000)
                        page.wait_for_timeout(1000)
                except Exception:
                    pass

            # Set privacy → Semua orang
            print("  Set privacy 'Semua orang' ...")
            for sel in [
                'label:has-text("Semua orang")',
                'label:has-text("Public")',
                'div[role="radio"]:has-text("Semua orang")',
                'span:has-text("Semua orang")',
            ]:
                try:
                    loc = page.locator(sel).first
                    if loc.count() and loc.is_visible():
                        loc.click()
                        break
                except Exception:
                    pass

            page.wait_for_timeout(2000)

            # Tunggu server selesai proses upload
            print("  Menunggu upload server selesai (~20 detik) ...")
            page.wait_for_timeout(20_000)

            # Debug: list tombol yang tersedia
            try:
                btns   = page.locator("button, div[role='button']").all()
                labels = []
                for b in btns[:30]:
                    try:
                        t = b.inner_text(timeout=400).strip()
                        if t and len(t) < 40:
                            labels.append(t)
                    except Exception:
                        pass
                print(f"  [DEBUG] Tombol terdeteksi: {labels}")
            except Exception:
                pass

            # Klik tombol Post
            print("  Mencoba klik tombol 'Post' ...")
            posted        = False
            post_selectors = [
                'button[data-e2e="post_video_button"]',
                'div[data-e2e="post_video_button"]',
                'button:has-text("Posting")',
                'button:has-text("Post")',
                'button:has-text("Publikasi")',
                'button:has-text("Kirim")',
                'div[role="button"]:has-text("Posting")',
                'button.TUXButton--primary',
            ]

            for sel in post_selectors:
                try:
                    btn = page.locator(sel).first
                    if btn.count() == 0:
                        continue
                    btn.wait_for(state="visible", timeout=3000)
                    # Tunggu tombol jadi enabled
                    for _ in range(20):
                        try:
                            if btn.is_enabled():
                                break
                        except Exception:
                            pass
                        page.wait_for_timeout(2000)
                    btn.scroll_into_view_if_needed()
                    btn.click(force=True, timeout=5000)
                    posted = True
                    print(f"  Tombol diklik: {sel}")
                    break
                except Exception:
                    continue

            if posted:
                page.wait_for_timeout(5000)
                # Verifikasi sukses
                success = False
                for _ in range(15):
                    url_now = page.url
                    content = page.content().lower()
                    if (
                        "success" in url_now
                        or "manage"  in url_now
                        or "berhasil" in content
                        or "your video is being uploaded" in content
                    ):
                        success = True
                        break
                    page.wait_for_timeout(2000)

                if success:
                    print("  SUKSES upload TikTok! ✅")
                else:
                    print("  WARNING: Klik post jalan tapi sukses belum terverifikasi.")
                    print("  >>> Cek browser manual, lalu tekan ENTER ...")
                    try:
                        input()
                    except EOFError:
                        pass
            else:
                print("\n  >>> GAGAL auto-klik. KLIK 'POST' MANUAL di browser.")
                print("  >>> Setelah post selesai, tekan ENTER di terminal.\n")
                try:
                    input()
                except EOFError:
                    pass

            browser.close()
            print("  Browser ditutup.")
            return posted

    except Exception as e:
        print(f"  [ERROR] Upload TikTok: {e}")
        return False


# =============================================================================
# MAIN PROGRAM
# =============================================================================

if __name__ == "__main__":
  try:
    print("=" * 55)
    print("  TikTok UGC Automation")
    print(f"  Keyword    : {KEYWORD}")
    print(f"  Max produk : {MAX_PRODUCTS}")
    print(f"  Auto upload: {'YA' if AUTO_UPLOAD else 'TIDAK'}")
    print("=" * 55 + "\n")

    # ── 1. Scraping produk ────────────────────────────────────────────────────
    products = scrape_products(KEYWORD, MAX_PRODUCTS)
    if not products:
        print("Tidak ada produk. Selesai.")
        sys.exit(0)

    processed = 0

    for i, product in enumerate(products, 1):
        pid = product["product_id"]
        print(f"\n{'─'*55}")
        print(f"[{i}/{len(products)}] {product['name']}")
        print(
            f"  Harga: Rp{product['price']:,.0f} | "
            f"Rating: {product['rating']} | "
            f"Terjual: {product['sold_count']:,}"
        )

        final_path = ""

        # ── 2. Download gambar ────────────────────────────────────────────────
        img_paths = download_product_images(product)
        if not img_paths:
            print("  [SKIP] Tidak ada foto produk — video tidak dibuat.")
            continue

        # ── 3. Generate script ────────────────────────────────────────────────
        script = generate_script(product)
        print(f"  Hook: {script['hook'][:65]}...")

        # ── 4. Voiceover ──────────────────────────────────────────────────────
        voice_path = generate_voiceover(script["voiceover"], pid)

        # ── 5. Buat video ─────────────────────────────────────────────────────
        print("  Membuat video ...")
        final_path = create_video(product, script, voice_path, img_paths)

        if not final_path:
            print("  [SKIP] Video tidak terbentuk.")
            continue

        # ── 6. Upload ke TikTok ───────────────────────────────────────────────
        script_for_caption = generate_script(product)
        upload_to_tiktok(final_path, script_for_caption["caption"])

        processed += 1

        if processed < len(products):
            delay = random.randint(8, 15)
            print(f"\nTunggu {delay} detik sebelum produk berikutnya ...\n")
            gc.collect()
            time.sleep(delay)

    print("\n" + "=" * 55)
    print(f"  Selesai! {processed}/{len(products)} produk diproses.")
    print(f"  Output: {os.path.abspath(OUTPUT_FOLDER)}")
    print("=" * 55)
  except Exception as _err:
    import traceback
    print("\n" + "!" * 55)
    print("  SCRIPT CRASH — salin error di bawah ini ke chat:")
    print("!" * 55)
    traceback.print_exc()
  finally:
    input("\nTekan Enter untuk keluar...")
