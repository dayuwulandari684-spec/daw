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

# Google Veo 3 (video AI — jauh lebih bagus dari PIL)
# Daftar gratis di: https://aistudio.google.com → Get API Key
GOOGLE_API_KEY = ""        # isi dengan API key dari Google AI Studio
USE_VEO3       = False     # ganti True setelah isi GOOGLE_API_KEY

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
    from PIL import Image, ImageDraw, ImageFont
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
    data = np.zeros((VIDEO_H, VIDEO_W, 3), dtype="uint8")
    a, b = np.array(c1), np.array(c2)
    for row in range(VIDEO_H):
        t = row / VIDEO_H
        data[row] = (a * (1 - t) + b * t).astype("uint8")
    return Image.fromarray(data)


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

def download_product_images(product: dict) -> list:
    """Download gambar produk ke lokal. Return list path PNG/JPG."""
    pid     = product["product_id"]
    img_dir = os.path.join("output", "images", pid)
    os.makedirs(img_dir, exist_ok=True)

    paths = []
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

def _slide_text(text: str, palette: list,
                emoji: str = "", font_size: int = 65) -> "Image.Image":
    img  = _gradient_bg(*palette)
    draw = ImageDraw.Draw(img)
    font = _get_font(font_size)

    if emoji:
        ef = _get_font(160)
        eb = draw.textbbox((0, 0), emoji, font=ef)
        ew = eb[2] - eb[0]
        draw.text(((VIDEO_W - ew) // 2, 140), emoji, font=ef, fill=(255, 255, 255))

    lines   = _wrap_text(text, font, VIDEO_W - 120, draw)
    total_h = len(lines) * (font_size + 18)
    y       = (VIDEO_H - total_h) // 2 + (80 if emoji else 0)

    for line in lines:
        bb = draw.textbbox((0, 0), line, font=font)
        tw = bb[2] - bb[0]
        x  = (VIDEO_W - tw) // 2
        draw.text((x + 3, y + 3), line, fill=(0, 0, 0, 100), font=font)
        draw.text((x,     y),     line, fill=(255, 255, 255), font=font)
        y += font_size + 18

    return img


def _slide_product(product: dict, img_path: str, palette: list) -> "Image.Image":
    canvas = _gradient_bg(*palette)

    text_y = VIDEO_H // 2 - 200
    if img_path and os.path.exists(img_path):
        try:
            prod_img = Image.open(img_path).convert("RGBA")
            prod_img.thumbnail(
                (int(VIDEO_W * 0.85), int(VIDEO_H * 0.52)), Image.LANCZOS
            )
            x = (VIDEO_W - prod_img.width) // 2
            canvas.paste(prod_img, (x, 80),
                         prod_img if prod_img.mode == "RGBA" else None)
            text_y = 80 + prod_img.height + 40
        except Exception:
            pass

    draw = ImageDraw.Draw(canvas)
    big  = _get_font(55)
    med  = _get_font(44)

    for line in _wrap_text(product["name"], big, VIDEO_W - 80, draw)[:2]:
        bb = draw.textbbox((0, 0), line, font=big)
        tw = bb[2] - bb[0]
        draw.text(((VIDEO_W - tw) // 2, text_y), line, fill="white", font=big)
        text_y += 68

    draw.text((80, text_y + 10),
              f"Rp{product['price']:,.0f}", fill="#FFD700", font=big)
    draw.text((80, text_y + 82),
              f"⭐ {product['rating']}  |  {product['sold_count']:,}+ terjual",
              fill="white", font=med)
    return canvas


def _slide_cta(cta: str) -> "Image.Image":
    img  = _gradient_bg((255, 69, 0), (220, 20, 60))
    draw = ImageDraw.Draw(img)
    big  = _get_font(72)
    med  = _get_font(52)

    ikon = _get_font(200)
    ib   = draw.textbbox((0, 0), "🛒", font=ikon)
    draw.text(((VIDEO_W - (ib[2] - ib[0])) // 2, 160),
              "🛒", font=ikon, fill="#FFD700")

    lines = _wrap_text(cta, big, VIDEO_W - 100, draw)
    y = 580
    for line in lines:
        bb = draw.textbbox((0, 0), line, font=big)
        tw = bb[2] - bb[0]
        draw.text(((VIDEO_W - tw) // 2 + 3, y + 3), line,
                  fill=(0, 0, 0, 80), font=big)
        draw.text(((VIDEO_W - tw) // 2,     y),     line,
                  fill="white",       font=big)
        y += 92

    btn_y = y + 60
    draw.rounded_rectangle(
        [120, btn_y, VIDEO_W - 120, btn_y + 160], radius=40, fill="#FFD700"
    )
    btn_text = "TAP KERANJANG DI BAWAH ⬇"
    bb = draw.textbbox((0, 0), btn_text, font=med)
    draw.text(
        ((VIDEO_W - (bb[2] - bb[0])) // 2, btn_y + 55),
        btn_text, fill="#1A1A1A", font=med,
    )
    return img


# =============================================================================
# BUAT VIDEO PAKAI GOOGLE VEO 3
# =============================================================================

def _build_veo3_prompt(product: dict, script: dict) -> str:
    """Buat prompt Veo 3 berdasarkan data produk."""
    name  = product["name"]
    price = product["price"]
    cat   = product.get("category", "produk")
    hook  = script["hook"]
    sol   = script["solution"][:120]

    return (
        f"Vertical TikTok UGC video (9:16 aspect ratio) for Indonesian audience. "
        f"An enthusiastic young Indonesian woman reviews '{name}'. "
        f"She holds the product close to camera, smiles and speaks excitedly. "
        f"Bright natural lighting, clean modern background. "
        f"Text overlay shows: '{name}' and 'Rp{price:,.0f}'. "
        f"The video feels authentic, relatable, and viral. "
        f"Opening line (spoken in Indonesian): '{hook}'. "
        f"She demonstrates: '{sol}'. "
        f"End with clear product shot and satisfied expression."
    )


def create_video_veo3(product: dict, script: dict, voice_path: str) -> str:
    """
    Generate video UGC pakai Google Veo 3.
    Fallback ke PIL + ffmpeg jika gagal.
    """
    if not GOOGLE_API_KEY:
        print("  [SKIP] GOOGLE_API_KEY kosong. Fallback ke PIL.")
        return ""

    try:
        from google import genai
        from google.genai.types import GenerateVideoConfig
    except ImportError:
        print("  [ERROR] google-genai belum diinstall. Jalankan: pip install google-genai")
        return ""

    pid    = product["product_id"]
    prompt = _build_veo3_prompt(product, script)
    print(f"  Veo3 prompt: {prompt[:80]}...")
    print("  Generating video dengan Google Veo 3 (bisa 1-3 menit)...")

    try:
        client    = genai.Client(api_key=GOOGLE_API_KEY)
        operation = client.models.generate_video(
            model  = "veo-3.0-generate-preview",
            prompt = prompt,
            config = GenerateVideoConfig(
                aspect_ratio     = "9:16",
                number_of_videos = 1,
                duration_seconds = 8,
            ),
        )

        # Polling sampai selesai (max 3 menit)
        waited = 0
        while not operation.done and waited < 180:
            time.sleep(10)
            waited += 10
            print(f"  Menunggu Veo3... ({waited}s)")
            operation = client.operations.get(operation.name)

        if not operation.done:
            print("  [TIMEOUT] Veo3 timeout. Fallback ke PIL.")
            return ""

        videos = getattr(operation.result, "generated_videos", [])
        if not videos:
            print("  [ERROR] Veo3 tidak menghasilkan video.")
            return ""

        # Simpan video mentah
        raw_path = os.path.join(OUTPUT_FOLDER, f"{pid}_veo3_raw.mp4")
        vid = videos[0].video

        if getattr(vid, "video_bytes", None):
            with open(raw_path, "wb") as f:
                f.write(vid.video_bytes)
        elif getattr(vid, "uri", None):
            urllib.request.urlretrieve(vid.uri, raw_path)
        else:
            print("  [ERROR] Format video Veo3 tidak dikenali.")
            return ""

        size_mb = os.path.getsize(raw_path) / 1e6
        print(f"  Video Veo3 OK: {raw_path} ({size_mb:.1f} MB)")

        # Tambahkan voiceover
        final_path = os.path.join(OUTPUT_FOLDER, f"{pid}_final.mp4")
        if voice_path and os.path.exists(voice_path):
            ok = run_ffmpeg([
                "-i", raw_path,
                "-i", voice_path,
                "-map", "0:v:0",
                "-map", "1:a:0",
                "-c:v", "copy",
                "-c:a", "aac", "-b:a", "128k",
                "-shortest",
                final_path,
            ], "veo3 + voiceover")
            if not ok:
                shutil.copy2(raw_path, final_path)
        else:
            shutil.copy2(raw_path, final_path)

        print(f"  Final: {final_path}")
        return final_path

    except Exception as e:
        print(f"  [ERROR] Veo3: {e}")
        return ""


# =============================================================================
# BUAT VIDEO (PIL + ffmpeg) — fallback jika tidak pakai Veo 3
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

    # 1. Hook
    s = _slide_text(script["hook"], palette, emoji="🔥", font_size=68)
    p = os.path.join(slides_d, "01_hook.png")
    s.save(p)
    slide_defs.append((p, SLIDE_DURATION))

    # 2. Problem
    s = _slide_text(script["problem"], [(80, 80, 80), (30, 30, 30)],
                    emoji="😩", font_size=60)
    p = os.path.join(slides_d, "02_problem.png")
    s.save(p)
    slide_defs.append((p, SLIDE_DURATION))

    # 3. Produk (satu slide per gambar, max 3)
    if img_paths:
        for i, ip in enumerate(img_paths[:3]):
            s = _slide_product(product, ip, palette)
            p = os.path.join(slides_d, f"03_product_{i}.png")
            s.save(p)
            slide_defs.append((p, SLIDE_DURATION))
    else:
        s = _slide_product(product, "", palette)
        p = os.path.join(slides_d, "03_product.png")
        s.save(p)
        slide_defs.append((p, SLIDE_DURATION))

    # 4. CTA
    s = _slide_cta(script["cta"])
    p = os.path.join(slides_d, "04_cta.png")
    s.save(p)
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
            "-c:v", "libx264", "-preset", "veryfast", "-crf", "23",
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

        final_path = os.path.join(OUTPUT_FOLDER, f"{pid}_final.mp4")

        if os.path.exists(final_path):
            print(f"  Video sudah ada ({final_path}), skip generate.")
        else:
            # ── 2. Download gambar ────────────────────────────────────────────
            img_paths = download_product_images(product)

            # ── 3. Generate script ────────────────────────────────────────────
            script = generate_script(product)
            print(f"  Hook: {script['hook'][:65]}...")

            # ── 4. Voiceover ──────────────────────────────────────────────────
            voice_path = generate_voiceover(script["voiceover"], pid)

            # ── 5. Buat video ─────────────────────────────────────────────────
            print("  Membuat video ...")
            if USE_VEO3 and GOOGLE_API_KEY:
                print("  Mode: Google Veo 3")
                final_path = create_video_veo3(product, script, voice_path)
                if not final_path:
                    print("  Veo3 gagal, fallback ke PIL + ffmpeg...")
                    final_path = create_video(product, script, voice_path, img_paths)
            else:
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
            time.sleep(delay)

    print("\n" + "=" * 55)
    print(f"  Selesai! {processed}/{len(products)} produk diproses.")
    print(f"  Output: {os.path.abspath(OUTPUT_FOLDER)}")
    print("=" * 55)
