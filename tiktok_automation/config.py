import os
from dotenv import load_dotenv

load_dotenv()

# ── TikTok Developer App (https://developers.tiktok.com) ──────────────────────
TIKTOK_CLIENT_KEY     = os.getenv("TIKTOK_CLIENT_KEY", "")
TIKTOK_CLIENT_SECRET  = os.getenv("TIKTOK_CLIENT_SECRET", "")
TIKTOK_ACCESS_TOKEN   = os.getenv("TIKTOK_ACCESS_TOKEN", "")
TIKTOK_REDIRECT_URI   = os.getenv("TIKTOK_REDIRECT_URI", "http://localhost:8080/callback")

# ── TikTok Shop Affiliate (https://affiliate.tiktok.com) ──────────────────────
TIKTOK_SHOP_APP_KEY    = os.getenv("TIKTOK_SHOP_APP_KEY", "")
TIKTOK_SHOP_APP_SECRET = os.getenv("TIKTOK_SHOP_APP_SECRET", "")
TIKTOK_SHOP_ACCESS_TOKEN = os.getenv("TIKTOK_SHOP_ACCESS_TOKEN", "")

# ── OpenAI (opsional – untuk script UGC lebih baik) ───────────────────────────
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")

# ── Direktori output ───────────────────────────────────────────────────────────
OUTPUT_DIR     = "output"
VIDEOS_DIR     = "output/videos"
AUDIO_DIR      = "output/audio"
IMAGES_DIR     = "output/images"
THUMBS_DIR     = "output/thumbnails"

# ── Pengaturan video (TikTok vertical) ────────────────────────────────────────
VIDEO_WIDTH    = 1080
VIDEO_HEIGHT   = 1920
VIDEO_FPS      = 30
SLIDE_DURATION = 4       # detik per slide produk
FADE_DURATION  = 0.5     # detik fade antar slide

# ── Pengaturan scraping ────────────────────────────────────────────────────────
MAX_PRODUCTS      = 5    # jumlah produk best-selling yang diambil
SCRAPE_HEADLESS   = True # False = buka browser (berguna debug)
SCRAPE_TIMEOUT_MS = 30_000

# ── Bahasa konten ──────────────────────────────────────────────────────────────
CONTENT_LANGUAGE = "id"  # "id" = Bahasa Indonesia, "en" = English
TTS_LANGUAGE     = "id"
TTS_TLD          = "co.id"  # domain untuk aksen Indonesia
