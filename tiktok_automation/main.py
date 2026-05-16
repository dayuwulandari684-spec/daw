"""
TikTok UGC Automation

Jalankan di CMD / Terminal:

  # Setup pertama kali
  python main.py --setup

  # Autentikasi TikTok (wajib sekali)
  python main.py --auth

  # Jalankan pipeline penuh (scrape → generate → upload)
  python main.py --run

  # Hanya scrape produk
  python main.py --scrape --category beauty

  # Hanya generate video (tanpa upload)
  python main.py --generate --product-id demo_001

  # Lihat daftar produk yang sudah di-scrape
  python main.py --list
"""

import argparse
import json
import os
import sys
import time
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs

# Colorama untuk output berwarna di CMD Windows
try:
    from colorama import init, Fore, Style
    init(autoreset=True)
    GREEN  = Fore.GREEN
    YELLOW = Fore.YELLOW
    RED    = Fore.RED
    CYAN   = Fore.CYAN
    BOLD   = Style.BRIGHT
    RESET  = Style.RESET_ALL
except ImportError:
    GREEN = YELLOW = RED = CYAN = BOLD = RESET = ""

from tqdm import tqdm

import config
from scraper  import ProductScraper, Product
from generator import ScriptGenerator, VoiceGenerator, VideoGenerator
from uploader  import TikTokUploader


# ── File state ─────────────────────────────────────────────────────────────────
STATE_FILE = Path("output/state.json")


def load_state() -> dict:
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text())
    return {"products": [], "generated": [], "uploaded": []}


def save_state(state: dict):
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False))


# ── Pipeline ───────────────────────────────────────────────────────────────────
def run_pipeline(category: str = "", limit: int = None):
    state  = load_state()
    limit  = limit or config.MAX_PRODUCTS

    print(f"\n{BOLD}{CYAN}═══════════════════════════════════════{RESET}")
    print(f"{BOLD}{CYAN}   TikTok UGC Automation Pipeline{RESET}")
    print(f"{BOLD}{CYAN}═══════════════════════════════════════{RESET}\n")

    # ── TAHAP 1: Scraping produk ───────────────────────────────────────────────
    print(f"{BOLD}[1/4] Scraping produk best-selling TikTok Shop...{RESET}")
    scraper  = ProductScraper(max_products=limit)
    products = scraper.scrape(category=category)

    if not products:
        print(f"{RED}Tidak ada produk yang ditemukan.{RESET}")
        return

    print(f"{GREEN}✓ {len(products)} produk ditemukan:{RESET}")
    for i, p in enumerate(products, 1):
        print(f"  {i}. {p.name[:60]} | Rp{p.price:,.0f} | ⭐{p.rating} | {p.sold_count:,} terjual")

    state["products"] = [_product_to_dict(p) for p in products]
    save_state(state)

    # ── TAHAP 2, 3, 4: Generate + Upload per produk ────────────────────────────
    script_gen = ScriptGenerator()
    voice_gen  = VoiceGenerator()
    video_gen  = VideoGenerator()
    uploader   = TikTokUploader()

    for product in products:
        print(f"\n{BOLD}{'─'*50}{RESET}")
        print(f"{BOLD}Produk: {product.name}{RESET}")

        # ── TAHAP 2: Generate script + voice ──────────────────────────────────
        print(f"\n{BOLD}[2/4] Membuat script UGC...{RESET}")
        script = script_gen.generate(product)
        print(f"  Hook    : {script.hook[:70]}...")
        print(f"  CTA     : {script.cta}")

        print(f"\n{BOLD}[3/4] Membuat voiceover...{RESET}")
        voiceover_text = script.full_voiceover()
        voice_path     = voice_gen.generate(voiceover_text, filename=product.product_id)
        print(f"  {GREEN}✓ Audio: {voice_path}{RESET}")

        # ── TAHAP 3: Generate video ────────────────────────────────────────────
        print(f"\n{BOLD}[3/4] Membuat video UGC...{RESET}")
        try:
            video_path = video_gen.create(
                product     = product,
                script      = script,
                voice_path  = voice_path,
                output_name = product.product_id,
            )
            print(f"  {GREEN}✓ Video: {video_path}{RESET}")
            state["generated"].append({"product_id": product.product_id, "video": video_path})
            save_state(state)
        except Exception as e:
            print(f"  {RED}✗ Video gagal: {e}{RESET}")
            continue

        # ── TAHAP 4: Upload ke TikTok + Keranjang Kuning ──────────────────────
        print(f"\n{BOLD}[4/4] Upload ke TikTok + pasang keranjang kuning...{RESET}")

        # Cari product ID di TikTok Shop untuk keranjang kuning
        shop_product_id = uploader.get_shop_product_id(product.name)
        product_ids     = [shop_product_id] if shop_product_id else []

        if not product_ids and product.product_id.startswith("demo_"):
            print(f"  {YELLOW}⚠ Demo mode: product tag tidak tersedia tanpa TikTok Shop API{RESET}")

        result = uploader.upload(
            video_path  = video_path,
            caption     = script.caption,
            product_ids = product_ids,
            hashtags    = script.hashtags,
        )

        if result.success:
            print(f"  {GREEN}✓ Upload sukses!{RESET}")
            print(f"  URL: {result.video_url}")
            state["uploaded"].append({
                "product_id": product.product_id,
                "publish_id": result.publish_id,
                "video_url" : result.video_url,
            })
        else:
            print(f"  {RED}✗ Upload gagal: {result.error}{RESET}")

        save_state(state)

    print(f"\n{BOLD}{GREEN}═══════════════════════════════════════{RESET}")
    print(f"{BOLD}{GREEN}   Pipeline selesai!{RESET}")
    print(f"{BOLD}{GREEN}═══════════════════════════════════════{RESET}\n")
    _print_summary(state)


# ── Auth flow ──────────────────────────────────────────────────────────────────
def run_auth():
    """Jalankan OAuth flow untuk mendapatkan TikTok access token."""
    uploader = TikTokUploader()
    if not uploader.client_key:
        print(f"{RED}ERROR: TIKTOK_CLIENT_KEY belum diset di .env{RESET}")
        print("1. Daftar di https://developers.tiktok.com")
        print("2. Buat app baru")
        print("3. Copy Client Key & Secret ke .env")
        return

    auth_url = uploader.get_auth_url()
    print(f"\n{BOLD}Membuka browser untuk login TikTok...{RESET}")
    print(f"URL: {auth_url}\n")
    webbrowser.open(auth_url)

    # Local server untuk menerima callback
    received_code = []

    class CallbackHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            qs   = parse_qs(urlparse(self.path).query)
            code = qs.get("code", [""])[0]
            if code:
                received_code.append(code)
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b"<h2>Berhasil! Tutup tab ini dan kembali ke terminal.</h2>")
            else:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b"Code tidak ditemukan")

        def log_message(self, *args):
            pass  # suppress log

    server = HTTPServer(("localhost", 8080), CallbackHandler)
    print(f"{YELLOW}Menunggu callback dari TikTok di port 8080...{RESET}")
    while not received_code:
        server.handle_request()

    code = received_code[0]
    print(f"\n{GREEN}Code diterima, menukar dengan access token...{RESET}")
    data = uploader.exchange_code_for_token(code)

    if data.get("access_token"):
        env_path = Path(".env")
        if env_path.exists():
            content = env_path.read_text()
            if "TIKTOK_ACCESS_TOKEN=" in content:
                lines = content.splitlines()
                lines = [
                    f"TIKTOK_ACCESS_TOKEN={data['access_token']}"
                    if l.startswith("TIKTOK_ACCESS_TOKEN=") else l
                    for l in lines
                ]
                env_path.write_text("\n".join(lines) + "\n")
            else:
                with open(env_path, "a") as f:
                    f.write(f"\nTIKTOK_ACCESS_TOKEN={data['access_token']}\n")
        print(f"\n{GREEN}✓ Access token berhasil disimpan ke .env!{RESET}")
        print(f"  Expires in: {data.get('expires_in', 'N/A')} detik")
    else:
        print(f"{RED}Gagal mendapatkan token: {data}{RESET}")


# ── Setup ──────────────────────────────────────────────────────────────────────
def run_setup():
    """Install dependencies dan siapkan lingkungan."""
    print(f"{BOLD}Setup TikTok Automation...{RESET}\n")

    steps = [
        ("Install Python packages", "pip install -r requirements.txt"),
        ("Install Playwright browsers", "playwright install chromium"),
    ]

    for name, cmd in steps:
        print(f"▶ {name}...")
        ret = os.system(cmd)
        if ret == 0:
            print(f"  {GREEN}✓ Selesai{RESET}")
        else:
            print(f"  {RED}✗ Gagal (kode: {ret}){RESET}")

    # Buat .env jika belum ada
    env_path = Path(".env")
    if not env_path.exists():
        import shutil
        if Path(".env.example").exists():
            shutil.copy(".env.example", ".env")
            print(f"\n{YELLOW}⚠ File .env dibuat dari .env.example{RESET}")
            print(f"  Isi nilai di .env dengan credentials TikTok kamu!")
        else:
            print(f"{RED}File .env.example tidak ditemukan{RESET}")

    print(f"\n{GREEN}Setup selesai!{RESET}")
    print(f"\nLangkah selanjutnya:")
    print(f"  1. Edit .env dan isi TIKTOK_CLIENT_KEY + TIKTOK_CLIENT_SECRET")
    print(f"  2. Jalankan: {BOLD}python main.py --auth{RESET}  (login ke TikTok)")
    print(f"  3. Jalankan: {BOLD}python main.py --run{RESET}   (mulai automation)")


# ── Utils ──────────────────────────────────────────────────────────────────────
def _product_to_dict(p: Product) -> dict:
    return {
        "product_id" : p.product_id,
        "name"       : p.name,
        "price"      : p.price,
        "rating"     : p.rating,
        "sold_count" : p.sold_count,
        "category"   : p.category,
        "product_url": p.product_url,
    }


def _print_summary(state: dict):
    prods    = len(state.get("products", []))
    generated = len(state.get("generated", []))
    uploaded  = len(state.get("uploaded", []))
    print(f"  Produk ditemukan : {prods}")
    print(f"  Video dibuat     : {generated}")
    print(f"  Berhasil diupload: {uploaded}")
    if state.get("uploaded"):
        print(f"\n  Link video:")
        for item in state["uploaded"]:
            print(f"  → {item.get('video_url', 'N/A')}")


# ── Entry point ────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="TikTok UGC Automation - Scrape, Generate, Upload",
        formatter_class=argparse.RawTextHelpFormatter,
    )

    parser.add_argument("--setup",    action="store_true",
                        help="Install dependencies dan siapkan environment")
    parser.add_argument("--auth",     action="store_true",
                        help="Login ke TikTok via OAuth (wajib sekali)")
    parser.add_argument("--run",      action="store_true",
                        help="Jalankan pipeline lengkap: scrape → generate → upload")
    parser.add_argument("--scrape",   action="store_true",
                        help="Hanya scrape produk, simpan ke state.json")
    parser.add_argument("--generate", action="store_true",
                        help="Hanya generate video dari produk di state.json")
    parser.add_argument("--list",     action="store_true",
                        help="Tampilkan daftar produk dan status")
    parser.add_argument("--category", default="",
                        help="Kategori produk: beauty, fashion, electronics, food")
    parser.add_argument("--limit",    type=int, default=None,
                        help="Jumlah produk yang diproses (default dari config)")

    args = parser.parse_args()

    if args.setup:
        run_setup()

    elif args.auth:
        run_auth()

    elif args.run:
        run_pipeline(category=args.category, limit=args.limit)

    elif args.scrape:
        print(f"{BOLD}Scraping produk...{RESET}")
        scraper  = ProductScraper(max_products=args.limit or config.MAX_PRODUCTS)
        products = scraper.scrape(category=args.category)
        state    = load_state()
        state["products"] = [_product_to_dict(p) for p in products]
        save_state(state)
        print(f"{GREEN}✓ {len(products)} produk disimpan ke output/state.json{RESET}")
        for i, p in enumerate(products, 1):
            print(f"  {i}. {p.name} | Rp{p.price:,.0f}")

    elif args.list:
        state = load_state()
        print(f"\n{BOLD}Status Pipeline:{RESET}")
        _print_summary(state)

    else:
        parser.print_help()
        print(f"\n{YELLOW}Contoh cepat:{RESET}")
        print(f"  python main.py --setup          # pertama kali")
        print(f"  python main.py --auth           # login TikTok")
        print(f"  python main.py --run            # jalankan semua")
        print(f"  python main.py --run --category beauty --limit 3")


if __name__ == "__main__":
    main()
