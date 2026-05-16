"""
Scraper produk best-selling TikTok Shop.

Strategi:
  1. Buka TikTok Shop via Playwright (handle JS rendering)
  2. Navigasi ke halaman Flash Sale / Best Seller
  3. Ambil data produk: nama, harga, rating, jumlah terjual, gambar, product_id
"""

import asyncio
import json
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

import requests
from playwright.async_api import async_playwright, Page, TimeoutError as PWTimeout

import config


@dataclass
class Product:
    product_id: str
    name: str
    price: float
    currency: str
    rating: float
    sold_count: int
    image_urls: List[str]
    shop_name: str
    product_url: str
    category: str = ""
    description: str = ""
    commission_rate: float = 0.0  # untuk afiliasi


class ProductScraper:
    TIKTOK_SHOP_URL  = "https://www.tiktok.com/shop"
    FLASH_SALE_URL   = "https://www.tiktok.com/shop/flash-sale"
    BEST_SELLER_URL  = "https://www.tiktok.com/shop/top"

    # TikTok internal API yang dipakai website-nya
    _API_BEST_SELL = (
        "https://www.tiktok.com/api/recommend/item_list/"
        "?aid=1988&count=20&type=5"  # type=5 = best selling
    )

    def __init__(self, max_products: int = config.MAX_PRODUCTS,
                 headless: bool = config.SCRAPE_HEADLESS):
        self.max_products = max_products
        self.headless     = headless
        self._products: List[Product] = []

    # ── Public entry point ─────────────────────────────────────────────────────
    def scrape(self, category: str = "") -> List[Product]:
        """Jalankan scraping secara sinkron (wrapper asyncio)."""
        return asyncio.run(self._scrape_async(category))

    # ── Async core ─────────────────────────────────────────────────────────────
    async def _scrape_async(self, category: str) -> List[Product]:
        async with async_playwright() as pw:
            browser = await pw.chromium.launch(
                headless=self.headless,
                args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
            )
            context = await browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/122.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1280, "height": 800},
                locale="id-ID",
            )
            page = await context.new_page()

            # Intercept network request untuk menangkap JSON produk
            captured: list = []
            async def _handle_response(resp):
                if "item_list" in resp.url or "product" in resp.url:
                    try:
                        body = await resp.json()
                        captured.append(body)
                    except Exception:
                        pass
            page.on("response", _handle_response)

            products = await self._scrape_shop_page(page, category, captured)

            await browser.close()
            return products[:self.max_products]

    async def _scrape_shop_page(self, page: Page, category: str,
                                 captured: list) -> List[Product]:
        """Buka TikTok Shop dan ekstrak produk best-selling."""
        products: List[Product] = []

        try:
            await page.goto(self.BEST_SELLER_URL,
                            timeout=config.SCRAPE_TIMEOUT_MS,
                            wait_until="domcontentloaded")
            await page.wait_for_timeout(3000)   # tunggu JS load
        except PWTimeout:
            print("[WARN] Timeout navigasi TikTok Shop, coba fallback...")

        # ── Coba parse dari request yang ter-intercept ─────────────────────
        for payload in captured:
            items = self._extract_items_from_payload(payload)
            products.extend(items)
            if len(products) >= self.max_products:
                break

        # ── Fallback: scrape DOM ───────────────────────────────────────────
        if not products:
            products = await self._scrape_dom(page, category)

        # ── Fallback 2: scrape via requests (tidak butuh browser) ─────────
        if not products:
            products = self._scrape_via_requests(category)

        return products

    async def _scrape_dom(self, page: Page, category: str) -> List[Product]:
        """Scrape elemen DOM TikTok Shop."""
        products = []
        try:
            # Tunggu card produk muncul
            await page.wait_for_selector(
                "[data-e2e='shop-product-card'], .product-card, [class*='ProductCard']",
                timeout=10_000,
            )
            cards = await page.query_selector_all(
                "[data-e2e='shop-product-card'], .product-card"
            )
            for card in cards[:self.max_products]:
                try:
                    name_el  = await card.query_selector("[class*='title'], h3, h2")
                    price_el = await card.query_selector("[class*='price'], [class*='Price']")
                    img_el   = await card.query_selector("img")
                    link_el  = await card.query_selector("a")

                    name  = await name_el.inner_text()  if name_el  else "Produk TikTok"
                    price_raw = await price_el.inner_text() if price_el else "0"
                    img   = await img_el.get_attribute("src") if img_el else ""
                    href  = await link_el.get_attribute("href") if link_el else ""

                    price = self._parse_price(price_raw)
                    pid   = self._extract_product_id(href)

                    products.append(Product(
                        product_id  = pid or f"dom_{len(products)}",
                        name        = name.strip(),
                        price       = price,
                        currency    = "IDR",
                        rating      = 4.5,
                        sold_count  = 0,
                        image_urls  = [img] if img else [],
                        shop_name   = "",
                        product_url = f"https://www.tiktok.com{href}" if href.startswith("/") else href,
                        category    = category,
                    ))
                except Exception as e:
                    print(f"[WARN] Gagal parse card: {e}")
        except Exception as e:
            print(f"[WARN] Scrape DOM gagal: {e}")
        return products

    def _scrape_via_requests(self, category: str) -> List[Product]:
        """
        Fallback ringan: ambil data dari TikTok web API tanpa browser.
        Karena TikTok ketat soal bot, ini mungkin perlu cookie/token.
        Di sini diimplementasikan sebagai demo placeholder.
        """
        print("[INFO] Menggunakan data demo (API butuh autentikasi)...")
        return self._demo_products(category)

    # ── Helper: parse payload JSON dari network intercept ─────────────────────
    def _extract_items_from_payload(self, payload: dict) -> List[Product]:
        products = []
        items = (
            payload.get("itemList")
            or payload.get("data", {}).get("itemList")
            or payload.get("data", {}).get("products")
            or []
        )
        for item in items:
            try:
                pid   = str(item.get("id") or item.get("itemId") or item.get("productId", ""))
                name  = item.get("desc") or item.get("title") or item.get("name", "")
                stats = item.get("stats") or item.get("statistics") or {}

                raw_price = (
                    item.get("price")
                    or item.get("priceMin")
                    or item.get("originPrice")
                    or 0
                )
                price = float(raw_price) / 100 if isinstance(raw_price, int) and raw_price > 1000 else float(raw_price)

                images = []
                cover  = item.get("video", {}).get("cover") or item.get("thumbnail")
                if cover:
                    images.append(cover)
                for img in item.get("images", []):
                    url = img if isinstance(img, str) else img.get("url", "")
                    if url:
                        images.append(url)

                products.append(Product(
                    product_id  = pid,
                    name        = name,
                    price       = price,
                    currency    = "IDR",
                    rating      = float(stats.get("commentCount", 0)) or 4.5,
                    sold_count  = int(stats.get("playCount") or stats.get("soldCount", 0)),
                    image_urls  = images,
                    shop_name   = item.get("author", {}).get("nickname", ""),
                    product_url = f"https://www.tiktok.com/shop/product/{pid}",
                ))
            except Exception as e:
                print(f"[WARN] Parse item gagal: {e}")
        return products

    # ── Helpers ────────────────────────────────────────────────────────────────
    @staticmethod
    def _parse_price(raw: str) -> float:
        digits = re.sub(r"[^\d.,]", "", raw).replace(",", "")
        try:
            return float(digits)
        except ValueError:
            return 0.0

    @staticmethod
    def _extract_product_id(url: str) -> str:
        m = re.search(r"/product/(\w+)", url)
        return m.group(1) if m else ""

    @staticmethod
    def _demo_products(category: str) -> List[Product]:
        """Data demo saat scraping gagal / untuk testing offline."""
        return [
            Product(
                product_id  = "demo_001",
                name        = "Serum Wajah Vitamin C Anti-Aging 30ml",
                price       = 89_000,
                currency    = "IDR",
                rating      = 4.8,
                sold_count  = 15_420,
                image_urls  = [],
                shop_name   = "BeautyStore Official",
                product_url = "https://www.tiktok.com/shop/product/demo_001",
                category    = category or "beauty",
                description = (
                    "Serum Vitamin C terbaik untuk kulit cerah, anti kusam, "
                    "dan anti penuaan. Cocok untuk semua jenis kulit."
                ),
                commission_rate = 15.0,
            ),
            Product(
                product_id  = "demo_002",
                name        = "Masker Rambut Keratin Treatment 200ml",
                price       = 65_000,
                currency    = "IDR",
                rating      = 4.7,
                sold_count  = 8_930,
                image_urls  = [],
                shop_name   = "HairCare Pro",
                product_url = "https://www.tiktok.com/shop/product/demo_002",
                category    = category or "beauty",
                description = (
                    "Masker rambut dengan formula Keratin + Argan Oil. "
                    "Rambut halus, berkilau, dan anti kusut hanya dalam 10 menit!"
                ),
                commission_rate = 12.0,
            ),
            Product(
                product_id  = "demo_003",
                name        = "Skincare Set Paket Lengkap Acne Care",
                price       = 175_000,
                currency    = "IDR",
                rating      = 4.9,
                sold_count  = 22_100,
                image_urls  = [],
                shop_name   = "GlowSkin Official",
                product_url = "https://www.tiktok.com/shop/product/demo_003",
                category    = category or "skincare",
                description = (
                    "Paket skincare lengkap: cleanser + toner + moisturizer. "
                    "Khusus kulit berjerawat. Sudah BPOM & dermatologically tested."
                ),
                commission_rate = 18.0,
            ),
        ]
