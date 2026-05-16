"""
TikTok Content Posting API v2 + Keranjang Kuning (TikTok Shop Product Tag).

Flow upload:
  1. POST /v2/post/publish/video/init/   → dapat upload_url + publish_id
  2. PUT  <upload_url>                   → upload binary video (chunk)
  3. POST /v2/post/publish/video/complete/ → finalisasi + metadata + product tag

Referensi resmi:
  https://developers.tiktok.com/doc/content-posting-api-get-started

Untuk Keranjang Kuning:
  - Butuh scope: video.publish + product.list (TikTok Shop Affiliate API)
  - Product ID diambil dari TikTok Shop API atau dari scraper
  - Di-pass dalam field `product_links` saat publish

Jika belum ada access token → jalankan OAuth helper (--auth flag di main.py).
"""

import hashlib
import json
import math
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional
from urllib.parse import urlencode

import requests

import config


@dataclass
class UploadResult:
    success: bool
    publish_id: str = ""
    video_url: str = ""
    error: str = ""


class TikTokUploader:
    API_BASE    = "https://open.tiktokapis.com/v2"
    AUTH_URL    = "https://www.tiktok.com/v2/auth/authorize/"
    TOKEN_URL   = "https://open.tiktokapis.com/v2/oauth/token/"
    CHUNK_SIZE  = 10 * 1024 * 1024  # 10 MB per chunk

    def __init__(self):
        self.client_key    = config.TIKTOK_CLIENT_KEY
        self.client_secret = config.TIKTOK_CLIENT_SECRET
        self.access_token  = config.TIKTOK_ACCESS_TOKEN
        self._session      = requests.Session()
        self._session.headers.update({"Content-Type": "application/json; charset=UTF-8"})

    # ── Public API ─────────────────────────────────────────────────────────────
    def upload(self, video_path: str, caption: str,
               product_ids: List[str] = None,
               hashtags: List[str]    = None) -> UploadResult:
        """
        Upload video ke TikTok dan pasang product tag (keranjang kuning).
        """
        if not self.access_token:
            return UploadResult(False, error=(
                "Access token belum ada. Jalankan: python main.py --auth"
            ))

        video_path = Path(video_path)
        if not video_path.exists():
            return UploadResult(False, error=f"File tidak ditemukan: {video_path}")

        file_size = video_path.stat().st_size
        print(f"[UPLOAD] File: {video_path.name} ({file_size / 1e6:.1f} MB)")

        # 1. Init upload
        init_resp = self._init_upload(file_size)
        if not init_resp.get("data"):
            return UploadResult(False, error=str(init_resp))

        upload_url = init_resp["data"]["upload_url"]
        publish_id = init_resp["data"]["publish_id"]
        print(f"[UPLOAD] Publish ID: {publish_id}")

        # 2. Upload video (chunked)
        ok = self._upload_chunks(video_path, upload_url, file_size)
        if not ok:
            return UploadResult(False, error="Chunk upload gagal")

        # 3. Complete / publish
        result = self._complete_upload(
            publish_id = publish_id,
            caption    = caption,
            hashtags   = hashtags or [],
            product_ids= product_ids or [],
        )

        if result.get("data", {}).get("publish_id"):
            video_url = f"https://www.tiktok.com/@me/video/{publish_id}"
            print(f"[UPLOAD] Sukses! URL: {video_url}")
            return UploadResult(True, publish_id=publish_id, video_url=video_url)
        else:
            return UploadResult(False, error=str(result))

    # ── Step 1: Init ───────────────────────────────────────────────────────────
    def _init_upload(self, file_size: int) -> dict:
        chunk_count = math.ceil(file_size / self.CHUNK_SIZE)
        payload = {
            "post_info": {
                "privacy_level": "PUBLIC_TO_EVERYONE",
                "disable_duet": False,
                "disable_comment": False,
                "disable_stitch": False,
                "video_cover_timestamp_ms": 1000,
            },
            "source_info": {
                "source"          : "FILE_UPLOAD",
                "video_size"      : file_size,
                "chunk_size"      : self.CHUNK_SIZE,
                "total_chunk_count": chunk_count,
            },
        }
        return self._post("/post/publish/video/init/", payload)

    # ── Step 2: Chunked upload ─────────────────────────────────────────────────
    def _upload_chunks(self, video_path: Path, upload_url: str,
                       file_size: int) -> bool:
        chunk_count = math.ceil(file_size / self.CHUNK_SIZE)
        with open(video_path, "rb") as f:
            for i in range(chunk_count):
                start = i * self.CHUNK_SIZE
                end   = min(start + self.CHUNK_SIZE, file_size) - 1
                chunk = f.read(self.CHUNK_SIZE)

                headers = {
                    "Content-Range" : f"bytes {start}-{end}/{file_size}",
                    "Content-Length": str(len(chunk)),
                    "Content-Type"  : "video/mp4",
                }
                resp = requests.put(upload_url, data=chunk, headers=headers, timeout=120)
                print(f"[UPLOAD] Chunk {i+1}/{chunk_count} → HTTP {resp.status_code}")
                if resp.status_code not in (200, 206):
                    print(f"[ERROR] Chunk gagal: {resp.text[:300]}")
                    return False
        return True

    # ── Step 3: Complete + product tag ────────────────────────────────────────
    def _complete_upload(self, publish_id: str, caption: str,
                         hashtags: List[str], product_ids: List[str]) -> dict:
        # Gabungkan caption + hashtag
        full_caption = caption
        if hashtags:
            ht_str = " ".join(h if h.startswith("#") else f"#{h}" for h in hashtags)
            if ht_str not in full_caption:
                full_caption = f"{full_caption}\n\n{ht_str}"

        payload: dict = {
            "publish_id"   : publish_id,
            "post_info"    : {
                "title"        : full_caption[:2200],  # TikTok max 2200 karakter
                "privacy_level": "PUBLIC_TO_EVERYONE",
            },
        }

        # Tambahkan product link (Keranjang Kuning) jika ada
        if product_ids:
            payload["product_links"] = [
                {"item_id": pid} for pid in product_ids
            ]
            print(f"[UPLOAD] Menambahkan keranjang kuning: {product_ids}")

        return self._post("/post/publish/video/complete/", payload)

    # ── HTTP helpers ───────────────────────────────────────────────────────────
    def _post(self, endpoint: str, payload: dict) -> dict:
        url     = self.API_BASE + endpoint
        headers = {
            "Authorization" : f"Bearer {self.access_token}",
            "Content-Type"  : "application/json; charset=UTF-8",
        }
        resp = self._session.post(url, json=payload, headers=headers, timeout=60)
        try:
            return resp.json()
        except Exception:
            return {"error": resp.text}

    # ── OAuth helper ───────────────────────────────────────────────────────────
    def get_auth_url(self) -> str:
        """Hasilkan URL untuk OAuth login pengguna."""
        params = {
            "client_key"   : self.client_key,
            "scope"        : "user.info.basic,video.publish,video.upload",
            "response_type": "code",
            "redirect_uri" : config.TIKTOK_REDIRECT_URI,
            "state"        : "tiktok_auto",
        }
        return self.AUTH_URL + "?" + urlencode(params)

    def exchange_code_for_token(self, code: str) -> dict:
        """Tukar authorization code dengan access token."""
        payload = {
            "client_key"   : self.client_key,
            "client_secret": self.client_secret,
            "code"         : code,
            "grant_type"   : "authorization_code",
            "redirect_uri" : config.TIKTOK_REDIRECT_URI,
        }
        resp = requests.post(self.TOKEN_URL, data=payload, timeout=30)
        data = resp.json()
        if data.get("access_token"):
            print(f"[AUTH] Access token berhasil! Simpan ke .env:")
            print(f"       TIKTOK_ACCESS_TOKEN={data['access_token']}")
        return data

    # ── TikTok Shop: Cari product_id untuk keranjang kuning ───────────────────
    def get_shop_product_id(self, search_keyword: str) -> Optional[str]:
        """
        Cari produk di TikTok Shop Affiliate untuk mendapatkan product_id
        yang bisa ditag sebagai keranjang kuning.
        """
        url     = "https://open-api.affiliate.tiktok.com/seller/202309/products/search"
        headers = {
            "x-Use-Encryption": "0",
            "Content-Type"    : "application/json",
        }
        payload = {
            "app_key"     : config.TIKTOK_SHOP_APP_KEY,
            "access_token": config.TIKTOK_SHOP_ACCESS_TOKEN,
            "keyword"     : search_keyword,
            "page_size"   : 5,
            "page_number" : 1,
            "sort_type"   : 6,  # 6 = sort by sales
        }
        try:
            resp = requests.post(url, json=payload, headers=headers, timeout=15)
            data = resp.json()
            items = data.get("data", {}).get("products", [])
            if items:
                pid = items[0].get("product_id") or items[0].get("item_id", "")
                print(f"[SHOP] Product ID ditemukan: {pid} untuk '{search_keyword}'")
                return str(pid)
        except Exception as e:
            print(f"[WARN] Gagal cari product ID: {e}")
        return None
