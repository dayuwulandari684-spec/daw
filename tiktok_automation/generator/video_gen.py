"""
Generator video UGC untuk TikTok (1080x1920 vertical, 30fps).

Pipeline:
  1. Download gambar produk (Pillow)
  2. Buat slide: background gradient + gambar produk + teks overlay
  3. Compose slide → video clips (MoviePy)
  4. Gabungkan clips + voiceover + musik latar
  5. Export MP4
"""

import os
import io
import random
from pathlib import Path
from typing import List, Optional

import numpy as np
import requests
from PIL import Image, ImageDraw, ImageFont, ImageFilter

try:
    from moviepy.editor import (
        VideoFileClip, ImageClip, AudioFileClip,
        CompositeVideoClip, concatenate_videoclips,
        TextClip, ColorClip,
    )
    from moviepy.audio.fx.all import audio_fadein, audio_fadeout, volumex
    MOVIEPY_OK = True
except ImportError:
    MOVIEPY_OK = False
    print("[WARN] moviepy tidak tersedia, video generation dinonaktifkan.")

import config
from scraper import Product
from generator.script_gen import UGCScript


# ── Palet warna per kategori ───────────────────────────────────────────────────
PALETTES = {
    "beauty"  : [(255, 182, 193), (255, 105, 180)],  # pink
    "fashion" : [(147, 112, 219), (75,  0, 130)],    # purple
    "electronics": [(30, 144, 255), (0, 0, 139)],    # blue
    "food"    : [(255, 165, 0),  (255, 69,  0)],     # orange
    "default" : [(255, 99,  71), (220, 20,  60)],    # red
}

W, H = config.VIDEO_WIDTH, config.VIDEO_HEIGHT  # 1080 x 1920


class VideoGenerator:
    def __init__(self, output_dir: str = config.VIDEOS_DIR,
                 images_dir: str = config.IMAGES_DIR):
        self.output_dir = Path(output_dir)
        self.images_dir = Path(images_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.images_dir.mkdir(parents=True, exist_ok=True)

    def create(self, product: Product, script: UGCScript,
               voice_path: str, output_name: str = None) -> str:
        """
        Buat video UGC lengkap.
        Return: path ke file MP4.
        """
        if not MOVIEPY_OK:
            raise RuntimeError("moviepy belum diinstall. Jalankan: pip install moviepy")

        output_name = output_name or f"{product.product_id}_ugc"
        output_path = self.output_dir / f"{output_name}.mp4"

        print(f"[VIDEO] Membuat video untuk: {product.name}")

        # 1. Download / siapkan gambar produk
        images = self._prepare_images(product)

        # 2. Hitung durasi dari voiceover
        audio_clip   = AudioFileClip(voice_path)
        total_secs   = max(audio_clip.duration + 2, 25)  # minimal 25 detik

        # 3. Buat frame-frame slide (numpy arrays)
        slides = self._build_slides(product, script, images)

        # 4. Durasi per slide
        n_slides     = len(slides)
        slide_dur    = total_secs / n_slides

        # 5. Buat video clips dari slides
        clips = []
        for i, frame in enumerate(slides):
            clip = (ImageClip(frame)
                    .set_duration(slide_dur)
                    .fadein(config.FADE_DURATION)
                    .fadeout(config.FADE_DURATION))
            clips.append(clip)

        video = concatenate_videoclips(clips, method="compose")

        # 6. Pasang audio
        audio_clip   = audio_clip.set_start(0)
        if audio_clip.duration < video.duration:
            audio_clip = audio_clip.audio_loop(duration=video.duration)
        else:
            audio_clip = audio_clip.subclip(0, video.duration)

        final = video.set_audio(audio_clip)

        # 7. Export
        print(f"[VIDEO] Exporting → {output_path.name}")
        final.write_videofile(
            str(output_path),
            fps     = config.VIDEO_FPS,
            codec   = "libx264",
            audio_codec = "aac",
            threads = 4,
            preset  = "fast",
            logger  = None,
        )
        audio_clip.close()
        final.close()
        print(f"[VIDEO] Selesai: {output_path}")
        return str(output_path)

    # ── Slide builder ──────────────────────────────────────────────────────────
    def _build_slides(self, product: Product, script: UGCScript,
                      images: List[Image.Image]) -> List[np.ndarray]:
        cat     = (product.category or "default").lower()
        palette = PALETTES.get(cat, PALETTES["default"])

        slides = []

        # Slide 1 – HOOK
        slides.append(self._slide_text(
            text=script.hook,
            bg_colors=palette,
            emoji="🔥",
            font_size=70,
        ))

        # Slide 2 – PROBLEM
        slides.append(self._slide_text(
            text=script.problem,
            bg_colors=[(80, 80, 80), (30, 30, 30)],
            emoji="😩",
            font_size=60,
        ))

        # Slide 3..N-1 – PRODUK (gambar + nama + harga)
        if images:
            for img in images[:3]:
                slides.append(self._slide_product(product, img, palette))
        else:
            slides.append(self._slide_product_no_image(product, palette))

        # Slide terakhir – CTA
        slides.append(self._slide_cta(script.cta, product))

        return slides

    def _slide_text(self, text: str, bg_colors: list,
                    emoji: str = "", font_size: int = 65) -> np.ndarray:
        img = self._gradient_bg(bg_colors)
        draw = ImageDraw.Draw(img)
        font = self._get_font(font_size)

        # Word wrap
        lines  = self._wrap_text(text, font, W - 120)
        total_h = len(lines) * (font_size + 20)
        y = (H - total_h) // 2 - 60

        for line in lines:
            bbox = draw.textbbox((0, 0), line, font=font)
            tw   = bbox[2] - bbox[0]
            x    = (W - tw) // 2
            # shadow
            draw.text((x+3, y+3), line, fill=(0, 0, 0, 120), font=font)
            draw.text((x, y),     line, fill=(255, 255, 255), font=font)
            y += font_size + 20

        if emoji:
            ef   = self._get_font(120)
            ebbox = draw.textbbox((0, 0), emoji, font=ef)
            ew   = ebbox[2] - ebbox[0]
            draw.text(((W - ew) // 2, 120), emoji, font=ef, fill=(255,255,255))

        return np.array(img)

    def _slide_product(self, product: Product, img: Image.Image,
                       palette: list) -> np.ndarray:
        canvas = self._gradient_bg(palette)

        # Produk image (centre, atas)
        img_resized = self._fit_image(img, int(W * 0.85), int(H * 0.55))
        x = (W - img_resized.width) // 2
        y = 100
        canvas.paste(img_resized, (x, y), img_resized if img_resized.mode == "RGBA" else None)

        draw  = ImageDraw.Draw(canvas)
        big   = self._get_font(55)
        small = self._get_font(45)

        # Nama produk
        name_lines = self._wrap_text(product.name, big, W - 80)
        ny = y + img_resized.height + 40
        for line in name_lines[:2]:
            bbox = draw.textbbox((0,0), line, font=big)
            tw   = bbox[2] - bbox[0]
            draw.text(((W-tw)//2, ny), line, fill="white", font=big)
            ny += 65

        # Harga + rating
        price_str = f"Rp{product.price:,.0f}"
        draw.text((80, ny + 10), price_str, fill="#FFD700", font=big)

        star_str = f"⭐ {product.rating}  |  {product.sold_count:,}+ terjual"
        draw.text((80, ny + 80), star_str, fill="white", font=small)

        return np.array(canvas)

    def _slide_product_no_image(self, product: Product, palette: list) -> np.ndarray:
        return self._slide_text(
            text=f"{product.name}\n\nRp{product.price:,.0f}\n⭐{product.rating}",
            bg_colors=palette,
            emoji="🛍️",
        )

    def _slide_cta(self, cta_text: str, product: Product) -> np.ndarray:
        img  = self._gradient_bg([(255, 69, 0), (220, 20, 60)])
        draw = ImageDraw.Draw(img)
        big  = self._get_font(75)
        med  = self._get_font(55)

        # Keranjang kuning icon (teks)
        draw.text((W//2 - 80, 200), "🛒", font=self._get_font(200), fill="#FFD700")

        # CTA text
        lines = self._wrap_text(cta_text, big, W - 100)
        y     = 560
        for line in lines:
            bbox = draw.textbbox((0,0), line, font=big)
            tw   = bbox[2] - bbox[0]
            draw.text(((W-tw)//2 + 3, y+3), line, fill=(0,0,0,100), font=big)
            draw.text(((W-tw)//2,     y),   line, fill="white",       font=big)
            y += 90

        # Tombol kuning
        btn_rect = [160, y + 60, W - 160, y + 180]
        draw.rounded_rectangle(btn_rect, radius=40, fill="#FFD700")
        btn_text = "TAP KERANJANG KUNING ⬇"
        bbox = draw.textbbox((0,0), btn_text, font=med)
        bw   = bbox[2] - bbox[0]
        draw.text(((W - bw)//2, y + 95), btn_text, fill="#1A1A1A", font=med)

        return np.array(img)

    # ── Image helpers ──────────────────────────────────────────────────────────
    def _prepare_images(self, product: Product) -> List[Image.Image]:
        images = []
        for i, url in enumerate(product.image_urls[:3]):
            try:
                resp = requests.get(url, timeout=10, headers={
                    "User-Agent": "Mozilla/5.0"
                })
                img = Image.open(io.BytesIO(resp.content)).convert("RGBA")
                images.append(img)
                print(f"[IMAGE] Downloaded: {url[:60]}...")
            except Exception as e:
                print(f"[WARN] Gagal download gambar {i+1}: {e}")
        return images

    @staticmethod
    def _fit_image(img: Image.Image, max_w: int, max_h: int) -> Image.Image:
        img.thumbnail((max_w, max_h), Image.LANCZOS)
        return img

    @staticmethod
    def _gradient_bg(colors: list) -> Image.Image:
        """Buat latar belakang gradient vertikal."""
        img  = Image.new("RGB", (W, H))
        data = np.zeros((H, W, 3), dtype=np.uint8)
        c1   = np.array(colors[0])
        c2   = np.array(colors[-1])
        for row in range(H):
            t = row / H
            data[row] = (c1 * (1 - t) + c2 * t).astype(np.uint8)
        return Image.fromarray(data)

    @staticmethod
    def _get_font(size: int):
        """Cari font yang tersedia di sistem."""
        font_paths = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
            "C:/Windows/Fonts/ariblk.ttf",   # Windows
            "C:/Windows/Fonts/arialbd.ttf",
        ]
        for fp in font_paths:
            if os.path.exists(fp):
                try:
                    return ImageFont.truetype(fp, size)
                except Exception:
                    continue
        return ImageFont.load_default()

    @staticmethod
    def _wrap_text(text: str, font, max_width: int) -> List[str]:
        """Bagi teks menjadi baris agar muat dalam max_width."""
        words  = text.split()
        lines  = []
        cur    = ""
        dummy_img  = Image.new("RGB", (1, 1))
        dummy_draw = ImageDraw.Draw(dummy_img)
        for word in words:
            test = f"{cur} {word}".strip()
            bbox = dummy_draw.textbbox((0, 0), test, font=font)
            if bbox[2] - bbox[0] <= max_width:
                cur = test
            else:
                if cur:
                    lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)
        return lines
