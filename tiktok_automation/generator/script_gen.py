"""
Generator script UGC untuk TikTok.

Struktur konten (AIDA framework):
  Hook (0-3 dtk)  → Attention: kalimat pembuka mengejutkan / relatable
  Problem (3-8 dtk) → Interest: masalah yang dialami target audiens
  Solution (8-20 dtk) → Desire: produk sebagai solusi + demo visual
  CTA (20-30 dtk) → Action: "Tap keranjang kuning di bawah!"

Jika OPENAI_API_KEY diset → gunakan GPT-4o.
Jika tidak → gunakan template bawaan.
"""

import random
from dataclasses import dataclass
from typing import List

import config
from scraper import Product


@dataclass
class UGCScript:
    hook: str
    problem: str
    solution: str
    cta: str
    hashtags: List[str]
    caption: str

    def full_voiceover(self) -> str:
        """Gabungkan semua bagian menjadi teks voiceover penuh."""
        return " ".join([self.hook, self.problem, self.solution, self.cta])


class ScriptGenerator:
    """Buat script UGC dari data produk."""

    # ── Template hook berdasarkan kategori ────────────────────────────────────
    _HOOKS = {
        "beauty": [
            "Kulit glowing cuma modal berapa? Ini dia rahasianya!",
            "Udah coba skincare mahal tapi kulit tetep kusam? Dengerin dulu ini!",
            "Ini bukan iklan, ini pengakuan jujur dari aku!",
            "Teman-teman, aku nemu sesuatu yang wajib kalian tau!",
            "STOP beli skincare mahal dulu! Coba ini dulu yang harganya bikin kaget.",
        ],
        "fashion": [
            "Outfit kece nggak harus mahal, ini buktinya!",
            "Tampil stylish under 100 ribu? Bisa banget!",
            "Koleksi terbaru yang lagi viral di TikTok, jangan sampai kehabisan!",
        ],
        "electronics": [
            "Gadget murah tapi kualitas premium? Ada nih!",
            "Ini alasan kenapa semua orang lagi rebutan beli ini!",
        ],
        "default": [
            "Nggak nyangka produk ini bisa sebagus ini!",
            "Honest review dari aku yang udah pakai selama seminggu!",
            "Ini dia yang lagi viral dan terjual ribuan dalam sehari!",
            "Kalau kamu sering {masalah}, ini solusinya!",
        ],
    }

    _PROBLEMS = {
        "beauty": [
            "Masalah kulit kusam, jerawat, dan noda hitam udah bikin nggak pede banget kan?",
            "Udah coba banyak produk tapi hasilnya mengecewakan?",
            "Skincare mahal belum tentu cocok, dan itu yang aku rasain juga.",
        ],
        "default": [
            "Aku tau persis masalah yang kamu hadapi setiap hari.",
            "Dan itu yang bikin aku nyoba produk ini pertama kali.",
        ],
    }

    _SOLUTIONS = {
        "beauty": [
            (
                "Tapi setelah pakai {name}, dalam 7 hari hasilnya udah keliatan! "
                "Kulit lebih cerah, lembut, dan glowing alami. "
                "Formulanya ringan, cepat meresap, dan nggak bikin bruntusan."
            ),
            (
                "{name} ini udah terbukti! Rating {rating} bintang dengan lebih dari "
                "{sold_count} yang udah beli. Bahan-bahannya aman, sudah teruji dermatologis."
            ),
        ],
        "default": [
            (
                "{name} ini beneran game changer! "
                "Harganya cuma {price} tapi kualitasnya beyond ekspektasi. "
                "Udah {sold_count} orang lebih yang beli dan puas!"
            ),
        ],
    }

    _CTAS = [
        "Tap keranjang kuning di bawah sekarang sebelum kehabisan!",
        "Klik keranjang kuning ya, aku kasih link langsung ke produknya!",
        "Jangan lupa tap keranjang di bawah video ini, link produknya ada di sana!",
        "Buruan, stok terbatas! Tap keranjang kuning sekarang!",
        "Link produknya ada di keranjang kuning di bawah, langsung order ya!",
    ]

    _HASHTAGS_BASE = [
        "#TikTokShop", "#FYP", "#Viral", "#Rekomendasi",
        "#ProdukViral", "#TikTokMadeMeBuyIt",
    ]

    _HASHTAGS_CATEGORY = {
        "beauty"     : ["#Skincare", "#BeautyTips", "#KulitGlowing", "#Perawatan"],
        "fashion"    : ["#OOTD", "#Fashion", "#StyleTips", "#OutfitInspiration"],
        "electronics": ["#Gadget", "#TechTips", "#Review"],
        "food"       : ["#Kuliner", "#FoodReview", "#Makanan"],
    }

    def generate(self, product: Product) -> UGCScript:
        """Hasilkan script UGC. Gunakan OpenAI jika tersedia."""
        if config.OPENAI_API_KEY:
            try:
                return self._generate_with_ai(product)
            except Exception as e:
                print(f"[WARN] OpenAI gagal ({e}), pakai template...")
        return self._generate_from_template(product)

    # ── Template-based ─────────────────────────────────────────────────────────
    def _generate_from_template(self, p: Product) -> UGCScript:
        cat = p.category.lower() if p.category else "default"
        hooks    = self._HOOKS.get(cat, self._HOOKS["default"])
        problems = self._PROBLEMS.get(cat, self._PROBLEMS["default"])
        solutions = self._SOLUTIONS.get(cat, self._SOLUTIONS["default"])

        hook     = random.choice(hooks)
        problem  = random.choice(problems)
        solution = random.choice(solutions).format(
            name       = p.name,
            rating     = p.rating,
            sold_count = f"{p.sold_count:,}",
            price      = f"Rp{p.price:,.0f}",
        )
        cta = random.choice(self._CTAS)

        hashtags = (
            self._HASHTAGS_BASE
            + self._HASHTAGS_CATEGORY.get(cat, [])
            + [f"#{word.capitalize()}" for word in p.name.split()[:2]]
        )

        caption = (
            f"{hook}\n\n"
            f"✨ {p.name}\n"
            f"💰 Harga: Rp{p.price:,.0f}\n"
            f"⭐ Rating: {p.rating}/5\n"
            f"🛒 Sudah {p.sold_count:,}+ terjual!\n\n"
            f"Tap keranjang kuning di bawah!\n\n"
            + " ".join(hashtags)
        )

        return UGCScript(
            hook=hook,
            problem=problem,
            solution=solution,
            cta=cta,
            hashtags=hashtags,
            caption=caption,
        )

    # ── AI-based (OpenAI GPT-4o) ───────────────────────────────────────────────
    def _generate_with_ai(self, p: Product) -> UGCScript:
        from openai import OpenAI
        client = OpenAI(api_key=config.OPENAI_API_KEY)

        prompt = f"""
Kamu adalah kreator konten TikTok Indonesia yang ahli membuat video UGC viral.
Buat script video TikTok 30 detik untuk produk berikut:

Nama produk : {p.name}
Harga       : Rp{p.price:,.0f}
Rating      : {p.rating}/5
Terjual     : {p.sold_count:,} pcs
Deskripsi   : {p.description}
Kategori    : {p.category}

Format WAJIB (JSON):
{{
  "hook"    : "kalimat pembuka 1-2 kalimat, menarik perhatian dalam 3 detik",
  "problem" : "masalah relatable target audiens, 1-2 kalimat",
  "solution": "produk sebagai solusi + manfaat utama, 3-4 kalimat",
  "cta"     : "ajakan tap keranjang kuning, 1 kalimat",
  "hashtags": ["#tag1", "#tag2", ...],
  "caption" : "caption lengkap dengan emoji dan hashtag"
}}

Gunakan Bahasa Indonesia yang casual, relatable, dan antusias.
Sebutkan keranjang kuning di CTA.
"""
        resp = client.chat.completions.create(
            model    = "gpt-4o-mini",
            messages = [{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
        )
        import json
        data = json.loads(resp.choices[0].message.content)
        return UGCScript(
            hook     = data["hook"],
            problem  = data["problem"],
            solution = data["solution"],
            cta      = data["cta"],
            hashtags = data["hashtags"],
            caption  = data["caption"],
        )
