# TikTok UGC Automation

## Mekanisme Kerja

```
┌─────────────────────────────────────────────────────────────────┐
│                     PIPELINE OTOMASI                            │
│                                                                 │
│  ┌──────────┐    ┌──────────┐    ┌──────────┐    ┌──────────┐  │
│  │  SCRAPER │───▶│GENERATOR │───▶│  VIDEO   │───▶│ UPLOADER │  │
│  │          │    │          │    │ CREATOR  │    │          │  │
│  │ TikTok   │    │ Script   │    │ MoviePy  │    │ TikTok   │  │
│  │ Shop     │    │ UGC(AI)  │    │ + PIL    │    │ API v2   │  │
│  │ Best     │    │ + TTS    │    │ + gTTS   │    │ + 🛒     │  │
│  │ Selling  │    │ Voiceover│    │ 1080x1920│    │ Keranjang│  │
│  └──────────┘    └──────────┘    └──────────┘    └──────────┘  │
│       │                │                │               │       │
│   Produk Info      Script UGC        Video MP4      Post URL   │
│   Nama, Harga      Hook+Problem      30 detik       + Cart Tag  │
│   Gambar, ID       Solution+CTA      Vertical                   │
└─────────────────────────────────────────────────────────────────┘
```

## Struktur Folder

```
tiktok_automation/
├── main.py              ← Entry point utama
├── config.py            ← Semua konfigurasi
├── requirements.txt     ← Dependencies Python
├── .env                 ← API keys (JANGAN dicommit ke git!)
├── .env.example         ← Template .env
│
├── scraper/
│   └── product_scraper.py   ← Scraping TikTok Shop
│
├── generator/
│   ├── script_gen.py    ← Buat script UGC (template / OpenAI)
│   ├── voice_gen.py     ← Text-to-Speech (gTTS, gratis)
│   └── video_gen.py     ← Buat video 1080×1920 (MoviePy)
│
├── uploader/
│   └── tiktok_uploader.py  ← Upload + Keranjang Kuning
│
└── output/
    ├── videos/          ← File MP4 hasil generate
    ├── audio/           ← File MP3 voiceover
    ├── images/          ← Gambar produk yang didownload
    └── state.json       ← Progress & history
```

## Setup Awal (Jalankan Sekali)

### 1. Install Python & Dependencies

```cmd
# Pastikan Python 3.10+ sudah terinstall
python --version

# Masuk ke folder project
cd tiktok_automation

# Setup otomatis (install semua package)
python main.py --setup
```

### 2. Daftar Akun Developer TikTok

1. Buka https://developers.tiktok.com
2. Buat akun / login dengan akun TikTok kamu
3. Klik **"Manage Apps"** → **"Create App"**
4. Isi nama app, deskripsi
5. Di bagian **Products**, aktifkan:
   - ✅ **Content Posting API** (untuk upload video)
   - ✅ **Login Kit** (untuk OAuth)
6. Set **Redirect URI**: `http://localhost:8080/callback`
7. Copy **Client Key** dan **Client Secret**

### 3. Setup TikTok Shop Affiliate (untuk Keranjang Kuning)

1. Buka https://affiliate.tiktok.com
2. Daftar sebagai **Creator Affiliate**
3. Di **Developer Settings** → buat app
4. Aktifkan scope: `product.list`
5. Copy **App Key** dan **App Secret**

### 4. Isi File .env

```cmd
# Copy template
copy .env.example .env

# Edit dengan Notepad
notepad .env
```

Isi values berikut:
```
TIKTOK_CLIENT_KEY=your_client_key
TIKTOK_CLIENT_SECRET=your_client_secret
TIKTOK_SHOP_APP_KEY=your_shop_app_key
TIKTOK_SHOP_APP_SECRET=your_shop_app_secret

# Opsional - untuk script lebih bagus:
OPENAI_API_KEY=sk-...
```

### 5. Login ke TikTok (OAuth)

```cmd
python main.py --auth
```

→ Browser akan terbuka → Login TikTok → Izinkan akses
→ Access token otomatis disimpan ke .env

---

## Cara Pakai

### Jalankan Pipeline Lengkap

```cmd
# Semua kategori, ambil 5 produk
python main.py --run

# Khusus kategori beauty, ambil 3 produk
python main.py --run --category beauty --limit 3

# Kategori lain: fashion, electronics, food
python main.py --run --category fashion
```

### Perintah Individual

```cmd
# Hanya scraping (lihat produk dulu sebelum buat video)
python main.py --scrape --category beauty

# Hanya generate video (dari produk yang sudah di-scrape)
python main.py --generate

# Lihat status / progress
python main.py --list
```

---

## Mekanisme Keranjang Kuning

**Keranjang Kuning** = fitur TikTok Shop di mana produk bisa ditag di video,
muncul sebagai ikon keranjang kuning yang bisa diklik penonton untuk langsung beli.

### Cara kerjanya dalam script ini:

```
1. Scraper ambil nama produk dari TikTok Shop
      ↓
2. Uploader cari product_id di TikTok Shop Affiliate API
   (berdasarkan nama produk yang paling mirip)
      ↓
3. Saat upload video, product_id dikirim di field "product_links"
      ↓
4. TikTok otomatis tampilkan keranjang kuning di video
```

### Syarat keranjang kuning aktif:
- ✅ Akun TikTok sudah join TikTok Shop Affiliate
- ✅ Produk tersedia di TikTok Shop
- ✅ App sudah punya scope `product.list`
- ✅ TIKTOK_SHOP_ACCESS_TOKEN sudah diisi

---

## Struktur Video yang Dibuat

```
[0-3 dtk]   HOOK      → Kalimat pembuka menarik + emoji
[3-8 dtk]   PROBLEM   → Masalah relatable target audiens  
[8-20 dtk]  PRODUK    → Gambar produk + nama + harga + rating
[20-30 dtk] CTA       → "Tap keranjang kuning di bawah!" 🛒
```

Spesifikasi:
- Resolusi : 1080 × 1920 (TikTok vertical)
- FPS      : 30
- Codec    : H.264 + AAC
- Durasi   : 25-35 detik
- Audio    : Google TTS (Bahasa Indonesia)

---

## Troubleshooting

| Masalah | Solusi |
|---------|--------|
| `playwright install` gagal | Coba: `python -m playwright install chromium` |
| Scraping dapat 0 produk | Normal jika TikTok blokir bot; akan pakai data demo |
| Upload gagal: "token expired" | Jalankan ulang `python main.py --auth` |
| Video hitam/blank | Pastikan FFmpeg terinstall: `pip install imageio-ffmpeg` |
| Keranjang kuning tidak muncul | Cek TIKTOK_SHOP_ACCESS_TOKEN di .env |
| gTTS error | Cek koneksi internet (butuh akses Google TTS) |

---

## Catatan Penting

> ⚠️ Scraping TikTok Shop melanggar ToS jika berlebihan.
> Gunakan secara wajar: maksimal 20-30 request/hari.
> 
> ⚠️ Access token TikTok berlaku 24 jam.
> Jalankan `--auth` kembali jika expired.
>
> ✅ Upload via Content Posting API adalah cara RESMI dan aman.
