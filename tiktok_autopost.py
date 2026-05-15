"""
TikTok AutoPost — versi HEADLESS untuk Task Scheduler / cron.

Alur:
  1. Cari video YouTube berdasarkan KEYWORD
  2. Download video (via yt-dlp)
  3. Proses dengan FFmpeg: clip 60 detik + tambah logo + burn subtitle
  4. Upload otomatis ke TikTok Studio (Playwright headless)
  5. Semua log tersimpan ke auto_run.log

Dependensi:
  pip install yt-dlp playwright
  playwright install chromium
  ffmpeg harus tersedia di PATH
"""

import os
import re
import sys
import time
import random
import builtins
import urllib.request
import subprocess
from datetime import datetime

import yt_dlp

# Pindah ke direktori script agar path relatif berfungsi
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)

# ============================================================
# KONFIGURASI — sesuaikan di sini
# ============================================================

KEYWORD       = "KLA Project live performance"   # kata kunci pencarian YouTube
MAX_VIDEO     = 1                                 # berapa video per run
SEARCH_COUNT  = 25                                # pool pencarian YouTube
SHUFFLE       = True                              # acak urutan hasil

CLIP_DURATION = 60                                # durasi clip (detik)
CLIP_POSITION = 0.40                              # mulai clip dari posisi ini (0-1)
MAX_DURATION  = 1800                              # lewati video > 30 menit
MAX_HEIGHT    = 720                               # resolusi maks download

OUTPUT_FOLDER = "output"
COOKIE_FILE   = "cookies.txt"                     # cookies YouTube (opsional)
LOGO_URL      = "https://klacorporation.com/en/images/LogoFront.png"
LOGO_FILE     = "logo_kla.png"

TIKTOK_COOKIE_FILE = "tiktok_cookies.txt"
ARTIST_NAME        = "KLa Project"
CAPTION_TEMPLATE   = "{song} - {artist} #KLAProject #KLaCorp #KLanese"
AUTO_UPLOAD_TIKTOK = True

LOG_FILE = "auto_run.log"

# ============================================================
# LOGGING — semua print() otomatis ke terminal + file
# ============================================================

_original_print = builtins.print

def _logged_print(*args, **kwargs):
    msg = " ".join(str(a) for a in args)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    _original_print(line, **kwargs)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

builtins.print = _logged_print

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# ============================================================
# DOWNLOAD LOGO
# ============================================================

def download_logo():
    if os.path.exists(LOGO_FILE):
        return True
    try:
        req = urllib.request.Request(LOGO_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req) as resp, open(LOGO_FILE, "wb") as f:
            f.write(resp.read())
        print(f"Logo disimpan: {LOGO_FILE}")
        return True
    except Exception as e:
        print(f"Gagal download logo: {e}")
        return False

# ============================================================
# COOKIES
# ============================================================

def get_cookie_config():
    if os.path.exists(COOKIE_FILE):
        return {"cookiefile": COOKIE_FILE}
    return {}

# ============================================================
# FFMPEG — clip + logo + subtitle
# ============================================================

def get_video_duration(path):
    try:
        r = subprocess.run(
            ["ffprobe", "-v", "error", "-show_entries", "format=duration",
             "-of", "default=noprint_wrappers=1:nokey=1", path],
            capture_output=True, text=True, check=True,
        )
        return float(r.stdout.strip() or 0)
    except Exception:
        return 0


def add_logo_and_subtitles(video_path):
    if not os.path.exists(LOGO_FILE):
        print("  Logo tidak ditemukan, skip overlay.")
        return None

    base = os.path.splitext(video_path)[0]
    duration = get_video_duration(video_path)

    if duration > CLIP_DURATION:
        clip_start = duration * CLIP_POSITION
        clip_dur   = min(CLIP_DURATION, duration - clip_start)
        print(f"  Clip: mulai {clip_start:.0f}s, durasi {clip_dur:.0f}s (total {duration:.0f}s)")
    else:
        clip_start = 0
        clip_dur   = duration
        print(f"  Video pendek ({duration:.0f}s), tidak di-clip.")

    # Cari subtitle dengan lirik nyata (bukan sekadar tag [Musik])
    subtitle_path = None
    for ext in [".id.vtt", ".en.vtt", ".id.srt", ".en.srt", ".vtt", ".srt"]:
        candidate = base + ext
        if not os.path.exists(candidate):
            continue
        try:
            with open(candidate, "r", encoding="utf-8") as f:
                content = f.read()
            music_tags = content.lower().count("[musik]") + content.lower().count("[music]")
            text_lines = [
                l for l in content.split("\n")
                if l.strip()
                and "-->" not in l
                and not l.startswith("WEBVTT")
                and not l.strip().isdigit()
                and not l.startswith(("Kind:", "Language:"))
            ]
            if text_lines and music_tags <= len(text_lines) * 0.3:
                subtitle_path = candidate
                print(f"  Subtitle layak: {os.path.basename(candidate)}")
                break
            else:
                print(f"  Subtitle isinya mostly [Musik] tag, skip.")
        except Exception:
            subtitle_path = candidate
            break

    output_path = base + "_final.mp4"

    canvas_filter = (
        "[0:v]scale=-2:1920,crop=1080:1920,boxblur=20:3,setsar=1[bg];"
        "[0:v]scale=1080:-2,setsar=1[fg];"
        "[bg][fg]overlay=(W-w)/2:(H-h)/2[canvas];"
        "[1:v]scale=200:-1[logo];"
        "[canvas][logo]overlay=W-w-30:40[withlogo]"
    )

    if subtitle_path:
        sub_esc = subtitle_path.replace("\\", "/").replace(":", "\\:")
        final_chain = (
            f"[withlogo]subtitles='{sub_esc}':force_style="
            f"'FontSize=20,PrimaryColour=&H00FFFFFF,"
            f"OutlineColour=&H00000000,BorderStyle=1,Outline=2,MarginV=120'"
            f",scale=1080:1920,setsar=1[outv]"
        )
    else:
        final_chain = "[withlogo]scale=1080:1920,setsar=1[outv]"

    vf = canvas_filter + ";" + final_chain

    cmd = [
        "ffmpeg", "-y", "-loglevel", "error",
        "-ss", f"{clip_start:.2f}",
        "-t",  f"{clip_dur:.2f}",
        "-i", video_path,
        "-i", LOGO_FILE,
        "-filter_complex", vf,
        "-map", "[outv]",
        "-map", "0:a?",
        "-c:v", "libx264", "-preset", "veryfast", "-crf", "23",
        "-c:a", "aac", "-b:a", "128k",
        "-pix_fmt", "yuv420p",
        "-movflags", "+faststart",
        output_path,
    ]

    print(f"  Render ffmpeg → {os.path.basename(output_path)}")
    try:
        subprocess.run(cmd, check=True, capture_output=True)
        print(f"  Output: {output_path}")
        return output_path
    except subprocess.CalledProcessError as e:
        print(f"  ffmpeg gagal (exit code {e.returncode})")
        return None

# ============================================================
# EXTRACT SONG TITLE dari judul YouTube
# ============================================================

def extract_song_title(raw_title):
    t = raw_title
    t = re.sub(r"_final.*$", "", t, flags=re.IGNORECASE)
    t = re.sub(r"\.(mp4|webm|mkv|m4a)$", "", t, flags=re.IGNORECASE)
    for bracket in [r"\([^)]*\)", r"\[[^\]]*\]", r"\{[^}]*\}"]:
        t = re.sub(bracket, "", t)
    noise = [
        "GrandKLakustik", "KLAKUSTIK", "KLakustik",
        "LIVE", "Live", "live", "Concert", "CONCERT", "concert",
        "Passion, Love & Culture", "Passion Love & Culture",
        "Remastered", "FULL VERSION", "Full Version",
        "Official Music Video", "Official Video", "Official",
        "Music Video", "MV", "HD", "HQ", "4K", "1080p",
        "PLC", "Show", "lyrics", "Lyrics",
    ]
    for n in noise:
        t = re.sub(re.escape(n), "", t, flags=re.IGNORECASE)
    t = t.replace("|", " ")
    parts = [p.strip(" -_") for p in t.split("-") if p.strip(" -_")]
    artist_lower = ARTIST_NAME.lower()
    song = next(
        (p for p in parts if artist_lower not in p.lower() and "kla" not in p.lower()),
        parts[0] if parts else raw_title,
    )
    return re.sub(r"\s+", " ", song).strip()

# ============================================================
# BACA COOKIES NETSCAPE FORMAT
# ============================================================

def _load_netscape_cookies(path):
    cookies = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split("\t")
            if len(parts) < 7:
                continue
            domain, _flag, cpath, secure, expires, name, value = parts[:7]
            try:
                expires_i = int(expires)
            except ValueError:
                expires_i = -1
            cookies.append({
                "name": name, "value": value,
                "domain": domain, "path": cpath,
                "secure": secure.upper() == "TRUE",
                "expires": expires_i if expires_i > 0 else -1,
            })
    return cookies

# ============================================================
# UPLOAD KE TIKTOK (Playwright headless)
# ============================================================

def upload_to_tiktok(video_path, title):
    if not AUTO_UPLOAD_TIKTOK:
        return False
    if not video_path or not os.path.exists(video_path):
        print("  Video tidak ada, skip upload.")
        return False
    if not os.path.exists(TIKTOK_COOKIE_FILE):
        print(f"  {TIKTOK_COOKIE_FILE} tidak ditemukan, skip upload.")
        return False

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("  playwright belum terinstall. Jalankan: pip install playwright && playwright install chromium")
        return False

    song    = extract_song_title(title)[:100]
    caption = CAPTION_TEMPLATE.format(song=song, artist=ARTIST_NAME)
    video_abs = os.path.abspath(video_path)
    print(f"  Song: '{song}' | Caption: {caption}")
    print(f"  Upload TikTok (headless): {os.path.basename(video_path)}")

    try:
        cookies = _load_netscape_cookies(TIKTOK_COOKIE_FILE)

        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
            )
            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1920, "height": 1080},
            )
            context.add_cookies(cookies)
            page = context.new_page()

            print("  Buka TikTok Studio upload...")
            page.goto("https://www.tiktok.com/tiktokstudio/upload", wait_until="domcontentloaded")
            page.wait_for_timeout(5000)

            def dismiss_overlays():
                for sel in ['div[data-test-id="overlay"]', "div.react-joyride__overlay"]:
                    try:
                        ov = page.locator(sel).first
                        if ov.count() > 0 and ov.is_visible(timeout=1000):
                            page.mouse.click(540, 300)
                            page.wait_for_timeout(1000)
                    except Exception:
                        pass
                for sel in [
                    'button:has-text("Mengerti")', 'button:has-text("Got it")',
                    'button:has-text("Skip")',     'button:has-text("Lewati")',
                    'button:has-text("OK")',        'button[aria-label="Close"]',
                ]:
                    try:
                        btn = page.locator(sel).first
                        if btn.count() > 0 and btn.is_visible(timeout=1000):
                            btn.click(timeout=2000)
                            page.wait_for_timeout(800)
                    except Exception:
                        pass

            dismiss_overlays()

            print("  Set file video...")
            page.locator('input[type="file"]').first.set_input_files(video_abs)

            print("  Proses upload (~20 detik)...")
            page.wait_for_timeout(20000)
            dismiss_overlays()

            print("  Isi caption...")
            caption_filled = False
            for attempt in range(3):
                try:
                    dismiss_overlays()
                    box = page.locator('div[contenteditable="true"]').first
                    box.click(force=True, timeout=5000)
                    page.wait_for_timeout(500)
                    page.keyboard.press("Control+A")
                    page.keyboard.press("Delete")
                    page.keyboard.type(caption, delay=15)
                    caption_filled = True
                    print(f"  Caption OK (attempt {attempt+1})")
                    break
                except Exception as e:
                    print(f"  Caption attempt {attempt+1} gagal: {e}")
                    page.wait_for_timeout(2000)
            if not caption_filled:
                print("  Caption gagal 3x, lanjut upload.")

            for sel in ['label:has-text("Semua orang")', 'label:has-text("Public")']:
                try:
                    loc = page.locator(sel).first
                    if loc.count() > 0:
                        loc.click(timeout=2000)
                        break
                except Exception:
                    pass

            print("  Tunggu server upload (20s)...")
            page.wait_for_timeout(20000)

            posted = False
            for sel in [
                'button[data-e2e="post_video_button"]',
                'div[data-e2e="post_video_button"]',
                'button:has-text("Posting")',
                'button:has-text("Post")',
                "button.TUXButton--primary",
            ]:
                try:
                    btn = page.locator(sel).first
                    if btn.count() == 0:
                        continue
                    btn.wait_for(state="visible", timeout=3000)
                    for _ in range(30):
                        try:
                            if btn.is_enabled():
                                break
                        except Exception:
                            pass
                        page.wait_for_timeout(2000)
                    btn.click(force=True, timeout=5000)
                    posted = True
                    print(f"  Tombol Post diklik: {sel}")
                    break
                except Exception:
                    continue

            if posted:
                page.wait_for_timeout(5000)
                success = False
                for _ in range(15):
                    try:
                        url     = page.url
                        content = page.content().lower()
                        if "upload/success" in url or "manage" in url or "berhasil" in content:
                            success = True
                            break
                    except Exception:
                        pass
                    page.wait_for_timeout(2000)
                print("  UPLOAD SUKSES (verified)" if success else
                      "  WARNING: diklik tapi belum verified — kemungkinan tetap berhasil")
            else:
                print("  GAGAL: Tombol Post tidak ditemukan")

            browser.close()
            return posted

    except Exception as e:
        print(f"  TikTok upload error: {e}")
        return False

# ============================================================
# DOWNLOAD VIDEO
# ============================================================

_downloaded = []

def _progress_hook(d):
    if d["status"] == "finished":
        _downloaded.append(d["filename"])


def download_video(url, index):
    global _downloaded
    _downloaded = []
    print(f"[{index}] Download: {url}")

    ydl_opts = {
        "format"             : f"bestvideo[height<={MAX_HEIGHT}]+bestaudio/best[height<={MAX_HEIGHT}]/best",
        "outtmpl"            : f"{OUTPUT_FOLDER}/%(title)s.%(ext)s",
        "merge_output_format": "mp4",
        "match_filter"       : yt_dlp.utils.match_filter_func(f"duration < {MAX_DURATION}"),
        "writeautomaticsub"  : False,
        "writesubtitles"     : True,
        "subtitleslangs"     : ["id", "en"],
        "subtitlesformat"    : "vtt",
        "ignoreerrors"       : True,
        "no_warnings"        : True,
        "sleep_interval_subtitles": 5,
        "extractor_retries"  : 3,
        "sleep_interval"     : 1,
        "max_sleep_interval" : 3,
        "noplaylist"         : True,
        "quiet"              : True,
        "progress_hooks"     : [_progress_hook],
    }
    ydl_opts.update(get_cookie_config())

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)

        if not info:
            print("  SKIP (tidak lolos filter)")
            return False

        merged_file = None
        for rd in info.get("requested_downloads", []):
            fp = rd.get("filepath") or rd.get("_filename")
            if fp and os.path.exists(fp) and fp.lower().endswith((".mp4", ".mkv", ".webm")):
                merged_file = fp
                break

        if not merged_file:
            t = info.get("title", "")
            for fname in os.listdir(OUTPUT_FOLDER):
                if t and t[:30] in fname and fname.lower().endswith(".mp4") and "_final" not in fname:
                    merged_file = os.path.join(OUTPUT_FOLDER, fname)
                    break

        if not merged_file or not os.path.exists(merged_file):
            print("  GAGAL: file hasil download tidak ditemukan")
            return False

        print(f"  File: {merged_file}")
        final_path = add_logo_and_subtitles(merged_file)
        title = os.path.splitext(os.path.basename(merged_file))[0]
        upload_to_tiktok(final_path, title)
        print("  SUKSES")
        return True

    except Exception as e:
        print(f"  GAGAL: {e}")
        return False

# ============================================================
# SCRAPING YouTube
# ============================================================

def scrape_keyword():
    print(f"Scraping keyword: '{KEYWORD}'")

    ydl_opts = {"quiet": True, "extract_flat": True, "no_warnings": True}
    ydl_opts.update(get_cookie_config())

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            result = ydl.extract_info(f"ytsearch{SEARCH_COUNT}:{KEYWORD}", download=False)

        videos = [v for v in result["entries"] if v]
        if SHUFFLE:
            random.shuffle(videos)
        print(f"Ditemukan {len(videos)} kandidat.")

        existing = set(os.path.splitext(f)[0] for f in os.listdir(OUTPUT_FOLDER))
        success_count = 0

        for i, video in enumerate(videos, start=1):
            if success_count >= MAX_VIDEO:
                break

            title = video.get("title", "")
            if any(title in name or name.startswith(title[:30]) for name in existing if title):
                print(f"[{i}] Skip (sudah ada): {title}")
                continue

            url = f"https://www.youtube.com/watch?v={video['id']}"
            if download_video(url, i):
                success_count += 1
            if success_count < MAX_VIDEO and i < len(videos):
                time.sleep(random.randint(3, 7))

    except Exception as e:
        print(f"ERROR scraping: {e}")

# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    print("=" * 50)
    print("=== TIKTOK AUTOPOST — HEADLESS ===")
    download_logo()
    scrape_keyword()
    print("Selesai.")
    print("=" * 50)
