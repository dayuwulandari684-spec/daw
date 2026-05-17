"""Jalankan file ini sekali untuk memperbaiki tiktok_ugc_auto.py"""
import re, os

target = "tiktok_ugc_auto.py"
if not os.path.exists(target):
    print("ERROR: tiktok_ugc_auto.py tidak ditemukan di folder ini.")
    input("Tekan Enter untuk keluar...")
    exit()

content = open(target, encoding="utf-8").read()

if "isinstance(payload, dict)" in content:
    print("File sudah ter-patch, tidak perlu diperbaiki lagi.")
    input("Tekan Enter untuk keluar...")
    exit()

# Cari baris "for payload in captured:" dan tambahkan pengecekan setelahnya
fixed = re.sub(
    r'(    for payload in captured:\n)',
    r'\1        if not isinstance(payload, dict):\n            continue\n',
    content,
    count=1
)

if fixed == content:
    print("Tidak bisa menemukan bagian yang perlu dipatch.")
    print("Coba hapus file tiktok_ugc_auto.py dan download ulang dari Claude.")
    input("Tekan Enter untuk keluar...")
    exit()

open(target, "w", encoding="utf-8").write(fixed)
print("PATCH BERHASIL! Sekarang jalankan: python tiktok_ugc_auto.py")
input("Tekan Enter untuk keluar...")
