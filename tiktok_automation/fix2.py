"""Fix2: patch tiktok_ugc_auto.py - handle Windows & Unix line endings"""
import os

fname = "tiktok_ugc_auto.py"
if not os.path.exists(fname):
    print(f"ERROR: {fname} tidak ditemukan di folder ini.")
    print(f"Folder saat ini: {os.getcwd()}")
    input("Tekan Enter untuk keluar...")
    exit()

with open(fname, "rb") as f:
    raw = f.read()

# Cek apakah sudah ter-patch
if b"isinstance(payload, dict)" in raw:
    print("File sudah ter-patch sebelumnya, tidak perlu diubah.")
    input("Tekan Enter untuk keluar...")
    exit()

# Deteksi line ending (Windows \r\n atau Unix \n)
nl = b"\r\n" if b"\r\n" in raw else b"\n"
lines = raw.split(nl)

result = []
patched = False

for i, line in enumerate(lines):
    result.append(line)
    # Cari baris "for payload in captured:"
    if line.strip() == b"for payload in captured:":
        next_line = lines[i+1].strip() if i + 1 < len(lines) else b""
        if b"isinstance" not in next_line:
            indent = len(line) - len(line.lstrip())
            check = b" " * (indent + 4) + b"if not isinstance(payload, dict):"
            cont  = b" " * (indent + 8) + b"continue"
            result.append(check)
            result.append(cont)
            patched = True
            print(f"  Patch diterapkan di baris {i+1}")

if patched:
    with open(fname, "wb") as f:
        f.write(nl.join(result))
    print("\nPATCH BERHASIL!")
    print("Sekarang jalankan: python tiktok_ugc_auto.py")
else:
    print("GAGAL: Tidak bisa menemukan lokasi patch.")
    print("Buka file manual, cari 'for payload in captured:',")
    print("tambahkan 2 baris di bawahnya:")
    print("        if not isinstance(payload, dict):")
    print("            continue")

input("\nTekan Enter untuk keluar...")
