"""fix3: wrap _parse_captured_responses dalam try/except"""
import os

fname = "tiktok_ugc_auto.py"
if not os.path.exists(fname):
    print(f"ERROR: {fname} tidak ditemukan. Pastikan fix3.py ada di folder yang sama.")
    input("Enter..."); exit()

with open(fname, "rb") as f:
    raw = f.read()

target = b"products = _parse_captured_responses(captured, keyword)"
if target not in raw:
    print("ERROR: Baris target tidak ditemukan.")
    input("Enter..."); exit()

nl   = b"\r\n" if b"\r\n" in raw else b"\n"
lines = raw.split(nl)

out     = []
patched = False
for line in lines:
    if line.strip() == target and not patched:
        ind = b" " * (len(line) - len(line.lstrip()))
        out.append(ind + b"try:")
        out.append(ind + b"    " + line.strip())
        out.append(ind + b"except Exception:")
        out.append(ind + b"    products = []")
        patched = True
    else:
        out.append(line)

if patched:
    with open(fname, "wb") as f:
        f.write(nl.join(out))
    print("PATCH BERHASIL! Jalankan: python tiktok_ugc_auto.py")
else:
    print("GAGAL: baris tidak ditemukan atau sudah ter-patch.")

input("Enter...")
