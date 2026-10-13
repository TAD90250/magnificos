#!/usr/bin/env python3
"""
Gets photos ready for the league website.

  PUBLIC photos (shown on the home page, no password):
      python prepare_photos.py --public "C:\\path\\to\\public_photos"
      (adds only NEW photos; the display order is the order in images/public/manifest.json,
       so move entries around in that file to reorder.  --reset starts over.)

  PRIVATE photos (encrypted, shown only on photos.html after the password is typed):
      python prepare_photos.py --private "C:\\path\\to\\private_photos"

Each photo is shrunk for the web (long side 1600 px; private photos also get a small
thumbnail).  Name files like  2019_champions dinner.jpg  -- a 4-digit year at the start
groups the photo under that year, and the rest of the name becomes the caption.

PRIVATE photos are encrypted here with AES-256-GCM (key = PBKDF2-SHA256 of your password,
600,000 rounds) and the page decrypts them in the visitor's browser.  The files that go to
GitHub are unreadable without the password -- file names and captions are encrypted too.
Run it again later with the SAME password to add photos: only new photos are encrypted, so
you only upload the new files.  Use --reset to start the vault over (new password).

Needs:  pip install pillow cryptography
The password is typed at a prompt and is never saved.  Do not put it in the repo.
"""
import argparse
import base64
import getpass
import hashlib
import json
import os
import re
import secrets
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
import league_config  # noqa: F401  (runs the Camaradas isolation guard)
SITE = os.path.join(HERE, "site")
PUBLIC_DIR = os.path.join(SITE, "images", "public")
PRIVATE_DIR = os.path.join(SITE, "private")
EXTS = {".jpg", ".jpeg", ".png", ".webp", ".heic", ".bmp", ".gif", ".tif", ".tiff"}
KDF_ITER = 600_000
FULL_PX, THUMB_PX = 1600, 480


def need(mod, pkg):
    try:
        return __import__(mod)
    except ImportError:
        sys.exit(f"Missing package '{pkg}'. Install it with:  pip install {pkg}")


def list_photos(folder):
    out = []
    for root, _, files in os.walk(folder):
        for f in sorted(files):
            if os.path.splitext(f)[1].lower() in EXTS:
                out.append(os.path.join(root, f))
    return sorted(out)


def year_and_caption(path):
    stem = os.path.splitext(os.path.basename(path))[0]
    m = re.match(r"^((?:19|20)\d\d)[\s_\-\.]*(.*)$", stem)
    year, rest = (int(m.group(1)), m.group(2)) if m else (None, stem)
    cap = re.sub(r"[_]+", " ", rest).strip(" -.")
    return year, cap


def shrink(path, px, quality):
    from io import BytesIO
    from PIL import Image, ImageOps
    with Image.open(path) as im:
        im = ImageOps.exif_transpose(im)
        if im.mode not in ("RGB",):
            im = im.convert("RGB")
        im.thumbnail((px, px), Image.LANCZOS)
        buf = BytesIO()
        im.save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
        return buf.getvalue()


# ---------------------------------------------------------------- public
def do_public(folder, reset):
    need("PIL", "pillow")
    photos = list_photos(folder)
    if not photos:
        sys.exit(f"No photos found in {folder}")
    mpath = os.path.join(PUBLIC_DIR, "manifest.json")
    if reset and os.path.isdir(PUBLIC_DIR):
        shutil.rmtree(PUBLIC_DIR)
    os.makedirs(PUBLIC_DIR, exist_ok=True)
    manifest = []
    if os.path.exists(mpath):
        with open(mpath, encoding="utf-8") as f:
            manifest = json.load(f)
    have = {m.get("src") for m in manifest if m.get("src")}
    nums = [int(m["file"][1:4]) for m in manifest if re.match(r"p\d{3}\.jpg$", m["file"])]
    n = max(nums, default=0)
    new = []
    for p in photos:
        with open(p, "rb") as f:
            h = hashlib.sha256(f.read()).hexdigest()[:20]
        if h in have:
            continue
        have.add(h)
        year, cap = year_and_caption(p)
        new.append((year, p, cap, h))
    new.sort(key=lambda t: (t[0] is None, t[0] or 0, t[1]))      # new photos: by year, undated last
    for year, p, cap, h in new:
        n += 1
        name = f"p{n:03d}.jpg"
        with open(os.path.join(PUBLIC_DIR, name), "wb") as f:
            f.write(shrink(p, FULL_PX, 82))
        manifest.append({"file": name, "year": year, "caption": cap, "src": h})
    with open(mpath, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    print(f"{len(new)} new public photos added; the home page now shows {len(manifest)}.")
    print("They appear in the order listed in images/public/manifest.json (new ones go at the end).")
    print("To change the order, move the entries around in that file. Then upload images/public to GitHub.")


# ---------------------------------------------------------------- private
def do_private(folder, reset):
    need("PIL", "pillow")
    need("cryptography", "cryptography")
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    photos = list_photos(folder)
    if not photos:
        sys.exit(f"No photos found in {folder}")
    vault_path = os.path.join(PRIVATE_DIR, "vault.json")
    existing = os.path.exists(vault_path) and not reset

    if reset and os.path.isdir(PRIVATE_DIR):
        shutil.rmtree(PRIVATE_DIR)
    os.makedirs(PRIVATE_DIR, exist_ok=True)

    def derive(pw, salt, iters):
        return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=iters).derive(pw.encode("utf-8"))

    def seal(key, data):
        iv = secrets.token_bytes(12)
        return iv + AESGCM(key).encrypt(iv, data, None)

    items = []
    if existing:
        with open(vault_path, encoding="utf-8") as f:
            vault = json.load(f)
        salt, iters = base64.b64decode(vault["salt"]), vault["iter"]
        pw = getpass.getpass("Vault password: ")
        key = derive(pw, salt, iters)
        blob = base64.b64decode(vault["data"])
        try:
            items = json.loads(AESGCM(key).decrypt(blob[:12], blob[12:], None))
        except Exception:
            sys.exit("That password does not open the existing vault. Use --reset to start over.")
        print(f"Vault opened: {len(items)} photos already in it.")
    else:
        print("Choose a password for the photo vault. Use a long passphrase (several random words).")
        pw = getpass.getpass("New password: ")
        if len(pw) < 10:
            sys.exit("Please use at least 10 characters.")
        if getpass.getpass("Type it again: ") != pw:
            sys.exit("The two passwords did not match.")
        salt, iters = secrets.token_bytes(16), KDF_ITER
        key = derive(pw, salt, iters)

    have = {it["src"] for it in items}
    added = 0
    for p in photos:
        with open(p, "rb") as f:
            h = hashlib.sha256(f.read()).hexdigest()[:20]
        if h in have:
            continue
        pid = secrets.token_hex(6)
        year, cap = year_and_caption(p)
        with open(os.path.join(PRIVATE_DIR, pid + ".enc"), "wb") as f:
            f.write(seal(key, shrink(p, FULL_PX, 82)))
        with open(os.path.join(PRIVATE_DIR, pid + ".t.enc"), "wb") as f:
            f.write(seal(key, shrink(p, THUMB_PX, 74)))
        items.append({"id": pid, "year": year, "caption": cap, "src": h})
        have.add(h)
        added += 1
        print(f"  encrypted {os.path.basename(p)}")

    vault = {"v": 1, "kdf": "PBKDF2-SHA256", "iter": iters, "salt": base64.b64encode(salt).decode(),
             "data": base64.b64encode(seal(key, json.dumps(items, ensure_ascii=False).encode("utf-8"))).decode()}
    with open(vault_path, "w", encoding="utf-8") as f:
        json.dump(vault, f)
    print(f"{added} new photos encrypted; vault now holds {len(items)}.")
    print("Upload the private folder to GitHub (new files plus vault.json).")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--public", help="folder of photos for the home page")
    ap.add_argument("--private", help="folder of photos for the password-protected page")
    ap.add_argument("--reset", action="store_true", help="throw away what is already there (public photos or the private vault) and start over")
    a = ap.parse_args()
    if not a.public and not a.private:
        ap.print_help()
        return
    if a.public:
        do_public(a.public, a.reset)
    if a.private:
        do_private(a.private, a.reset)


if __name__ == "__main__":
    main()
