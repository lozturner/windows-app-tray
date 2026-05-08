"""
gen_star_frames.py — generates 8 frames of an animated Mario-style star
into ../resources/star_NN.ico, each containing 16x16 + 32x32 sub-images
so Windows can pick the right size for the notification area.

Run:
    python tools/gen_star_frames.py

Outputs (relative to repo root):
    resources/star_00.ico .. star_07.ico
"""
from PIL import Image, ImageDraw
import math, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT  = os.path.normpath(os.path.join(HERE, "..", "resources"))
os.makedirs(OUT, exist_ok=True)

def star_points(cx, cy, r_outer, r_inner, n_points=5, rot=-math.pi/2):
    pts = []
    for i in range(n_points * 2):
        r = r_outer if i % 2 == 0 else r_inner
        a = rot + i * math.pi / n_points
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts

# Colour cycle — Mario-star rainbow
COLOURS = [
    (255, 80, 80),    # red
    (255, 160, 40),   # orange
    (255, 230, 40),   # yellow
    (120, 220, 80),   # green
    (60, 220, 220),   # cyan
    (60, 120, 255),   # blue
    (180, 90, 220),   # purple
    (255, 120, 200),  # pink
]

def render(size, fill, outline=(0, 0, 0, 255)):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    margin = max(1, size // 16)
    cx = cy = size / 2
    pts = star_points(cx, cy, size / 2 - margin, size / 4)
    d.polygon(pts, fill=fill + (255,), outline=outline)
    # tiny eye / highlight to look more Mario-ish
    eye = max(1, size // 8)
    d.ellipse((cx - eye*0.6, cy - eye*0.4, cx + eye*0.6, cy + eye*1.2),
              fill=(255, 255, 255, 255), outline=(0, 0, 0, 255))
    return img

def save_ico(path, base_colour):
    # Render at a high resolution so the embedded sub-sizes look crisp.
    big = render(256, base_colour)
    big.save(path, format="ICO",
             sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (256, 256)])

for i, c in enumerate(COLOURS):
    p = os.path.join(OUT, f"star_{i:02d}.ico")
    save_ico(p, c)
    print(f"wrote {p}")

print(f"\nDone — {len(COLOURS)} frames in {OUT}")
