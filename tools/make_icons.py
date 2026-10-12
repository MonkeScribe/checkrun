"""Draw the pixel-art app icon (a paper check + green arrow) and write assets/icon.png and icon.ico.

Run: python tools/make_icons.py   (needs Pillow; outputs are committed, so this is only for edits)
"""
from pathlib import Path

from PIL import Image

OUT = Path(__file__).resolve().parent.parent / "assets"
N = 32

K = (38, 38, 54, 255)       # outline
P = (246, 242, 224, 255)    # paper
PS = (222, 216, 190, 255)   # paper shade (bottom edge)
G = (150, 150, 165, 255)    # grey text lines
B = (58, 110, 200, 255)     # blue ink (signature / amount)
H = (92, 190, 110, 255)     # header stripe on the check
A = (47, 179, 68, 255)      # arrow green
AD = (24, 112, 44, 255)     # arrow outline
AL = (130, 225, 140, 255)   # arrow highlight

g = [[None] * N for _ in range(N)]


def px(x, y, c):
    g[y][x] = c


def rect(x0, y0, x1, y1, c):
    for y in range(y0, y1 + 1):
        for x in range(x0, x1 + 1):
            px(x, y, c)


# --- the check: x 1..20, y 8..24 -------------------------------------------
rect(1, 8, 20, 24, K)
rect(2, 9, 19, 23, P)
rect(2, 23, 19, 23, PS)
rect(2, 9, 19, 10, H)          # green header stripe
rect(14, 12, 18, 12, G)        # date line
rect(3, 15, 12, 15, G)         # payee line
rect(14, 14, 18, 16, B)        # amount box
rect(15, 15, 17, 15, P)
rect(3, 18, 9, 18, G)          # memo line
for x, y in [(11, 21), (12, 20), (13, 21), (14, 20), (15, 19), (16, 20), (17, 21)]:
    px(x, y, B)                # signature squiggle
rect(3, 21, 9, 21, G)

# --- the arrow: pointing right, x 22..31 -----------------------------------
# outline
rect(22, 14, 24, 19, AD)
for i, (y0, y1) in enumerate([(9, 24), (10, 23), (11, 22), (12, 21), (13, 20), (14, 19), (15, 18), (16, 17)]):
    rect(24 + i, y0, 24 + i, y1, AD)
# fill
rect(23, 15, 24, 18, A)
for i, (y0, y1) in enumerate([(11, 22), (12, 21), (13, 20), (14, 19), (15, 18)]):
    rect(25 + i, y0, 25 + i, y1, A)
rect(23, 15, 24, 15, AL)       # highlight on top of shaft
rect(25, 11, 25, 12, AL)

img = Image.new("RGBA", (N, N), (0, 0, 0, 0))
for y in range(N):
    for x in range(N):
        if g[y][x]:
            img.putpixel((x, y), g[y][x])

OUT.mkdir(exist_ok=True)
big = img.resize((1024, 1024), Image.NEAREST)
big.save(OUT / "icon.png")
big.save(OUT / "icon.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("wrote", OUT / "icon.png", OUT / "icon.ico")
