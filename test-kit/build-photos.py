"""
The three test photos, in three deliberately different shapes.

    pip install pillow && python3 test-kit/build-photos.py

Run rarely, by hand — see build-template.py.
"""

import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))

# Three deliberately different SHAPES, so a cover-crop is visible at a glance:
# a wide one loses its sides, a tall one loses top and bottom, a square one
# fills exactly. Different colours so the wrong photo on the wrong row is
# obvious from across the room.
SPEC = [
    ("ada.png",   (900, 450), (36, 82, 140),  "ADA"),    # wide
    ("grace.png", (450, 900), (150, 60, 40),  "GRACE"),  # tall
    ("alan.png",  (700, 700), (30, 110, 80),  "ALAN"),   # square
]

for name, (w, h), colour, label in SPEC:
    img = Image.new("RGB", (w, h), colour)
    d = ImageDraw.Draw(img)
    # A border and corner marks: if a corner mark is missing in the merged
    # deck, the picture was cropped, which is what `image` is meant to do.
    d.rectangle([8, 8, w - 9, h - 9], outline=(255, 255, 255), width=6)
    for cx, cy in [(40, 40), (w - 40, 40), (40, h - 40), (w - 40, h - 40)]:
        d.ellipse([cx - 22, cy - 22, cx + 22, cy + 22], fill=(255, 200, 40))
    d.text((w / 2 - len(label) * 14, h / 2 - 12), label, fill=(255, 255, 255))
    img.save(os.path.join(HERE, name))
    print(name, w, "x", h)
