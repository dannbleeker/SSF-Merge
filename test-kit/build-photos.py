"""
The three test photos, in three deliberately different shapes.

    pip install pillow && python3 test-kit/build-photos.py

Run rarely, by hand — see build-template.py.
"""

import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))

# The picture frame on slide 2 of the template, 3.20 x 2.40in. Everything below
# is placed against THIS number, because a cover-crop is defined by the frame it
# crops into: a marker placed without reference to it cannot be read.
FRAME_RATIO = 3.20 / 2.40

# Three deliberately different SHAPES, so a cover-crop is visible at a glance:
# a wide one loses its sides, a tall one loses top and bottom, a square one
# loses a little of each end. Different colours so the wrong photo on the wrong
# row is obvious from across the room.
SPEC = [
    ("ada.png", (900, 450), (36, 82, 140), "ADA"),  # wide
    ("grace.png", (450, 900), (150, 60, 40), "GRACE"),  # tall
    ("alan.png", (700, 700), (30, 110, 80), "ALAN"),  # square
]

DOT_R = 26


def safe_box(w, h):
    """
    The part of this image a correct cover-crop into the frame keeps.

    Anything outside it is gone when the merge does its job, so it is the only
    place a marker meant to be SEEN can live.
    """
    if w / h > FRAME_RATIO:  # wider than the frame: the sides go
        vis_w = h * FRAME_RATIO
        return ((w - vis_w) / 2, 0, (w + vis_w) / 2, h)
    vis_h = w / FRAME_RATIO  # taller than the frame: top and bottom go
    return (0, (h - vis_h) / 2, w, (h + vis_h) / 2)


for name, (w, h), colour, label in SPEC:
    img = Image.new("RGB", (w, h), colour)
    d = ImageDraw.Draw(img)

    # The white border sits at the image's own edge, and it is what says WHICH
    # AXIS was cropped — the one marker that separates the three outcomes. Cover
    # keeps only the two edges on the long axis, so a wide image keeps top and
    # bottom and a tall one keeps left and right. A stretch keeps all four. A
    # letterbox keeps all four and adds bands of background.
    d.rectangle([8, 8, w - 9, h - 9], outline=(255, 255, 255), width=6)

    # The dots sit INSIDE the safe box, and they are round on purpose.
    #
    # They used to sit at the image's corners, where a correct crop removes
    # every one of them: they lay within 5% of each corner, and the crop takes a
    # third of the wide image's width and nearly two thirds of the tall one's
    # height. So docs/TEST-KIT.md asked a tester to confirm dots that could
    # never appear, and the round of 2026-08-28 read their absence as a failure
    # until the arithmetic was run. A check that cannot pass says nothing about
    # the code.
    #
    # Inside the safe box they survive every correct merge, which lets them
    # answer the question the corners could not: a circle drawn as an oval is a
    # stretch, and nothing else does that.
    x0, y0, x1, y1 = safe_box(w, h)
    m = DOT_R + 22
    for cx in (x0 + m, x1 - m):
        for cy in (y0 + m, y1 - m):
            d.ellipse([cx - DOT_R, cy - DOT_R, cx + DOT_R, cy + DOT_R], fill=(255, 200, 40))

    d.text((w / 2 - len(label) * 14, h / 2 - 12), label, fill=(255, 255, 255))
    img.save(os.path.join(HERE, name))
    print(f"{name} {w}x{h}  safe box x[{x0:.0f}-{x1:.0f}] y[{y0:.0f}-{y1:.0f}]")
