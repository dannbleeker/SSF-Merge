"""Tiny real JPEGs carrying each EXIF orientation, for test/image-orient.test.ts.

Real files rather than bytes assembled in the test: a decoder checked only
against its own encoder proves the two agree and nothing else. Pillow writes
the segment the way a camera does.

    python scripts/build-exif-fixtures.py
"""

from PIL import Image

OUT = "test/fixtures/exif"

for tag in [1, 2, 3, 4, 5, 6, 7, 8]:
    im = Image.new("RGB", (8, 4), (200, 60, 60))
    exif = Image.Exif()
    exif[274] = tag
    im.save(f"{OUT}/orientation-{tag}.jpg", "JPEG", exif=exif, quality=40)

# The common case: a JPEG with no EXIF at all.
Image.new("RGB", (8, 4), (60, 120, 200)).save(f"{OUT}/no-exif.jpg", "JPEG", quality=40)

# And a PNG, which cannot carry EXIF and must simply answer "nothing".
Image.new("RGB", (8, 4), (60, 200, 120)).save(f"{OUT}/plain.png", "PNG")

print("wrote fixtures to", OUT)
