"""Generate red PDF icon (icon.ico) — run once."""
from PIL import Image, ImageDraw, ImageFont

SIZES = [16, 32, 48, 64, 128, 256]
RED = (220, 30, 30, 255)
WHITE = (255, 255, 255, 255)


def make(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    r = max(2, size // 8)
    d.rounded_rectangle([(0, 0), (size - 1, size - 1)], radius=r, fill=RED)
    # "PDF" text
    try:
        font = ImageFont.truetype("arialbd.ttf", int(size * 0.42))
    except Exception:
        font = ImageFont.load_default()
    text = "PDF"
    bbox = d.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    d.text(((size - tw) / 2 - bbox[0], (size - th) / 2 - bbox[1]),
           text, fill=WHITE, font=font)
    return img


images = [make(s) for s in SIZES]
images[0].save("icon.ico", sizes=[(s, s) for s in SIZES])
print("Wrote icon.ico")
