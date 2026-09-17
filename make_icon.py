"""Generate app icon (icon.ico) — sharp, high-contrast with gradient and subtle depth."""
from PIL import Image, ImageDraw, ImageFont, ImageFilter

SIZES = [16, 20, 24, 32, 40, 48, 64, 96, 128, 256]
RED_TOP = (230, 32, 32, 255)
RED_BOT = (180, 12, 12, 255)
RED_EDGE = (120, 8, 8, 255)
WHITE = (255, 255, 255, 255)
SUPERSAMPLE = 8


def load_font(size):
    # Impact เป็นตัวหนาหนาเห็นชัดเจน
    for name in ("impact.ttf", "arialbd.ttf", "seguisb.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            continue
    return ImageFont.load_default()


def gradient_fill(size, top, bottom):
    """สร้าง gradient แนวตั้ง"""
    grad = Image.new("RGBA", (1, size), 0)
    for y in range(size):
        t = y / (size - 1)
        r = int(top[0] * (1 - t) + bottom[0] * t)
        g = int(top[1] * (1 - t) + bottom[1] * t)
        b = int(top[2] * (1 - t) + bottom[2] * t)
        grad.putpixel((0, y), (r, g, b, 255))
    return grad.resize((size, size))


def make_hires(size):
    s = size * SUPERSAMPLE
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    radius = max(6, s // 8)

    # 1. ตัวพื้นเรียบสี gradient
    grad = gradient_fill(s, RED_TOP, RED_BOT)
    mask = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [(0, 0), (s - 1, s - 1)], radius=radius, fill=255)
    img.paste(grad, (0, 0), mask)

    # 2. ขอบเข้ม (bevel รอบนอก)
    edge_w = max(2, s // 80)
    d.rounded_rectangle(
        [(0, 0), (s - 1, s - 1)],
        radius=radius, outline=RED_EDGE, width=edge_w)

    # 3. ไฮไลต์บน (แถบสว่างบางๆ ด้านบน)
    hl = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    hl_d = ImageDraw.Draw(hl)
    hl_d.rounded_rectangle(
        [(edge_w * 2, edge_w * 2), (s - 1 - edge_w * 2, s * 0.45)],
        radius=radius, fill=(255, 255, 255, 40))
    img = Image.alpha_composite(img, hl)
    d = ImageDraw.Draw(img)

    # 4. ตัวอักษร PDF/P
    if size <= 20:
        text = "P"
        font = load_font(int(s * 0.82))
        y_offset = 0
    elif size <= 40:
        text = "PDF"
        font = load_font(int(s * 0.5))
        y_offset = s * 0.02
    else:
        text = "PDF"
        font = load_font(int(s * 0.52))
        y_offset = s * 0.02

    bbox = d.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    tx = (s - tw) / 2 - bbox[0]
    ty = (s - th) / 2 - bbox[1] + y_offset

    # เงาตัวอักษรบางๆ ให้ลอย
    shadow_offset = max(1, s // 200)
    d.text((tx + shadow_offset, ty + shadow_offset),
           text, fill=(0, 0, 0, 90), font=font)
    d.text((tx, ty), text, fill=WHITE, font=font)

    return img.resize((size, size), Image.LANCZOS)


images = [make_hires(s) for s in SIZES]
images[0].save("icon.ico", sizes=[(s, s) for s in SIZES])
images[-1].save("icon_preview.png")
print(f"Wrote icon.ico with sizes: {SIZES}")
