"""Generate app icon (icon.ico) — bold, high-contrast, supersampled."""
from PIL import Image, ImageDraw, ImageFont

SIZES = [16, 20, 24, 32, 40, 48, 64, 96, 128, 256]
RED = (215, 25, 25, 255)
RED_DARK = (150, 15, 15, 255)
WHITE = (255, 255, 255, 255)
SUPERSAMPLE = 4  # render ที่ 4× แล้ว downsample เพื่อ anti-aliasing ที่คมชัด


def load_font(size):
    for name in ("arialbd.ttf", "seguisb.ttf", "impact.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            continue
    return ImageFont.load_default()


def make_hires(size):
    s = size * SUPERSAMPLE
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # ขอบมนอย่างเดียว ไม่มีมุมพับ — ใหญ่ ชัด อ่านง่าย
    radius = max(6, s // 7)
    d.rounded_rectangle([(0, 0), (s - 1, s - 1)], radius=radius, fill=RED)

    # เงาด้านล่างเล็กน้อยเพื่อความลึก
    inner_rect = [(s * 0.04, s * 0.04), (s - s * 0.04 - 1, s - s * 0.04 - 1)]
    # (ข้ามการใส่เงาถ้าเล็ก)

    # ตัวอักษร — เลือกตามขนาด
    if size <= 20:
        text = "P"
        font = load_font(int(s * 0.78))
    elif size <= 40:
        text = "PDF"
        font = load_font(int(s * 0.44))
    else:
        text = "PDF"
        font = load_font(int(s * 0.46))

    bbox = d.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    tx = (s - tw) / 2 - bbox[0]
    ty = (s - th) / 2 - bbox[1]
    d.text((tx, ty), text, fill=WHITE, font=font)

    # ปรับ downsample ด้วย LANCZOS ให้คมชัด
    return img.resize((size, size), Image.LANCZOS)


images = [make_hires(s) for s in SIZES]
# บันทึกเป็น .ico รวมทุกขนาด
images[0].save("icon.ico", sizes=[(s, s) for s in SIZES])
# บันทึก preview PNG 256px ด้วย (สำหรับดูตัวอย่าง)
images[-1].save("icon_preview.png")
print(f"Wrote icon.ico with sizes: {SIZES}")
