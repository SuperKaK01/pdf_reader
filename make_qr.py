"""Generate feedback QR code (qr_feedback.png). Re-run after changing FEEDBACK_URL."""
import qrcode
from qrcode.constants import ERROR_CORRECT_H

# TODO: แทนที่ URL นี้ด้วยลิงก์ Google Form จริงเมื่อสร้างเสร็จ
FEEDBACK_URL = "https://forms.gle/YOUR-FORM-ID-HERE"

qr = qrcode.QRCode(
    version=None,
    error_correction=ERROR_CORRECT_H,
    box_size=10,
    border=2,
)
qr.add_data(FEEDBACK_URL)
qr.make(fit=True)
img = qr.make_image(fill_color="#b81d13", back_color="white")
img.save("qr_feedback.png")
print(f"Wrote qr_feedback.png for {FEEDBACK_URL}")
