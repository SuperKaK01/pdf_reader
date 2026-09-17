# PDF Reader

โปรแกรมอ่านและแก้ไข PDF บน Windows — ครบเครื่องในไฟล์เดียว
A full-featured PDF viewer & editor for Windows, packed into one file.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL_v3-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows_10%2F11-lightgrey)]()
[![Python](https://img.shields.io/badge/python-3.9%2B-yellow)]()

---

## 📥 ดาวน์โหลด / Download

โหลด `.exe` พร้อมใช้จาก **[Releases](https://github.com/SuperKaK01/pdf_reader/releases)** — ไม่ต้องลง Python

Download the ready-to-run `.exe` from the [Releases page](https://github.com/SuperKaK01/pdf_reader/releases) — no Python required.

---

## 🎯 ทำไมโปรเจกต์นี้ / Why?

- 🇹🇭 **เมนูภาษาไทย** — โปรแกรม PDF ที่รองรับคนไทยจริงๆ
- 📦 **ไม่ต้องลง Python** — .exe เดียวใช้ได้เลย 40 MB
- 🆓 **ฟรี ไม่มีลายน้ำ ไม่มีขายลิขสิทธิ์** — ทดแทน Adobe Acrobat ระดับพื้นฐาน
- 🔒 **Redact จริงๆ** — ลบข้อความออกจาก PDF ได้ ไม่ใช่แค่บัง (สำคัญสำหรับหน่วยงาน)
- ✏️ **ครบทุกอย่างที่ต้องแก้** — ไฮไลต์, เขียน, เซ็น, ใส่รูป, คอมเมนต์, กรอกฟอร์ม

---

## ✨ ฟีเจอร์ / Features

### การอ่านและนำทาง (Viewing & Navigation)
- 📖 **Multi-tab** — เปิดหลายไฟล์พร้อมกัน
- 🖼️ **Thumbnail sidebar** — คลิกกระโดดไปหน้าที่ต้องการ (toggle ได้)
- 🖱️ **Smooth scroll** — เลื่อนต่อเนื่องหลายหน้า
- 🔍 **ซูม / Fit width** — Ctrl+MouseWheel หรือปุ่ม +/−
- ⛶ **Full Screen** — F11

### การเลือกและคัดลอกข้อความ (Selection & Copy)
- ✂️ **ลากเลือกทีละอักษร** เหมือน Adobe
- 📑 **เลือกข้ามหน้าได้** (cross-page selection)
- 🖱️ **คลิกขวา → Copy** พร้อม Select page / Clear
- คัดลอกอัตโนมัติเมื่อปล่อยเมาส์

### การแก้ไข (Editing)
- 🖍️ **ไฮไลต์** — ลากคลุมเน้นสีเหลือง
- ✏️ **ขีดเขียนอิสระ** — freehand drawing
- 💬 **คอมเมนต์** — sticky note ใน PDF
- 🖼️ **ใส่รูปภาพ**
- ✍️ **เซ็นชื่อ** — วาดลายเซ็นแล้ววางในเอกสาร
- 🗑️ **ลบข้อความ (Redact)** — ลบข้อมูลออกจริง copy ก็ไม่ติด
- 🔄 **หมุนหน้า** — หน้าเดียว หรือทุกหน้า
- 📝 **กรอก PDF Form** — text field, checkbox, combobox

### ประวัติและความปลอดภัย (Safety)
- ↶ **Undo / Redo** — สูงสุด 20 การแก้ไข (Ctrl+Z / Ctrl+Y)
- 💾 **เตือนบันทึกก่อนปิด** — ไม่พลาดข้อมูล
- 🔒 **กันเขียนทับไฟล์ต้นฉบับ** — บันทึกเป็นไฟล์ใหม่เสมอ

### เครื่องมือเสริม (Utilities)
- 📝 **สร้าง PDF จากข้อความ** (รองรับภาษาไทย)
- 🖼️ **รวมรูปภาพเป็น PDF**
- 🔗 **Merge PDFs**
- ✂️ **Split PDF** — แยกทีละหน้า
- 📄 **Word → PDF** (ต้องมี Microsoft Word)
- 🖨️ **พิมพ์** — ผ่านเครื่องพิมพ์ default หรือ Edge fallback

---

## 💻 ความต้องการของระบบ / Requirements

- Windows 10 / 11 (64-bit)
- ~ 60 MB พื้นที่ว่าง
- Microsoft Word (สำหรับ Word→PDF เท่านั้น)

---

## 🚀 ติดตั้งและใช้งาน / Installation

### วิธีที่ 1: ใช้ .exe (แนะนำ)
1. โหลด `PDFReader.exe` จาก [Releases](https://github.com/SuperKaK01/pdf_reader/releases)
2. ดับเบิลคลิกเปิดใช้งาน
3. (ทางเลือก) ตั้งเป็นโปรแกรมเริ่มต้นสำหรับ .pdf ได้ที่ Windows Settings

### วิธีที่ 2: รันจากซอร์สโค้ด
```bash
git clone https://github.com/SuperKaK01/pdf_reader.git
cd pdf_reader
pip install pymupdf Pillow
python pdf_reader.py
```

### วิธีที่ 3: build .exe เอง
```bash
pip install pyinstaller
pyinstaller PDFReader.spec
# ผลลัพธ์อยู่ใน dist/PDFReader.exe
```

---

## ⌨️ คีย์ลัด / Keyboard Shortcuts

| คีย์ / Key | ทำอะไร / Action |
|---|---|
| `Ctrl+O` | เปิดไฟล์ / Open |
| `Ctrl+W` | ปิดแท็บ / Close tab |
| `Ctrl+Tab` | สลับแท็บ / Switch tab |
| `Ctrl+Z` / `Ctrl+Y` | Undo / Redo |
| `Ctrl+F` | ค้นหา / Search |
| `Ctrl+P` | พิมพ์ / Print |
| `←` / `→` | เปลี่ยนหน้า / Prev / Next page |
| `Ctrl+MouseWheel` | ซูม / Zoom |
| `F11` | Full Screen |
| `Esc` | ออกจากโหมด / Exit mode / fullscreen |

---

## 🚧 ข้อจำกัดที่ทราบ / Known Limitations

**อยู่ระหว่างพัฒนา:**
- การพิมพ์ในโปรแกรม — ปัจจุบันใช้ Microsoft Edge เป็น fallback หาก default handler ไม่รองรับ

**แผนอนาคต (Planned):**
- OCR สำหรับ PDF สแกน (Tesseract)
- ลายเซ็นดิจิทัลแบบ certificate (Digital Signature)
- รองรับ PDF form ทุก type (radio, listbox, digital signature widgets)
- Batch processing หลายไฟล์พร้อมกัน
- Dark mode

**โดยดีไซน์:**
- ใช้ได้บน Windows เท่านั้น (โค้ดพร้อม cross-platform แต่ยังไม่ได้ test บน Mac/Linux)
- ไม่รองรับ PDF ที่ encrypted ด้วยรหัสผ่าน

---

## 📚 Libraries ที่ใช้ / Dependencies

| Library | หน้าที่ | License |
|---|---|---|
| [PyMuPDF](https://pymupdf.readthedocs.io/) | อ่าน/แก้ไข/render PDF | **AGPL-3.0** |
| [Pillow](https://pillow.readthedocs.io/) | จัดการรูป, สร้าง icon | HPND |
| [PyInstaller](https://pyinstaller.org/) | Build .exe | GPL + runtime exception |

รายละเอียดเต็ม: [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)

---

## 📄 License

โปรแกรมนี้เผยแพร่ภายใต้ **GNU Affero General Public License v3.0 (AGPL-3.0)**
เพราะใช้ PyMuPDF ซึ่งเป็น AGPL-3.0

Released under **GNU AGPL-3.0** because it uses PyMuPDF (AGPL).

หาก fork หรือ redistribute → ต้องเปิดซอร์สโค้ดของคุณภายใต้ AGPL-3.0 เช่นกัน
If you fork or redistribute, your code must also be AGPL-3.0.

ดูรายละเอียดเต็มใน [LICENSE](LICENSE)

---

## 👤 ผู้พัฒนา / Author

**Thanagrid C.** — thanagrid.c@ku.th
Faculty of Architecture, Kasetsart University
Version 1.0

---

## 🤝 Contributing

ยินดีรับ issue / pull request ครับ
Issues and pull requests are welcome!

หากมี bug หรือ feature request → เปิด [Issue](https://github.com/SuperKaK01/pdf_reader/issues)

---

## 🌟 หากชอบโปรเจกต์นี้

กด ⭐ Star ให้กำลังใจได้นะครับ!
If you like this project, please give it a ⭐ Star!
