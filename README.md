# PDF Reader

โปรแกรมอ่านและจัดการไฟล์ PDF แบบ Desktop สำหรับ Windows | Lightweight desktop PDF viewer & toolkit for Windows

เขียนด้วย Python + tkinter ครบใน 1 ไฟล์ ไม่มี dependency ยุ่งยาก
Written in Python + tkinter, single-file, minimal dependencies.

---

## 📸 Screenshots

> วางไฟล์รูปในโฟลเดอร์ `docs/` แล้วอัปเดตลิงก์ด้านล่าง
> Put screenshots in `docs/` folder and update the links below

<p align="center">
  <img src="docs/screenshot-main.png" width="700" alt="Main window"/>
  <br><em>หน้าจอหลัก อ่าน PDF พร้อม continuous scroll / Main viewer with continuous scroll</em>
</p>

<p align="center">
  <img src="docs/screenshot-create.png" width="500" alt="Create PDF from text"/>
  <br><em>สร้าง PDF จากข้อความ (รองรับภาษาไทย) / Create PDF from text (Thai supported)</em>
</p>

---

## 🎯 ทำไมโปรเจกต์นี้ / Why this project?

- 🇹🇭 **เมนูภาษาไทย** — ตัวเลือกส่วนใหญ่บน GitHub เป็นภาษาอังกฤษล้วน
  Thai menu — most alternatives on GitHub are English-only
- 📦 **1 ไฟล์ Python** — ไม่ต้องรัน Docker, ไม่ต้อง build
  Single Python file — no Docker, no build step
- ⚙️ **Auto-install dependencies ครั้งเดียว** — ผู้ใช้ไม่รำคาญ
  One-time auto-install — no repeated prompts
- 🎨 **UI นิ่ง** — smooth scroll, จัดกลาง, fit toggle
  Polished UI — smooth scrolling, centered layout, fit-width toggle

---

## ✨ ฟีเจอร์ / Features

| ฟีเจอร์ / Feature | รายละเอียด / Details |
|---|---|
| 📖 อ่าน PDF / View PDF | Continuous scroll, smooth scrolling, ซูม / zoom, Fit width toggle |
| 🔍 ค้นหาข้อความ / Search | ไฮไลต์ทุกจุดที่พบ กระโดดทีละที่ / Highlights all matches, jump one by one |
| ✏️ สร้าง PDF จากข้อความ / Text → PDF | รองรับภาษาไทย / Thai supported (Tahoma) |
| 🖼️ รวมรูป → PDF / Images → PDF | เลือกหลายไฟล์ / Multiple images at once |
| 🔗 Merge PDFs | รวม PDF หลายไฟล์ / Combine multiple PDFs |
| ✂️ Split PDF | แยกทุกหน้าเป็นไฟล์ย่อย / Split into single-page files |
| 📄 Word → PDF | ต้องมี Microsoft Word / Requires MS Word |
| 🖨️ พิมพ์ / Print | ส่งไปเครื่องพิมพ์ default / Sends to default printer |
| ⛶ Full Screen | F11 toggle |

---

## 💻 ความต้องการของระบบ / Requirements

- Windows 10 / 11
- Python 3.9+ ([download](https://www.python.org/downloads/) — check **"Add Python to PATH"**)
- Microsoft Word (สำหรับ Word→PDF เท่านั้น / only for Word→PDF)

---

## 🚀 วิธีติดตั้งและใช้งาน / Installation

**ภาษาไทย:**
1. โคลนหรือดาวน์โหลด repo นี้
2. ดับเบิลคลิก `run.bat`
   - ครั้งแรกจะติดตั้ง dependencies อัตโนมัติ (PyMuPDF, Pillow, docx2pdf, pywin32)
   - ครั้งต่อไปเปิดโปรแกรมได้ทันทีโดยไม่มีหน้าต่าง cmd

**English:**
1. Clone or download this repo
2. Double-click `run.bat`
   - First run installs dependencies automatically
   - Subsequent runs launch silently without a console window

Or manually:
```bash
pip install pymupdf Pillow docx2pdf pywin32
python pdf_reader.py
```

---

## ⌨️ คีย์ลัด / Keyboard Shortcuts

| คีย์ / Key | ทำอะไร / Action |
|---|---|
| `Ctrl+O` | เปิดไฟล์ / Open file |
| `Ctrl+P` | พิมพ์ / Print |
| `Ctrl+F` | ค้นหา / Focus search |
| `←` / `→` | หน้าก่อนหน้า/ถัดไป / Prev/Next page |
| `Ctrl+MouseWheel` | ซูม / Zoom |
| `F11` | Full Screen |
| `Esc` | ออกจาก Full Screen / Exit Full Screen |

---

## 📚 Libraries ที่ใช้ / Dependencies

| Library | หน้าที่ / Purpose | License |
|---|---|---|
| [PyMuPDF](https://pymupdf.readthedocs.io/) | อ่าน/สร้าง PDF / PDF I/O & rendering | AGPL-3.0 |
| [Pillow](https://pillow.readthedocs.io/) | จัดการรูปภาพ / Image processing | HPND |
| [docx2pdf](https://github.com/AlJohri/docx2pdf) | Word → PDF | MIT |
| [pywin32](https://github.com/mhammond/pywin32) | Windows API | PSF |

---

## 📄 License

โปรแกรมนี้เผยแพร่ภายใต้ **GNU Affero General Public License v3.0 (AGPL-3.0)**
Released under **GNU AGPL-3.0** because it uses PyMuPDF (AGPL).

หาก fork หรือ redistribute → ต้องเปิดซอร์สโค้ดของคุณภายใต้ AGPL-3.0 เช่นกัน
If you fork or redistribute, your code must also be AGPL-3.0.

ดูรายละเอียดเต็มใน / See full text in [LICENSE](LICENSE)

---

## 👤 ผู้พัฒนา / Author

thanagrid.c@ku.th

## 🤝 Contributing

ยินดีรับ issue / pull request ครับ
Issues and pull requests are welcome!
