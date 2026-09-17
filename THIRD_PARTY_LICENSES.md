# Third-Party Licenses

โปรแกรม **PDF Reader** นี้ใช้ไลบรารีจากบุคคลที่สาม (third-party) ดังต่อไปนี้:

| Library | Purpose | License | Copyright |
|---------|---------|---------|-----------|
| [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/) | อ่าน/แก้ไข PDF, render, ค้นหา, redact, form widgets | **AGPL-3.0** | © Artifex Software, Inc. |
| [Pillow (PIL)](https://python-pillow.org/) | จัดการรูปภาพและสร้าง icon | HPND | © Jeffrey A. Clark and contributors |
| [PyInstaller](https://pyinstaller.org/) | สร้าง .exe (build-time เท่านั้น) | GPL + runtime exception | © PyInstaller Development Team |
| [Python](https://www.python.org/) | Runtime | PSF License | © Python Software Foundation |
| [Tk / Tcl](https://www.tcl.tk/) | GUI toolkit (ผ่าน tkinter) | BSD-like | © Regents of the University of California et al. |

## หมายเหตุสำคัญ

### PyMuPDF — AGPL-3.0
PyMuPDF เป็น library แบบ **AGPL-3.0 copyleft** ซึ่งเป็นเหตุผลหลักที่โปรแกรมนี้ต้องเผยแพร่ภายใต้ AGPL-3.0 ด้วย
หากต้องการใช้ในเชิงพาณิชย์แบบปิดโค้ด สามารถซื้อ commercial license ได้ที่ [artifex.com](https://artifex.com/products/pymupdf-pro/)

### PyInstaller runtime exception
PyInstaller มี exception ที่อนุญาตให้ผลลัพธ์ (.exe) ไม่ต้องเป็น GPL ตาม PyInstaller เอง — แต่ยังต้องเคารพ license ของ library ที่ bundled ไป (เช่น PyMuPDF ยังคงเป็น AGPL)

### สรุป
- โค้ดต้นฉบับ (source code): AGPL-3.0
- Binary distribution (.exe): AGPL-3.0
- ผู้ที่แจกจ่ายซ้ำต้องเผยแพร่โค้ดตามเงื่อนไข AGPL-3.0

รายละเอียดเต็ม: ดูไฟล์ `LICENSE`
