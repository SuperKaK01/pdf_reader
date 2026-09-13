"""
PDF Reader - Desktop App
Copyright (C) 2026 thanagrid.c@ku.th
Licensed under GNU AGPL-3.0 — see LICENSE file for details.

ต้องติดตั้ง: pip install PyMuPDF Pillow docx2pdf pywin32
รัน: python pdf_reader.py
"""
import os
import sys
import subprocess

# ---------- First-run setup: ติดตั้ง dependencies ครั้งเดียว ----------
REQUIRED = ["pymupdf", "Pillow", "docx2pdf", "pywin32"]
_MARKER = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".setup_done")

def _ensure_deps():
    if os.path.exists(_MARKER):
        return
    missing = []
    checks = {"pymupdf": "pymupdf", "Pillow": "PIL",
              "docx2pdf": "docx2pdf", "pywin32": "win32api"}
    for pkg, mod in checks.items():
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"[Setup] กำลังติดตั้ง: {', '.join(missing)} ...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
        except Exception as e:
            print(f"[Setup] ผิดพลาด: {e}")
            input("กด Enter เพื่อออก...")
            sys.exit(1)
    open(_MARKER, "w").close()  # จำว่าติดตั้งแล้ว
    print("[Setup] เสร็จสิ้น — ครั้งต่อไปจะไม่ถามอีก")

_ensure_deps()

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
try:
    import pymupdf as fitz
except ImportError:
    import fitz
from PIL import Image, ImageTk


class PDFReader(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Reader")
        self.geometry("1000x750")

        self.doc = None
        self.doc_path = None
        self.page_index = 0
        self.zoom = 1.25
        self.search_hits = []      # list of (page_index, rect)
        self.search_pos = -1
        self.photos = []           # PhotoImage per page (กันโดน GC)
        self.page_offsets = []     # y-offset ของแต่ละหน้าใน canvas
        self.page_gap = 30

        self._build_menu()
        self._build_toolbar()
        self._build_canvas()
        self._build_statusbar()
        self._bind_keys()

    def _build_menu(self):
        menubar = tk.Menu(self)

        def add_top(label, cmd):
            m = tk.Menu(menubar, tearoff=0)
            m.add_command(label=label, command=cmd)
            menubar.add_cascade(label=label, menu=m)

        # แต่ละอันเป็นเมนูแยกบนแถบเมนู
        menubar.add_command(label="PDF จากข้อความ", command=self.create_text_pdf)
        menubar.add_command(label="รวมรูปเป็น PDF", command=self.images_to_pdf)
        menubar.add_command(label="Merge PDFs", command=self.merge_pdfs)
        menubar.add_command(label="Split PDF", command=self.split_pdf)
        menubar.add_command(label="Word → PDF", command=self.word_to_pdf)
        menubar.add_command(label="🖨 พิมพ์", command=self.print_pdf)
        menubar.add_command(label="⛶ Full Screen", command=self.toggle_fullscreen)
        self.config(menu=menubar)

    def toggle_fullscreen(self, _event=None):
        is_full = bool(self.attributes("-fullscreen"))
        self.attributes("-fullscreen", not is_full)
        # re-render เพื่อจัดกลางใหม่ตามขนาดใหม่
        self.after(50, self.render)

    def print_pdf(self):
        if not self.doc_path:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        try:
            import os
            os.startfile(self.doc_path, "print")
            self.status.config(text="ส่งไปยังเครื่องพิมพ์แล้ว")
        except Exception as e:
            messagebox.showerror("ผิดพลาด",
                                 f"พิมพ์ไม่ได้: {e}\n\n"
                                 "ต้องมีโปรแกรมเปิด PDF ที่รองรับ Print เช่น Adobe Reader / Edge")

    # ---------- Create PDF features ----------
    def _thai_font(self):
        """หา path ฟอนต์ไทย (Windows)"""
        import os
        candidates = [
            r"C:\Windows\Fonts\tahoma.ttf",
            r"C:\Windows\Fonts\leelawad.ttf",
            r"C:\Windows\Fonts\cordia.ttc",
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
        return None

    def create_text_pdf(self):
        dlg = tk.Toplevel(self)
        dlg.title("สร้าง PDF จากข้อความ")
        dlg.geometry("600x500")
        ttk.Label(dlg, text="พิมพ์ข้อความ (ขึ้นบรรทัดใหม่ได้):").pack(anchor="w", padx=8, pady=4)
        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=6)
        txt = tk.Text(dlg, wrap="word", font=("Tahoma", 12))
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        def save():
            content = txt.get("1.0", "end-1c")
            if not content.strip():
                return
            path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                filetypes=[("PDF", "*.pdf")])
            if not path:
                return
            try:
                doc = fitz.open()
                page = doc.new_page()  # A4
                font_path = self._thai_font()
                rect = fitz.Rect(50, 50, page.rect.width - 50, page.rect.height - 50)
                if font_path:
                    page.insert_textbox(rect, content, fontsize=14,
                                        fontfile=font_path, fontname="thai")
                else:
                    page.insert_textbox(rect, content, fontsize=14)
                doc.save(path)
                doc.close()
                messagebox.showinfo("สำเร็จ", f"บันทึกที่:\n{path}")
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("ผิดพลาด", str(e))

        ttk.Button(btn_frame, text="บันทึกเป็น PDF", command=save).pack(side=tk.RIGHT, padx=10)
        ttk.Button(btn_frame, text="ยกเลิก", command=dlg.destroy).pack(side=tk.RIGHT)

    def images_to_pdf(self):
        paths = filedialog.askopenfilenames(
            title="เลือกรูปภาพ (เรียงตามลำดับที่ต้องการ)",
            filetypes=[("รูปภาพ", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff")])
        if not paths:
            return
        out = filedialog.asksaveasfilename(defaultextension=".pdf",
                                           filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        try:
            doc = fitz.open()
            for p in paths:
                img_doc = fitz.open(p)
                pdfbytes = img_doc.convert_to_pdf()
                img_doc.close()
                img_pdf = fitz.open("pdf", pdfbytes)
                doc.insert_pdf(img_pdf)
                img_pdf.close()
            doc.save(out)
            doc.close()
            messagebox.showinfo("สำเร็จ", f"รวม {len(paths)} รูปเป็น PDF ที่:\n{out}")
        except Exception as e:
            messagebox.showerror("ผิดพลาด", str(e))

    def merge_pdfs(self):
        paths = filedialog.askopenfilenames(
            title="เลือกไฟล์ PDF ที่ต้องการรวม (ตามลำดับ)",
            filetypes=[("PDF", "*.pdf")])
        if not paths or len(paths) < 2:
            messagebox.showwarning("แจ้ง", "เลือกอย่างน้อย 2 ไฟล์")
            return
        out = filedialog.asksaveasfilename(defaultextension=".pdf",
                                           filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        try:
            merged = fitz.open()
            for p in paths:
                src = fitz.open(p)
                merged.insert_pdf(src)
                src.close()
            merged.save(out)
            merged.close()
            messagebox.showinfo("สำเร็จ", f"รวม {len(paths)} ไฟล์แล้ว:\n{out}")
        except Exception as e:
            messagebox.showerror("ผิดพลาด", str(e))

    def word_to_pdf(self):
        paths = filedialog.askopenfilenames(
            title="เลือกไฟล์ Word (.docx)",
            filetypes=[("Word", "*.docx *.doc")])
        if not paths:
            return
        try:
            from docx2pdf import convert
        except Exception as e:
            messagebox.showerror("ผิดพลาด",
                                 f"import docx2pdf ไม่ได้: {e}\n"
                                 "ลองลบไฟล์ .setup_done แล้วเปิดโปรแกรมใหม่")
            return

        out_dir = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึก PDF")
        if not out_dir:
            return
        # fix สำหรับ pythonw: docx2pdf ต้องการ stdout/stderr
        import io
        if sys.stdout is None:
            sys.stdout = io.StringIO()
        if sys.stderr is None:
            sys.stderr = io.StringIO()

        # progress dialog
        prog = tk.Toplevel(self)
        prog.title("กำลังแปลง...")
        prog.geometry("380x120")
        prog.transient(self)
        prog.grab_set()
        prog.resizable(False, False)
        lbl = ttk.Label(prog, text="กำลังเปิด Microsoft Word...", padding=10)
        lbl.pack()
        bar = ttk.Progressbar(prog, mode="indeterminate", length=340)
        bar.pack(padx=15, pady=5)
        bar.start(12)
        sub = ttk.Label(prog, text="โปรดรอสักครู่ อย่าปิดหน้าต่างนี้", foreground="#666")
        sub.pack()
        prog.update()

        import threading
        result = {"error": None, "done": 0}

        def worker():
            try:
                import os
                for i, p in enumerate(paths):
                    out = os.path.join(out_dir,
                                       os.path.splitext(os.path.basename(p))[0] + ".pdf")
                    self.after(0, lbl.config,
                               {"text": f"กำลังแปลง ({i+1}/{len(paths)}): {os.path.basename(p)}"})
                    convert(p, out)
                    result["done"] += 1
            except Exception as e:
                result["error"] = str(e)

        t = threading.Thread(target=worker, daemon=True)
        t.start()

        def check():
            if t.is_alive():
                self.after(100, check)
                return
            bar.stop()
            prog.destroy()
            if result["error"]:
                messagebox.showerror("ผิดพลาด",
                                     f"{result['error']}\n\nต้องมี Microsoft Word ติดตั้งในเครื่อง")
            else:
                messagebox.showinfo("สำเร็จ",
                                    f"แปลง {result['done']} ไฟล์แล้วที่:\n{out_dir}")
        self.after(100, check)

    def split_pdf(self):
        if not self.doc:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        out_dir = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึก")
        if not out_dir:
            return
        try:
            import os
            base = "page"
            for i in range(len(self.doc)):
                new = fitz.open()
                new.insert_pdf(self.doc, from_page=i, to_page=i)
                new.save(os.path.join(out_dir, f"{base}_{i+1:03d}.pdf"))
                new.close()
            messagebox.showinfo("สำเร็จ", f"แยกเป็น {len(self.doc)} ไฟล์ที่:\n{out_dir}")
        except Exception as e:
            messagebox.showerror("ผิดพลาด", str(e))

    # ---------- UI ----------
    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=6)
        bar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(bar, text="เปิดไฟล์", command=self.open_file).pack(side=tk.LEFT)
        ttk.Separator(bar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)

        ttk.Button(bar, text="◀", width=3, command=self.prev_page).pack(side=tk.LEFT)
        self.page_var = tk.StringVar(value="0 / 0")
        self.page_entry = ttk.Entry(bar, width=10, justify="center")
        self.page_entry.pack(side=tk.LEFT, padx=4)
        self.page_entry.bind("<Return>", self._go_to_page)
        ttk.Button(bar, text="▶", width=3, command=self.next_page).pack(side=tk.LEFT)

        ttk.Separator(bar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)
        ttk.Button(bar, text="−", width=3, command=self.zoom_out).pack(side=tk.LEFT)
        self.zoom_lbl = ttk.Label(bar, text="125%", width=6, anchor="center")
        self.zoom_lbl.pack(side=tk.LEFT)
        ttk.Button(bar, text="+", width=3, command=self.zoom_in).pack(side=tk.LEFT)
        ttk.Button(bar, text="Fit", command=self.fit_width).pack(side=tk.LEFT, padx=(4, 0))

        ttk.Separator(bar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)
        self.search_entry = ttk.Entry(bar, width=25)
        self.search_entry.pack(side=tk.LEFT)
        self.search_entry.bind("<Return>", lambda e: self.search())
        ttk.Button(bar, text="ค้นหา", command=self.search).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="ถัดไป", command=self.next_hit).pack(side=tk.LEFT)

    def _build_canvas(self):
        frame = ttk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(frame, bg="#e5e5e5", highlightthickness=0)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.canvas.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self._scroll_target = None
        self._scroll_animating = False

        def _on_wheel(e):
            # หมุนล้อ 1 ครั้ง = เลื่อน 120px แบบ smooth
            step_px = -(e.delta / 120) * 120
            self._smooth_scroll(step_px)
        self.canvas.bind("<MouseWheel>", _on_wheel)
        self.canvas.bind("<Control-MouseWheel>",
                         lambda e: self.zoom_in() if e.delta > 0 else self.zoom_out())
        self.canvas.bind("<Configure>", lambda e: self.render() if self.doc else None)

    def _build_statusbar(self):
        self.status = ttk.Label(self, text="เลือก 'เปิดไฟล์' เพื่อเริ่ม", anchor="w", padding=4)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    def _bind_keys(self):
        self.bind("<Left>", lambda e: self.prev_page())
        self.bind("<Right>", lambda e: self.next_page())
        self.bind("<Control-o>", lambda e: self.open_file())
        self.bind("<Control-p>", lambda e: self.print_pdf())
        self.bind("<F11>", self.toggle_fullscreen)
        self.bind("<Escape>", lambda e: self.attributes("-fullscreen", False) or self.after(50, self.render))
        self.bind("<Control-f>", lambda e: self.search_entry.focus_set())
        self.bind("<Control-plus>", lambda e: self.zoom_in())
        self.bind("<Control-minus>", lambda e: self.zoom_out())

    # ---------- Actions ----------
    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        try:
            self.doc = fitz.open(path)
            self.doc_path = path
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"เปิดไฟล์ไม่ได้:\n{e}")
            return
        self.title(f"PDF Reader — {path}")
        self.page_index = 0
        self.search_hits = []
        self.render()

    def render(self):
        if not self.doc:
            return
        self.canvas.delete("all")
        self.photos = []
        self.page_offsets = []

        cw = self.canvas.winfo_width() or 900
        mat = fitz.Matrix(self.zoom, self.zoom)
        y = 0
        max_w = 0

        for i in range(len(self.doc)):
            pix = self.doc[i].get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            photo = ImageTk.PhotoImage(img)
            self.photos.append(photo)

            offset_x = max((cw - pix.width) // 2, 0)
            self.canvas.create_image(offset_x, y, image=photo, anchor="nw")

            # ไฮไลต์ผลค้นหาบนหน้านี้
            for j, (p, rect) in enumerate(self.search_hits):
                if p != i:
                    continue
                x0 = rect.x0 * self.zoom + offset_x
                y0 = rect.y0 * self.zoom + y
                x1 = rect.x1 * self.zoom + offset_x
                y1 = rect.y1 * self.zoom + y
                color = "#ff9900" if j == self.search_pos else "#ffee00"
                self.canvas.create_rectangle(x0, y0, x1, y1, outline=color, width=2)

            self.page_offsets.append((y, y + pix.height))
            y += pix.height + self.page_gap
            max_w = max(max_w, pix.width)

        scroll_w = max(max_w, cw)
        self.canvas.configure(scrollregion=(0, 0, scroll_w, y))

        total = len(self.doc)
        self.page_var.set(f"{self.page_index + 1} / {total}")
        self.page_entry.delete(0, tk.END)
        self.page_entry.insert(0, str(self.page_index + 1))
        self.zoom_lbl.config(text=f"{int(self.zoom * 100)}%")
        self.status.config(text=f"หน้า {self.page_index + 1} จาก {total}")

    def _smooth_scroll(self, delta_px):
        """เพิ่มระยะทางที่จะเลื่อนไปยังเป้าหมาย แล้วอนิเมทให้ค่อยๆ ถึง"""
        if not self.page_offsets:
            return
        total_h = self.page_offsets[-1][1]
        if total_h <= 0:
            return
        cur_top = self.canvas.yview()[0] * total_h
        if self._scroll_target is None:
            self._scroll_target = cur_top
        self._scroll_target = max(0, min(total_h, self._scroll_target + delta_px))
        if not self._scroll_animating:
            self._scroll_animating = True
            self._animate_scroll()

    def _animate_scroll(self):
        if self._scroll_target is None or not self.page_offsets:
            self._scroll_animating = False
            return
        total_h = self.page_offsets[-1][1]
        cur = self.canvas.yview()[0] * total_h
        diff = self._scroll_target - cur
        if abs(diff) < 1:
            self.canvas.yview_moveto(self._scroll_target / total_h)
            self._scroll_animating = False
            self._scroll_target = None
            self._update_current_page()
            return
        # easing: ขยับ 25% ของระยะที่เหลือทุกเฟรม
        new_cur = cur + diff * 0.25
        self.canvas.yview_moveto(new_cur / total_h)
        self._update_current_page()
        self.after(15, self._animate_scroll)

    def _scroll_to_page(self, index):
        """เลื่อน canvas ไปที่ต้นหน้าที่ index"""
        if not self.page_offsets or index < 0 or index >= len(self.page_offsets):
            return
        top_y = self.page_offsets[index][0]
        total_h = self.page_offsets[-1][1]
        if total_h > 0:
            self.canvas.yview_moveto(top_y / total_h)

    def _update_current_page(self):
        """อัพเดต page_index จาก scroll position ปัจจุบัน"""
        if not self.page_offsets:
            return
        top_frac = self.canvas.yview()[0]
        total_h = self.page_offsets[-1][1]
        y = top_frac * total_h
        for i, (y0, y1) in enumerate(self.page_offsets):
            if y < y1 - 20:
                if self.page_index != i:
                    self.page_index = i
                    total = len(self.doc)
                    self.page_var.set(f"{i + 1} / {total}")
                    self.page_entry.delete(0, tk.END)
                    self.page_entry.insert(0, str(i + 1))
                    self.status.config(text=f"หน้า {i + 1} จาก {total}")
                return

    def next_page(self):
        if self.doc and self.page_index < len(self.doc) - 1:
            self.page_index += 1
            self._scroll_to_page(self.page_index)

    def prev_page(self):
        if self.doc and self.page_index > 0:
            self.page_index -= 1
            self._scroll_to_page(self.page_index)

    def _go_to_page(self, _event=None):
        if not self.doc:
            return
        try:
            n = int(self.page_entry.get()) - 1
            if 0 <= n < len(self.doc):
                self.page_index = n
                self._scroll_to_page(n)
        except ValueError:
            pass

    def zoom_in(self):
        self.zoom = min(self.zoom + 0.25, 6.0)
        self.render()

    def zoom_out(self):
        self.zoom = max(self.zoom - 0.25, 0.25)
        self.render()

    def fit_width(self):
        if not self.doc:
            return
        page = self.doc[self.page_index]
        cw = self.canvas.winfo_width() or 900
        fit_zoom = max(0.25, (cw - 40) / page.rect.width)
        # toggle: ถ้าอยู่ที่ fit อยู่แล้ว → กลับไป zoom เดิม
        if abs(self.zoom - fit_zoom) < 0.01 and getattr(self, "_prev_zoom", None):
            self.zoom = self._prev_zoom
            self._prev_zoom = None
        else:
            self._prev_zoom = self.zoom
            self.zoom = fit_zoom
        self.render()

    def search(self):
        if not self.doc:
            return
        term = self.search_entry.get().strip()
        if not term:
            return
        self.search_hits = []
        for i in range(len(self.doc)):
            for rect in self.doc[i].search_for(term):
                self.search_hits.append((i, rect))
        if not self.search_hits:
            self.status.config(text=f"ไม่พบ '{term}'")
            self.search_pos = -1
            self.render()
            return
        self.search_pos = 0
        self._jump_to_hit()

    def next_hit(self):
        if not self.search_hits:
            return
        self.search_pos = (self.search_pos + 1) % len(self.search_hits)
        self._jump_to_hit()

    def _jump_to_hit(self):
        p, _rect = self.search_hits[self.search_pos]
        self.page_index = p
        self.render()
        self._scroll_to_page(p)
        self.status.config(text=f"ผลลัพธ์ {self.search_pos + 1}/{len(self.search_hits)}")


if __name__ == "__main__":
    import traceback
    try:
        PDFReader().mainloop()
    except Exception:
        traceback.print_exc()
        input("\nกด Enter เพื่อปิด...")
