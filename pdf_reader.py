"""
PDF Reader - Desktop App
Developed by Thanagrid C. <Thanagrid.c@ku.th>
Copyright (c) 2026
Licensed under the MIT License — see LICENSE file for details.

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
    open(_MARKER, "w").close()
    print("[Setup] เสร็จสิ้น — ครั้งต่อไปจะไม่ถามอีก")

_ensure_deps()

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
try:
    import pymupdf as fitz
except ImportError:
    import fitz
from PIL import Image, ImageTk


# =========================================================
#  PDFTab — แต่ละแท็บมี canvas + doc + state ของตัวเอง
# =========================================================
class PDFTab(ttk.Frame):
    def __init__(self, parent, app, path):
        super().__init__(parent)
        self.app = app
        self.doc = fitz.open(path)
        self.doc_path = path
        self.page_index = 0
        self.zoom = 1.25
        self._prev_zoom = None
        self.search_hits = []
        self.search_pos = -1
        self.photos = []
        self.page_offsets = []
        self.page_gap = 30
        self._scroll_target = None
        self._scroll_animating = False

        # signature
        self.sign_mode = False
        self.current_stroke = None
        self.strokes = []

        # dirty flag — ตั้ง True เมื่อมีการแก้ไข
        self.dirty = False

        # undo/redo stacks — เก็บ snapshot ของ doc ก่อนแก้ไข
        self.undo_stack = []
        self.redo_stack = []
        self.max_history = 20

        # ===== Thumbnail sidebar (ซ้าย) =====
        self.thumb_width = 140
        self.thumb_photos = []
        self.thumb_items = []   # (canvas_id, page_index, y_top, y_bottom)
        self._thumb_current = None

        thumb_frame = tk.Frame(self, bg="#2b2b2b", width=self.thumb_width + 20)
        thumb_frame.pack(side=tk.LEFT, fill=tk.Y)
        thumb_frame.pack_propagate(False)
        self.thumb_frame = thumb_frame
        self.thumb_visible = True

        self.thumb_canvas = tk.Canvas(thumb_frame, bg="#2b2b2b",
                                      highlightthickness=0, width=self.thumb_width + 4)
        thumb_vsb = ttk.Scrollbar(thumb_frame, orient="vertical",
                                  command=self.thumb_canvas.yview)
        self.thumb_canvas.configure(yscrollcommand=thumb_vsb.set)
        thumb_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.thumb_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.thumb_canvas.bind("<MouseWheel>",
                               lambda e: self.thumb_canvas.yview_scroll(int(-e.delta / 120), "units"))
        self.thumb_canvas.bind("<Button-1>", self._on_thumb_click)

        # ===== Main canvas + scrollbars (ขวา) =====
        main_frame = ttk.Frame(self)
        main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(main_frame, bg="#e5e5e5", highlightthickness=0)
        vsb = ttk.Scrollbar(main_frame, orient="vertical", command=self.canvas.yview)
        hsb = ttk.Scrollbar(main_frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        def _on_wheel(e):
            self._smooth_scroll(-(e.delta / 120) * 120)
        self.canvas.bind("<MouseWheel>", _on_wheel)
        self.canvas.bind("<Control-MouseWheel>",
                         lambda e: self.zoom_in() if e.delta > 0 else self.zoom_out())
        self.canvas.bind("<Configure>", lambda e: self.render() if self.doc else None)

        # text selection (เมื่อไม่มี mode ใดทำงาน)
        self._sel_start = None       # (page, char_idx)
        self._sel_end = None
        self._sel_highlight_ids = []
        self._sel_text = ""
        self._chars_cache = {}
        self._words_cache = {}       # legacy (ยังเผื่อไว้)
        self._bind_selection()

        self.after(80, self.render_thumbnails)

    def _bind_selection(self):
        self.canvas.bind("<ButtonPress-1>", self._sel_press)
        self.canvas.bind("<B1-Motion>", self._sel_motion)
        self.canvas.bind("<ButtonRelease-1>", self._sel_release)
        self.canvas.bind("<Button-3>", self._sel_context_menu)
        self.canvas.config(cursor="xterm")

    def _any_mode_active(self):
        return any(getattr(self, m, False) for m in
                   ("sign_mode", "image_mode", "hl_mode",
                    "draw_mode", "cmt_mode", "redact_mode"))

    def _get_chars(self, page_idx):
        """คืน list [(bbox, char, block, line), ...] เรียงตามลำดับการอ่าน"""
        if page_idx in self._chars_cache:
            return self._chars_cache[page_idx]
        chars = []
        try:
            d = self.doc[page_idx].get_text("rawdict")
        except Exception:
            d = {"blocks": []}
        for b_idx, block in enumerate(d.get("blocks", [])):
            if block.get("type", 1) != 0:  # ข้ามรูป
                continue
            for l_idx, line in enumerate(block.get("lines", [])):
                for span in line.get("spans", []):
                    for ch in span.get("chars", []):
                        bbox = ch.get("bbox")
                        c = ch.get("c", "")
                        if bbox and c:
                            chars.append((bbox, c, b_idx, l_idx))
        self._chars_cache[page_idx] = chars
        return chars

    def _char_index_at(self, page_idx, x, y):
        chars = self._get_chars(page_idx)
        if not chars:
            return None
        for i, (bbox, _, _, _) in enumerate(chars):
            x0, y0, x1, y1 = bbox
            if x0 <= x <= x1 and y0 <= y <= y1:
                return i
        # หาบรรทัดที่ครอบ y ก่อน แล้วเลือกตัวใกล้ x
        line_cands = [i for i, (bb, _, _, _) in enumerate(chars) if bb[1] <= y <= bb[3]]
        if line_cands:
            return min(line_cands, key=lambda i: abs((chars[i][0][0] + chars[i][0][2]) / 2 - x))
        best, best_d = None, float("inf")
        for i, (bbox, _, _, _) in enumerate(chars):
            cx = (bbox[0] + bbox[2]) / 2
            cy = (bbox[1] + bbox[3]) / 2
            d = (x - cx) ** 2 + ((y - cy) * 3) ** 2
            if d < best_d:
                best_d, best = d, i
        return best

    def _clear_sel_highlight(self):
        for hid in self._sel_highlight_ids:
            self.canvas.delete(hid)
        self._sel_highlight_ids = []

    def _draw_sel_highlight(self, page_idx, i0, i1):
        """วาดสีน้ำเงินคลุมช่วง char index i0..i1 — รวมช่วงในบรรทัดเดียวเป็น rect เดียว"""
        self._clear_sel_highlight()
        chars = self._get_chars(page_idx)
        if not chars or page_idx >= len(self.page_offsets):
            return
        lo, hi = min(i0, i1), max(i0, i1)
        page_y_top, _ = self.page_offsets[page_idx]
        cw = self.canvas.winfo_width() or 900
        page_w_px = self.doc[page_idx].rect.width * self.zoom
        offset_x = max((cw - page_w_px) // 2, 0)
        # รวมเป็นช่วงตามบรรทัด
        run = None  # [x0,y0,x1,y1]
        prev_line = None
        for i in range(lo, hi + 1):
            bbox, _, b, l = chars[i]
            key = (b, l)
            if run and key == prev_line:
                run[0] = min(run[0], bbox[0])
                run[1] = min(run[1], bbox[1])
                run[2] = max(run[2], bbox[2])
                run[3] = max(run[3], bbox[3])
            else:
                if run:
                    self._flush_sel_rect(run, page_y_top, offset_x)
                run = list(bbox)
                prev_line = key
        if run:
            self._flush_sel_rect(run, page_y_top, offset_x)

    def _flush_sel_rect(self, r, page_y_top, offset_x):
        cx0 = r[0] * self.zoom + offset_x
        cy0 = r[1] * self.zoom + page_y_top
        cx1 = r[2] * self.zoom + offset_x
        cy1 = r[3] * self.zoom + page_y_top
        hid = self.canvas.create_rectangle(
            cx0, cy0, cx1, cy1, outline="",
            fill="#3a7bd5", stipple="gray50")
        self._sel_highlight_ids.append(hid)

    def _selected_text(self, page_idx, i0, i1):
        chars = self._get_chars(page_idx)
        if not chars:
            return ""
        lo, hi = min(i0, i1), max(i0, i1)
        parts = []
        prev_line = None
        for i in range(lo, hi + 1):
            _, c, b, l = chars[i]
            key = (b, l)
            if prev_line is not None and key != prev_line:
                parts.append("\n")
            parts.append(c)
            prev_line = key
        return "".join(parts).strip()

    def _sel_press(self, e):
        self._clear_sel_highlight()
        if self._any_mode_active():
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        info = self._canvas_to_pdf(cx, cy)
        if info is None:
            self._sel_start = None
            return
        p, px, py = info
        idx = self._char_index_at(p, px, py)
        if idx is None:
            self._sel_start = None
            return
        self._sel_start = (p, idx)
        self._sel_end = (p, idx)

    def _sel_motion(self, e):
        if self._any_mode_active() or not self._sel_start:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        info = self._canvas_to_pdf(cx, cy)
        if not info:
            return
        p, px, py = info
        if p != self._sel_start[0]:
            return
        idx = self._char_index_at(p, px, py)
        if idx is None:
            return
        self._sel_end = (p, idx)
        self._draw_sel_highlight(p, self._sel_start[1], idx)

    def _sel_release(self, e):
        if self._any_mode_active() or not self._sel_start or not self._sel_end:
            return
        p, i0 = self._sel_start
        _, i1 = self._sel_end
        text = self._selected_text(p, i0, i1)
        self._sel_text = text
        if text:
            try:
                self.clipboard_clear()
                self.clipboard_append(text)
                self.app.status.config(
                    text=f"คัดลอกแล้ว ({len(text)} ตัวอักษร) — Ctrl+V วางที่อื่นได้")
            except Exception:
                pass

    def _sel_context_menu(self, e):
        if self._any_mode_active():
            return
        m = tk.Menu(self.canvas, tearoff=0)
        has_sel = bool(self._sel_text)
        m.add_command(label="📋 คัดลอก", state=("normal" if has_sel else "disabled"),
                      command=self._copy_selection)
        m.add_command(label="เลือกทั้งหน้า", command=self._select_current_page)
        m.add_separator()
        m.add_command(label="ล้างการเลือก", command=self._clear_selection)
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()

    def _copy_selection(self):
        if not self._sel_text:
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(self._sel_text)
            self.app.status.config(text=f"คัดลอกแล้ว ({len(self._sel_text)} ตัวอักษร)")
        except Exception:
            pass

    def _select_current_page(self):
        p = self.page_index
        chars = self._get_chars(p)
        if not chars:
            return
        self._sel_start = (p, 0)
        self._sel_end = (p, len(chars) - 1)
        self._draw_sel_highlight(p, 0, len(chars) - 1)
        self._sel_text = self._selected_text(p, 0, len(chars) - 1)
        self._copy_selection()

    def _clear_selection(self):
        self._clear_sel_highlight()
        self._sel_start = None
        self._sel_end = None
        self._sel_text = ""

    # ---------- Rendering ----------
    def render(self):
        if not self.doc:
            return
        self.canvas.delete("all")
        self.photos = []
        self.page_offsets = []
        self._sel_highlight_ids = []

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

        self.canvas.configure(scrollregion=(0, 0, max(max_w, cw), y))
        self.app._sync_toolbar()

    # ---------- Undo / Redo ----------
    def _snapshot(self):
        """เก็บ state ปัจจุบันเพื่อ undo ก่อนแก้ไข doc"""
        try:
            self.undo_stack.append(self.doc.tobytes(garbage=0, deflate=False))
            if len(self.undo_stack) > self.max_history:
                self.undo_stack.pop(0)
            self.redo_stack.clear()
        except Exception:
            pass

    def _reload_from_bytes(self, data):
        cur_page = self.page_index
        try:
            self.doc.close()
        except Exception:
            pass
        self.doc = fitz.open(stream=data, filetype="pdf")
        self.page_index = min(cur_page, len(self.doc) - 1)
        self._words_cache = {}
        self._chars_cache = {}
        self.render()
        self.render_thumbnails()

    def undo(self):
        if not self.undo_stack:
            return False
        try:
            current = self.doc.tobytes(garbage=0, deflate=False)
            self.redo_stack.append(current)
            data = self.undo_stack.pop()
            self._reload_from_bytes(data)
            self.dirty = bool(self.undo_stack)
            return True
        except Exception:
            return False

    def redo(self):
        if not self.redo_stack:
            return False
        try:
            current = self.doc.tobytes(garbage=0, deflate=False)
            self.undo_stack.append(current)
            data = self.redo_stack.pop()
            self._reload_from_bytes(data)
            self.dirty = True
            return True
        except Exception:
            return False

    # ---------- Rotation ----------
    def rotate_current(self, delta):
        if not self.doc:
            return
        self._snapshot()
        page = self.doc[self.page_index]
        page.set_rotation((page.rotation + delta) % 360)
        self._words_cache.pop(self.page_index, None)
        self._chars_cache.pop(self.page_index, None)
        self.dirty = True
        self.render()
        self.render_thumbnails()

    def rotate_all(self, delta):
        if not self.doc:
            return
        self._snapshot()
        for p in self.doc:
            p.set_rotation((p.rotation + delta) % 360)
        self._words_cache = {}
        self._chars_cache = {}
        self.dirty = True
        self.render()
        self.render_thumbnails()

    # ---------- Scrolling ----------
    def _smooth_scroll(self, delta_px):
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
        new_cur = cur + diff * 0.25
        self.canvas.yview_moveto(new_cur / total_h)
        self._update_current_page()
        self.after(15, self._animate_scroll)

    def _scroll_to_page(self, index):
        if not self.page_offsets or index < 0 or index >= len(self.page_offsets):
            return
        top_y = self.page_offsets[index][0]
        try:
            scroll_h = float(self.canvas.cget("scrollregion").split()[-1])
        except (ValueError, IndexError):
            scroll_h = self.page_offsets[-1][1]
        if scroll_h > 0:
            self.canvas.yview_moveto(top_y / scroll_h)
        self._scroll_target = None
        self._scroll_animating = False
        self._highlight_thumb(index)

    def _update_current_page(self):
        if not self.page_offsets:
            return
        top_frac = self.canvas.yview()[0]
        total_h = self.page_offsets[-1][1]
        y = top_frac * total_h
        for i, (y0, y1) in enumerate(self.page_offsets):
            if y < y1 - 20:
                if self.page_index != i:
                    self.page_index = i
                    self.app._sync_toolbar()
                self._highlight_thumb(i)
                return

    # ---------- Thumbnails ----------
    def render_thumbnails(self):
        if not self.doc:
            return
        self.thumb_canvas.delete("all")
        self.thumb_photos = []
        self.thumb_items = []

        pad_x = 10
        pad_y = 8
        y = pad_y
        max_w = self.thumb_width

        for i in range(len(self.doc)):
            page = self.doc[i]
            r = page.rect
            scale = self.thumb_width / max(r.width, 1)
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            photo = ImageTk.PhotoImage(img)
            self.thumb_photos.append(photo)

            rect_id = self.thumb_canvas.create_rectangle(
                pad_x - 2, y - 2, pad_x + pix.width + 2, y + pix.height + 2,
                outline="", width=2, tags=(f"thumb{i}",))
            self.thumb_canvas.create_image(pad_x, y, image=photo, anchor="nw",
                                           tags=(f"thumb{i}",))
            self.thumb_canvas.create_text(
                pad_x + pix.width / 2, y + pix.height + 10,
                text=str(i + 1), fill="#cccccc", font=("Segoe UI", 9),
                tags=(f"thumb{i}",))

            self.thumb_items.append((rect_id, i, y - 2, y + pix.height + 2))
            y += pix.height + 26
            max_w = max(max_w, pix.width + 2 * pad_x)

        self.thumb_canvas.configure(scrollregion=(0, 0, max_w, y))
        self._highlight_thumb(self.page_index)

    def _highlight_thumb(self, page_idx):
        if page_idx == self._thumb_current:
            return
        self._thumb_current = page_idx
        for rect_id, i, _, _ in self.thumb_items:
            color = "#4a90e2" if i == page_idx else ""
            self.thumb_canvas.itemconfig(rect_id, outline=color)
        # เลื่อน thumbnail ให้เห็นหน้าปัจจุบัน
        for rect_id, i, y_top, y_bot in self.thumb_items:
            if i == page_idx:
                total = float(self.thumb_canvas.cget("scrollregion").split()[-1] or 1)
                view = self.thumb_canvas.yview()
                view_h = (view[1] - view[0]) * total
                if y_top < view[0] * total or y_bot > view[0] * total + view_h:
                    self.thumb_canvas.yview_moveto(max(0, (y_top - 10) / total))
                break

    def toggle_thumbnails(self):
        if self.thumb_visible:
            self.thumb_frame.pack_forget()
            self.thumb_visible = False
        else:
            self.thumb_frame.pack(side=tk.LEFT, fill=tk.Y, before=self.canvas.master)
            self.thumb_visible = True
        self.render()

    def _on_thumb_click(self, event):
        y = self.thumb_canvas.canvasy(event.y)
        for _, i, y_top, y_bot in self.thumb_items:
            if y_top <= y <= y_bot:
                self.go_to_page(i)
                return

    # ---------- Navigation ----------
    def next_page(self):
        if self.page_index < len(self.doc) - 1:
            self.page_index += 1
            self._scroll_to_page(self.page_index)
            self.app._sync_toolbar()

    def prev_page(self):
        if self.page_index > 0:
            self.page_index -= 1
            self._scroll_to_page(self.page_index)
            self.app._sync_toolbar()

    def go_to_page(self, n):
        if 0 <= n < len(self.doc):
            self.page_index = n
            self._scroll_to_page(n)
            self.app._sync_toolbar()

    # ---------- Zoom ----------
    def zoom_in(self):
        self.zoom = min(self.zoom + 0.25, 6.0)
        self.render()

    def zoom_out(self):
        self.zoom = max(self.zoom - 0.25, 0.25)
        self.render()

    def fit_width(self):
        page = self.doc[self.page_index]
        cw = self.canvas.winfo_width() or 900
        fit_zoom = max(0.25, (cw - 40) / page.rect.width)
        if abs(self.zoom - fit_zoom) < 0.01 and self._prev_zoom:
            self.zoom = self._prev_zoom
            self._prev_zoom = None
        else:
            self._prev_zoom = self.zoom
            self.zoom = fit_zoom
        self.render()

    # ---------- Search ----------
    def search(self, term):
        self.search_hits = []
        for i in range(len(self.doc)):
            for rect in self.doc[i].search_for(term):
                self.search_hits.append((i, rect))
        if not self.search_hits:
            self.search_pos = -1
            self.render()
            return False
        self.search_pos = 0
        self._jump_to_hit()
        return True

    def next_hit(self):
        if not self.search_hits:
            return
        self.search_pos = (self.search_pos + 1) % len(self.search_hits)
        self._jump_to_hit()

    def _jump_to_hit(self):
        p, _ = self.search_hits[self.search_pos]
        self.page_index = p
        self.render()
        self._scroll_to_page(p)
        self.app._sync_toolbar()

    # ---------- Signature ----------
    def toggle_sign_mode(self):
        self.sign_mode = not self.sign_mode
        if self.sign_mode:
            self.canvas.config(cursor="crosshair")
            self.current_stroke = None
            self.strokes = []
            # เฟส 1: ยังไม่ได้กรอบ → ต้องลากกรอบก่อน
            # เฟส 2: มีกรอบแล้ว → เซ็นในกรอบได้
            self.sign_phase = "frame"
            self.sign_frame = None  # (page_idx, x0, y0, x1, y1) — PDF coords
            self._frame_start = None
            self._frame_rect_id = None
            self.canvas.bind("<ButtonPress-1>", self._sign_press)
            self.canvas.bind("<B1-Motion>", self._sign_motion)
            self.canvas.bind("<ButtonRelease-1>", self._sign_release)
        else:
            self.sign_cleanup()

    def sign_cleanup(self):
        self.sign_mode = False
        self.sign_phase = None
        self.sign_frame = None
        self.canvas.unbind("<ButtonPress-1>")
        self.canvas.unbind("<B1-Motion>")
        self.canvas.unbind("<ButtonRelease-1>")
        self._bind_selection()
        self.render()

    # ---------- Insert image ----------
    def start_image_mode(self):
        self.image_mode = True
        self.canvas.config(cursor="crosshair")
        self._img_start = None
        self._img_rect_id = None
        self.canvas.bind("<ButtonPress-1>", self._img_press)
        self.canvas.bind("<B1-Motion>", self._img_motion)
        self.canvas.bind("<ButtonRelease-1>", self._img_release)

    def _img_cleanup(self):
        self.image_mode = False
        self.canvas.unbind("<ButtonPress-1>")
        self.canvas.unbind("<B1-Motion>")
        self.canvas.unbind("<ButtonRelease-1>")
        self._bind_selection()
        if self._img_rect_id:
            self.canvas.delete(self._img_rect_id)
            self._img_rect_id = None

    def _img_press(self, e):
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        info = self._canvas_to_pdf(cx, cy)
        if info is None:
            return
        p, px, py = info
        self._img_start = (p, px, py, cx, cy)

    def _img_motion(self, e):
        if not self._img_start:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        _, _, _, scx, scy = self._img_start
        if self._img_rect_id:
            self.canvas.delete(self._img_rect_id)
        self._img_rect_id = self.canvas.create_rectangle(
            scx, scy, cx, cy, outline="#009933", width=2, dash=(6, 4))

    # ---------- Highlight ----------
    def toggle_highlight_mode(self):
        self.hl_mode = not getattr(self, "hl_mode", False)
        if self.hl_mode:
            self._exit_other_modes(keep="hl")
            self.canvas.config(cursor="crosshair")
            self._hl_start = None
            self._hl_rect_id = None
            self.canvas.bind("<ButtonPress-1>", self._hl_press)
            self.canvas.bind("<B1-Motion>", self._hl_motion)
            self.canvas.bind("<ButtonRelease-1>", self._hl_release)
        else:
            self._hl_cleanup()

    def _hl_cleanup(self):
        self.hl_mode = False
        self.canvas.unbind("<ButtonPress-1>")
        self.canvas.unbind("<B1-Motion>")
        self.canvas.unbind("<ButtonRelease-1>")
        self._bind_selection()

    def _hl_press(self, e):
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        info = self._canvas_to_pdf(cx, cy)
        if info is None:
            return
        p, px, py = info
        self._hl_start = (p, px, py, cx, cy)

    def _hl_motion(self, e):
        if not self._hl_start:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        _, _, _, scx, scy = self._hl_start
        if self._hl_rect_id:
            self.canvas.delete(self._hl_rect_id)
        self._hl_rect_id = self.canvas.create_rectangle(
            scx, scy, cx, cy, outline="", fill="#fff275", stipple="gray50")

    def _hl_release(self, e):
        if not self._hl_start:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        p, sx, sy, _, _ = self._hl_start
        self._hl_start = None
        if self._hl_rect_id:
            self.canvas.delete(self._hl_rect_id)
            self._hl_rect_id = None
        info = self._canvas_to_pdf(cx, cy)
        if not info or info[0] != p:
            return
        _, ex, ey = info
        x0, x1 = min(sx, ex), max(sx, ex)
        y0, y1 = min(sy, ey), max(sy, ey)
        if (x1 - x0) < 5 or (y1 - y0) < 5:
            return
        self._snapshot()
        page = self.doc[p]
        page.draw_rect(fitz.Rect(x0, y0, x1, y1),
                       color=None, fill=(1, 0.95, 0.3), fill_opacity=0.4,
                       overlay=True)
        self.dirty = True
        self.render()

    # ---------- Redact (ลบข้อความ) ----------
    def toggle_redact_mode(self):
        self.redact_mode = not getattr(self, "redact_mode", False)
        if self.redact_mode:
            self._exit_other_modes(keep="redact")
            self.canvas.config(cursor="crosshair")
            self._rd_start = None
            self._rd_rect_id = None
            self.canvas.bind("<ButtonPress-1>", self._rd_press)
            self.canvas.bind("<B1-Motion>", self._rd_motion)
            self.canvas.bind("<ButtonRelease-1>", self._rd_release)
        else:
            self._rd_cleanup()

    def _rd_cleanup(self):
        self.redact_mode = False
        self.canvas.unbind("<ButtonPress-1>")
        self.canvas.unbind("<B1-Motion>")
        self.canvas.unbind("<ButtonRelease-1>")
        self._bind_selection()

    def _rd_press(self, e):
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        info = self._canvas_to_pdf(cx, cy)
        if info is None:
            return
        p, px, py = info
        self._rd_start = (p, px, py, cx, cy)

    def _rd_motion(self, e):
        if not self._rd_start:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        _, _, _, scx, scy = self._rd_start
        if self._rd_rect_id:
            self.canvas.delete(self._rd_rect_id)
        self._rd_rect_id = self.canvas.create_rectangle(
            scx, scy, cx, cy, outline="#c00000", width=2, fill="#000000", stipple="gray50")

    def _rd_release(self, e):
        if not self._rd_start:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        p, sx, sy, _, _ = self._rd_start
        self._rd_start = None
        if self._rd_rect_id:
            self.canvas.delete(self._rd_rect_id)
            self._rd_rect_id = None
        info = self._canvas_to_pdf(cx, cy)
        if not info or info[0] != p:
            return
        _, ex, ey = info
        x0, x1 = min(sx, ex), max(sx, ex)
        y0, y1 = min(sy, ey), max(sy, ey)
        if (x1 - x0) < 3 or (y1 - y0) < 3:
            return
        self._snapshot()
        page = self.doc[p]
        # ใส่ redact annotation แล้ว apply เพื่อลบข้อความในบริเวณนั้น
        page.add_redact_annot(fitz.Rect(x0, y0, x1, y1), fill=(1, 1, 1))
        page.apply_redactions()
        self.dirty = True
        self.render()
        self.render_thumbnails()

    # ---------- Draw (freehand) ----------
    def toggle_draw_mode(self):
        self.draw_mode = not getattr(self, "draw_mode", False)
        if self.draw_mode:
            self._exit_other_modes(keep="draw")
            self.canvas.config(cursor="pencil")
            self._draw_current = None
            self._draw_strokes = []
            self.canvas.bind("<ButtonPress-1>", self._draw_press)
            self.canvas.bind("<B1-Motion>", self._draw_motion)
            self.canvas.bind("<ButtonRelease-1>", self._draw_release)
        else:
            self._draw_commit()

    def _draw_commit(self):
        # เขียนเส้นลง doc
        if getattr(self, "_draw_strokes", None):
            self._snapshot()
            for page_idx, pts in self._draw_strokes:
                page = self.doc[page_idx]
                for i in range(len(pts) - 1):
                    p1 = fitz.Point(*pts[i])
                    p2 = fitz.Point(*pts[i+1])
                    page.draw_line(p1, p2, color=(0.9, 0.1, 0.1), width=1.8)
            self._draw_strokes = []
            self.dirty = True
        self.draw_mode = False
        self.canvas.unbind("<ButtonPress-1>")
        self.canvas.unbind("<B1-Motion>")
        self.canvas.unbind("<ButtonRelease-1>")
        self._bind_selection()
        self.render()

    def _draw_press(self, e):
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        info = self._canvas_to_pdf(cx, cy)
        if info is None:
            return
        p, px, py = info
        self._draw_current = {"page": p, "canvas": [(cx, cy)], "pdf": [(px, py)]}

    def _draw_motion(self, e):
        if not self._draw_current:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        info = self._canvas_to_pdf(cx, cy)
        if info is None or info[0] != self._draw_current["page"]:
            return
        p, px, py = info
        last_cx, last_cy = self._draw_current["canvas"][-1]
        self.canvas.create_line(last_cx, last_cy, cx, cy,
                                fill="#e60000", width=2, capstyle="round", smooth=True)
        self._draw_current["canvas"].append((cx, cy))
        self._draw_current["pdf"].append((px, py))

    def _draw_release(self, _e):
        if self._draw_current and len(self._draw_current["pdf"]) > 1:
            self._draw_strokes.append(
                (self._draw_current["page"], self._draw_current["pdf"]))
        self._draw_current = None

    # ---------- Comment ----------
    def toggle_comment_mode(self):
        self.cmt_mode = not getattr(self, "cmt_mode", False)
        if self.cmt_mode:
            self._exit_other_modes(keep="cmt")
            self.canvas.config(cursor="question_arrow")
            self.canvas.bind("<ButtonPress-1>", self._cmt_click)
        else:
            self._cmt_cleanup()

    def _cmt_cleanup(self):
        self.cmt_mode = False
        self.canvas.unbind("<ButtonPress-1>")
        self._bind_selection()

    def _cmt_click(self, e):
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        info = self._canvas_to_pdf(cx, cy)
        if info is None:
            return
        p, px, py = info

        # popup ให้พิมพ์ข้อความ
        dlg = tk.Toplevel(self)
        dlg.title("เพิ่มคอมเมนต์")
        dlg.geometry("400x200")
        dlg.transient(self.winfo_toplevel())
        ttk.Label(dlg, text="พิมพ์คอมเมนต์:").pack(anchor="w", padx=8, pady=4)
        txt = tk.Text(dlg, wrap="word", font=("Tahoma", 11), height=6)
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        txt.focus_set()

        def ok():
            content = txt.get("1.0", "end-1c").strip()
            if content:
                self._snapshot()
                page = self.doc[p]
                annot = page.add_text_annot(fitz.Point(px, py), content)
                annot.set_info(title="Note")
                annot.update()
                self.dirty = True
            dlg.destroy()
            self.render()

        bf = ttk.Frame(dlg); bf.pack(side=tk.BOTTOM, fill=tk.X, pady=6)
        ttk.Button(bf, text="ตกลง", command=ok).pack(side=tk.RIGHT, padx=8)
        ttk.Button(bf, text="ยกเลิก", command=dlg.destroy).pack(side=tk.RIGHT)

    def _exit_other_modes(self, keep):
        """ปิด mode อื่นๆ ก่อนเปิด mode ใหม่"""
        if keep != "sign" and getattr(self, "sign_mode", False):
            self.sign_cleanup()
        if keep != "img" and getattr(self, "image_mode", False):
            self._img_cleanup()
        if keep != "hl" and getattr(self, "hl_mode", False):
            self._hl_cleanup()
        if keep != "draw" and getattr(self, "draw_mode", False):
            self._draw_commit()
        if keep != "cmt" and getattr(self, "cmt_mode", False):
            self._cmt_cleanup()
        if keep != "redact" and getattr(self, "redact_mode", False):
            self._rd_cleanup()

    def _img_release(self, e):
        if not self._img_start:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        p, sx, sy, _, _ = self._img_start
        info = self._canvas_to_pdf(cx, cy)
        self._img_start = None
        if not info or info[0] != p:
            self._img_cleanup()
            return
        _, ex, ey = info
        x0, x1 = min(sx, ex), max(sx, ex)
        y0, y1 = min(sy, ey), max(sy, ey)
        if (x1 - x0) < 20 or (y1 - y0) < 20:
            self._img_cleanup()
            self.app.status.config(text="กรอบเล็กเกินไป ลองใหม่")
            return

        img_path = filedialog.askopenfilename(
            title="เลือกรูปภาพ",
            filetypes=[("รูปภาพ", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff")])
        if not img_path:
            self._img_cleanup()
            return

        out = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialfile="with_image.pdf")
        if not out:
            self._img_cleanup()
            return

        try:
            page = self.doc[p]
            rect = fitz.Rect(x0, y0, x1, y1)
            page.insert_image(rect, filename=img_path)
            self.dirty = True
            self.doc.save(out, garbage=3, deflate=True)
            self.dirty = False
            messagebox.showinfo("สำเร็จ", f"บันทึกแล้วที่:\n{out}")
        except Exception as e:
            messagebox.showerror("ผิดพลาด", str(e))
        finally:
            self._img_cleanup()
            self.render()

    def _canvas_to_pdf(self, cx, cy):
        if not self.page_offsets:
            return None
        for i, (y0, y1) in enumerate(self.page_offsets):
            if y0 <= cy <= y1:
                cw = self.canvas.winfo_width() or 900
                page_w = self.doc[i].rect.width * self.zoom
                offset_x = max((cw - page_w) // 2, 0)
                return (i, (cx - offset_x) / self.zoom, (cy - y0) / self.zoom)
        return None

    def _pdf_to_canvas(self, page_idx, px, py):
        """แปลง PDF coord กลับเป็น canvas coord"""
        y0 = self.page_offsets[page_idx][0]
        cw = self.canvas.winfo_width() or 900
        page_w = self.doc[page_idx].rect.width * self.zoom
        offset_x = max((cw - page_w) // 2, 0)
        return (px * self.zoom + offset_x, py * self.zoom + y0)

    def _in_frame(self, page_idx, px, py):
        """เช็คว่าจุด (px,py) อยู่ในกรอบไหม"""
        if not self.sign_frame:
            return False
        fp, x0, y0, x1, y1 = self.sign_frame
        return fp == page_idx and x0 <= px <= x1 and y0 <= py <= y1

    def _sign_press(self, e):
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        info = self._canvas_to_pdf(cx, cy)
        if info is None:
            return
        p, px, py = info

        if self.sign_phase == "frame":
            # เริ่มลากกรอบ
            self._frame_start = (p, px, py, cx, cy)
        else:
            # เฟส draw: ต้องอยู่ในกรอบ
            if not self._in_frame(p, px, py):
                self.app.status.config(text="⚠ ต้องเซ็นในกรอบเท่านั้น")
                return
            self.current_stroke = {"page": p, "canvas": [(cx, cy)], "pdf": [(px, py)]}

    def _sign_motion(self, e):
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)

        if self.sign_phase == "frame" and self._frame_start:
            # แสดง preview กรอบขณะลาก
            p, sx, sy, scx, scy = self._frame_start
            if self._frame_rect_id:
                self.canvas.delete(self._frame_rect_id)
            self._frame_rect_id = self.canvas.create_rectangle(
                scx, scy, cx, cy, outline="#cc0000", width=2, dash=(6, 4),
                tags="frame_preview")
            return

        if self.sign_phase == "draw" and self.current_stroke:
            info = self._canvas_to_pdf(cx, cy)
            if info is None or info[0] != self.current_stroke["page"]:
                return
            p, px, py = info
            if not self._in_frame(p, px, py):
                return  # ออกนอกกรอบ → หยุดวาด
            last_cx, last_cy = self.current_stroke["canvas"][-1]
            self.canvas.create_line(last_cx, last_cy, cx, cy,
                                    fill="#0033cc", width=2, capstyle="round",
                                    smooth=True, tags="signature")
            self.current_stroke["canvas"].append((cx, cy))
            self.current_stroke["pdf"].append((px, py))

    def _sign_release(self, e):
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)

        if self.sign_phase == "frame" and self._frame_start:
            p, sx, sy, _, _ = self._frame_start
            info = self._canvas_to_pdf(cx, cy)
            if info and info[0] == p:
                _, ex, ey = info
                x0, x1 = min(sx, ex), max(sx, ex)
                y0, y1 = min(sy, ey), max(sy, ey)
                if (x1 - x0) > 20 and (y1 - y0) > 10:  # กรอบต้องใหญ่พอ
                    self.sign_frame = (p, x0, y0, x1, y1)
                    self.sign_phase = "draw"
                    self.canvas.config(cursor="pencil")
                    # วาดกรอบถาวร
                    if self._frame_rect_id:
                        self.canvas.delete(self._frame_rect_id)
                    scx0, scy0 = self._pdf_to_canvas(p, x0, y0)
                    scx1, scy1 = self._pdf_to_canvas(p, x1, y1)
                    self._frame_rect_id = self.canvas.create_rectangle(
                        scx0, scy0, scx1, scy1, outline="#cc0000", width=1.5,
                        tags="sign_frame")
                    self.app.status.config(
                        text="✓ กำหนดกรอบแล้ว — เซ็นในกรอบ | Enter=บันทึก | Esc=ยกเลิก")
                else:
                    if self._frame_rect_id:
                        self.canvas.delete(self._frame_rect_id)
                        self._frame_rect_id = None
                    self.app.status.config(text="กรอบเล็กเกินไป ลองใหม่")
            self._frame_start = None
            return

        if self.sign_phase == "draw":
            if self.current_stroke and len(self.current_stroke["pdf"]) > 1:
                self.strokes.append((self.current_stroke["page"], self.current_stroke["pdf"]))
            self.current_stroke = None

    def sign_save(self):
        if not self.strokes:
            messagebox.showinfo("แจ้ง", "ยังไม่มีลายเซ็น")
            return
        out = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialfile="signed.pdf")
        if not out:
            return
        try:
            # วาดเส้นลายเซ็น
            for page_idx, pts in self.strokes:
                page = self.doc[page_idx]
                for i in range(len(pts) - 1):
                    p1 = fitz.Point(pts[i][0], pts[i][1])
                    p2 = fitz.Point(pts[i+1][0], pts[i+1][1])
                    page.draw_line(p1, p2, color=(0, 0.2, 0.8), width=1.5)
            self.dirty = True
            # ไม่วาดกรอบลง PDF — เป็นแค่ helper บนจอ
            self.doc.save(out, garbage=3, deflate=True)
            self.dirty = False
            messagebox.showinfo("สำเร็จ", f"บันทึกแล้วที่:\n{out}")
            self.sign_cleanup()
        except Exception as e:
            messagebox.showerror("ผิดพลาด", str(e))

    def close(self):
        if self.doc:
            self.doc.close()
            self.doc = None


# =========================================================
#  PDFReader — Main window มี Notebook + toolbar
# =========================================================
class PDFReader(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Reader")
        self.geometry("1100x780")
        self._set_app_icon()

        self.tabs = []  # list of PDFTab

        self._build_menu()
        self._build_toolbar()
        self._build_notebook()
        self._build_statusbar()
        self._bind_keys()
        self.protocol("WM_DELETE_WINDOW", self._on_app_close)

    # ---------- UI ----------
    def _build_menu(self):
        menubar = tk.Menu(self)
        self._menu_labels = {
            "sign":   "✍ เซ็นชื่อ",
            "image":  "🖼 ใส่รูป",
            "hl":     "🖍 ไฮไลต์",
            "draw":   "✏ ขีดเขียน",
            "cmt":    "💬 คอมเมนต์",
            "redact": "🗑 ลบข้อความ",
        }
        self._menu_index = {}

        menubar.add_command(label="PDF จากข้อความ", command=self.create_text_pdf)
        menubar.add_command(label="รวมรูปเป็น PDF", command=self.images_to_pdf)
        menubar.add_command(label="Merge PDFs", command=self.merge_pdfs)
        menubar.add_command(label="Split PDF", command=self.split_pdf)
        menubar.add_command(label="Word → PDF", command=self.word_to_pdf)
        menubar.add_command(label="🖨 พิมพ์", command=self.print_pdf)

        menubar.add_command(label=self._menu_labels["sign"], command=self.toggle_sign_mode)
        self._menu_index["sign"] = menubar.index("end")
        menubar.add_command(label=self._menu_labels["image"], command=self.insert_image)
        self._menu_index["image"] = menubar.index("end")
        menubar.add_command(label=self._menu_labels["hl"], command=self.toggle_highlight_mode)
        self._menu_index["hl"] = menubar.index("end")
        menubar.add_command(label=self._menu_labels["draw"], command=self.toggle_draw_mode)
        self._menu_index["draw"] = menubar.index("end")
        menubar.add_command(label=self._menu_labels["cmt"], command=self.toggle_comment_mode)
        self._menu_index["cmt"] = menubar.index("end")
        menubar.add_command(label=self._menu_labels["redact"], command=self.toggle_redact_mode)
        self._menu_index["redact"] = menubar.index("end")

        menubar.add_command(label="📑 Thumbnail", command=self._toggle_thumbnails)
        menubar.add_command(label="↺ หมุนซ้าย", command=lambda: self._rotate(-90, all_pages=False))
        menubar.add_command(label="↻ หมุนขวา", command=lambda: self._rotate(90, all_pages=False))
        menubar.add_command(label="↺↺ หมุนซ้ายทุกหน้า", command=lambda: self._rotate(-90, all_pages=True))
        menubar.add_command(label="↻↻ หมุนขวาทุกหน้า", command=lambda: self._rotate(90, all_pages=True))

        menubar.add_command(label="💾 บันทึก", command=self.save_doc)
        menubar.add_command(label="⛶ Full Screen", command=self.toggle_fullscreen)
        menubar.add_command(label="ℹ About", command=self._show_about)
        self._menubar = menubar
        self.config(menu=menubar)

    def _update_menu_indicators(self):
        t = self._active_tab()
        modes = {
            "sign":  bool(t and getattr(t, "sign_mode", False)),
            "image": bool(t and getattr(t, "image_mode", False)),
            "hl":    bool(t and getattr(t, "hl_mode", False)),
            "draw":  bool(t and getattr(t, "draw_mode", False)),
            "cmt":   bool(t and getattr(t, "cmt_mode", False)),
            "redact": bool(t and getattr(t, "redact_mode", False)),
        }
        # อัปเดต menu labels
        for key, is_active in modes.items():
            idx = self._menu_index.get(key)
            if idx is not None:
                base = self._menu_labels[key]
                lbl = f"● {base} (กำลังใช้)" if is_active else base
                self._menubar.entryconfig(idx, label=lbl)

    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=6)
        bar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(bar, text="＋ เปิดไฟล์", command=self.open_file).pack(side=tk.LEFT)
        ttk.Separator(bar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)

        self.undo_btn = ttk.Button(bar, text="↶ Undo", command=self.undo)
        self.undo_btn.pack(side=tk.LEFT)
        self.redo_btn = ttk.Button(bar, text="↷ Redo", command=self.redo)
        self.redo_btn.pack(side=tk.LEFT, padx=(2, 0))
        ttk.Separator(bar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)

        ttk.Button(bar, text="◀", width=3, command=self._prev_page).pack(side=tk.LEFT)
        self.page_entry = ttk.Entry(bar, width=10, justify="center")
        self.page_entry.pack(side=tk.LEFT, padx=4)
        self.page_entry.bind("<Return>", self._go_to_page)
        ttk.Button(bar, text="▶", width=3, command=self._next_page).pack(side=tk.LEFT)

        ttk.Separator(bar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)
        ttk.Button(bar, text="−", width=3, command=self._zoom_out).pack(side=tk.LEFT)
        self.zoom_lbl = ttk.Label(bar, text="—", width=6, anchor="center")
        self.zoom_lbl.pack(side=tk.LEFT)
        ttk.Button(bar, text="+", width=3, command=self._zoom_in).pack(side=tk.LEFT)
        ttk.Button(bar, text="Fit", command=self._fit_width).pack(side=tk.LEFT, padx=(4, 0))

        ttk.Separator(bar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=6)
        self.search_entry = ttk.Entry(bar, width=25)
        self.search_entry.pack(side=tk.LEFT)
        self.search_entry.bind("<Return>", lambda e: self._search())
        ttk.Button(bar, text="ค้นหา", command=self._search).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="ถัดไป", command=self._next_hit).pack(side=tk.LEFT)

    def _build_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self._sync_toolbar())

        # หน้าจอเริ่มต้น
        self.welcome = ttk.Frame(self.notebook)
        ttk.Label(self.welcome, text="ยินดีต้อนรับ", font=("Segoe UI", 20)).pack(pady=40)
        ttk.Label(self.welcome, text="กด '＋ เปิดไฟล์' เพื่อเริ่มอ่าน PDF\nเปิดได้หลายไฟล์พร้อมกันเป็นแท็บ",
                  justify="center", foreground="#666").pack()
        ttk.Button(self.welcome, text="＋ เปิดไฟล์",
                   command=self.open_file).pack(pady=20)
        self.notebook.add(self.welcome, text="เริ่มต้น")

        # คลิกขวาที่แท็บ → ปิด
        self.notebook.bind("<Button-3>", self._on_tab_right_click)
        # คลิกกลาง (ลูกกลิ้ง) → ปิด
        self.notebook.bind("<Button-2>", self._on_tab_middle_click)
        # คลิกซ้ายที่กากบาท ✕ → ปิด
        self.notebook.bind("<ButtonRelease-1>", self._on_tab_left_click)

    def _build_statusbar(self):
        self.status = ttk.Label(self, text="เลือก 'เปิดไฟล์' เพื่อเริ่ม", anchor="w", padding=4)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    def _bind_keys(self):
        self.bind("<Left>", lambda e: self._prev_page())
        self.bind("<Right>", lambda e: self._next_page())
        self.bind("<Control-o>", lambda e: self.open_file())
        self.bind("<Control-w>", lambda e: self._close_current_tab())
        self.bind("<Control-Tab>", lambda e: self._cycle_tab(1))
        self.bind("<Control-Shift-Tab>", lambda e: self._cycle_tab(-1))
        self.bind("<Control-p>", lambda e: self.print_pdf())
        self.bind("<F11>", self.toggle_fullscreen)
        self.bind("<Return>", self._on_enter)
        self.bind("<Escape>", self._on_esc)
        self.bind("<Control-f>", lambda e: self.search_entry.focus_set())
        self.bind("<Control-plus>", lambda e: self._zoom_in())
        self.bind("<Control-minus>", lambda e: self._zoom_out())
        self.bind("<Control-z>", lambda e: self.undo())
        self.bind("<Control-y>", lambda e: self.redo())
        self.bind("<Control-Shift-Z>", lambda e: self.redo())

    # ---------- Active tab helpers ----------
    def _active_tab(self):
        idx = self.notebook.index(self.notebook.select()) if self.notebook.tabs() else -1
        if 0 <= idx - 1 < len(self.tabs):  # -1 เพราะ welcome tab อยู่ index 0
            return self.tabs[idx - 1]
        # ลอง lookup ตาม widget
        try:
            widget = self.notebook.nametowidget(self.notebook.select())
            for t in self.tabs:
                if str(t) == str(widget):
                    return t
        except Exception:
            pass
        return None

    def _sync_toolbar(self):
        t = self._active_tab()
        if not t:
            self.zoom_lbl.config(text="—")
            self.page_entry.delete(0, tk.END)
            self.status.config(text="ไม่มีไฟล์เปิดอยู่")
            self.title("PDF Reader")
            self.undo_btn.state(["disabled"])
            self.redo_btn.state(["disabled"])
            return
        total = len(t.doc)
        self.page_entry.delete(0, tk.END)
        self.page_entry.insert(0, str(t.page_index + 1))
        self.zoom_lbl.config(text=f"{int(t.zoom * 100)}%")
        self.status.config(text=f"หน้า {t.page_index + 1} จาก {total} | {t.doc_path}")
        self.title(f"PDF Reader — {os.path.basename(t.doc_path)}")
        self.undo_btn.state(["!disabled"] if t.undo_stack else ["disabled"])
        self.redo_btn.state(["!disabled"] if t.redo_stack else ["disabled"])

    # ---------- Toolbar actions (delegate to active tab) ----------
    def _prev_page(self):
        t = self._active_tab()
        if t: t.prev_page()

    def _next_page(self):
        t = self._active_tab()
        if t: t.next_page()

    def _go_to_page(self, _e=None):
        t = self._active_tab()
        if not t:
            return
        try:
            n = int(self.page_entry.get()) - 1
            t.go_to_page(n)
        except ValueError:
            pass

    def _zoom_in(self):
        t = self._active_tab()
        if t:
            t.zoom_in()
            self._sync_toolbar()

    def _zoom_out(self):
        t = self._active_tab()
        if t:
            t.zoom_out()
            self._sync_toolbar()

    def _fit_width(self):
        t = self._active_tab()
        if t:
            t.fit_width()
            self._sync_toolbar()

    def _search(self):
        t = self._active_tab()
        if not t:
            return
        term = self.search_entry.get().strip()
        if not term:
            return
        found = t.search(term)
        if not found:
            self.status.config(text=f"ไม่พบ '{term}'")
        else:
            self.status.config(text=f"ผลลัพธ์ 1/{len(t.search_hits)}")

    def _next_hit(self):
        t = self._active_tab()
        if not t or not t.search_hits:
            return
        t.next_hit()
        self.status.config(text=f"ผลลัพธ์ {t.search_pos + 1}/{len(t.search_hits)}")

    # ---------- Tabs management ----------
    def open_file(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if not paths:
            return
        for p in paths:
            self._open_in_new_tab(p)

    def _open_in_new_tab(self, path):
        try:
            tab = PDFTab(self.notebook, self, path)
        except Exception as e:
            messagebox.showerror("ผิดพลาด", f"เปิดไฟล์ไม่ได้:\n{e}")
            return
        # ซ่อน welcome tab เมื่อมีไฟล์แรก
        if self.welcome in [self.notebook.nametowidget(t) for t in self.notebook.tabs()]:
            self.notebook.forget(self.welcome)
        name = os.path.basename(path)
        if len(name) > 30:
            name = name[:27] + "..."
        self.notebook.add(tab, text=f" {name}  ✕ ")
        self.tabs.append(tab)
        self.notebook.select(tab)
        self.after(50, tab.render)

    def _on_tab_right_click(self, event):
        try:
            idx = self.notebook.index(f"@{event.x},{event.y}")
        except tk.TclError:
            return
        self._close_tab_at(idx)

    def _on_tab_left_click(self, event):
        try:
            idx = self.notebook.index(f"@{event.x},{event.y}")
        except tk.TclError:
            return
        # หาขอบขวาของแท็บ: เลื่อน x ไปทางขวาจนกว่า tab index จะเปลี่ยน
        x = event.x
        limit = event.x + 300
        while x < limit:
            try:
                if self.notebook.index(f"@{x + 1},{event.y}") != idx:
                    break
            except tk.TclError:
                break
            x += 1
        # ถ้าคลิกใกล้ขอบขวา (บริเวณ ✕) → ปิด
        if x - event.x <= 22:
            self._close_tab_at(idx)
            return "break"

    def _on_tab_middle_click(self, event):
        try:
            idx = self.notebook.index(f"@{event.x},{event.y}")
        except tk.TclError:
            return
        self._close_tab_at(idx)

    def _on_app_close(self):
        dirty_tabs = [t for t in self.tabs if getattr(t, "dirty", False)]
        if dirty_tabs:
            names = "\n".join(f"• {os.path.basename(t.doc_path)}" for t in dirty_tabs)
            ans = messagebox.askyesnocancel(
                "ยังไม่ได้บันทึก",
                f"มีไฟล์ที่แก้ไขแล้วยังไม่ได้บันทึก:\n\n{names}\n\n"
                "ต้องการบันทึกก่อนออกหรือไม่?")
            if ans is None:
                return  # Cancel — ไม่ปิดโปรแกรม
            if ans:
                for t in dirty_tabs:
                    self.notebook.select(t)
                    if not self.save_doc():
                        return  # ผู้ใช้ยกเลิก save → หยุด
        self.destroy()

    def _close_tab_at(self, idx):
        try:
            widget = self.notebook.nametowidget(self.notebook.tabs()[idx])
        except (IndexError, tk.TclError):
            return
        if widget is self.welcome:
            return
        # หา PDFTab
        for t in self.tabs:
            if str(t) == str(widget):
                if getattr(t, "dirty", False):
                    ans = messagebox.askyesnocancel(
                        "ยังไม่ได้บันทึก",
                        f"'{os.path.basename(t.doc_path)}' มีการแก้ไขที่ยังไม่ได้บันทึก\n\n"
                        "ต้องการบันทึกก่อนปิดหรือไม่?")
                    if ans is None:  # Cancel — ไม่ปิด
                        return
                    if ans:  # Yes — บันทึกก่อน
                        self.notebook.select(t)
                        saved = self.save_doc()
                        if not saved:
                            return  # ยกเลิก save → ไม่ปิดแท็บ
                t.close()
                self.tabs.remove(t)
                break
        self.notebook.forget(widget)
        # ถ้าไม่มีแท็บเหลือ → โชว์ welcome
        if not self.tabs:
            self.notebook.add(self.welcome, text="เริ่มต้น")
        self._sync_toolbar()

    def _close_current_tab(self):
        try:
            idx = self.notebook.index(self.notebook.select())
            self._close_tab_at(idx)
        except tk.TclError:
            pass

    def _cycle_tab(self, direction):
        tabs = self.notebook.tabs()
        if len(tabs) <= 1:
            return
        cur = self.notebook.index(self.notebook.select())
        new = (cur + direction) % len(tabs)
        self.notebook.select(tabs[new])

    # ---------- Keys ----------
    def _on_enter(self, _e):
        t = self._active_tab()
        if t and t.sign_mode:
            t.sign_save()

    def _on_esc(self, _e):
        t = self._active_tab()
        if not t:
            self.attributes("-fullscreen", False)
            return
        if t.sign_mode:
            t.sign_cleanup()
        elif getattr(t, "image_mode", False):
            t._img_cleanup()
            self.status.config(text="ยกเลิกใส่รูป")
        elif getattr(t, "hl_mode", False):
            t._hl_cleanup()
            self.status.config(text="ออกจากไฮไลต์")
        elif getattr(t, "draw_mode", False):
            t._draw_commit()
            self.status.config(text="ออกจากขีดเขียน")
        elif getattr(t, "cmt_mode", False):
            t._cmt_cleanup()
            self.status.config(text="ออกจากคอมเมนต์")
        elif getattr(t, "redact_mode", False):
            t._rd_cleanup()
            self.status.config(text="ออกจากลบข้อความ")
        else:
            self.attributes("-fullscreen", False)
            self.after(50, t.render)
        self._update_menu_indicators()

    def toggle_fullscreen(self, _event=None):
        is_full = bool(self.attributes("-fullscreen"))
        self.attributes("-fullscreen", not is_full)
        t = self._active_tab()
        if t:
            self.after(50, t.render)

    # ---------- Print / Sign ----------
    def print_pdf(self):
        t = self._active_tab()
        if not t:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        path = t.doc_path
        # ลอง Print verb ก่อน
        try:
            os.startfile(path, "print")
            self.status.config(text="ส่งไปยังเครื่องพิมพ์แล้ว")
            return
        except OSError:
            pass
        # Fallback: เปิดใน Microsoft Edge เพื่อให้ผู้ใช้กด Ctrl+P
        edge_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
        for edge in edge_paths:
            if os.path.isfile(edge):
                try:
                    import subprocess
                    subprocess.Popen([edge, path])
                    self.status.config(text="เปิดใน Edge แล้ว — กด Ctrl+P เพื่อพิมพ์")
                    messagebox.showinfo(
                        "พิมพ์ผ่านเบราว์เซอร์",
                        "เปิดไฟล์ใน Microsoft Edge แล้ว\n"
                        "กด Ctrl+P ในหน้าต่างนั้นเพื่อพิมพ์")
                    return
                except Exception:
                    pass
        messagebox.showerror(
            "พิมพ์ไม่ได้",
            "โปรแกรมเริ่มต้นสำหรับ PDF ไม่รองรับการสั่งพิมพ์โดยตรง\n\n"
            "วิธีแก้:\n"
            "1. เปลี่ยน default PDF ไปเป็น Edge หรือ Adobe Reader\n"
            "   (Settings → Apps → Default apps → .pdf)\n"
            "2. หรือคลิกขวาที่ไฟล์ → Open with → เลือกโปรแกรมที่พิมพ์ได้")

    def _set_app_icon(self):
        # หา icon.ico ทั้งใน dev และตอน build เป็น exe
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        for name in ("icon.ico", os.path.join(base, "icon.ico")):
            if os.path.isfile(name):
                try:
                    self.iconbitmap(name)
                    return
                except Exception:
                    pass

    def _show_about(self):
        messagebox.showinfo(
            "เกี่ยวกับโปรแกรม",
            "PDF Reader\n"
            "\n"
            "พัฒนาโดย: Thanagrid C.\n"
            "อีเมล: Thanagrid.c@ku.th\n"
            "\n"
            "License: MIT\n"
            "© 2026"
        )

    def undo(self):
        t = self._active_tab()
        if not t:
            return
        if t.undo():
            self.status.config(text=f"Undo ({len(t.undo_stack)} เหลือ)")
        else:
            self.status.config(text="ไม่มีอะไรให้ Undo")
        self._sync_toolbar()

    def redo(self):
        t = self._active_tab()
        if not t:
            return
        if t.redo():
            self.status.config(text=f"Redo ({len(t.redo_stack)} เหลือ)")
        else:
            self.status.config(text="ไม่มีอะไรให้ Redo")
        self._sync_toolbar()

    def _toggle_thumbnails(self):
        t = self._active_tab()
        if not t:
            return
        t.toggle_thumbnails()

    def _rotate(self, delta, all_pages=False):
        t = self._active_tab()
        if not t:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        if all_pages:
            t.rotate_all(delta)
            self.status.config(text=f"หมุนทุกหน้า {delta:+d}°")
        else:
            t.rotate_current(delta)
            self.status.config(text=f"หมุนหน้า {t.page_index + 1} {delta:+d}°")

    def save_doc(self):
        t = self._active_tab()
        if not t:
            messagebox.showwarning("แจ้ง", "ไม่มีไฟล์เปิดอยู่")
            return
        src_dir = os.path.dirname(t.doc_path)
        src_name = os.path.basename(t.doc_path)
        default_name = src_name.replace(".pdf", "_edited.pdf")
        out = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialdir=src_dir, initialfile=default_name)
        if not out:
            return
        # กันเขียนทับไฟล์ต้นฉบับ
        if os.path.normcase(os.path.abspath(out)) == os.path.normcase(os.path.abspath(t.doc_path)):
            messagebox.showwarning(
                "แจ้ง",
                "ไม่สามารถบันทึกทับไฟล์ต้นฉบับได้\nกรุณาตั้งชื่อไฟล์ใหม่")
            return self.save_doc()
        try:
            t.doc.save(out, garbage=3, deflate=True)
            t.dirty = False
            messagebox.showinfo("สำเร็จ", f"บันทึกเป็นไฟล์ใหม่ที่:\n{out}")
            return True
        except Exception as e:
            messagebox.showerror("ผิดพลาด", str(e))
            return False

    def toggle_highlight_mode(self):
        t = self._active_tab()
        if not t:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        t.toggle_highlight_mode()
        self.status.config(text="ไฮไลต์: ลากคลุมพื้นที่ | Esc=ออก | 💾 บันทึกก่อนปิด")
        self._update_menu_indicators()

    def toggle_redact_mode(self):
        t = self._active_tab()
        if not t:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        t.toggle_redact_mode()
        self.status.config(text="ลบข้อความ: ลากคลุมข้อความที่จะลบ | Esc=ออก | 💾 บันทึกก่อนปิด")
        self._update_menu_indicators()

    def toggle_draw_mode(self):
        t = self._active_tab()
        if not t:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        t.toggle_draw_mode()
        self.status.config(text="ขีดเขียน: ลากเมาส์ขีดได้ทั่วหน้า | Esc=ออก | 💾 บันทึกก่อนปิด")
        self._update_menu_indicators()

    def toggle_comment_mode(self):
        t = self._active_tab()
        if not t:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        t.toggle_comment_mode()
        self.status.config(text="คอมเมนต์: คลิกจุดที่ต้องการวางโน้ต | Esc=ออก")
        self._update_menu_indicators()

    def insert_image(self):
        t = self._active_tab()
        if not t:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        t.start_image_mode()
        self.status.config(text="ใส่รูป: ลากเมาส์สร้างกรอบ → เลือกรูป | Esc=ยกเลิก")
        self._update_menu_indicators()

    def toggle_sign_mode(self):
        t = self._active_tab()
        if not t:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        t.toggle_sign_mode()
        if t.sign_mode:
            self.status.config(text="ขั้นที่ 1: ลากเมาส์เพื่อสร้างกรอบเซ็นชื่อ | Esc=ยกเลิก")
        else:
            self.status.config(text="ออกจากโหมดเซ็นชื่อ")
        self._update_menu_indicators()

    # ---------- Create PDF ----------
    def _thai_font(self):
        for p in [r"C:\Windows\Fonts\tahoma.ttf",
                  r"C:\Windows\Fonts\leelawad.ttf",
                  r"C:\Windows\Fonts\cordia.ttc"]:
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
                page = doc.new_page()
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
            title="เลือกรูปภาพ",
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
            title="เลือก PDF ที่ต้องการรวม", filetypes=[("PDF", "*.pdf")])
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

    def split_pdf(self):
        t = self._active_tab()
        if not t:
            messagebox.showwarning("แจ้ง", "เปิดไฟล์ PDF ก่อน")
            return
        out_dir = filedialog.askdirectory(title="เลือกโฟลเดอร์บันทึก")
        if not out_dir:
            return
        try:
            for i in range(len(t.doc)):
                new = fitz.open()
                new.insert_pdf(t.doc, from_page=i, to_page=i)
                new.save(os.path.join(out_dir, f"page_{i+1:03d}.pdf"))
                new.close()
            messagebox.showinfo("สำเร็จ", f"แยก {len(t.doc)} ไฟล์ที่:\n{out_dir}")
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
            messagebox.showerror("ผิดพลาด", f"import docx2pdf ไม่ได้: {e}")
            return
        out_dir = filedialog.askdirectory(title="เลือกโฟลเดอร์บันทึก PDF")
        if not out_dir:
            return
        import io
        if sys.stdout is None:
            sys.stdout = io.StringIO()
        if sys.stderr is None:
            sys.stderr = io.StringIO()

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
        ttk.Label(prog, text="โปรดรอสักครู่", foreground="#666").pack()
        prog.update()

        import threading
        result = {"error": None, "done": 0}

        def worker():
            try:
                for i, p in enumerate(paths):
                    out = os.path.join(out_dir,
                                       os.path.splitext(os.path.basename(p))[0] + ".pdf")
                    self.after(0, lbl.config,
                               {"text": f"กำลังแปลง ({i+1}/{len(paths)})"})
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
                                     f"{result['error']}\n\nต้องมี Microsoft Word")
            else:
                messagebox.showinfo("สำเร็จ",
                                    f"แปลง {result['done']} ไฟล์แล้วที่:\n{out_dir}")
        self.after(100, check)


if __name__ == "__main__":
    import traceback
    try:
        app = PDFReader()
        # เปิดไฟล์ที่ Windows ส่งมาผ่าน argv (double-click / Open with)
        for arg in sys.argv[1:]:
            if arg.lower().endswith(".pdf") and os.path.isfile(arg):
                app.after(100, lambda p=arg: app._open_in_new_tab(p))
        app.mainloop()
    except Exception:
        traceback.print_exc()
        input("\nกด Enter เพื่อปิด...")
