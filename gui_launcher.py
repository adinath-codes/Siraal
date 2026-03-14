"""
gui_launcher_2d_apple.py — Siraal Manufacturing Engine
2D Drafting Engine — Apple Design + Glassmorphism + Glow
Aesthetic: macOS Sonoma / visionOS vibes — frosted panels, SF-style type,
           luminous accents, animated progress, depth shadows.
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import queue
import math
import time

# ── Apple / visionOS palette ─────────────────────────────────────────────────
C = {
    # backgrounds
    "bg_deep":      "#050A12",
    "bg_mid":       "#0A1628",
    # glass panels  (simulate with semi-opaque colours — no true alpha in tk)
    "glass_dark":   "#0D1F35",
    "glass_mid":    "#122340",
    "glass_light":  "#19304F",
    "glass_card":   "#1A3356",
    # accents
    "blue":         "#2997FF",   # Apple blue
    "blue_glow":    "#1A6BCC",
    "blue_dim":     "#0A3A70",
    "mint":         "#63E6BE",
    "teal":         "#30D5C8",
    "indigo":       "#5E5CE6",
    "gold":         "#FFD60A",
    # text
    "text_primary": "#F5F5F7",
    "text_sec":     "#98A0A8",
    "text_dim":     "#4A5568",
    # status
    "success":      "#30D158",
    "warn":         "#FFD60A",
    "error":        "#FF453A",
    # borders / separators
    "border_glow":  "#2997FF",
    "border_dim":   "#1E3A5F",
    "sep":          "#1E3A5F",
}

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ── SF-style font stack (falls back gracefully) ───────────────────────────────
# Global font scale — bump everything up for readability
_FS = 4   # add this to every sf() call base

def sf(size=12, weight="normal"):
    for fam in ("SF Pro Display", "SF Pro Text", ".AppleSystemUIFont",
                "Helvetica Neue", "Segoe UI", "Calibri"):
        return ctk.CTkFont(fam, size + _FS, weight)


# ── Glow Canvas Widget ────────────────────────────────────────────────────────
class GlowBar(tk.Canvas):
    """Animated progress bar with luminous glow trail."""
    def __init__(self, master, **kw):
        h = kw.pop("height", 6)
        super().__init__(master, height=h, bg=C["glass_dark"],
                         highlightthickness=0, bd=0, **kw)
        self._val = 0.0
        self._target = 0.0
        self._animating = False
        self._h = h
        self.bind("<Configure>", self._draw)

    def set(self, val: float):
        self._target = max(0.0, min(1.0, val))
        if not self._animating:
            self._animating = True
            self._step()

    def _step(self):
        diff = self._target - self._val
        if abs(diff) < 0.002:
            self._val = self._target
            self._animating = False
            self._draw()
            return
        self._val += diff * 0.18
        self._draw()
        self.after(16, self._step)

    def _draw(self, *_):
        self.delete("all")
        w = self.winfo_width()
        if w < 2: return
        h = self._h
        r = h // 2
        # track
        self.create_rounded_rect(0, 0, w, h, r, fill=C["glass_mid"], outline="")
        # fill
        fw = max(0, int(w * self._val))
        if fw > 4:
            self.create_rounded_rect(0, 0, fw, h, r, fill=C["blue"], outline="")
            # glow head
            gx = fw
            for i, alpha in [(10, C["blue_dim"]), (6, C["blue_glow"]), (3, C["blue"])]:
                self.create_rounded_rect(max(0, gx-i), 0, gx+i, h, r,
                                         fill=alpha, outline="")

    def create_rounded_rect(self, x1, y1, x2, y2, r, **kw):
        r = min(r, (x2-x1)//2, (y2-y1)//2)
        if r < 1:
            self.create_rectangle(x1, y1, x2, y2, **kw); return
        pts = [
            x1+r, y1,  x2-r, y1,
            x2,   y1,  x2,   y1+r,
            x2,   y2-r, x2, y2,
            x2-r, y2,  x1+r, y2,
            x1,   y2,  x1,   y2-r,
            x1,   y1+r, x1, y1,
        ]
        self.create_polygon(pts, smooth=True, **kw)


# ── Pulse dot ─────────────────────────────────────────────────────────────────
class PulseDot(tk.Canvas):
    """Animated breathing status dot."""
    def __init__(self, master, color=C["success"], **kw):
        super().__init__(master, width=14, height=14,
                         bg=C["glass_dark"], highlightthickness=0, bd=0, **kw)
        self._color = color
        self._t = 0
        self._active = False
        self._draw()

    def set_color(self, color: str):
        self._color = color
        self._draw()

    def start_pulse(self):
        self._active = True
        self._pulse()

    def stop_pulse(self):
        self._active = False
        self._draw()

    def _pulse(self):
        if not self._active: return
        self._t += 0.12
        self._draw()
        self.after(40, self._pulse)

    def _draw(self, *_):
        self.delete("all")
        cx, cy, r = 7, 7, 4
        if self._active:
            scale = 0.6 + 0.4 * abs(math.sin(self._t))
            outer = int(r * (1 + scale))
            self.create_oval(cx-outer, cy-outer, cx+outer, cy+outer,
                             fill=self._color, outline="", stipple="gray50")
        self.create_oval(cx-r, cy-r, cx+r, cy+r,
                         fill=self._color, outline="")


# ── Separator ─────────────────────────────────────────────────────────────────
def sep(parent, pady=(8, 4)):
    f = ctk.CTkFrame(parent, fg_color=C["sep"], height=1, corner_radius=0)
    f.pack(fill="x", padx=16, pady=pady)


# ── Section label ─────────────────────────────────────────────────────────────
def section_label(parent, text):
    sep(parent)
    ctk.CTkLabel(parent, text=text.upper(),
                 font=sf(10, "bold"),
                 text_color=C["text_dim"]).pack(anchor="w", padx=18, pady=(2, 3))


# ═════════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# =════════════════════════════════════════════════════════════════════════════
class SiraalApple2D(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Siraal Manufacturing Engine")
        self.geometry("1500x920")
        self.minsize(1280, 800)
        self.configure(fg_color=C["bg_deep"])
        self._q: queue.Queue = queue.Queue()
        self._validator = None
        self._build_bg_canvas()
        self._build_titlebar()
        self._build_body()
        self.after(100, self._poll)

    # ── Gradient background canvas ────────────────────────────────────────────
    def _build_bg_canvas(self):
        self._bg = tk.Canvas(self, bg=C["bg_deep"], highlightthickness=0)
        self._bg.place(relwidth=1, relheight=1)
        self._bg.bind("<Configure>", self._draw_bg)

    def _draw_bg(self, *_):
        self._bg.delete("all")
        w = self._bg.winfo_width()
        h = self._bg.winfo_height()
        # subtle radial "light leak" top-left in indigo/blue
        for i in range(30, 0, -1):
            r = i * 22
            alpha_hex = format(int(i * 3), '02x')
            try:
                self._bg.create_oval(-r, -r, r, r,
                                     fill=C["bg_mid"], outline="")
            except Exception:
                pass
        # bottom-right warm glow
        for i in range(20, 0, -1):
            r = i * 18
            self._bg.create_oval(w - r, h - r, w + r, h + r,
                                 fill=C["glass_dark"], outline="")

    # ── macOS-style title bar ─────────────────────────────────────────────────
    def _build_titlebar(self):
        bar = ctk.CTkFrame(self, fg_color=C["glass_dark"],
                           corner_radius=0, height=64)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)

        # Traffic lights (decorative)
        dots = ctk.CTkFrame(bar, fg_color="transparent")
        dots.pack(side="left", padx=16, pady=18)
        for col in ("#FF5F57", "#FFBD2E", "#28CA41"):
            c2 = tk.Canvas(dots, width=14, height=14, bg=C["glass_dark"],
                           highlightthickness=0)
            c2.pack(side="left", padx=3)
            c2.create_oval(1, 1, 13, 13, fill=col, outline="")

        # Title
        ctk.CTkLabel(bar, text="Siraal Manufacturing Engine",
                     font=sf(17, "bold"),
                     text_color=C["text_primary"]).pack(side="left", padx=14)
        ctk.CTkLabel(bar, text="2D Drafting Engine",
                     font=sf(13),
                     text_color=C["blue"]).pack(side="left", padx=4)

        # Right badges
        badge_frame = ctk.CTkFrame(bar, fg_color="transparent")
        badge_frame.pack(side="right", padx=20)
        for txt, col in [("IS 11669", C["text_dim"]),
                         ("ISO 128", C["text_dim"]),
                         ("1st Angle", C["mint"]),
                         ("v3.0", C["gold"])]:
            b = ctk.CTkFrame(badge_frame, fg_color=C["glass_mid"],
                             corner_radius=12)
            b.pack(side="left", padx=4)
            ctk.CTkLabel(b, text=txt, font=sf(10),
                         text_color=col).pack(padx=10, pady=4)

        # Thin glowing bottom border
        glow_line = tk.Canvas(bar, height=1, bg=C["blue_dim"],
                              highlightthickness=0)
        glow_line.place(relx=0, rely=1.0, relwidth=1.0, anchor="sw")

    # ── Body: left panel + right panel ───────────────────────────────────────
    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=14, pady=10)
        body.columnconfigure(0, weight=0, minsize=310)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        # Left glass panel
        lp = ctk.CTkFrame(body, fg_color=C["glass_dark"],
                          corner_radius=16,
                          border_width=1, border_color=C["border_dim"],
                          width=310)
        lp.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        lp.pack_propagate(False)
        self._build_left(lp)

        # Right glass panel
        rp = ctk.CTkFrame(body, fg_color=C["glass_dark"],
                          corner_radius=16,
                          border_width=1, border_color=C["border_dim"])
        rp.grid(row=0, column=1, sticky="nsew")
        self._build_right(rp)

    # ── LEFT PANEL ────────────────────────────────────────────────────────────
    def _build_left(self, p):
        # Hero icon + title
        hero = ctk.CTkFrame(p, fg_color=C["glass_mid"], corner_radius=12)
        hero.pack(fill="x", padx=14, pady=(16, 8))
        icon_c = tk.Canvas(hero, width=52, height=52,
                           bg=C["glass_mid"], highlightthickness=0)
        icon_c.pack(pady=(12, 6))
        # Glowing pencil-square icon
        icon_c.create_rectangle(7, 7, 45, 45, fill=C["blue_dim"],
                                 outline=C["blue"], width=1)
        icon_c.create_line(12, 40, 20, 28, 32, 16, fill=C["blue"], width=2,
                           smooth=True)
        icon_c.create_line(32, 16, 40, 24, fill=C["blue"], width=2)
        icon_c.create_line(12, 40, 18, 37, fill=C["mint"], width=1)

        ctk.CTkLabel(hero, text="2D Drafting Engine",
                     font=sf(15, "bold"),
                     text_color=C["text_primary"]).pack(pady=(0, 2))
        ctk.CTkLabel(hero, text="Plates · Gears · Shafts · Ring Gears",
                     font=sf(11),
                     text_color=C["text_sec"]).pack(pady=(0, 12))

        # BOM FILE
        section_label(p, "BOM File")
        self._file_var = ctk.StringVar(value="excels/demo.xlsx")
        ef = ctk.CTkFrame(p, fg_color=C["glass_mid"], corner_radius=8,
                          border_width=1, border_color=C["border_dim"])
        ef.pack(fill="x", padx=14, pady=(0, 4))
        ctk.CTkEntry(ef, textvariable=self._file_var,
                     fg_color="transparent", border_width=0,
                     text_color=C["text_primary"],
                     height=36, font=sf(11)
                     ).pack(fill="x", padx=4)
        ctk.CTkButton(p, text="Browse…", height=34,
                      fg_color=C["glass_mid"],
                      hover_color=C["glass_light"],
                      border_color=C["border_dim"], border_width=1,
                      text_color=C["text_sec"], font=sf(11),
                      corner_radius=8,
                      command=self._browse).pack(fill="x", padx=14, pady=(0, 8))

        # FILTERS
        section_label(p, "Filters")
        ctk.CTkLabel(p, text="Part Type", font=sf(11),
                     text_color=C["text_dim"]).pack(anchor="w", padx=18)
        self._type_var = ctk.StringVar(value="All")
        ctk.CTkComboBox(p, variable=self._type_var,
                        values=["All","Plate","Spur_Gear","Ring_Gear",
                                "Stepped_Shaft","Flanged_Shaft"],
                        fg_color=C["glass_mid"],
                        button_color=C["glass_light"],
                        dropdown_fg_color=C["glass_card"],
                        border_color=C["border_dim"],
                        text_color=C["text_primary"],
                        font=sf(11), height=34,
                        corner_radius=8
                        ).pack(fill="x", padx=14, pady=(2, 6))
        ctk.CTkLabel(p, text="Priority", font=sf(11),
                     text_color=C["text_dim"]).pack(anchor="w", padx=18)
        self._prio_var = ctk.StringVar(value="All")
        ctk.CTkComboBox(p, variable=self._prio_var,
                        values=["All","High","Medium","Low"],
                        fg_color=C["glass_mid"],
                        button_color=C["glass_light"],
                        dropdown_fg_color=C["glass_card"],
                        border_color=C["border_dim"],
                        text_color=C["text_primary"],
                        font=sf(11), height=34,
                        corner_radius=8
                        ).pack(fill="x", padx=14, pady=(2, 10))

        # ACTIONS
        section_label(p, "Actions")

        # ① Validate
        ctk.CTkButton(p, text="①  Validate BOM",
                      height=40,
                      fg_color=C["glass_mid"],
                      hover_color=C["glass_light"],
                      border_color=C["teal"],
                      border_width=1,
                      text_color=C["teal"],
                      font=sf(12, "bold"),
                      corner_radius=10,
                      command=self._run_validate
                      ).pack(fill="x", padx=14, pady=(0, 6))

        # ② RUN (glowing Apple blue button)
        self._btn_run = ctk.CTkButton(
            p, text="②  Run 2D CAD Batch",
            height=46,
            fg_color=C["blue"],
            hover_color=C["blue_glow"],
            text_color="#FFFFFF",
            font=sf(14, "bold"),
            corner_radius=12,
            command=self._run_batch)
        self._btn_run.pack(fill="x", padx=14, pady=(0, 6))

        # ③ Save
        ctk.CTkButton(p, text="③  Save Report",
                      height=36,
                      fg_color=C["glass_mid"],
                      hover_color=C["glass_light"],
                      border_color=C["border_dim"],
                      border_width=1,
                      text_color=C["text_sec"],
                      font=sf(11),
                      corner_radius=8,
                      command=self._save_report
                      ).pack(fill="x", padx=14, pady=(0, 12))

        # STATUS
        sep(p, pady=(4, 4))
        status_row = ctk.CTkFrame(p, fg_color="transparent")
        status_row.pack(fill="x", padx=14, pady=(4, 0))
        self._dot = PulseDot(status_row, color=C["success"])
        self._dot.pack(side="left", padx=(0, 6))
        self._status_lbl = ctk.CTkLabel(status_row, text="Ready",
                                        font=sf(12, "bold"),
                                        text_color=C["success"])
        self._status_lbl.pack(side="left")

        # Progress bar
        self._prog = GlowBar(p, height=7)
        self._prog.pack(fill="x", padx=14, pady=(8, 2))
        self._prog_lbl = ctk.CTkLabel(p, text="",
                                      font=sf(10),
                                      text_color=C["text_dim"])
        self._prog_lbl.pack(pady=(0, 14))

    # ── RIGHT PANEL ───────────────────────────────────────────────────────────
    def _build_right(self, p):
        p.columnconfigure(0, weight=1)
        p.rowconfigure(0, weight=0)  # table
        p.rowconfigure(1, weight=1)  # log

        # ── BOM Table ─────────────────────────────────────────────────────────
        tbl_hdr = ctk.CTkFrame(p, fg_color="transparent")
        tbl_hdr.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 2))
        ctk.CTkLabel(tbl_hdr, text="📋  Bill of Materials",
                     font=sf(13, "bold"),
                     text_color=C["text_primary"]).pack(side="left")
        ctk.CTkLabel(tbl_hdr, text="QUEUED",
                     font=sf(10),
                     text_color=C["text_dim"]).pack(side="right")

        tbl_frame = ctk.CTkFrame(p, fg_color=C["glass_mid"],
                                 corner_radius=10,
                                 border_width=1, border_color=C["border_dim"])
        tbl_frame.grid(row=0, column=0, sticky="ew", padx=14, pady=(0, 6))
        self._tbl_box = tk.Text(tbl_frame, height=7,
                                bg=C["glass_mid"], fg=C["text_primary"],
                                font=("Menlo", 11) if self._has_font("Menlo")
                                     else ("Cascadia Code", 11),
                                relief="flat", padx=10, pady=6,
                                state="disabled",
                                insertbackground=C["blue"],
                                selectbackground=C["blue_dim"])
        sb_t = ctk.CTkScrollbar(tbl_frame, command=self._tbl_box.yview,
                                fg_color=C["glass_card"],
                                button_color=C["glass_light"])
        self._tbl_box.configure(yscrollcommand=sb_t.set)
        sb_t.pack(side="right", fill="y")
        self._tbl_box.pack(fill="both", expand=True)

        # ── Log ───────────────────────────────────────────────────────────────
        log_hdr = ctk.CTkFrame(p, fg_color="transparent")
        log_hdr.grid(row=1, column=0, sticky="ew", padx=14, pady=(4, 2))
        ctk.CTkLabel(log_hdr, text="⚙  Batch Log",
                     font=sf(13, "bold"),
                     text_color=C["text_primary"]).pack(side="left")
        self._clear_btn = ctk.CTkButton(
            log_hdr, text="Clear", width=60, height=26,
            fg_color=C["glass_mid"], hover_color=C["glass_light"],
            border_width=1, border_color=C["border_dim"],
            text_color=C["text_dim"], font=sf(10),
            corner_radius=6, command=self._clear_log)
        self._clear_btn.pack(side="right")

        log_frame = ctk.CTkFrame(p, fg_color=C["bg_deep"],
                                 corner_radius=10,
                                 border_width=1, border_color=C["border_dim"])
        log_frame.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        self._log_box = tk.Text(log_frame,
                                bg=C["bg_deep"],
                                fg="#4DD2FF",   # cyan log text = Apple terminal feel
                                font=("Menlo", 11) if self._has_font("Menlo")
                                     else ("Cascadia Code", 11),
                                relief="flat", padx=10, pady=6,
                                state="disabled",
                                insertbackground=C["blue"],
                                selectbackground=C["blue_dim"])
        # colour tags
        self._log_box.tag_configure("ok",   foreground=C["success"])
        self._log_box.tag_configure("warn", foreground=C["warn"])
        self._log_box.tag_configure("err",  foreground=C["error"])
        self._log_box.tag_configure("sys",  foreground=C["indigo"])
        self._log_box.tag_configure("dim",  foreground=C["text_dim"])

        sb_l = ctk.CTkScrollbar(log_frame, command=self._log_box.yview,
                                fg_color=C["glass_dark"],
                                button_color=C["glass_mid"])
        self._log_box.configure(yscrollcommand=sb_l.set)
        sb_l.pack(side="right", fill="y")
        self._log_box.pack(fill="both", expand=True)

        self._log("[SYSTEM] Siraal 2D Engine — standby.", "sys")
        self._log("[SYSTEM] Awaiting BOM file…", "dim")

    # ── Helpers ───────────────────────────────────────────────────────────────
    @staticmethod
    def _has_font(name: str) -> bool:
        import tkinter.font as tkfont
        try:
            return name in tkfont.families()
        except Exception:
            return False

    def _log(self, msg: str, tag: str = ""):
        self._log_box.configure(state="normal")
        if tag:
            self._log_box.insert("end", msg + "\n", tag)
        else:
            self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def _update_table(self, parts, statuses):
        hdr = (f"{'#':<3}  {'Part No':<18} {'Type':<16} "
               f"{'Material':<12} {'Prio':<7}  Status\n")
        sep_line = "─" * 72 + "\n"
        body = "".join(
            f"{i+1:<3}  {str(p.get('Part_Number',''))[:17]:<18} "
            f"{str(p.get('Part_Type',''))[:15]:<16} "
            f"{str(p.get('Material',''))[:11]:<12} "
            f"{str(p.get('Priority',''))[:6]:<7}  "
            f"{statuses.get(p.get('Part_Number',''),'⏳ Queued')}\n"
            for i, p in enumerate(parts))
        self._tbl_box.configure(state="normal")
        self._tbl_box.delete("1.0", "end")
        self._tbl_box.insert("end", hdr + sep_line + body)
        self._tbl_box.configure(state="disabled")

    def _set_status(self, text: str, color: str):
        self._status_lbl.configure(text=text, text_color=color)
        self._dot.set_color(color)

    def _browse(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("All", "*.*")])
        if p:
            self._file_var.set(p)

    # ── Actions ───────────────────────────────────────────────────────────────
    def _run_validate(self):
        self._log(f"\n[BOM] {self._file_var.get()}", "dim")
        threading.Thread(target=self._t_validate, daemon=True).start()

    def _t_validate(self):
        from validator import EngineeringValidator
        v = EngineeringValidator(self._file_var.get(),
                                 log_callback=lambda m: self._q.put(("log", m)))
        v.run_checks()
        self._validator = v
        st = {p["Part_Number"]: "⏳ Queued" for p in v.valid_parts}
        self._q.put(("table", (v.valid_parts, st)))
        col = C["success"] if v.error_count == 0 else C["warn"]
        self._q.put(("status", (f"{len(v.valid_parts)} valid  ·  {v.error_count} errors", col)))

    def _run_batch(self):
        self._btn_run.configure(state="disabled")
        self._prog.set(0)
        self._dot.start_pulse()
        threading.Thread(target=self._t_batch, daemon=True).start()

    def _t_batch(self):
        from validator import EngineeringValidator
        from autocad_engine import AutoCADController

        self._q.put(("log_tag", ("\n[BOM] " + self._file_var.get(), "dim")))
        self._q.put(("progress", (0.05, "Validating…")))

        v = EngineeringValidator(self._file_var.get(),
                                 log_callback=lambda m: self._q.put(("log", m)))
        v.run_checks()
        self._validator = v

        parts = v.valid_parts
        ft = self._type_var.get()
        fp = self._prio_var.get()
        if ft != "All":
            parts = [p for p in parts if p.get("Part_Type") == ft]
        if fp != "All":
            parts = [p for p in parts if p.get("Priority") == fp]

        if not parts:
            self._q.put(("status", ("⚠ No parts after filters", C["warn"])))
            self._q.put(("btn", None)); return

        st = {p["Part_Number"]: "⏳ Queued" for p in parts}
        self._q.put(("table", (parts, dict(st))))
        self._q.put(("progress", (0.20, "Connecting to AutoCAD…")))

        try:
            eng = AutoCADController(log_callback=lambda m: self._q.put(("log", m)))
            idx = [0]
            orig = eng._log_info

            def tlog(msg):
                orig(msg)
                if "[*] GENERATING:" in msg and idx[0] < len(parts):
                    if idx[0] > 0:
                        st[parts[idx[0]-1]["Part_Number"]] = "✔ Done"
                    st[parts[idx[0]]["Part_Number"]] = "⚙ Drawing…"
                    self._q.put(("table", (parts, dict(st))))
                    self._q.put(("progress", (
                        0.20 + 0.75 * (idx[0] / len(parts)),
                        f"Drawing {parts[idx[0]]['Part_Number']} "
                        f"({idx[0]+1}/{len(parts)})")))
                    idx[0] += 1
                elif "ERROR drafting" in msg:
                    for k, v2 in st.items():
                        if v2 == "⚙ Drawing…": st[k] = "✘ Error"
                    self._q.put(("table", (parts, dict(st))))
            eng._log_info = tlog
            eng.generate_batch(parts)

            for k in st:
                if st[k] in ("⏳ Queued", "⚙ Drawing…"): st[k] = "✔ Done"
            self._q.put(("table", (parts, dict(st))))
            self._q.put(("progress", (1.0, "Complete!")))
            self._q.put(("status",
                (f"✔  {len(parts)} parts drawn", C["success"])))
            self._q.put(("log_tag", ("[✔] Batch complete.", "ok")))

        except Exception as e:
            self._q.put(("log_tag", (f"[ERROR] {e}", "err")))
            self._q.put(("status", ("Engine error", C["error"])))

        self._q.put(("btn", None))

    def _save_report(self):
        if not self._validator:
            messagebox.showinfo("Info", "Run validation first."); return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text", "*.txt")],
            initialfile="validation_2d_report.txt")
        if path:
            with open(path, "w") as f:
                f.write(self._validator.summary_report())
            messagebox.showinfo("Saved", f"Report saved:\n{path}")

    # ── Queue poller ──────────────────────────────────────────────────────────
    def _poll(self):
        try:
            while True:
                k, d = self._q.get_nowait()
                if k == "log":
                    tag = ("ok" if "[✔]" in d or "OK" in d
                           else "err" if "ERROR" in d or "✘" in d
                           else "warn" if "WARN" in d or "⚠" in d
                           else "sys" if "[SYSTEM]" in d
                           else "")
                    self._log(str(d), tag)
                elif k == "log_tag":
                    self._log(d[0], d[1])
                elif k == "table":
                    self._update_table(*d)
                elif k == "progress":
                    self._prog.set(d[0])
                    self._prog_lbl.configure(text=d[1])
                elif k == "status":
                    self._set_status(d[0], d[1])
                elif k == "btn":
                    self._btn_run.configure(state="normal")
                    self._dot.stop_pulse()
        except queue.Empty:
            pass
        self.after(100, self._poll)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = SiraalApple2D()
    app.mainloop()