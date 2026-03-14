"""
siraal_hub.py  —  Siraal Master Command Center
Apple Design + Glassmorphism + Glow Edition
============================================
Website-style dashboard. Sidebar nav. Canvas-rendered glass panels and
animated ambient glow orbs — all in pure CustomTkinter + tkinter.Canvas.
"""

import customtkinter as ctk
import tkinter as tk
import subprocess, sys, os, json, math, datetime
from tkinter import messagebox

ctk.set_appearance_mode("Dark")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# APPLE DARK PALETTE
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
C = {
    "bg":               "#080A12",
    "sidebar":          "#0C0F1A",
    "topbar":           "#0C0F1A",
    "glass":            "#101828",
    "glass2":           "#131C2E",
    "glass_border_ctk": "#1E2D44",
    "blue":             "#3B82F6",
    "blue_dim":         "#1E3A5F",
    "blue_glow":        "#0F1F3A",
    "teal":             "#2DD4BF",
    "teal_dim":         "#0D3530",
    "gold":             "#F59E0B",
    "gold_dim":         "#3D2800",
    "violet":           "#A78BFA",
    "violet_dim":       "#2E1F5E",
    "red":              "#F87171",
    "red_dim":          "#3D1515",
    "text":             "#F2F4F8",
    "text2":            "#8896AA",
    "text3":            "#3A4A60",
    "white":            "#FFFFFF",
}

WIN_W, WIN_H = 1020, 660
SIDEBAR_W    = 230
TOPBAR_H     = 58
CORNER       = 16


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# COLOR UTILITY
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def lerp_color(c1, c2, t):
    r1,g1,b1 = int(c1[1:3],16), int(c1[3:5],16), int(c1[5:7],16)
    r2,g2,b2 = int(c2[1:3],16), int(c2[3:5],16), int(c2[5:7],16)
    r = int(r1 + (r2-r1)*t)
    g = int(g1 + (g2-g1)*t)
    b = int(b1 + (b2-b1)*t)
    return f"#{r:02X}{g:02X}{b:02X}"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# AMBIENT GLOW CANVAS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class AmbientCanvas(tk.Canvas):
    ORB_DEFS = [
        (0.25, 0.30, 280, "#1A3A6E",  0.0003,  0.0002),
        (0.70, 0.60, 260, "#1A1A4E", -0.0002,  0.0003),
        (0.55, 0.15, 200, "#0D3030",  0.0004, -0.0002),
        (0.15, 0.75, 220, "#2D1060", -0.0003, -0.0003),
        (0.85, 0.25, 190, "#3D1A00",  0.0002,  0.0004),
    ]
    STEPS = 18

    def __init__(self, parent, **kw):
        super().__init__(parent, highlightthickness=0, bd=0, **kw)
        self._t = 0.0
        self._orbs = [list(o) for o in self.ORB_DEFS]
        self._after_id = None
        self.bind("<Configure>", lambda e: self._draw())
        self._draw()

    def _draw(self):
        if self._after_id:
            self.after_cancel(self._after_id)
        w = self.winfo_width()  or WIN_W
        h = self.winfo_height() or WIN_H
        self.delete("all")
        self.configure(bg=C["bg"])
        self._t += 1

        for orb in self._orbs:
            rx, ry, radius, color, sx, sy = orb
            cx = (rx + math.sin(self._t * sx * 60) * 0.12) * w
            cy = (ry + math.cos(self._t * sy * 60) * 0.12) * h
            for i in range(self.STEPS, 0, -1):
                t  = i / self.STEPS
                r  = int(radius * t)
                tc = lerp_color(color, C["bg"], t ** 0.6)
                self.create_oval(cx-r, cy-r, cx+r, cy+r,
                                 fill=tc, outline="")

        # Subtle grid
        for gx in range(0, w, 52):
            self.create_line(gx, 0, gx, h, fill="#111A26", width=1)
        for gy in range(0, h, 52):
            self.create_line(0, gy, w, gy, fill="#111A26", width=1)

        self._after_id = self.after(50, self._draw)

    def stop(self):
        if self._after_id:
            self.after_cancel(self._after_id)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# GLOWING SEPARATOR
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class GlowSep(tk.Canvas):
    def __init__(self, parent, color, bg_color=None, **kw):
        kw.setdefault("height", 1)
        self._bg_color = bg_color or C["sidebar"]
        super().__init__(parent, bg=self._bg_color,
                         highlightthickness=0, **kw)
        self._color = color
        self.bind("<Configure>", lambda e: self._draw())
        self._draw()

    def _draw(self):
        w = self.winfo_width() or 230
        self.delete("all")
        steps = 32
        for i in range(steps):
            t   = abs(i / steps - 0.5) * 2
            col = lerp_color(self._color, self._bg_color, t ** 0.35)
            x   = int(i / steps * w)
            x2  = int((i+1) / steps * w)
            self.create_line(x, 0, x2, 0, fill=col, width=1)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SIDEBAR NAV ITEM
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class NavItem(ctk.CTkFrame):
    def __init__(self, parent, icon, label, accent, command, active=False, **kw):
        super().__init__(parent, fg_color="transparent",
                         corner_radius=12, cursor="hand2", **kw)
        self._accent  = accent
        self._command = command
        self._active  = active

        self._pill = ctk.CTkFrame(self, corner_radius=10,
                                   fg_color=C["blue_glow"] if active else "transparent")
        self._pill.pack(fill="x", padx=6, pady=2)

        row = ctk.CTkFrame(self._pill, fg_color="transparent")
        row.pack(fill="x")

        self._bar = ctk.CTkFrame(row, fg_color=accent if active else "transparent",
                                  width=3, corner_radius=2)
        self._bar.pack(side="left", fill="y", pady=8)

        self._icon_lbl = ctk.CTkLabel(row, text=icon,
                                       font=ctk.CTkFont("Segoe UI Emoji", 15),
                                       text_color=accent if active else C["text3"],
                                       width=34)
        self._icon_lbl.pack(side="left", pady=10)

        self._lbl = ctk.CTkLabel(row, text=label,
                                  font=ctk.CTkFont("Segoe UI", 12,
                                                   "bold" if active else "normal"),
                                  text_color=C["text"] if active else C["text2"],
                                  anchor="w")
        self._lbl.pack(side="left", fill="x", expand=True, padx=(6, 10))

        for w in [self, self._pill, row, self._icon_lbl, self._lbl, self._bar]:
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<Button-1>", lambda e: self._command())

    def _on_enter(self, e=None):
        if not self._active:
            self._pill.configure(fg_color=C["glass2"])
            self._icon_lbl.configure(text_color=self._accent)
            self._lbl.configure(text_color=C["text"])

    def _on_leave(self, e=None):
        if not self._active:
            self._pill.configure(fg_color="transparent")
            self._icon_lbl.configure(text_color=C["text3"])
            self._lbl.configure(text_color=C["text2"])

    def set_active(self, v):
        self._active = v
        self._pill.configure(fg_color=C["blue_glow"] if v else "transparent")
        self._bar.configure(fg_color=self._accent if v else "transparent")
        self._icon_lbl.configure(text_color=self._accent if v else C["text3"])
        self._lbl.configure(
            font=ctk.CTkFont("Segoe UI", 12, "bold" if v else "normal"),
            text_color=C["text"] if v else C["text2"])


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STAT CARD
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class StatCard(ctk.CTkFrame):
    def __init__(self, parent, icon, title, value, sub, accent, dim, **kw):
        super().__init__(parent, fg_color=C["glass"], corner_radius=CORNER,
                         border_color=C["glass_border_ctk"], border_width=1, **kw)
        ctk.CTkFrame(self, fg_color=accent, height=2,
                     corner_radius=1).pack(fill="x", side="top")

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=14)

        ib = ctk.CTkFrame(body, fg_color=dim, corner_radius=10, width=40, height=40)
        ib.pack(side="left", anchor="n")
        ib.pack_propagate(False)
        ctk.CTkLabel(ib, text=icon,
                     font=ctk.CTkFont("Segoe UI Emoji", 18),
                     text_color=accent).pack(expand=True)

        txt = ctk.CTkFrame(body, fg_color="transparent")
        txt.pack(side="left", padx=(12, 0), fill="both", expand=True)
        ctk.CTkLabel(txt, text=title.upper(),
                     font=ctk.CTkFont("Consolas", 8, "bold"),
                     text_color=C["text2"]).pack(anchor="w")
        ctk.CTkLabel(txt, text=value,
                     font=ctk.CTkFont("Segoe UI", 22, "bold"),
                     text_color=C["text"]).pack(anchor="w", pady=(2, 0))
        ctk.CTkLabel(txt, text=sub,
                     font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["text3"]).pack(anchor="w")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# LAUNCH CARD
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class LaunchCard(ctk.CTkFrame):
    def __init__(self, parent, icon, title, desc, badge,
                 accent, dim, command, **kw):
        super().__init__(parent, fg_color=C["glass"], corner_radius=CORNER,
                         border_color=C["glass_border_ctk"], border_width=1,
                         cursor="hand2", **kw)
        self._accent  = accent
        self._command = command

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=16)

        iw = ctk.CTkFrame(body, fg_color=dim, corner_radius=14,
                          width=54, height=54)
        iw.pack(side="left", anchor="center")
        iw.pack_propagate(False)
        ctk.CTkLabel(iw, text=icon,
                     font=ctk.CTkFont("Segoe UI Emoji", 24),
                     text_color=accent).pack(expand=True)

        mid = ctk.CTkFrame(body, fg_color="transparent")
        mid.pack(side="left", fill="both", expand=True, padx=(16, 0))

        top_row = ctk.CTkFrame(mid, fg_color="transparent")
        top_row.pack(fill="x")
        ctk.CTkLabel(top_row, text=title,
                     font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["text"]).pack(side="left")
        bf = ctk.CTkFrame(top_row, fg_color=dim, corner_radius=6)
        bf.pack(side="left", padx=(10, 0))
        ctk.CTkLabel(bf, text=badge,
                     font=ctk.CTkFont("Consolas", 8, "bold"),
                     text_color=accent, padx=7, pady=2).pack()

        ctk.CTkLabel(mid, text=desc,
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["text2"],
                     anchor="w", justify="left",
                     wraplength=360).pack(anchor="w", pady=(5, 0))

        btn_txt_color = "#000000" if accent in (C["teal"], C["gold"]) else C["white"]
        ctk.CTkButton(body,
                      text="Open  ›",
                      font=ctk.CTkFont("Segoe UI", 11, "bold"),
                      fg_color=accent, hover_color=dim,
                      text_color=btn_txt_color,
                      corner_radius=20, width=80, height=32,
                      command=command).pack(side="right", anchor="center")

        for w in [self, body, mid, top_row, iw]:
            w.bind("<Enter>", self._hov_on)
            w.bind("<Leave>", self._hov_off)
            w.bind("<Button-1>", lambda e: command())

    def _hov_on(self, e=None):
        self.configure(fg_color=C["glass2"], border_color=self._accent)

    def _hov_off(self, e=None):
        self.configure(fg_color=C["glass"], border_color=C["glass_border_ctk"])


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ACTIVITY LOG
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class ActivityLog(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=C["glass"], corner_radius=CORNER,
                         border_color=C["glass_border_ctk"], border_width=1, **kw)
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=16, pady=(14, 0))
        ctk.CTkLabel(hdr, text="ACTIVITY LOG",
                     font=ctk.CTkFont("Consolas", 9, "bold"),
                     text_color=C["text2"]).pack(side="left")
        self._live = ctk.CTkLabel(hdr, text="● LIVE",
                                   font=ctk.CTkFont("Consolas", 8, "bold"),
                                   text_color=C["teal"])
        self._live.pack(side="right")
        ctk.CTkFrame(self, fg_color=C["glass_border_ctk"],
                     height=1).pack(fill="x", padx=16, pady=(8, 0))
        self._sf = ctk.CTkScrollableFrame(
            self, fg_color="transparent",
            scrollbar_button_color=C["glass_border_ctk"])
        self._sf.pack(fill="both", expand=True, padx=4, pady=4)
        self._blink = True
        self._pulse()

    def _pulse(self):
        self._blink = not self._blink
        self._live.configure(
            text_color=C["teal"] if self._blink else C["teal_dim"])
        self.after(900, self._pulse)

    def add(self, msg, accent=None, sym="›"):
        ts  = datetime.datetime.now().strftime("%H:%M:%S")
        row = ctk.CTkFrame(self._sf, fg_color="transparent")
        row.pack(fill="x", pady=3)
        ctk.CTkLabel(row, text=sym,
                     font=ctk.CTkFont("Segoe UI", 12),
                     text_color=accent or C["text2"],
                     width=16).pack(side="left")
        ctk.CTkLabel(row, text=msg,
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["text"], anchor="w").pack(side="left", padx=(6, 0))
        ctk.CTkLabel(row, text=ts,
                     font=ctk.CTkFont("Consolas", 9),
                     text_color=C["text3"]).pack(side="right", padx=4)
        try:
            self._sf._parent_canvas.yview_moveto(1.0)
        except Exception:
            pass


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# LIVE CLOCK
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class LiveClock(ctk.CTkLabel):
    def __init__(self, parent, **kw):
        super().__init__(parent, text="",
                         font=ctk.CTkFont("Consolas", 11),
                         text_color=C["text2"], **kw)
        self._tick()

    def _tick(self):
        self.configure(
            text=datetime.datetime.now().strftime("%a %d %b  %H:%M:%S"))
        self.after(1000, self._tick)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# MAIN APPLICATION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class SiraalHub(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Siraal  —  Command Center")
        self.geometry(f"{WIN_W}x{WIN_H}")
        self.minsize(WIN_W, WIN_H)
        self.configure(fg_color=C["bg"])
        self.eval("tk::PlaceWindow . center")
        self._ensure_rules_file()
        self._nav_items = []
        self._pages     = {}
        self._active    = None
        self._build_layout()
        self._nav_to("dashboard")

    # ── Build Layout ──────────────────────────────────────────────────────────
    def _build_layout(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Ambient background
        self._amb = AmbientCanvas(self, bg=C["bg"])
        self._amb.place(x=0, y=0, relwidth=1, relheight=1)

        # ── MAIN COLUMN ──
        main = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        main.grid(row=0, column=0, sticky="nsew")
        main.grid_rowconfigure(1, weight=1)
        main.grid_columnconfigure(0, weight=1)

        # Top bar
        topbar = ctk.CTkFrame(main, fg_color=C["topbar"],
                              height=TOPBAR_H, corner_radius=0)
        topbar.grid(row=0, column=0, sticky="ew")
        topbar.grid_propagate(False)

        GlowSep(topbar, C["blue_dim"], bg_color=C["topbar"]
                ).place(relx=0, rely=1.0, relwidth=1, anchor="sw")

        tb_l = ctk.CTkFrame(topbar, fg_color="transparent")
        tb_l.pack(side="left", padx=24, pady=10)
        self._page_title = ctk.CTkLabel(tb_l, text="Dashboard",
                                         font=ctk.CTkFont("Segoe UI", 17, "bold"),
                                         text_color=C["text"])
        self._page_title.pack(anchor="w")
        self._breadcrumb = ctk.CTkLabel(tb_l, text="Home",
                                         font=ctk.CTkFont("Consolas", 9),
                                         text_color=C["text3"])
        self._breadcrumb.pack(anchor="w")

        tb_r = ctk.CTkFrame(topbar, fg_color="transparent")
        tb_r.pack(side="right", padx=20)
        LiveClock(tb_r).pack(side="right", padx=(12, 0))

        self._pill = ctk.CTkFrame(tb_r, fg_color=C["blue_glow"],
                                   corner_radius=20)
        self._pill.pack(side="right")
        self._pill_dot = ctk.CTkLabel(self._pill, text="●",
                                       font=ctk.CTkFont("Segoe UI", 9),
                                       text_color=C["blue"], padx=4)
        self._pill_dot.pack(side="left", padx=(8, 0))
        self._pill_lbl = ctk.CTkLabel(self._pill, text="READY",
                                       font=ctk.CTkFont("Consolas", 9, "bold"),
                                       text_color=C["blue"])
        self._pill_lbl.pack(side="left", padx=(0, 10), pady=5)

        # Content area
        self._content = ctk.CTkFrame(main, fg_color="transparent",
                                      corner_radius=0)
        self._content.grid(row=1, column=0, sticky="nsew")
        self._content.grid_rowconfigure(0, weight=1)
        self._content.grid_columnconfigure(0, weight=1)

        # Pages
        self._pages["dashboard"] = self._build_dashboard(self._content)
        self._pages["2d"]   = self._build_module_page(
            self._content, "2D Engineering Workspace",
            "Flat-pattern layout, DXF export and 2D gear/component design engine.",
            C["teal"], C["teal_dim"], "📐", self._launch_2d)
        self._pages["3d"]   = self._build_module_page(
            self._content, "3D Solid Modelling Studio",
            "Volumetric part design, assembly tree, STL and STEP export pipeline.",
            C["gold"], C["gold_dim"], "🧊", self._launch_3d)
        self._pages["rules"] = self._build_module_page(
            self._content, "Factory Rule Editor",
            "Define, edit and validate manufacturing constraints and tolerance rules.",
            C["violet"], C["violet_dim"], "🛠", self._launch_rules)

    def _sb_section(self, parent, label):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", padx=18, pady=(18, 4))
        ctk.CTkLabel(f, text=label,
                     font=ctk.CTkFont("Consolas", 8, "bold"),
                     text_color=C["text3"]).pack(anchor="w")

    # ── Dashboard ─────────────────────────────────────────────────────────────
    def _build_dashboard(self, parent):
        page = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(0, weight=1)

        sf = ctk.CTkScrollableFrame(
            page, fg_color="transparent",
            scrollbar_button_color=C["glass_border_ctk"],
            scrollbar_button_hover_color=C["blue_dim"])
        sf.grid(row=0, column=0, sticky="nsew")
        sf.grid_columnconfigure((0, 1, 2), weight=1)

        # Hero banner
        banner = ctk.CTkFrame(sf, fg_color=C["glass"],
                              corner_radius=20,
                              border_color=C["blue_dim"],
                              border_width=1)
        banner.grid(row=0, column=0, columnspan=3, sticky="ew",
                    padx=24, pady=(24, 0))
        ctk.CTkFrame(banner, fg_color=C["blue"], height=2,
                     corner_radius=1).pack(fill="x", side="top")

        bi = ctk.CTkFrame(banner, fg_color="transparent")
        bi.pack(fill="x", padx=24, pady=18)

        bl = ctk.CTkFrame(bi, fg_color="transparent")
        bl.pack(side="left", fill="both", expand=True)
        ctk.CTkLabel(bl, text="Good day, Engineer.",
                     font=ctk.CTkFont("Segoe UI", 20, "bold"),
                     text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(bl,
                     text="Siraal Command Center is online. All modules are ready to launch.",
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["text2"]).pack(anchor="w", pady=(4, 0))

        ctk.CTkLabel(bi, text="⚙",
                     font=ctk.CTkFont("Segoe UI Emoji", 48),
                     text_color=C["blue_glow"]).pack(side="right", padx=(0, 10))

        # Stat cards
        self._sec_lbl(sf, "QUICK STATUS", row=1)
        for col, (icon, title, val, sub, acc, dim) in enumerate([
            ("📐", "2D Module",  "Ready",  "Flat-pattern engine active", C["teal"],   C["teal_dim"]),
            ("🧊", "3D Module",  "Ready",  "Solid modelling engine active", C["gold"], C["gold_dim"]),
            ("🛠", "Rule Set",   "Ready", "custom_rules.json active",   C["violet"], C["violet_dim"]),
        ]):
            StatCard(sf, icon=icon, title=title, value=val,
                     sub=sub, accent=acc, dim=dim).grid(
                row=2, column=col, sticky="ew",
                padx=(24 if col == 0 else 8, 8 if col < 2 else 24),
                pady=(0, 16))

        # Launch cards
        self._sec_lbl(sf, "LAUNCH MODULE", row=3)
        for i, (icon, title, desc, badge, acc, dim, cmd) in enumerate([
            ("📐", "2D Engineering Workspace",
             "Open the flat-pattern editor, DXF export tools and 2D layout engine.",
             "2D",  C["teal"],   C["teal_dim"],   self._launch_2d),
            ("🧊", "3D Solid Modelling Studio",
             "Launch the volumetric part designer, assembly navigator and export pipeline.",
             "3D",  C["gold"],   C["gold_dim"],   self._launch_3d),
            ("🛠", "Factory Rule Editor",
             "Manage validation rules, tolerance limits and manufacturing constraints.",
             "CFG", C["violet"], C["violet_dim"], self._launch_rules),
        ]):
            LaunchCard(sf, icon=icon, title=title, desc=desc, badge=badge,
                       accent=acc, dim=dim, command=cmd).grid(
                row=4+i, column=0, columnspan=3, sticky="ew",
                padx=24, pady=(0, 10))

       
        return page

    def _sec_lbl(self, parent, text, row):
        ctk.CTkLabel(parent, text=text,
                     font=ctk.CTkFont("Consolas", 9, "bold"),
                     text_color=C["text3"]).grid(
            row=row, column=0, columnspan=3, sticky="w",
            padx=26, pady=(12, 6))

    # ── Module Page ───────────────────────────────────────────────────────────
    def _build_module_page(self, parent, title, desc, accent, dim, icon, cmd):
        page = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(0, weight=1)

        card = ctk.CTkFrame(page, fg_color=C["glass"],
                            corner_radius=24,
                            border_color=accent,
                            border_width=1,
                            width=460, height=360)
        card.place(relx=0.5, rely=0.5, anchor="center")
        card.pack_propagate(False)

        ctk.CTkFrame(card, fg_color=accent, height=2,
                     corner_radius=1).pack(fill="x", side="top")

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(expand=True, fill="both", padx=40, pady=36)

        ib = ctk.CTkFrame(inner, fg_color=dim, corner_radius=20,
                          width=72, height=72)
        ib.pack()
        ib.pack_propagate(False)
        ctk.CTkLabel(ib, text=icon,
                     font=ctk.CTkFont("Segoe UI Emoji", 34),
                     text_color=accent).pack(expand=True)

        ctk.CTkLabel(inner, text=title,
                     font=ctk.CTkFont("Segoe UI", 19, "bold"),
                     text_color=C["text"]).pack(pady=(18, 6))
        ctk.CTkLabel(inner, text=desc,
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["text2"],
                     wraplength=360, justify="center").pack()

        btn_color = "#000000" if accent in (C["teal"], C["gold"]) else C["white"]
        ctk.CTkButton(inner, text="Launch  ›",
                      font=ctk.CTkFont("Segoe UI", 13, "bold"),
                      fg_color=accent, hover_color=dim,
                      text_color=btn_color,
                      corner_radius=22, height=44, width=200,
                      command=cmd).pack(pady=(24, 0))

        back = ctk.CTkLabel(inner, text="← Back to Dashboard",
                            font=ctk.CTkFont("Segoe UI", 11),
                            text_color=C["text3"], cursor="hand2")
        back.pack(pady=(12, 0))
        back.bind("<Button-1>", lambda e: self._nav_to("dashboard"))
        back.bind("<Enter>", lambda e: back.configure(text_color=C["blue"]))
        back.bind("<Leave>", lambda e: back.configure(text_color=C["text3"]))

        return page

    # ── Navigation ────────────────────────────────────────────────────────────
    def _nav_to(self, key):
        titles = {
            "dashboard": ("Dashboard",          "Home"),
            "2d":        ("2D Engineering",     "Home  ›  2D Workspace"),
            "3d":        ("3D Solid Studio",    "Home  ›  3D Studio"),
            "rules":     ("Factory Rules",      "Home  ›  Administration"),
        }
        if self._active:
            self._pages[self._active].grid_remove()
        self._pages[key].grid(row=0, column=0, sticky="nsew")
        self._active = key
        t, bc = titles.get(key, (key, "Home"))
        self._page_title.configure(text=t)
        self._breadcrumb.configure(text=bc)
        for k, item in self._nav_items:
            item.set_active(k == key)

    # ── Launchers ─────────────────────────────────────────────────────────────
    _PILL = {
        "teal":   (None,        None),   # filled at runtime
        "gold":   (None,        None),
        "violet": (None,        None),
        "red":    (None,        None),
    }

    def _launch(self, filename, label, accent, pill_bg):
        if not os.path.exists(filename):
            messagebox.showerror("Module Not Found",
                                 f"'{filename}' not found in working directory.")
            self._set_status("NOT FOUND", C["red"], C["red_dim"])
            if hasattr(self, "_log"):
                self._log.add(f"FAILED: {filename} missing", C["red"], "✕")
            return
        subprocess.Popen([sys.executable, filename])
        self._set_status("LAUNCHED", accent, pill_bg)
        if hasattr(self, "_log"):
            self._log.add(f"Launched: {label}", accent, "▶")

    def _set_status(self, text, color, pill_bg):
        self._pill_lbl.configure(text=text,  text_color=color)
        self._pill_dot.configure(text_color=color)
        self._pill.configure(fg_color=pill_bg)
        self.after(4000, lambda: (
            self._pill_lbl.configure(text="READY",   text_color=C["blue"]),
            self._pill_dot.configure(text_color=C["blue"]),
            self._pill.configure(fg_color=C["blue_glow"]),
        ))

    def _launch_2d(self):
        self._launch("gui_launcher.py",    "2D Workspace",  C["teal"],   C["teal_dim"])
    def _launch_3d(self):
        self._launch("gui_launcher_3d.py", "3D Studio",     C["gold"],   C["gold_dim"])
    def _launch_rules(self):
        self._launch("gui_launcher_val.py","Factory Rules", C["violet"], C["violet_dim"])

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _ensure_rules_file(self):
        if not os.path.exists("custom_rules.json"):
            default = {"rules": [{
                "rule_id":         "EXAMPLE_RULE_01",
                "target_type":     "Spur_Gear_3D",
                "target_material": "ALL",
                "condition":       "P3 > 200",
                "severity":        "ERROR",
                "message":         "Face width (P3) exceeds 200mm machine limit."
            }]}
            with open("custom_rules.json", "w") as f:
                json.dump(default, f, indent=4)

    def destroy(self):
        if hasattr(self, "_amb"):
            self._amb.stop()
        super().destroy()


if __name__ == "__main__":
    app = SiraalHub()
    app.mainloop()