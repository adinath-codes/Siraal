"""
gui_launcher.py — Siraal Manufacturing Engine — Dual-Mode GUI
Tab 1: 2D Drafting Engine  (autocad_engine.py  + validator.py)
Tab 2: 3D Solid Builder     (autocad_engine_3d.py + validator_3d.py)

PS compliance:
  • 3D solids → Layout viewports (Front/Top/Right/ISO)
  • ISO title block on all sheets
  • CAD naming, layers, drafting standards (IS 11669 / ISO 128)
  • Maintainable via Excel templates
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import queue

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

COL = {
    "bg":       "#0D1B2A",
    "panel":    "#1A2B3C",
    "card":     "#1E3448",
    "accent2d": "#1ABC9C",
    "accent3d": "#E67E22",
    "gold":     "#F4C842",
    "success":  "#2ECC71",
    "error":    "#E74C3C",
    "warn":     "#F39C12",
    "text":     "#ECF0F1",
    "subtext":  "#95A5A6",
    "border":   "#2C4A6B",
}


class SiraalGUI(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("SIRAAL Manufacturing Engine  |  TN-IMPACT 2026")
        self.geometry("1380x860")
        self.minsize(1200, 780)
        self.configure(fg_color=COL["bg"])
        self._q2d: queue.Queue = queue.Queue()
        self._q3d: queue.Queue = queue.Queue()
        self._2d_validator = None
        self._3d_validator = None
        self._build_header()
        self._build_tabs()
        self._build_2d_tab()
        self._build_3d_tab()
        self.after(120, self._poll_2d)
        self.after(130, self._poll_3d)

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color=COL["panel"], corner_radius=0, height=56)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="⚙  SIRAAL MANUFACTURING ENGINE",
                     font=ctk.CTkFont("Calibri", 20, "bold"),
                     text_color=COL["gold"]).pack(side="left", padx=20, pady=10)
        ctk.CTkLabel(hdr,
                     text="TN-IMPACT 2026  |  IS 11669 / ISO 128  |  1st Angle Projection",
                     font=ctk.CTkFont("Calibri", 11),
                     text_color=COL["subtext"]).pack(side="left", padx=4)
        ctk.CTkLabel(hdr, text="2D + 3D  v3.0",
                     font=ctk.CTkFont("Calibri", 10, "bold"),
                     text_color=COL["accent3d"]).pack(side="right", padx=20)

    # ── Tabs ──────────────────────────────────────────────────────────────────
    def _build_tabs(self):
        self._tabs = ctk.CTkTabview(
            self, fg_color=COL["bg"],
            segmented_button_fg_color=COL["panel"],
            segmented_button_selected_color=COL["accent2d"],
            segmented_button_selected_hover_color=COL["accent2d"],
            segmented_button_unselected_color=COL["panel"],
            text_color=COL["text"],
            corner_radius=0)
        self._tabs.pack(fill="both", expand=True)
        self._tabs.add("  📐  2D Drafting Engine  ")
        self._tabs.add("  🧊  3D Solid Builder   ")

    # ════════════════════════════════════════════════════════════════════════
    # 2D TAB
    # ════════════════════════════════════════════════════════════════════════
    def _build_2d_tab(self):
        tab = self._tabs.tab("  📐  2D Drafting Engine  ")
        tab.configure(fg_color=COL["bg"])
        root = ctk.CTkFrame(tab, fg_color=COL["bg"])
        root.pack(fill="both", expand=True, padx=10, pady=8)
        root.columnconfigure(0, weight=0, minsize=280)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)
        lp = ctk.CTkFrame(root, fg_color=COL["panel"], corner_radius=10, width=280)
        lp.grid(row=0, column=0, sticky="ns", padx=(0, 8))
        lp.pack_propagate(False)
        self._build_2d_left(lp)
        rp = ctk.CTkFrame(root, fg_color=COL["panel"], corner_radius=10)
        rp.grid(row=0, column=1, sticky="nsew")
        self._build_2d_right(rp)

    def _build_2d_left(self, p):
        ctk.CTkLabel(p, text="📐  2D DRAFTING ENGINE",
                     font=ctk.CTkFont("Calibri", 13, "bold"),
                     text_color=COL["accent2d"]).pack(pady=(16, 4), padx=12)
        ctk.CTkLabel(p, text="Plates · Gears · Shafts · Ring Gears",
                     font=ctk.CTkFont("Calibri", 9),
                     text_color=COL["subtext"]).pack(pady=(0, 12))

        self._section(p, "BOM FILE")
        self._2d_file_var = ctk.StringVar(value="excels/demo.xlsx")
        ctk.CTkEntry(p, textvariable=self._2d_file_var, fg_color=COL["card"],
                     border_color=COL["border"], text_color=COL["text"],
                     height=30, font=ctk.CTkFont("Calibri", 9)
                     ).pack(fill="x", padx=12, pady=(0, 4))
        ctk.CTkButton(p, text="Browse…", height=28, fg_color=COL["card"],
                      hover_color=COL["border"], text_color=COL["subtext"],
                      font=ctk.CTkFont("Calibri", 9),
                      command=self._browse_2d).pack(fill="x", padx=12, pady=(0, 12))

        self._section(p, "FILTERS")
        ctk.CTkLabel(p, text="Part Type", font=ctk.CTkFont("Calibri", 9),
                     text_color=COL["subtext"]).pack(anchor="w", padx=14)
        self._2d_type_var = ctk.StringVar(value="All")
        ctk.CTkComboBox(p, variable=self._2d_type_var,
                        values=["All","Plate","Spur_Gear","Ring_Gear",
                                "Stepped_Shaft","Flanged_Shaft"],
                        fg_color=COL["card"], button_color=COL["border"],
                        border_color=COL["border"], text_color=COL["text"],
                        font=ctk.CTkFont("Calibri", 9), height=28
                        ).pack(fill="x", padx=12, pady=(2, 6))
        ctk.CTkLabel(p, text="Priority", font=ctk.CTkFont("Calibri", 9),
                     text_color=COL["subtext"]).pack(anchor="w", padx=14)
        self._2d_prio_var = ctk.StringVar(value="All")
        ctk.CTkComboBox(p, variable=self._2d_prio_var,
                        values=["All","High","Medium","Low"],
                        fg_color=COL["card"], button_color=COL["border"],
                        border_color=COL["border"], text_color=COL["text"],
                        font=ctk.CTkFont("Calibri", 9), height=28
                        ).pack(fill="x", padx=12, pady=(2, 12))

        self._section(p, "ACTIONS")
        ctk.CTkButton(p, text="① Validate 2D BOM", height=34,
                      fg_color=COL["card"], hover_color="#27AE60",
                      border_color=COL["accent2d"], border_width=1,
                      text_color=COL["accent2d"],
                      font=ctk.CTkFont("Calibri", 10, "bold"),
                      command=self._run_2d_validate
                      ).pack(fill="x", padx=12, pady=(0, 6))
        self._btn_2d_run = ctk.CTkButton(
            p, text="② RUN 2D CAD BATCH", height=38,
            fg_color=COL["accent2d"], hover_color="#17A589",
            text_color="#000000", font=ctk.CTkFont("Calibri", 11, "bold"),
            command=self._run_2d_batch)
        self._btn_2d_run.pack(fill="x", padx=12, pady=(0, 6))
        ctk.CTkButton(p, text="③ Save Report", height=30,
                      fg_color=COL["card"], hover_color=COL["border"],
                      text_color=COL["subtext"], font=ctk.CTkFont("Calibri", 9),
                      command=self._save_2d_report
                      ).pack(fill="x", padx=12, pady=(0, 12))

        self._2d_status = ctk.CTkLabel(p, text="● Ready",
                                       font=ctk.CTkFont("Calibri", 10, "bold"),
                                       text_color=COL["success"])
        self._2d_status.pack(pady=4)
        self._2d_prog = ctk.CTkProgressBar(p, fg_color=COL["card"],
                                           progress_color=COL["accent2d"], height=8)
        self._2d_prog.set(0)
        self._2d_prog.pack(fill="x", padx=12, pady=(0, 2))
        self._2d_prog_lbl = ctk.CTkLabel(p, text="", font=ctk.CTkFont("Calibri", 8),
                                         text_color=COL["subtext"])
        self._2d_prog_lbl.pack(pady=(0, 12))

    def _build_2d_right(self, p):
        p.columnconfigure(0, weight=1)
        p.rowconfigure(0, weight=0)
        p.rowconfigure(1, weight=1)
        p.rowconfigure(2, weight=1)
        tf = ctk.CTkFrame(p, fg_color=COL["card"], corner_radius=8)
        tf.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
        self._2d_tbl = tk.Text(tf, height=8, bg=COL["card"], fg=COL["text"],
                               font=("Cascadia Code", 8), relief="flat",
                               padx=8, pady=4, state="disabled")
        sb = ctk.CTkScrollbar(tf, command=self._2d_tbl.yview)
        self._2d_tbl.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._2d_tbl.pack(fill="both", expand=True)

        ctk.CTkLabel(p, text="  ⚙ BATCH LOG", font=ctk.CTkFont("Calibri", 10, "bold"),
                     text_color=COL["accent2d"]).grid(row=1, column=0, sticky="w", padx=14)
        lf = ctk.CTkFrame(p, fg_color=COL["bg"], corner_radius=8)
        lf.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self._2d_log_box = tk.Text(lf, bg="#0A0F14", fg="#00FF88",
                                   font=("Cascadia Code", 8), relief="flat",
                                   padx=8, pady=4, state="disabled")
        sb2 = ctk.CTkScrollbar(lf, command=self._2d_log_box.yview)
        self._2d_log_box.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right", fill="y")
        self._2d_log_box.pack(fill="both", expand=True)
        self._log2d("[SYSTEM] Siraal 2D Engine — standby.\n")

    # ════════════════════════════════════════════════════════════════════════
    # 3D TAB
    # ════════════════════════════════════════════════════════════════════════
    def _build_3d_tab(self):
        tab = self._tabs.tab("  🧊  3D Solid Builder   ")
        tab.configure(fg_color=COL["bg"])
        root = ctk.CTkFrame(tab, fg_color=COL["bg"])
        root.pack(fill="both", expand=True, padx=10, pady=8)
        root.columnconfigure(0, weight=0, minsize=280)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)
        lp = ctk.CTkFrame(root, fg_color=COL["panel"], corner_radius=10, width=280)
        lp.grid(row=0, column=0, sticky="ns", padx=(0, 8))
        lp.pack_propagate(False)
        self._build_3d_left(lp)
        rp = ctk.CTkFrame(root, fg_color=COL["panel"], corner_radius=10)
        rp.grid(row=0, column=1, sticky="nsew")
        self._build_3d_right(rp)

    def _build_3d_left(self, p):
        ctk.CTkLabel(p, text="🧊  3D SOLID BUILDER",
                     font=ctk.CTkFont("Calibri", 13, "bold"),
                     text_color=COL["accent3d"]).pack(pady=(16, 4), padx=12)
        ctk.CTkLabel(p, text="Box · Cylinder · Cone · Sphere\nFlanged Boss · Extruded · Revolved",
                     font=ctk.CTkFont("Calibri", 9), text_color=COL["subtext"],
                     justify="center").pack(pady=(0, 6))

        badge = ctk.CTkFrame(p, fg_color="#1A3A1A", corner_radius=6)
        badge.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkLabel(badge,
                     text="PS: 3D Solids → Layout Tabs\nFront · Top · Right · ISO Views\n+ IS 11669 Title Block per Sheet",
                     font=ctk.CTkFont("Calibri", 8), text_color="#2ECC71",
                     justify="center").pack(pady=6, padx=6)

        self._section(p, "3D BOM FILE")
        self._3d_file_var = ctk.StringVar(value="excels/demo_3d.xlsx")
        ctk.CTkEntry(p, textvariable=self._3d_file_var, fg_color=COL["card"],
                     border_color=COL["border"], text_color=COL["text"],
                     height=30, font=ctk.CTkFont("Calibri", 9)
                     ).pack(fill="x", padx=12, pady=(0, 4))
        ctk.CTkButton(p, text="Browse…", height=28, fg_color=COL["card"],
                      hover_color=COL["border"], text_color=COL["subtext"],
                      font=ctk.CTkFont("Calibri", 9),
                      command=self._browse_3d).pack(fill="x", padx=12, pady=(0, 12))

        self._section(p, "FILTERS")
        ctk.CTkLabel(p, text="Part Type", font=ctk.CTkFont("Calibri", 9),
                     text_color=COL["subtext"]).pack(anchor="w", padx=14)
        self._3d_type_var = ctk.StringVar(value="All")
        ctk.CTkComboBox(p, variable=self._3d_type_var,
                        values=["All","Box","Cylinder","Cone","Sphere",
                                "Flanged_Boss","Extruded_Profile","Revolved_Part"],
                        fg_color=COL["card"], button_color=COL["border"],
                        border_color=COL["border"], text_color=COL["text"],
                        font=ctk.CTkFont("Calibri", 9), height=28
                        ).pack(fill="x", padx=12, pady=(2, 6))
        ctk.CTkLabel(p, text="Priority", font=ctk.CTkFont("Calibri", 9),
                     text_color=COL["subtext"]).pack(anchor="w", padx=14)
        self._3d_prio_var = ctk.StringVar(value="All")
        ctk.CTkComboBox(p, variable=self._3d_prio_var,
                        values=["All","High","Medium","Low"],
                        fg_color=COL["card"], button_color=COL["border"],
                        border_color=COL["border"], text_color=COL["text"],
                        font=ctk.CTkFont("Calibri", 9), height=28
                        ).pack(fill="x", padx=12, pady=(2, 12))

        self._section(p, "ACTIONS")
        ctk.CTkButton(p, text="① Validate 3D BOM", height=34,
                      fg_color=COL["card"], hover_color="#D35400",
                      border_color=COL["accent3d"], border_width=1,
                      text_color=COL["accent3d"],
                      font=ctk.CTkFont("Calibri", 10, "bold"),
                      command=self._run_3d_validate
                      ).pack(fill="x", padx=12, pady=(0, 6))
        self._btn_3d_run = ctk.CTkButton(
            p, text="② BUILD 3D SOLIDS + VIEWS", height=38,
            fg_color=COL["accent3d"], hover_color="#CA6F1E",
            text_color="#000000", font=ctk.CTkFont("Calibri", 11, "bold"),
            command=self._run_3d_batch)
        self._btn_3d_run.pack(fill="x", padx=12, pady=(0, 6))
        ctk.CTkButton(p, text="③ Save 3D Report", height=30,
                      fg_color=COL["card"], hover_color=COL["border"],
                      text_color=COL["subtext"], font=ctk.CTkFont("Calibri", 9),
                      command=self._save_3d_report
                      ).pack(fill="x", padx=12, pady=(0, 8))

        info = ctk.CTkFrame(p, fg_color=COL["card"], corner_radius=6)
        info.pack(fill="x", padx=12, pady=(0, 8))
        for line in ["📁 output_3d/Master_3D_Assembly.dwg",
                     "📁 output_3d/DXF_Files/*_3D.dxf",
                     "📋 Layout per part: DRW_<PARTNO>",
                     "👁 Views: Front · Top · Right · ISO"]:
            ctk.CTkLabel(info, text=line, font=ctk.CTkFont("Calibri", 8),
                         text_color=COL["subtext"],
                         anchor="w").pack(anchor="w", padx=8, pady=1)

        self._3d_status = ctk.CTkLabel(p, text="● Ready",
                                       font=ctk.CTkFont("Calibri", 10, "bold"),
                                       text_color=COL["success"])
        self._3d_status.pack(pady=4)
        self._3d_prog = ctk.CTkProgressBar(p, fg_color=COL["card"],
                                           progress_color=COL["accent3d"], height=8)
        self._3d_prog.set(0)
        self._3d_prog.pack(fill="x", padx=12, pady=(0, 2))
        self._3d_prog_lbl = ctk.CTkLabel(p, text="", font=ctk.CTkFont("Calibri", 8),
                                         text_color=COL["subtext"])
        self._3d_prog_lbl.pack(pady=(0, 12))

    def _build_3d_right(self, p):
        p.columnconfigure(0, weight=1)
        p.rowconfigure(0, weight=0)
        p.rowconfigure(1, weight=1)
        p.rowconfigure(2, weight=1)
        tf = ctk.CTkFrame(p, fg_color=COL["card"], corner_radius=8)
        tf.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
        self._3d_tbl = tk.Text(tf, height=8, bg=COL["card"], fg=COL["text"],
                               font=("Cascadia Code", 8), relief="flat",
                               padx=8, pady=4, state="disabled")
        sb = ctk.CTkScrollbar(tf, command=self._3d_tbl.yview)
        self._3d_tbl.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._3d_tbl.pack(fill="both", expand=True)

        ctk.CTkLabel(p, text="  🧊 3D BUILD LOG", font=ctk.CTkFont("Calibri", 10, "bold"),
                     text_color=COL["accent3d"]).grid(row=1, column=0, sticky="w", padx=14)
        lf = ctk.CTkFrame(p, fg_color=COL["bg"], corner_radius=8)
        lf.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self._3d_log_box = tk.Text(lf, bg="#0A0A0F", fg="#FFA500",
                                   font=("Cascadia Code", 8), relief="flat",
                                   padx=8, pady=4, state="disabled")
        sb2 = ctk.CTkScrollbar(lf, command=self._3d_log_box.yview)
        self._3d_log_box.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right", fill="y")
        self._3d_log_box.pack(fill="both", expand=True)
        self._log3d("[SYSTEM] Siraal 3D Engine — standby.\n")
        self._log3d("[PS] Output: 3D Solids + per-part Layout tabs\n")
        self._log3d("[PS] Standard: IS 11669 / ISO 128 / 1st Angle\n")

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _section(self, parent, label):
        ctk.CTkFrame(parent, fg_color=COL["border"], height=1,
                     corner_radius=0).pack(fill="x", padx=12, pady=(8, 0))
        ctk.CTkLabel(parent, text=label, font=ctk.CTkFont("Calibri", 8, "bold"),
                     text_color=COL["subtext"]).pack(anchor="w", padx=14, pady=(2, 2))

    def _log2d(self, msg):
        self._2d_log_box.configure(state="normal")
        self._2d_log_box.insert("end", msg + ("" if msg.endswith("\n") else "\n"))
        self._2d_log_box.see("end")
        self._2d_log_box.configure(state="disabled")

    def _log3d(self, msg):
        self._3d_log_box.configure(state="normal")
        self._3d_log_box.insert("end", msg + ("" if msg.endswith("\n") else "\n"))
        self._3d_log_box.see("end")
        self._3d_log_box.configure(state="disabled")

    def _tbl2d(self, parts, statuses):
        hdr = f"{'#':<3} {'Part No':<18} {'Type':<16} {'Material':<12} {'Prio':<7} {'Status'}\n"
        sep = "─" * 72 + "\n"
        body = "".join(
            f"{i+1:<3} {str(p.get('Part_Number',''))[:17]:<18} "
            f"{str(p.get('Part_Type',''))[:15]:<16} "
            f"{str(p.get('Material',''))[:11]:<12} "
            f"{str(p.get('Priority',''))[:6]:<7} "
            f"{statuses.get(p.get('Part_Number',''),'⏳ Queued')}\n"
            for i, p in enumerate(parts))
        self._2d_tbl.configure(state="normal")
        self._2d_tbl.delete("1.0", "end")
        self._2d_tbl.insert("end", hdr + sep + body)
        self._2d_tbl.configure(state="disabled")

    def _tbl3d(self, parts, statuses):
        hdr = f"{'#':<3} {'Part No':<18} {'Type':<18} {'Material':<12} {'Layout':<16} {'Status'}\n"
        sep = "─" * 80 + "\n"
        body = "".join(
            f"{i+1:<3} {str(p.get('Part_Number',''))[:17]:<18} "
            f"{str(p.get('Part_Type',''))[:17]:<18} "
            f"{str(p.get('Material',''))[:11]:<12} "
            f"DRW_{str(p.get('Part_Number',''))[:10]:<16} "
            f"{statuses.get(p.get('Part_Number',''),'⏳ Queued')}\n"
            for i, p in enumerate(parts))
        self._3d_tbl.configure(state="normal")
        self._3d_tbl.delete("1.0", "end")
        self._3d_tbl.insert("end", hdr + sep + body)
        self._3d_tbl.configure(state="disabled")

    def _browse_2d(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("All", "*.*")])
        if p:
            self._2d_file_var.set(p)

    def _browse_3d(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        if p:
            self._3d_file_var.set(p)

    # ════════════════════════════════════════════════════════════════════════
    # 2D ACTIONS
    # ════════════════════════════════════════════════════════════════════════
    def _run_2d_validate(self):
        self._log2d(f"\n[BOM] {self._2d_file_var.get()}")
        threading.Thread(target=self._t_2d_validate, daemon=True).start()

    def _t_2d_validate(self):
        from validator import EngineeringValidator
        v = EngineeringValidator(self._2d_file_var.get(),
                                 log_callback=lambda m: self._q2d.put(("log", m)))
        v.run_checks()
        self._2d_validator = v
        st = {p["Part_Number"]: "⏳ Queued" for p in v.valid_parts}
        self._q2d.put(("table", (v.valid_parts, st)))
        col = COL["success"] if v.error_count == 0 else COL["warn"]
        self._q2d.put(("status", (f"✔ {len(v.valid_parts)} valid | {v.error_count} errors", col)))

    def _run_2d_batch(self):
        self._btn_2d_run.configure(state="disabled")
        self._2d_prog.set(0)
        threading.Thread(target=self._t_2d_batch, daemon=True).start()

    def _t_2d_batch(self):
        from validator import EngineeringValidator
        from autocad_engine import AutoCADController
        self._q2d.put(("log", f"\n[BOM] {self._2d_file_var.get()}"))
        self._q2d.put(("progress", (0.05, "Validating…")))
        v = EngineeringValidator(self._2d_file_var.get(),
                                 log_callback=lambda m: self._q2d.put(("log", m)))
        v.run_checks()
        self._2d_validator = v
        parts = v.valid_parts
        if self._2d_type_var.get() != "All":
            parts = [p for p in parts if p.get("Part_Type") == self._2d_type_var.get()]
        if self._2d_prio_var.get() != "All":
            parts = [p for p in parts if p.get("Priority") == self._2d_prio_var.get()]
        if not parts:
            self._q2d.put(("status", ("⚠ No parts after filters", COL["warn"])))
            self._q2d.put(("btn", None)); return
        st = {p["Part_Number"]: "⏳ Queued" for p in parts}
        self._q2d.put(("table", (parts, dict(st))))
        self._q2d.put(("progress", (0.25, "Connecting to AutoCAD…")))
        try:
            eng = AutoCADController(log_callback=lambda m: self._q2d.put(("log", m)))
            idx = [0]
            orig = eng._log_info
            def tlog(msg):
                orig(msg)
                if "[*] GENERATING:" in msg and idx[0] < len(parts):
                    if idx[0] > 0:
                        st[parts[idx[0]-1]["Part_Number"]] = "✔ Done"
                    st[parts[idx[0]]["Part_Number"]] = "⚙ Drawing…"
                    self._q2d.put(("table", (parts, dict(st))))
                    self._q2d.put(("progress", (0.25 + 0.70*(idx[0]/len(parts)),
                        f"Drawing {parts[idx[0]]['Part_Number']} ({idx[0]+1}/{len(parts)})")))
                    idx[0] += 1
                elif "ERROR drafting" in msg:
                    for k, v2 in st.items():
                        if v2 == "⚙ Drawing…": st[k] = "✘ Error"
                    self._q2d.put(("table", (parts, dict(st))))
            eng._log_info = tlog
            eng.generate_batch(parts)
            for k in st:
                if st[k] in ("⏳ Queued","⚙ Drawing…"): st[k] = "✔ Done"
            self._q2d.put(("table", (parts, dict(st))))
            self._q2d.put(("progress", (1.0, "Complete!")))
            self._q2d.put(("status", (f"✔ {len(parts)} parts drawn", COL["success"])))
        except Exception as e:
            self._q2d.put(("log", f"[ERROR] {e}"))
            self._q2d.put(("status", ("● Engine error", COL["error"])))
        self._q2d.put(("btn", None))

    def _save_2d_report(self):
        if not self._2d_validator:
            messagebox.showinfo("Info", "Run validation first."); return
        path = filedialog.asksaveasfilename(defaultextension=".txt",
            filetypes=[("Text","*.txt")], initialfile="validation_2d_report.txt")
        if path:
            with open(path, "w") as f: f.write(self._2d_validator.summary_report())
            messagebox.showinfo("Saved", f"Saved:\n{path}")

    # ════════════════════════════════════════════════════════════════════════
    # 3D ACTIONS
    # ════════════════════════════════════════════════════════════════════════
    def _run_3d_validate(self):
        self._log3d(f"\n[BOM] {self._3d_file_var.get()}")
        threading.Thread(target=self._t_3d_validate, daemon=True).start()

    def _t_3d_validate(self):
        from validator_3d import Validator3D
        v = Validator3D(self._3d_file_var.get(),
                        log_callback=lambda m: self._q3d.put(("log", m)))
        v.run_checks()
        self._3d_validator = v
        st = {p["Part_Number"]: "⏳ Queued" for p in v.valid_parts}
        self._q3d.put(("table", (v.valid_parts, st)))
        col = COL["success"] if v.error_count == 0 else COL["warn"]
        self._q3d.put(("status", (f"✔ {len(v.valid_parts)} valid | {v.error_count} errors", col)))

    def _run_3d_batch(self):
        self._btn_3d_run.configure(state="disabled")
        self._3d_prog.set(0)
        threading.Thread(target=self._t_3d_batch, daemon=True).start()

    def _t_3d_batch(self):
        from validator_3d import Validator3D
        from autocad_engine_3d import AutoCAD3DGearEngine as AutoCAD3DController
        self._q3d.put(("log", f"\n[BOM] {self._3d_file_var.get()}"))
        self._q3d.put(("progress", (0.05, "Validating 3D BOM…")))
        v = Validator3D(self._3d_file_var.get(),
                        log_callback=lambda m: self._q3d.put(("log", m)))
        v.run_checks()
        self._3d_validator = v
        parts = v.valid_parts
        if self._3d_type_var.get() != "All":
            parts = [p for p in parts if p.get("Part_Type") == self._3d_type_var.get()]
        if self._3d_prio_var.get() != "All":
            parts = [p for p in parts if p.get("Priority") == self._3d_prio_var.get()]
        if not parts:
            self._q3d.put(("status", ("⚠ No parts after filters", COL["warn"])))
            self._q3d.put(("btn", None)); return
        st = {p["Part_Number"]: "⏳ Queued" for p in parts}
        self._q3d.put(("table", (parts, dict(st))))
        self._q3d.put(("progress", (0.20, "Booting 3D engine…")))
        try:
            eng = AutoCAD3DController(log_callback=lambda m: self._q3d.put(("log", m)))
            idx = [0]
            orig = eng._log
            def tlog(msg):
                orig(msg)
                if "3D GENERATING:" in msg and idx[0] < len(parts):
                    if idx[0] > 0:
                        st[parts[idx[0]-1]["Part_Number"]] = "✔ Done"
                    st[parts[idx[0]]["Part_Number"]] = "⚙ Building…"
                    self._q3d.put(("table", (parts, dict(st))))
                    self._q3d.put(("progress", (0.20 + 0.75*(idx[0]/len(parts)),
                        f"Building {parts[idx[0]]['Part_Number']} ({idx[0]+1}/{len(parts)})")))
                    idx[0] += 1
                elif "ERROR" in msg:
                    for k, v2 in st.items():
                        if v2 == "⚙ Building…": st[k] = "✘ Error"
                    self._q3d.put(("table", (parts, dict(st))))
            eng._log = tlog
            eng.generate_3d_batch(parts)
            for k in st:
                if st[k] in ("⏳ Queued","⚙ Building…"): st[k] = "✔ Done"
            self._q3d.put(("table", (parts, dict(st))))
            self._q3d.put(("progress", (1.0, "3D Build Complete!")))
            self._q3d.put(("status",
                (f"✔ {len(parts)} solids + layouts", COL["success"])))
        except Exception as e:
            import traceback
            self._q3d.put(("log", f"[ERROR] {e}\n{traceback.format_exc()}"))
            self._q3d.put(("status", ("● 3D Engine error", COL["error"])))
        self._q3d.put(("btn", None))

    def _save_3d_report(self):
        if not self._3d_validator:
            messagebox.showinfo("Info", "Run 3D validation first."); return
        path = filedialog.asksaveasfilename(defaultextension=".txt",
            filetypes=[("Text","*.txt")], initialfile="validation_3d_report.txt")
        if path:
            with open(path, "w") as f: f.write(self._3d_validator.summary_report())
            messagebox.showinfo("Saved", f"Saved:\n{path}")

    # ── Queue pollers ─────────────────────────────────────────────────────────
    def _poll_2d(self):
        try:
            while True:
                k, d = self._q2d.get_nowait()
                if k == "log":     self._log2d(str(d))
                elif k == "table": self._tbl2d(*d)
                elif k == "progress":
                    self._2d_prog.set(d[0])
                    self._2d_prog_lbl.configure(text=d[1])
                elif k == "status":
                    self._2d_status.configure(text=f"● {d[0]}", text_color=d[1])
                elif k == "btn":   self._btn_2d_run.configure(state="normal")
        except queue.Empty:
            pass
        self.after(120, self._poll_2d)

    def _poll_3d(self):
        try:
            while True:
                k, d = self._q3d.get_nowait()
                if k == "log":     self._log3d(str(d))
                elif k == "table": self._tbl3d(*d)
                elif k == "progress":
                    self._3d_prog.set(d[0])
                    self._3d_prog_lbl.configure(text=d[1])
                elif k == "status":
                    self._3d_status.configure(text=f"● {d[0]}", text_color=d[1])
                elif k == "btn":   self._btn_3d_run.configure(state="normal")
        except queue.Empty:
            pass
        self.after(130, self._poll_3d)


if __name__ == "__main__":
    SiraalGUI().mainloop()