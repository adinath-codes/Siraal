"""
gui_launcher.py — Siraal Grand Unified Manufacturing Engine
Advanced CustomTkinter GUI: live sync, part status table, progress bar,
type/priority filters, validation report viewer, theme switching
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
import time
import logging

from validator      import EngineeringValidator
from autocad_engine import AutoCADController

# ── Theme ─────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"
logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)


# ════════════════════════════════════════════════════════════════════════════
class DesignSyncApp(ctk.CTk):
# ════════════════════════════════════════════════════════════════════════════

    def __init__(self):
        super().__init__()
        self.title("DesignSync AI — Siraal Engine v3.0")
        self.geometry("1050x750")
        self.resizable(True, True)

        self.excel_path  = None
        self.last_mtime  = 0
        self.is_watching = False
        self._valid_parts = []

        self._build_ui()

    # ── UI Construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color="#1a1a2e", corner_radius=0)
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text="DesignSync AI",
                     font=ctk.CTkFont(size=30, weight="bold"),
                     text_color="#00d4ff").pack(side="left", padx=20, pady=14)
        ctk.CTkLabel(hdr, text="TN-IMPACT 2026  |  Grand Unified Manufacturing Engine v3.0",
                     text_color="#aaaaaa").pack(side="left")

        # Theme toggle
        self._theme_var = ctk.StringVar(value="Dark")
        ctk.CTkOptionMenu(hdr, values=["Dark", "Light", "System"],
                          variable=self._theme_var,
                          command=self._change_theme,
                          width=90).pack(side="right", padx=14)

        # ── Body (left panel + right panel) ───────────────────────────────────
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=10, pady=8)

        left = ctk.CTkFrame(body, width=340)
        left.pack(side="left", fill="y", padx=(0, 6))
        left.pack_propagate(False)

        right = ctk.CTkFrame(body)
        right.pack(side="left", fill="both", expand=True)

        self._build_left(left)
        self._build_right(right)

    def _build_left(self, parent):
        ctk.CTkLabel(parent, text="📁  BOM FILE",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=14, pady=(12, 2))

        ctk.CTkButton(parent, text="Select Excel / CSV BOM", height=38,
                      command=self._select_file).pack(padx=14, fill="x")

        self._lbl_file = ctk.CTkLabel(parent, text="No file selected",
                                       text_color="#ffcc00", wraplength=300)
        self._lbl_file.pack(padx=14, pady=(4, 10))

        # ── Filters ───────────────────────────────────────────────────────────
        ctk.CTkLabel(parent, text="⚙  FILTERS",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=14, pady=(8, 2))

        ctk.CTkLabel(parent, text="Part Type", text_color="#aaaaaa").pack(anchor="w", padx=14)
        self._type_filter = ctk.CTkOptionMenu(
            parent, values=["ALL", "Plate", "Spur_Gear", "Stepped_Shaft",
                             "Flanged_Shaft", "Ring_Gear"])
        self._type_filter.pack(padx=14, fill="x")

        ctk.CTkLabel(parent, text="Priority", text_color="#aaaaaa").pack(anchor="w", padx=14, pady=(6, 0))
        self._prio_filter = ctk.CTkOptionMenu(
            parent, values=["ALL", "High", "Medium", "Low"])
        self._prio_filter.pack(padx=14, fill="x")

        # ── Live Sync ─────────────────────────────────────────────────────────
        ctk.CTkLabel(parent, text="🔄  LIVE SYNC",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=14, pady=(14, 2))
        self._sync_sw = ctk.CTkSwitch(parent, text="Watch file for changes",
                                       command=self._toggle_sync)
        self._sync_sw.pack(padx=14, anchor="w")

        ctk.CTkLabel(parent, text="Poll interval (s)",
                     text_color="#aaaaaa").pack(anchor="w", padx=14, pady=(6, 0))
        self._poll_var = ctk.StringVar(value="3")
        ctk.CTkEntry(parent, textvariable=self._poll_var, width=80).pack(anchor="w", padx=14)

        # ── Actions ───────────────────────────────────────────────────────────
        ctk.CTkLabel(parent, text="▶  ACTIONS",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=14, pady=(16, 2))

        self._btn_validate = ctk.CTkButton(
            parent, text="1 — Validate BOM Only", height=36,
            fg_color="#6c5ce7", hover_color="#4b3fa8",
            command=self._run_validate_only)
        self._btn_validate.pack(padx=14, fill="x", pady=2)

        self._btn_run = ctk.CTkButton(
            parent, text="2 — RUN FULL CAD BATCH", height=42,
            fg_color="#00b894", hover_color="#00856e",
            font=ctk.CTkFont(weight="bold"),
            command=self._run_batch_thread)
        self._btn_run.pack(padx=14, fill="x", pady=4)

        self._btn_report = ctk.CTkButton(
            parent, text="3 — Save Validation Report", height=34,
            fg_color="#636e72", hover_color="#2d3436",
            command=self._save_report)
        self._btn_report.pack(padx=14, fill="x", pady=2)

        self._btn_clear = ctk.CTkButton(
            parent, text="Clear Log", height=30, fg_color="#b2bec3",
            hover_color="#dfe6e9", text_color="#2d3436",
            command=self._clear_log)
        self._btn_clear.pack(padx=14, fill="x", pady=(12, 2))

        # ── Status indicator ──────────────────────────────────────────────────
        self._status_lbl = ctk.CTkLabel(parent, text="● Idle",
                                          text_color="#aaaaaa")
        self._status_lbl.pack(padx=14, pady=10, anchor="w")

    def _build_right(self, parent):
        # ── Progress bar ──────────────────────────────────────────────────────
        self._progress = ctk.CTkProgressBar(parent, width=600)
        self._progress.pack(fill="x", padx=10, pady=(8, 2))
        self._progress.set(0)
        self._progress_lbl = ctk.CTkLabel(parent, text="Ready", text_color="#aaaaaa")
        self._progress_lbl.pack(anchor="w", padx=10)

        # ── Part status table ─────────────────────────────────────────────────
        ctk.CTkLabel(parent, text="📋  PART QUEUE", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(8, 2))

        tbl_frame = ctk.CTkFrame(parent, height=180)
        tbl_frame.pack(fill="x", padx=10, pady=(0, 6))
        tbl_frame.pack_propagate(False)

        self._tbl = ctk.CTkTextbox(tbl_frame, font=("Consolas", 10), height=180)
        self._tbl.pack(fill="both", expand=True)
        self._tbl.configure(state="disabled")
        
        # Trigger the empty state immediately on boot
        self._update_tbl([]) 

        # ── Log terminal ──────────────────────────────────────────────────────
        ctk.CTkLabel(parent, text="🖥  ENGINE LOG", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(6, 2))

        self._log_box = ctk.CTkTextbox(parent, font=("Consolas", 11))
        self._log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self._log_box.configure(state="disabled")
        self._log("[SYSTEM] Siraal Engine v3.0 — standby.")
    # ── Helpers ───────────────────────────────────────────────────────────────
    def _log(self, msg: str):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("0.0", "end")
        self._log_box.configure(state="disabled")

    def _status(self, text: str, color: str = "#aaaaaa"):
        self._status_lbl.configure(text=text, text_color=color)

    def _set_progress(self, value: float, label: str = ""):
        self._progress.set(value)
        if label:
            self._progress_lbl.configure(text=label)

    def _change_theme(self, mode: str):
        ctk.set_appearance_mode(mode)

    def _write_tbl_header(self):
        self._tbl.configure(state="normal")
        self._tbl.delete("0.0", "end")
        hdr = f"{'#':<3} {'Part No':<16} {'Type':<16} {'Material':<14} {'P1':>7} {'P2':>6} {'Priority':<10} {'Status'}"
        sep = "─" * 90
        self._tbl.insert("end", hdr + "\n" + sep + "\n")
        self._tbl.configure(state="disabled")

    def _update_tbl(self, parts: list, statuses: dict = None):
        statuses = statuses or {}
        self._tbl.configure(state="normal")
        self._write_tbl_header()
        
        # --- NEW EMPTY STATE LOGIC ---
        if not parts:
            empty_msg = "\n\n" + " "*24 + "📁 [ AWAITING BOM DATA / NO PARTS QUEUED ]\n"
            empty_msg += " "*28 + "Please select a valid Excel or CSV file."
            self._tbl.insert("end", empty_msg)
        # -----------------------------
        else:
            for i, p in enumerate(parts, 1):
                st = statuses.get(p["Part_Number"], "⏳ Queued")
                row = (f"{i:<3} {p['Part_Number']:<16} {p['Part_Type']:<16} "
                       f"{p['Material']:<14} {p['Param_1']:>7.1f} {p['Param_2']:>6.1f} "
                       f"{p.get('Priority','Med'):<10} {st}")
                self._tbl.insert("end", row + "\n")
                
        self._tbl.configure(state="disabled")
    # ── File selection ────────────────────────────────────────────────────────
    def _select_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Data Files", "*.xlsx *.csv"), ("All", "*.*")])
        if path:
            self.excel_path = path
            self._lbl_file.configure(text=os.path.basename(path))
            self._log(f"[*] BOM loaded: {path}")
            if self.is_watching:
                self.last_mtime = os.path.getmtime(path)

    # ── Validate-only run ─────────────────────────────────────────────────────
    def _run_validate_only(self):
        if not self.excel_path:
            messagebox.showwarning("No File", "Select a BOM file first.")
            return
        self._btn_validate.configure(state="disabled")
        threading.Thread(target=self._validate_task, daemon=True).start()

    def _validate_task(self):
        self._status("● Validating…", "#f9ca24")
        self._set_progress(0.2, "Validating BOM…")
        v = EngineeringValidator(self.excel_path, log_callback=self._log)
        v.run_checks()
        self._valid_parts = self._apply_filters(v.valid_parts)
        self._update_tbl(self._valid_parts)
        self._set_progress(1.0, f"Validation done — {len(self._valid_parts)} valid parts")
        self._status(f"● {len(self._valid_parts)} valid parts", "#00cec9")
        self._btn_validate.configure(state="normal")

# ── Full batch run ────────────────────────────────────────────────────────
    def _run_batch_thread(self):
        if not self.excel_path:
            messagebox.showwarning("No File", "Select a BOM file first.")
            return
        self._btn_run.configure(state="disabled")
        threading.Thread(target=self._batch_task, daemon=True).start()

    def _get_next_session_name(self) -> str:
        """Determines the next available session folder name."""
        base_dir = os.path.join(os.getcwd(), "CNC_Machine_Files")
        os.makedirs(base_dir, exist_ok=True)
        session_num = 1
        while os.path.exists(os.path.join(base_dir, f"Session_{session_num}")):
            session_num += 1
        return f"Session_{session_num}"

    def _batch_task(self):
        self._status("● Running pipeline…", "#fdcb6e")
        self._set_progress(0.0, "Starting…")

        # Phase 1 — validate
        self._set_progress(0.1, "Phase 1: Validation…")
        v = EngineeringValidator(self.excel_path, log_callback=self._log)
        ok = v.run_checks()
        if not ok:
            self._log("[-] No valid parts — aborting.")
            self._status("● Failed", "#d63031")
            self._btn_run.configure(state="normal")
            return

        parts = self._apply_filters(v.valid_parts)
        if not parts:
            self._log("[-] All parts filtered out — change filter settings.")
            self._btn_run.configure(state="normal")
            return

        self._valid_parts = parts
        self._update_tbl(parts)

        # Generate a new session folder name for this run
        current_session = self._get_next_session_name()
        self._log(f"[*] Starting new run: {current_session}")

        # Phase 2 — AutoCAD
        self._set_progress(0.3, "Phase 2: Connecting to AutoCAD…")
        
        # --- NEW CALLBACK TO UPDATE GUI FROM THE ENGINE ---
        statuses = {}
        def _update_ui_status(pno, st, prog_ratio):
            statuses[pno] = st
            self._update_tbl(parts, statuses)
            if prog_ratio is not None:
                self._set_progress(0.3 + 0.5 * prog_ratio, f"Generating {pno}…")
        # --------------------------------------------------

        try:
            engine = AutoCADController(log_callback=self._log, session_name=current_session)
            
            # Pass the ENTIRE list at once, and let the engine handle the loop!
            engine.generate_batch(parts, status_callback=_update_ui_status)
            
            # Export DXF files at the end
            self._set_progress(0.85, "Exporting isolated DXF files…")
            engine._export_all_dxf(parts)

        except Exception as e:
            self._log(f"[-] AutoCAD engine failure: {e}")
            self._status("● AutoCAD error", "#d63031")
            self._btn_run.configure(state="normal")
            return

        self._set_progress(1.0, f"Complete — {len(parts)} parts generated in {current_session}")
        self._status("● Pipeline complete ✔", "#00b894")
        self._btn_run.configure(state="normal")
    # ── Save report ───────────────────────────────────────────────────────────
    def _save_report(self):
        if not self.excel_path:
            messagebox.showwarning("No File", "Select a BOM file first.")
            return
        v = EngineeringValidator(self.excel_path, log_callback=self._log)
        v.run_checks()
        path = os.path.join(os.path.dirname(self.excel_path), "validation_report.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write(v.summary_report())
        self._log(f"[+] Report saved: {path}")
        messagebox.showinfo("Saved", f"Report saved to:\n{path}")

    # ── Filters ───────────────────────────────────────────────────────────────
    def _apply_filters(self, parts: list) -> list:
        t = self._type_filter.get()
        p = self._prio_filter.get()
        out = parts
        if t != "ALL":
            out = [x for x in out if x["Part_Type"] == t]
        if p != "ALL":
            out = [x for x in out if x.get("Priority", "Medium") == p]
        return out

    # ── Live Sync watcher ─────────────────────────────────────────────────────
    def _toggle_sync(self):
        if self._sync_sw.get() == 1:
            if not self.excel_path:
                self._log("[-] Select a BOM file before enabling Live Sync.")
                self._sync_sw.deselect()
                return
            self.is_watching = True
            self.last_mtime  = os.path.getmtime(self.excel_path)
            self._log("[*] Live Sync enabled — watching for file changes…")
            self._watch_loop()
        else:
            self.is_watching = False
            self._log("[*] Live Sync disabled.")

    def _watch_loop(self):
        if not self.is_watching:
            return
        try:
            mtime = os.path.getmtime(self.excel_path)
            if mtime > self.last_mtime:
                self.last_mtime = mtime
                self._log("[!] File change detected!")
                if messagebox.askyesno("Live Sync",
                                       f"{os.path.basename(self.excel_path)} changed.\n"
                                       "Regenerate AutoCAD batch now?"):
                    self._run_batch_thread()
        except Exception:
            pass

        try:
            interval_ms = max(1000, int(float(self._poll_var.get()) * 1000))
        except ValueError:
            interval_ms = 3000

        self.after(interval_ms, self._watch_loop)


if __name__ == "__main__":
    app = DesignSyncApp()
    app.mainloop()
