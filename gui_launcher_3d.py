"""
gui_launcher_3d.py  —  Siraal Manufacturing Engine  v6.0
======================================================
Professional 4-Tab GUI (Integrated with robust Validator3D & Cost Engine):
  Tab 1 — EXCEL MODE      : Load BOM xlsx → Validate → Run 3D batch → Cost PDF
  Tab 2 — MANUAL MODE     : Type gear specs → Auto-builds Excel → Runs batch → Cost PDF
  Tab 3 — AI SHAPE CREATOR: Text-to-CAD → Gemini API → CSG JSON Template → 3D Preview
  Tab 4 — AI BOM COPILOT  : Generative AI for editing mass Excel BOMs
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading, os, queue, json, math, time, datetime, traceback
from pathlib import Path
import pythoncom

# ── Optional imports (non-fatal) ─────────────────────────────────────────────
try:    import pandas as pd;         PANDAS_OK = True
except: PANDAS_OK = False
try:    import openpyxl;             OPENPYXL_OK = True
except: OPENPYXL_OK = False
try:
    from watchdog.observers import Observer
    from watchdog.events   import FileSystemEventHandler
    WATCHDOG_OK = True
except: WATCHDOG_OK = False

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# DESIGN SYSTEM
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

C = {
    "void":      "#080C10", "base":      "#0D1117", "surface":   "#111820",
    "elevated":  "#161E28", "card":      "#1C2534", "border":    "#243040",
    "border2":   "#2E3E52", "gold":      "#F0B429", "gold_dim":  "#A07820",
    "amber":     "#E8821A", "teal":      "#1ABC9C", "teal_dim":  "#0E8870",
    "violet":    "#8B5CF6", "violet_dim":"#5B3FA6", "ok":        "#22C55E",
    "warn":      "#F59E0B", "error":     "#EF4444", "info":      "#3B82F6",
    "text":      "#E8EDF2", "text2":     "#8B99AA", "text3":     "#5A6A7A",
}

FONT_HEADER  = ("Segoe UI",  20, "bold")
FONT_TITLE   = ("Segoe UI",  13, "bold")
FONT_BODY    = ("Segoe UI",  10)
FONT_SMALL   = ("Segoe UI",   9)
FONT_MONO    = ("Cascadia Code", 8)
FONT_MONO_M  = ("Cascadia Code", 9)

def _dim(hex_color: str, alpha_hex: str, bg: str = "#0D1117") -> str:
    a  = int(alpha_hex, 16) / 255.0
    fg = [int(hex_color.lstrip("#")[i*2:i*2+2], 16) for i in range(3)]
    bv = [int(bg.lstrip("#")[i*2:i*2+2],        16) for i in range(3)]
    b  = [int(fg[i]*a + bv[i]*(1-a))             for i in range(3)]
    return "#{:02X}{:02X}{:02X}".format(*b)

def _badge_bg(hex_color: str) -> str:
    return _dim(hex_color, "1A")

GEAR_TYPES   = ["Spur_Gear_3D","Helical_Gear","Ring_Gear_3D",
                "Bevel_Gear","Worm","Worm_Wheel"]
MATERIALS    = ["Steel-4140","Steel-1020","Al-6061","Brass-C360","Nylon-66","Ti-6Al-4V"]
PRIORITIES   = ["High","Medium","Low"]

def _profile_shift(Z, PA=20.0):
    a = math.radians(PA)
    z_min = 2.0 / math.sin(a)**2
    return round(max(0.0, (z_min-Z)/z_min), 4) if Z < z_min else 0.0

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# REALTIME EXCEL WATCHER
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class ExcelWatcher:
    def __init__(self, on_change):
        self._on_change = on_change
        self._observer  = None
        self._path      = None
        self._last_snap = {}

    def _snapshot(self):
        if not (self._path and PANDAS_OK): return {}
        try:
            xl = pd.read_excel(self._path, sheet_name=None, dtype=str)
            snap = {}
            for sname, df in xl.items():
                snap[sname] = df.fillna("").to_dict(orient="records")
            return snap
        except Exception: return {}

    def start(self, path):
        self.stop()
        self._path = path
        self._last_snap = self._snapshot()
        if not WATCHDOG_OK: return

        class _H(FileSystemEventHandler):
            def __init__(self_, ev_cb): self_.ev_cb = ev_cb
            def on_modified(self_, ev):
                if Path(ev.src_path).resolve() == Path(path).resolve(): self_.ev_cb()

        self._observer = Observer()
        self._observer.schedule(_H(self._check), str(Path(path).parent), recursive=False)
        self._observer.start()

    def _check(self):
        time.sleep(0.4)
        new_snap = self._snapshot()
        diffs = []
        for sheet in set(list(self._last_snap.keys()) + list(new_snap.keys())):
            old_rows = self._last_snap.get(sheet, [])
            new_rows = new_snap.get(sheet, [])
            if old_rows != new_rows:
                diffs.append(f"Sheet [{sheet}]: {abs(len(new_rows)-len(old_rows))} row delta")
                for i, (o, n) in enumerate(zip(old_rows, new_rows)):
                    for k in set(list(o.keys())+list(n.keys())):
                        ov, nv = o.get(k,""), n.get(k,"")
                        if ov != nv: diffs.append(f"  Row {i+1}  [{k}]  {ov!r} → {nv!r}")
        if diffs:
            self._last_snap = new_snap
            self._on_change(diffs)

    def stop(self):
        if self._observer:
            try: self._observer.stop(); self._observer.join(timeout=1)
            except Exception: pass
            self._observer = None

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# EXCEL WRITER (Legacy for Tabs 1-3. Tab 4 uses Copilot Backend)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def write_bom_excel(parts, out_path):
    if not OPENPYXL_OK: raise ImportError("openpyxl not installed")
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils  import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "BOM_Gears"

    FG_HEADER = "0D3A5C"; FG_SUB = "1A2B3C"; FG_ALT = "1C2D3E"
    FG_WARN = "3A2000"; GOLD = "F0B429"; WHITE = "E8EDF2"; GREY = "8B99AA"
    thin = Side(style="thin", color="2E3E52")
    bd   = Border(left=thin, right=thin, top=thin, bottom=thin)
    def hfill(hex_): return PatternFill("solid", fgColor=hex_)

    ws.merge_cells("A1:P1")
    c = ws["A1"]
    c.value = "SIRAAL 3D GEAR ENGINE — GEAR BILL OF MATERIALS"
    c.font  = Font(name="Segoe UI", bold=True, color=GOLD, size=12)
    c.fill  = hfill(FG_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:P2")
    c2 = ws["A2"]
    c2.value = ("IS 2535 / ISO 1328  |  Involute 20°  |  1st Angle  |  TN-IMPACT 2026  "
                "|  P1=Teeth Z (N_starts for Worm)  P2=Module m(mm)  "
                "P3=Face Width(mm)  P4=Bore Dia(mm) or Ring Thk")
    c2.font  = Font(name="Segoe UI", color=GREY, size=8)
    c2.fill  = hfill(FG_SUB)
    c2.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 15

    COLS = ["#","Part_Number","Part_Type","Material",
            "Param_1\n(Z / Starts)","Param_2\n(Module m)",
            "Param_3\n(Face Width)","Param_4\n(Bore Dia)",
            "Qty","Priority","Enabled","Description",
            "Mass_kg","Est_Cost\n(Rs)","Notes","Profile_Shift_x"]
    WIDTHS = [4,20,16,12,10,10,10,10,5,8,7,30,8,12,30,10]

    for ci, (col, w) in enumerate(zip(COLS, WIDTHS), 1):
        cell = ws.cell(row=3, column=ci, value=col)
        cell.font  = Font(name="Segoe UI", bold=True, color=WHITE, size=9)
        cell.fill  = hfill("1A3A5C")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = bd
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[3].height = 30

    for ri, p in enumerate(parts, 4):
        ptype = p.get("Part_Type","Spur_Gear_3D")
        Z     = int(float(p.get("Param_1",20)))
        m     = float(p.get("Param_2",3))
        fw    = float(p.get("Param_3",30))
        bd_d  = float(p.get("Param_4",20))
        mat   = p.get("Material","Steel-4140")
        x     = _profile_shift(Z) if ptype in ("Spur_Gear_3D","Helical_Gear","Worm_Wheel") else 0.0

        import math as _m
        r_a = Z*m/2+m*(1+x); r_b = bd_d/2
        vol = _m.pi*(r_a**2-max(r_b,0)**2)*fw * 0.92
        DENS = {"Steel-1020":7.87e-3,"Steel-4140":7.85e-3,"Al-6061":2.70e-3,
                "Brass-C360":8.50e-3,"Nylon-66":1.14e-3,"Ti-6Al-4V":4.43e-3}
        COST = {"Steel-1020":125,"Steel-4140":185,"Al-6061":265,
                "Brass-C360":520,"Nylon-66":415,"Ti-6Al-4V":3800}
        mass = round(vol * DENS.get(mat,7.85e-3)/1e6, 3)
        cost = round(mass * COST.get(mat,185), 2)

        row_fill = hfill(FG_WARN if x > 0 else (FG_ALT if ri%2==0 else FG_SUB))
        vals = [
            ri-3, p.get("Part_Number",""), ptype, mat,
            Z, m, fw, bd_d,
            p.get("Quantity",1), p.get("Priority","High"), "YES",
            p.get("Description",""), mass, cost, p.get("Notes",""), x
        ]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.fill = row_fill
            cell.font = Font(name="Segoe UI", color=WHITE, size=9)
            cell.alignment = Alignment(vertical="center", horizontal="center" if ci in (1,5,6,7,8,9,10,11,16) else "left")
            cell.border = bd
        ws.row_dimensions[ri].height = 16

    n = len(parts)
    tr = 4 + n
    ws.merge_cells(f"A{tr}:D{tr}")
    c = ws.cell(row=tr, column=1, value="TOTALS (all enabled):")
    c.font = Font(name="Segoe UI", bold=True, color=GOLD, size=9)
    c.fill = hfill(FG_HEADER)
    c.alignment = Alignment(horizontal="right")
    for ci in range(1,17): ws.cell(row=tr,column=ci).fill = hfill(FG_HEADER)

    ws.freeze_panes = "B4"
    wb.save(out_path)
    return out_path

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SHARED WIDGETS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _divider(parent, label="", color=C["border2"]):
    fr = ctk.CTkFrame(parent, fg_color="transparent", height=22)
    fr.pack(fill="x", padx=12, pady=(6,0))
    fr.pack_propagate(False)
    ctk.CTkFrame(fr, fg_color=color, height=1, corner_radius=0).place(relx=0, rely=0.5, relwidth=1.0, anchor="w")
    if label:
        lbl = ctk.CTkLabel(fr, text=f"  {label}  ", font=ctk.CTkFont("Segoe UI", 8, "bold"), text_color=C["text3"], fg_color=C["surface"])
        lbl.place(relx=0.03, rely=0.5, anchor="w")

def _log_widget(parent):
    fr = ctk.CTkFrame(parent, fg_color=C["void"], corner_radius=8)
    txt = tk.Text(fr, bg=C["void"], fg="#7DFF9A", font=FONT_MONO_M, relief="flat", padx=10, pady=8, state="disabled", cursor="arrow", wrap="word", selectbackground=C["border"])
    sb = ctk.CTkScrollbar(fr, command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    txt.pack(fill="both", expand=True)
    txt.tag_config("ok",    foreground="#22C55E")
    txt.tag_config("warn",  foreground="#F59E0B")
    txt.tag_config("err",   foreground="#EF4444")
    txt.tag_config("info",  foreground="#60A5FA")
    txt.tag_config("head",  foreground=C["gold"])
    txt.tag_config("ai",    foreground="#C084FC")
    txt.tag_config("cost",  foreground="#A78BFA") # Violet for cost engine
    return fr, txt

def _append_log(txt_widget, msg, tag=""):
    txt_widget.configure(state="normal")
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}\n"
    if not tag:
        if any(x in msg for x in ("✔","OK","Done","complete","COMPLETE")): tag="ok"
        elif any(x in msg for x in ("⚠","WARN","warn","WARNING")):         tag="warn"
        elif any(x in msg for x in ("✘","ERROR","error","fail","FAIL")):    tag="err"
        elif any(x in msg for x in ("SYSTEM","BOM","LOADING","SAVING")):    tag="info"
        elif any(x in msg for x in ("AI","Gemini","Designing","model")):    tag="ai"
        elif any(x in msg for x in ("PDF","Economic","ESG","Cost")):        tag="cost"
        elif msg.startswith("╔") or msg.startswith("║") or msg.startswith("╚"): tag="head"
    txt_widget.insert("end", line, tag or "")
    txt_widget.see("end")
    txt_widget.configure(state="disabled")

def _pill_button(parent, text, cmd, color, width=None, height=34):
    kw = dict(height=height, fg_color=_dim(color,"33"), hover_color=_dim(color,"55"),
              border_color=color, border_width=1, text_color=color, font=ctk.CTkFont("Segoe UI", 10, "bold"),
              corner_radius=8, command=cmd)
    if width: kw["width"] = width
    return ctk.CTkButton(parent, text=text, **kw)

def _run_button(parent, text, cmd, color, height=40):
    return ctk.CTkButton(parent, text=text, height=height, fg_color=color, hover_color=color,
                         text_color=C["void"], font=ctk.CTkFont("Segoe UI", 11, "bold"), corner_radius=8, command=cmd)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# GEAR ROW WIDGET  (used in Manual Mode)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class GearRow(ctk.CTkFrame):
    _counter = 0
    def __init__(self, parent, on_delete, on_change, main_app): 
        super().__init__(parent, fg_color=C["card"], corner_radius=8)
        GearRow._counter += 1
        self.main_app = main_app 
        self._on_delete = on_delete; self._on_change = on_change
        self.pack(fill="x", pady=(0, 4))
        self._build()

    def _field(self, parent, label, var, width=72, choices=None):
        fr = ctk.CTkFrame(parent, fg_color="transparent")
        fr.pack(side="left", padx=(0,6))
        ctk.CTkLabel(fr, text=label, font=ctk.CTkFont("Segoe UI",7), text_color=C["text3"]).pack(anchor="w")
        
        if label == "Type":
            w = ctk.CTkComboBox(fr, variable=var, 
                                values=self.main_app._get_all_part_types(), 
                                width=width, height=26,
                                fg_color=C["elevated"], border_color=C["border2"],
                                button_color=C["border2"], text_color=C["text"],
                                font=ctk.CTkFont("Segoe UI",9), 
                                command=lambda _: self._on_change())
            w.bind("<Button-1>", lambda e: w.configure(values=self.main_app._get_all_part_types()))
            w.pack()
        elif choices:
            w = ctk.CTkComboBox(fr, variable=var, values=choices, width=width, height=26,
                                fg_color=C["elevated"], border_color=C["border2"],
                                button_color=C["border2"], text_color=C["text"],
                                font=ctk.CTkFont("Segoe UI",9), command=lambda _: self._on_change())
            w.pack()
        else:
            w = ctk.CTkEntry(fr, textvariable=var, width=width, height=26,
                             fg_color=C["elevated"], border_color=C["border2"],
                             text_color=C["text"], font=ctk.CTkFont("Segoe UI",9))
            w.pack()
            var.trace_add("write", lambda *_: self._on_change())
        return w

    def _build(self):
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=8, pady=(6,2))
        n = GearRow._counter

        self.v_pno  = ctk.StringVar(value=f"GR-{n:03d}")
        self.v_type = ctk.StringVar(value="Spur_Gear_3D")
        self.v_mat  = ctk.StringVar(value="Steel-4140")
        self.v_z    = ctk.StringVar(value="20"); self.v_m    = ctk.StringVar(value="3")
        self.v_fw   = ctk.StringVar(value="30"); self.v_bd   = ctk.StringVar(value="20")
        self.v_qty  = ctk.StringVar(value="1"); self.v_prio = ctk.StringVar(value="High")
        self.v_desc = ctk.StringVar(value="")

        self._field(top,"Part No",  self.v_pno,  92)
        self._field(top,"Type",     self.v_type, 130, None) 
        self._field(top,"Material", self.v_mat,  110, MATERIALS)
        self._field(top,"Z / Starts",self.v_z,   58)
        self._field(top,"Module m", self.v_m,    52)
        self._field(top,"Face W",   self.v_fw,   52)
        self._field(top,"Bore Ø",   self.v_bd,   52)
        self._field(top,"Qty",      self.v_qty,  40)
        self._field(top,"Priority", self.v_prio, 80, PRIORITIES)

        tail = ctk.CTkFrame(top, fg_color="transparent")
        tail.pack(side="left", padx=(6,0))
        ctk.CTkButton(tail, text="✕", width=26, height=26, fg_color="#3A1B20", hover_color="#672529", text_color=C["error"], border_color=C["error"], border_width=1, corner_radius=6, font=ctk.CTkFont("Segoe UI",9,"bold"), command=self._on_delete).pack()

        desc_row = ctk.CTkFrame(self, fg_color="transparent")
        desc_row.pack(fill="x", padx=8, pady=(0,6))
        ctk.CTkLabel(desc_row, text="Description:", font=ctk.CTkFont("Segoe UI",7), text_color=C["text3"]).pack(side="left", padx=(0,4))
        ctk.CTkEntry(desc_row, textvariable=self.v_desc, height=22, fg_color=C["elevated"], border_color=C["border2"], text_color=C["text2"], font=ctk.CTkFont("Segoe UI",8)).pack(side="left", fill="x", expand=True)

        self._warn_lbl = ctk.CTkLabel(desc_row, text="", font=ctk.CTkFont("Segoe UI",7), text_color=C["warn"])
        self._warn_lbl.pack(side="right", padx=(4,0))

        for v in (self.v_z, self.v_m, self.v_bd, self.v_fw): v.trace_add("write", lambda *_: self._live_validate())

    def _live_validate(self):
        try:
            Z=int(self.v_z.get()); m=float(self.v_m.get()); bd=float(self.v_bd.get()); fw=float(self.v_fw.get())
            warns=[]
            x=_profile_shift(Z)
            if x>0: warns.append(f"⚠ Profile shift x={x:.3f}")
            if Z<6: warns.append("⚠ Z<6 critical")
            if bd/2>=Z*m/2: warns.append("⚠ bore≥PCD/2")
            if fw/m>12: warns.append("⚠ fw/m>12")
            if fw/m<6:  warns.append("⚠ fw/m<6")
            self._warn_lbl.configure(text="  ".join(warns))
        except Exception: self._warn_lbl.configure(text="")

    def get_part(self):
        return {"Part_Number": self.v_pno.get(), "Part_Type": self.v_type.get(), "Material": self.v_mat.get(), "Param_1": self.v_z.get(), "Param_2": self.v_m.get(), "Param_3": self.v_fw.get(), "Param_4": self.v_bd.get(), "Qty": self.v_qty.get(), "Priority": self.v_prio.get(), "Description": self.v_desc.get(), "Enabled": "YES"}

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# MAIN WINDOW
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class SiraalGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("SIRAAL Manufacturing Engine  |  TN-IMPACT 2026")
        self.geometry("1500x920")
        self.minsize(1200, 800)
        self.configure(fg_color=C["base"])

        self._q1 = queue.Queue(); self._q2 = queue.Queue(); self._q3 = queue.Queue()
        self._q4 = queue.Queue() # Queue for Tab 4 AI Copilot
        self._watcher  = ExcelWatcher(self._on_excel_change)
        self._watch_active = False
        self._gear_rows: list[GearRow] = []
        self._pending_copilot_data = None # Store AI generated edits

        self._build_header()
        self._build_tabs()
        self._build_excel_tab()
        self._build_manual_tab()
        self._build_ai_tab()
        self._build_copilot_tab() # Tab 4
        self._build_watcher_bar()

        self.after(100, self._poll_q1)
        self.after(110, self._poll_q2)
        self.after(120, self._poll_q3)
        self.after(130, self._poll_q4)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # Tab listener
        self._tabs.configure(command=self._on_tab_change)

    # ─────────────────────────────────────────────────────────────────────────
    # MISSING UI HELPER METHODS ADDED HERE
    # ─────────────────────────────────────────────────────────────────────────
    def _update_tbl(self, widget, parts, states):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("end", f"{'PART NUMBER':<24} {'TYPE':<18} {'STATUS'}\n", "head")
        widget.insert("end", "-"*60 + "\n", "head")
        for p in parts:
            pno = p.get("Part_Number", "")
            pt = p.get("Part_Type", "")
            st = states.get(pno, "⏳ Queued")
            line = f"{pno:<24} {pt:<18} {st}\n"
            if "✔" in st: widget.insert("end", line, "ok")
            elif "✘" in st: widget.insert("end", line, "err")
            elif "⚙" in st: widget.insert("end", line, "warn")
            else: widget.insert("end", line)
        widget.configure(state="disabled")

    def _clear_log(self, widget):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.configure(state="disabled")

    def _on_tab_change(self):
        if "MANUAL MODE" in self._tabs.get():
            self._refresh_custom_buttons()

    def _on_close(self):
        self._watcher.stop()
        self.destroy()

    def _filter_manual_rows(self, *args):
        query = self._manual_search_var.get().lower()
        for row in self._gear_rows:
            if query in row.v_pno.get().lower() or query in row.v_type.get().lower():
                row.pack(fill="x", pady=(0, 4))
            else:
                row.pack_forget()

    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color=C["void"], corner_radius=0, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        logo_fr = ctk.CTkFrame(hdr, fg_color="transparent")
        logo_fr.pack(side="left", padx=(18,0), pady=8)
        ctk.CTkLabel(logo_fr, text="⚙", font=ctk.CTkFont("Segoe UI", 22), text_color=C["gold"]).pack(side="left", padx=(0,6))
        ctk.CTkLabel(logo_fr, text="SIRAAL  MANUFACTURING  ENGINE", font=ctk.CTkFont("Segoe UI", 15, "bold"), text_color=C["text"]).pack(side="left")
        ctk.CTkLabel(logo_fr, text=" v6.0", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"]).pack(side="left", pady=(4,0))
        meta = ctk.CTkFrame(hdr, fg_color="transparent")
        meta.pack(side="right", padx=18)
        
        # --- NEW GLOBAL COST REPORT BUTTON ---
        self._btn_global_cost = ctk.CTkButton(meta, text="📊 ESG & COST REPORT", fg_color=C["violet"], hover_color="#6D48C4", text_color=C["void"], font=ctk.CTkFont("Segoe UI", 10, "bold"), height=26, command=self._run_global_cost)
        self._btn_global_cost.pack(side="left", padx=(0, 15))
        
        for badge, col in [("IS 2535/ISO 1328", C["teal"]), ("1st Angle", C["gold"]), ("TN-IMPACT 2026", C["amber"])]:
            ctk.CTkLabel(meta, text=badge, font=ctk.CTkFont("Segoe UI",8,"bold"), text_color=col, fg_color=_badge_bg(col), corner_radius=4, padx=6, pady=2).pack(side="left", padx=3)
        ctk.CTkFrame(self, fg_color=C["border"], height=1, corner_radius=0).pack(fill="x")

    def _build_tabs(self):
        self._tabs = ctk.CTkTabview(
            self, fg_color=C["base"], segmented_button_fg_color=C["void"],
            segmented_button_selected_color=C["elevated"], segmented_button_selected_hover_color=C["card"],
            segmented_button_unselected_color=C["void"], segmented_button_unselected_hover_color=C["surface"],
            text_color=C["text2"], text_color_disabled=C["text3"], corner_radius=0, border_width=0,
        )
        self._tabs.pack(fill="both", expand=True, padx=0, pady=0)
        self._T1 = "  📂  EXCEL MODE  "
        self._T2 = "  🔧  MANUAL MODE  "
        self._T3 = "  🤖  AI SHAPE CREATOR  "
        self._T4 = "  🧠  AI BOM COPILOT  " # Tab 4
        for t in (self._T1, self._T2, self._T3, self._T4): self._tabs.add(t)

    def _get_all_part_types(self):
        native = ["Spur_Gear_3D", "Helical_Gear", "Ring_Gear_3D", "Bevel_Gear", "Worm", "Worm_Wheel"]
        customs = []
        base_dir = os.path.dirname(os.path.abspath(__file__))
        t_path = os.path.join(base_dir, "templates")
        if os.path.exists(t_path):
            for f in os.listdir(t_path):
                if f.startswith("Custom_") and f.endswith(".json"):
                    customs.append(f.replace(".json", ""))
        return native + sorted(customs)

    def _get_custom_template_names(self):
        customs = []
        base_dir = os.path.dirname(os.path.abspath(__file__))
        t_path = os.path.join(base_dir, "templates")
        if os.path.exists(t_path):
            for f in os.listdir(t_path):
                if f.startswith("Custom_") and f.endswith(".json"):
                    customs.append(f.replace(".json", ""))
        return sorted(customs)

    def _refresh_custom_buttons(self):
        for child in self._custom_btn_fr.winfo_children():
            child.destroy()
            
        templates = self._get_custom_template_names()
        if not templates:
            ctk.CTkLabel(self._custom_btn_fr, text="No AI templates. Use Tab 3 to create one.", 
                         font=ctk.CTkFont("Segoe UI", 9, "italic"), text_color=C["text3"]).pack(pady=10, padx=10)
            return

        for name in templates:
            display_name = name.replace("Custom_", "")
            btn = _pill_button(self._custom_btn_fr, f"🤖 {display_name}", 
                               lambda n=name: self._add_row(n), C["violet"], height=28)
            btn.pack(side="left", padx=4)

    def _build_watcher_bar(self):
        bar = ctk.CTkFrame(self, fg_color=C["void"], corner_radius=0, height=38)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        ctk.CTkFrame(self, fg_color=C["border"], height=1, corner_radius=0).pack(fill="x", side="bottom")
        ctk.CTkLabel(bar, text="⚡  REALTIME EXCEL CHECKER", font=ctk.CTkFont("Segoe UI",9,"bold"), text_color=C["gold"]).pack(side="left", padx=(12,8), pady=8)
        self._watch_path_var = ctk.StringVar(value="")
        self._watch_entry = ctk.CTkEntry(bar, textvariable=self._watch_path_var, width=320, height=26, fg_color=C["surface"], border_color=C["border2"], text_color=C["text2"], font=ctk.CTkFont("Segoe UI",9), placeholder_text="Select Excel file to watch…")
        self._watch_entry.pack(side="left", padx=(0,6))
        ctk.CTkButton(bar, text="Browse", width=64, height=26, fg_color=C["elevated"], hover_color=C["card"], text_color=C["text2"], font=ctk.CTkFont("Segoe UI",8), command=self._browse_watch).pack(side="left", padx=(0,6))
        self._watch_toggle = ctk.CTkButton(bar, text="▶  START WATCH", width=120, height=26, fg_color="#0F3331", hover_color="#114A43", border_color=C["teal"], border_width=1, text_color=C["teal"], font=ctk.CTkFont("Segoe UI",9,"bold"), command=self._toggle_watch)
        self._watch_toggle.pack(side="left", padx=(0,12))
        self._watch_indicator = ctk.CTkLabel(bar, text="● IDLE", font=ctk.CTkFont("Segoe UI",9,"bold"), text_color=C["text3"])
        self._watch_indicator.pack(side="left")
        self._watch_diff_fr = ctk.CTkFrame(bar, fg_color=C["surface"], corner_radius=6, height=26)
        self._watch_diff_fr.pack(side="left", fill="x", expand=True, padx=(12,12), pady=6)
        self._watch_diff_lbl = ctk.CTkLabel(self._watch_diff_fr, text="No changes detected", font=ctk.CTkFont("Cascadia Code", 8), text_color=C["text3"], anchor="w")
        self._watch_diff_lbl.pack(fill="x", padx=8)

    def _browse_watch(self):
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls"),("All","*.*")])
        if p: self._watch_path_var.set(p)

    def _toggle_watch(self):
        if self._watch_active:
            self._watcher.stop()
            self._watch_active = False
            self._watch_toggle.configure(text="▶  START WATCH", fg_color="#0F3331", border_color=C["teal"], text_color=C["teal"])
            self._watch_indicator.configure(text="● IDLE", text_color=C["text3"])
        else:
            p = self._watch_path_var.get()
            if not p or not os.path.exists(p):
                messagebox.showwarning("Watch", "Please select a valid Excel file first.")
                return
            self._watcher.start(p)
            self._watch_active = True
            self._watch_toggle.configure(text="⏹  STOP WATCH", fg_color="#3A1B20", border_color=C["error"], text_color=C["error"])
            self._watch_indicator.configure(text=f"● WATCHING  {os.path.basename(p)}", text_color=C["ok"])

    def _on_excel_change(self, diffs):
        preview = " │ ".join(diffs[:3])
        if len(diffs) > 3: preview += f"  +{len(diffs)-3} more"
        
        # We must route the UI updates back to the main thread via self.after
        def _update_ui():
            self._watch_diff_lbl.configure(text=f"⚡ {datetime.datetime.now().strftime('%H:%M:%S')}  {preview}", text_color=C["gold"])
            for log in (self._log1_txt, self._log2_txt, self._log3_txt, self._log4_txt):
                _append_log(log, f"⚡ Excel changed: {preview}", "warn")
            
            # Show the alert pop-up directly on the main thread
            messagebox.showinfo(
                "Watched BOM Modified", 
                "The watched Excel file has been modified externally.\n\nPlease Re-Run the validation and generation engine to apply these changes."
            )
            
        self.after(0, _update_ui)

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 1 — EXCEL MODE
    # ═════════════════════════════════════════════════════════════════════════
    def _build_excel_tab(self):
        tab = self._tabs.tab(self._T1)
        tab.configure(fg_color=C["base"])
        root = ctk.CTkFrame(tab, fg_color=C["base"])
        root.pack(fill="both", expand=True, padx=12, pady=10)
        root.columnconfigure(0, weight=0, minsize=270); root.columnconfigure(1, weight=1); root.rowconfigure(0, weight=1)
        lp = ctk.CTkFrame(root, fg_color=C["surface"], corner_radius=10, width=270)
        lp.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        lp.pack_propagate(False)
        self._excel_left(lp)
        rp = ctk.CTkFrame(root, fg_color=C["surface"], corner_radius=10)
        rp.grid(row=0, column=1, sticky="nsew")
        rp.rowconfigure(1, weight=1); rp.columnconfigure(0, weight=1)
        self._excel_right(rp)

    def _excel_left(self, p):
        ctk.CTkFrame(p, fg_color=C["teal"], height=3, corner_radius=2).pack(fill="x", padx=0, pady=0)
        ctk.CTkLabel(p, text="📂  EXCEL MODE", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color=C["teal"]).pack(pady=(14,2), padx=12)
        ctk.CTkLabel(p, text="Load a BOM xlsx · Validate · Generate 3D", font=ctk.CTkFont("Segoe UI",9), text_color=C["text3"]).pack(pady=(0,10))

        _divider(p, "BOM FILE", C["teal"])
        self._e1_file = ctk.StringVar(value="excels/demo_gears_3d.xlsx")
        ctk.CTkEntry(p, textvariable=self._e1_file, fg_color=C["card"], border_color=C["border2"], text_color=C["text"], height=28, font=ctk.CTkFont("Segoe UI",9)).pack(fill="x", padx=12, pady=(4,3))
        row = ctk.CTkFrame(p, fg_color="transparent")
        row.pack(fill="x", padx=12, pady=(0,2))
        _pill_button(row,"Browse", self._browse_e1, C["teal"], height=28).pack(side="left", padx=(0,4))
        
        _divider(p, "FILTERS", C["teal"])
        ctk.CTkLabel(p, text="Part Type", font=ctk.CTkFont("Segoe UI",8), text_color=C["text3"]).pack(anchor="w", padx=14, pady=(4,0))
        self._e1_type = ctk.StringVar(value="All")
        ctk.CTkComboBox(p, variable=self._e1_type, values=["All"]+GEAR_TYPES, fg_color=C["card"], button_color=C["border2"], border_color=C["border2"], text_color=C["text"], font=ctk.CTkFont("Segoe UI",9), height=28).pack(fill="x", padx=12, pady=(2,4))
        ctk.CTkLabel(p, text="Priority", font=ctk.CTkFont("Segoe UI",8), text_color=C["text3"]).pack(anchor="w", padx=14)
        self._e1_prio = ctk.StringVar(value="All")
        ctk.CTkComboBox(p, variable=self._e1_prio, values=["All"]+PRIORITIES, fg_color=C["card"], button_color=C["border2"], border_color=C["border2"], text_color=C["text"], font=ctk.CTkFont("Segoe UI",9), height=28).pack(fill="x", padx=12, pady=(2,10))

        _divider(p, "ACTIONS", C["teal"])
        _pill_button(p, "① Validate BOM", self._validate_e1, C["teal"], height=32).pack(fill="x", padx=12, pady=(6,4))
        
        self._btn_e1 = _run_button(p, "② GENERATE 3D PARTS", self._run_e1, C["teal"])
        self._btn_e1.pack(fill="x", padx=12, pady=(0,4))

        self._e1_status = ctk.CTkLabel(p, text="● Ready", font=ctk.CTkFont("Segoe UI",10,"bold"), text_color=C["ok"])
        self._e1_status.pack(pady=4)
        self._e1_prog = ctk.CTkProgressBar(p, fg_color=C["card"], progress_color=C["teal"], height=6, corner_radius=3)
        self._e1_prog.set(0)
        self._e1_prog.pack(fill="x", padx=12, pady=(0,2))
        self._e1_prog_lbl = ctk.CTkLabel(p, text="", font=ctk.CTkFont("Segoe UI",8), text_color=C["text3"])
        self._e1_prog_lbl.pack(pady=(0,6))
        
        info = ctk.CTkFrame(p, fg_color=C["card"], corner_radius=6)
        info.pack(fill="x", padx=12, pady=(0,10))
        self._e1_stats = ctk.CTkLabel(info, text="No BOM loaded", font=ctk.CTkFont("Cascadia Code",8), text_color=C["text3"], justify="left")
        self._e1_stats.pack(padx=8, pady=6, anchor="w")

    def _excel_right(self, p):
        tbl_fr = ctk.CTkFrame(p, fg_color=C["card"], corner_radius=8)
        tbl_fr.grid(row=0, column=0, sticky="ew", padx=10, pady=(10,4))
        ctk.CTkLabel(tbl_fr, text="  PART LIST", font=ctk.CTkFont("Segoe UI",9,"bold"), text_color=C["teal"]).pack(anchor="w", padx=4, pady=(4,0))
        self._e1_tbl = tk.Text(tbl_fr, height=8, bg=C["card"], fg=C["text"], font=FONT_MONO, relief="flat", padx=8, pady=4, state="disabled", cursor="arrow")
        sb = ctk.CTkScrollbar(tbl_fr, command=self._e1_tbl.yview)
        self._e1_tbl.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._e1_tbl.pack(fill="both", expand=True)

        log_lbl = ctk.CTkFrame(p, fg_color="transparent")
        log_lbl.grid(row=1, column=0, sticky="ew", padx=10, pady=(4,0))
        ctk.CTkLabel(log_lbl, text="  ⚙ ENGINE LOG", font=ctk.CTkFont("Segoe UI",9,"bold"), text_color=C["teal"]).pack(side="left")
        _pill_button(log_lbl, "Clear", lambda: self._clear_log(self._log1_txt), C["text3"], height=22, width=50).pack(side="right", padx=6)
        log_fr, self._log1_txt = _log_widget(p)
        log_fr.grid(row=2, column=0, sticky="nsew", padx=10, pady=(2,10))
        p.rowconfigure(2, weight=1)
        _append_log(self._log1_txt, "SYSTEM  Siraal Excel Mode — standby.", "info")

    def _browse_e1(self):
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls"),("All","*.*")])
        if p: self._e1_file.set(p)

    def _validate_e1(self):
        threading.Thread(target=self._t_validate_e1, daemon=True).start()

    def _t_validate_e1(self):
        path = self._e1_file.get()
        self._q1.put(("log", f"\n[BOM] VALIDATING {path}"))
        if not os.path.exists(path):
            self._q1.put(("log", f"✘ File not found: {path}")); return
            
        try:
            from validator_3d import Validator3D 
            v = Validator3D(path, log_callback=lambda m: self._q1.put(("log", m)))
            v.run_checks()
            parts = v.valid_parts
        except Exception as e:
            self._q1.put(("log", f"✘ Validation error: {e}")); return
            
        pt = self._e1_type.get(); pr = self._e1_prio.get()
        if pt != "All": parts = [p for p in parts if p.get("Part_Type")==pt]
        if pr != "All": parts = [p for p in parts if p.get("Priority")==pr]
        
        self._q1.put(("stats", f"Parts: {len(parts)}  (filtered)\nFile: {os.path.basename(path)}"))
        st = {p["Part_Number"]: "⏳ Queued" for p in parts}
        self._q1.put(("table1", (parts, dict(st))))
        col = C["ok"] if v.error_count == 0 else C["warn"]
        self._q1.put(("status1", (f"✔ {len(parts)} valid | {v.error_count} errors", col)))

    def _run_e1(self):
        self._btn_e1.configure(state="disabled")
        threading.Thread(target=self._t_run_e1, daemon=True).start()

    def _t_run_e1(self):
        path = self._e1_file.get()
        self._q1.put(("prog1",(0.05,"Validating…")))
        try:
            from validator_3d import Validator3D
            from autocad_engine_3d import AutoCAD3DGearEngine
            
            v = Validator3D(path, log_callback=lambda m: self._q1.put(("log", m)))
            v.run_checks()
            parts = v.valid_parts
            
            pt = self._e1_type.get(); pr = self._e1_prio.get()
            if pt != "All": parts = [p for p in parts if p.get("Part_Type")==pt]
            if pr != "All": parts = [p for p in parts if p.get("Priority")==pr]
            
            if not parts:
                self._q1.put(("status1", ("⚠ No parts after filters", C["warn"])))
                self._q1.put(("btn1", None)); return
                
            st = {p["Part_Number"]: "⏳ Queued" for p in parts}
            self._q1.put(("table1", (parts, dict(st))))
            self._q1.put(("prog1",(0.15,"Starting AutoCAD…")))
            
            eng = AutoCAD3DGearEngine(log_cb=lambda m: self._q1.put(("log",m)))
            idx = [0]
            orig = eng._log
            
            def tlog(msg):
                orig(msg)
                if "3D GENERATING:" in msg and idx[0] < len(parts):
                    if idx[0] > 0:
                        st[parts[idx[0]-1]["Part_Number"]] = "✔ Done"
                    st[parts[idx[0]]["Part_Number"]] = "⚙ Building…"
                    self._q1.put(("table1", (parts, dict(st))))
                    self._q1.put(("prog1", (0.20 + 0.75*(idx[0]/len(parts)), f"Building {parts[idx[0]]['Part_Number']} ({idx[0]+1}/{len(parts)})")))
                    idx[0] += 1
                elif "ERROR" in msg or "fail" in msg.lower():
                    for k, v2 in st.items():
                        if v2 == "⚙ Building…": st[k] = "✘ Error"
                    self._q1.put(("table1", (parts, dict(st))))
                    
            eng._log = tlog
            self._q1.put(("prog1",(0.25,"Generating parts…")))
            eng.generate_3d_batch(parts)
            
            for k in st:
                if st[k] in ("⏳ Queued","⚙ Building…"): st[k] = "✔ Done"
            self._q1.put(("table1", (parts, dict(st))))
            
            self._q1.put(("prog1",(1.0,"Complete!")))
            self._q1.put(("status1",(f"✔ {len(parts)} parts built", C["ok"])))
        except Exception as e:
            self._q1.put(("log", f"✘ {e}\n{traceback.format_exc()}"))
            self._q1.put(("status1",("● Error", C["error"])))
        self._q1.put(("btn1", None))

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 2 — MANUAL MODE
    # ═════════════════════════════════════════════════════════════════════════
    def _build_manual_tab(self):
        tab = self._tabs.tab(self._T2)
        tab.configure(fg_color=C["base"])
        root = ctk.CTkFrame(tab, fg_color=C["base"])
        root.pack(fill="both", expand=True, padx=12, pady=10)
        root.columnconfigure(0, weight=0, minsize=260); root.columnconfigure(1, weight=1); root.rowconfigure(0, weight=1)
        lp = ctk.CTkFrame(root, fg_color=C["surface"], corner_radius=10, width=260)
        lp.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        lp.pack_propagate(False)
        self._manual_left(lp)
        rp = ctk.CTkFrame(root, fg_color=C["surface"], corner_radius=10)
        rp.grid(row=0, column=1, sticky="nsew")
        rp.rowconfigure(1, weight=1); rp.columnconfigure(0, weight=1)
        self._manual_right(rp)

    def _manual_left(self, p):
        ctk.CTkFrame(p, fg_color=C["gold"], height=3, corner_radius=2).pack(fill="x")
        ctk.CTkLabel(p, text="🔧  MANUAL MODE", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color=C["gold"]).pack(pady=(14,2), padx=12)
        ctk.CTkLabel(p, text="Build gear specs manually → Excel → 3D", font=ctk.CTkFont("Segoe UI",9), text_color=C["text3"]).pack(pady=(0,10))

        _divider(p, "OUTPUT", C["gold"])
        ctk.CTkLabel(p, text="Save Excel to:", font=ctk.CTkFont("Segoe UI",8), text_color=C["text3"]).pack(anchor="w", padx=14, pady=(4,0))
        self._m_out = ctk.StringVar(value="output_manual/manual_bom.xlsx")
        ctk.CTkEntry(p, textvariable=self._m_out, fg_color=C["card"], border_color=C["border2"], text_color=C["text"], height=28, font=ctk.CTkFont("Segoe UI",9)).pack(fill="x", padx=12, pady=(2,4))
        _pill_button(p,"Choose…", self._browse_m_out, C["text3"], height=26).pack(fill="x", padx=12, pady=(0,8))

        _divider(p, "LIVE STATS", C["gold"])
        self._m_stats_lbl = ctk.CTkLabel(p, text="0 parts  |  0 warnings", font=ctk.CTkFont("Cascadia Code",9), text_color=C["text2"])
        self._m_stats_lbl.pack(pady=(6,2), padx=12, anchor="w")
        self._m_warn_fr = ctk.CTkScrollableFrame(p, fg_color=C["card"], corner_radius=6, height=80)
        self._m_warn_fr.pack(fill="x", padx=12, pady=(0,10))
        self._m_warn_labels = []

        _divider(p, "ACTIONS", C["gold"])
        self._btn_m_save = _pill_button(p, "① Save Excel", self._save_manual_excel, C["gold"], height=32)
        self._btn_m_save.pack(fill="x", padx=12, pady=(0,4))
        
        self._btn_m_run = _run_button(p, "② VALIDATE & GENERATE 3D", self._run_manual, C["gold"])
        self._btn_m_run.pack(fill="x", padx=12, pady=(0,4))

        self._m_status = ctk.CTkLabel(p, text="● Ready", font=ctk.CTkFont("Segoe UI",10,"bold"), text_color=C["ok"])
        self._m_status.pack(pady=4)
        self._m_prog = ctk.CTkProgressBar(p, fg_color=C["card"], progress_color=C["gold"], height=6, corner_radius=3)
        self._m_prog.set(0)
        self._m_prog.pack(fill="x", padx=12, pady=(0,2))
        self._m_prog_lbl = ctk.CTkLabel(p, text="", font=ctk.CTkFont("Segoe UI",8), text_color=C["text3"])
        self._m_prog_lbl.pack(pady=(0,8))

    def _manual_right(self, p):
        search_fr = ctk.CTkFrame(p, fg_color="transparent")
        search_fr.grid(row=0, column=0, sticky="ew", padx=10, pady=(10,0))
        self._manual_search_var = ctk.StringVar()
        self._manual_search_var.trace_add("write", self._filter_manual_rows)
        ctk.CTkEntry(search_fr, textvariable=self._manual_search_var, 
                    placeholder_text="🔍 Filter rows by Part No or Type...", 
                    width=400, height=28).pack(side="left", fill="x", expand=True)

        tb = ctk.CTkFrame(p, fg_color="transparent")
        tb.grid(row=1, column=0, sticky="ew", padx=10, pady=(10,2))
        ctk.CTkLabel(tb, text=" STANDARD GEARS  ", font=ctk.CTkFont("Segoe UI",10,"bold"), text_color=C["gold"]).pack(side="left")
        for gtype, col in [("Spur_Gear_3D", C["gold"]), ("Helical_Gear", C["amber"]), ("Ring_Gear_3D", C["teal"]), ("Bevel_Gear", C["info"])]:
            _pill_button(tb, f"＋ {gtype.split('_')[0]}", lambda t=gtype: self._add_row(t), col, height=26).pack(side="left", padx=2)
        
        custom_lbl_fr = ctk.CTkFrame(p, fg_color="transparent", height=20)
        custom_lbl_fr.grid(row=2, column=0, sticky="ew", padx=12, pady=(10,0))
        ctk.CTkFrame(custom_lbl_fr, fg_color=C["violet"], height=1).place(relx=0, rely=0.5, relwidth=1.0)
        ctk.CTkLabel(custom_lbl_fr, text="  CUSTOM BUILDERS (AI)  ", font=ctk.CTkFont("Segoe UI", 8, "bold"), 
                     text_color=C["text3"], fg_color=C["surface"]).place(relx=0.03, rely=0.5, anchor="w")

        self._custom_btn_fr = ctk.CTkScrollableFrame(p, fg_color="transparent", orientation="horizontal", height=50)
        self._custom_btn_fr.grid(row=3, column=0, sticky="ew", padx=10, pady=(2,5))
        
        self._rows_outer = ctk.CTkScrollableFrame(p, fg_color=C["card"], corner_radius=8)
        self._rows_outer.grid(row=4, column=0, sticky="nsew", padx=10, pady=(5,4))
        p.rowconfigure(4, weight=3) 
        
        log_fr, self._log2_txt = _log_widget(p)
        log_fr.grid(row=5, column=0, sticky="nsew", padx=10, pady=(2,10))
        p.rowconfigure(5, weight=1)
        
        self._refresh_custom_buttons()
        self._add_row("Spur_Gear_3D")

    def _add_row(self, gear_type="Spur_Gear_3D"):
        def _del():
            self._gear_rows.remove(row); row.destroy(); self._update_manual_stats()
        
        row = GearRow(self._rows_outer, _del, self._update_manual_stats, self) 
        row.v_type.set(gear_type)
        self._gear_rows.append(row)
        self._update_manual_stats()

    def _add_worm_pair(self):
        self._add_row("Worm"); self._add_row("Worm_Wheel")
        n = len(self._gear_rows)
        base = f"GR-WP-{n:03d}"
        self._gear_rows[-2].v_pno.set(f"{base}-WORM")
        self._gear_rows[-1].v_pno.set(f"{base}-WHEEL")
        self._gear_rows[-1].v_mat.set("Brass-C360")

    def _clear_rows(self):
        for r in self._gear_rows: r.destroy()
        self._gear_rows.clear()
        self._update_manual_stats()

    def _update_manual_stats(self):
        total = len(self._gear_rows)
        warns = []
        for r in self._gear_rows:
            w = r._warn_lbl.cget("text")
            if w: warns.append(f"{r.v_pno.get()}: {w}")
        self._m_stats_lbl.configure(text=f"{total} parts  |  {len(warns)} warnings")
        for lbl in self._m_warn_labels:
            try: lbl.destroy()
            except Exception: pass
        self._m_warn_labels.clear()
        for w in warns[:8]:
            lbl = ctk.CTkLabel(self._m_warn_fr, text=w, font=ctk.CTkFont("Segoe UI",8), text_color=C["warn"], anchor="w")
            lbl.pack(anchor="w", padx=4, pady=1)
            self._m_warn_labels.append(lbl)

    def _get_manual_parts(self):
        return [r.get_part() for r in self._gear_rows]

    def _browse_m_out(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")], initialfile="manual_bom.xlsx")
        if p: self._m_out.set(p)

    def _save_manual_excel(self):
        parts = self._get_manual_parts()
        if not parts:
            messagebox.showwarning("Save","No parts defined."); return
        out = self._m_out.get()
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        try:
            write_bom_excel(parts, out)
            _append_log(self._log2_txt, f"✔ Excel saved: {out}", "ok")
            messagebox.showinfo("Saved", f"Excel saved:\n{out}")
        except Exception as e:
            _append_log(self._log2_txt, f"✘ Save failed: {e}", "err")

    def _run_manual(self):
        self._btn_m_run.configure(state="disabled")
        threading.Thread(target=self._t_run_manual, daemon=True).start()

    def _t_run_manual(self):
        parts = self._get_manual_parts()
        if not parts:
            self._q2.put(("log","✘ No parts defined")); self._q2.put(("btn2",None)); return
        out = self._m_out.get()
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        self._q2.put(("prog2",(0.1,"Saving Excel…")))
        
        try:
            write_bom_excel(parts, out)
            self._q2.put(("log",f"✔ Excel: {out}"))
        except Exception as e:
            self._q2.put(("log",f"✘ Excel write: {e}")); self._q2.put(("btn2",None)); return
            
        self._q2.put(("prog2",(0.2,"Validating…")))
        try:
            from validator_3d import Validator3D
            from autocad_engine_3d import AutoCAD3DGearEngine
            
            v = Validator3D(out, log_callback=lambda m: self._q2.put(("log", m)))
            v.run_checks()
            valid_parts = v.valid_parts
            
            if not valid_parts:
                self._q2.put(("log","✘ Validation failed for all parts. Check warnings.")); self._q2.put(("btn2",None)); return
                
            self._q2.put(("prog2",(0.3,"Starting AutoCAD…")))
            eng = AutoCAD3DGearEngine(log_cb=lambda m: self._q2.put(("log",m)))
            
            self._q2.put(("prog2",(0.4,"Generating 3D parts…")))
            eng.generate_3d_batch(valid_parts)
            
            self._q2.put(("prog2",(1.0,"Complete!")))
            self._q2.put(("status2",(f"✔ {len(valid_parts)} parts built", C["ok"])))
        except Exception as e:
            self._q2.put(("log",f"✘ {e}\n{traceback.format_exc()}"))
            self._q2.put(("status2",("● Error", C["error"])))
        self._q2.put(("btn2",None))

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 3 — AI SHAPE CREATOR
    # ═════════════════════════════════════════════════════════════════════════
    def _build_ai_tab(self):
        tab = self._tabs.tab(self._T3)
        tab.configure(fg_color=C["base"])
        root = ctk.CTkFrame(tab, fg_color=C["base"])
        root.pack(fill="both", expand=True, padx=12, pady=10)
        root.columnconfigure(0, weight=0, minsize=270); root.columnconfigure(1, weight=1); root.rowconfigure(0, weight=1)
        lp = ctk.CTkFrame(root, fg_color=C["surface"], corner_radius=10, width=270)
        lp.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        lp.pack_propagate(False)
        self._ai_left(lp)
        rp = ctk.CTkFrame(root, fg_color=C["surface"], corner_radius=10)
        rp.grid(row=0, column=1, sticky="nsew")
        rp.rowconfigure(1, weight=2); rp.rowconfigure(3, weight=1); rp.columnconfigure(0, weight=1)
        self._ai_right(rp)

    def _ai_left(self, p):
        ctk.CTkFrame(p, fg_color=C["violet"], height=3, corner_radius=2).pack(fill="x")
        ctk.CTkLabel(p, text="🤖  AI SHAPE CREATOR", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color=C["violet"]).pack(pady=(14,2), padx=12)
        ctk.CTkLabel(p, text="Text-to-CAD: Generate CSG JSON templates automatically using Google Gemini.", font=ctk.CTkFont("Segoe UI",9), text_color=C["text3"], wraplength=220).pack(pady=(0,8))

        _divider(p, "PART CONFIG", C["violet"])
        ctk.CTkLabel(p, text="Part Name (Must be Custom_...):", font=ctk.CTkFont("Segoe UI",8), text_color=C["text3"]).pack(anchor="w", padx=14, pady=(4,0))
        self._ai_part_name = ctk.StringVar(value="Sensor_Mount")
        ctk.CTkEntry(p, textvariable=self._ai_part_name, fg_color=C["card"], border_color=C["border2"], text_color=C["text"], height=28, font=ctk.CTkFont("Segoe UI",9)).pack(fill="x", padx=12, pady=(4,14))
        
        _divider(p, "ACTIONS", C["violet"])
        self._btn_ai_gen = _run_button(p, "① GENERATE JSON TEMPLATE", self._run_ai_generate, C["violet"])
        self._btn_ai_gen.pack(fill="x", padx=12, pady=(6,8))
        self._btn_ai_run = _run_button(p, "② VALIDATE & BUILD 3D", self._run_ai_3d, C["violet"])
        self._btn_ai_run.pack(fill="x", padx=12, pady=(0,8))

        self._ai_status = ctk.CTkLabel(p, text="● Ready", font=ctk.CTkFont("Segoe UI",10,"bold"), text_color=C["ok"])
        self._ai_status.pack(pady=4)
        self._ai_prog = ctk.CTkProgressBar(p, fg_color=C["card"], progress_color=C["violet"], height=6, corner_radius=3)
        self._ai_prog.set(0)
        self._ai_prog.pack(fill="x", padx=12, pady=(0,2))
        self._ai_prog_lbl = ctk.CTkLabel(p, text="", font=ctk.CTkFont("Segoe UI",8), text_color=C["text3"])
        self._ai_prog_lbl.pack(pady=(0,8))

    def _ai_right(self, p):
        ctk.CTkLabel(p, text="  💬 DESCRIBE YOUR CUSTOM SHAPE", font=ctk.CTkFont("Segoe UI",10,"bold"), text_color=C["violet"]).grid(row=0,column=0,sticky="w",padx=10,pady=(10,2))
        prompt_fr = ctk.CTkFrame(p, fg_color=C["card"], corner_radius=8)
        prompt_fr.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,6))
        self._ai_prompt = tk.Text(prompt_fr, bg=C["card"], fg=C["text"], font=("Segoe UI",11), relief="flat", padx=12, pady=10, wrap="word", insertbackground=C["violet"], selectbackground="#2E2552")
        self._ai_prompt.pack(fill="both", expand=True)
        self._ai_prompt.insert("end", "Example: A rectangular base where length is P1, width is P2, and thickness is P3. Subtract a large cylinder hole in the exact center with diameter P4.")

        ctk.CTkLabel(p, text="  🤖 AI CSG RECIPE COMPILER LOG", font=ctk.CTkFont("Segoe UI",9,"bold"), text_color=C["violet"]).grid(row=2,column=0,sticky="w",padx=10,pady=(4,2))
        log_fr, self._log3_txt = _log_widget(p)
        log_fr.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0,10))
        _append_log(self._log3_txt, "SYSTEM  AI Mode — Describe the shape math using P1, P2, P3, and P4.", "info")

    def _run_ai_generate(self):
        self._btn_ai_gen.configure(state="disabled")
        threading.Thread(target=self._t_ai_generate, daemon=True).start()

    def _t_ai_generate(self):
        description = self._ai_prompt.get("1.0","end").strip()
        part_name = self._ai_part_name.get().strip().replace(" ", "_")
        
        # Pull API key securely from environment variables instead of GUI
        key = os.environ.get("GEMINI_API_KEY", "").strip()
        model_val = "gemini-2.5-flash"  
        
        if not description: self._q3.put(("log","✘ Enter a description first.")); self._q3.put(("btn_ai_gen",None)); return
        if not part_name: self._q3.put(("log","✘ Enter a Part Name.")); self._q3.put(("btn_ai_gen",None)); return
        
        if not key: 
            self._q3.put(("log","✘ System Error: GEMINI_API_KEY environment variable is missing.")); 
            self._q3.put(("btn_ai_gen",None)); 
            return

        self._q3.put(("prog3",(0.3,"Sending instructions to Gemini…")))
        
        def gui_logger(msg):
            self._q3.put(("log", msg))

        try:
            # IMPORT THE MODULAR AI GENERATOR
            from genai_creator import generate_siraal_shape
            
            success = generate_siraal_shape(
                part_name=part_name, 
                description=description, 
                api_key=key, 
                model_name=model_val, 
                log_cb=gui_logger
            )
            
            if success:
                self._q3.put(("status3",(f"✔ Template Ready", C["ok"])))
                self._q3.put(("prog3",(1.0,"JSON Template ready!")))
            else:
                self._q3.put(("prog3",(0,"Generation Failed")))
                
        except ImportError:
            self._q3.put(("log", "✘ Error: genai_creator.py not found in the directory."))
            self._q3.put(("prog3",(0,"Import Error")))
        except Exception as e:
            self._q3.put(("log", f"✘ System error: {e}"))
            self._q3.put(("prog3",(0,"Error")))

        self._q3.put(("btn_ai_gen",None))

    def _run_ai_3d(self):
        part_name = self._ai_part_name.get().strip().replace(" ", "_")
        if part_name.startswith("Custom_"): part_name = part_name[7:]
        target_json = os.path.join("templates", f"Custom_{part_name}.json")
        
        if not os.path.exists(target_json):
            messagebox.showwarning("AI Mode", f"Run Step ① first to generate {target_json}")
            return
            
        self._btn_ai_run.configure(state="disabled")
        threading.Thread(target=self._t_ai_3d, args=(part_name,), daemon=True).start()

    def _t_ai_3d(self, part_name):
        pythoncom.CoInitialize()
        self._q3.put(("prog3", (0.1, "Creating preview BOM…")))
        
        preview_parts = [{
            "Part_Number": f"AI-PREV-{part_name.upper()[:10]}",
            "Part_Type": f"Custom_{part_name}",
            "Material": "Al-6061",
            "Param_1": "100",  # Default P1
            "Param_2": "100",  # Default P2
            "Param_3": "20",   # Default P3
            "Param_4": "15",   # Default P4
            "Quantity": 1,
            "Priority": "High",
            "Description": "AI Generated CSG Preview",
            "Enabled": "YES"
        }]
        
        out_path = os.path.abspath("output_ai/ai_preview_bom.xlsx")
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        
        try:
            write_bom_excel(preview_parts, out_path)
            self._q3.put(("log", f"SYSTEM  Auto-generated preview BOM for Custom_{part_name}"))
        except Exception as e:
            self._q3.put(("log", f"✘ Failed to create preview BOM: {e}"))
            self._q3.put(("btn_ai_run", None))
            return
            
        self._q3.put(("prog3", (0.3, "Starting AutoCAD Engine…")))
        try:
            from validator_3d import Validator3D
            from autocad_engine_3d import AutoCAD3DGearEngine
            
            v = Validator3D(out_path, log_callback=lambda m: self._q3.put(("log", m)))
            v.run_checks()
            valid_parts = v.valid_parts
            
            if not valid_parts:
                self._q3.put(("log", "✘ Validation failed for preview."))
                self._q3.put(("btn_ai_run", None))
                return
                
            eng = AutoCAD3DGearEngine(log_cb=lambda m: self._q3.put(("log", m)))
            self._q3.put(("prog3", (0.6, "Building 3D Preview…")))
            eng.generate_3d_batch(valid_parts)
            
            self._q3.put(("prog3", (1.0, "Preview Complete!")))
            self._q3.put(("status3", (f"✔ 3D Preview Built", C["ok"])))
            
        except Exception as e:
            import traceback
            self._q3.put(("log", f"✘ {e}\n{traceback.format_exc()}"))
            self._q3.put(("status3", ("● Error", C["error"])))
            
        self._q3.put(("btn_ai_run", None))

    # ═════════════════════════════════════════════════════════════════════════
    # TAB 4 — AI BOM COPILOT (Generative Excel Editing)
    # ═════════════════════════════════════════════════════════════════════════
    def _build_copilot_tab(self):
        tab = self._tabs.tab(self._T4)
        tab.configure(fg_color=C["base"])
        root = ctk.CTkFrame(tab, fg_color=C["base"])
        root.pack(fill="both", expand=True, padx=12, pady=10)
        root.columnconfigure(0, weight=0, minsize=290); root.columnconfigure(1, weight=1); root.rowconfigure(0, weight=1)
        
        lp = ctk.CTkFrame(root, fg_color=C["surface"], corner_radius=10, width=290)
        lp.grid(row=0, column=0, sticky="nsew", padx=(0,8))
        lp.pack_propagate(False)
        self._copilot_left(lp)
        
        rp = ctk.CTkFrame(root, fg_color=C["surface"], corner_radius=10)
        rp.grid(row=0, column=1, sticky="nsew")
        rp.rowconfigure(0, weight=2); rp.rowconfigure(2, weight=1); rp.columnconfigure(0, weight=1)
        self._copilot_right(rp)

    def _copilot_left(self, p):
        ctk.CTkFrame(p, fg_color="#3B82F6", height=3, corner_radius=2).pack(fill="x") 
        ctk.CTkLabel(p, text="🧠  AI BOM COPILOT", font=ctk.CTkFont("Segoe UI",13,"bold"), text_color="#3B82F6").pack(pady=(14,2), padx=12)
        ctk.CTkLabel(p, text="Intelligently edit mass Excel BOMs using Natural Language commands.", font=ctk.CTkFont("Segoe UI",9), text_color=C["text3"], wraplength=260).pack(pady=(0,10))

        _divider(p, "1. TARGET BOM", "#3B82F6")
        self._cp_file = ctk.StringVar(value="")
        ctk.CTkEntry(p, textvariable=self._cp_file, fg_color=C["card"], border_color=C["border2"], text_color=C["text"], height=28, font=ctk.CTkFont("Segoe UI",9), placeholder_text="Select existing BOM...").pack(fill="x", padx=12, pady=(4,3))
        
        row = ctk.CTkFrame(p, fg_color="transparent")
        row.pack(fill="x", padx=12, pady=(0,8))
        _pill_button(row,"Browse", self._browse_cp, "#3B82F6", height=28).pack(side="left")

        _divider(p, "2. INSTRUCTIONS", "#3B82F6")
        ctk.CTkLabel(p, text="What should the AI change?", font=ctk.CTkFont("Segoe UI",8), text_color=C["text3"]).pack(anchor="w", padx=14, pady=(4,0))
        
        self._cp_prompt = tk.Text(p, bg=C["card"], fg=C["text"], font=("Segoe UI",10), relief="flat", padx=8, pady=8, wrap="word", insertbackground="#3B82F6", height=6)
        self._cp_prompt.pack(fill="x", padx=12, pady=(2,12))
        self._cp_prompt.insert("end", "Example: Change the material of all Spur Gears to Al-6061 and increase their Face Width by 5mm.")

        _divider(p, "ACTIONS", "#3B82F6")
        self._btn_cp_preview = _run_button(p, "① PREVIEW AI CHANGES", self._run_cp_preview, "#3B82F6", height=32)
        self._btn_cp_preview.pack(fill="x", padx=12, pady=(6,4))
        
        self._btn_cp_apply = _pill_button(p, "② APPROVE & SAVE EXCEL", self._run_cp_apply, C["ok"], height=32)
        self._btn_cp_apply.pack(fill="x", padx=12, pady=(0,8))
        self._btn_cp_apply.configure(state="disabled") 

        self._cp_status = ctk.CTkLabel(p, text="● Waiting for instructions", font=ctk.CTkFont("Segoe UI",10,"bold"), text_color=C["text3"])
        self._cp_status.pack(pady=4)
        self._cp_prog = ctk.CTkProgressBar(p, fg_color=C["card"], progress_color="#3B82F6", height=6, corner_radius=3)
        self._cp_prog.set(0)
        self._cp_prog.pack(fill="x", padx=12, pady=(0,2))
        
        self._cp_prog_lbl = ctk.CTkLabel(p, text="", font=ctk.CTkFont("Segoe UI",8), text_color=C["text3"])
        self._cp_prog_lbl.pack(pady=(0,8))

    def _copilot_right(self, p):
        diff_lbl_fr = ctk.CTkFrame(p, fg_color="transparent")
        diff_lbl_fr.grid(row=0, column=0, sticky="ew", padx=10, pady=(10,2))
        ctk.CTkLabel(diff_lbl_fr, text="  📊 DATA DIFF VIEWER", font=ctk.CTkFont("Segoe UI",10,"bold"), text_color="#3B82F6").pack(side="left")
        
        diff_fr = ctk.CTkFrame(p, fg_color=C["card"], corner_radius=8)
        diff_fr.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0,6))
        
        self._cp_diff_txt = tk.Text(diff_fr, bg=C["card"], fg=C["ok"], font=FONT_MONO, relief="flat", padx=8, pady=8, state="disabled", wrap="none")
        sb_y = ctk.CTkScrollbar(diff_fr, command=self._cp_diff_txt.yview)
        sb_x = ctk.CTkScrollbar(diff_fr, command=self._cp_diff_txt.xview, orientation="horizontal")
        self._cp_diff_txt.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self._cp_diff_txt.pack(fill="both", expand=True)
        
        log_lbl_fr = ctk.CTkFrame(p, fg_color="transparent")
        log_lbl_fr.grid(row=2, column=0, sticky="ew", padx=10, pady=(4,2))
        ctk.CTkLabel(log_lbl_fr, text="  🧠 COPILOT REASONING LOG", font=ctk.CTkFont("Segoe UI",9,"bold"), text_color="#3B82F6").pack(side="left")
        
        log_fr, self._log4_txt = _log_widget(p)
        log_fr.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0,10))
        _append_log(self._log4_txt, "SYSTEM  AI Copilot initialized. Awaiting user commands.", "info")

    def _browse_cp(self):
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls"),("All","*.*")])
        if p: self._cp_file.set(p)

    def _run_cp_preview(self):
        self._btn_cp_preview.configure(state="disabled")
        self._btn_cp_apply.configure(state="disabled")
        threading.Thread(target=self._t_cp_preview, daemon=True).start()

    def _t_cp_preview(self):
        excel_path = self._cp_file.get()
        prompt = self._cp_prompt.get("1.0", "end").strip()
        key = os.environ.get("GEMINI_API_KEY", "").strip()

        if not excel_path or not os.path.exists(excel_path):
            self._q4.put(("log", "✘ Error: Select a valid Excel file first."))
            self._q4.put(("preview_fail", None))
            return
        
        self._q4.put(("prog4", (0.2, "Reading Excel & Analyzing...")))
        self._q4.put(("status4", ("● Processing", C["info"])))

        def t_log(msg): self._q4.put(("log", msg))

        try:
            # 🔗 CONNECT TO NEW COPILOT BACKEND
            from ai_bom_copilot import preview_bom_edits
            success, new_data, diff_text = preview_bom_edits(excel_path, prompt, key, t_log)

            if success:
                self._q4.put(("diff", diff_text))
                self._q4.put(("preview_success", new_data))
                self._q4.put(("prog4", (1.0, "Preview Ready!")))
                self._q4.put(("status4", ("✔ Ready for Approval", C["ok"])))
            else:
                self._q4.put(("preview_fail", None))
                self._q4.put(("status4", ("● Error", C["error"])))
                
        except ImportError:
            self._q4.put(("log", "✘ Error: ai_bom_copilot.py not found in directory."))
            self._q4.put(("preview_fail", None))
        except Exception as e:
            self._q4.put(("log", f"✘ Crash: {e}"))
            self._q4.put(("preview_fail", None))

    def _run_cp_apply(self):
        if not hasattr(self, '_pending_copilot_data') or not self._pending_copilot_data:
            return
            
        out_path = self._cp_file.get()
        prompt = self._cp_prompt.get("1.0", "end").strip()

        # Disable button and run save in thread to prevent GUI freezing
        self._btn_cp_apply.configure(state="disabled")
        threading.Thread(target=self._t_cp_apply, args=(out_path, prompt), daemon=True).start()

    def _t_cp_apply(self, out_path, prompt):
        try:
            self._q4.put(("log", "SYSTEM  Committing approved changes to Excel..."))
            
            # 🔗 CONNECT TO NEW COPILOT BACKEND (Saves preserving formulas)
            from ai_bom_copilot import commit_bom_edits
            success, msg = commit_bom_edits(
                excel_path=out_path,
                new_dicts=self._pending_copilot_data,
                log_cb=lambda m: self._q4.put(("log", m)),
                author="AI Copilot via GUI",
                prompt=prompt
            )
            
            if success:
                self._q4.put(("apply_success", msg))
            else:
                self._q4.put(("apply_fail", msg))
                
        except Exception as e:
            self._q4.put(("log", f"✘ System error during save: {e}"))
            self._q4.put(("apply_fail", str(e)))

    # ─────────────────────────────────────────────────────────────────────────
    # GLOBAL COST ENGINE INTEGRATION
    # ─────────────────────────────────────────────────────────────────────────
    def _get_current_log_q(self):
        """Returns the queue and progress key for the currently active tab."""
        curr = self._tabs.get()
        if "EXCEL" in curr: return self._q1, "prog1"
        elif "MANUAL" in curr: return self._q2, "prog2"
        elif "SHAPE" in curr: return self._q3, "prog3"
        elif "COPILOT" in curr: return self._q4, "prog4"
        return self._q1, "prog1"

    def _run_global_cost(self):
        self._btn_global_cost.configure(state="disabled")
        threading.Thread(target=self._t_run_global_cost, daemon=True).start()

    def _t_run_global_cost(self):
        q, prog_key = self._get_current_log_q()
        try:
            curr_tab = self._tabs.get()
            parts = []
            
            # 1. Determine data source based on active tab
            if "MANUAL" in curr_tab:
                parts = self._get_manual_parts()
                if not parts:
                    q.put(("log", "✘ Error: No parts defined in Manual Mode."))
                    return
            else:
                path = ""
                if "EXCEL" in curr_tab: path = self._e1_file.get()
                elif "COPILOT" in curr_tab: path = self._cp_file.get()
                
                # If path is empty/invalid, or we're on the AI Shape tab, prompt the user
                if not path or not os.path.exists(path):
                    path = filedialog.askopenfilename(title="Select BOM for ESG & Cost Analysis", filetypes=[("Excel", "*.xlsx *.xls")])
                    if not path: return
                        
                q.put(("log", f"SYSTEM  Generating Economic & ESG Report for {os.path.basename(path)}..."))
                
                from validator_3d import Validator3D
                v = Validator3D(path, log_callback=lambda m: None) # Silent validation
                v.run_checks()
                parts = v.valid_parts

            if not parts:
                q.put(("log", "✘ Error: No valid parts found to analyze."))
                return

            # 2. Run Cost Engine
            q.put((prog_key, (0.3, "Analyzing costs & carbon footprint...")))
            from cost_engine import CostEngine
            
            metal_key = os.environ.get("METALPRICE_API_KEY", "")
            gemini_key = os.environ.get("GEMINI_API_KEY", "").strip()
            q.put(("log", "SYSTEM  Fetching live metal prices and running cost analysis with AI..."))
            engine = CostEngine(metal_api_key=metal_key, gemini_api_key=gemini_key)
            
            engine.fetch_live_metal_prices() 
            
            out_pdf = os.path.abspath(f"output_reports/Siraal_ESG_Cost_Report_{int(time.time())}.pdf")
            success = engine.export_pdf_report(parts, out_pdf)
            
            if success:
                q.put(("log", f"✔ PDF Report Generated: {out_pdf}"))
                q.put((prog_key, (1.0, "Report Ready!")))
                try:
                    os.startfile(out_pdf) 
                except Exception as e: 
                    q.put(("log", f"Could not auto-open PDF: {e}"))
            else:
                q.put(("log", "✘ Failed to generate PDF (is fpdf2 installed?)"))

        except Exception as e:
            q.put(("log", f"✘ Crash during report generation: {e}\n{traceback.format_exc()}"))
        finally:
            self.after(0, lambda: self._btn_global_cost.configure(state="normal"))

    # ─────────────────────────────────────────────────────────────────────────
    # QUEUE POLLERS
    # ─────────────────────────────────────────────────────────────────────────
    def _poll_q1(self):
        try:
            while True:
                k,d = self._q1.get_nowait()
                if k=="log":    _append_log(self._log1_txt, str(d))
                elif k=="table1": self._update_tbl(self._e1_tbl, d[0], d[1])
                elif k=="stats":  self._e1_stats.configure(text=str(d))
                elif k=="prog1":  self._e1_prog.set(d[0]); self._e1_prog_lbl.configure(text=d[1])
                elif k=="status1": self._e1_status.configure(text=f"● {d[0]}", text_color=d[1])
                elif k=="btn1": self._btn_e1.configure(state="normal")
        except queue.Empty: pass
        except Exception as e: print(f"UI Queue 1 Error: {e}") # Prevents UI freezing if a log fails
        finally: self.after(100, self._poll_q1)

    def _poll_q2(self):
        try:
            while True:
                k,d = self._q2.get_nowait()
                if k=="log":    _append_log(self._log2_txt, str(d))
                elif k=="prog2": self._m_prog.set(d[0]); self._m_prog_lbl.configure(text=d[1])
                elif k=="status2": self._m_status.configure(text=f"● {d[0]}", text_color=d[1])
                elif k=="btn2": self._btn_m_run.configure(state="normal")
        except queue.Empty: pass
        except Exception as e: print(f"UI Queue 2 Error: {e}")
        finally: self.after(110, self._poll_q2)

    def _poll_q3(self):
        try:
            while True:
                k,d = self._q3.get_nowait()
                if k=="log":    _append_log(self._log3_txt, str(d))
                elif k=="prog3": self._ai_prog.set(d[0]); self._ai_prog_lbl.configure(text=d[1])
                elif k=="status3": self._ai_status.configure(text=f"● {d[0]}", text_color=d[1])
                elif k=="btn_ai_gen": self._btn_ai_gen.configure(state="normal")
                elif k=="btn_ai_run": self._btn_ai_run.configure(state="normal")
        except queue.Empty: pass
        except Exception as e: print(f"UI Queue 3 Error: {e}")
        finally: self.after(120, self._poll_q3)
        
    def _poll_q4(self):
        try:
            while True:
                k,d = self._q4.get_nowait()
                if k=="log":    _append_log(self._log4_txt, str(d))
                elif k=="prog4": self._cp_prog.set(d[0]); self._cp_prog_lbl.configure(text=d[1])
                elif k=="status4": self._cp_status.configure(text=f"{d[0]}", text_color=d[1])
                elif k=="diff":
                    self._cp_diff_txt.configure(state="normal")
                    self._cp_diff_txt.delete("1.0", "end")
                    self._cp_diff_txt.insert("end", d)
                    self._cp_diff_txt.configure(state="disabled")
                elif k=="preview_success":
                    self._pending_copilot_data = d
                    self._btn_cp_preview.configure(state="normal")
                    self._btn_cp_apply.configure(state="normal") 
                elif k=="preview_fail":
                    self._btn_cp_preview.configure(state="normal")
                    self._btn_cp_apply.configure(state="disabled")
                elif k=="apply_success":
                    messagebox.showinfo("Copilot Success", d)
                    self._cp_status.configure(text="● Excel Saved", text_color=C["ok"])
                    self._btn_cp_preview.configure(state="normal")
                elif k=="apply_fail":
                    self._btn_cp_apply.configure(state="normal")
                    self._btn_cp_preview.configure(state="normal")
        except queue.Empty: pass
        except Exception as e: print(f"UI Queue 4 Error: {e}")
        finally: self.after(130, self._poll_q4)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if __name__ == "__main__":
    app = SiraalGUI()
    app.mainloop()