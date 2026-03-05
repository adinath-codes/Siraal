"""
autocad_engine.py — Siraal Grand Unified Manufacturing Engine
Advanced AutoCAD COM controller: Plate, Spur_Gear, Stepped_Shaft, Flanged_Shaft, Ring_Gear
ISO title block • Auto-dimensioning • DXF/DWG export • Live market pricing
"""

import win32com.client
import pythoncom
import math
import os
import time
import logging
import requests
from typing import Callable, Optional, List, Dict, Tuple

logger = logging.getLogger("Siraal.CADEngine")

# ── AutoCAD SaveAs format constants ─────────────────────────────────────────
DWG_SAVE_FORMATS = [67, 64, 61]   # R2018, R2013, R2010 — tries in order
DXF_FORMAT       = 12

# ── ISO standard lineweights (in hundredths of mm) ──────────────────────────
LW = {"thin": 13, "medium": 30, "thick": 50}

# ── Material database (Rs./kg) ───────────────────────────────────────────────
DEFAULT_MATERIAL_DB: Dict[str, dict] = {
    "Steel-1020": {"density": 7.87, "cost_per_kg": 125.00,  "color": 7},
    "Steel-4140": {"density": 7.85, "cost_per_kg": 185.00,  "color": 7},
    "Al-6061":    {"density": 2.70, "cost_per_kg": 265.00,  "color": 4},
    "Brass-C360": {"density": 8.50, "cost_per_kg": 520.00,  "color": 2},
    "Nylon-66":   {"density": 1.14, "cost_per_kg": 415.00,  "color": 3},
    "Ti-6Al-4V":  {"density": 4.43, "cost_per_kg": 3800.00, "color": 6},
}

LAYERS = [
    # (name, ACI-color, lineweight, linetype)
    ("01_Visible_Geometry",  7, LW["thick"],  "Continuous"),
    ("02_Hidden_Geometry",   8, LW["thin"],   "HIDDEN"),
    ("03_Centerlines",       1, LW["thin"],   "CENTER"),
    ("04_Dimensions",        3, LW["thin"],   "Continuous"),
    ("05_Title_Block",       7, LW["medium"], "Continuous"),
    ("06_Hatch",             8, LW["thin"],   "Continuous"),
    ("07_Notes",             2, LW["thin"],   "Continuous"),
]

# ════════════════════════════════════════════════════════════════════════════
class AutoCADController:
# ════════════════════════════════════════════════════════════════════════════

    def __init__(self, log_callback: Optional[Callable] = None, session_name: str = "Session_Default"):
        self._log = log_callback or print
        self.session_name = session_name
        self.material_db = {k: v.copy() for k, v in DEFAULT_MATERIAL_DB.items()}
        self._log(f"[*] Booting Siraal Engine v3.0 for {self.session_name}...")

        self._clear_gen_py_cache()

        self.acad = win32com.client.dynamic.Dispatch("AutoCAD.Application")
        self.acad.Visible = True

        self.doc = self.acad.Documents.Add()
        self.ms = win32com.client.dynamic.Dispatch(self.doc.ModelSpace)
        self._log("[*] Fresh Master AutoCAD document created.")

        self._setup_document_env(self.doc)
        self._update_live_prices()
    def _setup_document_env(self, doc):
        """Applies ISO dimension styles, linetypes, and layers to a specific document."""
        doc.SetVariable("LWDISPLAY", 1)
        doc.SetVariable("DIMSCALE",  1)
        doc.SetVariable("DIMASZ",    5)
        doc.SetVariable("DIMTXT",    5)
        doc.SetVariable("DIMEXE",    3)
        doc.SetVariable("DIMEXO",    3)
        doc.SetVariable("DIMDLI",   10)

        for lt in ("CENTER", "HIDDEN", "DASHED"):
            try:
                doc.Linetypes.Load(lt, "acad.lin")
            except Exception:
                pass

        for name, color, weight, linetype in LAYERS:
            try:
                layer = doc.Layers.Add(name)
                layer.Color, layer.Lineweight, layer.Linetype = color, weight, linetype
            except Exception:
                pass

    @staticmethod
    def _clear_gen_py_cache():
        import shutil
        try:
            import win32com as _wc
            gen_py_path = os.path.join(os.path.dirname(_wc.__file__), "gen_py")
            if os.path.exists(gen_py_path):
                shutil.rmtree(gen_py_path, ignore_errors=True)
                print(f"[*] Cleared win32com gen_py cache.")
        except Exception:
            pass 

    # ── Logging ──────────────────────────────────────────────────────────────
    def _log_info(self, msg):
        logger.info(msg)
        self._log(msg)

    # ── COM Primitives ────────────────────────────────────────────────────────
    def _pnt(self, x, y, z=0):
        return win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))

    def _arr(self, coords):
        return win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(c) for c in coords])

    # ── Draw helpers ──────────────────────────────────────────────────────────
    def _line(self, x1, y1, x2, y2, layer="01_Visible_Geometry"):
        obj = self.ms.AddLine(self._pnt(x1, y1), self._pnt(x2, y2))
        obj.Layer = layer
        return obj

    def _circle(self, cx, cy, r, layer="01_Visible_Geometry"):
        obj = self.ms.AddCircle(self._pnt(cx, cy), float(r))
        obj.Layer = layer
        return obj

    def _pline(self, coords, closed=True, layer="01_Visible_Geometry"):
        obj = self.ms.AddLightWeightPolyline(self._arr(coords))
        obj.Closed = closed
        obj.Layer  = layer
        return obj

    def _text(self, txt, x, y, height=5, layer="05_Title_Block"):
        obj = self.ms.AddText(str(txt), self._pnt(x, y), float(height))
        obj.Layer = layer
        return obj

    def _rect(self, x, y, w, h, layer="01_Visible_Geometry"):
        return self._pline([x, y, x+w, y, x+w, y+h, x, y+h], closed=True, layer=layer)

    def _center_mark(self, cx, cy, r, layer="03_Centerlines"):
        ext = max(r * 1.35, r + 10)
        self._line(cx - ext, cy, cx + ext, cy, layer)
        self._line(cx, cy - ext, cx, cy + ext, layer)

    # ── Dimensioning ──────────────────────────────────────────────────────────
    def _dim_linear(self, x1, y1, x2, y2, text_pt, layer="04_Dimensions"):
        try:
            dim = self.ms.AddDimAligned(
                self._pnt(x1, y1), self._pnt(x2, y2), self._pnt(*text_pt))
            dim.Layer = layer
            return dim
        except Exception:
            pass

    def _dim_diameter(self, cx, cy, r, angle_deg=0, layer="04_Dimensions"):
        try:
            angle = math.radians(angle_deg)
            dim = self.ms.AddDimDiametric(
                self._pnt(cx + r * math.cos(angle), cy + r * math.sin(angle)),
                self._pnt(cx - r * math.cos(angle), cy - r * math.sin(angle)),
                float(r * 1.8))
            dim.Layer = layer
            return dim
        except Exception:
            pass

    # ── Live Market Pricing ───────────────────────────────────────────────────
    def _update_live_prices(self):
        api_key = os.environ.get("METALPRICE_API_KEY", "")
        if not api_key:
            self._log_info("[-] No METALPRICE_API_KEY env var — using offline prices.")
            return
        try:
            url = (f"https://api.metalpriceapi.com/v1/latest"
                   f"?api_key={api_key}&currencies=INR,XAG,XAU")
            r = requests.get(url, timeout=8)
            data = r.json()
            if data.get("success"):
                rates = data["rates"]
                inr = rates.get("INR", 1)
                xag = rates.get("XAG")
                if xag and inr:
                    al_inr_kg = (1 / xag) * inr / 0.0283495
                    self.material_db["Al-6061"]["cost_per_kg"] = round(al_inr_kg, 2)
                    self._log_info(f"[+] Live: Al-6061 → ₹{al_inr_kg:.2f}/kg")
        except Exception as e:
            self._log_info(f"[-] Live pricing unavailable: {e}")

    # ── ERP Calculations ──────────────────────────────────────────────────────
    def _calc_specs(self, part_type, p1, p2, p3, p4, material) -> Tuple[float, float, float]:
        vol = 0.0
        pt = part_type.strip()

        if pt == "Spur_Gear":
            od, bd = (p1 * p2 + 2 * p2) / 2, p4 / 2
            vol = math.pi * (od**2 - bd**2) * p3
        elif pt == "Ring_Gear":
            inner_r = (p1 * p2) / 2.0 - p2
            outer_r = inner_r + p4
            vol = math.pi * (outer_r**2 - inner_r**2) * p3
        elif pt == "Stepped_Shaft":
            vol = math.pi * ((p2/2)**2) * p1 + math.pi * ((p4/2)**2) * p3
        elif pt == "Flanged_Shaft":
            vol = (math.pi * ((p2/2)**2) * p1) + (math.pi * ((p3/2)**2) * p4)
        else: 
            vol = p1 * p2 * p3 - 4 * math.pi * ((p4/2)**2) * p3

        vol_cm3 = max(vol, 0) / 1000.0
        mat = self.material_db.get(material, self.material_db["Steel-1020"])
        mass  = round((vol_cm3 * mat["density"]) / 1000.0, 3)
        cost  = round(mass * mat["cost_per_kg"], 2)
        return vol_cm3, mass, cost

    def _bbox(self, ptype, p1, p2, p3, p4) -> Tuple[float, float]:
        DIM_CLEAR = 80.0
        if ptype == "Plate": return p1 + DIM_CLEAR, p2 + DIM_CLEAR
        elif ptype == "Spur_Gear":
            d = ((p1 * p2) / 2.0 + p2) * 2
            return d + DIM_CLEAR, d + DIM_CLEAR
        elif ptype == "Ring_Gear":
            d = ((p1 * p2) / 2.0 - p2 + p4) * 2
            return d + DIM_CLEAR, d + DIM_CLEAR
        elif ptype == "Stepped_Shaft": return p1 + p3 + DIM_CLEAR, max(p2, p4) + DIM_CLEAR
        elif ptype == "Flanged_Shaft": return p1 + p4 + DIM_CLEAR, max(p2, p3) + DIM_CLEAR
        return 300.0, 200.0

    # ── Shape Drawers ─────────────────────────────────────────────────────────
    def _draw_plate(self, ox, oy, L, W, hole_d):
        self._rect(ox, oy, L, W)
        inset = max(20.0, hole_d * 1.5)
        holes = [(ox+inset, oy+inset), (ox+L-inset, oy+inset), (ox+inset, oy+W-inset), (ox+L-inset, oy+W-inset)]
        for hx, hy in holes:
            self._circle(hx, hy, hole_d / 2)
            self._center_mark(hx, hy, hole_d / 2)
        self._dim_linear(ox, oy-35, ox+L, oy-35, (ox+L/2, oy-52))
        self._dim_linear(ox+L+35, oy, ox+L+35, oy+W, (ox+L+55, oy+W/2))
        if hole_d > 0: self._dim_diameter(holes[0][0], holes[0][1], hole_d / 2, 45)

    def _draw_spur_gear(self, cx, cy, Z, m, bore_d):
        Z, m = float(int(Z)), float(m)
        pitch_r, outer_r, root_r, bore_r = (Z*m)/2, (Z*m)/2 + m, (Z*m)/2 - 1.25*m, bore_d/2.0
        self._center_mark(cx, cy, outer_r)
        self.ms.AddCircle(self._pnt(cx, cy), pitch_r).Layer = "03_Centerlines"
        self._circle(cx, cy, bore_r)
        kw_w, kw_h = bore_d * 0.25, bore_d * 0.12
        self._rect(cx - kw_w/2, cy + bore_r - kw_h/2, kw_w, kw_h)
        ang = (2 * math.pi) / Z; th = ang / 2.0
        for i in range(int(Z)):
            a = i * ang
            pts = [
                (cx + root_r * math.cos(a),             cy + root_r * math.sin(a)),
                (cx + outer_r * math.cos(a + th*0.2), cy + outer_r * math.sin(a + th*0.2)),
                (cx + outer_r * math.cos(a + th*0.8), cy + outer_r * math.sin(a + th*0.8)),
                (cx + root_r * math.cos(a + th),      cy + root_r * math.sin(a + th)),
                (cx + root_r * math.cos(a + ang),     cy + root_r * math.sin(a + ang)),
            ]
            for j in range(len(pts)-1): self._line(*pts[j], *pts[j+1])
        self._dim_diameter(cx, cy, outer_r, 30)
        self._dim_diameter(cx, cy, bore_r,  135)

    def _draw_ring_gear(self, cx, cy, Z, m, ring_t):
        Z, m = float(int(Z)), float(m)
        pitch_r = (Z * m) / 2
        inner_r, outer_r = pitch_r - m, pitch_r - m + ring_t
        self._center_mark(cx, cy, outer_r)
        self.ms.AddCircle(self._pnt(cx, cy), pitch_r).Layer = "03_Centerlines"
        self._circle(cx, cy, outer_r)
        self._circle(cx, cy, inner_r)
        ang = (2 * math.pi) / Z; th = ang / 2.0
        for i in range(int(Z)):
            a = i * ang
            pts = [
                (cx + inner_r * math.cos(a),                 cy + inner_r * math.sin(a)),
                (cx + (inner_r-m) * math.cos(a + th*0.25), cy + (inner_r-m) * math.sin(a + th*0.25)),
                (cx + (inner_r-m) * math.cos(a + th*0.75), cy + (inner_r-m) * math.sin(a + th*0.75)),
                (cx + inner_r * math.cos(a + th),          cy + inner_r * math.sin(a + th)),
                (cx + inner_r * math.cos(a + ang),         cy + inner_r * math.sin(a + ang)),
            ]
            for j in range(len(pts)-1): self._line(*pts[j], *pts[j+1])
        self._dim_diameter(cx, cy, outer_r, 30)
        self._dim_diameter(cx, cy, inner_r, 60)

    def _draw_stepped_shaft(self, ox, cy, L1, D1, L2, D2):
        r1, r2 = D1/2, D2/2
        self._rect(ox, cy-r1, L1, D1)
        self._rect(ox+L1, cy-r2, L2, D2)
        self._line(ox-25, cy, ox+L1+L2+25, cy, "03_Centerlines")
        ch = min(4.0, min(D1,D2)*0.05)
        self._line(ox, cy+r1, ox+ch, cy+r1-ch)
        self._line(ox, cy-r1, ox+ch, cy-r1+ch)
        self._line(ox+L1, cy+r2, ox+L1, cy+r1)
        self._line(ox+L1, cy-r2, ox+L1, cy-r1)
        self._dim_linear(ox, cy-r1-35, ox+L1, cy-r1-35, (ox+L1/2, cy-r1-55))
        self._dim_linear(ox+L1, cy-r2-35, ox+L1+L2, cy-r2-35, (ox+L1+L2/2, cy-r2-55))
        self._dim_linear(ox-35, cy-r1, ox-35, cy+r1, (ox-60, cy))
        self._dim_linear(ox+L1+L2+35, cy-r2, ox+L1+L2+35, cy+r2, (ox+L1+L2+60, cy))

    def _draw_flanged_shaft(self, ox, cy, sl, sd, fod, ft):
        rs, rf = sd/2, fod/2
        self._rect(ox, cy-rs, sl, sd)
        fx = ox + sl
        self._rect(fx, cy-rf, ft, fod)
        self._line(ox-25, cy, fx+ft+25, cy, "03_Centerlines")
        bpcd = (rf+rs)/2
        br = max(2.0, min((rf-rs)/3.5, 6.0))
        self.ms.AddCircle(self._pnt(fx+ft/2, cy), bpcd).Layer = "03_Centerlines"
        for i in range(6):
            a = math.radians(i*60)
            bx, by = fx+ft/2 + bpcd*math.cos(a), cy + bpcd*math.sin(a)
            self._circle(bx, by, br)
            self._center_mark(bx, by, br)
        self._dim_linear(ox, cy-rs-35, fx, cy-rs-35, (ox+sl/2, cy-rs-55))
        self._dim_linear(ox-35, cy-rs, ox-35, cy+rs, (ox-60, cy))
        self._dim_linear(fx-35, cy-rf, fx-35, cy+rf, (fx-65, cy))

    # ── ISO Title Block ───────────────────────────────────────────────────────
    def _draw_title_block(self, ox, oy, w, part, mass, cost, vol):
        H, mid, ROW_H = 55.0, ox + w * 0.5, 55.0 / 3.0
        self._rect(ox, oy, w, H, "05_Title_Block")
        self._line(ox, oy+ROW_H, ox+w, oy+ROW_H, "05_Title_Block")
        self._line(ox, oy+ROW_H*2, ox+w, oy+ROW_H*2, "05_Title_Block")
        self._line(mid, oy+ROW_H, mid, oy+ROW_H*2, "05_Title_Block")
        self._line(ox+w*0.72, oy+ROW_H*2, ox+w*0.72, oy+H, "05_Title_Block")
        self._line(ox+w*0.25, oy, ox+w*0.25, oy+ROW_H, "05_Title_Block")
        self._line(ox+w*0.65, oy, ox+w*0.65, oy+ROW_H, "05_Title_Block")
        self._text("SIRAAL MANUFACTURING SYSTEMS  |  TN-IMPACT 2026", ox+8, oy+ROW_H*2+5, 6, "05_Title_Block")
        self._text("REV: A", ox+w*0.72+6, oy+ROW_H*2+5, 5, "05_Title_Block")
        RS, base = ROW_H / 3.5, oy + ROW_H + 2
        left  = [(f"PART NO : {part['Part_Number']}", ox+6, base+RS*2),
                 (f"TYPE    : {part['Part_Type']}", ox+6, base+RS*1),
                 (f"MATL    : {part['Material']}", ox+6, base+RS*0)]
        right = [(f"MASS    : {mass} kg", mid+6, base+RS*2),
                 (f"VOLUME  : {vol:.2f} cm3", mid+6, base+RS*1),
                 (f"COST    : Rs. {cost:,.2f}", mid+6, base+RS*0)]
        for txt, tx, ty in left + right: self._text(txt, tx, ty, 4.5, "05_Title_Block")
        self._text(f"P1={part['Param_1']}  P2={part['Param_2']}", ox+4, oy+3, 4.0, "05_Title_Block")
        self._text(f"P3={part['Param_3']}  P4={part['Param_4']}", ox+w*0.25+4, oy+3, 4.0, "05_Title_Block")
        self._text(f"QTY:{int(part.get('Quantity',1))}  {part.get('Priority','Med')}", ox+w*0.65+4, oy+3, 4.0, "05_Title_Block")

    def _draw_border(self, ox, oy, w, h, pno):
        self._rect(ox, oy, w, h, "05_Title_Block")
        self._rect(ox+8, oy+8, w-16, h-16, "05_Title_Block")
        self._text(pno, ox+14, oy+h-22, 7, "05_Title_Block")

# ── Strict Isolated DXF Export ───────────────────────────────────────────
# ── Strict Isolated DXF Export (Stabilized) ──────────────────────────────
    def _export_all_dxf(self, parts: List[dict]):
        # Use the session name provided by the GUI
        out_dir = os.path.join(os.getcwd(), "CNC_Machine_Files", self.session_name)
        os.makedirs(out_dir, exist_ok=True)
        
        self._log_info(f"\n[*] Exporting {len(parts)} strictly isolated CNC DXF files to '{out_dir}'...")

        for part in parts:
            pno   = part.get("Part_Number", "UNK")
            ptype = str(part.get("Part_Type", "Plate")).strip()
            p1, p2, p3, p4 = (float(part.get(f"Param_{i}", 100)) for i in range(1, 5))

            # Backup context BEFORE attempting anything
            orig_ms, orig_doc = self.ms, self.doc
            tmp_doc = None

            try:
                # 1. Create a brand new document
                tmp_doc = self.acad.Documents.Add()
                
                # STABILIZER 1: Let AutoCAD catch its breath and initialize the document
                time.sleep(0.5) 
                
                # STABILIZER 2: Explicitly force AutoCAD focus to the new document
                tmp_doc.Activate() 
                time.sleep(0.2)

                tmp_ms  = win32com.client.dynamic.Dispatch(tmp_doc.ModelSpace)
                
                # 2. Swap the drawing context FIRST
                self.ms, self.doc = tmp_ms, tmp_doc

                # 3. Inject identical layers/dims
                self._setup_document_env(tmp_doc)

                # 4. Draw geometry precisely at origin (0,0)
                geo_w, geo_h = self._bbox(ptype, p1, p2, p3, p4)
                cx, cy = geo_w / 2.0, geo_h / 2.0

                if ptype == "Plate":           self._draw_plate(0, 0, p1, p2, p4)
                elif ptype == "Spur_Gear":     self._draw_spur_gear(cx, cy, p1, p2, p4)
                elif ptype == "Ring_Gear":     self._draw_ring_gear(cx, cy, p1, p2, p4)
                elif ptype == "Stepped_Shaft": self._draw_stepped_shaft(0, cy, p1, p2, p3, p4)
                elif ptype == "Flanged_Shaft": self._draw_flanged_shaft(0, cy, p1, p2, p3, p4)

                # Frame the view perfectly for the CNC export
                self.acad.ZoomExtents()

                # 5. Export clean DXF
                path = os.path.abspath(os.path.join(out_dir, f"{pno}_CNC.dxf"))
                if os.path.exists(path):
                    try:
                        os.remove(path)
                    except: pass # File might be locked by OS
                
                tmp_doc.SaveAs(path, DXF_FORMAT)
                self._log_info(f"    ✔ {pno}_CNC.dxf")
                
                # STABILIZER 3: Give the hard drive time to finish writing the DXF file
                time.sleep(0.5) 

            except Exception as e:
                self._log_info(f"    ✘ DXF failed for {pno}: {e}")

            finally:
                # 6. Safely restore Master Document context (ALWAYS runs, even on crash)
                self.ms, self.doc = orig_ms, orig_doc
                
                # Shift AutoCAD's focus back to the master assembly
                if orig_doc:
                    try:
                        orig_doc.Activate()
                    except: pass
                
                # Close the temp document without saving DWG changes
                if tmp_doc:
                    try:
                        tmp_doc.Close(False)
                    except Exception: pass
                
                # STABILIZER 4: Breather before generating the next document
                time.sleep(0.3)
    def _save_dwg(self, path: str):
        abs_path = os.path.abspath(path)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        for fmt in DWG_SAVE_FORMATS:
            try:
                self.doc.SaveAs(abs_path, fmt)
                self._log_info(f"[+] Saved DWG (fmt={fmt}): {abs_path}")
                return
            except Exception:
                continue
        self.doc.SaveAs(abs_path)

    
    # ── Master Batch Generator ────────────────────────────────────────────────
    def generate_batch(self, parts: List[dict], status_callback: Optional[Callable] = None):
        TB_H, PAD, GAP_TB, BORD, PART_GAP = 55.0, 20.0, 8.0, 10.0, 60.0
        current_x = 0.0
        failed = []

        for i, part in enumerate(parts):
            pno   = part.get("Part_Number", "UNK")
            ptype = str(part.get("Part_Type", "Plate")).strip()
            mat   = part.get("Material", "Steel-1020")
            p1, p2, p3, p4 = (float(part.get(f"Param_{i}", 100)) for i in range(1, 5))

            self._log_info(f"\n[*] GENERATING: {pno} | {ptype} | {mat}")
            
            # Trigger the Tkinter UI to say "Drawing..."
            if status_callback:
                status_callback(pno, "⚙ Drawing…", i / len(parts))

            try:
                vol, mass, cost = self._calc_specs(ptype, p1, p2, p3, p4, mat)
                self._log_info(f"    ERP → Mass: {mass} kg | Cost: Rs.{cost:,.2f} | Vol: {vol:.2f} cm3")

                geo_w, geo_h = self._bbox(ptype, p1, p2, p3, p4)
                frame_w = geo_w + PAD*2 + BORD*2
                frame_h = geo_h + PAD*2 + BORD*2 + GAP_TB + TB_H

                gc_ox, gc_oy = current_x + BORD + PAD, BORD + TB_H + GAP_TB + PAD
                tb_ox, tb_oy, tb_w = current_x + BORD, BORD, frame_w - BORD*2

                self._draw_border(current_x, 0, frame_w, frame_h, pno)

                if ptype == "Plate":
                    ox, oy = gc_ox + (geo_w - p1) / 2.0, gc_oy + (geo_h - p2) / 2.0
                    self._log_info("    -> Plate with corner holes...")
                    self._draw_plate(ox, oy, p1, p2, p4)

                elif ptype == "Spur_Gear":
                    cx, cy = gc_ox + geo_w/2.0, gc_oy + geo_h/2.0
                    self._log_info("    -> Involute spur gear...")
                    self._draw_spur_gear(cx, cy, p1, p2, p4)

                elif ptype == "Ring_Gear":
                    cx, cy = gc_ox + geo_w/2.0, gc_oy + geo_h/2.0
                    self._log_info("    -> Ring gear...")
                    self._draw_ring_gear(cx, cy, p1, p2, p4)

                elif ptype == "Stepped_Shaft":
                    ox, cy = gc_ox + (geo_w - (p1 + p3)) / 2.0, gc_oy + geo_h / 2.0
                    self._log_info("    -> Stepped shaft...")
                    self._draw_stepped_shaft(ox, cy, p1, p2, p3, p4)

                elif ptype == "Flanged_Shaft":
                    ox, cy = gc_ox + (geo_w - (p1 + p4)) / 2.0, gc_oy + geo_h / 2.0
                    self._log_info("    -> Flanged shaft...")
                    self._draw_flanged_shaft(ox, cy, p1, p2, p3, p4)

                self._draw_title_block(tb_ox, tb_oy, tb_w, part, mass, cost, vol)

                # ADVANCE THE X-AXIS FOR THE NEXT PART
                current_x += frame_w + PART_GAP
                time.sleep(0.05)
                
                # Trigger the Tkinter UI to say "Done"
                if status_callback:
                    status_callback(pno, "✔ Done", None)

            except Exception as e:
                self._log_info(f"  ERROR drafting {pno}: {e}")
                failed.append(pno)
                if status_callback:
                    status_callback(pno, "✘ Error", None)

        self.acad.ZoomExtents()
        
        # Save Master Assembly into the Session Folder
        master_out_dir = os.path.join(os.getcwd(), "CNC_Machine_Files", self.session_name)
        os.makedirs(master_out_dir, exist_ok=True)
        master_path = os.path.join(master_out_dir, "Master_Assembly.dwg")
        
        self._save_dwg(master_path)
        self._log_info(f"\n[+] Master DWG saved in session folder: {master_path}")
