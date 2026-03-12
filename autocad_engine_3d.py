"""
autocad_engine_3d.py  —  Siraal 3D Gear Engine  v6.2
=====================================================
BUGS FIXED & UPGRADES FROM v6.1:
─────────────────────────────────
  UPGRADE: Dynamic Folder Architecture.
    - Creates a unique 'Session_[TIMESTAMP]' folder for each batch run.
    - Creates an isolated sub-folder for every single part generated.
    - Master Assembly DWG is saved neatly inside the Session folder.
"""

import win32com.client
import win32com.client.dynamic
import pythoncom
import math
import os
import time
import logging
import shutil
from typing import Callable, List, Optional, Tuple
import pythoncom
logger = logging.getLogger("Siraal.GearEngine")

# ── AutoCAD COM Boolean constants ─────────────────────────────────────────────
AC_UNION     = 0
AC_INTERSECT = 1
AC_SUBTRACT  = 2  # CRITICAL FIX: 2 is Subtraction in AutoCAD COM

DWG_SAVE_FORMATS = [67, 64, 61, 60, 48]   # 2023→2018→2013→2010→2007
DXF_FORMAT       = 12                      # R2010 DXF

MATERIAL_DB = {
    "Steel-1020": {"density": 7.87e-3, "cost_per_kg":  125.0, "color": (190,190,200)},
    "Steel-4140": {"density": 7.85e-3, "cost_per_kg":  185.0, "color": (255,185, 15)},
    "Steel-EN36": {"density": 7.85e-3, "cost_per_kg":  220.0, "color": (210,200,180)},
    "Al-6061":    {"density": 2.70e-3, "cost_per_kg":  265.0, "color": (200,215,235)},
    "Al-7075":    {"density": 2.81e-3, "cost_per_kg":  390.0, "color": (185,205,225)},
    "Brass-C360": {"density": 8.50e-3, "cost_per_kg":  520.0, "color": (210,155, 30)},
    "Bronze-C93": {"density": 8.83e-3, "cost_per_kg":  680.0, "color": (180,120, 40)},
    "Nylon-66":   {"density": 1.14e-3, "cost_per_kg":  415.0, "color": (240,230,205)},
    "Ti-6Al-4V":  {"density": 4.43e-3, "cost_per_kg": 3800.0, "color": (160,160,175)},
    "Cast-Iron":  {"density": 7.20e-3, "cost_per_kg":   95.0, "color": (100,100,100)},
}

LAYERS = [
    ("GEAR_SOLID",   2,  70, "Continuous"),
    ("GEAR_TEETH",   4,  70, "Continuous"),
    ("GEAR_BLANK",   6,  50, "Continuous"),
    ("GEAR_BORE",    1,  35, "Continuous"),
    ("WORK_GEOM",    8,  13, "Continuous"),
    ("TITLE_BLOCK",  7,  35, "Continuous"),
    ("VIEW_BORDER",  2,  18, "Continuous"),
    ("HIDDEN_LINE",  8,  18, "HIDDEN"),
    ("CENTRE_LINE",  1,  13, "CENTER"),
]

VIEWPORTS = {
    "FRONT": {"eye":(0,-1, 0),"up":(0,0,1),"ox":12, "oy":148,"w":125,"h":105},
    "TOP":   {"eye":(0, 0, 1),"up":(0,1,0),"ox":148,"oy":148,"w":125,"h":105},
    "RIGHT": {"eye":(1, 0, 0),"up":(0,0,1),"ox":284,"oy":148,"w":120,"h":105},
    "ISO":   {"eye":(1,-1, 1),"up":(0,0,1),"ox":12, "oy":58, "w":185,"h": 82},
}


# ══════════════════════════════════════════════════════════════════════════════
#  INVOLUTE MATHEMATICS  (pure Python, no AutoCAD dependency)
# ══════════════════════════════════════════════════════════════════════════════

def _inv_pt(base_r: float, t: float) -> Tuple[float, float]:
    return (base_r*(math.cos(t) + t*math.sin(t)),
            base_r*(math.sin(t) - t*math.cos(t)))

def profile_shift_x(Z: int, PA_deg: float = 20.0) -> float:
    alpha = math.radians(PA_deg)
    z_min = 2.0 / math.sin(alpha)**2
    return round(max(0.0, (z_min - Z) / z_min), 4) if Z < z_min else 0.0

def single_tooth_flat(Z: int, m: float, angle: float,
                       PA_deg: float = 20.0, x: float = 0.0,
                       N: int = 48) -> List[float]:
    alpha    = math.radians(PA_deg)
    pitch_r  = Z * m / 2.0
    base_r   = pitch_r * math.cos(alpha)
    outer_r  = pitch_r + m * (1.0 + x)
    root_r   = max(pitch_r - m * (1.25 - x), base_r * 0.05, m * 0.3)

    t_max   = math.sqrt(max((outer_r / base_r)**2 - 1.0, 0.0))
    t_min   = math.sqrt(max((max(root_r, base_r) / base_r)**2 - 1.0, 1e-9))
    t_pitch = math.sqrt(max((pitch_r / base_r)**2 - 1.0, 0.0))

    phi_pitch  = t_pitch - math.atan(t_pitch)
    tooth_half = math.pi / (2*Z) + 2*x*math.tan(alpha) / Z
    r_off      = tooth_half + phi_pitch - t_pitch

    def rpt(t):
        ix, iy = _inv_pt(base_r, t)
        r = math.hypot(ix, iy);  a = math.atan2(iy, ix) + r_off
        return r*math.cos(a), r*math.sin(a)

    def lpt(t):
        ix, iy = _inv_pt(base_r, t)
        r = math.hypot(ix, iy);  a = math.atan2(-iy, ix) - r_off
        return r*math.cos(a), r*math.sin(a)

    pts: List[Tuple[float,float]] = []

    for k in range(N + 1):
        t = t_min + (t_max - t_min) * k / N
        pts.append(rpt(t))

    rt = rpt(t_max);  lt = lpt(t_max)
    a_rt = math.atan2(rt[1], rt[0]);  a_lt = math.atan2(lt[1], lt[0])
    da   = a_lt - a_rt
    if da >  math.pi: da -= 2*math.pi
    if da < -math.pi: da += 2*math.pi
    N_tip = max(12, N // 3)
    for k in range(1, N_tip):
        a = a_rt + da * k / N_tip
        pts.append((outer_r*math.cos(a), outer_r*math.sin(a)))

    for k in range(N, -1, -1):
        t = t_min + (t_max - t_min) * k / N
        pts.append(lpt(t))

    close_r = root_r * 0.75
    lf = lpt(t_min);  rf = rpt(t_min)
    a_lf = math.atan2(lf[1], lf[0]);  a_rf = math.atan2(rf[1], rf[0])
    da_r = a_rf - a_lf
    if da_r >  math.pi: da_r -= 2*math.pi
    if da_r < -math.pi: da_r += 2*math.pi
    N_root = max(8, N // 5)
    for k in range(1, N_root + 1):
        a = a_lf + da_r * k / N_root
        pts.append((close_r*math.cos(a), close_r*math.sin(a)))

    ca, sa = math.cos(angle), math.sin(angle)
    flat: List[float] = []
    for px, py in pts:
        flat.append(px*ca - py*sa)
        flat.append(px*sa + py*ca)
    return flat


# ══════════════════════════════════════════════════════════════════════════════
#  ENGINE CLASS
# ══════════════════════════════════════════════════════════════════════════════

class AutoCAD3DGearEngine:

    def __init__(self, log_cb: Optional[Callable] = None):
        pythoncom.CoInitialize()
        self._log_cb = log_cb or print
        self._log("╔════════════════════════════════════════════════════╗")
        self._log("║  SIRAAL GEAR ENGINE v6.2 — DYNAMIC ARCHITECTURE    ║")
        self._log("║  Auto-creates Session & Part Isolation Folders     ║")
        self._log("╚════════════════════════════════════════════════════╝")
        self._purge_gen_py()

        self.acad = win32com.client.dynamic.Dispatch("AutoCAD.Application")
        self.acad.Visible = True
        self.doc  = self.acad.Documents.Add()
        self.ms   = win32com.client.dynamic.Dispatch(self.doc.ModelSpace)
        self._log("[*] AutoCAD connected — new document.")

        for var, val in [("ISOLINES",32),("FACETRES",5),("DISPSILH",1),
                         ("LWDISPLAY",1),("DIMSCALE",1),("DELOBJ",1)]:
            try: self.doc.SetVariable(var, val)
            except Exception: pass
        for lt in ("CENTER","HIDDEN"):
            try: self.doc.Linetypes.Load(lt,"acad.lin")
            except Exception: pass
        for name,color,weight,lt in LAYERS:
            self._mk_layer(name,color,weight,lt)
        self._log("[*] Initialisation complete.\n")

    # ── Utilities ───────────────────────────────────────────────────────────

    def _do(self, func, *args):
        """Robust COM Auto-Retry wrapper. Prevents 'Call was rejected by callee'."""
        for _ in range(30):
            try: return func(*args)
            except Exception as e:
                if "rejected" in str(e).lower() or "-2147418111" in str(e):
                    time.sleep(0.15) # Wait for AutoCAD to finish background task
                else:
                    raise e
        return None

    @staticmethod
    def _purge_gen_py():
        try:
            import win32com as _w
            p = os.path.join(os.path.dirname(_w.__file__),"gen_py")
            if os.path.exists(p): shutil.rmtree(p,ignore_errors=True)
        except Exception: pass

    def _log(self, msg: str):
        logger.info(msg); self._log_cb(msg)

    def _pt(self, x, y, z=0.0):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, (float(x),float(y),float(z)))
    def _arr(self, flat):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, [float(v) for v in flat])
    def _vec(self, x, y, z):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8, (float(x),float(y),float(z)))
    def _obj_arr(self, obj_list):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_DISPATCH, list(obj_list))

    def _mk_layer(self, name, color, weight, lt):
        try:
            ly = self.doc.Layers.Add(name)
            ly.Color,ly.Lineweight,ly.Linetype = color,weight,lt
        except Exception: pass

    def _lyr(self, obj, layer):
        try: obj.Layer = layer
        except Exception: pass

    def _rgb(self, obj, r, g, b):
        try:
            tc = obj.TrueColor
            tc.Red,tc.Green,tc.Blue = int(r),int(g),int(b)
            tc.ColorMethod = 0xC8
            obj.TrueColor = tc
        except Exception: pass

    def _mat_color(self, obj, mat: str):
        db = MATERIAL_DB.get(mat, MATERIAL_DB["Steel-4140"])
        self._rgb(obj, *db["color"])

    def _del(self, obj):
        try: obj.Delete()
        except Exception: pass

    # ══════════════════════════════════════════════════════════════════════════
    #  BOOLEAN OPERATIONS  (CORRECT CONSTANTS)
    # ══════════════════════════════════════════════════════════════════════════

    def _union(self, base, tool):
        if base is None or tool is None: return base
        try: self._do(base.Boolean, AC_UNION, tool)
        except Exception as e: self._log(f"      [!] UNION failed: {e}")
        return base

    def _subtract(self, base, tool):
        if base is None or tool is None: return base
        try: self._do(base.Boolean, AC_SUBTRACT, tool)
        except Exception as e: self._log(f"      [!] SUBTRACT failed: {e}")
        return base
    # ══════════════════════════════════════════════════════════════════════════
    #  UNIVERSAL CSG COMPILER (JSON RECIPES)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_from_recipe(self, cx, cy, p1, p2, p3, p4, recipe_steps):
        """Reads a list of JSON CSG steps and dynamically builds ANY part."""
        base_solid = None
        
        # Secure variable dictionary for evaluating Excel parameters
        variables = {"P1": float(p1), "P2": float(p2), "P3": float(p3), "P4": float(p4)}

        for step in recipe_steps:
            def val(key, default=0.0):
                if key not in step: return default
                try: return float(eval(str(step[key]), {"__builtins__": None}, variables))
                except Exception: return default

            action = str(step.get("action", "ADD")).upper()
            shape  = str(step.get("shape", "cylinder")).lower()
            
            x_off = val("x_offset")
            y_off = val("y_offset")
            z_pos = val("z")
            
            curr_x = cx + x_off
            curr_y = cy + y_off
            
            temp_solid = None
            if shape == "cylinder":
                r = val("radius"); h = val("height")
                temp_solid = self._cyl(curr_x, curr_y, z_pos, r, h, "GEAR_SOLID")
            elif shape == "box":
                l = val("length"); w = val("width"); h = val("height")
                temp_solid = self._box(curr_x - l/2, curr_y - w/2, z_pos, l, w, h, "GEAR_SOLID")
            elif shape == "sphere":
                r = val("radius")
                temp_solid = self._solid_sphere(curr_x, curr_y, r)

            if not temp_solid: continue

            if action == "BASE":
                base_solid = temp_solid
            elif action == "ADD" and base_solid:
                self._union(base_solid, temp_solid)
            elif action == "SUBTRACT" and base_solid:
                self._subtract(base_solid, temp_solid)
                
        self._log(f"      [Recipe] ✔ Custom part generated from CSG template")
        return base_solid
    # ══════════════════════════════════════════════════════════════════════════
    #  CORE SOLID-FROM-PROFILE PIPELINE
    # ══════════════════════════════════════════════════════════════════════════

    def _lwpl(self, coords2d: List[float], z_elev: float = 0.0):
        try:
            pl = self._do(self.ms.AddLightWeightPolyline, self._arr(coords2d))
            if pl:
                pl.Closed    = True
                pl.Layer     = "WORK_GEOM"
                pl.Elevation = float(z_elev)
            return pl
        except Exception as e:
            self._log(f"      [!] AddLWPL: {e}"); return None

    def _region(self, pl):
        try:
            ra = self._do(self.ms.AddRegion, self._obj_arr([pl]))
            if ra and len(ra) > 0: return ra[0]
            return None
        except Exception as e:
            self._log(f"      [!] AddRegion: {e}"); return None

    def _extrude(self, reg, h: float, taper: float = 0.0, layer: str = "GEAR_TEETH"):
        try:
            s = self._do(self.ms.AddExtrudedSolid, reg, float(h), float(taper))
            if s: self._lyr(s, layer)
            return s
        except Exception as e:
            self._log(f"      [!] AddExtrudedSolid: {e}"); return None

    def _profile_solid(self, coords2d: List[float], h: float, z: float = 0.0, taper: float = 0.0, layer: str = "GEAR_TEETH"):
        pl = self._lwpl(coords2d, z)
        if pl is None: return None
        reg = self._region(pl)
        self._del(pl)
        if reg is None: return None
        s = self._extrude(reg, h, taper, layer)
        self._del(reg)
        return s

    # ── Primitives ────────────────────────────────────────────────────────────

    def _cyl(self, cx, cy, z, r, h, layer="GEAR_SOLID"):
        if r <= 0 or h <= 0: return None
        try:
            s = self._do(self.ms.AddCylinder, self._pt(cx,cy,z), float(r), float(h))
            if s: self._lyr(s, layer)
            return s
        except Exception as e: return None

    def _box(self, x, y, z, L, W, H, layer="GEAR_SOLID"):
        try:
            s = self._do(self.ms.AddBox, self._pt(x,y,z), float(L),float(W),float(H))
            if s: self._lyr(s, layer)
            return s
        except Exception as e: return None

    def _annulus(self, cx, cy, z, r_out, r_in, h, layer="GEAR_SOLID"):
        outer = self._cyl(cx,cy,z,r_out,h,layer)
        if outer and r_in > 0.5 and r_in < r_out:
            inner = self._cyl(cx,cy,z-2.0,r_in,h+4.0,layer) # Deep pierce internal cut
            if inner: self._subtract(outer,inner)
        return outer

    # ══════════════════════════════════════════════════════════════════════════
    #  PER-TOOTH GEAR DISC BUILDER
    # ══════════════════════════════════════════════════════════════════════════

    def _build_gear_disc(self, cx: float, cy: float, Z: int, m: float, face_w: float, PA_deg: float = 20.0, x: float = 0.0, z0: float = 0.0, angle_offset: float = 0.0, N: int = 48, layer: str = "GEAR_TEETH") -> Optional[object]:
        pitch_r = Z * m / 2.0
        alpha   = math.radians(PA_deg)
        base_r  = pitch_r * math.cos(alpha)
        root_r  = max(pitch_r - m*(1.25-x), base_r*0.05, m*0.3)
        span    = 2.0 * math.pi / Z

        root_disc = self._cyl(cx,cy,z0,root_r,face_w,layer)
        if root_disc is None: return None

        ok = 0; fail = 0
        for i in range(Z):
            ang      = i * span + angle_offset
            flat_loc = single_tooth_flat(Z, m, ang, PA_deg, x, N)
            flat_wld = []
            for k in range(0, len(flat_loc), 2):
                flat_wld.append(flat_loc[k]   + cx)
                flat_wld.append(flat_loc[k+1] + cy)

            ts = self._profile_solid(flat_wld, face_w, z=z0, layer=layer)
            if ts is not None:
                self._union(root_disc, ts)
                ok += 1
            else:
                fail += 1

        self._log(f"      [Teeth] ✔ {ok}/{Z} involute  {fail} fail")
        return root_disc

    # ── Bore + DIN 6885-1 keyway ──────────────────────────────────────────────

    def _bore_kw(self, solid, cx, cy, z_bot, z_top, bore_d):
        if solid is None or bore_d < 1.0: return
        h  = (z_top - z_bot) + 20.0 # Extend by 20mm to prevent peg errors
        br = bore_d / 2.0
        bc = self._cyl(cx, cy, z_bot - 10.0, br, h, "GEAR_BORE")
        if bc: self._subtract(solid, bc)
        
        kw_b = bore_d/4.0;  kw_t = bore_d/8.0
        kw = self._box(cx-kw_b/2, cy+br-kw_t, z_bot - 10.0, kw_b, kw_t+br, h, "GEAR_BORE")
        if kw: self._subtract(solid, kw)
        self._log(f"      [Bore] ✔ Ø{bore_d:.1f}  kw {kw_b:.1f}×{kw_t:.1f}")

    # ── Blank features ────────────────────────────────────────────────────────

    def _blank_p(self, Z, m, bore_d, face_w):
        pitch_r = Z*m/2.0
        root_r  = max(pitch_r-1.25*m, m*0.3)
        br      = bore_d/2.0
        hub_r = min(max(br*1.7, br+m*1.3, br+5.0), root_r*0.54)
        hub_r = max(hub_r, br+2.0)
        boss_h = min(face_w*0.28, 12.0)
        web_t    = max(face_w*0.40, 5.0)
        recess_d = max((face_w - web_t)/2.0 - 0.5, 0.0)
        recess_r = root_r - m*0.65
        ann   = root_r - hub_r
        has_h = False;  n_h=0;  bc_r=hub_r;  hr=0.0
        if ann > 4.0*4.5:
            bc_r   = (hub_r + root_r*0.80)/2.0
            max_hr = min(bc_r-hub_r-2.5, root_r-bc_r-2.5, ann*0.28)
            hr     = max(4.0, max_hr)
            n_h    = 6 if Z<=36 else 8
            while n_h >= 4:
                if 2*math.pi*bc_r/n_h - 2*hr >= 3.0: break
                n_h -= 2
            has_h = (hr >= 4.0) and (n_h >= 4)
        return dict(hub_r=hub_r, boss_h=boss_h, recess_d=recess_d, recess_r=recess_r, has_h=has_h, n_h=n_h, bc_r=bc_r, hr=hr)

    def _blank(self, solid, cx, cy, Z, m, bore_d, face_w, z0=0.0):
        if solid is None: return
        p  = self._blank_p(Z, m, bore_d, face_w)
        fz = z0 + face_w

        if p["hub_r"] > bore_d/2.0 + 1.0 and p["boss_h"] > 0.5:
            boss = self._cyl(cx, cy, fz, p["hub_r"], p["boss_h"], "GEAR_BLANK")
            if boss: self._union(solid, boss)

        rd = p["recess_d"];  rr = p["recess_r"]
        if rd > 1.5 and rr > p["hub_r"] + 3.0:
            rc_front = self._annulus(cx,cy, fz-rd, rr, p["hub_r"]+0.5, rd+2.0, "GEAR_BLANK")
            if rc_front: self._subtract(solid, rc_front)

            rc_back  = self._annulus(cx,cy, z0-2.0, rr, p["hub_r"]+0.5, rd+2.0, "GEAR_BLANK")
            if rc_back:  self._subtract(solid, rc_back)

        if p["has_h"]:
            total_h = face_w + p["boss_h"] + 20.0
            z_start = z0 - 10.0
            for i in range(p["n_h"]):
                ang = i * 2*math.pi/p["n_h"] + math.pi/p["n_h"]
                hx  = cx + p["bc_r"]*math.cos(ang)
                hy  = cy + p["bc_r"]*math.sin(ang)
                hc  = self._cyl(hx, hy, z_start, p["hr"], total_h, "GEAR_BLANK")
                if hc: self._subtract(solid, hc)

    # ══════════════════════════════════════════════════════════════════════════
    #  GEAR 1 — SPUR GEAR
    # ══════════════════════════════════════════════════════════════════════════

    def _gear_spur(self, cx, cy, Z, m, face_w, bore_d, PA_deg=20.0):
        x       = profile_shift_x(Z, PA_deg)
        pitch_r = Z*m/2.0
        outer_r = pitch_r+m*(1.0+x)
        bore_r  = bore_d/2.0
        hub_r   = max(bore_r*1.7, bore_r+m*1.3)
        hub_h   = face_w*0.55

        N     = max(24, min(48, Z*2)) # Dynamic resolution
        solid = self._build_gear_disc(cx,cy,Z,m,face_w,PA_deg,x,z0=0.0,N=N)
        if solid is None:
            solid = self._cyl(cx,cy,0,outer_r,face_w,"GEAR_TEETH")

        if hub_r > bore_r+0.5:
            hub = self._cyl(cx,cy,-hub_h, hub_r, hub_h, "GEAR_BLANK")
            if hub: self._union(solid, hub)

        self._bore_kw(solid, cx, cy, -hub_h, face_w, bore_d)
        self._blank(solid, cx, cy, Z, m, bore_d, face_w, z0=0.0)
        return solid

    # ══════════════════════════════════════════════════════════════════════════
    #  GEAR 2 — HELICAL GEAR 
    # ══════════════════════════════════════════════════════════════════════════

    def _gear_helical(self, cx, cy, Z, m, face_w, bore_d, helix_deg=15.0, PA_deg=20.0, right_hand=True):
        x           = profile_shift_x(Z, PA_deg)
        pitch_r     = Z*m/2.0
        outer_r     = pitch_r+m*(1.0+x)
        bore_r      = bore_d/2.0
        hub_r       = max(bore_r*1.7, bore_r+m*1.3)
        hub_h       = face_w*0.55
        total_twist = (face_w/pitch_r)*math.tan(math.radians(helix_deg))
        sign        = 1.0 if right_hand else -1.0
        N_SL        = 20
        OVLP        = 0.12
        N           = max(24, min(48, Z*2))

        slice_h = face_w/N_SL;  ext_h = slice_h*(1+OVLP)
        base = None
        for sl in range(N_SL):
            ang_off = sign * sl * (total_twist/N_SL)
            sl_disc = self._build_gear_disc(cx,cy,Z,m,ext_h,PA_deg,x, z0=sl*slice_h, angle_offset=ang_off, N=N)
            if sl_disc is None:
                sl_disc = self._cyl(cx,cy,sl*slice_h,outer_r,ext_h,"GEAR_TEETH")
            if base is None: base = sl_disc
            elif sl_disc: self._union(base, sl_disc)

        if base is None:
            base = self._cyl(cx,cy,0,outer_r,face_w,"GEAR_TEETH")

        if hub_r > bore_r+0.5:
            hub = self._cyl(cx,cy,-hub_h,hub_r,hub_h,"GEAR_BLANK")
            if hub: self._union(base, hub)

        self._bore_kw(base,cx,cy,-hub_h,face_w,bore_d)
        self._blank(base,cx,cy,Z,m,bore_d,face_w,z0=0.0)
        return base

    # ══════════════════════════════════════════════════════════════════════════
    #  GEAR 3 — RING (INTERNAL) GEAR
    # ══════════════════════════════════════════════════════════════════════════

    def _gear_ring(self, cx, cy, Z, m, face_w, ring_thk, PA_deg=20.0):
        x       = profile_shift_x(Z, PA_deg)
        pitch_r = Z*m/2.0
        outer_r = pitch_r - m + ring_thk
        if outer_r <= pitch_r - m:
            outer_r = pitch_r + ring_thk*0.5

        disc = self._cyl(cx,cy,0,outer_r,face_w,"GEAR_TEETH")
        if disc is None: return None

        N    = max(24, min(48, Z*2))
        span = 2.0*math.pi/Z
        for i in range(Z):
            ang = i*span
            flat_loc = single_tooth_flat(Z,m,ang,PA_deg,x,N)
            flat_wld = []
            for k in range(0,len(flat_loc),2):
                flat_wld.append(flat_loc[k]+cx); flat_wld.append(flat_loc[k+1]+cy)
            vs = self._profile_solid(flat_wld, face_w + 4.0, z=-2.0, layer="WORK_GEOM") # Deep pierce void
            if vs:
                self._subtract(disc, vs)

        return disc

    # ══════════════════════════════════════════════════════════════════════════
    #  GEAR 4 — BEVEL GEAR
    # ══════════════════════════════════════════════════════════════════════════

    def _gear_bevel(self, cx, cy, Z, m, face_w, bore_d, cone_deg=45.0, PA_deg=20.0):
        cr      = math.radians(cone_deg)
        pitch_r = Z*m/2.0
        back_r  = pitch_r
        front_r = max(back_r - face_w*math.sin(cr), m*0.5)
        cone_h  = face_w*math.cos(cr)
        mean_r  = (back_r+front_r)/2.0
        m_v     = m*mean_r/back_r 
        x_v     = profile_shift_x(Z, PA_deg)
        bore_r  = bore_d/2.0
        hub_r   = max(bore_r*1.5, bore_r+m*0.8)
        hub_h   = cone_h*0.40

        try:
            frustum = self._do(self.ms.AddFrustum, self._pt(cx,cy,0),float(back_r),float(cone_h),float(front_r))
            self._lyr(frustum,"GEAR_TEETH")
        except Exception:
            try:
                frustum = self._do(self.ms.AddCone, self._pt(cx,cy,0),float(back_r),float(cone_h))
                self._lyr(frustum,"GEAR_TEETH")
            except Exception as e:
                return self._cyl(cx,cy,0,back_r,cone_h,"GEAR_TEETH")

        N=max(24,min(48,Z*2)); span=2*math.pi/Z
        for i in range(Z):
            ang = i*span
            tx  = cx+mean_r*math.cos(ang);  ty = cy+mean_r*math.sin(ang)
            flat_loc = single_tooth_flat(Z,m_v,0.0,PA_deg,x_v,N)
            flat_wld = []
            for k in range(0,len(flat_loc),2):
                flat_wld.append(flat_loc[k]+tx); flat_wld.append(flat_loc[k+1]+ty)
            th_h = cone_h*0.70;  th_z = cone_h*0.15
            ts = self._profile_solid(flat_wld,th_h,z=th_z,layer="GEAR_TEETH")
            if ts:
                self._union(frustum, ts)

        if hub_r>bore_r+0.5 and hub_r<back_r:
            hub=self._cyl(cx,cy,-hub_h,hub_r,hub_h,"GEAR_BLANK")
            if hub: self._union(frustum, hub)
        self._bore_kw(frustum,cx,cy,-hub_h,cone_h,bore_d)
        return frustum

    # ══════════════════════════════════════════════════════════════════════════
    #  GEAR 5 — WORM
    # ══════════════════════════════════════════════════════════════════════════

    def _gear_worm(self, cx, cy, n_starts, m, length, bore_d, lead_deg=15.0):
        bore_r   = bore_d/2.0
        shaft_r  = max(bore_r+m*1.5, m*2.5)
        thread_r = m*0.55
        lead     = n_starts*math.pi*m
        n_turns  = length/lead if lead>0 else 2
        fl_r     = shaft_r+m*0.45;  fl_h=m*0.80

        shaft = self._cyl(cx,cy,0,shaft_r,length,"GEAR_TEETH")
        if shaft is None: return None

        helix_r = shaft_r+thread_r*0.30

        for s in range(n_starts):
            ang0=s*2*math.pi/n_starts; success=False
            path_sp=None; prof_pl=None
            try:
                npts=120; pts3d=[]
                for j in range(npts+1):
                    frac=j/npts; ang=ang0+2*math.pi*n_turns*frac
                    pts3d+=[cx+helix_r*math.cos(ang), cy+helix_r*math.sin(ang), length*frac]
                tang0=[-helix_r*math.sin(ang0), helix_r*math.cos(ang0), lead/(2*math.pi)]
                ang_end=ang0+2*math.pi*n_turns
                tang1=[-helix_r*math.sin(ang_end), helix_r*math.cos(ang_end), lead/(2*math.pi)]
                path_sp = self._do(self.ms.AddSpline, self._arr(pts3d),self._vec(*tang0),self._vec(*tang1))
                self._lyr(path_sp,"WORK_GEOM")
                
                x0=cx+helix_r*math.cos(ang0); y0=cy+helix_r*math.sin(ang0)
                nc=20; cpts=[]
                for j in range(nc):
                    a=2*math.pi*j/nc; cpts+=[x0+thread_r*math.cos(a), y0+thread_r*math.sin(a)]
                prof_pl = self._do(self.ms.AddLightWeightPolyline, self._arr(cpts))
                prof_pl.Closed=True; prof_pl.Layer="WORK_GEOM"; prof_pl.Elevation=0.0
                
                ts = self._do(self.ms.AddExtrudedSolidAlongPath, prof_pl,path_sp)
                self._lyr(ts,"GEAR_TEETH")
                self._union(shaft, ts)
                self._del(prof_pl); self._del(path_sp)
                success=True
            except Exception as e:
                self._del(path_sp); self._del(prof_pl)
            
            if not success:
                nrid=max(8,int(length/lead*12))
                for j in range(nrid):
                    frac=j/nrid; ang=ang0+2*math.pi*n_turns*frac; z_j=length*frac
                    rx=cx+helix_r*math.cos(ang); ry=cy+helix_r*math.sin(ang)
                    rd=self._cyl(rx,ry,z_j-thread_r,thread_r,thread_r*2.1,"GEAR_TEETH")
                    if rd: self._union(shaft,rd)

        for z_fl,h_fl in [(-fl_h,fl_h),(length,fl_h)]:
            fl=self._cyl(cx,cy,z_fl,fl_r,h_fl,"GEAR_BLANK")
            if fl: self._union(shaft,fl)
        
        self._bore_kw(shaft, cx, cy, -fl_h, length+fl_h, bore_d)
        return shaft

    # ══════════════════════════════════════════════════════════════════════════
    #  GEAR 6 — WORM WHEEL
    # ══════════════════════════════════════════════════════════════════════════

    def _gear_worm_wheel(self, cx, cy, Z, m, face_w, bore_d):
        x        = profile_shift_x(Z)
        pitch_r  = Z*m/2.0
        outer_r  = pitch_r+m*(1.0+x)
        bore_r   = bore_d/2.0
        hub_r    = max(bore_r*1.7, bore_r+m*1.2)
        hub_h    = face_w*0.50
        worm_sr  = max(bore_r+m*1.5, m*2.5)
        cd       = pitch_r+worm_sr
        t_minor  = worm_sr+m

        N     = max(24, min(48, Z*2))
        solid = self._build_gear_disc(cx,cy,Z,m,face_w,20.0,x,z0=0.0,N=N)
        if solid is None:
            solid = self._cyl(cx,cy,0,outer_r,face_w,"GEAR_TEETH")

        try:
            tor = self._do(self.ms.AddTorus, self._pt(cx,cy+cd,face_w/2.0),float(cd),float(t_minor))
            self._lyr(tor,"WORK_GEOM")
            self._subtract(solid, tor) 
        except Exception as e: pass

        if hub_r>bore_r+0.5:
            hub=self._cyl(cx,cy,face_w,hub_r,hub_h,"GEAR_BLANK")
            if hub: self._union(solid, hub)
        
        self._bore_kw(solid,cx,cy,0,face_w+hub_h,bore_d)
        self._blank(solid,cx,cy,Z,m,bore_d,face_w,z0=0.0)
        return solid

    # ══════════════════════════════════════════════════════════════════════════
    #  NON-GEAR SOLIDS
    # ══════════════════════════════════════════════════════════════════════════

    def _solid_box(self,cx,cy,L,W,H):
        return self._box(cx-L/2,cy-W/2,0,L,W,H,"GEAR_SOLID")

    def _solid_sphere(self,cx,cy,r):
        try:
            s = self._do(self.ms.AddSphere, self._pt(cx,cy,r),float(r))
            self._lyr(s,"GEAR_SOLID"); return s
        except Exception as e: return None

    def _solid_cylinder(self,cx,cy,r_out,r_in,h):
        return self._annulus(cx,cy,0,r_out,r_in,h,"GEAR_SOLID")

    def _solid_cone(self,cx,cy,r,h):
        try:
            s = self._do(self.ms.AddCone, self._pt(cx,cy,0),float(r),float(h))
            self._lyr(s,"GEAR_SOLID"); return s
        except Exception:
            return self._cyl(cx,cy,0,r,h,"GEAR_SOLID")
    def _parametric_plate(self, cx, cy, L, W, H, hole_dia):
        # 1. Main Plate
        plate = self._box(cx - L/2, cy - W/2, 0, L, W, H, "GEAR_SOLID")
        if not plate: return None
        
        # 2. Hole Pattern (4 corners, offset by 15mm from edges)
        offset = 15.0
        hr = hole_dia / 2.0
        if L > 40 and W > 40 and hr > 0:
            for dx in [-1, 1]:
                for dy in [-1, 1]:
                    hx = cx + dx * (L/2 - offset)
                    hy = cy + dy * (W/2 - offset)
                    hole = self._cyl(hx, hy, -5.0, hr, H + 10.0, "GEAR_BORE")
                    if hole: self._subtract(plate, hole)
        
        self._log(f"      [Plate] ✔ {L}x{W}x{H} with 4x Ø{hole_dia} holes")
        return plate
    # ══════════════════════════════════════════════════════════════════════════
    #  ESSENTIAL MECHANICAL COMPONENTS (Hole Patterns & Assembly)
    # ══════════════════════════════════════════════════════════════════════════

    def _solid_flange(self, cx, cy, od, id_bore, thk, n_holes):
        """Parametric Flange with automatic Bolt Hole Circle (BHC) pattern."""
        flange = self._cyl(cx, cy, 0, od/2.0, thk, "GEAR_SOLID")
        if not flange: return None

        # 1. Center Bore
        if id_bore > 0:
            bore = self._cyl(cx, cy, -5.0, id_bore/2.0, thk+10.0, "GEAR_BORE")
            if bore: self._subtract(flange, bore)

        # 2. Bolt Hole Pattern (Polar Coordinates)
        n_holes = int(n_holes)
        if n_holes > 0:
            pcd = (od + id_bore) / 2.0  # Pitch Circle Diameter in the middle
            hole_r = min(10.0, (od - pcd) * 0.35) # Auto-scale hole size
            if hole_r > 1.0:
                span = 2 * math.pi / n_holes
                ok = 0
                for i in range(n_holes):
                    ang = i * span
                    hx = cx + (pcd/2.0) * math.cos(ang)
                    hy = cy + (pcd/2.0) * math.sin(ang)
                    hc = self._cyl(hx, hy, -5.0, hole_r, thk+10.0, "GEAR_BORE")
                    if hc: 
                        self._subtract(flange, hc)
                        ok += 1
                self._log(f"      [Flange] ✔ {ok}/{n_holes} holes drilled on Ø{pcd:.1f} BHC")

        return flange

    def _solid_stepped_shaft(self, cx, cy, d1, length, d2, l2):
        """Power transmission shaft with end-steps and a central keyway."""
        # 1. Main central shaft
        shaft = self._cyl(cx, cy, 0, d1/2.0, length, "GEAR_SOLID")
        if not shaft: return None

        # 2. Add stepped down ends (Z-axis stacking)
        if d2 > 0 and l2 > 0:
            step_top = self._cyl(cx, cy, length, d2/2.0, l2, "GEAR_SOLID")
            if step_top: self._union(shaft, step_top)
            
            step_bot = self._cyl(cx, cy, -l2, d2/2.0, l2, "GEAR_SOLID")
            if step_bot: self._union(shaft, step_bot)

        # 3. Mill a DIN keyway into the main body
        kw_w = d1/4.0; kw_d = d1/8.0
        kw = self._box(cx - kw_w/2, cy + d1/2.0 - kw_d, length*0.2, kw_w, d1/2.0, length*0.6, "GEAR_BORE")
        if kw: self._subtract(shaft, kw)
        
        self._log(f"      [Shaft] ✔ Stepped ends added, Keyway milled")
        return shaft

    def _solid_l_bracket(self, cx, cy, L, W, H, T):
        """Structural L-Bracket with multi-axis mounting holes."""
        # 1. Base Plate (XY Plane)
        base = self._box(cx - L/2, cy - W/2, 0, L, W, T, "GEAR_SOLID")
        if not base: return None

        # 2. Upright Wall (YZ Plane)
        upright = self._box(cx - L/2, cy - W/2, T, T, W, H - T, "GEAR_SOLID")
        if upright: self._union(base, upright)

        hole_r = max(T * 0.4, 3.0)
        offset = T + hole_r + 4.0

        # 3. Drill Base Holes (Vertical Z-axis punches)
        if L > offset*3 and W > offset*3:
            for dy in [1, -1]:
                hx = cx + L/2 - offset
                hy = cy + dy * (W/2 - offset)
                hc = self._cyl(hx, hy, -5.0, hole_r, T+10.0, "GEAR_BORE")
                if hc: self._subtract(base, hc)

            # 4. Drill Upright Holes (Horizontal X-axis punches via Rotate3D)
            for dy in [1, -1]:
                hx = cx - L/2 - 5.0
                hy = cy + dy * (W/2 - offset)
                hz = H - offset
                hc = self._cyl(hx, hy, hz, hole_r, T+10.0, "GEAR_BORE")
                if hc:
                    # Rotate the punch 90 degrees around the Y-axis to point horizontally
                    self._do(hc.Rotate3D, self._pt(hx, hy, hz), self._pt(hx, hy+1, hz), math.pi/2)
                    self._subtract(base, hc)
        
        self._log(f"      [Bracket] ✔ {L}x{W}x{H} with multi-axis mounting holes")
        return base
    # ══════════════════════════════════════════════════════════════════════════
    #  ERP & BATCH RUNNER
    # ══════════════════════════════════════════════════════════════════════════

    def _erp(self,ptype,p1,p2,p3,p4,mat):
        pi=math.pi; Z,m,fw,bd=p1,p2,p3,p4; v=0.0
        if ptype in("Spur_Gear_3D","Helical_Gear","Worm_Wheel"):
            v=pi*((Z*m/2+m)**2-(bd/2)**2)*fw*0.88
        elif ptype=="Ring_Gear_3D":
            inn=Z*m/2-m; v=pi*((inn+bd)**2-inn**2)*fw
        elif ptype=="Bevel_Gear":
            cr=math.radians(45); br=Z*m/2
            fr=max(br-fw*math.sin(cr),0); h=fw*math.cos(cr)
            v=(pi/3)*h*(br**2+br*fr+fr**2)*0.82
        elif ptype=="Worm":
            sr=bd/2+m*1.5 if bd>1 else m*2.5; v=pi*sr**2*fw*0.78
        elif ptype=="Box": v=p1*p2*p3
        elif ptype=="Cylinder": v=pi*(p1**2-p2**2)*p3
        elif ptype=="Sphere": v=4/3*pi*p1**3
        elif ptype=="Flange":
            v = pi * ((p1/2)**2 - (p2/2)**2) * p3 * 0.90 # 10% volume reduction for holes
        elif ptype=="Stepped_Shaft":
            v = pi * (p1/2)**2 * p3 + 2 * pi * (p2/2)**2 * p4
        elif ptype=="L_Bracket":
            v = (p1 * p2 * p4) + (p4 * p2 * (p3 - p4))
        else: v=pi*p1**2*p3
        db=MATERIAL_DB.get(mat,MATERIAL_DB["Steel-4140"])
        mass=round(max(v,0)*db["density"]/1e6,3)
        cost=round(mass*db["cost_per_kg"],2)
        return max(v,0),mass,cost

    def _make_layout(self,part,mass,cost,vol):
        pno=str(part.get("Part_Number","PART")); name=f"DRW_{pno}"[:31]
        try:    layout=self.doc.Layouts.Add(name)
        except Exception:
            try:    layout=self.doc.Layouts.Item(name)
            except Exception: return ""
        try:
            self.doc.ActiveLayout=layout; ps=self.doc.PaperSpace
        except Exception as e:
            return name

        def pl(x,y,w,h,lyr="TITLE_BLOCK"):
            try:
                o=ps.AddLightWeightPolyline(self._arr([x,y,x+w,y,x+w,y+h,x,y+h]))
                o.Closed=True; o.Layer=lyr
            except Exception: pass
        def tx(t,x,y,h=4.0,lyr="TITLE_BLOCK"):
            try: o=ps.AddText(str(t),self._pt(x,y),float(h)); o.Layer=lyr
            except Exception: pass
        def ln(x1,y1,x2,y2,lyr="TITLE_BLOCK"):
            try: o=ps.AddLine(self._pt(x1,y1),self._pt(x2,y2)); o.Layer=lyr
            except Exception: pass

        M=10.0; pl(0,0,420,297); pl(M,M+62,420-2*M,297-2*M-62)
        for vn,cfg in VIEWPORTS.items():
            try:
                ox,oy,w,h=cfg["ox"],cfg["oy"],cfg["w"],cfg["h"]
                vp=ps.AddViewport(self._pt(ox+w/2,oy+h/2),float(w),float(h))
                try: vp.Layer="VIEW_BORDER"
                except Exception: pass
                try:
                    vp.Direction=self._vec(*cfg["eye"]); vp.UpVector=self._vec(*cfg["up"])
                except Exception: pass
                try: vp.StandardScale=2
                except Exception: pass
                pl(ox,oy,w,h,"VIEW_BORDER"); tx(vn,ox+2,oy-7,3.5,"VIEW_BORDER")
            except Exception: pass

        TB=62.0; TY=M; TW=420-2*M; c1=M+TW*0.40; c2=M+TW*0.70
        pl(M,TY,TW,TB)
        for y_ in [TY+TB*2/3,TY+TB/3]: ln(M,y_,M+TW,y_)
        ln(c1,TY+TB*2/3,c1,TY+TB); ln(c2,TY+TB/3,c2,TY+TB)
        dh=TB/3/3.5; r3=TY+TB*2/3+3
        tx("SIRAAL MANUFACTURING SYSTEMS  |  TN-IMPACT 2026",M+5,TY+TB*2/3+9,6.5)
        tx(f"REV:A  |  3D INVOLUTE GEAR  |  PA=20°",c2+4,TY+TB*2/3+5,4.0)
        tx(f"PART NO : {pno}",M+5,r3+dh*2,4.5)
        tx(f"TYPE    : {part.get('Part_Type','')}",M+5,r3+dh,4.5)
        tx(f"MATL    : {part.get('Material','')}",M+5,r3,4.5)
        tx(f"MASS  : {mass} kg",c1+5,r3+dh*2,4.5)
        tx(f"VOL   : {vol/1000:.1f} cm³",c1+5,r3+dh,4.5)
        tx(f"COST  : Rs.{cost:,.0f}",c1+5,r3,4.5)
        r2=TY+TB/3+3
        tx(f"Z={part.get('Param_1','')} m={part.get('Param_2','')} "
           f"FW={part.get('Param_3','')} BD={part.get('Param_4','')}",M+5,r2+dh*2,4.0)
        tx(f"QTY={part.get('Quantity',1)} PRI={part.get('Priority','')}",M+5,r2+dh,4.0)
        tx("ISO 53/DIN 867/DIN 3992  1st ANGLE  mm",M+5,r2,4.0)
        tx("IS 2535/ISO 1328/DIN 3961  SCALE 1:1",c1+5,r2+dh*2,4.0)
        tx(f"TOLERANCE: ±0.05 mm (Unless specified)", c1+5, r2-dh, 3.5)
        tx(f"SIRAAL ENGINE v6.2  |  Session Folders",c1+5,r2+dh,4.0)
        try: self.doc.ActiveLayout=self.doc.Layouts.Item("Model")
        except Exception: pass
        return name

    # ══════════════════════════════════════════════════════════════════════════
    #  DXF EXPORT (CLONE & SWEEP METHOD)
    # ══════════════════════════════════════════════════════════════════════════

    def _export_dxf(self, pno, solid, orig_cx, orig_cy, out_dir):
        """
        Clones the solid from the Master Document into a new DXF file. 
        Bypasses COM tuple bugs by simply sweeping the new ModelSpace 
        and moving everything to the origin.
        """
        if solid is None:
            return
            
        os.makedirs(out_dir, exist_ok=True)
        orig_doc = self.doc
        orig_ms = self.ms
        tmp_doc = None
        
        try:
            time.sleep(0.2)
            tmp_doc = self.acad.Documents.Add()
            tmp_ms = win32com.client.dynamic.Dispatch(tmp_doc.ModelSpace)
            
            # 1. Copy the solid to the new document
            obj_array = self._obj_arr([solid])
            self._do(orig_doc.CopyObjects, obj_array, tmp_ms)
            
            # 2. THE FIX: Sweep the new ModelSpace
            # Since it's a fresh file, the ONLY thing inside is our cloned gear!
            pt_current = self._pt(orig_cx, orig_cy, 0.0)
            pt_origin  = self._pt(0.0, 0.0, 0.0)
            
            for obj in tmp_ms:
                self._do(obj.Move, pt_current, pt_origin)
            
            # 3. Save as DXF
            path = os.path.abspath(os.path.join(out_dir, f"{pno}_3D.dxf"))
            if os.path.exists(path): os.remove(path)
            
            time.sleep(0.2)
            self._do(tmp_doc.SaveAs, path, DXF_FORMAT)
            self._log(f"          [DXF] ✔ {pno}_3D.dxf (Cloned & Centered)")
            
        except Exception as e: 
            self._log(f"          [DXF] {pno}: {e}")
        finally:
            self.doc = orig_doc
            self.ms = orig_ms
            try: 
                if tmp_doc:
                    self._do(tmp_doc.Close, False)
                    time.sleep(0.5)
            except Exception: pass
    def _dispatch(self,cx,cy,ptype,p1,p2,p3,p4):
        Z,m,fw,bd=p1,p2,p3,p4
        if   ptype=="Spur_Gear_3D":   return self._gear_spur(cx,cy,int(Z),m,fw,bd)
        elif ptype=="Helical_Gear":    return self._gear_helical(cx,cy,int(Z),m,fw,bd)
        elif ptype=="Ring_Gear_3D":    return self._gear_ring(cx,cy,int(Z),m,fw,bd)
        elif ptype=="Bevel_Gear":      return self._gear_bevel(cx,cy,int(Z),m,fw,bd)
        elif ptype=="Worm":            return self._gear_worm(cx,cy,int(Z),m,fw,bd)
        elif ptype=="Worm_Wheel":      return self._gear_worm_wheel(cx,cy,int(Z),m,fw,bd)
        elif ptype=="Box":             return self._solid_box(cx,cy,Z,m,fw)
        elif ptype=="Cylinder":        return self._solid_cylinder(cx,cy,Z,m,fw)
        elif ptype=="Sphere":          return self._solid_sphere(cx,cy,Z)
        elif ptype=="Cone":            return self._solid_cone(cx,cy,Z,fw)
        elif ptype=="Mounting_Plate":  return self._parametric_plate(cx,cy,Z,m,fw,bd)
        elif ptype=="Flange":          return self._solid_flange(cx,cy,p1,p2,p3,p4)
        elif ptype=="Stepped_Shaft":   return self._solid_stepped_shaft(cx,cy,p1,p2,p3,p4)
        elif ptype=="L_Bracket":       return self._solid_l_bracket(cx,cy,p1,p2,p3,p4)
        elif ptype in("Flanged_Boss","Extruded_Profile","Revolved_Part"):
            return self._solid_cylinder(cx,cy,max(Z,m),0,fw)
        # 2. Check for Universal JSON CSG Recipes
        template_path = os.path.abspath(os.path.join(os.getcwd(), "templates", f"{ptype}.json"))
        if os.path.exists(template_path):
            import json
            try:
                with open(template_path, 'r') as f:
                    recipe = json.load(f)
                return self._build_from_recipe(cx, cy, Z, m, fw, bd, recipe.get("Steps", []))
            except Exception as e:
                self._log(f"      [!] Recipe Error ({ptype}.json): {e}")
                return None
                
        return None
    def _footprint(self,ptype,p1,p2,p3,p4):
        Z,m,fw,bd=p1,p2,p3,p4
        if ptype in("Spur_Gear_3D","Helical_Gear","Worm_Wheel"):
            x=profile_shift_x(int(Z))
            return (Z*m/2+m*(1+x))*2, fw
        elif ptype=="Ring_Gear_3D":
            return (Z*m/2-m+bd)*2, fw
        elif ptype=="Bevel_Gear":
            return (Z*m/2+m)*2, fw
        elif ptype=="Worm":
            sr=bd/2+m*1.5 if bd>1 else m*2.5; return (sr+m*0.5)*2, fw
        elif ptype=="Box": return Z, fw
        elif ptype=="Sphere": return Z*2, Z*2
        elif ptype=="Flange":          return p1 + 20, p3
        elif ptype=="Stepped_Shaft":   return max(p1, p3) + 20, p2 + 2*p4
        elif ptype=="L_Bracket":       return p1 + 20, max(p2, p3)
        else: return max(Z,m)*2+20, fw+20

    def generate_3d_batch(self,parts:List[dict]):
        GAP=90.0; cx=0.0; layouts=[]; failed=[]
        
        # --- NEW SESSION FOLDER LOGIC ---
        session_id = time.strftime("%Y%m%d_%H%M%S")
        session_dir = os.path.abspath(os.path.join(os.getcwd(), "output_3d", f"Session_{session_id}"))
        os.makedirs(session_dir, exist_ok=True)
        
        SEP="═"*62
        self._log(f"\n{SEP}")
        self._log(f"  SIRAAL ENGINE v6.2 — BATCH: {len(parts)} parts")
        self._log(f"  SESSION DIR: {session_dir}")
        self._log(f"{SEP}\n")

        for idx,part in enumerate(parts,1):
            pno  =str(part.get("Part_Number", f"P{idx:03d}")).strip()
            ptype=str(part.get("Part_Type",   "Spur_Gear_3D")).strip()
            mat  =str(part.get("Material",    "Steel-4140")).strip()
            
            # --- NEW PART FOLDER LOGIC ---
            part_dir = os.path.join(session_dir, pno)
            os.makedirs(part_dir, exist_ok=True)

            try:
                p1=float(part.get("Param_1",20)); p2=float(part.get("Param_2",3))
                p3=float(part.get("Param_3",30)); p4=float(part.get("Param_4",20))
            except Exception:
                self._log(f"  [!] {pno}: bad params — skip"); failed.append(pno); continue

            self._log(f"[{idx:02d}/{len(parts)}] ▶ {pno}  {ptype}  {mat}")
            try:
                vol,mass,cost=self._erp(ptype,p1,p2,p3,p4,mat)
                
                # --- SAVE ORIGINAL X COORDINATE ---
                orig_cx = cx
                
                solid=self._dispatch(orig_cx, 0.0, ptype, p1, p2, p3, p4)
                if solid is None:
                    self._log(f"          [!] no solid"); failed.append(pno); continue
                self._mat_color(solid,mat)
                
                # Advance the cursor for the next gear in the Master Assembly
                fw,_=self._footprint(ptype,p1,p2,p3,p4); cx+=fw+GAP
                
                try:
                    ln=self._make_layout(part,mass,cost,vol)
                    if ln: layouts.append(ln); self._log(f"          [Layout] ✔ {ln}")
                except Exception as e: self._log(f"          [Layout] {e}")
                
                # --- NEW EXPORT CALL: CLONING ---
                try: self._export_dxf(pno, solid, orig_cx, 0.0, part_dir)
                except Exception as e: self._log(f"          [DXF] {e}")
                
                self._log(f"          ✔ {pno} DONE")
                time.sleep(1.0)
            except Exception as e:
                import traceback
                self._log(f"  [✘] {pno}: {e}\n{traceback.format_exc()}")
                failed.append(pno)

        try: self.acad.ZoomExtents()
        except Exception: pass
        for v,val in [("ISOLINES",32),("FACETRES",5),("SHADEMODE",2)]:
            try: self.doc.SetVariable(v,val)
            except Exception: pass

        # --- SAVE MASTER TO SESSION FOLDER ---
        master = os.path.join(session_dir, "Master_3D_Assembly.dwg")
        for fmt in DWG_SAVE_FORMATS:
            try: self.doc.SaveAs(master,fmt); self._log(f"\n[+] Master DWG saved: {master}"); break
            except Exception: continue

        self._log(f"\n{SEP}")
        self._log(f"  DONE  built={len(parts)-len(failed)}/{len(parts)}  "
                  f"layouts={len(layouts)}")
        if failed: self._log(f"  failed={failed}")
        self._log(f"{SEP}\n")

# ══════════════════════════════════════════════════════════════════════════════
#  Excel BOM loader
# ══════════════════════════════════════════════════════════════════════════════

def load_bom_from_excel(xlsx_path:str)->List[dict]:
    try: import pandas as pd
    except ImportError: raise ImportError("pip install pandas openpyxl")
    df=pd.read_excel(xlsx_path,sheet_name="BOM_Gears",header=2,dtype=str)
    df.columns=[str(c).strip() for c in df.columns]
    def g(row,keys,default=""):
        ks=[keys] if isinstance(keys,str) else keys
        for k in ks:
            v=row.get(k)
            if v is not None and str(v).strip() not in("","nan"): return str(v).strip()
        return default
    parts=[]
    for _,row in df.iterrows():
        en=g(row,"Enabled","YES").upper()
        if en not in("YES","Y","1","TRUE"): continue
        pno=g(row,"Part_Number")
        if not pno or pno.lower() in("nan","part_number"): continue
        parts.append({
            "Part_Number": pno,
            "Part_Type":   g(row,"Part_Type","Spur_Gear_3D"),
            "Material":    g(row,"Material","Steel-4140"),
            "Param_1":     g(row,["Param_1\n(Z / Starts)","Param_1"],"20"),
            "Param_2":     g(row,["Param_2\n(Module m)","Param_2"],"3"),
            "Param_3":     g(row,["Param_3\n(Face Width)","Param_3"],"30"),
            "Param_4":     g(row,["Param_4\n(Bore Dia)","Param_4"],"20"),
            "Quantity":    g(row,["Qty","Quantity"],"1"),
            "Priority":    g(row,"Priority","High"),
            "Description": g(row,"Description"),
        })
    return parts

if __name__=="__main__":
    import sys
    logging.basicConfig(level=logging.INFO,format="%(asctime)s  %(message)s",datefmt="%H:%M:%S")
    if len(sys.argv)>1:
        parts=load_bom_from_excel(sys.argv[1])
        print(f"Loaded {len(parts)} parts from {sys.argv[1]}")
    else:
        parts=[
            {"Part_Number":"GR-001-SPUR",   "Part_Type":"Spur_Gear_3D",
             "Material":"Steel-4140","Param_1":"24","Param_2":"3","Param_3":"30","Param_4":"20","Quantity":"1","Priority":"High"},
            {"Part_Number":"GR-002-HELICAL","Part_Type":"Helical_Gear",
             "Material":"Steel-4140","Param_1":"30","Param_2":"3","Param_3":"40","Param_4":"22","Quantity":"1","Priority":"High"},
        ]
    AutoCAD3DGearEngine().generate_3d_batch(parts)