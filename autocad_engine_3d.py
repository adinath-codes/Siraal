"""
autocad_engine_3d.py  —  Siraal 3D Gear Engine  v7.0
=====================================================
v7.0 UPGRADES over v6.2:
─────────────────────────
  UPGRADE 1: Advanced Recipe Compiler
    - Shapes: cylinder, box, sphere, cone, torus, frustum,
              extrude_profile (2D point list), revolve (2D cross-section)
    - Actions: BASE, ADD, SUBTRACT, FILLET, CHAMFER,
               PATTERN_CIRCULAR, PATTERN_LINEAR, MIRROR
    - Rotation: every step supports rotate_axis + rotate_deg
    - Eval namespace fixed: math, pi, abs available in expressions
    - Per-step error logging (no more silent failures)

  UPGRADE 2: Tolerant Template Lookup in _dispatch()
    - Tries ptype.json, Custom_ptype.json, and case-insensitive match

  All v6.2 features preserved 100%.
"""

import win32com.client
import win32com.client.dynamic
import pythoncom
import math
import os
import time
import json
import logging
import shutil
from typing import Callable, Dict, List, Optional, Tuple
logger = logging.getLogger("Siraal.GearEngine")

# ── AutoCAD COM Boolean constants ─────────────────────────────────────────────
AC_UNION     = 0
AC_INTERSECT = 1
AC_SUBTRACT  = 2

DWG_SAVE_FORMATS = [67, 64, 61, 60, 48]
DXF_FORMAT       = 12

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

# Safe eval namespace — every expression the AI writes can use these
_EVAL_NS = {
    "__builtins__": None,
    "math":    math,
    "pi":      math.pi,
    "abs":     abs,
    "min":     min,
    "max":     max,
    "sqrt":    math.sqrt,
    "sin":     math.sin,
    "cos":     math.cos,
    "tan":     math.tan,
    "asin":    math.asin,
    "acos":    math.acos,
    "atan":    math.atan,
    "atan2":   math.atan2,
    "hypot":   math.hypot,
    "ceil":    math.ceil,
    "floor":   math.floor,
    "log":     math.log,
    "log10":   math.log10,
    "exp":     math.exp,
    "radians": math.radians,
    "degrees": math.degrees,
    "pow":     pow,
    "round":   round,
}


# ══════════════════════════════════════════════════════════════════════════════
#  INVOLUTE MATHEMATICS
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
        r = math.hypot(ix, iy); a = math.atan2(iy, ix) + r_off
        return r*math.cos(a), r*math.sin(a)

    def lpt(t):
        ix, iy = _inv_pt(base_r, t)
        r = math.hypot(ix, iy); a = math.atan2(-iy, ix) - r_off
        return r*math.cos(a), r*math.sin(a)

    pts: List[Tuple[float,float]] = []
    for k in range(N + 1):
        t = t_min + (t_max - t_min) * k / N
        pts.append(rpt(t))

    rt = rpt(t_max); lt = lpt(t_max)
    a_rt = math.atan2(rt[1], rt[0]); a_lt = math.atan2(lt[1], lt[0])
    da = a_lt - a_rt
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
    lf = lpt(t_min); rf = rpt(t_min)
    a_lf = math.atan2(lf[1], lf[0]); a_rf = math.atan2(rf[1], rf[0])
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
        self._log("║  SIRAAL GEAR ENGINE v7.0 — AI SHAPE COMPILER       ║")
        self._log("║  Advanced Recipe: Revolve · Sweep · Pattern · Fillet║")
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

    # ── Utilities ────────────────────────────────────────────────────────────

    def _do(self, func, *args):
        for _ in range(30):
            try: return func(*args)
            except Exception as e:
                if "rejected" in str(e).lower() or "-2147418111" in str(e):
                    time.sleep(0.15)
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
        return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,
                                       (float(x),float(y),float(z)))
    def _arr(self, flat):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,
                                       [float(v) for v in flat])
    def _vec(self, x, y, z):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,
                                       (float(x),float(y),float(z)))
    def _obj_arr(self, obj_list):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_DISPATCH,
                                       list(obj_list))

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

    # ── Boolean ops ──────────────────────────────────────────────────────────

    def _union(self, base, tool):
        if base is None or tool is None: return base
        try: self._do(base.Boolean, AC_UNION, tool)
        except Exception as e: self._log(f"      [!] UNION: {e}")
        return base

    def _subtract(self, base, tool):
        if base is None or tool is None: return base
        try: self._do(base.Boolean, AC_SUBTRACT, tool)
        except Exception as e: self._log(f"      [!] SUBTRACT: {e}")
        return base

    # ── Profile pipeline ─────────────────────────────────────────────────────

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

    def _profile_solid(self, coords2d: List[float], h: float, z: float = 0.0,
                        taper: float = 0.0, layer: str = "GEAR_TEETH"):
        pl = self._lwpl(coords2d, z)
        if pl is None: return None
        reg = self._region(pl)
        self._del(pl)
        if reg is None: return None
        s = self._extrude(reg, h, taper, layer)
        self._del(reg)
        return s

    # ── Revolve helper ────────────────────────────────────────────────────────

    def _revolve_profile(self, coords2d: List[float], z_elev: float,
                          axis_pt: Tuple, axis_dir: Tuple,
                          degrees: float = 360.0,
                          layer: str = "GEAR_SOLID") -> Optional[object]:
        """
        Revolve a closed 2D polyline around an axis.
        coords2d  — flat [x,y, x,y, ...] list of world coordinates
        axis_pt   — (x,y,z) point on the rotation axis
        axis_dir  — (dx,dy,dz) direction vector of the axis
        degrees   — arc of revolution
        """
        pl = self._lwpl(coords2d, z_elev)
        if pl is None: return None
        reg = self._region(pl)
        self._del(pl)
        if reg is None: return None
        try:
            s = self._do(self.ms.AddRevolvedSolid,
                         reg,
                         self._pt(*axis_pt),
                         self._vec(*axis_dir),
                         math.radians(degrees))
            if s: self._lyr(s, layer)
            self._del(reg)
            return s
        except Exception as e:
            self._log(f"      [!] AddRevolvedSolid: {e}")
            self._del(reg)
            return None

    # ── Primitives ────────────────────────────────────────────────────────────

    def _cyl(self, cx, cy, z, r, h, layer="GEAR_SOLID"):
        if r <= 0 or h <= 0: return None
        try:
            s = self._do(self.ms.AddCylinder, self._pt(cx,cy,z), float(r), float(h))
            if s: self._lyr(s, layer)
            return s
        except Exception: return None

    def _box(self, x, y, z, L, W, H, layer="GEAR_SOLID"):
        try:
            s = self._do(self.ms.AddBox, self._pt(x,y,z), float(L),float(W),float(H))
            if s: self._lyr(s, layer)
            return s
        except Exception: return None

    def _annulus(self, cx, cy, z, r_out, r_in, h, layer="GEAR_SOLID"):
        outer = self._cyl(cx,cy,z,r_out,h,layer)
        if outer and r_in > 0.5 and r_in < r_out:
            inner = self._cyl(cx,cy,z-2.0,r_in,h+4.0,layer)
            if inner: self._subtract(outer,inner)
        return outer

    def _solid_box(self,cx,cy,L,W,H):
        return self._box(cx-L/2,cy-W/2,0,L,W,H,"GEAR_SOLID")

    def _solid_sphere(self,cx,cy,r):
        try:
            s = self._do(self.ms.AddSphere, self._pt(cx,cy,r),float(r))
            self._lyr(s,"GEAR_SOLID"); return s
        except Exception: return None

    def _solid_cylinder(self,cx,cy,r_out,r_in,h):
        return self._annulus(cx,cy,0,r_out,r_in,h,"GEAR_SOLID")

    def _solid_cone(self,cx,cy,r,h):
        try:
            s = self._do(self.ms.AddCone, self._pt(cx,cy,0),float(r),float(h))
            self._lyr(s,"GEAR_SOLID"); return s
        except Exception:
            return self._cyl(cx,cy,0,r,h,"GEAR_SOLID")

    # ══════════════════════════════════════════════════════════════════════════
    #  UNIVERSAL CSG + ADVANCED RECIPE COMPILER  (v7.0)
    # ══════════════════════════════════════════════════════════════════════════

    def _eval_expr(self, expr, variables: dict, context: str = ""):
        """Safe expression evaluator. variables dict may contain P1-P4 + any V1-V9 SET_VARs."""
        ns = dict(_EVAL_NS)
        ns.update(variables)          # P1-P4 + V1-V9 all live here
        try:
            return float(eval(str(expr), ns))
        except Exception as e:
            if context:
                self._log(f"      [!] eval '{expr}' ({context}): {e}")
            return 0.0

    def _eval_pts(self, raw_pts: list, cx: float, cy: float,
                  variables: dict) -> List[float]:
        """
        Convert a raw list of alternating x,y expressions into world-space
        flat coordinates.  Expressions may reference P1-P4 and math.
        """
        ns = dict(_EVAL_NS); ns.update(variables)
        wld: List[float] = []
        for k in range(0, len(raw_pts) - 1, 2):
            try:
                px = float(eval(str(raw_pts[k]),   ns))
                py = float(eval(str(raw_pts[k+1]), ns))
            except Exception as e:
                self._log(f"      [!] point eval [{k}]: {e}")
                px, py = 0.0, 0.0
            wld.append(cx + px)
            wld.append(cy + py)
        return wld

    def _apply_rotation(self, solid, cx: float, cy: float, z: float,
                         rot_axis: str, rot_deg: float):
        """Rotate `solid` around an axis through (cx,cy,z)."""
        if solid is None or rot_deg == 0.0 or not rot_axis:
            return
        axis_ends = {
            "X": (self._pt(cx,   cy,   z), self._pt(cx+1, cy,   z)),
            "Y": (self._pt(cx,   cy,   z), self._pt(cx,   cy+1, z)),
            "Z": (self._pt(cx,   cy,   z), self._pt(cx,   cy,   z+1)),
        }
        pts = axis_ends.get(rot_axis.upper())
        if pts:
            try:
                self._do(solid.Rotate3D, pts[0], pts[1], math.radians(rot_deg))
            except Exception as e:
                self._log(f"      [!] Rotate3D ({rot_axis} {rot_deg}°): {e}")

    def _build_from_recipe(self, cx: float, cy: float,
                            p1, p2, p3, p4,
                            recipe_steps: list) -> Optional[object]:
        """
        Single-pass recipe compiler.  Every step executes immediately in
        sequence — no deferred Pass 2.

        KEY FIXES over previous version
        ────────────────────────────────
        FIX 1 — PATTERN_CIRCULAR / PATTERN_LINEAR
          OLD (broken): Copied the entire base_solid N times, growing
                        complexity exponentially. 36 slots = ~149 minutes.
          NEW (correct): Reads the PREVIOUS geometry step, builds N-1
                         additional copies of just that tool, unions them
                         into a single combined_cutter, then applies ONE
                         Boolean to base_solid.  O(n) not O(2^n).

        FIX 2 — FILLET / CHAMFER
          OLD (broken): SendCommand fires asynchronously with no selection,
                        AutoCAD waits for user input → hangs forever.
          NEW (correct): Logged as a note only.  These operations require
                         interactive edge-selection in AutoCAD COM and cannot
                         be automated reliably.  The geometry is otherwise
                         complete and correct.

        FIX 3 — Single pass
          OLD: Two-pass system where patterns ran after all geometry was
               already applied to base_solid → patterned wrong object.
          NEW: Steps execute in document order.  PATTERN immediately follows
               the step whose tool it should replicate.

        Shapes:  cylinder, box, sphere, cone, torus, frustum,
                 extrude_profile, revolve
        Actions: BASE, ADD, SUBTRACT,
                 PATTERN_CIRCULAR, PATTERN_LINEAR, MIRROR,
                 FILLET (note only), CHAMFER (note only)
        """
        base_solid  = None
        last_tool   = None   # the solid created by the most recent geometry step
        last_action = None   # "ADD" or "SUBTRACT" for that tool
        last_step   = None   # the step dict (for re-evaluating coords in pattern)

        variables = {
            "P1": float(p1), "P2": float(p2),
            "P3": float(p3), "P4": float(p4),
        }

        # ── helpers ───────────────────────────────────────────────────────────

        def _make_solid(step, step_i) -> Optional[object]:
            """Build one primitive solid from a step dict. Returns the solid."""
            shape    = str(step.get("shape", "cylinder")).lower()
            x_off    = self._eval_expr(step.get("x_offset", 0), variables, f"s{step_i} x_off")
            y_off    = self._eval_expr(step.get("y_offset", 0), variables, f"s{step_i} y_off")
            z_pos    = self._eval_expr(step.get("z",        0), variables, f"s{step_i} z")
            curr_x   = cx + x_off
            curr_y   = cy + y_off
            ctx      = f"step {step_i} {shape}"
            s        = None

            if shape == "cylinder":
                r = self._eval_expr(step.get("radius", "P1/2"), variables, ctx)
                h = self._eval_expr(step.get("height", "P3"),   variables, ctx)
                # overlap: extend shape into existing solid for guaranteed Boolean fusion
                ovlp = self._eval_expr(step.get("overlap", 0), variables, ctx)
                s = self._cyl(curr_x, curr_y, z_pos - ovlp, r, h + ovlp, "GEAR_SOLID")

            elif shape == "box":
                l = self._eval_expr(step.get("length", "P1"), variables, ctx)
                w = self._eval_expr(step.get("width",  "P2"), variables, ctx)
                h = self._eval_expr(step.get("height", "P3"), variables, ctx)
                ovlp = self._eval_expr(step.get("overlap", 0), variables, ctx)
                # origin="corner"  → box starts at (curr_x, curr_y)   — matches extrude_profile
                # origin="centre"  → box centred on (curr_x, curr_y)  — default legacy
                origin = str(step.get("origin", "centre")).lower()
                if origin == "corner":
                    s = self._box(curr_x,       curr_y - w/2, z_pos - ovlp,
                                   l, w, h + ovlp, "GEAR_SOLID")
                else:
                    s = self._box(curr_x - l/2, curr_y - w/2, z_pos - ovlp,
                                   l, w, h + ovlp, "GEAR_SOLID")

            elif shape == "sphere":
                r = self._eval_expr(step.get("radius", "P1/2"), variables, ctx)
                try:
                    s = self._do(self.ms.AddSphere,
                                  self._pt(curr_x, curr_y, z_pos + r), float(r))
                    if s: self._lyr(s, "GEAR_SOLID")
                except Exception as e:
                    self._log(f"      [!] Sphere {ctx}: {e}")

            elif shape == "cone":
                r = self._eval_expr(step.get("radius", "P1/2"), variables, ctx)
                h = self._eval_expr(step.get("height", "P3"),   variables, ctx)
                try:
                    s = self._do(self.ms.AddCone,
                                  self._pt(curr_x, curr_y, z_pos), float(r), float(h))
                    if s: self._lyr(s, "GEAR_SOLID")
                except Exception as e:
                    self._log(f"      [!] Cone {ctx}: {e}")

            elif shape == "torus":
                maj = self._eval_expr(step.get("major_radius", "P1/2"), variables, ctx)
                mnr = self._eval_expr(step.get("minor_radius", "P2/2"), variables, ctx)
                try:
                    s = self._do(self.ms.AddTorus,
                                  self._pt(curr_x, curr_y, z_pos),
                                  float(maj), float(mnr))
                    if s: self._lyr(s, "GEAR_SOLID")
                except Exception as e:
                    self._log(f"      [!] Torus {ctx}: {e}")

            elif shape == "frustum":
                rb = self._eval_expr(step.get("radius_bottom", "P1/2"), variables, ctx)
                rt = self._eval_expr(step.get("radius_top",    "P2/2"), variables, ctx)
                h  = self._eval_expr(step.get("height",        "P3"),   variables, ctx)
                try:
                    s = self._do(self.ms.AddFrustum,
                                  self._pt(curr_x, curr_y, z_pos),
                                  float(rb), float(h), float(rt))
                    if s: self._lyr(s, "GEAR_SOLID")
                except Exception as e:
                    self._log(f"      [!] Frustum {ctx}: {e}")
                    s = self._cyl(curr_x, curr_y, z_pos, rb, h, "GEAR_SOLID")

            elif shape == "extrude_profile":
                raw_pts = step.get("points", [])
                h       = self._eval_expr(step.get("height",     "P3"), variables, ctx)
                taper   = self._eval_expr(step.get("taper_angle", 0),   variables, ctx)
                if len(raw_pts) >= 4:
                    wld = self._eval_pts(raw_pts, curr_x, curr_y, variables)
                    s   = self._profile_solid(wld, h, z=z_pos, taper=taper, layer="GEAR_SOLID")
                else:
                    self._log(f"      [!] extrude_profile needs ≥4 points ({ctx})")

            elif shape == "revolve":
                raw_pts = step.get("profile_points", [])
                deg     = self._eval_expr(step.get("degrees", 360),  variables, ctx)
                axis_s  = str(step.get("axis", "Y")).upper()
                if len(raw_pts) >= 4:
                    wld = self._eval_pts(raw_pts, curr_x, curr_y, variables)
                    ax_map = {
                        "X": ((curr_x, curr_y, z_pos), (1, 0, 0)),
                        "Y": ((curr_x, curr_y, z_pos), (0, 1, 0)),
                        "Z": ((curr_x, curr_y, z_pos), (0, 0, 1)),
                    }
                    ax_pt, ax_dir = ax_map.get(axis_s, ax_map["Y"])
                    s = self._revolve_profile(wld, z_pos, ax_pt, ax_dir, deg,
                                               layer="GEAR_SOLID")
                else:
                    self._log(f"      [!] revolve needs ≥4 profile_points ({ctx})")

            elif shape == "pipe":
                r_out = self._eval_expr(step.get("outer_radius", "P1/2"),  variables, ctx)
                r_in  = self._eval_expr(step.get("inner_radius", "P2/2"),  variables, ctx)
                h     = self._eval_expr(step.get("height",       "P3"),    variables, ctx)
                ovlp  = self._eval_expr(step.get("overlap", 0), variables, ctx)
                outer = self._cyl(curr_x, curr_y, z_pos - ovlp, r_out, h + ovlp, "GEAR_SOLID")
                if outer and r_in > 0.5 and r_in < r_out:
                    bore = self._cyl(curr_x, curr_y, z_pos - ovlp - 2, r_in, h + ovlp + 4, "GEAR_SOLID")
                    if bore: self._subtract(outer, bore)
                s = outer

            elif shape == "polygon_prism":
                # Regular N-sided prism — sides, radius (circumscribed), height
                sides = max(3, int(self._eval_expr(step.get("sides", 6),     variables, ctx)))
                r     = self._eval_expr(step.get("radius", "P1/2"),           variables, ctx)
                h     = self._eval_expr(step.get("height", "P3"),             variables, ctx)
                poly_pts: List[float] = []
                for k in range(sides):
                    a = 2 * math.pi * k / sides
                    poly_pts.append(curr_x + r * math.cos(a))
                    poly_pts.append(curr_y + r * math.sin(a))
                s = self._profile_solid(poly_pts, h, z=z_pos, layer="GEAR_SOLID")

            elif shape == "ellipsoid":
                # Sphere approximated as revolved half-ellipse — rx, ry, rz
                rx = self._eval_expr(step.get("rx", "P1/2"), variables, ctx)
                ry = self._eval_expr(step.get("ry", "P2/2"), variables, ctx)
                rz = self._eval_expr(step.get("rz", "P3/2"), variables, ctx)
                base_r = max(rx, ry, rz)
                if abs(rx - ry) < 1.0 and abs(ry - rz) < 1.0:
                    # Nearly spherical — sphere + uniform scale
                    try:
                        sph = self._do(self.ms.AddSphere,
                                       self._pt(curr_x, curr_y, z_pos + base_r), float(base_r))
                        if sph:
                            self._lyr(sph, "GEAR_SOLID")
                            if abs(rx - base_r) > 0.5:
                                self._do(sph.ScaleEntity,
                                         self._pt(curr_x, curr_y, z_pos + base_r),
                                         float(rx / base_r))
                            s = sph
                    except Exception as e:
                        self._log(f"      [!] Ellipsoid sphere {ctx}: {e}")
                else:
                    # Non-uniform: revolve half-ellipse cross-section around Z
                    n_e = 32
                    ep: List[float] = []
                    for k in range(n_e + 1):
                        theta = math.pi * k / n_e
                        ep.append(curr_x + rx * math.sin(theta))
                        ep.append(curr_y + rz * math.cos(theta))
                    s = self._revolve_profile(ep, z_pos,
                                               (curr_x, curr_y, z_pos), (0, 1, 0),
                                               360.0, "GEAR_SOLID")

            elif shape == "spring":
                # Helical spring swept along helix — coil_radius, wire_radius, pitch, turns
                cr    = self._eval_expr(step.get("coil_radius", "P1/2"), variables, ctx)
                wr    = self._eval_expr(step.get("wire_radius", "P2/2"), variables, ctx)
                ptch  = self._eval_expr(step.get("pitch",       "P3"),   variables, ctx)
                nt    = self._eval_expr(step.get("turns",       "P4"),   variables, ctx)
                n_seg = max(48, int(nt * 24))
                sp3d: List[float] = []
                for k in range(n_seg + 1):
                    frac = k / n_seg
                    ang  = 2 * math.pi * nt * frac
                    sp3d += [curr_x + cr * math.cos(ang),
                              curr_y + cr * math.sin(ang),
                              z_pos  + ptch * nt * frac]
                t0   = [-cr * math.sin(0),                  cr * math.cos(0),                  ptch / (2*math.pi)]
                ae   = 2 * math.pi * nt
                t1   = [-cr * math.sin(ae), cr * math.cos(ae), ptch / (2*math.pi)]
                sp_path = None; sp_prof = None
                try:
                    sp_path = self._do(self.ms.AddSpline,
                                       self._arr(sp3d),
                                       self._vec(*t0), self._vec(*t1))
                    self._lyr(sp_path, "WORK_GEOM")
                    x0 = curr_x + cr; sp_cpts: List[float] = []
                    for k in range(20):
                        a = 2 * math.pi * k / 20
                        sp_cpts += [x0 + wr * math.cos(a), curr_y + wr * math.sin(a)]
                    sp_prof = self._do(self.ms.AddLightWeightPolyline, self._arr(sp_cpts))
                    sp_prof.Closed = True; sp_prof.Elevation = z_pos; sp_prof.Layer = "WORK_GEOM"
                    s = self._do(self.ms.AddExtrudedSolidAlongPath, sp_prof, sp_path)
                    if s: self._lyr(s, "GEAR_SOLID")
                    self._del(sp_prof); self._del(sp_path)
                except Exception as e:
                    self._log(f"      [!] Spring {ctx}: {e}")
                    self._del(sp_path); self._del(sp_prof)

            else:
                self._log(f"      [!] Unknown shape '{shape}' (step {step_i})")

            # Apply optional rotation
            if s is not None:
                rot_axis = str(step.get("rotate_axis", "")).upper()
                rot_deg  = self._eval_expr(step.get("rotate_deg", 0), variables, ctx)
                if rot_deg != 0.0 and rot_axis:
                    self._apply_rotation(s, curr_x, curr_y, z_pos, rot_axis, rot_deg)

            return s

        # ── main step loop ────────────────────────────────────────────────────
        for step_i, step in enumerate(recipe_steps):
            action = str(step.get("action", "ADD")).upper()
            ctx    = f"step {step_i} {action}"

            # ── Geometry steps ────────────────────────────────────────────────
            if action in ("BASE", "ADD", "SUBTRACT", "INTERSECT"):
                temp = _make_solid(step, step_i)
                if temp is None:
                    self._log(f"      [!] No solid from {ctx} — skipping")
                    continue

                # Store for the PATTERN step that may follow
                last_tool   = temp
                last_action = action
                last_step   = step

                if action == "BASE":
                    if base_solid is not None:
                        self._del(base_solid)
                    base_solid = temp

                elif action == "ADD":
                    if base_solid is None:
                        self._log(f"      [!] ADD before BASE — promoting")
                        base_solid = temp
                    else:
                        self._union(base_solid, temp)

                elif action == "SUBTRACT":
                    if base_solid is None:
                        self._log(f"      [!] SUBTRACT before BASE — skipping")
                        self._del(temp)
                        last_tool = None
                    else:
                        self._subtract(base_solid, temp)

                elif action == "INTERSECT":
                    if base_solid is None:
                        self._log(f"      [!] INTERSECT before BASE — promoting to BASE")
                        base_solid = temp
                    else:
                        try:
                            self._do(base_solid.Boolean, AC_INTERSECT, temp)
                            self._log(f"      [Intersect] ✔")
                        except Exception as e:
                            self._log(f"      [!] INTERSECT failed: {e}")
                            self._del(temp)
                        last_tool = None; last_action = None

            # ── PATTERN_CIRCULAR ──────────────────────────────────────────────
            # FIX: build N-1 extra copies of last_tool only, union into one
            # combined cutter, apply a single Boolean to base_solid.
            elif action == "PATTERN_CIRCULAR":
                if last_tool is None or last_action not in ("ADD", "SUBTRACT"):
                    self._log(f"      [!] PATTERN_CIRCULAR at {ctx}: no preceding "
                               "ADD/SUBTRACT step to pattern — skipped")
                    continue

                count       = max(2, int(self._eval_expr(
                                  step.get("count", 4), variables, ctx)))
                total_angle = self._eval_expr(step.get("total_angle", 360), variables, ctx)
                pcx = cx + self._eval_expr(step.get("center_x", 0), variables, ctx)
                pcy = cy + self._eval_expr(step.get("center_y", 0), variables, ctx)
                span = math.radians(total_angle / count)

                # Rebuild a PRISTINE template tool.
                # CRITICAL: never union anything INTO pristine.
                # Every clone must come from pristine so combined stays linear.
                pristine = _make_solid(last_step, step_i)
                if pristine is None:
                    self._log(f"      [!] PATTERN_CIRCULAR: could not recreate tool")
                    continue

                self._log(f"      [Pattern_Circular] building {count} instances…")

                # Instance 0 — fresh copy at angle 0 (original position)
                combined = self._do(pristine.Copy)
                if combined is None:
                    self._del(pristine)
                    self._log(f"      [!] PATTERN_CIRCULAR: Copy() returned None")
                    continue

                ok = 1
                for i in range(1, count):
                    # Always copy PRISTINE — size stays constant, no explosion
                    clone = self._do(pristine.Copy)
                    if clone:
                        try:
                            self._do(clone.Rotate3D,
                                     self._pt(pcx, pcy, 0),
                                     self._pt(pcx, pcy, 1),
                                     span * i)
                            self._union(combined, clone)
                            ok += 1
                        except Exception as e:
                            self._log(f"      [!] Pattern copy {i}: {e}")
                            self._del(clone)
                    # Pulse every 8 steps so AutoCAD stays responsive
                    if i % 8 == 0:
                        try: pythoncom.PumpWaitingMessages()
                        except Exception: pass

                # Pristine served its purpose — remove from scene
                self._del(pristine)

                # ONE Boolean on base — this is the whole point
                if last_action == "SUBTRACT":
                    self._subtract(base_solid, combined)
                else:
                    self._union(base_solid, combined)

                self._log(f"      [Pattern_Circular] ✔ {ok}/{count} instances  "
                           f"action={last_action}  1 Boolean op")

                # Invalidate last_tool so a second PATTERN doesn't re-use it
                last_tool   = None
                last_action = None

            # ── PATTERN_LINEAR ────────────────────────────────────────────────
            elif action == "PATTERN_LINEAR":
                if last_tool is None or last_action not in ("ADD", "SUBTRACT"):
                    self._log(f"      [!] PATTERN_LINEAR at {ctx}: no preceding "
                               "ADD/SUBTRACT step — skipped")
                    continue

                count = max(2, int(self._eval_expr(step.get("count", 2), variables, ctx)))
                dx = self._eval_expr(step.get("dx", 0), variables, ctx)
                dy = self._eval_expr(step.get("dy", 0), variables, ctx)
                dz = self._eval_expr(step.get("dz", 0), variables, ctx)

                pristine = _make_solid(last_step, step_i)
                if pristine is None:
                    self._log(f"      [!] PATTERN_LINEAR: could not recreate tool")
                    continue

                self._log(f"      [Pattern_Linear] building {count} instances…")

                # Instance 0 at original position
                combined = self._do(pristine.Copy)
                if combined is None:
                    self._del(pristine); continue

                ok = 1
                for i in range(1, count):
                    clone = self._do(pristine.Copy)   # always from pristine
                    if clone:
                        try:
                            self._do(clone.Move,
                                     self._pt(0, 0, 0),
                                     self._pt(dx*i, dy*i, dz*i))
                            self._union(combined, clone)
                            ok += 1
                        except Exception as e:
                            self._log(f"      [!] Linear copy {i}: {e}")
                            self._del(clone)

                self._del(pristine)

                if last_action == "SUBTRACT":
                    self._subtract(base_solid, combined)
                else:
                    self._union(base_solid, combined)

                self._log(f"      [Pattern_Linear] ✔ {ok}/{count} instances  "
                           f"dx={dx} dy={dy} dz={dz}")
                last_tool = None; last_action = None

            # ── MIRROR ────────────────────────────────────────────────────────
            elif action == "MIRROR":
                if base_solid is None:
                    self._log(f"      [!] MIRROR: no base solid yet — skipped")
                    continue
                plane = str(step.get("plane", "XZ")).upper()
                plane_pts = {
                    "XZ": ((cx,cy,0),(cx+1,cy,0),(cx,cy,1)),
                    "YZ": ((cx,cy,0),(cx,cy+1,0),(cx,cy,1)),
                    "XY": ((cx,cy,0),(cx+1,cy,0),(cx+1,cy+1,0)),
                }
                if plane in plane_pts:
                    p1m, p2m, p3m = plane_pts[plane]
                    try:
                        mirror_obj = self._do(base_solid.Copy)
                        if mirror_obj:
                            self._do(mirror_obj.Mirror3D,
                                     self._pt(*p1m), self._pt(*p2m), self._pt(*p3m))
                            self._union(base_solid, mirror_obj)
                            self._log(f"      [Mirror] plane={plane} ✔")
                    except Exception as e:
                        self._log(f"      [!] Mirror {plane}: {e}")
                else:
                    self._log(f"      [!] MIRROR: unknown plane '{plane}'")

            # ── FILLET / CHAMFER — logged note, not executed ──────────────────
            # AutoCAD COM does not support automated edge-selection on 3D solids.
            # SendCommand is asynchronous and hangs waiting for user input.
            # Fillets must be applied manually in AutoCAD after build.
            elif action == "FILLET":
                r = self._eval_expr(step.get("radius", 1.0), variables, ctx)
                self._log(f"      [Fillet] NOTE: r={r:.1f}mm — apply manually "
                           f"in AutoCAD (FILLETEDGE). Geometry otherwise complete.")

            elif action == "CHAMFER":
                d = self._eval_expr(step.get("distance", 1.0), variables, ctx)
                self._log(f"      [Chamfer] NOTE: d={d:.1f}mm — apply manually "
                           f"in AutoCAD (CHAMFEREDGE). Geometry otherwise complete.")

            # ── SET_VAR — compute an intermediate value into the variable dict ──
            elif action == "SET_VAR":
                var_name = str(step.get("var", "V1"))
                expr     = step.get("expr", "0")
                result   = self._eval_expr(expr, variables, ctx)
                variables[var_name] = result
                self._log(f"      [SetVar] {var_name} = {result:.4f}  (expr: {expr})")

            # ── SET_VAR — compute an intermediate value into the variable dict ──
            elif action == "SCALE":
                if base_solid is None:
                    self._log(f"      [!] SCALE: no base solid — skipped"); continue
                factor = self._eval_expr(step.get("factor", 1.0), variables, ctx)
                try:
                    self._do(base_solid.ScaleEntity,
                              self._pt(cx, cy, 0), float(factor))
                    self._log(f"      [Scale] ✔ factor={factor:.4f}")
                except Exception as e:
                    self._log(f"      [!] SCALE failed: {e}")

            # ── SLICE — cut base_solid with a plane (big-box subtraction method) ──
            # plane_normal: "X"|"Y"|"Z"|[nx,ny,nz]  plane_z: float  keep: "+"|"-"
            elif action == "SLICE":
                if base_solid is None:
                    self._log(f"      [!] SLICE: no base solid — skipped"); continue
                normal_s = step.get("plane_normal", "Z")
                plane_z  = self._eval_expr(step.get("plane_z",  "P3/2"), variables, ctx)
                keep     = str(step.get("keep", "+")).strip()
                slab_size = max(
                    self._eval_expr(step.get("slab_size", "P1*4"), variables, ctx),
                    500.0
                )
                half = slab_size / 2.0
                # Build a slab on the unwanted side of the plane
                if str(normal_s).upper() == "Z":
                    if keep == "+":   # keep above plane_z → delete below
                        slab = self._box(cx - half, cy - half, plane_z - slab_size,
                                          slab_size, slab_size, slab_size, "GEAR_SOLID")
                    else:             # keep below → delete above
                        slab = self._box(cx - half, cy - half, plane_z,
                                          slab_size, slab_size, slab_size, "GEAR_SOLID")
                elif str(normal_s).upper() == "X":
                    if keep == "+":
                        slab = self._box(cx - slab_size, cy - half, -half,
                                          slab_size, slab_size, slab_size, "GEAR_SOLID")
                    else:
                        slab = self._box(cx + plane_z, cy - half, -half,
                                          slab_size, slab_size, slab_size, "GEAR_SOLID")
                elif str(normal_s).upper() == "Y":
                    if keep == "+":
                        slab = self._box(cx - half, cy - slab_size, -half,
                                          slab_size, slab_size, slab_size, "GEAR_SOLID")
                    else:
                        slab = self._box(cx - half, cy + plane_z, -half,
                                          slab_size, slab_size, slab_size, "GEAR_SOLID")
                else:
                    slab = None
                    self._log(f"      [!] SLICE: plane_normal must be X, Y, or Z")
                if slab:
                    self._subtract(base_solid, slab)
                    self._log(f"      [Slice] ✔ plane={normal_s}  z={plane_z:.2f}  keep={keep}")

            # ── SWEEP — sweep a 2D profile along a 3D path ───────────────────────
            # profile_points: flat [x,y,...] 2D   path_points: flat [x,y,z,...] 3D
            elif action == "SWEEP":
                if base_solid is None:
                    self._log(f"      [!] SWEEP: no base solid to union into — "
                               "SWEEP creates ADD geometry"); # allow standalone
                raw_prof = step.get("profile_points", [])
                raw_path = step.get("path_points",    [])
                z_elev   = self._eval_expr(step.get("z", 0), variables, ctx)
                if len(raw_prof) < 4 or len(raw_path) < 6:
                    self._log(f"      [!] SWEEP: need ≥4 profile pts and ≥6 (2×xyz) path pts")
                    continue
                # Build profile region
                wld_prof = self._eval_pts(raw_prof, cx, cy, variables)
                pl = self._lwpl(wld_prof, z_elev)
                if pl is None: continue
                reg = self._region(pl); self._del(pl)
                if reg is None: continue
                # Build 3D path spline
                ns2 = dict(_EVAL_NS); ns2.update(variables)
                pts3d: List[float] = []
                for k in range(0, len(raw_path) - 2, 3):
                    try:
                        pts3d.append(cx + float(eval(str(raw_path[k]),   ns2)))
                        pts3d.append(cy + float(eval(str(raw_path[k+1]), ns2)))
                        pts3d.append(      float(eval(str(raw_path[k+2]), ns2)))
                    except Exception as e:
                        self._log(f"      [!] SWEEP path pt {k}: {e}")
                if len(pts3d) < 9:
                    self._del(reg)
                    self._log(f"      [!] SWEEP: too few valid path pts"); continue
                n3 = len(pts3d) // 3
                tx = pts3d[3] - pts3d[0]; ty = pts3d[4] - pts3d[1]; tz = pts3d[5] - pts3d[2]
                ex = pts3d[-3] - pts3d[-6]; ey = pts3d[-2] - pts3d[-5]; ez = pts3d[-1] - pts3d[-4]
                path_sp = None
                try:
                    path_sp = self._do(self.ms.AddSpline,
                                       self._arr(pts3d),
                                       self._vec(tx, ty, tz), self._vec(ex, ey, ez))
                    self._lyr(path_sp, "WORK_GEOM")
                    swept = self._do(self.ms.AddExtrudedSolidAlongPath, reg, path_sp)
                    self._del(reg); self._del(path_sp)
                    if swept:
                        self._lyr(swept, "GEAR_SOLID")
                        sweep_action = str(step.get("boolean", "ADD")).upper()
                        last_tool = swept; last_action = sweep_action; last_step = step
                        if base_solid is None:
                            base_solid = swept
                        elif sweep_action == "SUBTRACT":
                            self._subtract(base_solid, swept)
                        else:
                            self._union(base_solid, swept)
                        self._log(f"      [Sweep] ✔  boolean={sweep_action}")
                    else:
                        self._del(reg); self._del(path_sp)
                except Exception as e:
                    self._log(f"      [!] SWEEP failed: {e}")
                    self._del(path_sp); self._del(reg)

            # ── ARRAY_GRID — rectangular rows × cols grid of last tool ────────────
            elif action == "ARRAY_GRID":
                if last_tool is None or last_action not in ("ADD", "SUBTRACT"):
                    self._log(f"      [!] ARRAY_GRID: no preceding ADD/SUBTRACT — skipped")
                    continue
                rows = max(1, int(self._eval_expr(step.get("rows", 2), variables, ctx)))
                cols = max(1, int(self._eval_expr(step.get("cols", 2), variables, ctx)))
                dx   = self._eval_expr(step.get("dx", "P1/4"), variables, ctx)
                dy   = self._eval_expr(step.get("dy", "P2/4"), variables, ctx)
                # Origin offset so the grid is centred
                ox   = self._eval_expr(step.get("start_x", 0), variables, ctx)
                oy   = self._eval_expr(step.get("start_y", 0), variables, ctx)
                total = rows * cols
                self._log(f"      [ArrayGrid] {rows}×{cols}={total} instances…")
                pristine = _make_solid(last_step, step_i)
                if pristine is None:
                    self._log(f"      [!] ARRAY_GRID: could not recreate tool"); continue
                combined = self._do(pristine.Copy)
                if combined is None: self._del(pristine); continue
                # Move instance-0 to grid origin
                self._do(combined.Move, self._pt(0,0,0), self._pt(ox, oy, 0))
                ok = 1
                for r in range(rows):
                    for c in range(cols):
                        if r == 0 and c == 0: continue
                        clone = self._do(pristine.Copy)
                        if clone:
                            try:
                                self._do(clone.Move,
                                         self._pt(0, 0, 0),
                                         self._pt(ox + c*dx, oy + r*dy, 0))
                                self._union(combined, clone); ok += 1
                            except Exception as e:
                                self._log(f"      [!] Grid ({r},{c}): {e}")
                                self._del(clone)
                self._del(pristine)
                if last_action == "SUBTRACT": self._subtract(base_solid, combined)
                else:                          self._union(base_solid, combined)
                self._log(f"      [ArrayGrid] ✔ {ok}/{total}  dx={dx}  dy={dy}")
                last_tool = None; last_action = None

            # ── HELIX_ARRAY — N copies distributed along a helix (Z + rotation) ──
            # Great for: cooling fins, thread start positions, stator slots
            elif action == "HELIX_ARRAY":
                if last_tool is None or last_action not in ("ADD", "SUBTRACT"):
                    self._log(f"      [!] HELIX_ARRAY: no preceding ADD/SUBTRACT — skipped")
                    continue
                count  = max(2, int(self._eval_expr(step.get("count",  8),    variables, ctx)))
                dz     = self._eval_expr(step.get("dz",   "P3/8"), variables, ctx)
                da_deg = self._eval_expr(step.get("da_deg", 45),   variables, ctx)
                pcx    = cx + self._eval_expr(step.get("center_x", 0), variables, ctx)
                pcy    = cy + self._eval_expr(step.get("center_y", 0), variables, ctx)
                self._log(f"      [HelixArray] {count} instances  dz={dz}  da={da_deg}°…")
                pristine = _make_solid(last_step, step_i)
                if pristine is None:
                    self._log(f"      [!] HELIX_ARRAY: could not recreate tool"); continue
                combined = self._do(pristine.Copy)
                if combined is None: self._del(pristine); continue
                ok = 1
                for i in range(1, count):
                    clone = self._do(pristine.Copy)
                    if clone:
                        try:
                            # Rotate around Z at part centre
                            self._do(clone.Rotate3D,
                                     self._pt(pcx, pcy, 0),
                                     self._pt(pcx, pcy, 1),
                                     math.radians(da_deg * i))
                            # Then translate up Z
                            self._do(clone.Move, self._pt(0,0,0), self._pt(0, 0, dz * i))
                            self._union(combined, clone); ok += 1
                        except Exception as e:
                            self._log(f"      [!] HelixArray {i}: {e}")
                            self._del(clone)
                self._del(pristine)
                if last_action == "SUBTRACT": self._subtract(base_solid, combined)
                else:                          self._union(base_solid, combined)
                self._log(f"      [HelixArray] ✔ {ok}/{count}")
                last_tool = None; last_action = None

            # ── PATTERN_RADIAL — copies at explicit angle list ────────────────────
            # Use when spacing is NOT uniform (e.g. crankshaft throws at 0/90/270/360)
            elif action == "PATTERN_RADIAL":
                if last_tool is None or last_action not in ("ADD", "SUBTRACT"):
                    self._log(f"      [!] PATTERN_RADIAL: no preceding ADD/SUBTRACT — skipped")
                    continue
                angles_raw = step.get("angles", [])
                if not isinstance(angles_raw, list) or len(angles_raw) < 1:
                    self._log(f"      [!] PATTERN_RADIAL: 'angles' must be a list of degrees")
                    continue
                pcx = cx + self._eval_expr(step.get("center_x", 0), variables, ctx)
                pcy = cy + self._eval_expr(step.get("center_y", 0), variables, ctx)
                ns3 = dict(_EVAL_NS); ns3.update(variables)
                angles_deg = []
                for a in angles_raw:
                    try: angles_deg.append(float(eval(str(a), ns3)))
                    except Exception: angles_deg.append(0.0)
                self._log(f"      [PatternRadial] {len(angles_deg)} angles: {angles_deg}…")
                pristine = _make_solid(last_step, step_i)
                if pristine is None:
                    self._log(f"      [!] PATTERN_RADIAL: could not recreate tool"); continue
                combined = self._do(pristine.Copy)
                if combined is None: self._del(pristine); continue
                ok = 1
                for ang_deg in angles_deg[1:]:    # instance-0 is already at angles_deg[0]
                    clone = self._do(pristine.Copy)
                    if clone:
                        try:
                            self._do(clone.Rotate3D,
                                     self._pt(pcx, pcy, 0), self._pt(pcx, pcy, 1),
                                     math.radians(ang_deg - angles_deg[0]))
                            self._union(combined, clone); ok += 1
                        except Exception as e:
                            self._log(f"      [!] RadialCopy {ang_deg}°: {e}")
                            self._del(clone)
                self._del(pristine)
                if last_action == "SUBTRACT": self._subtract(base_solid, combined)
                else:                          self._union(base_solid, combined)
                self._log(f"      [PatternRadial] ✔ {ok}/{len(angles_deg)}")
                last_tool = None; last_action = None

            else:
                self._log(f"      [!] Unknown action '{action}' at step {step_i}")

        if base_solid is None:
            self._log("      [!] Recipe produced no solid — check for BASE step")
            return None

        self._log("      [Recipe] ✔ Custom part built from JSON template")
        return base_solid

    # ══════════════════════════════════════════════════════════════════════════
    #  PER-TOOTH GEAR DISC
    # ══════════════════════════════════════════════════════════════════════════

    def _build_gear_disc(self, cx, cy, Z, m, face_w, PA_deg=20.0, x=0.0,
                          z0=0.0, angle_offset=0.0, N=48,
                          layer="GEAR_TEETH") -> Optional[object]:
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
                self._union(root_disc, ts); ok += 1
            else:
                fail += 1

        self._log(f"      [Teeth] ✔ {ok}/{Z} involute  {fail} fail")
        return root_disc

    # ── Bore + DIN 6885 keyway ────────────────────────────────────────────────

    def _bore_kw(self, solid, cx, cy, z_bot, z_top, bore_d):
        if solid is None or bore_d < 1.0: return
        h  = (z_top - z_bot) + 20.0
        br = bore_d / 2.0
        bc = self._cyl(cx, cy, z_bot - 10.0, br, h, "GEAR_BORE")
        if bc: self._subtract(solid, bc)
        kw_b = bore_d/4.0; kw_t = bore_d/8.0
        kw = self._box(cx-kw_b/2, cy+br-kw_t, z_bot-10.0, kw_b, kw_t+br, h, "GEAR_BORE")
        if kw: self._subtract(solid, kw)
        self._log(f"      [Bore] ✔ Ø{bore_d:.1f}  kw {kw_b:.1f}×{kw_t:.1f}")

    # ── Blank features ────────────────────────────────────────────────────────

    def _blank_p(self, Z, m, bore_d, face_w):
        pitch_r = Z*m/2.0
        root_r  = max(pitch_r-1.25*m, m*0.3)
        br      = bore_d/2.0
        hub_r   = min(max(br*1.7, br+m*1.3, br+5.0), root_r*0.54)
        hub_r   = max(hub_r, br+2.0)
        boss_h  = min(face_w*0.28, 12.0)
        web_t   = max(face_w*0.40, 5.0)
        recess_d= max((face_w - web_t)/2.0 - 0.5, 0.0)
        recess_r= root_r - m*0.65
        ann     = root_r - hub_r
        has_h   = False; n_h=0; bc_r=hub_r; hr=0.0
        if ann > 4.0*4.5:
            bc_r   = (hub_r + root_r*0.80)/2.0
            max_hr = min(bc_r-hub_r-2.5, root_r-bc_r-2.5, ann*0.28)
            hr     = max(4.0, max_hr)
            n_h    = 6 if Z<=36 else 8
            while n_h >= 4:
                if 2*math.pi*bc_r/n_h - 2*hr >= 3.0: break
                n_h -= 2
            has_h = (hr >= 4.0) and (n_h >= 4)
        return dict(hub_r=hub_r, boss_h=boss_h, recess_d=recess_d,
                    recess_r=recess_r, has_h=has_h, n_h=n_h, bc_r=bc_r, hr=hr)

    def _blank(self, solid, cx, cy, Z, m, bore_d, face_w, z0=0.0):
        if solid is None: return
        p  = self._blank_p(Z, m, bore_d, face_w)
        fz = z0 + face_w
        if p["hub_r"] > bore_d/2.0 + 1.0 and p["boss_h"] > 0.5:
            boss = self._cyl(cx, cy, fz, p["hub_r"], p["boss_h"], "GEAR_BLANK")
            if boss: self._union(solid, boss)
        rd = p["recess_d"]; rr = p["recess_r"]
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
    #  GEAR BUILDERS
    # ══════════════════════════════════════════════════════════════════════════

    def _gear_spur(self, cx, cy, Z, m, face_w, bore_d, PA_deg=20.0):
        x       = profile_shift_x(Z, PA_deg)
        pitch_r = Z*m/2.0
        outer_r = pitch_r+m*(1.0+x)
        bore_r  = bore_d/2.0
        hub_r   = max(bore_r*1.7, bore_r+m*1.3)
        hub_h   = face_w*0.55
        N       = max(24, min(48, Z*2))
        solid   = self._build_gear_disc(cx,cy,Z,m,face_w,PA_deg,x,z0=0.0,N=N)
        if solid is None:
            solid = self._cyl(cx,cy,0,outer_r,face_w,"GEAR_TEETH")
        if hub_r > bore_r+0.5:
            hub = self._cyl(cx,cy,-hub_h,hub_r,hub_h,"GEAR_BLANK")
            if hub: self._union(solid, hub)
        self._bore_kw(solid, cx, cy, -hub_h, face_w, bore_d)
        self._blank(solid, cx, cy, Z, m, bore_d, face_w, z0=0.0)
        return solid

    def _gear_helical(self, cx, cy, Z, m, face_w, bore_d, helix_deg=15.0,
                       PA_deg=20.0, right_hand=True):
        x           = profile_shift_x(Z, PA_deg)
        pitch_r     = Z*m/2.0
        outer_r     = pitch_r+m*(1.0+x)
        bore_r      = bore_d/2.0
        hub_r       = max(bore_r*1.7, bore_r+m*1.3)
        hub_h       = face_w*0.55
        total_twist = (face_w/pitch_r)*math.tan(math.radians(helix_deg))
        sign        = 1.0 if right_hand else -1.0
        N_SL        = 20; OVLP = 0.12
        N           = max(24, min(48, Z*2))
        slice_h = face_w/N_SL; ext_h = slice_h*(1+OVLP)
        base = None
        for sl in range(N_SL):
            ang_off = sign * sl * (total_twist/N_SL)
            sl_disc = self._build_gear_disc(cx,cy,Z,m,ext_h,PA_deg,x,
                                             z0=sl*slice_h,angle_offset=ang_off,N=N)
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

    def _gear_ring(self, cx, cy, Z, m, face_w, ring_thk, PA_deg=20.0):
        x       = profile_shift_x(Z, PA_deg)
        pitch_r = Z*m/2.0
        outer_r = pitch_r - m + ring_thk
        if outer_r <= pitch_r - m:
            outer_r = pitch_r + ring_thk*0.5
        disc = self._cyl(cx,cy,0,outer_r,face_w,"GEAR_TEETH")
        if disc is None: return None
        N = max(24, min(48, Z*2)); span = 2.0*math.pi/Z
        for i in range(Z):
            ang = i*span
            flat_loc = single_tooth_flat(Z,m,ang,PA_deg,x,N)
            flat_wld = []
            for k in range(0,len(flat_loc),2):
                flat_wld.append(flat_loc[k]+cx); flat_wld.append(flat_loc[k+1]+cy)
            vs = self._profile_solid(flat_wld, face_w+4.0, z=-2.0, layer="WORK_GEOM")
            if vs: self._subtract(disc, vs)
        return disc

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
            frustum = self._do(self.ms.AddFrustum, self._pt(cx,cy,0),
                               float(back_r),float(cone_h),float(front_r))
            self._lyr(frustum,"GEAR_TEETH")
        except Exception:
            try:
                frustum = self._do(self.ms.AddCone, self._pt(cx,cy,0),
                                   float(back_r),float(cone_h))
                self._lyr(frustum,"GEAR_TEETH")
            except Exception:
                return self._cyl(cx,cy,0,back_r,cone_h,"GEAR_TEETH")
        N=max(24,min(48,Z*2)); span=2*math.pi/Z
        for i in range(Z):
            ang = i*span
            tx  = cx+mean_r*math.cos(ang); ty = cy+mean_r*math.sin(ang)
            flat_loc = single_tooth_flat(Z,m_v,0.0,PA_deg,x_v,N)
            flat_wld = []
            for k in range(0,len(flat_loc),2):
                flat_wld.append(flat_loc[k]+tx); flat_wld.append(flat_loc[k+1]+ty)
            th_h = cone_h*0.70; th_z = cone_h*0.15
            ts = self._profile_solid(flat_wld,th_h,z=th_z,layer="GEAR_TEETH")
            if ts: self._union(frustum, ts)
        if hub_r>bore_r+0.5 and hub_r<back_r:
            hub=self._cyl(cx,cy,-hub_h,hub_r,hub_h,"GEAR_BLANK")
            if hub: self._union(frustum, hub)
        self._bore_kw(frustum,cx,cy,-hub_h,cone_h,bore_d)
        return frustum

    def _gear_worm(self, cx, cy, n_starts, m, length, bore_d, lead_deg=15.0):
        bore_r   = bore_d/2.0
        shaft_r  = max(bore_r+m*1.5, m*2.5)
        thread_r = m*0.55
        lead     = n_starts*math.pi*m
        n_turns  = length/lead if lead>0 else 2
        fl_r     = shaft_r+m*0.45; fl_h=m*0.80
        shaft    = self._cyl(cx,cy,0,shaft_r,length,"GEAR_TEETH")
        if shaft is None: return None
        helix_r  = shaft_r+thread_r*0.30
        for s in range(n_starts):
            ang0=s*2*math.pi/n_starts; path_sp=None; prof_pl=None
            try:
                npts=120; pts3d=[]
                for j in range(npts+1):
                    frac=j/npts; ang=ang0+2*math.pi*n_turns*frac
                    pts3d+=[cx+helix_r*math.cos(ang), cy+helix_r*math.sin(ang), length*frac]
                tang0=[-helix_r*math.sin(ang0), helix_r*math.cos(ang0), lead/(2*math.pi)]
                ang_end=ang0+2*math.pi*n_turns
                tang1=[-helix_r*math.sin(ang_end), helix_r*math.cos(ang_end), lead/(2*math.pi)]
                path_sp = self._do(self.ms.AddSpline,self._arr(pts3d),
                                   self._vec(*tang0),self._vec(*tang1))
                self._lyr(path_sp,"WORK_GEOM")
                x0=cx+helix_r*math.cos(ang0); y0=cy+helix_r*math.sin(ang0)
                nc=20; cpts=[]
                for j in range(nc):
                    a=2*math.pi*j/nc; cpts+=[x0+thread_r*math.cos(a), y0+thread_r*math.sin(a)]
                prof_pl = self._do(self.ms.AddLightWeightPolyline, self._arr(cpts))
                prof_pl.Closed=True; prof_pl.Layer="WORK_GEOM"; prof_pl.Elevation=0.0
                ts = self._do(self.ms.AddExtrudedSolidAlongPath, prof_pl, path_sp)
                self._lyr(ts,"GEAR_TEETH")
                self._union(shaft, ts)
                self._del(prof_pl); self._del(path_sp)
            except Exception:
                self._del(path_sp); self._del(prof_pl)
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

    def _gear_worm_wheel(self, cx, cy, Z, m, face_w, bore_d):
        x       = profile_shift_x(Z)
        pitch_r = Z*m/2.0
        outer_r = pitch_r+m*(1.0+x)
        bore_r  = bore_d/2.0
        hub_r   = max(bore_r*1.7, bore_r+m*1.2)
        hub_h   = face_w*0.50
        worm_sr = max(bore_r+m*1.5, m*2.5)
        cd      = pitch_r+worm_sr
        t_minor = worm_sr+m
        N       = max(24, min(48, Z*2))
        solid   = self._build_gear_disc(cx,cy,Z,m,face_w,20.0,x,z0=0.0,N=N)
        if solid is None:
            solid = self._cyl(cx,cy,0,outer_r,face_w,"GEAR_TEETH")
        try:
            tor = self._do(self.ms.AddTorus,
                           self._pt(cx,cy+cd,face_w/2.0),float(cd),float(t_minor))
            self._lyr(tor,"WORK_GEOM")
            self._subtract(solid, tor)
        except Exception: pass
        if hub_r>bore_r+0.5:
            hub=self._cyl(cx,cy,face_w,hub_r,hub_h,"GEAR_BLANK")
            if hub: self._union(solid, hub)
        self._bore_kw(solid,cx,cy,0,face_w+hub_h,bore_d)
        self._blank(solid,cx,cy,Z,m,bore_d,face_w,z0=0.0)
        return solid

    # ══════════════════════════════════════════════════════════════════════════
    #  TURBINE DISC  (industry-level — native Python, no AI recipe needed)
    # ══════════════════════════════════════════════════════════════════════════
    #
    #  P1 = OD (outer diameter, mm)        e.g. 300
    #  P2 = bore diameter (mm)             e.g. 60
    #  P3 = rim / hub thickness (mm)       e.g. 45
    #  P4 = number of blade slots          e.g. 36
    #
    #  What this builds:
    #    1. Disc body with proper web thinning + hub boss (CSG annular recesses)
    #    2. 3-lobe fir-tree blade slots cut into the rim (via extrude_profile)
    #    3. 8 balance holes at 60 % radius
    #    4. DIN-style shaft bore
    # ══════════════════════════════════════════════════════════════════════════

    def _turbine_disc(self, cx: float, cy: float,
                      OD: float, bore_d: float,
                      thickness: float, n_slots: int) -> Optional[object]:

        OD        = float(OD)
        bore_d    = float(bore_d)
        thickness = float(thickness)
        n_slots   = max(6, int(n_slots))

        rim_r  = OD / 2.0
        bore_r = bore_d / 2.0

        # ── Derived proportions (match real turbine disc ratios) ──────────────
        hub_r      = max(bore_r * 1.65, bore_r + thickness * 0.45)
        hub_r      = min(hub_r, rim_r * 0.24)        # hub never wider than 24 % OD
        hub_r      = max(hub_r, bore_r + 3.0)
        boss_h     = thickness * 0.38                 # hub boss protrudes above disc face
        web_inner_r = hub_r   * 1.18                 # web starts just outside hub
        web_outer_r = rim_r   * 0.74                 # web ends well before rim
        web_recess  = thickness * 0.34               # how deep the recess goes on each face
        rim_inner_r = rim_r   * 0.80                 # inner edge of the solid rim band

        self._log(f"      [Turbine] OD={OD}  bore={bore_d}  t={thickness}  "
                  f"slots={n_slots}  hub_r={hub_r:.1f}  web_recess={web_recess:.1f}")

        # ── 1. BASE: full disc cylinder ───────────────────────────────────────
        disc = self._cyl(cx, cy, 0, rim_r, thickness, "GEAR_SOLID")
        if disc is None:
            self._log("      [!] Turbine base cylinder failed"); return None

        # ── 2. Web thinning — annular recesses on both faces ─────────────────
        # Top recess
        top_recess = self._annulus(cx, cy,
                                    thickness - web_recess,
                                    web_outer_r, web_inner_r,
                                    web_recess + 3.0, "GEAR_BLANK")
        if top_recess: self._subtract(disc, top_recess)

        # Bottom recess
        bot_recess = self._annulus(cx, cy,
                                    -3.0,
                                    web_outer_r, web_inner_r,
                                    web_recess + 3.0, "GEAR_BLANK")
        if bot_recess: self._subtract(disc, bot_recess)

        self._log("      [Turbine] ✔ Web thinning applied")

        # ── 3. Hub boss — raised cylindrical boss on top face ─────────────────
        boss = self._cyl(cx, cy, thickness, hub_r, boss_h, "GEAR_BLANK")
        if boss: self._union(disc, boss)
        self._log(f"      [Turbine] ✔ Hub boss Ø{hub_r*2:.1f} × {boss_h:.1f}mm")

        # ── 4. Shaft bore ─────────────────────────────────────────────────────
        full_h = thickness + boss_h + 20.0
        bore_cyl = self._cyl(cx, cy, -10.0, bore_r, full_h, "GEAR_BORE")
        if bore_cyl: self._subtract(disc, bore_cyl)
        self._log(f"      [Turbine] ✔ Bore Ø{bore_d:.1f}")

        # ── 5. Fir-tree blade slots ───────────────────────────────────────────
        # 3-lobe profile (top-view, extruded through disc thickness)
        # Geometry based on standard turbine fir-tree proportions:
        #   outer lobe → neck → middle lobe → neck → root lobe

        slot_depth = rim_r * 0.52      # total radial penetration from rim surface

        # Lobe widths (tangential, widest at outside)
        w_out  = rim_r * 0.092         # outer lobe  (~13.8 mm for OD=300)
        w_mid  = rim_r * 0.062         # middle lobe (~9.3 mm)
        w_root = rim_r * 0.044         # root lobe   (~6.6 mm)
        w_neck = rim_r * 0.036         # neck width  (~5.4 mm)

        # Depth zones (radial)
        d_out   = slot_depth * 0.27
        d_neck1 = slot_depth * 0.09
        d_mid   = slot_depth * 0.27
        d_neck2 = slot_depth * 0.09
        d_root  = slot_depth * 0.28    # slightly deeper root

        # Cumulative depths
        y1  = -d_out
        y2  = y1  - d_neck1
        y3  = y2  - d_mid
        y4  = y3  - d_neck2
        y5  = y4  - d_root            # deepest point

        # 2D profile in local space:
        #   local X = tangential direction (±)
        #   local Y = radial inward (0 = rim surface, negative = into disc)
        # Closed polygon, clockwise from top-left corner:
        ft_local = [
            # outer lobe
            -w_out/2,  0,
            -w_out/2,  y1,
            # neck 1 (narrow)
            -w_neck/2, y1,
            -w_neck/2, y2,
            # middle lobe
            -w_mid/2,  y2,
            -w_mid/2,  y3,
            # neck 2
            -w_neck/2, y3,
            -w_neck/2, y4,
            # root lobe
            -w_root/2, y4,
            -w_root/2, y5,
            # root bottom (flat)
             w_root/2, y5,
            # mirror right side upward
             w_root/2, y4,
             w_neck/2, y4,
             w_neck/2, y3,
             w_mid/2,  y3,
             w_mid/2,  y2,
             w_neck/2, y2,
             w_neck/2, y1,
             w_out/2,  y1,
             w_out/2,  0,
        ]

        # Convert to world XY at angle = 0  (slot template at +X axis of disc)
        # local tangential (X) → world Y offset from cy
        # local radial    (Y) → world X offset from (cx + rim_r)
        #   local_y = 0  → rim surface   → wld_x = cx + rim_r
        #   local_y < 0  → inward        → wld_x = cx + rim_r + local_y  (smaller)
        ft_world = []
        for k in range(0, len(ft_local), 2):
            lx = ft_local[k]       # tangential
            ly = ft_local[k + 1]   # radial (0 = rim, negative = inward)
            ft_world.append(cx + rim_r + ly)   # ly ≤ 0 → inside disc ✓
            ft_world.append(cy + lx)

        # Build the template slot solid (slightly over-height for clean Boolean)
        slot_tmpl = self._profile_solid(ft_world, thickness + 10.0,
                                         z=-5.0, layer="GEAR_BORE")
        if slot_tmpl is None:
            self._log("      [!] Turbine: fir-tree profile_solid failed — "
                      "falling back to rectangular slots")
            # Fallback: simple rectangular slot at rim
            sw = w_out; sd = slot_depth
            slot_tmpl = self._box(cx + rim_r - sd, cy - sw/2, -5.0,
                                   sd, sw, thickness + 10.0, "GEAR_BORE")

        if slot_tmpl is not None:
            # Pattern using pristine-copy approach (no exponential explosion)
            pristine = self._do(slot_tmpl.Copy)
            if pristine is None:
                self._del(slot_tmpl)
                pristine = self._profile_solid(ft_world, thickness + 10.0,
                                                z=-5.0, layer="GEAR_BORE")

            combined = self._do(pristine.Copy)   # instance 0 at angle 0
            span     = 2.0 * math.pi / n_slots
            ok       = 1

            for i in range(1, n_slots):
                clone = self._do(pristine.Copy)
                if clone:
                    try:
                        self._do(clone.Rotate3D,
                                 self._pt(cx, cy, 0),
                                 self._pt(cx, cy, 1),
                                 span * i)
                        self._union(combined, clone)
                        ok += 1
                    except Exception as e:
                        self._log(f"      [!] Slot pattern {i}: {e}")
                        self._del(clone)
                if i % 8 == 0:
                    try: pythoncom.PumpWaitingMessages()
                    except Exception: pass

            self._del(pristine)
            self._del(slot_tmpl)
            self._subtract(disc, combined)
            self._log(f"      [Turbine] ✔ {ok}/{n_slots} fir-tree slots cut")

        # ── 6. Balance holes (8×) ─────────────────────────────────────────────
        n_bal    = 8
        bal_r    = rim_r * 0.026      # hole radius
        bal_ring = rim_r * 0.60       # circle of hole centres

        bal_tmpl = self._cyl(cx + bal_ring, cy, -5.0, bal_r,
                              thickness + 10.0, "GEAR_BORE")
        if bal_tmpl is not None:
            pristine_b  = self._do(bal_tmpl.Copy)
            combined_b  = self._do(pristine_b.Copy)
            b_span      = 2.0 * math.pi / n_bal
            ok_b        = 1

            for i in range(1, n_bal):
                clone = self._do(pristine_b.Copy)
                if clone:
                    try:
                        self._do(clone.Rotate3D,
                                 self._pt(cx, cy, 0),
                                 self._pt(cx, cy, 1),
                                 b_span * i)
                        self._union(combined_b, clone)
                        ok_b += 1
                    except Exception as e:
                        self._del(clone)

            self._del(pristine_b)
            self._del(bal_tmpl)
            self._subtract(disc, combined_b)
            self._log(f"      [Turbine] ✔ {ok_b}/{n_bal} balance holes")

        self._log(f"      [Turbine] ✔ COMPLETE — "
                  f"OD={OD}  bore={bore_d}  slots={n_slots}")
        return disc

    # ══════════════════════════════════════════════════════════════════════════
    #  INDUSTRY BUILDERS  (Bharat Forge + L&T demo suite)
    # ══════════════════════════════════════════════════════════════════════════

    # ─────────────────────────────────────────────────────────────────────────
    #  1. TURBINE BLADE  (Bharat Forge — Aerospace)
    #     P1 = blade span mm       (root to tip)
    #     P2 = root chord mm       (widest cross-section)
    #     P3 = tip chord mm        (narrowest cross-section, < P2)
    #     P4 = aerodynamic twist ° (root to tip, typical 10-30°)
    #
    #  Architecture: N thin slices each with an aerofoil extrude_profile,
    #  each slice rotated by twist/N — same multi-slice approach as helical gear.
    #  NACA 65-series approximation (modified joukowski profile via 20 points).
    # ─────────────────────────────────────────────────────────────────────────

    @staticmethod
    def _naca_profile(chord: float, n: int = 20) -> List[float]:
        """
        Returns a flat [x,y,...] closed aerofoil profile in local coords.
        Profile is centred at (chord/2, 0), chord runs along X.
        Uses a NACA-65 thickness distribution (max t/c ≈ 10%).
        """
        tc   = 0.10   # thickness-to-chord ratio
        pts: List[Tuple[float,float]] = []
        # Upper surface (leading edge → trailing edge)
        for i in range(n + 1):
            xn = i / n   # normalised 0→1
            # NACA symmetric thickness half-distribution (5-coeff)
            yt = 5 * tc * chord * (
                0.2969 * math.sqrt(xn)
                - 0.1260 * xn
                - 0.3516 * xn**2
                + 0.2843 * xn**3
                - 0.1015 * xn**4
            )
            pts.append((xn * chord, yt))
        # Lower surface (trailing edge → leading edge)
        for i in range(n, -1, -1):
            xn = i / n
            yt = 5 * tc * chord * (
                0.2969 * math.sqrt(xn)
                - 0.1260 * xn
                - 0.3516 * xn**2
                + 0.2843 * xn**3
                - 0.1015 * xn**4
            )
            pts.append((xn * chord, -yt))
        flat: List[float] = []
        for px, py in pts:
            flat.append(px); flat.append(py)
        return flat

    def _turbine_blade(self, cx: float, cy: float,
                       span: float, root_chord: float,
                       tip_chord: float, twist_deg: float) -> Optional[object]:

        span       = max(float(span),        50.0)
        root_chord = max(float(root_chord),  20.0)
        tip_chord  = max(float(tip_chord),   10.0)
        twist_deg  = float(twist_deg)

        N_SL    = 16                   # slices — enough for smooth taper + twist
        OVLP    = 0.08                 # 8 % height overlap for clean union
        sl_h    = span / N_SL
        ext_h   = sl_h * (1.0 + OVLP)

        self._log(f"      [Blade] span={span}  rootC={root_chord}  "
                  f"tipC={tip_chord}  twist={twist_deg}°  {N_SL} slices")

        blade = None
        for sl in range(N_SL):
            frac   = sl / N_SL
            chord  = root_chord + (tip_chord - root_chord) * frac   # linear taper
            angle  = math.radians(twist_deg * frac)                  # progressive twist

            prof_local = self._naca_profile(chord, n=20)             # 84 floats

            # Rotate profile around its own centroid (chord/2, 0) by twist angle
            cx_p = chord / 2.0
            prof_wld: List[float] = []
            for k in range(0, len(prof_local), 2):
                px = prof_local[k] - cx_p   # centre around 0
                py = prof_local[k + 1]
                rx = px * math.cos(angle) - py * math.sin(angle)
                ry = px * math.sin(angle) + py * math.cos(angle)
                prof_wld.append(cx + rx)
                prof_wld.append(cy + ry)

            z_base = sl * sl_h
            s = self._profile_solid(prof_wld, ext_h, z=z_base, layer="GEAR_TEETH")
            if s is None:
                continue
            if blade is None:
                blade = s
            else:
                self._union(blade, s)

        if blade is None:
            self._log("      [!] Blade: all slices failed — returning cylinder fallback")
            return self._cyl(cx, cy, 0, root_chord / 2.0, span, "GEAR_TEETH")

        # ── Root platform (dovetail base block) ───────────────────────────────
        plat_h = span * 0.08
        plat_w = root_chord * 1.30
        plat_d = root_chord * 0.40
        plat   = self._box(cx - plat_w/2, cy - plat_d/2, -plat_h,
                            plat_w, plat_d, plat_h, "GEAR_BLANK")
        if plat: self._union(blade, plat)

        # ── Dovetail root (trapezoid narrowing toward base) ───────────────────
        dtail_pts = [
            cx - plat_w*0.45, cy - plat_d*0.45,
            cx + plat_w*0.45, cy - plat_d*0.45,
            cx + plat_w*0.30, cy + plat_d*0.45,
            cx - plat_w*0.30, cy + plat_d*0.45,
        ]
        dtail = self._profile_solid(dtail_pts, plat_h * 1.6,
                                     z = -(plat_h * 2.6), layer="GEAR_BLANK")
        if dtail: self._union(blade, dtail)

        self._log(f"      [Blade] ✔ COMPLETE — {N_SL} aerofoil slices + dovetail root")
        return blade

    # ─────────────────────────────────────────────────────────────────────────
    #  COMPLETE TURBINE STAGE  (disc + all blades assembled)
    #
    #  P1 = disc OD mm            e.g. 400
    #  P2 = bore diameter mm      e.g. 80
    #  P3 = disc thickness mm     e.g. 55
    #  P4 = number of blades      e.g. 24
    #
    #  What this builds:
    #    Disc   — web thinning, hub boss, bore, balance holes, fir-tree slots
    #    Blades — NACA 4-series cambered aerofoil (24 pts per section, 16 slices)
    #             progressive taper root→tip, 32° aerodynamic twist,
    #             root platform with angel-wing seal flanges,
    #             tip shroud with interlocking Z-notch,
    #             3 internal cooling holes per blade
    #    All blades patterned and unioned to disc as single solid
    # ─────────────────────────────────────────────────────────────────────────

    @staticmethod
    def _naca4_cambered(chord: float, tc: float = 0.10,
                         mc: float = 0.045, pc: float = 0.40,
                         n: int = 24) -> List[float]:
        """
        NACA 4-digit cambered aerofoil in local coords.
        chord runs 0→chord along X.  Camber makes upper/lower surfaces asymmetric.
        tc = max thickness / chord   (0.10 = 10 %)
        mc = max camber / chord      (0.045 = 4.5 %)
        pc = chordwise position of max camber  (0.40 = 40 % chord)
        """
        upper: List[Tuple[float,float]] = []
        lower: List[Tuple[float,float]] = []

        for i in range(n + 1):
            xn = i / n
            # Mean camber line and slope
            if xn <= pc:
                yc  = mc / pc**2 * (2.0 * pc * xn - xn**2)
                dyc = mc / pc**2 * (2.0 * pc - 2.0 * xn)
            else:
                yc  = mc / (1.0 - pc)**2 * ((1.0 - 2.0*pc) + 2.0*pc*xn - xn**2)
                dyc = mc / (1.0 - pc)**2 * (2.0*pc - 2.0*xn)

            # NACA thickness distribution
            yt = 5.0 * tc * (
                0.2969 * math.sqrt(max(xn, 1e-9))
                - 0.1260 * xn
                - 0.3516 * xn**2
                + 0.2843 * xn**3
                - 0.1015 * xn**4
            )
            theta = math.atan(dyc)
            upper.append((chord * (xn - yt * math.sin(theta)),
                           chord * (yc  + yt * math.cos(theta))))
            lower.append((chord * (xn + yt * math.sin(theta)),
                           chord * (yc  - yt * math.cos(theta))))

        flat: List[float] = []
        for px, py in upper:
            flat.append(px); flat.append(py)
        for px, py in reversed(lower):
            flat.append(px); flat.append(py)
        return flat

    def _turbine_stage(self, cx: float, cy: float,
                        disc_od: float, bore_d: float,
                        disc_t: float, n_blades: int) -> Optional[object]:

        disc_od  = max(float(disc_od),  150.0)
        bore_d   = max(float(bore_d),    20.0)
        disc_t   = max(float(disc_t),    20.0)
        n_blades = max(8, int(n_blades))

        rim_r  = disc_od / 2.0
        bore_r = bore_d  / 2.0

        # ── Blade proportions (scaled to disc) ───────────────────────────────
        blade_span   = rim_r * 0.78          # span above disc face
        root_chord   = disc_t * 1.30         # chord at root
        tip_chord    = root_chord * 0.55     # chord at tip (42 % taper)
        stagger_root = 38.0                  # aero stagger at root (deg)
        twist_total  = 32.0                  # tip twist relative to root (deg)

        # Platform (angel-wing seal flanges at blade root)
        pitch_arc    = 2.0 * math.pi * rim_r / n_blades
        plat_arc     = pitch_arc * 0.85      # platform covers 85 % of pitch
        plat_h       = disc_t * 0.08
        plat_radial  = root_chord * 1.08     # platform radial extent

        # Tip shroud (interlocking Z-lock)
        shroud_h     = root_chord * 0.09
        shroud_arc   = plat_arc * 1.05       # shroud slightly wider than platform
        shroud_radial = root_chord * 0.95

        # Cooling holes
        cool_r       = root_chord * 0.048
        cool_offsets = [0.22, 0.50, 0.78]    # fractional span positions

        self._log(f"      [Stage] OD={disc_od}  n={n_blades}  "
                  f"span={blade_span:.0f}  chord={root_chord:.0f}→{tip_chord:.0f}  "
                  f"twist={twist_total}°")

        # ══ STEP 1: DISC ═════════════════════════════════════════════════════
        self._log("      [Stage] Building disc…")
        disc = self._turbine_disc(cx, cy, disc_od, bore_d, disc_t, n_blades)
        if disc is None:
            self._log("      [!] Stage: disc failed"); return None

        # ══ STEP 2: BUILD ONE BLADE TEMPLATE AT ANGLE = 0 ═══════════════════
        # Blade profile is in XY plane at each Z slice.
        # Root sits at (cx + rim_r, cy), chord is staggered in XY.
        # Span rises in +Z from disc_t (top face) to disc_t + blade_span.

        self._log("      [Stage] Building blade template…")

        N_SL  = 16
        OVLP  = 0.07
        sl_h  = blade_span / N_SL
        ext_h = sl_h * (1.0 + OVLP)

        blade = None
        for sl in range(N_SL):
            frac    = sl / N_SL
            chord   = root_chord + (tip_chord - root_chord) * frac  # linear taper
            stagger = math.radians(stagger_root + twist_total * frac) # progressive twist

            prof = self._naca4_cambered(chord, tc=0.10, mc=0.045, pc=0.40, n=24)

            # Centre profile at leading edge, apply stagger rotation,
            # then translate to blade root at (cx + rim_r, cy)
            cx_le = 0.0   # leading edge is at x=0
            prof_wld: List[float] = []
            for k in range(0, len(prof), 2):
                px = prof[k] - chord * 0.25   # pivot at quarter-chord
                py = prof[k + 1]
                # Rotate by stagger (chord goes from tangential toward radial)
                rx =  px * math.cos(stagger) - py * math.sin(stagger)
                ry =  px * math.sin(stagger) + py * math.cos(stagger)
                prof_wld.append(cx + rim_r + rx)
                prof_wld.append(cy + ry)

            z_base = disc_t + sl * sl_h
            s = self._profile_solid(prof_wld, ext_h, z=z_base, layer="GEAR_TEETH")
            if s is None: continue
            if blade is None: blade = s
            else:             self._union(blade, s)

        if blade is None:
            self._log("      [!] Stage: all blade slices failed — cylinder fallback")
            blade = self._cyl(cx + rim_r, cy, disc_t,
                               root_chord / 2.0, blade_span, "GEAR_TEETH")
        if blade is None:
            self._log("      [!] Stage: blade fallback also failed"); return disc

        # ── Root platform (angel-wing seals) ──────────────────────────────────
        # Flat platform at disc face, slightly above disc_t, full pitch coverage
        plat_x0 = cx + rim_r - plat_radial * 0.48
        plat_y0 = cy - plat_arc / 2.0
        plat = self._box(plat_x0, plat_y0, disc_t - plat_h,
                          plat_radial, plat_arc, plat_h, "GEAR_BLANK")
        if plat: self._union(blade, plat)

        # Angel-wing: thin forward seal flange (overhangs disc OD slightly)
        aw_h = plat_h * 0.35;  aw_t = plat_h * 0.45
        aw_fwd = self._box(cx + rim_r * 1.015, plat_y0 - aw_t,
                            disc_t - plat_h, aw_t * 2.5, plat_arc + aw_t * 2, aw_h,
                            "GEAR_BLANK")
        if aw_fwd: self._union(blade, aw_fwd)

        # ── Neck fillet transition (short frustum-like thickening above platform) ─
        neck_h = disc_t * 0.06
        neck   = self._cyl(cx + rim_r, cy, disc_t,
                            root_chord * 0.18, neck_h, "GEAR_BLANK")
        if neck: self._union(blade, neck)

        # ── Tip shroud with Z-lock notch ──────────────────────────────────────
        tip_z  = disc_t + blade_span
        tip_cx = cx + rim_r + blade_span * 0.025  # very slight radial sweep at tip

        shroud = self._box(tip_cx - shroud_radial * 0.48, cy - shroud_arc / 2.0,
                            tip_z, shroud_radial, shroud_arc, shroud_h, "GEAR_BLANK")
        if shroud: self._union(blade, shroud)

        # Z-notch (creates interlocking Z-profile between adjacent shrouds)
        notch_w = shroud_arc * 0.28;  notch_d = shroud_radial * 0.38
        notch = self._box(tip_cx + shroud_radial * 0.08,
                           cy + shroud_arc * 0.22,
                           tip_z - 1.0,
                           notch_d, notch_w, shroud_h + 2.0, "GEAR_BLANK")
        if notch: self._subtract(blade, notch)

        # Snubber rib (adds the positive Z-lock on the other edge)
        rib = self._box(tip_cx - shroud_radial * 0.46,
                         cy - shroud_arc * 0.50,
                         tip_z, notch_d, notch_w, shroud_h * 1.30, "GEAR_BLANK")
        if rib: self._union(blade, rib)

        # ── Internal cooling holes ────────────────────────────────────────────
        # Three cylindrical passages drilled perpendicular to blade span (in Z).
        # Drilled through the blade at 22 %, 50 %, 78 % of span.
        # Holes are oriented along Y (tangential) for film-cooling exit at TE.
        for frac_z in cool_offsets:
            cool_z   = disc_t + blade_span * frac_z
            # Build bore along Z then Rotate3D 90° around X at blade centre
            # to make it run through the profile in Y direction
            cool_cyl = self._cyl(cx + rim_r, cy,
                                   cool_z - root_chord * 0.8,
                                   cool_r, root_chord * 1.6, "GEAR_BORE")
            if cool_cyl:
                try:
                    self._do(cool_cyl.Rotate3D,
                              self._pt(cx + rim_r, cy, cool_z),
                              self._pt(cx + rim_r + 1, cy, cool_z),
                              math.pi / 2.0)
                    self._subtract(blade, cool_cyl)
                except Exception as e:
                    self._log(f"      [!] Cooling hole {frac_z}: {e}")
                    self._del(cool_cyl)

        # ── Trailing-edge cooling slot (very thin Z-extrusion slot at TE) ────
        te_slot_w = root_chord * 0.02
        te_slot_h = blade_span * 0.55
        # Trailing edge position after stagger: approximately at
        # (cx + rim_r + root_chord*0.75*sin(stagger), cy - root_chord*0.75*cos(stagger))
        s_rad = math.radians(stagger_root)
        te_x  = cx + rim_r + root_chord * 0.68 * math.cos(s_rad)
        te_y  = cy          + root_chord * 0.68 * math.sin(s_rad)
        te_slot = self._box(te_x - te_slot_w / 2, te_y - te_slot_w / 2,
                             disc_t + blade_span * 0.22,
                             te_slot_w, te_slot_w, te_slot_h, "GEAR_BORE")
        if te_slot: self._subtract(blade, te_slot)

        self._log("      [Stage] ✔ Blade template complete "
                  "(NACA4 cambered, platform, tip shroud, 3 cooling holes)")

        # ══ STEP 3: PATTERN ALL BLADES AND UNION TO DISC ═════════════════════
        self._log(f"      [Stage] Patterning {n_blades} blades…")

        pristine = self._do(blade.Copy)
        if pristine is None:
            self._log("      [!] Stage: Copy() on blade returned None")
            self._union(disc, blade)
            return disc

        combined_blades = self._do(pristine.Copy)   # instance 0 at angle 0
        if combined_blades is None:
            self._del(pristine); self._del(blade)
            return disc

        b_span = 2.0 * math.pi / n_blades
        ok     = 1

        for i in range(1, n_blades):
            clone = self._do(pristine.Copy)
            if clone:
                try:
                    self._do(clone.Rotate3D,
                              self._pt(cx, cy, 0),
                              self._pt(cx, cy, 1),
                              b_span * i)
                    self._union(combined_blades, clone)
                    ok += 1
                except Exception as e:
                    self._log(f"      [!] Blade {i}: {e}")
                    self._del(clone)
            if i % 6 == 0:
                try: pythoncom.PumpWaitingMessages()
                except Exception: pass

        self._del(pristine)
        self._del(blade)

        # ── Single Boolean: all blades → disc ─────────────────────────────────
        self._union(disc, combined_blades)
        self._log(f"      [Stage] ✔ {ok}/{n_blades} blades unioned to disc")

        # ══ STEP 4: APPLY MATERIAL COLOUR ════════════════════════════════════
        # Nickel superalloy colour (warm silver-gold — IN718 appearance)
        self._rgb(disc, 210, 185, 140)

        self._log(f"      [Stage] ✔ COMPLETE — "
                  f"tip Ø{disc_od + 2*blade_span:.0f}mm  "
                  f"height={disc_t + blade_span:.0f}mm")
        return disc

    # ─────────────────────────────────────────────────────────────────────────
    #  2. CRANKSHAFT  (Bharat Forge — Automotive / Heavy Engine)
    #     P1 = main journal diameter mm
    #     P2 = stroke mm               (= 2 × throw radius)
    #     P3 = rod journal diameter mm
    #     P4 = number of cylinders     (4 or 6 or 8)
    #
    #  Architecture: central spine cylinder + off-axis rod journals
    #  at parametric throw positions + counterweights + oil bore.
    # ─────────────────────────────────────────────────────────────────────────

    def _crankshaft(self, cx: float, cy: float,
                    mj_dia: float, stroke: float,
                    rj_dia: float, n_cyl: int) -> Optional[object]:

        mj_dia  = max(float(mj_dia), 30.0)
        stroke  = max(float(stroke), 20.0)
        rj_dia  = max(float(rj_dia), 20.0)
        n_cyl   = max(2, int(n_cyl))

        mj_r    = mj_dia  / 2.0
        rj_r    = rj_dia  / 2.0
        throw   = stroke  / 2.0          # eccentric offset

        # Journal geometry
        mj_len  = mj_dia  * 0.85         # main journal length
        rj_len  = rj_dia  * 1.00         # rod journal length
        web_thk = mj_dia  * 0.55         # crank web thickness
        web_w   = mj_dia  * 1.80         # web width (tangential)
        cw_r    = mj_dia  * 1.05         # counterweight outer radius
        cw_thk  = web_thk                # same thickness as web
        oil_r   = mj_r   * 0.22          # central oil bore

        # Total shaft length
        unit    = mj_len + web_thk       # one main + one web unit
        total_L = unit * (n_cyl + 1)     # n+1 main journals, n throws

        self._log(f"      [Crank] n={n_cyl}  Ø{mj_dia}mj  Ø{rj_dia}rj  "
                  f"stroke={stroke}  L={total_L:.0f}mm")

        # ── 1. Main spine ─────────────────────────────────────────────────────
        shaft = self._cyl(cx, cy, 0, mj_r, total_L, "GEAR_SOLID")
        if shaft is None: return None

        # ── Firing-order angles for each cylinder ─────────────────────────────
        # Even firing for inline: each throw offset by 720°/n_cyl
        throw_angles: List[float] = []
        if n_cyl == 4:
            throw_angles = [0, 180, 180, 0]        # classic inline-4
        elif n_cyl == 6:
            throw_angles = [0, 120, 240, 240, 120, 0]
        elif n_cyl == 8:
            throw_angles = [0, 90, 270, 180, 180, 270, 90, 0]
        else:
            throw_angles = [i * 360.0 / n_cyl for i in range(n_cyl)]

        # ── 2. Rod journals + webs + counterweights ───────────────────────────
        for i in range(n_cyl):
            z_web1   = (i + 1) * mj_len + i * (rj_len + 2 * web_thk)
            z_rj     = z_web1 + web_thk
            z_web2   = z_rj   + rj_len
            ang      = math.radians(throw_angles[i % len(throw_angles)])
            off_x    = cx + throw * math.cos(ang)
            off_y    = cy + throw * math.sin(ang)
            cw_x     = cx - throw * math.cos(ang) * 0.70   # counterweight opposite
            cw_y     = cy - throw * math.sin(ang) * 0.70

            # Rod journal
            rj = self._cyl(off_x, off_y, z_rj, rj_r, rj_len, "GEAR_SOLID")
            if rj: self._union(shaft, rj)

            # Web 1 (main → rod)
            w1 = self._box(cx - web_w/2, cy - web_thk/2, z_web1,
                            web_w, web_thk, web_thk, "GEAR_BLANK")
            if w1: self._union(shaft, w1)

            # Web 2 (rod → main)
            w2 = self._box(cx - web_w/2, cy - web_thk/2, z_web2,
                            web_w, web_thk, web_thk, "GEAR_BLANK")
            if w2: self._union(shaft, w2)

            # Counterweight (half-disc opposite the throw)
            cw_pts = []
            for k in range(9):
                a = math.pi + math.pi * k / 8   # half-circle on the opposite side
                cw_pts.append(cx + cw_r * math.cos(a))
                cw_pts.append(cy + cw_r * math.sin(a))
            # Close with the flat edge through the shaft centre
            cw_pts += [cx + mj_r, cy, cx - mj_r, cy]
            cw = self._profile_solid(cw_pts, cw_thk, z=z_web1, layer="GEAR_BLANK")
            if cw:
                # Rotate counterweight to correct angular position
                try:
                    self._do(cw.Rotate3D,
                              self._pt(cx, cy, z_web1),
                              self._pt(cx, cy, z_web1 + 1),
                              ang + math.pi)
                except Exception: pass
                self._union(shaft, cw)

            try: pythoncom.PumpWaitingMessages()
            except Exception: pass

        # ── 3. Front and rear flanges ─────────────────────────────────────────
        fl_r = mj_r * 1.35; fl_h = mj_dia * 0.18
        for z_fl in [- fl_h, total_L]:
            fl = self._cyl(cx, cy, z_fl, fl_r, fl_h, "GEAR_BLANK")
            if fl: self._union(shaft, fl)

        # ── 4. Central oil drilling bore ──────────────────────────────────────
        oil = self._cyl(cx, cy, -fl_h - 2, oil_r, total_L + fl_h * 2 + 4, "GEAR_BORE")
        if oil: self._subtract(shaft, oil)

        self._log(f"      [Crank] ✔ COMPLETE — {n_cyl} throws  L={total_L:.0f}mm")
        return shaft

    # ─────────────────────────────────────────────────────────────────────────
    #  3. HEAT EXCHANGER TUBESHEET  (L&T — Nuclear / Power)
    #     P1 = shell OD mm
    #     P2 = tube OD mm             (each individual tube)
    #     P3 = tubesheet thickness mm
    #     P4 = tube pitch mm          (centre-to-centre spacing)
    #
    #  Architecture: large disc + ARRAY_GRID of tube bores
    #  + bolt hole ring + gasketed flange face.
    # ─────────────────────────────────────────────────────────────────────────

    def _heat_exchanger_tubesheet(self, cx: float, cy: float,
                                   shell_od: float, tube_od: float,
                                   thickness: float, pitch: float) -> Optional[object]:

        shell_od  = max(float(shell_od),  200.0)
        tube_od   = max(float(tube_od),   12.0)
        thickness = max(float(thickness), 20.0)
        pitch     = max(float(pitch),     tube_od * 1.25)

        shell_r   = shell_od / 2.0
        tube_r    = tube_od  / 2.0

        # ── 1. Base disc ──────────────────────────────────────────────────────
        flange_od = shell_od * 1.18
        disc = self._cyl(cx, cy, 0, flange_od / 2.0, thickness, "GEAR_SOLID")
        if disc is None: return None

        # Raised inner landing face (gasketed surface)
        land_h = thickness * 0.12
        land   = self._cyl(cx, cy, thickness, shell_r, land_h, "GEAR_BLANK")
        if land: self._union(disc, land)

        self._log(f"      [HX] OD={shell_od}  tube_od={tube_od}  pitch={pitch}  "
                  f"t={thickness}")

        # ── 2. Tube bore array — triangular pitch grid ────────────────────────
        # How many rows and columns fit inside the shell
        usable_r = shell_r * 0.88
        n_half   = max(1, int(usable_r / pitch))
        row_pitch = pitch * math.sqrt(3) / 2.0   # triangular pitch row spacing

        # Build one pristine bore template at the first position
        first_x = cx + pitch                     # offset from centre
        first_y = cy
        tmpl = self._cyl(first_x, first_y, -5.0,
                          tube_r, thickness + land_h + 10.0, "GEAR_BORE")
        if tmpl is None:
            self._log("      [!] HX: tube bore template failed"); return disc

        combined = self._do(tmpl.Copy)
        if combined is None: self._del(tmpl); return disc

        ok = 1; total_expected = 0

        for row in range(-n_half, n_half + 1):
            # Triangular pitch: odd rows offset by pitch/2
            x_offset = (pitch / 2.0) if (row % 2 != 0) else 0.0
            y_pos    = cy + row * row_pitch

            for col in range(-n_half, n_half + 1):
                x_pos = cx + col * pitch + x_offset
                # Only place tube if centre is inside the usable circle
                if math.hypot(x_pos - cx, y_pos - cy) > usable_r:
                    continue
                # Skip the template position (already in combined)
                if abs(x_pos - first_x) < 0.5 and abs(y_pos - first_y) < 0.5:
                    continue
                total_expected += 1
                clone = self._do(tmpl.Copy)
                if clone:
                    try:
                        self._do(clone.Move,
                                  self._pt(first_x, first_y, 0),
                                  self._pt(x_pos,   y_pos,   0))
                        self._union(combined, clone)
                        ok += 1
                    except Exception as e:
                        self._del(clone)
                if ok % 50 == 0:
                    try: pythoncom.PumpWaitingMessages()
                    except Exception: pass

        self._del(tmpl)
        self._subtract(disc, combined)
        self._log(f"      [HX] ✔ {ok} tube bores drilled  (triangular pitch)")

        # ── 3. Bolt hole ring ─────────────────────────────────────────────────
        n_bolts  = max(8, int(shell_od / 40) * 4)   # ~1 bolt per 40mm circumference
        bolt_r   = shell_od * 0.030
        bolt_pcd = (flange_od + shell_od) / 2.0
        bolt_tmpl = self._cyl(cx + bolt_pcd/2, cy, -5.0,
                               bolt_r, thickness + land_h + 10.0, "GEAR_BORE")
        if bolt_tmpl:
            pristine_b = self._do(bolt_tmpl.Copy)
            combined_b = self._do(pristine_b.Copy)
            b_span = 2.0 * math.pi / n_bolts
            ok_b = 1
            for i in range(1, n_bolts):
                clone = self._do(pristine_b.Copy)
                if clone:
                    try:
                        self._do(clone.Rotate3D,
                                  self._pt(cx, cy, 0), self._pt(cx, cy, 1),
                                  b_span * i)
                        self._union(combined_b, clone); ok_b += 1
                    except Exception: self._del(clone)
            self._del(pristine_b); self._del(bolt_tmpl)
            self._subtract(disc, combined_b)
            self._log(f"      [HX] ✔ {ok_b}/{n_bolts} bolt holes on Ø{bolt_pcd:.0f}mm PCD")

        self._log(f"      [HX] ✔ COMPLETE")
        return disc

    # ─────────────────────────────────────────────────────────────────────────
    #  4. CENTRIFUGAL PUMP IMPELLER  (L&T — Naval / Submarine)
    #     P1 = impeller OD mm
    #     P2 = bore diameter mm
    #     P3 = blade height mm        (axial width)
    #     P4 = number of blades
    #
    #  Architecture: hub cylinder + one backward-curved blade swept along
    #  a logarithmic spiral spline + PATTERN_CIRCULAR for all blades
    #  + front and back shroud discs.
    # ─────────────────────────────────────────────────────────────────────────

    def _pump_impeller(self, cx: float, cy: float,
                       OD: float, bore_d: float,
                       height: float, n_blades: int) -> Optional[object]:

        OD       = max(float(OD),     100.0)
        bore_d   = max(float(bore_d),  20.0)
        height   = max(float(height),  20.0)
        n_blades = max(3, int(n_blades))

        rim_r    = OD    / 2.0
        bore_r   = bore_d / 2.0
        hub_r    = max(bore_r * 1.6, bore_r + height * 0.3)
        hub_r    = min(hub_r, rim_r * 0.30)
        blade_t  = rim_r * 0.055          # blade thickness
        blade_h  = height * 0.78          # blades don't fill full axial height
        shroud_t = height * 0.11          # front/back shroud thickness

        self._log(f"      [Impeller] OD={OD}  bore={bore_d}  h={height}  "
                  f"blades={n_blades}  hub_r={hub_r:.1f}")

        # ── 1. Hub cylinder ───────────────────────────────────────────────────
        hub = self._cyl(cx, cy, 0, hub_r, height, "GEAR_SOLID")
        if hub is None: return None

        # ── 2. Back shroud ────────────────────────────────────────────────────
        back = self._cyl(cx, cy, 0, rim_r, shroud_t, "GEAR_SOLID")
        if back: self._union(hub, back)

        # ── 3. Front shroud ───────────────────────────────────────────────────
        front = self._cyl(cx, cy, height - shroud_t, rim_r, shroud_t, "GEAR_SOLID")
        if front: self._union(hub, front)

        # ── 4. One backward-curved blade (logarithmic spiral) ─────────────────
        # Blade runs from hub_r to rim_r, sweeping backward by 70°
        # We approximate the spiral with a swept spline + thin extrude_profile.
        # For COM reliability, build as a thick extrude_profile in XY (top view)
        # then treat as a constant-height solid between the shrouds.

        n_pts   = 16
        wrap    = math.radians(70.0)    # backward sweep angle
        blade_pts: List[float] = []

        # Outer edge of blade (from hub to rim, spiraling backward)
        for k in range(n_pts + 1):
            t   = k / n_pts
            r   = hub_r + (rim_r - hub_r) * t
            ang = -wrap * t               # backward sweep (negative = backward)
            blade_pts.append(cx + r * math.cos(ang))
            blade_pts.append(cy + r * math.sin(ang))

        # Inner edge (same spiral, offset tangentially by blade_t)
        for k in range(n_pts, -1, -1):
            t   = k / n_pts
            r   = hub_r + (rim_r - hub_r) * t
            ang = -wrap * t
            # Offset by blade_t perpendicular to the spiral direction
            tang_ang = ang - math.pi/2      # tangent to spiral
            blade_pts.append(cx + r * math.cos(ang) + blade_t * math.cos(tang_ang))
            blade_pts.append(cy + r * math.sin(ang) + blade_t * math.sin(tang_ang))

        blade_tmpl = self._profile_solid(blade_pts, blade_h,
                                          z=shroud_t, layer="GEAR_TEETH")
        if blade_tmpl is None:
            self._log("      [!] Impeller: blade profile_solid failed — using simple box")
            blade_tmpl = self._box(cx + hub_r, cy - blade_t/2, shroud_t,
                                    rim_r - hub_r, blade_t, blade_h, "GEAR_TEETH")

        if blade_tmpl is not None:
            pristine_bl = self._do(blade_tmpl.Copy)
            combined_bl = self._do(pristine_bl.Copy)
            b_span      = 2.0 * math.pi / n_blades
            ok_b        = 1
            for i in range(1, n_blades):
                clone = self._do(pristine_bl.Copy)
                if clone:
                    try:
                        self._do(clone.Rotate3D,
                                  self._pt(cx, cy, 0), self._pt(cx, cy, 1),
                                  b_span * i)
                        self._union(combined_bl, clone); ok_b += 1
                    except Exception: self._del(clone)
                if i % 4 == 0:
                    try: pythoncom.PumpWaitingMessages()
                    except Exception: pass
            self._del(pristine_bl); self._del(blade_tmpl)
            self._union(hub, combined_bl)
            self._log(f"      [Impeller] ✔ {ok_b}/{n_blades} blades added")

        # ── 5. Shaft bore ─────────────────────────────────────────────────────
        bore_cyl = self._cyl(cx, cy, -5.0, bore_r, height + 10.0, "GEAR_BORE")
        if bore_cyl: self._subtract(hub, bore_cyl)

        # Keyway
        kw_b = bore_d / 4.0; kw_t = bore_d / 8.0
        kw = self._box(cx - kw_b/2, cy + bore_r - kw_t, -5.0,
                        kw_b, kw_t + bore_r, height + 10.0, "GEAR_BORE")
        if kw: self._subtract(hub, kw)

        self._log(f"      [Impeller] ✔ COMPLETE — OD={OD}  {n_blades} blades")
        return hub

    # ─────────────────────────────────────────────────────────────────────────
    #  5. ROCKET MOTOR CASING  (L&T — ISRO / Defence)
    #     P1 = casing OD mm
    #     P2 = wall thickness mm      (structural shell wall)
    #     P3 = cylinder length mm     (straight body section)
    #     P4 = dome height mm         (hemispherical / elliptical end cap)
    #
    #  Architecture: revolve the outer profile (cylinder + dome),
    #  subtract the inner bore (shell = outer - inner with wall thickness),
    #  add mounting flanges and nozzle throat boss.
    # ─────────────────────────────────────────────────────────────────────────

    def _rocket_casing(self, cx: float, cy: float,
                       OD: float, wall_t: float,
                       length: float, dome_h: float) -> Optional[object]:

        OD      = max(float(OD),     150.0)
        wall_t  = max(float(wall_t),   3.0)
        length  = max(float(length),  200.0)
        dome_h  = max(float(dome_h),   OD * 0.35)

        r_outer = OD / 2.0
        r_inner = r_outer - wall_t

        self._log(f"      [Rocket] OD={OD}  wall={wall_t}  L={length}  "
                  f"dome={dome_h:.1f}  r_inner={r_inner:.1f}")

        # ── 1. Outer shell — revolve profile around Z axis ────────────────────
        # Profile (X=radial, Y=axial, revolve around Y axis of revolution):
        #   base flange → straight cylinder → dome shoulder → dome cap
        flange_h   = wall_t * 2.5
        flange_ext = wall_t * 1.8        # extra radial extent for flange
        dome_r     = (r_outer**2 + dome_h**2) / (2.0 * dome_h)   # sphere radius for cap

        # Build dome top as ellipsoid approximation via N-point arc
        n_dome = 20
        dome_profile: List[Tuple[float,float]] = []

        # Start at outer radius, base of cylinder
        dome_profile.append((0.0,                  0.0))          # axis bottom
        dome_profile.append((r_outer + flange_ext, 0.0))          # flange OD
        dome_profile.append((r_outer + flange_ext, flange_h))     # flange top
        dome_profile.append((r_outer,              flange_h))     # into cylinder
        dome_profile.append((r_outer,              flange_h + length))   # cylinder top

        # Dome arc from cylinder OD down to axis
        z_dome_base = flange_h + length
        for k in range(n_dome + 1):
            theta = math.pi/2 * k / n_dome        # 0 → π/2
            dome_profile.append((
                r_outer * math.cos(theta),
                z_dome_base + dome_h * math.sin(theta)
            ))
        dome_profile.append((0.0, z_dome_base + dome_h))     # apex, close to axis

        # Flatten for revolve
        outer_flat: List[float] = []
        for px, py in dome_profile:
            outer_flat.append(cx + px)
            outer_flat.append(cy + py)

        casing = self._revolve_profile(
            outer_flat, 0.0,
            (cx, cy, 0), (0, 0, 1),
            360.0, "GEAR_SOLID"
        )

        if casing is None:
            # Fallback: cylinder + frustum dome
            self._log("      [!] Rocket revolve failed — using cylinder fallback")
            casing = self._cyl(cx, cy, 0, r_outer, length + dome_h, "GEAR_SOLID")
            if casing is None: return None

        # ── 2. Inner bore — subtract hollow ───────────────────────────────────
        # Inner has same dome shape but with wall_t subtracted throughout
        inner_dome: List[Tuple[float,float]] = []
        inner_dome.append((0.0,     0.0))
        inner_dome.append((r_inner, 0.0))
        inner_dome.append((r_inner, flange_h + length))
        z_dome_base = flange_h + length
        for k in range(n_dome + 1):
            theta = math.pi/2 * k / n_dome
            inner_dome.append((
                r_inner * math.cos(theta),
                z_dome_base + (dome_h - wall_t) * math.sin(theta)
            ))
        inner_dome.append((0.0, z_dome_base + dome_h - wall_t))

        inner_flat: List[float] = []
        for px, py in inner_dome:
            inner_flat.append(cx + px)
            inner_flat.append(cy + py)

        inner_solid = self._revolve_profile(
            inner_flat, 0.0,
            (cx, cy, 0), (0, 0, 1),
            360.0, "GEAR_SOLID"
        )
        if inner_solid:
            self._subtract(casing, inner_solid)
            self._log(f"      [Rocket] ✔ Shell wall={wall_t}mm")
        else:
            # Fallback: simple bore
            bore = self._cyl(cx, cy, -5.0, r_inner, length + dome_h + 10.0, "GEAR_SOLID")
            if bore: self._subtract(casing, bore)

        # ── 3. Nozzle throat boss (aft end) ───────────────────────────────────
        nozzle_r   = r_outer * 0.38
        nozzle_h   = dome_h  * 0.55
        nozzle_bore = nozzle_r * 0.52
        throat = self._cyl(cx, cy, -(nozzle_h), nozzle_r, nozzle_h, "GEAR_BLANK")
        if throat: self._union(casing, throat)
        throat_bore = self._cyl(cx, cy, -(nozzle_h) - 5, nozzle_bore,
                                  nozzle_h + length + dome_h + 10.0, "GEAR_BORE")
        if throat_bore: self._subtract(casing, throat_bore)

        # ── 4. Forward mounting ring ──────────────────────────────────────────
        fwd_ring_r  = r_outer + flange_ext + wall_t * 0.5
        fwd_ring_h  = wall_t * 1.6
        fwd_z       = flange_h + length - fwd_ring_h / 2.0
        fwd_ring    = self._cyl(cx, cy, fwd_z, fwd_ring_r, fwd_ring_h, "GEAR_BLANK")
        if fwd_ring: self._union(casing, fwd_ring)

        # Bolt holes in fwd ring
        n_bolts  = 12
        bolt_r   = wall_t * 0.8
        bolt_pcd = r_outer + flange_ext * 0.6
        bolt_tmpl = self._cyl(cx + bolt_pcd, cy, fwd_z - 5.0,
                               bolt_r, fwd_ring_h + 10.0, "GEAR_BORE")
        if bolt_tmpl:
            p_bt   = self._do(bolt_tmpl.Copy)
            c_bt   = self._do(p_bt.Copy)
            bspan  = 2.0 * math.pi / n_bolts; ok_bt = 1
            for i in range(1, n_bolts):
                clone = self._do(p_bt.Copy)
                if clone:
                    try:
                        self._do(clone.Rotate3D,
                                  self._pt(cx,cy,fwd_z), self._pt(cx,cy,fwd_z+1),
                                  bspan * i)
                        self._union(c_bt, clone); ok_bt += 1
                    except Exception: self._del(clone)
            self._del(p_bt); self._del(bolt_tmpl)
            self._subtract(casing, c_bt)
            self._log(f"      [Rocket] ✔ {ok_bt} forward mount bolts")

        self._log(f"      [Rocket] ✔ COMPLETE — "
                  f"OD={OD}  wall={wall_t}  L={length}  dome={dome_h:.0f}")
        return casing

    # ══════════════════════════════════════════════════════════════════════════
    #  NON-GEAR SOLIDS (original v6.2)
    # ══════════════════════════════════════════════════════════════════════════

    def _parametric_plate(self, cx, cy, L, W, H, hole_dia):
        plate = self._box(cx-L/2, cy-W/2, 0, L, W, H, "GEAR_SOLID")
        if not plate: return None
        offset = 15.0; hr = hole_dia/2.0
        if L > 40 and W > 40 and hr > 0:
            for dx in [-1, 1]:
                for dy in [-1, 1]:
                    hx = cx + dx*(L/2-offset); hy = cy + dy*(W/2-offset)
                    hole = self._cyl(hx, hy, -5.0, hr, H+10.0, "GEAR_BORE")
                    if hole: self._subtract(plate, hole)
        self._log(f"      [Plate] ✔ {L}x{W}x{H} with 4x Ø{hole_dia} holes")
        return plate

    def _solid_flange(self, cx, cy, od, id_bore, thk, n_holes):
        flange = self._cyl(cx, cy, 0, od/2.0, thk, "GEAR_SOLID")
        if not flange: return None
        if id_bore > 0:
            bore = self._cyl(cx, cy, -5.0, id_bore/2.0, thk+10.0, "GEAR_BORE")
            if bore: self._subtract(flange, bore)
        n_holes = int(n_holes)
        if n_holes > 0:
            pcd = (od + id_bore)/2.0
            hole_r = min(10.0, (od - pcd)*0.35)
            if hole_r > 1.0:
                span = 2*math.pi/n_holes; ok = 0
                for i in range(n_holes):
                    ang = i*span
                    hx = cx+(pcd/2.0)*math.cos(ang); hy = cy+(pcd/2.0)*math.sin(ang)
                    hc = self._cyl(hx, hy, -5.0, hole_r, thk+10.0, "GEAR_BORE")
                    if hc: self._subtract(flange, hc); ok += 1
                self._log(f"      [Flange] ✔ {ok}/{n_holes} holes on Ø{pcd:.1f} BHC")
        return flange

    def _solid_stepped_shaft(self, cx, cy, d1, length, d2, l2):
        shaft = self._cyl(cx, cy, 0, d1/2.0, length, "GEAR_SOLID")
        if not shaft: return None
        if d2 > 0 and l2 > 0:
            step_top = self._cyl(cx, cy, length, d2/2.0, l2, "GEAR_SOLID")
            if step_top: self._union(shaft, step_top)
            step_bot = self._cyl(cx, cy, -l2, d2/2.0, l2, "GEAR_SOLID")
            if step_bot: self._union(shaft, step_bot)
        kw_w = d1/4.0; kw_d = d1/8.0
        kw = self._box(cx-kw_w/2, cy+d1/2.0-kw_d, length*0.2, kw_w, d1/2.0, length*0.6, "GEAR_BORE")
        if kw: self._subtract(shaft, kw)
        self._log("      [Shaft] ✔ Stepped ends + keyway")
        return shaft

    def _solid_l_bracket(self, cx, cy, L, W, H, T):
        base = self._box(cx-L/2, cy-W/2, 0, L, W, T, "GEAR_SOLID")
        if not base: return None
        upright = self._box(cx-L/2, cy-W/2, T, T, W, H-T, "GEAR_SOLID")
        if upright: self._union(base, upright)
        hole_r = max(T*0.4, 3.0); offset = T+hole_r+4.0
        if L > offset*3 and W > offset*3:
            for dy in [1, -1]:
                hx = cx+L/2-offset; hy = cy+dy*(W/2-offset)
                hc = self._cyl(hx, hy, -5.0, hole_r, T+10.0, "GEAR_BORE")
                if hc: self._subtract(base, hc)
            for dy in [1, -1]:
                hx = cx-L/2-5.0; hy = cy+dy*(W/2-offset); hz = H-offset
                hc = self._cyl(hx, hy, hz, hole_r, T+10.0, "GEAR_BORE")
                if hc:
                    self._do(hc.Rotate3D, self._pt(hx,hy,hz),
                             self._pt(hx,hy+1,hz), math.pi/2)
                    self._subtract(base, hc)
        self._log(f"      [Bracket] ✔ {L}x{W}x{H}")
        return base

    # ══════════════════════════════════════════════════════════════════════════
    #  DISPATCH  (v7.0 — tolerant template lookup)
    # ══════════════════════════════════════════════════════════════════════════

    def _dispatch(self, cx, cy, ptype, p1, p2, p3, p4):
        Z,m,fw,bd = p1,p2,p3,p4

        # ── 1. Named built-in types ────────────────────────────────────────
        dispatch_map = {
            "Spur_Gear_3D":   lambda: self._gear_spur(cx,cy,int(Z),m,fw,bd),
            "Helical_Gear":   lambda: self._gear_helical(cx,cy,int(Z),m,fw,bd),
            "Ring_Gear_3D":   lambda: self._gear_ring(cx,cy,int(Z),m,fw,bd),
            "Bevel_Gear":     lambda: self._gear_bevel(cx,cy,int(Z),m,fw,bd),
            "Worm":           lambda: self._gear_worm(cx,cy,int(Z),m,fw,bd),
            "Worm_Wheel":     lambda: self._gear_worm_wheel(cx,cy,int(Z),m,fw,bd),
            "Box":            lambda: self._solid_box(cx,cy,Z,m,fw),
            "Cylinder":       lambda: self._solid_cylinder(cx,cy,Z,m,fw),
            "Sphere":         lambda: self._solid_sphere(cx,cy,Z),
            "Cone":           lambda: self._solid_cone(cx,cy,Z,fw),
            "Mounting_Plate": lambda: self._parametric_plate(cx,cy,Z,m,fw,bd),
            "Flange":         lambda: self._solid_flange(cx,cy,p1,p2,p3,p4),
            "Stepped_Shaft":  lambda: self._solid_stepped_shaft(cx,cy,p1,p2,p3,p4),
            "L_Bracket":      lambda: self._solid_l_bracket(cx,cy,p1,p2,p3,p4),
            "Flanged_Boss":   lambda: self._solid_cylinder(cx,cy,max(Z,m),0,fw),
            "Extruded_Profile":lambda: self._solid_cylinder(cx,cy,max(Z,m),0,fw),
            "Revolved_Part":  lambda: self._solid_cylinder(cx,cy,max(Z,m),0,fw),
            # ── Industry components ──────────────────────────────────────────
            "Turbine_Disc":        lambda: self._turbine_disc(cx,cy,p1,p2,p3,int(p4)),
            "Turbine_Stage":       lambda: self._turbine_stage(cx,cy,p1,p2,p3,int(p4)),
            "Turbine_Blade":       lambda: self._turbine_blade(cx,cy,p1,p2,p3,p4),
            "Crankshaft":          lambda: self._crankshaft(cx,cy,p1,p2,p3,int(p4)),
            "HX_Tubesheet":        lambda: self._heat_exchanger_tubesheet(cx,cy,p1,p2,p3,p4),
            "Pump_Impeller":       lambda: self._pump_impeller(cx,cy,p1,p2,p3,int(p4)),
            "Rocket_Casing":       lambda: self._rocket_casing(cx,cy,p1,p2,p3,p4),
        }
        if ptype in dispatch_map:
            return dispatch_map[ptype]()

        # ── 2. JSON template lookup (tolerant naming) ──────────────────────
        templates_dir = os.path.abspath(os.path.join(os.getcwd(), "templates"))
        candidates = [
            ptype,
            f"Custom_{ptype}",
            ptype.replace("Custom_", ""),
        ]
        # Also try case-insensitive match against files in the templates dir
        if os.path.isdir(templates_dir):
            existing = {f.lower(): f for f in os.listdir(templates_dir)
                        if f.endswith(".json")}
            for c in list(candidates):
                key = f"{c}.json".lower()
                if key in existing:
                    candidates.insert(0, existing[key][:-5])  # strip .json
                    break

        for candidate in candidates:
            template_path = os.path.join(templates_dir, f"{candidate}.json")
            if os.path.exists(template_path):
                try:
                    with open(template_path, "r") as f:
                        recipe = json.load(f)
                    self._log(f"      [Template] Loading '{candidate}.json'")
                    return self._build_from_recipe(cx, cy, Z, m, fw, bd,
                                                    recipe.get("Steps", []))
                except Exception as e:
                    self._log(f"      [!] Template error '{candidate}.json': {e}")
                    return None

        self._log(f"      [!] No builder found for type '{ptype}'")
        return None

    # ══════════════════════════════════════════════════════════════════════════
    #  ERP, LAYOUT, DXF, BATCH  (unchanged from v6.2)
    # ══════════════════════════════════════════════════════════════════════════

    def _erp(self, ptype, p1, p2, p3, p4, mat):
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
            v=pi*((p1/2)**2-(p2/2)**2)*p3*0.90
        elif ptype=="Stepped_Shaft":
            v=pi*(p1/2)**2*p3+2*pi*(p2/2)**2*p4
        elif ptype=="L_Bracket":
            v=(p1*p2*p4)+(p4*p2*(p3-p4))
        elif ptype=="Turbine_Disc":
            v = math.pi*(p1/2)**2*p3*0.72
        elif ptype=="Turbine_Stage":
            # disc + n_blades
            disc_v  = math.pi*(p1/2)**2*p3*0.72
            avg_c   = (p3*1.30 + p3*1.30*0.55)/2
            blade_v = avg_c * avg_c * 0.10 * (p1/2*0.78) * 0.85
            v = disc_v + blade_v * int(p4)
        elif ptype=="Turbine_Blade":
            # Volume ≈ average chord × 10% thickness × span
            avg_chord = (p2 + p3) / 2.0
            v = avg_chord * avg_chord * 0.10 * p1 * 0.85
        elif ptype=="Crankshaft":
            # Approximate: main journals + rod journals + webs
            mj_r = p1/2; rj_r = p3/2; n = int(p4)
            unit = p1*0.85 + p1*0.55
            total_l = unit * (n + 1)
            v = math.pi*mj_r**2*total_l*0.65
        elif ptype=="HX_Tubesheet":
            # Disc volume minus tube bores
            disc_v = math.pi*(p1*0.59)**2*p3
            usable_r = p1/2*0.88
            n_tubes = int(math.pi*usable_r**2 / (p4**2))
            tube_v  = math.pi*(p2/2)**2*p3 * n_tubes
            v = max(disc_v - tube_v, disc_v * 0.40)
        elif ptype=="Pump_Impeller":
            v = math.pi*(p1/2)**2*p3*0.55
        elif ptype=="Rocket_Casing":
            r_o = p1/2; r_i = r_o - p2
            v = math.pi*(r_o**2 - r_i**2)*p3 + (2/3)*math.pi*r_o**2*p4*0.82
        else: v=pi*p1**2*p3
        db=MATERIAL_DB.get(mat,MATERIAL_DB["Steel-4140"])
        mass=round(max(v,0)*db["density"]/1e6,3)
        cost=round(mass*db["cost_per_kg"],2)
        return max(v,0),mass,cost

    def _make_layout(self, part, mass, cost, vol):
        pno=str(part.get("Part_Number","PART")); name=f"DRW_{pno}"[:31]
        try:    layout=self.doc.Layouts.Add(name)
        except Exception:
            try:    layout=self.doc.Layouts.Item(name)
            except Exception: return ""
        try:
            self.doc.ActiveLayout=layout; ps=self.doc.PaperSpace
        except Exception: return name

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
        tx(f"TOLERANCE: +/-0.05 mm (Unless specified)",c1+5,r2-dh,3.5)
        tx(f"SIRAAL ENGINE v7.0  |  AI Shape Compiler",c1+5,r2+dh,4.0)
        try: self.doc.ActiveLayout=self.doc.Layouts.Item("Model")
        except Exception: pass
        return name

    def _export_dxf(self, pno, solid, orig_cx, orig_cy, out_dir):
        if solid is None: return
        os.makedirs(out_dir, exist_ok=True)
        orig_doc = self.doc; orig_ms = self.ms; tmp_doc = None
        try:
            time.sleep(0.2)
            tmp_doc = self.acad.Documents.Add()
            tmp_ms  = win32com.client.dynamic.Dispatch(tmp_doc.ModelSpace)
            obj_array = self._obj_arr([solid])
            self._do(orig_doc.CopyObjects, obj_array, tmp_ms)
            pt_current = self._pt(orig_cx, orig_cy, 0.0)
            pt_origin  = self._pt(0.0, 0.0, 0.0)
            for obj in tmp_ms:
                self._do(obj.Move, pt_current, pt_origin)
            path = os.path.abspath(os.path.join(out_dir, f"{pno}_3D.dxf"))
            if os.path.exists(path): os.remove(path)
            time.sleep(0.2)
            self._do(tmp_doc.SaveAs, path, DXF_FORMAT)
            self._log(f"          [DXF] ✔ {pno}_3D.dxf")
        except Exception as e:
            self._log(f"          [DXF] {pno}: {e}")
        finally:
            self.doc = orig_doc; self.ms = orig_ms
            try:
                if tmp_doc: self._do(tmp_doc.Close, False); time.sleep(0.5)
            except Exception: pass

    def _footprint(self, ptype, p1, p2, p3, p4):
        Z,m,fw,bd=p1,p2,p3,p4
        if ptype in("Spur_Gear_3D","Helical_Gear","Worm_Wheel"):
            x=profile_shift_x(int(Z)); return (Z*m/2+m*(1+x))*2, fw
        elif ptype=="Ring_Gear_3D": return (Z*m/2-m+bd)*2, fw
        elif ptype=="Bevel_Gear":   return (Z*m/2+m)*2, fw
        elif ptype=="Worm":
            sr=bd/2+m*1.5 if bd>1 else m*2.5; return (sr+m*0.5)*2, fw
        elif ptype=="Box": return Z, fw
        elif ptype=="Sphere": return Z*2, Z*2
        elif ptype=="Flange":       return p1+20, p3
        elif ptype=="Stepped_Shaft":return max(p1,p3)+20, p2+2*p4
        elif ptype=="L_Bracket":    return p1+20, max(p2,p3)
        elif ptype=="Turbine_Disc":   return p1+20, p3
        elif ptype=="Turbine_Stage":  return p1 + 2*(p1/2*0.78) + 20, p3 + p1/2*0.78
        elif ptype=="Turbine_Blade":  return p2+20, p1       # p2=root_chord, p1=span
        elif ptype=="Crankshaft":
            n = int(p4); unit = p1*0.85 + p1*0.55
            return (p1*1.35 + p2)*2 + 40, unit*(n+1)
        elif ptype=="HX_Tubesheet":   return p1*1.18+40, p3
        elif ptype=="Pump_Impeller":  return p1+20, p3
        elif ptype=="Rocket_Casing":  return p1*1.25+40, p3+p4
        else: return max(Z,m)*2+20, fw+20

    def generate_3d_batch(self, parts: List[dict]):
        GAP=90.0; cx=0.0; layouts=[]; failed=[]
        session_id  = time.strftime("%Y%m%d_%H%M%S")
        session_dir = os.path.abspath(
            os.path.join(os.getcwd(), "output_3d", f"Session_{session_id}"))
        os.makedirs(session_dir, exist_ok=True)
        SEP="═"*62
        self._log(f"\n{SEP}")
        self._log(f"  SIRAAL ENGINE v7.0 — BATCH: {len(parts)} parts")
        self._log(f"  SESSION DIR: {session_dir}")
        self._log(f"{SEP}\n")

        for idx, part in enumerate(parts, 1):
            pno  = str(part.get("Part_Number", f"P{idx:03d}")).strip()
            ptype= str(part.get("Part_Type",   "Spur_Gear_3D")).strip()
            mat  = str(part.get("Material",    "Steel-4140")).strip()
            part_dir = os.path.join(session_dir, pno)
            os.makedirs(part_dir, exist_ok=True)
            try:
                p1=float(part.get("Param_1",20)); p2=float(part.get("Param_2",3))
                p3=float(part.get("Param_3",30)); p4=float(part.get("Param_4",20))
            except Exception:
                self._log(f"  [!] {pno}: bad params — skip"); failed.append(pno); continue

            self._log(f"[{idx:02d}/{len(parts)}] ▶ {pno}  {ptype}  {mat}")
            try:
                vol,mass,cost = self._erp(ptype,p1,p2,p3,p4,mat)
                orig_cx = cx
                solid   = self._dispatch(orig_cx, 0.0, ptype, p1, p2, p3, p4)
                if solid is None:
                    self._log(f"          [!] no solid"); failed.append(pno); continue
                self._mat_color(solid, mat)
                fw,_ = self._footprint(ptype,p1,p2,p3,p4); cx += fw+GAP
                try:
                    ln = self._make_layout(part,mass,cost,vol)
                    if ln: layouts.append(ln); self._log(f"          [Layout] ✔ {ln}")
                except Exception as e: self._log(f"          [Layout] {e}")
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

        master = os.path.join(session_dir, "Master_3D_Assembly.dwg")
        for fmt in DWG_SAVE_FORMATS:
            try:
                self.doc.SaveAs(master, fmt)
                self._log(f"\n[+] Master DWG saved: {master}"); break
            except Exception: continue

        self._log(f"\n{SEP}")
        self._log(f"  DONE  built={len(parts)-len(failed)}/{len(parts)}"
                  f"  layouts={len(layouts)}")
        if failed: self._log(f"  failed={failed}")
        self._log(f"{SEP}\n")


# ══════════════════════════════════════════════════════════════════════════════
#  Excel BOM loader
# ══════════════════════════════════════════════════════════════════════════════

def load_bom_from_excel(xlsx_path: str) -> List[dict]:
    try: import pandas as pd
    except ImportError: raise ImportError("pip install pandas openpyxl")
    df = pd.read_excel(xlsx_path, sheet_name="BOM_Gears", header=2, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    def g(row, keys, default=""):
        ks = [keys] if isinstance(keys, str) else keys
        for k in ks:
            v = row.get(k)
            if v is not None and str(v).strip() not in ("","nan"): return str(v).strip()
        return default
    parts = []
    for _, row in df.iterrows():
        en = g(row,"Enabled","YES").upper()
        if en not in ("YES","Y","1","TRUE"): continue
        pno = g(row,"Part_Number")
        if not pno or pno.lower() in ("nan","part_number"): continue
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


if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s  %(message)s", datefmt="%H:%M:%S")
    if len(sys.argv) > 1:
        parts = load_bom_from_excel(sys.argv[1])
        print(f"Loaded {len(parts)} parts from {sys.argv[1]}")
    else:
        parts = [
            {"Part_Number":"GR-001-SPUR","Part_Type":"Spur_Gear_3D",
             "Material":"Steel-4140","Param_1":"24","Param_2":"3",
             "Param_3":"30","Param_4":"20","Quantity":"1","Priority":"High"},
            {"Part_Number":"GR-002-HELICAL","Part_Type":"Helical_Gear",
             "Material":"Steel-4140","Param_1":"30","Param_2":"3",
             "Param_3":"40","Param_4":"22","Quantity":"1","Priority":"High"},
        ]
    AutoCAD3DGearEngine().generate_3d_batch(parts)