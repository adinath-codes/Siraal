"""
validator_3d.py — Siraal 3D Gear Validator v2.0
Validates demo_gears_3d.xlsx (sheet: BOM_Gears) before 3D gear generation.
Supports: Spur_Gear_3D, Helical_Gear, Ring_Gear_3D, Bevel_Gear, Worm, Worm_Wheel
Also supports legacy solid types for backwards compatibility.
"""
import os
import math
import pandas as pd
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional, Callable


class Severity(Enum):
    ERROR   = "ERROR"
    WARNING = "WARNING"
    INFO    = "INFO"


@dataclass
class ValidationIssue:
    part_no:  str
    severity: Severity
    rule:     str
    message:  str


SEVERITY_ICON = {Severity.ERROR:"✘", Severity.WARNING:"⚠", Severity.INFO:"ℹ"}

MATERIAL_DB = {
    "Steel-1020", "Steel-4140", "Al-6061",
    "Brass-C360", "Nylon-66",  "Ti-6Al-4V",
}

# All gear + legacy solid types the engine can build
GEAR_TYPES = {
    "Spur_Gear_3D", "Helical_Gear", "Ring_Gear_3D",
    "Bevel_Gear", "Worm", "Worm_Wheel",
}
SOLID_TYPES = {
    "Box", "Cylinder", "Cone", "Sphere",
    "Flanged_Boss", "Extruded_Profile", "Revolved_Part",
}
OTHER_TYPES = {
   'Flange','Stepped_Shaft','L_Bracket'
}
ALL_TYPES = GEAR_TYPES | SOLID_TYPES | OTHER_TYPES

# Sheet names to try — handles both old and new Excel layouts
SHEET_CANDIDATES = ["BOM_Gears", "BOM_3D", "BOM", "Parts"]


class Validator3D:

    def __init__(self, file_path: str, log_callback: Optional[Callable] = None):
        self.file_path     = file_path
        self._log_cb       = log_callback or print
        self.issues:       List[ValidationIssue] = []
        self.valid_parts:  List[dict]            = []
        self.error_count   = 0
        self.warning_count = 0

    def _log(self, msg):
        self._log_cb(msg)

    def _add(self, pno, sev, rule, msg):
        icon = SEVERITY_ICON[sev]
        self._log(f"  {icon} [{sev.value}] {pno} | {rule}: {msg}")
        self.issues.append(ValidationIssue(pno, sev, rule, msg))
        if sev == Severity.ERROR:   self.error_count   += 1
        elif sev == Severity.WARNING: self.warning_count += 1

    def _load(self) -> Optional[pd.DataFrame]:
        """Try each sheet candidate; normalise column names (strip newlines)."""
        for sheet in SHEET_CANDIDATES:
            try:
                df = pd.read_excel(self.file_path, sheet_name=sheet,
                                   engine="openpyxl", header=2)
                # Normalise headers: strip newlines and trailing whitespace
                df.columns = [str(c).split("\n")[0].strip() for c in df.columns]

                df = df.dropna(subset=["Part_Number"])
                df = df[df["Part_Number"].astype(str).str.strip().str.len() > 0]
                df = df[~df["Part_Number"].astype(str).str.lower().str.startswith("total")]
                df = df.reset_index(drop=True)
                self._log(f"[*] Loaded {len(df)} rows from sheet '{sheet}' in '{os.path.basename(self.file_path)}'")
                return df
            except Exception:
                continue
        self._log(f"[✘] FATAL: Could not read any BOM sheet from '{self.file_path}'")
        self._log(f"    Tried sheets: {SHEET_CANDIDATES}")
        return None

    # ── Per-type geometry rules ───────────────────────────────────────────────

    def _check_spur(self, pno, Z, m, fw, bore_d):
        err = False
        pitch_r = Z * m / 2.0
        bore_r  = bore_d / 2.0
        outer_r = pitch_r + m
        root_r  = pitch_r - 1.25 * m

        if Z < 6:
            self._add(pno, Severity.ERROR, "SPUR_Z_MIN", f"Z={Z} < 6 — too few teeth, gear invalid"); err=True
        elif Z < 17:
            self._add(pno, Severity.WARNING, "SPUR_UNDERCUT", f"Z={Z} < 17 — involute undercut; profile shift recommended")
        if m <= 0:
            self._add(pno, Severity.ERROR, "SPUR_MODULE", f"module m={m} must be > 0"); err=True
        if fw <= 0:
            self._add(pno, Severity.ERROR, "SPUR_FACEWIDTH", f"FaceWidth={fw} must be > 0"); err=True
        if bore_r >= pitch_r:
            self._add(pno, Severity.ERROR, "SPUR_BORE",
                      f"bore_r={bore_r:.1f} >= pitch_r={pitch_r:.1f} — bore too large for Z×m"); err=True
        if bore_r > 0 and bore_d < 5:
            self._add(pno, Severity.WARNING, "SPUR_BORE_SMALL", f"Bore_d={bore_d}mm very small — keyway may not fit")
        if fw > 12 * m:
            self._add(pno, Severity.WARNING, "SPUR_FW_RATIO", f"FaceWidth/m={fw/m:.1f} > 12 — excessive face width")
        return err

    def _check_helical(self, pno, Z, m, fw, bore_d):
        # Helical gears use wider face widths than spur — limit is 20×m not 12×m
        err = False
        pitch_r = Z * m / 2.0
        bore_r  = bore_d / 2.0
        if Z < 6:
            self._add(pno, Severity.ERROR, "HEL_Z_MIN", f"Z={Z} < 6 — too few teeth"); err=True
        elif Z < 17:
            self._add(pno, Severity.WARNING, "HEL_UNDERCUT", f"Z={Z} < 17 — profile shift recommended")
        if m <= 0:
            self._add(pno, Severity.ERROR, "HEL_MODULE", f"module m={m} must be > 0"); err=True
        if fw <= 0:
            self._add(pno, Severity.ERROR, "HEL_FACEWIDTH", f"FaceWidth={fw} must be > 0"); err=True
        if bore_r >= pitch_r:
            self._add(pno, Severity.ERROR, "HEL_BORE",
                      f"bore_r={bore_r:.1f} >= pitch_r={pitch_r:.1f}"); err=True
        if fw > 20 * m:
            self._add(pno, Severity.WARNING, "HEL_FW_RATIO",
                      f"FaceWidth/m={fw/m:.1f} > 20 — very wide for helical gear")
        return err

    def _check_ring(self, pno, Z, m, fw, ring_thk):
        err = False
        pitch_r  = Z * m / 2.0
        inner_r  = pitch_r - m   # addendum tip on inside
        outer_r  = inner_r + ring_thk

        if Z < 20:
            self._add(pno, Severity.WARNING, "RING_Z_MIN",
                      f"Z={Z} < 20 — internal gear may have meshing issues with small planet"); 
        if m <= 0:
            self._add(pno, Severity.ERROR, "RING_MODULE", f"module m={m} must be > 0"); err=True
        if fw <= 0:
            self._add(pno, Severity.ERROR, "RING_FACEWIDTH", f"FaceWidth={fw} must be > 0"); err=True
        if ring_thk <= m:
            self._add(pno, Severity.ERROR, "RING_THICKNESS",
                      f"ring_thk={ring_thk} <= m={m} — wall too thin, disc will collapse"); err=True
        if inner_r <= 0:
            self._add(pno, Severity.ERROR, "RING_INNER_R",
                      f"inner_r=pitch_r−m={inner_r:.1f} <= 0 — impossible geometry"); err=True
        return err

    def _check_bevel(self, pno, Z, m, fw, bore_d):
        err = False
        cone_r  = math.radians(45)
        back_r  = Z * m / 2.0
        front_r = back_r - fw * math.sin(cone_r)
        height  = fw * math.cos(cone_r)
        bore_r  = bore_d / 2.0

        if Z < 6:
            self._add(pno, Severity.ERROR, "BEVEL_Z_MIN", f"Z={Z} < 6 — too few teeth"); err=True
        if m <= 0:
            self._add(pno, Severity.ERROR, "BEVEL_MODULE", f"module m={m} must be > 0"); err=True
        if front_r <= 0:
            self._add(pno, Severity.ERROR, "BEVEL_FRONTCONE",
                      f"front_r={front_r:.1f}<=0 — FaceWidth={fw} too large for pitch_r={back_r:.1f}; "
                      f"reduce FW to < {back_r/math.sin(cone_r):.1f}"); err=True
        if bore_r >= back_r:
            self._add(pno, Severity.ERROR, "BEVEL_BORE",
                      f"bore_r={bore_r:.1f} >= back_r={back_r:.1f}"); err=True
        if height <= 0:
            self._add(pno, Severity.ERROR, "BEVEL_HEIGHT", f"cone height={height:.1f}<=0"); err=True
        return err

    def _check_worm(self, pno, n_starts, m, length, bore_d):
        err = False
        bore_r  = bore_d / 2.0
        shaft_r = bore_r + m * 1.5 if bore_r > 1.0 else m * 3.0
        axp     = math.pi * m

        if n_starts < 1:
            self._add(pno, Severity.ERROR, "WORM_STARTS", f"N_starts={n_starts} must be >= 1"); err=True
        if m <= 0:
            self._add(pno, Severity.ERROR, "WORM_MODULE", f"module m={m} must be > 0"); err=True
        if length <= 0:
            self._add(pno, Severity.ERROR, "WORM_LENGTH", f"length={length} must be > 0"); err=True
        if length < axp:
            self._add(pno, Severity.WARNING, "WORM_LENGTH_SHORT",
                      f"length={length:.1f} < 1 axial pitch={axp:.1f} — only partial thread")
        if bore_r > 0 and bore_r >= shaft_r:
            self._add(pno, Severity.ERROR, "WORM_BORE",
                      f"bore_r={bore_r:.1f} >= shaft_r={shaft_r:.1f} — bore too large"); err=True
        if n_starts > 6:
            self._add(pno, Severity.WARNING, "WORM_STARTS_HIGH",
                      f"N_starts={n_starts} > 6 — high lead angle, may not be self-locking")
        return err

    def _check_worm_wheel(self, pno, Z, m, fw, bore_d):
        err = False
        pitch_r = Z * m / 2.0
        outer_r = pitch_r + m
        bore_r  = bore_d / 2.0

        if Z < 20:
            self._add(pno, Severity.WARNING, "WWHEEL_Z_MIN",
                      f"Z={Z} < 20 — small worm wheel; contact ratio low")
        if m <= 0:
            self._add(pno, Severity.ERROR, "WWHEEL_MODULE", f"module m={m} must be > 0"); err=True
        if fw <= 0:
            self._add(pno, Severity.ERROR, "WWHEEL_FACEWIDTH", f"FaceWidth={fw} must be > 0"); err=True
        if bore_r >= pitch_r:
            self._add(pno, Severity.ERROR, "WWHEEL_BORE",
                      f"bore_r={bore_r:.1f} >= pitch_r={pitch_r:.1f}"); err=True
        if fw > outer_r * 1.5:
            self._add(pno, Severity.WARNING, "WWHEEL_FW_WIDE",
                      f"FaceWidth={fw} > 1.5×outer_r={outer_r:.1f} — throat may disappear")
        return err

    # ── Legacy solid checks (unchanged) ──────────────────────────────────────

    def _check_box(self, pno, p1, p2, p3, p4):
        err = False
        if p1<=0 or p2<=0 or p3<=0:
            self._add(pno, Severity.ERROR, "BOX_DIMS", "All dims (P1 L, P2 W, P3 H) must be > 0"); err=True
        if p4>0 and p4>=min(p1,p2,p3)/2:
            self._add(pno, Severity.ERROR, "BOX_FILLET", f"fillet_R={p4} >= min_dim/2={min(p1,p2,p3)/2:.1f}"); err=True
        return err

    def _check_cylinder(self, pno, p1, p2, p3, p4):
        err = False
        if p1<=0: self._add(pno, Severity.ERROR, "CYL_OUTER_R", "Outer_R must be > 0"); err=True
        if p2>=p1: self._add(pno, Severity.ERROR, "CYL_BORE", f"Bore_R={p2} >= Outer_R={p1}"); err=True
        if p3<=0: self._add(pno, Severity.ERROR, "CYL_HEIGHT", "Height must be > 0"); err=True
        if (p1-p2)<5: self._add(pno, Severity.WARNING, "CYL_WALL", f"wall={p1-p2:.1f}mm < 5mm — thin wall")
        return err

    # ── Main run ──────────────────────────────────────────────────────────────

    def run_checks(self) -> bool:
        self._log("")
        self._log("─" * 62)
        self._log("  SIRAAL 3D GEAR VALIDATOR v2.0")
        self._log("  Supports: Spur · Helical · Ring · Bevel · Worm · Worm Wheel")
        self._log("─" * 62)

        required = {"Part_Number","Part_Type","Material",
                    "Param_1","Param_2","Param_3","Param_4","Enabled"}
        df = self._load()
        if df is None:
            return False

        missing = required - set(df.columns)
        if missing:
            self._log(f"[✘] FATAL: Missing columns: {missing}")
            self._log(f"    Columns found: {list(df.columns)}")
            return False

        self._log(f"[*] Validating {len(df)} rows...\n")

        for _, row in df.iterrows():
            pno   = str(row["Part_Number"]).strip()
            ptype = str(row.get("Part_Type","")).strip()
            mat   = str(row.get("Material","")).strip()
            ena   = str(row.get("Enabled","YES")).strip().upper()

            if ena != "YES":
                self._log(f"  ⊘ [{pno}] Skipped (Enabled={ena})")
                continue

            try:
                p1 = float(row.get("Param_1", 0) or 0)
                p2 = float(row.get("Param_2", 0) or 0)
                p3 = float(row.get("Param_3", 0) or 0)
                p4 = float(row.get("Param_4", 0) or 0)
            except Exception:
                self._add(pno, Severity.ERROR, "PARAM_PARSE",
                          "Non-numeric P1–P4 — check Excel values")
                continue

            has_error = False

            # Material check
            if mat not in MATERIAL_DB:
                self._add(pno, Severity.WARNING, "UNKNOWN_MATERIAL",
                          f"'{mat}' not in DB — Steel-1020 fallback")

            # Type check
            if ptype not in ALL_TYPES and not ptype.startswith("Custom_"):
                self._add(pno, Severity.ERROR, "UNKNOWN_TYPE",
                          f"'{ptype}' not recognised. Valid: {sorted(ALL_TYPES)}")
                continue

            # Per-type geometry rules
            if   ptype == "Spur_Gear_3D":
                has_error = self._check_spur(pno, int(p1), p2, p3, p4)
            elif ptype == "Helical_Gear":
                has_error = self._check_helical(pno, int(p1), p2, p3, p4)
            elif ptype == "Ring_Gear_3D":
                # P4 = ring thickness for Ring gear
                has_error = self._check_ring(pno, int(p1), p2, p3, p4)
            elif ptype == "Bevel_Gear":
                has_error = self._check_bevel(pno, int(p1), p2, p3, p4)
            elif ptype == "Worm":
                has_error = self._check_worm(pno, int(p1), p2, p3, p4)
            elif ptype == "Worm_Wheel":
                has_error = self._check_worm_wheel(pno, int(p1), p2, p3, p4)
            # Legacy solid types
            elif ptype == "Box":
                has_error = self._check_box(pno, p1, p2, p3, p4)
            elif ptype == "Cylinder":
                has_error = self._check_cylinder(pno, p1, p2, p3, p4)

            if not has_error:
                row_dict = row.to_dict()
                row_dict["_validated"] = True
                self.valid_parts.append(row_dict)
                self._log(f"  ✔ [{pno}] {ptype} | {mat} | QUEUED")

        self._log("")
        self._log("─" * 62)
        self._log(f"  RESULT: {len(self.valid_parts)} valid | "
                  f"{self.error_count} errors | {self.warning_count} warnings")
        self._log("─" * 62)
        return self.error_count == 0