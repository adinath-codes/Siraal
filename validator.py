"""
validator.py — Siraal Grand Unified Manufacturing Engine
Advanced Engineering Validator: supports Plate, Spur_Gear, Stepped_Shaft, Flanged_Shaft, Ring_Gear
"""

import pandas as pd
import os
import math
import logging
from dataclasses import dataclass, field
from typing import List, Optional, Callable
from enum import Enum

logger = logging.getLogger("Siraal.Validator")


class Severity(Enum):
    INFO    = "INFO"
    WARNING = "WARNING"
    ERROR   = "ERROR"


@dataclass
class ValidationIssue:
    part_no:  str
    severity: Severity
    rule:     str
    message:  str

    def __str__(self):
        icon = {"INFO": "ℹ", "WARNING": "⚠", "ERROR": "✘"}[self.severity.value]
        return f"  {icon} [{self.severity.value}] {self.part_no} | {self.rule}: {self.message}"


# ─────────────────────────────────────────────────────────────
# TYPE-SPECIFIC RULE SETS
# ─────────────────────────────────────────────────────────────
PART_TYPE_SCHEMA = {
    "Plate": {
        "params": ["Param_1", "Param_2", "Param_3", "Param_4"],
        "labels": ["Length (mm)", "Width (mm)", "Thickness (mm)", "Hole Dia (mm)"],
        "ranges": [(10, 5000), (10, 5000), (1, 500), (1, 200)],
    },
    "Spur_Gear": {
        "params": ["Param_1", "Param_2", "Param_3", "Param_4"],
        "labels": ["Teeth Count", "Module (mm)", "Face Width (mm)", "Bore Dia (mm)"],
        "ranges": [(6, 400), (0.5, 25), (5, 500), (5, 200)],
    },
    "Stepped_Shaft": {
        "params": ["Param_1", "Param_2", "Param_3", "Param_4"],
        "labels": ["Step-1 Length (mm)", "Step-1 Dia (mm)", "Step-2 Length (mm)", "Step-2 Dia (mm)"],
        "ranges": [(10, 2000), (5, 500), (10, 2000), (5, 500)],
    },
    "Flanged_Shaft": {
        "params": ["Param_1", "Param_2", "Param_3", "Param_4"],
        "labels": ["Shaft Length (mm)", "Shaft Dia (mm)", "Flange OD (mm)", "Flange Thickness (mm)"],
        "ranges": [(20, 3000), (5, 500), (20, 1000), (3, 200)],
    },
    "Ring_Gear": {
        "params": ["Param_1", "Param_2", "Param_3", "Param_4"],
        "labels": ["Teeth Count", "Module (mm)", "Face Width (mm)", "Ring Thickness (mm)"],
        "ranges": [(12, 500), (0.5, 25), (5, 500), (3, 150)],
    },
}

VALID_MATERIALS = {"Steel-1020", "Al-6061", "Nylon-66", "Steel-4140", "Brass-C360", "Ti-6Al-4V"}
REQUIRED_COLUMNS = {"Part_Number", "Part_Type", "Material", "Param_1", "Param_2", "Param_3", "Param_4"}


# ─────────────────────────────────────────────────────────────
# RULE FUNCTIONS — each returns Optional[ValidationIssue]
# ─────────────────────────────────────────────────────────────
def _rule_param_ranges(part_no: str, part_type: str, params: dict) -> List[ValidationIssue]:
    issues = []
    schema = PART_TYPE_SCHEMA.get(part_type)
    if not schema:
        return issues
    for pkey, label, (lo, hi) in zip(schema["params"], schema["labels"], schema["ranges"]):
        val = params.get(pkey)
        if val is None:
            continue
        if not (lo <= val <= hi):
            issues.append(ValidationIssue(part_no, Severity.ERROR, "PARAM_RANGE",
                f"{label} = {val} is outside allowed range [{lo}, {hi}]"))
    return issues


def _rule_plate_structural(part_no, p1, p2, p3, p4) -> List[ValidationIssue]:
    issues = []
    shortest = min(p1, p2)
    margin = 15.0
    if p4 >= (shortest - margin):
        issues.append(ValidationIssue(part_no, Severity.ERROR, "PLATE_INTEGRITY",
            f"Hole Ø{p4}mm compromises shortest plate wall ({shortest}mm). Min clearance = {margin}mm"))
    if p3 < (p4 / 5.0):
        issues.append(ValidationIssue(part_no, Severity.WARNING, "PLATE_THICKNESS",
            f"Plate thickness {p3}mm is very thin relative to hole Ø{p4}mm — risk of deformation"))
    aspect = max(p1, p2) / min(p1, p2)
    if aspect > 10:
        issues.append(ValidationIssue(part_no, Severity.WARNING, "PLATE_ASPECT",
            f"Extreme aspect ratio {aspect:.1f}:1 — check flatness tolerance"))
    return issues


def _rule_gear_geometry(part_no, teeth, module, face_w, bore_dia) -> List[ValidationIssue]:
    issues = []
    pitch_dia = teeth * module
    root_dia  = pitch_dia - (2 * 1.25 * module)
    outer_dia = pitch_dia + (2 * module)

    if bore_dia >= root_dia * 0.85:
        issues.append(ValidationIssue(part_no, Severity.ERROR, "GEAR_BORE",
            f"Bore Ø{bore_dia}mm exceeds 85% of root dia ({root_dia:.2f}mm) — hub wall too thin"))
    if face_w < module * 8:
        issues.append(ValidationIssue(part_no, Severity.WARNING, "GEAR_FACEWIDTH",
            f"Face width {face_w}mm < 8×module ({module*8}mm) — consider increasing for load capacity"))
    if face_w > module * 16:
        issues.append(ValidationIssue(part_no, Severity.WARNING, "GEAR_FACEWIDTH",
            f"Face width {face_w}mm > 16×module — diminishing returns, axial load risk"))
    if teeth < 17:
        issues.append(ValidationIssue(part_no, Severity.INFO, "GEAR_UNDERCUTTING",
            f"Z={teeth} < 17 — standard profile may cause undercutting; consider profile shift"))
    return issues


def _rule_shaft_geometry(part_no, l1, d1, l2, d2) -> List[ValidationIssue]:
    issues = []
    step_ratio = max(d1, d2) / min(d1, d2)
    if step_ratio > 3.0:
        issues.append(ValidationIssue(part_no, Severity.WARNING, "SHAFT_STEP",
            f"Diameter step ratio {step_ratio:.2f}:1 is very large — stress concentration risk"))
    if d1 == d2:
        issues.append(ValidationIssue(part_no, Severity.INFO, "SHAFT_UNIFORM",
            "Both diameters are equal — this is a plain shaft, not a stepped shaft"))
    total_l = l1 + l2
    slenderness = total_l / max(d1, d2)
    if slenderness > 20:
        issues.append(ValidationIssue(part_no, Severity.WARNING, "SHAFT_SLENDERNESS",
            f"Slenderness ratio {slenderness:.1f} — consider intermediate bearings or reduced length"))
    return issues


def _rule_flanged_shaft(part_no, shaft_l, shaft_d, flange_od, flange_t) -> List[ValidationIssue]:
    issues = []
    if flange_od <= shaft_d:
        issues.append(ValidationIssue(part_no, Severity.ERROR, "FLANGE_OD",
            f"Flange OD {flange_od}mm must be larger than shaft Ø{shaft_d}mm"))
    flange_wall = (flange_od - shaft_d) / 2
    if flange_wall < 5:
        issues.append(ValidationIssue(part_no, Severity.ERROR, "FLANGE_WALL",
            f"Flange radial wall = {flange_wall:.1f}mm — too thin, minimum 5mm"))
    if flange_t < shaft_d * 0.15:
        issues.append(ValidationIssue(part_no, Severity.WARNING, "FLANGE_THICKNESS",
            f"Flange thickness {flange_t}mm is thin relative to shaft Ø{shaft_d}mm"))
    return issues


def _rule_ring_gear(part_no, teeth, module, face_w, ring_t) -> List[ValidationIssue]:
    issues = []
    pitch_dia = teeth * module
    if ring_t < module * 3:
        issues.append(ValidationIssue(part_no, Severity.ERROR, "RING_WALL",
            f"Ring thickness {ring_t}mm < 3×module ({module*3}mm) — insufficient for tooth root"))
    if teeth < 20:
        issues.append(ValidationIssue(part_no, Severity.WARNING, "RING_TEETH",
            f"Ring gear with Z={teeth} is unusual — verify mesh ratio with pinion"))
    return issues


# ─────────────────────────────────────────────────────────────
# MAIN VALIDATOR CLASS
# ─────────────────────────────────────────────────────────────
class EngineeringValidator:
    def __init__(self, file_path: str, log_callback: Optional[Callable] = None):
        self.file_path    = file_path
        self.valid_parts: List[dict]           = []
        self.all_issues:  List[ValidationIssue] = []
        self._log_cb = log_callback or print

    # ── public log ──────────────────────────────────────────
    def log(self, message: str):
        logger.info(message)
        self._log_cb(message)

    # ── helpers ─────────────────────────────────────────────
    def _add(self, issue: ValidationIssue):
        self.all_issues.append(issue)
        self.log(str(issue))

    def _add_many(self, issues):
        for i in issues:
            self._add(i)

    @property
    def error_count(self):
        return sum(1 for i in self.all_issues if i.severity == Severity.ERROR)

    @property
    def warning_count(self):
        return sum(1 for i in self.all_issues if i.severity == Severity.WARNING)

    # ── load ─────────────────────────────────────────────────
    def _load_dataframe(self) -> Optional[pd.DataFrame]:
        ext = os.path.splitext(self.file_path)[1].lower()
        try:
            if ext == ".csv":
                df = pd.read_csv(self.file_path)
            else:
                # Row 1 = banner, Row 2 = sub-header, Row 3 = real column headers
                # pandas header=2 means use 0-indexed row 2 (the 3rd row) as header
                df = pd.read_excel(
                    self.file_path,
                    sheet_name="BOM",
                    engine="openpyxl",
                    header=2      # ← skip banner + sub-header rows
                )
            # Drop completely empty rows and the totals row (non-numeric Part_Number)
            df = df.dropna(subset=["Part_Number"])
            df = df[df["Part_Number"].astype(str).str.strip().str.len() > 0]
            # Remove totals / label rows that have non-part data in Part_Number
            df = df[~df["Part_Number"].astype(str).str.lower().str.startswith("total")]
            df = df.reset_index(drop=True)
            self.log(f"[*] Loaded {len(df)} rows from '{os.path.basename(self.file_path)}'")
            return df
        except Exception as e:
            self.log(f"[✘] FATAL: Could not read file — {e}")
            return None

    def _check_schema(self, df: pd.DataFrame) -> bool:
        missing = REQUIRED_COLUMNS - set(df.columns)
        if missing:
            self.log(f"[✘] FATAL: Missing required columns: {missing}")
            return False
        return True

    # ── per-part validation ───────────────────────────────────
    def _validate_part(self, row: pd.Series, idx: int) -> Optional[dict]:
        part_no   = str(row.get("Part_Number", f"Row_{idx+1}")).strip()
        part_type = str(row.get("Part_Type", "Plate")).strip()
        material  = str(row.get("Material", "Steel-1020")).strip()
        qty       = int(row.get("Quantity", 1))

        # Skip intentionally disabled rows
        if str(row.get("Enabled", "YES")).strip().upper() in ("NO", "FALSE", "0", "SKIP"):
            self.log(f"  ⊘ [{part_no}] Skipped (Enabled = No)")
            return None

        # Unknown part type
        if part_type not in PART_TYPE_SCHEMA:
            self._add(ValidationIssue(part_no, Severity.ERROR, "UNKNOWN_TYPE",
                f"'{part_type}' is not a recognised part type. Valid: {list(PART_TYPE_SCHEMA.keys())}"))
            return None

        # Material
        if material not in VALID_MATERIALS:
            self._add(ValidationIssue(part_no, Severity.WARNING, "UNKNOWN_MATERIAL",
                f"'{material}' not in material DB — fallback to Steel-1020"))

        # Parse numeric params
        params = {}
        try:
            for key in ["Param_1", "Param_2", "Param_3", "Param_4"]:
                val = row[key]
                if pd.isna(val):
                    raise ValueError(f"{key} is blank")
                params[key] = float(val)
                if params[key] <= 0:
                    raise ValueError(f"{key} must be > 0 (got {params[key]})")
        except (ValueError, KeyError) as e:
            self._add(ValidationIssue(part_no, Severity.ERROR, "PARAM_PARSE", str(e)))
            return None

        p1, p2, p3, p4 = params["Param_1"], params["Param_2"], params["Param_3"], params["Param_4"]

        # Range checks
        self._add_many(_rule_param_ranges(part_no, part_type, params))

        # Type-specific structural rules
        if part_type == "Plate":
            self._add_many(_rule_plate_structural(part_no, p1, p2, p3, p4))
        elif part_type == "Spur_Gear":
            self._add_many(_rule_gear_geometry(part_no, p1, p2, p3, p4))
        elif part_type == "Stepped_Shaft":
            self._add_many(_rule_shaft_geometry(part_no, p1, p2, p3, p4))
        elif part_type == "Flanged_Shaft":
            self._add_many(_rule_flanged_shaft(part_no, p1, p2, p3, p4))
        elif part_type == "Ring_Gear":
            self._add_many(_rule_ring_gear(part_no, p1, p2, p3, p4))

        # Bail if this part has errors
        part_errors = [i for i in self.all_issues
                       if i.part_no == part_no and i.severity == Severity.ERROR]
        if part_errors:
            return None

        # ✓ Valid
        return {
            "Part_Number": part_no,
            "Part_Type":   part_type,
            "Material":    material,
            "Quantity":    qty,
            "Priority":    str(row.get("Priority", "Medium")).strip(),
            "Description": str(row.get("Description", "")).strip(),
            "Param_1": p1, "Param_2": p2, "Param_3": p3, "Param_4": p4,
        }

    # ── entry point ───────────────────────────────────────────
    def run_checks(self) -> bool:
        self.log("\n" + "─"*60)
        self.log("  SIRAAL ENGINEERING VALIDATOR v3.0")
        self.log("─"*60)

        df = self._load_dataframe()
        if df is None:
            return False

        if not self._check_schema(df):
            return False

        # Sort by Priority: High → Medium → Low
        priority_order = {"High": 0, "Medium": 1, "Low": 2}
        if "Priority" in df.columns:
            df["_prio"] = df["Priority"].map(priority_order).fillna(1)
            df = df.sort_values("_prio").drop(columns=["_prio"])

        for idx, row in df.iterrows():
            result = self._validate_part(row, idx)
            if result:
                self.valid_parts.append(result)
                self.log(f"  ✔ [{result['Part_Number']}] {result['Part_Type']} | {result['Material']} | QUEUED")

        self.log("\n" + "─"*60)
        self.log(f"  RESULT: {len(self.valid_parts)} valid | {self.error_count} errors | {self.warning_count} warnings")
        self.log("─"*60 + "\n")
        return len(self.valid_parts) > 0

    def summary_report(self) -> str:
        lines = ["=" * 60, "VALIDATION SUMMARY REPORT", "=" * 60]
        lines.append(f"File    : {self.file_path}")
        lines.append(f"Valid   : {len(self.valid_parts)} parts queued for generation")
        lines.append(f"Errors  : {self.error_count}")
        lines.append(f"Warnings: {self.warning_count}")
        if self.all_issues:
            lines.append("\nISSUES DETECTED:")
            for issue in self.all_issues:
                lines.append(str(issue))
        return "\n".join(lines)