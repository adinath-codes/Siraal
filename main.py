"""
main.py — Siraal Grand Unified Manufacturing Engine
CLI entry point: parses Excel/CSV → validates → generates AutoCAD batch
"""

import argparse
import logging
import os
import sys
import time

from validator      import EngineeringValidator
from autocad_engine import AutoCADController

# ── Logging setup ────────────────────────────────────────────────────────────
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"
logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)
logger = logging.getLogger("Siraal.Main")

DEMO_FILE = "excels/demo.xlsx"
BANNER = """
╔══════════════════════════════════════════════════════════════╗
║         SIRAAL GRAND UNIFIED MANUFACTURING ENGINE v3.0       ║
║              TN-IMPACT 2026  |  AutoCAD COM Pipeline         ║
╚══════════════════════════════════════════════════════════════╝
"""


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Siraal: Excel BOM → AutoCAD DWG batch generator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  uv run main.py                          # demo file auto-mode
  uv run main.py --file excels/bom.xlsx   # specific BOM
  uv run main.py --file bom.xlsx --dry-run # validate only, no AutoCAD
  uv run main.py --file bom.xlsx --report  # save validation report
        """
    )
    p.add_argument("--file",    default=None,
                   help="Path to Excel (.xlsx) or CSV BOM file")
    p.add_argument("--dry-run", action="store_true",
                   help="Run validation only — do not open AutoCAD")
    p.add_argument("--report",  action="store_true",
                   help="Save validation_report.txt alongside the BOM file")
    p.add_argument("--filter-type", default=None,
                   help="Only generate parts of this type (e.g. Spur_Gear)")
    p.add_argument("--filter-priority", default=None,
                   help="Only generate parts with this priority (High/Medium/Low)")
    p.add_argument("--no-dxf", action="store_true",
                   help="Skip per-part DXF export (faster)")
    p.add_argument("--verbose", action="store_true",
                   help="Set log level to DEBUG")
    return p


def _resolve_file(path: str | None) -> str:
    if path and os.path.exists(path):
        return path
    if path:
        logger.error(f"File not found: {path}")
        sys.exit(1)
    if os.path.exists(DEMO_FILE):
        logger.info(f"No --file given → using demo file: {DEMO_FILE}")
        return DEMO_FILE
    logger.error("No file given and no demo file found. Run: uv run temp.py")
    sys.exit(1)


def _apply_filters(parts: list, type_filter: str | None,
                   prio_filter: str | None) -> list:
    out = parts
    if type_filter:
        out = [p for p in out if p["Part_Type"] == type_filter]
        logger.info(f"Filter --type={type_filter}: {len(out)} parts retained")
    if prio_filter:
        out = [p for p in out if p.get("Priority", "Medium") == prio_filter]
        logger.info(f"Filter --priority={prio_filter}: {len(out)} parts retained")
    return out


def generate_drawing(file_path: str, args: argparse.Namespace):
    print(BANNER)
    t0 = time.perf_counter()
    sep = "─" * 64

    # ── PHASE 1: Validation ──────────────────────────────────────────────────
    logger.info(sep)
    logger.info("  PHASE 1 / 3 — ENGINEERING VALIDATION")
    logger.info(sep)

    validator = EngineeringValidator(file_path)
    ok = validator.run_checks()

    if args.report:
        report_path = os.path.join(
            os.path.dirname(file_path),
            "validation_report.txt"
        )
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(validator.summary_report())
        logger.info(f"[+] Validation report saved: {report_path}")

    if not ok:
        logger.error("PROCESS HALTED — no valid parts after validation.")
        sys.exit(2)

    # ── PHASE 2: Filter ──────────────────────────────────────────────────────
    parts = _apply_filters(
        validator.valid_parts,
        args.filter_type,
        args.filter_priority
    )
    if not parts:
        logger.error("No parts remain after filtering. Exiting.")
        sys.exit(2)

    if args.dry_run:
        logger.info("[DRY-RUN] Skipping AutoCAD — validation passed.")
        print(validator.summary_report())
        return

    # ── PHASE 3: AutoCAD Batch Generation ───────────────────────────────────
    logger.info(sep)
    logger.info("  PHASE 2 / 3 — AutoCAD COM PIPELINE")
    logger.info(sep)

    engine = AutoCADController(log_callback=logger.info)
    engine.generate_batch(parts)

    # ── PHASE 3: Summary ─────────────────────────────────────────────────────
    elapsed = time.perf_counter() - t0
    logger.info(sep)
    logger.info(f"  PHASE 3 / 3 — PIPELINE COMPLETE")
    logger.info(f"  Parts generated : {len(parts)}")
    logger.info(f"  Elapsed time    : {elapsed:.2f}s")
    logger.info(f"  Errors          : {validator.error_count}")
    logger.info(f"  Warnings        : {validator.warning_count}")
    logger.info(sep)


def main():
    parser = build_parser()
    args   = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    file_path = _resolve_file(args.file)
    generate_drawing(file_path, args)


if __name__ == "__main__":
    main()