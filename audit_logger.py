"""
audit_logger.py — Siraal Enterprise Compliance & Traceability
==============================================================
Maintains an immutable, time-stamped audit trail of all high-value 
factory actions. Required for ISO 9001 / AS9100 compliance traceability.

Upgrades: 
- RotatingFileHandler (prevents disk-fill crashes)
- Absolute path resolution
- Extended severity levels
"""

import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path

# 1. Enterprise Pathing: Always resolve to the project root, no matter where the script is run from.
BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True) # Thread-safe directory creation

AUDIT_FILE = LOG_DIR / "siraal_audit_trail.log"

def setup_audit_logger() -> logging.Logger:
    """Configures a rotating, production-safe logger for compliance tracking."""
    logger = logging.getLogger("SiraalAudit")
    logger.setLevel(logging.INFO)
    
    # Prevent duplicate handlers if imported across multiple modules
    if not logger.handlers:
        # 2. Log Rotation: Max 5MB per file, keeps the last 10 backups.
        # This prevents the server from crashing due to an oversized text file.
        fh = RotatingFileHandler(
            filename=AUDIT_FILE, 
            mode='a', 
            maxBytes=5 * 1024 * 1024, # 5 MB
            backupCount=10, 
            encoding='utf-8'
        )
        
        # Professional standard timestamp formatting
        formatter = logging.Formatter(
            fmt='[%(asctime)s] [%(levelname)-8s] | USER: %(user_role)-10s | ACTION: %(action_category)-20s | DETAILS: %(message)s', 
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        fh.setFormatter(formatter)
        logger.addHandler(fh)
        
    return logger

# Initialize the singleton logger
_audit = setup_audit_logger()

def log_event(user_role: str, action_category: str, details: str, severity: str = "INFO"):
    """
    Safely logs an action to the audit trail with enterprise formatting.
    
    Args:
        user_role (str): The role executing the action (e.g., 'ADMIN', 'SYSTEM', 'AI_ENGINE').
        action_category (str): The specific module or action type (e.g., 'RULE_OVERRIDE').
        details (str): Human-readable explanation of the event.
        severity (str): 'INFO', 'WARNING', 'ERROR', or 'CRITICAL'.
    """
    # Injecting contextual metadata into the formatter
    extra_context = {
        'user_role': user_role,
        'action_category': action_category
    }
    
    severity = severity.upper()
    
    if severity == "WARNING":
        _audit.warning(details, extra=extra_context)
    elif severity == "ERROR":
        _audit.error(details, extra=extra_context)
    elif severity == "CRITICAL":
        _audit.critical(details, extra=extra_context)
    else:
        _audit.info(details, extra=extra_context)

# Quick test if run directly
if __name__ == "__main__":
    print(f"Testing Enterprise Audit Logger. Writing to: {AUDIT_FILE}")
    log_event("SYSTEM", "BOOT", "Siraal Engine Initialized successfully.")
    log_event("ADMIN", "RULE_OVERRIDE", "Bypassed maximum face width constraint.", severity="WARNING")
    log_event("AI_AGENT", "JSON_GENERATION", "Failed to compile V8_Engine.json due to missing parameter P2.", severity="ERROR")
    print("Check the logs/siraal_audit_trail.log file!")