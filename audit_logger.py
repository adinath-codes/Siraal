"""
audit_logger.py — Siraal Enterprise Compliance & Traceability
==============================================================
Maintains an immutable, time-stamped audit trail of all high-value 
factory actions (Rule overrides, AI approvals, Batch generation).
Required for ISO 9001 / AS9100 compliance traceability.
"""

import logging
import os

# Ensure a logs directory exists
os.makedirs("logs", exist_ok=True)
AUDIT_FILE = os.path.join("logs", "siraal_audit_trail.log")

def setup_audit_logger():
    logger = logging.getLogger("SiraalAudit")
    logger.setLevel(logging.INFO)
    
    # Prevent adding duplicate handlers if imported multiple times
    if not logger.handlers:
        fh = logging.FileHandler(AUDIT_FILE, mode='a', encoding='utf-8')
        # Professional standard timestamp formatting
        formatter = logging.Formatter('%(asctime)s | %(levelname)-7s | %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        fh.setFormatter(formatter)
        logger.addHandler(fh)
        
    return logger

_audit = setup_audit_logger()

def log_event(user_role: str, action_category: str, details: str, is_warning: bool = False):
    """
    Safely logs an action to the audit trail.
    Example: log_event("ADMIN", "RULE_DELETED", "Deleted rule SHOP_001")
    """
    message = f"USER: {user_role:<8} | ACTION: {action_category:<18} | DETAILS: {details}"
    
    if is_warning:
        _audit.warning(message)
    else:
        _audit.info(message)

# Quick test if run directly
if __name__ == "__main__":
    print(f"Testing Audit Logger. Writing to: {AUDIT_FILE}")
    log_event("SYSTEM", "BOOT", "Siraal Engine Initialized")
    log_event("ADMIN", "RULE_OVERRIDE", "Deleted maximum face width constraint.", is_warning=True)
    print("Check the logs/siraal_audit_trail.log file!")