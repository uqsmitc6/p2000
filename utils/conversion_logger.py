"""
Conversion history logger for UQ Slide Converter.

Stores a summary of each conversion run persistently on Azure.
Enables tracking improvement over time: did the latest code changes
reduce critical issues for a given file?

Storage: JSONL file alongside api_costs.jsonl on Azure's persistent
/home volume. Each line is one conversion run.

Usage:
    from utils.conversion_logger import log_conversion, get_conversion_history

    # After conversion + verification:
    log_conversion(report, filename="Day 3 Customer Journey.pptx")

    # In the admin panel:
    history = get_conversion_history()
    history_for_file = get_conversion_history(filename="Day 3 Customer Journey.pptx")
"""

import json
import logging
import os
from datetime import datetime, timezone
from pathlib import Path

logger = logging.getLogger("uqslide.conversion_logger")

# Reuse the same storage directory as cost_logger
_DEFAULT_DIR = "/home/data/uq-slide-converter"
_FALLBACK_DIR = "/tmp/uq-slide-converter"


def _resolve_log_dir() -> Path:
    explicit = os.environ.get("COST_LOG_DIR", "")
    if explicit:
        return Path(explicit)

    home_dir = Path(_DEFAULT_DIR)
    try:
        home_dir.mkdir(parents=True, exist_ok=True)
        test_file = home_dir / ".write_test"
        test_file.write_text("ok")
        test_file.unlink()
        return home_dir
    except (OSError, PermissionError):
        fallback = Path(_FALLBACK_DIR)
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback


LOG_DIR = _resolve_log_dir()
CONVERSION_LOG_FILE = LOG_DIR / "conversion_history.jsonl"


def log_conversion(report: dict, filename: str = "") -> dict:
    """
    Log a conversion run summary.

    Extracts key metrics and issue details from the report dict
    produced by convert_presentation().

    Returns the logged entry dict.
    """
    try:
        # Count issues by severity
        verification = report.get("verification", [])
        issues_by_severity = {"critical": 0, "major": 0, "minor": 0, "ok": 0}
        issue_details = []

        for v in verification:
            sev = v.get("severity", "unknown")
            if sev in issues_by_severity:
                issues_by_severity[sev] += 1

            if v.get("pass") is False:
                issue_details.append({
                    "slide": v.get("source_slide"),
                    "handler": v.get("handler", ""),
                    "severity": sev,
                    "issues": v.get("issues", []),
                })

        # Also capture issues from the errors list (content loss, etc.)
        errors = report.get("errors", [])

        # Count by handler — which handlers produce the most issues
        handler_issues = {}
        for detail in issue_details:
            h = detail.get("handler", "unknown")
            if h not in handler_issues:
                handler_issues[h] = 0
            handler_issues[h] += 1

        v_summary = report.get("verification_summary", {})

        entry = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "filename": filename,
            "slides_converted": report.get("slides_converted", 0),
            "slides_flagged": report.get("slides_flagged", 0),
            "slides_skipped": report.get("slides_skipped", 0),
            "api_calls": report.get("api_calls", 0),
            "verification_total": v_summary.get("total", 0),
            "verification_passed": v_summary.get("passed", 0),
            "verification_issues": v_summary.get("issues_found", 0),
            "issues_by_severity": issues_by_severity,
            "issues_by_handler": handler_issues,
            "issue_details": issue_details,
            "errors": errors,
        }

        # Write to persistent JSONL
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONVERSION_LOG_FILE, "a") as f:
            f.write(json.dumps(entry) + "\n")

        logger.info(
            "Conversion logged: %s — %d converted, %d issues (%d critical, %d major)",
            filename,
            entry["slides_converted"],
            entry["verification_issues"],
            issues_by_severity["critical"],
            issues_by_severity["major"],
        )

        return entry

    except Exception as e:
        logger.warning("Failed to log conversion: %s", e)
        return {}


def get_conversion_history(filename: str = None) -> list[dict]:
    """
    Read conversion history. Returns list of entries, newest first.

    If filename is provided, returns only entries for that file.
    """
    if not CONVERSION_LOG_FILE.exists():
        return []

    entries = []
    with open(CONVERSION_LOG_FILE, "r") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    entry = json.loads(line)
                    if filename and entry.get("filename") != filename:
                        continue
                    entries.append(entry)
                except json.JSONDecodeError:
                    continue

    entries.reverse()  # Newest first
    return entries


def get_file_progression(filename: str) -> list[dict]:
    """
    Get the issue progression for a specific file across all conversions.

    Returns a list of summaries in chronological order, showing how
    issue counts changed over time.
    """
    entries = get_conversion_history(filename=filename)
    entries.reverse()  # Chronological order

    progression = []
    for entry in entries:
        sev = entry.get("issues_by_severity", {})
        progression.append({
            "timestamp": entry.get("timestamp", "")[:19].replace("T", " "),
            "converted": entry.get("slides_converted", 0),
            "critical": sev.get("critical", 0),
            "major": sev.get("major", 0),
            "minor": sev.get("minor", 0),
            "total_issues": entry.get("verification_issues", 0),
            "errors": len(entry.get("errors", [])),
        })

    return progression


def clear_conversion_history():
    """Delete the conversion history log."""
    if CONVERSION_LOG_FILE.exists():
        CONVERSION_LOG_FILE.unlink()
        logger.info("Conversion history cleared: %s", CONVERSION_LOG_FILE)
