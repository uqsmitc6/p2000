"""
API cost tracking for UQ Slide Converter.

Logs every Anthropic API call with token counts and estimated cost.
Data is stored in a JSON Lines file that persists between sessions.

Usage:
    from utils.cost_logger import log_api_call, get_cost_summary

    # After an API call:
    log_api_call(
        message=response,           # anthropic Message object
        purpose="classification",   # or "verification"
        slide_info="Slide 5/34",
        model="claude-sonnet-4-6",
    )

    # In the admin panel:
    summary = get_cost_summary()
"""

import json
import logging
import os
import time
from datetime import datetime, timezone
from pathlib import Path

logger = logging.getLogger("uqslide.cost_logger")

# Persistent log file — stored alongside the app
COST_LOG_DIR = Path(os.environ.get("COST_LOG_DIR", "/tmp/uq-slide-converter"))
COST_LOG_FILE = COST_LOG_DIR / "api_costs.jsonl"

# Pricing per million tokens (USD) — updated April 2026
# https://docs.anthropic.com/en/docs/about-claude/pricing
MODEL_PRICING = {
    "claude-sonnet-4-6": {"input": 3.00, "output": 15.00},
    "claude-sonnet-4-5-20250514": {"input": 3.00, "output": 15.00},
    "claude-haiku-4-5-20251001": {"input": 0.80, "output": 4.00},
    "claude-opus-4-6": {"input": 15.00, "output": 75.00},
}

# Default fallback pricing if model not recognised
DEFAULT_PRICING = {"input": 3.00, "output": 15.00}


def _ensure_log_dir():
    """Create log directory if needed."""
    COST_LOG_DIR.mkdir(parents=True, exist_ok=True)


def log_api_call(
    message,
    purpose: str,
    slide_info: str = "",
    model: str = "",
    filename: str = "",
) -> dict:
    """
    Log an API call's token usage and estimated cost.

    Args:
        message: anthropic.types.Message object (has .usage attribute)
        purpose: "classification" or "verification"
        slide_info: e.g. "Slide 5/34"
        model: model string used for the call
        filename: source PPTX filename being processed

    Returns:
        dict with the logged entry (for immediate display if needed)
    """
    try:
        usage = message.usage
        input_tokens = usage.input_tokens
        output_tokens = usage.output_tokens

        # Look up pricing
        pricing = MODEL_PRICING.get(model, DEFAULT_PRICING)
        input_cost = (input_tokens / 1_000_000) * pricing["input"]
        output_cost = (output_tokens / 1_000_000) * pricing["output"]
        total_cost = input_cost + output_cost

        entry = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "model": model or message.model,
            "purpose": purpose,
            "slide_info": slide_info,
            "filename": filename,
            "input_tokens": input_tokens,
            "output_tokens": output_tokens,
            "total_tokens": input_tokens + output_tokens,
            "input_cost_usd": round(input_cost, 6),
            "output_cost_usd": round(output_cost, 6),
            "total_cost_usd": round(total_cost, 6),
        }

        # Append to log file
        _ensure_log_dir()
        with open(COST_LOG_FILE, "a") as f:
            f.write(json.dumps(entry) + "\n")

        logger.debug(
            "API cost: %s %s — %d in / %d out tokens — $%.4f",
            purpose, slide_info, input_tokens, output_tokens, total_cost,
        )

        return entry

    except Exception as e:
        logger.warning("Failed to log API cost: %s", e)
        return {}


def get_cost_log() -> list[dict]:
    """Read all logged API calls. Returns list of dicts, newest first."""
    if not COST_LOG_FILE.exists():
        return []

    entries = []
    with open(COST_LOG_FILE, "r") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    entries.append(json.loads(line))
                except json.JSONDecodeError:
                    continue

    entries.reverse()  # Newest first
    return entries


def get_cost_summary(entries: list[dict] = None) -> dict:
    """
    Compute aggregate cost statistics.

    Returns:
        {
            "total_calls": int,
            "total_input_tokens": int,
            "total_output_tokens": int,
            "total_cost_usd": float,
            "by_purpose": {"classification": {...}, "verification": {...}},
            "by_model": {"claude-sonnet-4-6": {...}},
            "by_date": {"2026-04-12": {...}},
        }
    """
    if entries is None:
        entries = get_cost_log()

    summary = {
        "total_calls": len(entries),
        "total_input_tokens": 0,
        "total_output_tokens": 0,
        "total_cost_usd": 0.0,
        "by_purpose": {},
        "by_model": {},
        "by_date": {},
    }

    for e in entries:
        inp = e.get("input_tokens", 0)
        out = e.get("output_tokens", 0)
        cost = e.get("total_cost_usd", 0.0)

        summary["total_input_tokens"] += inp
        summary["total_output_tokens"] += out
        summary["total_cost_usd"] += cost

        # By purpose
        purpose = e.get("purpose", "unknown")
        if purpose not in summary["by_purpose"]:
            summary["by_purpose"][purpose] = {"calls": 0, "tokens": 0, "cost_usd": 0.0}
        summary["by_purpose"][purpose]["calls"] += 1
        summary["by_purpose"][purpose]["tokens"] += inp + out
        summary["by_purpose"][purpose]["cost_usd"] += cost

        # By model
        model = e.get("model", "unknown")
        if model not in summary["by_model"]:
            summary["by_model"][model] = {"calls": 0, "tokens": 0, "cost_usd": 0.0}
        summary["by_model"][model]["calls"] += 1
        summary["by_model"][model]["tokens"] += inp + out
        summary["by_model"][model]["cost_usd"] += cost

        # By date
        ts = e.get("timestamp", "")
        date_str = ts[:10] if len(ts) >= 10 else "unknown"
        if date_str not in summary["by_date"]:
            summary["by_date"][date_str] = {"calls": 0, "tokens": 0, "cost_usd": 0.0}
        summary["by_date"][date_str]["calls"] += 1
        summary["by_date"][date_str]["tokens"] += inp + out
        summary["by_date"][date_str]["cost_usd"] += cost

    summary["total_cost_usd"] = round(summary["total_cost_usd"], 4)

    return summary


def clear_cost_log():
    """Delete the cost log file. Called from admin panel if needed."""
    if COST_LOG_FILE.exists():
        COST_LOG_FILE.unlink()
        logger.info("Cost log cleared")
