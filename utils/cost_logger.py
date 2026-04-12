"""
API cost tracking for UQ Slide Converter.

Logs every Anthropic API call with token counts and estimated cost.
Dual storage:
  1. Local JSONL file (ephemeral on Render, good for same-session reads)
  2. Google Sheets via Apps Script webhook (persistent, survives redeploys)

The Google Sheets backend is optional — set GOOGLE_SHEETS_WEBHOOK_URL in
environment variables to enable it. If unset, only local logging is used.

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
import threading
import time
from datetime import datetime, timezone
from pathlib import Path

logger = logging.getLogger("uqslide.cost_logger")

# --- Local JSONL storage (ephemeral on Render) ---
COST_LOG_DIR = Path(os.environ.get("COST_LOG_DIR", "/tmp/uq-slide-converter"))
COST_LOG_FILE = COST_LOG_DIR / "api_costs.jsonl"

# --- Google Sheets webhook (persistent) ---
SHEETS_WEBHOOK_URL = os.environ.get("GOOGLE_SHEETS_WEBHOOK_URL", "")

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

# Cache for Sheets data (avoids hammering the API on every admin panel load)
_sheets_cache = {"data": None, "fetched_at": 0}
SHEETS_CACHE_TTL = 60  # seconds


def _ensure_log_dir():
    """Create log directory if needed."""
    COST_LOG_DIR.mkdir(parents=True, exist_ok=True)


def _post_to_sheets(entry: dict):
    """
    POST a cost entry to Google Sheets via Apps Script webhook.
    Runs in a background thread so it never blocks the conversion.
    """
    if not SHEETS_WEBHOOK_URL:
        return

    def _do_post():
        try:
            import urllib.request
            import urllib.error

            data = json.dumps(entry).encode("utf-8")
            req = urllib.request.Request(
                SHEETS_WEBHOOK_URL,
                data=data,
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            # Apps Script redirects on POST — follow it
            response = urllib.request.urlopen(req, timeout=10)
            body = response.read().decode("utf-8")

            # Apps Script returns 302 redirect; urllib follows it automatically
            # for GET but not POST. Handle the redirect manually.
            if response.status == 302:
                redirect_url = response.getheader("Location")
                if redirect_url:
                    req2 = urllib.request.Request(redirect_url, data=data,
                                                  headers={"Content-Type": "application/json"},
                                                  method="POST")
                    urllib.request.urlopen(req2, timeout=10)

            logger.debug("Sheets POST OK: %s", body[:100])
        except Exception as e:
            logger.warning("Sheets POST failed (non-blocking): %s", e)

    thread = threading.Thread(target=_do_post, daemon=True)
    thread.start()


def _fetch_from_sheets() -> list[dict]:
    """
    GET all cost entries from Google Sheets. Returns list of dicts (newest first).
    Uses a short TTL cache to avoid excessive requests.
    """
    if not SHEETS_WEBHOOK_URL:
        return []

    now = time.time()
    if _sheets_cache["data"] is not None and (now - _sheets_cache["fetched_at"]) < SHEETS_CACHE_TTL:
        return _sheets_cache["data"]

    try:
        import urllib.request
        import urllib.error

        req = urllib.request.Request(SHEETS_WEBHOOK_URL, method="GET")
        response = urllib.request.urlopen(req, timeout=15)

        # Apps Script may redirect GET requests too
        body = response.read().decode("utf-8")
        entries = json.loads(body)

        if isinstance(entries, list):
            entries.reverse()  # Newest first
            _sheets_cache["data"] = entries
            _sheets_cache["fetched_at"] = now
            logger.debug("Fetched %d entries from Sheets", len(entries))
            return entries
        else:
            logger.warning("Sheets GET returned non-list: %s", str(entries)[:100])
            return []

    except Exception as e:
        logger.warning("Sheets GET failed: %s", e)
        # Return stale cache if available
        if _sheets_cache["data"] is not None:
            return _sheets_cache["data"]
        return []


def log_api_call(
    message,
    purpose: str,
    slide_info: str = "",
    model: str = "",
    filename: str = "",
) -> dict:
    """
    Log an API call's token usage and estimated cost.

    Writes to both local JSONL (immediate) and Google Sheets (background POST).

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

        # 1. Local JSONL (immediate, ephemeral)
        _ensure_log_dir()
        with open(COST_LOG_FILE, "a") as f:
            f.write(json.dumps(entry) + "\n")

        # 2. Google Sheets (background, persistent)
        _post_to_sheets(entry)

        logger.debug(
            "API cost: %s %s — %d in / %d out tokens — $%.4f",
            purpose, slide_info, input_tokens, output_tokens, total_cost,
        )

        return entry

    except Exception as e:
        logger.warning("Failed to log API cost: %s", e)
        return {}


def get_cost_log() -> list[dict]:
    """
    Read all logged API calls. Returns list of dicts, newest first.

    Strategy:
    - If Google Sheets is configured, fetch from there (persistent, canonical)
    - Fall back to local JSONL if Sheets is unavailable or unconfigured
    """
    # Try Sheets first (persistent across deploys)
    if SHEETS_WEBHOOK_URL:
        sheets_entries = _fetch_from_sheets()
        if sheets_entries:
            return sheets_entries

    # Fall back to local JSONL
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
    """Delete the local cost log file. Called from admin panel if needed."""
    if COST_LOG_FILE.exists():
        COST_LOG_FILE.unlink()
        logger.info("Local cost log cleared")
