"""Shared, fence-tolerant JSON parsing for Claude responses.

Claude usually returns bare JSON, but intermittently wraps it in ```json
code fences or adds stray prose around it. The old per-agent parsers only
stripped fences sitting at the exact start/end of the string, so any
deviation raised "Invalid JSON from Claude" and failed the whole job
(e.g. product #278, Oct 23 2026 — the API call succeeded, the parse died).

This parser tries progressively more forgiving strategies and, on total
failure, logs the raw response so the failure is diagnosable instead of a
dead end.
"""
import json
import re
from typing import Any, Dict

from app.utils.logger import logger

_FENCE_OPEN = re.compile(r"^\s*```(?:json)?\s*\n?", re.IGNORECASE)
_FENCE_CLOSE = re.compile(r"\n?\s*```\s*$")


def _strip_fences(text: str) -> str:
    """Remove one leading and one trailing markdown code fence, if present."""
    text = _FENCE_OPEN.sub("", text.strip())
    text = _FENCE_CLOSE.sub("", text)
    return text.strip()


def parse_claude_json(raw_output: str) -> Dict[str, Any]:
    """Parse a Claude response into a dict, tolerating markdown fences
    and surrounding prose. Raises ValueError if no JSON object is found."""
    if not raw_output or not raw_output.strip():
        raise ValueError("Generated content is not valid JSON (empty response)")

    stripped = raw_output.strip()
    defenced = _strip_fences(stripped)

    candidates = [stripped]
    if defenced != stripped:
        candidates.append(defenced)
    # Extraction fallback: outermost {...} span — handles prose before/after
    # the JSON and fences in unexpected positions.
    start = defenced.find("{")
    end = defenced.rfind("}")
    if start != -1 and end > start:
        span = defenced[start:end + 1]
        if span not in candidates:
            candidates.append(span)

    last_error = None
    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                return parsed
            last_error = ValueError("Parsed JSON is not an object")
        except json.JSONDecodeError as e:
            last_error = e

    logger.error(f"Invalid JSON from Claude: {last_error}")
    logger.error(f"Raw Claude output (first 1000 chars): {raw_output[:1000]}")
    raise ValueError("Generated content is not valid JSON")
