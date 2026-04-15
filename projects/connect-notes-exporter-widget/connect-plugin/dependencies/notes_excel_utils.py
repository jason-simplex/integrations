# :coding: utf-8
import os
import re
import tempfile
import hashlib
from datetime import datetime, timezone
from typing import Iterable, Optional, Tuple

try:
    from zoneinfo import ZoneInfo  # py>=3.9
except Exception:  # pragma: no cover
    ZoneInfo = None

INVALID_SHEET_CHARS_RE = re.compile(r"[:\\\\/?*\\[\\]]")


def sanitize_sheet_name(name: str, suffix: str = "", max_len: int = 31) -> str:
    base = (name or "").strip()
    base = INVALID_SHEET_CHARS_RE.sub("_", base)
    base = base.replace("'", "_")
    base = re.sub(r"\\s+", " ", base).strip()
    if suffix:
        suffix = "_" + suffix.strip("_")
    # Reserve suffix space.
    keep = max_len - len(suffix)
    if keep < 1:
        keep = 1
    base = base[:keep]
    return (base + suffix)[:max_len].strip()


def short_hash(text: str, length: int = 6) -> str:
    h = hashlib.sha1((text or "").encode("utf-8")).hexdigest()
    return h[:length]


def utc_to_local_str(value, fmt: str = "%Y-%m-%d %H:%M:%S") -> str:
    """Convert ftrack UTC datetime string -> local time string.

    ftrack dates are typically ISO strings, sometimes ending with 'Z'.
    """
    if not value:
        return ""
    if isinstance(value, datetime):
        dt = value
    else:
        s = str(value).strip()
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        try:
            dt = datetime.fromisoformat(s)
        except Exception:
            return str(value)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    local_tz = datetime.now().astimezone().tzinfo
    if local_tz is None and ZoneInfo is not None:
        local_tz = ZoneInfo("UTC")
    if local_tz is None:
        local_tz = timezone.utc
    return dt.astimezone(local_tz).strftime(fmt)


def is_image_filename(filename: str) -> bool:
    ext = os.path.splitext((filename or "").lower())[1]
    return ext in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")


def ensure_tmp_dir(prefix: str = "notes_export_") -> str:
    return tempfile.mkdtemp(prefix=prefix)


def safe_filename(name: str, max_len: int = 120) -> str:
    s = (name or "").strip()
    s = re.sub(r"[^a-zA-Z0-9._-]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    if not s:
        s = "file"
    return s[:max_len]
