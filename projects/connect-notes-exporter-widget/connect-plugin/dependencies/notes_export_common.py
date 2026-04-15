# :coding: utf-8
"""Reusable helpers for ftrack note export flows."""

from __future__ import annotations

from typing import Dict, List, Optional


def esc(value: str) -> str:
    return (value or "").replace('"', '\\"')


def chunked(values: List[str], size: int) -> List[List[str]]:
    return [values[i : i + size] for i in range(0, len(values), size)]


def normalize_note_parent_type(value: str) -> str:
    """Normalize Note.parent_type to match entity type keys used by exporters."""
    s = (value or "").strip()
    if not s:
        return ""
    s = s.replace("_", "").replace(" ", "").lower()
    mapping = {
        "task": "Task",
        "assetversion": "AssetVersion",
        "typedcontext": "TypedContext",
        "context": "TypedContext",
        "project": "Project",
        "milestone": "TypedContext",
    }
    if s in mapping:
        return mapping[s]
    return s[:1].upper() + s[1:]


def ensure_single_xlsx_filename(name: str) -> str:
    """Ensure filename ends with a single .xlsx (case-insensitive)."""
    s = (name or "").strip()
    if not s:
        return "export.xlsx"
    lower = s.lower()
    while lower.endswith(".xlsx.xlsx"):
        s = s[:-5]
        lower = s.lower()
    if not lower.endswith(".xlsx"):
        s += ".xlsx"
    return s


def author_id_from_note(note: dict) -> str:
    """Return author id from a Note entity/projection without non-schema attributes."""
    try:
        author = note.get("author")
        if author and hasattr(author, "get"):
            return author.get("id") or ""
    except Exception:
        pass
    try:
        return note.get("author.id") or ""
    except Exception:
        return ""


def label_names_from_note(note: dict) -> List[str]:
    """Extract label names from note['note_label_links'] across projections."""
    out: List[str] = []
    try:
        links = note.get("note_label_links") or []
    except Exception:
        links = []

    for link in links:
        name = ""
        try:
            if hasattr(link, "get"):
                label = link.get("label")
                if label is not None:
                    if hasattr(label, "get"):
                        name = label.get("name") or ""
                    elif isinstance(label, dict):
                        name = label.get("name") or ""
                if not name:
                    name = link.get("label.name") or ""
            elif isinstance(link, dict):
                label = link.get("label")
                if isinstance(label, dict):
                    name = label.get("name") or ""
                if not name:
                    name = link.get("label.name") or ""
        except Exception:
            name = ""

        if name:
            out.append(str(name))

    seen = set()
    uniq: List[str] = []
    for name in out:
        if name not in seen:
            uniq.append(name)
            seen.add(name)
    return uniq


def get_object_type_name(entity) -> str:
    """Return ObjectType name: Shot / Asset Build / Folder / Task ..."""
    try:
        value = entity.get("object_type.name")
        if value:
            return str(value)
    except Exception:
        pass
    try:
        obj_type = entity.get("object_type")
        if obj_type:
            try:
                return str(obj_type.get("name") or obj_type["name"])
            except Exception:
                return str(obj_type)
    except Exception:
        pass
    return "TypedContext"


def get_pipeline_type_name(entity) -> str:
    """Return pipeline Type name: Character / Prop / Env ..."""
    try:
        value = entity.get("type.name")
        if value:
            return str(value)
    except Exception:
        pass
    try:
        pipeline_type = entity.get("type")
        if pipeline_type:
            try:
                return str(pipeline_type.get("name") or pipeline_type["name"])
            except Exception:
                return str(pipeline_type)
    except Exception:
        pass
    return ""


def get_task_type_name(entity) -> str:
    """Best-effort task type name (Audio/Comp/Env...)."""
    value = get_pipeline_type_name(entity)
    if value:
        return value
    try:
        value = entity.get("task_type.name")
        if value:
            return str(value)
    except Exception:
        pass
    try:
        task_type = entity.get("task_type")
        if task_type:
            try:
                return str(task_type.get("name") or task_type["name"])
            except Exception:
                return str(task_type)
    except Exception:
        pass
    return ""


def get_display_type(entity) -> str:
    """UI display style: ObjectType (TypeName)."""
    object_type = get_object_type_name(entity)
    type_name = get_pipeline_type_name(entity)
    if type_name and type_name.lower() != object_type.lower():
        return f"{object_type} ({type_name})"
    return object_type


def is_task_entity(entity) -> bool:
    try:
        return getattr(entity, "entity_type", None) == "Task"
    except Exception:
        return False


def thumbnail_url_from_value(value) -> Optional[str]:
    if not value:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, dict):
        for key in ("url", "original", "small", "medium", "large", "thumbnail"):
            url = value.get(key)
            if isinstance(url, str) and url:
                return url
        for nested in value.values():
            if isinstance(nested, dict) and isinstance(nested.get("url"), str):
                return nested.get("url")
    return None


def is_milestone_entity(entity) -> bool:
    try:
        object_type = get_object_type_name(entity) or ""
        if object_type.lower().startswith("milestone"):
            return True
    except Exception:
        pass
    try:
        return getattr(entity, "entity_type", None) == "Milestone"
    except Exception:
        return False

