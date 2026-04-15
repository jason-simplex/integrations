# :coding: utf-8
import json
import logging
import os
import pathlib
import threading
import tempfile
import urllib.request
import datetime
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from typing import Dict, List, Optional, Set, Tuple

import ftrack_api

import sys as _sys
_deps = pathlib.Path(__file__).resolve().parent.parent / "dependencies"
if str(_deps) not in _sys.path:
    _sys.path.insert(0, str(_deps))

from notes_excel_utils import (
    sanitize_sheet_name,
    short_hash,
    utc_to_local_str,
    is_image_filename,
    safe_filename,
    ensure_tmp_dir,
)

logger = logging.getLogger("connect_notes_exporter")

from ui_utils import show_info, show_warn, qta_icon_pixmap, set_button_icon_qta

from ftrack_connect.qt import QtWidgets, QtCore, QtGui
import ftrack_connect.ui.application


def _esc(s: str) -> str:
    return (s or "").replace('"', '\\"')


def _chunked(values: List[str], size: int) -> List[List[str]]:
    return [values[i : i + size] for i in range(0, len(values), size)]

def _normalize_parent_type(value: str) -> str:
    """Normalize Note.parent_type to match entity type keys used by this tool."""
    s = (value or "").strip()
    if not s:
        return ""
    s = s.replace("_", "").replace(" ", "").lower()
    mapping = {
        "task": "Task",
        "assetversion": "AssetVersion",
        "typedcontext": "TypedContext",
        "context": "TypedContext",  # notes on context are treated as TypedContext nodes in our tree
        "project": "Project",
        "milestone": "TypedContext",  # milestone behaves like a leaf context node
    }
    if s in mapping:
        return mapping[s]
    # Title-case fallback (handles unknown types reasonably)
    return s[:1].upper() + s[1:]


def _ensure_single_xlsx(name: str) -> str:
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

def _author_id_from_note(note: dict) -> str:
    """Return author id from a Note entity/projection without relying on non-schema attributes."""
    try:
        au = note.get("author")
        if au and hasattr(au, "get"):
            return au.get("id") or ""
    except Exception:
        pass
    try:
        return note.get("author.id") or ""
    except Exception:
        return ""


def _label_names_from_note(note: dict) -> List[str]:
    """Extract label names from note['note_label_links'] robustly across projections."""
    out: List[str] = []
    try:
        links = note.get("note_label_links") or []
    except Exception:
        links = []

    for lnk in links:
        name = ""
        # Best effort: link may be dict-like or an ftrack entity.
        try:
            if hasattr(lnk, "get"):
                label = lnk.get("label")
                if label is not None:
                    if hasattr(label, "get"):
                        name = label.get("name") or ""
                    elif isinstance(label, dict):
                        name = label.get("name") or ""
                if not name:
                    name = lnk.get("label.name") or ""
            elif isinstance(lnk, dict):
                label = lnk.get("label")
                if isinstance(label, dict):
                    name = label.get("name") or ""
                if not name:
                    name = lnk.get("label.name") or ""
        except Exception:
            name = ""

        if name:
            out.append(str(name))

    # De-dup while preserving order.
    seen = set()
    uniq: List[str] = []
    for n in out:
        if n not in seen:
            uniq.append(n)
            seen.add(n)
    return uniq


def _emu(px: int) -> int:
    # English Metric Units (EMU): 1 px ~= 9525 EMUs (Excel/OOXML convention)
    return int(px * 9525)


def _inject_images_into_xlsx(xlsx_path: str, placements: List[tuple]) -> int:
    """Inject images into an existing xlsx without Pillow/openpyxl image APIs.

    placements: list of (sheet_title, row_1based, col_1based, image_path)
    Returns number of images injected.
    """
    if not placements:
        return 0

    # Group by sheet title.
    # Each placement can be either:
    #   (sheet_title, row_1based, col_1based, image_path)
    # or
    #   (sheet_title, row_1based, col_1based, image_path, y_offset_px)
    by_sheet: Dict[str, List[Tuple[int, int, str, int]]] = defaultdict(list)
    for p in placements:
        if len(p) == 4:
            sheet_title, r, c, img_path = p
            y_off = 0
        else:
            sheet_title, r, c, img_path, y_off = p
        by_sheet[str(sheet_title)].append((int(r), int(c), str(img_path), int(y_off or 0)))

    tmp_out = xlsx_path + ".tmp"

    # Namespaces
    NS = {
        "w": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }

    for k, v in NS.items():
        ET.register_namespace(k, v)

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        # Map sheet title -> sheet xml path via workbook.xml
        wb_root = ET.fromstring(zin.read("xl/workbook.xml"))
        sheets_el = wb_root.find("w:sheets", NS)
        if sheets_el is None:
            return 0
        title_to_sheet_path: Dict[str, str] = {}
        wb_rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.get("Id"): rel.get("Target") for rel in wb_rels.findall("rel:Relationship", NS)}

        def _norm_target(t: str) -> str:
            t = (t or "").strip()
            if t.startswith("/"):
                t = t[1:]
            if t.startswith("xl/"):
                return t
            if t.startswith("worksheets/") or t.startswith("drawings/") or t.startswith("media/"):
                return "xl/" + t
            return "xl/" + t

        for sh in sheets_el.findall("w:sheet", NS):
            name = sh.get("name")
            rid = sh.get(f"{{{NS['r']}}}id")
            target = rid_to_target.get(rid)
            if name and target:
                title_to_sheet_path[name] = _norm_target(target)

        existing = set(zin.namelist())
        next_drawing = 1
        while f"xl/drawings/drawing{next_drawing}.xml" in existing:
            next_drawing += 1
        next_image = 1
        while any(n.startswith(f"xl/media/image{next_image}.") for n in existing):
            next_image += 1

        # Update [Content_Types].xml if needed.
        ct_root = ET.fromstring(zin.read("[Content_Types].xml"))
        def _ensure_default(ext: str, ctype: str):
            for d in ct_root.findall("ct:Default", NS):
                if (d.get("Extension") or "").lower() == ext.lower():
                    return
            ET.SubElement(ct_root, f"{{{NS['ct']}}}Default", {"Extension": ext, "ContentType": ctype})

        def _ensure_override(part_name: str, ctype: str):
            for o in ct_root.findall("ct:Override", NS):
                if o.get("PartName") == part_name:
                    return
            ET.SubElement(ct_root, f"{{{NS['ct']}}}Override", {"PartName": part_name, "ContentType": ctype})

        # Precompute which existing files we will replace, to avoid duplicate entries in zip.
        sheet_paths_to_modify: Set[str] = set()
        sheet_rels_to_modify: Set[str] = set()
        for title in by_sheet.keys():
            sp = title_to_sheet_path.get(title)
            if not sp:
                continue
            sheet_paths_to_modify.add(sp)
            sheet_rels_to_modify.add(sp.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels")

        skip_copy = set(sheet_paths_to_modify) | set(sheet_rels_to_modify) | {"[Content_Types].xml"}

        with zipfile.ZipFile(tmp_out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            # Copy all original files except those we will overwrite.
            for item in zin.infolist():
                if item.filename in skip_copy:
                    continue
                zout.writestr(item, zin.read(item.filename))

            injected = 0
            for title, imgs in by_sheet.items():
                sheet_path = title_to_sheet_path.get(title)
                if not sheet_path:
                    continue

                sheet_root = ET.fromstring(zin.read(sheet_path))

                # Create / update sheet rels
                sheet_rels_path = sheet_path.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
                if sheet_rels_path in existing:
                    rel_root = ET.fromstring(zin.read(sheet_rels_path))
                else:
                    rel_root = ET.Element(f"{{{NS['rel']}}}Relationships")

                # New drawing part
                drawing_idx = next_drawing
                next_drawing += 1
                drawing_name = f"drawing{drawing_idx}.xml"
                drawing_path = f"xl/drawings/{drawing_name}"
                drawing_rels_path = f"xl/drawings/_rels/{drawing_name}.rels"

                # Relationship id for drawing in sheet rels
                rid = f"rId_drawing_{drawing_idx}"
                ET.SubElement(
                    rel_root,
                    f"{{{NS['rel']}}}Relationship",
                    {
                        "Id": rid,
                        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                        "Target": f"../drawings/{drawing_name}",
                    },
                )

                # Add <drawing r:id="..."/> into sheet xml if not present
                has_drawing = sheet_root.find("w:drawing", NS) is not None
                if not has_drawing:
                    ET.SubElement(sheet_root, f"{{{NS['w']}}}drawing", {f"{{{NS['r']}}}id": rid})

                # Build drawing xml
                wsDr = ET.Element(f"{{{NS['xdr']}}}wsDr")
                drawing_rels = ET.Element(f"{{{NS['rel']}}}Relationships")

                pic_id = 1
                for row1, col1, img_file, y_off_px in imgs:
                    ext = (os.path.splitext(img_file)[1] or ".png").lower().lstrip(".")
                    if ext == "jpg":
                        ext = "jpeg"
                    media_name = f"image{next_image}.{ext}"
                    next_image += 1
                    media_path = f"xl/media/{media_name}"
                    with open(img_file, "rb") as f:
                        zout.writestr(media_path, f.read())

                    if ext == "png":
                        _ensure_default("png", "image/png")
                    elif ext in ("jpg", "jpeg"):
                        _ensure_default("jpeg", "image/jpeg")
                    elif ext == "gif":
                        _ensure_default("gif", "image/gif")
                    elif ext == "webp":
                        _ensure_default("webp", "image/webp")
                    else:
                        _ensure_default(ext, f"image/{ext}")

                    img_rid = f"rId{pic_id}"
                    ET.SubElement(
                        drawing_rels,
                        f"{{{NS['rel']}}}Relationship",
                        {
                            "Id": img_rid,
                            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                            "Target": f"../media/{media_name}",
                        },
                    )

                    col0 = max(col1 - 1, 0)
                    row0 = max(row1 - 1, 0)

                    # Preserve aspect ratio by computing extents from real image size.
                    w_px, h_px = _get_image_size(img_file)
                    max_w, max_h = 240, 240
                    scale = min(max_w / float(w_px), max_h / float(h_px), 1.0)
                    cx = _emu(int(w_px * scale))
                    cy = _emu(int(h_px * scale))

                    anchor = ET.SubElement(wsDr, f"{{{NS['xdr']}}}oneCellAnchor")
                    frm = ET.SubElement(anchor, f"{{{NS['xdr']}}}from")
                    ET.SubElement(frm, f"{{{NS['xdr']}}}col").text = str(col0)
                    ET.SubElement(frm, f"{{{NS['xdr']}}}colOff").text = "0"
                    ET.SubElement(frm, f"{{{NS['xdr']}}}row").text = str(row0)
                    ET.SubElement(frm, f"{{{NS['xdr']}}}rowOff").text = str(_emu(y_off_px))
                    ET.SubElement(anchor, f"{{{NS['xdr']}}}ext", {"cx": str(cx), "cy": str(cy)})

                    pic = ET.SubElement(anchor, f"{{{NS['xdr']}}}pic")
                    nv = ET.SubElement(pic, f"{{{NS['xdr']}}}nvPicPr")
                    ET.SubElement(nv, f"{{{NS['xdr']}}}cNvPr", {"id": str(pic_id), "name": media_name})
                    ET.SubElement(nv, f"{{{NS['xdr']}}}cNvPicPr")
                    blipFill = ET.SubElement(pic, f"{{{NS['xdr']}}}blipFill")
                    ET.SubElement(blipFill, f"{{{NS['a']}}}blip", {f"{{{NS['r']}}}embed": img_rid})
                    ET.SubElement(blipFill, f"{{{NS['a']}}}stretch")
                    spPr = ET.SubElement(pic, f"{{{NS['xdr']}}}spPr")
                    xfrm = ET.SubElement(spPr, f"{{{NS['a']}}}xfrm")
                    ET.SubElement(xfrm, f"{{{NS['a']}}}off", {"x": "0", "y": "0"})
                    ET.SubElement(xfrm, f"{{{NS['a']}}}ext", {"cx": str(cx), "cy": str(cy)})
                    ET.SubElement(spPr, f"{{{NS['a']}}}prstGeom", {"prst": "rect"})
                    ET.SubElement(anchor, f"{{{NS['xdr']}}}clientData")

                    pic_id += 1
                    injected += 1

                zout.writestr(drawing_path, ET.tostring(wsDr, encoding="utf-8", xml_declaration=True))
                zout.writestr(drawing_rels_path, ET.tostring(drawing_rels, encoding="utf-8", xml_declaration=True))
                _ensure_override(f"/xl/drawings/{drawing_name}", "application/vnd.openxmlformats-officedocument.drawing+xml")

                zout.writestr(sheet_path, ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True))
                zout.writestr(sheet_rels_path, ET.tostring(rel_root, encoding="utf-8", xml_declaration=True))

            # Write updated content types
            zout.writestr("[Content_Types].xml", ET.tostring(ct_root, encoding="utf-8", xml_declaration=True))

    os.replace(tmp_out, xlsx_path)
    return injected

def _get_object_type_name(entity) -> str:
    """Return ObjectType name: Shot / Asset Build / Folder / Task ..."""
    try:
        v = entity.get("object_type.name")
        if v:
            return str(v)
    except Exception:
        pass
    try:
        t = entity.get("object_type")
        if t:
            try:
                return str(t.get("name") or t["name"])
            except Exception:
                return str(t)
    except Exception:
        pass
    return "TypedContext"


def _get_pipeline_type_name(entity) -> str:
    """Return pipeline Type name: Character / Prop / Env ..."""
    try:
        v = entity.get("type.name")
        if v:
            return str(v)
    except Exception:
        pass
    try:
        t = entity.get("type")
        if t:
            try:
                return str(t.get("name") or t["name"])
            except Exception:
                return str(t)
    except Exception:
        pass
    return ""

def _get_task_type_name(entity) -> str:
    """Best-effort task type name (Audio/Comp/Env...)."""
    # Common: Task.type.name
    v = _get_pipeline_type_name(entity)
    if v:
        return v
    # Some schemas: Task.task_type.name
    try:
        v2 = entity.get("task_type.name")
        if v2:
            return str(v2)
    except Exception:
        pass
    try:
        tt = entity.get("task_type")
        if tt:
            try:
                return str(tt.get("name") or tt["name"])
            except Exception:
                return str(tt)
    except Exception:
        pass
    return ""


def _get_display_type(entity) -> str:
    """UI display style: ObjectType (TypeName). Example: Asset Build (Character)."""
    ot = _get_object_type_name(entity)
    tn = _get_pipeline_type_name(entity)
    if tn and tn.lower() != ot.lower():
        return f"{ot} ({tn})"
    return ot


def _is_task_entity(entity) -> bool:
    try:
        return getattr(entity, "entity_type", None) == "Task"
    except Exception:
        return False


def _thumbnail_url_from_value(value) -> Optional[str]:
    if not value:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, dict):
        for k in ("url", "original", "small", "medium", "large", "thumbnail"):
            v = value.get(k)
            if isinstance(v, str) and v:
                return v
        # Some schemas store {size:{url:..}}
        for v in value.values():
            if isinstance(v, dict) and isinstance(v.get("url"), str):
                return v.get("url")
    return None

def _looks_like_image_header(head: bytes) -> bool:
    if not head:
        return False
    if head.startswith(b"\x89PNG\r\n\x1a\n"):
        return True
    if head[:3] == b"\xff\xd8\xff":
        return True
    if head[:6] in (b"GIF87a", b"GIF89a"):
        return True
    # WEBP: RIFF....WEBP
    if head.startswith(b"RIFF") and len(head) >= 12 and head[8:12] == b"WEBP":
        return True
    # BMP
    if head[:2] == b"BM":
        return True
    return False

def _get_image_size(path: str) -> Tuple[int, int]:
    """Get image (width, height) without Pillow.

    Supports PNG/JPEG/GIF/BMP/WEBP (best-effort).
    """
    try:
        with open(path, "rb") as f:
            data = f.read(512)
    except Exception:
        return (240, 240)

    # PNG
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        try:
            w = int.from_bytes(data[16:20], "big")
            h = int.from_bytes(data[20:24], "big")
            return (max(w, 1), max(h, 1))
        except Exception:
            return (240, 240)

    # GIF
    if data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        w = int.from_bytes(data[6:8], "little")
        h = int.from_bytes(data[8:10], "little")
        return (max(w, 1), max(h, 1))

    # BMP
    if data[:2] == b"BM" and len(data) >= 26:
        w = int.from_bytes(data[18:22], "little")
        h = int.from_bytes(data[22:26], "little")
        return (max(w, 1), max(h, 1))

    # WEBP (RIFF....WEBP)
    if data.startswith(b"RIFF") and len(data) >= 32 and data[8:12] == b"WEBP":
        # VP8X
        if data[12:16] == b"VP8X" and len(data) >= 30:
            w = 1 + int.from_bytes(data[24:27], "little")
            h = 1 + int.from_bytes(data[27:30], "little")
            return (max(w, 1), max(h, 1))
        # VP8L (lossless)
        if data[12:16] == b"VP8L" and len(data) >= 25:
            b0 = data[21:25]
            bits = int.from_bytes(b0, "little")
            w = (bits & 0x3FFF) + 1
            h = ((bits >> 14) & 0x3FFF) + 1
            return (max(w, 1), max(h, 1))

    # JPEG: parse markers to find SOF
    if data[:3] == b"\xff\xd8\xff":
        try:
            with open(path, "rb") as f:
                f.read(2)  # SOI
                while True:
                    b = f.read(1)
                    if not b:
                        break
                    if b != b"\xff":
                        continue
                    marker = f.read(1)
                    if not marker or marker == b"\xd9":
                        break
                    # Skip padding
                    while marker == b"\xff":
                        marker = f.read(1)
                    # SOF0/SOF2 etc
                    if marker in (b"\xc0", b"\xc1", b"\xc2", b"\xc3", b"\xc5", b"\xc6", b"\xc7", b"\xc9", b"\xca", b"\xcb", b"\xcd", b"\xce", b"\xcf"):
                        ln = int.from_bytes(f.read(2), "big")
                        _ = f.read(1)  # precision
                        h = int.from_bytes(f.read(2), "big")
                        w = int.from_bytes(f.read(2), "big")
                        return (max(w, 1), max(h, 1))
                    else:
                        ln = int.from_bytes(f.read(2), "big")
                        if ln < 2:
                            break
                        f.seek(ln - 2, 1)
        except Exception:
            return (240, 240)

    return (240, 240)


class _ThumbSignals(QtCore.QObject):
    icon_ready = QtCore.Signal(str, object)  # (entity_id, QIcon)


class _ThumbTask(QtCore.QRunnable):
    def __init__(self, session: ftrack_api.Session, entity_id: str, url: str, size: int, signals: _ThumbSignals):
        super(_ThumbTask, self).__init__()
        self._session = session
        self._entity_id = entity_id
        self._url = url
        self._size = size
        self._signals = signals

    def run(self):
        try:
            tmpdir = ensure_tmp_dir("thumb_")
            out_path = os.path.join(tmpdir, f"{self._entity_id}.png")
            urllib.request.urlretrieve(self._url, out_path)  # nosec
            pm = QtGui.QPixmap(out_path)
            if not pm.isNull():
                pm = pm.scaled(self._size, self._size, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                self._signals.icon_ready.emit(self._entity_id, QtGui.QIcon(pm))
        except Exception:
            return


def _is_milestone_entity(entity) -> bool:
    try:
        ot = _get_object_type_name(entity) or ""
        if ot.lower().startswith("milestone"):
            return True
    except Exception:
        pass
    try:
        return getattr(entity, "entity_type", None) == "Milestone"
    except Exception:
        return False


class _IndexNode:
    __slots__ = ("key", "name", "type_name", "parent_key", "children", "has_notes", "sheet_name", "path_text")
    def __init__(self, key: str, name: str, type_name: str, parent_key: Optional[str]):
        self.key = key
        self.name = name
        self.type_name = type_name
        self.parent_key = parent_key
        self.children: List[str] = []
        self.has_notes: bool = False
        self.sheet_name: Optional[str] = None
        self.path_text: str = ""


class NotesExporter:
    def __init__(self, session: ftrack_api.Session):
        self.session = session

    # -----------------------
    # Data discovery helpers
    # -----------------------
    def _get_project(self, project_id: str):
        return self.session.get("Project", project_id)

    def _fetch_all_typed_contexts(self, project_id: str) -> List[dict]:
        q = (
            "select id, name, parent_id, object_type.name, type.name "
            "from TypedContext "
            f'where project_id is "{_esc(project_id)}" and object_type.name is_not "Task"'
        )
        # TypedContext query may also return Task entities (Task inherits TypedContext).
        # For tree traversal purposes we exclude Task here; Task nodes are handled separately.
        res = self.session.query(q).all()
        out = []
        for e in res:
            try:
                if getattr(e, "entity_type", None) == "Task":
                    continue
            except Exception:
                pass
            out.append(e)
        return out

    def _fetch_tasks_by_parents(self, project_id: str, parent_ids: List[str]) -> List[dict]:
        out: List[dict] = []
        for batch in _chunked(parent_ids, 200):
            ids = ",".join(f'"{_esc(i)}"' for i in batch)
            q = (
                "select id, name, parent_id, type.name from Task "
                f'where project_id is "{_esc(project_id)}" and parent_id in ({ids})'
            )
            out.extend(self.session.query(q).all())
        return out

    def _fetch_assets_by_context_ids(self, project_id: str, context_ids: List[str]) -> List[dict]:
        out: List[dict] = []
        for batch in _chunked(context_ids, 200):
            ids = ",".join(f'"{_esc(i)}"' for i in batch)
            q = (
                "select id, name, context_id from Asset "
                f'where project_id is "{_esc(project_id)}" and context_id in ({ids})'
            )
            out.extend(self.session.query(q).all())
        return out

    def _fetch_asset_versions_by_asset_ids(self, project_id: str, asset_ids: List[str]) -> List[dict]:
        out: List[dict] = []
        for batch in _chunked(asset_ids, 200):
            ids = ",".join(f'"{_esc(i)}"' for i in batch)
            q = (
                "select id, version, asset_id, task_id from AssetVersion "
                f'where project_id is "{_esc(project_id)}" and asset_id in ({ids})'
            )
            out.extend(self.session.query(q).all())
        return out

    def _fetch_asset_versions_by_task_ids(self, project_id: str, task_ids: List[str]) -> List[dict]:
        """Fetch AssetVersions related to tasks (Task-basis mode)."""
        out: List[dict] = []
        for batch in _chunked(task_ids, 200):
            ids = ",".join(f'"{_esc(i)}"' for i in batch)
            q = (
                "select id, version, asset_id, task_id from AssetVersion "
                f'where project_id is "{_esc(project_id)}" and task_id in ({ids})'
            )
            out.extend(self.session.query(q).all())
        return out

    def _fetch_note_components_by_note_ids(self, note_ids: List[str]) -> Dict[str, List[dict]]:
        """Batch fetch NoteComponent -> Component for notes.

        Some ftrack API projections do not populate note['note_components'] reliably.
        This method provides an explicit, batch-friendly fallback.
        """
        mapping: Dict[str, List[dict]] = defaultdict(list)  # type: ignore[assignment]
        for batch in _chunked(note_ids, 200):
            ids = ",".join(f'"{_esc(i)}"' for i in batch)
            q = (
                "select id, note_id, component_id, component.name from NoteComponent "
                f"where note_id in ({ids})"
            )
            try:
                for nc in self.session.query(q).all():
                    nid = nc.get("note_id")
                    if nid:
                        mapping[nid].append(nc)
            except Exception:
                continue
        return dict(mapping)

    # -----------------------
    # Note fetching helpers
    # -----------------------
    def _fetch_notes_by_parent_ids(self, project_id: str, parent_ids: List[str]) -> List[dict]:
        """Fetch notes by parent_id only (robust across parent_type casing / polymorphism)."""
        notes: List[dict] = []
        for batch in _chunked(parent_ids, 200):
            ids = ",".join(f'"{_esc(i)}"' for i in batch)
            q = (
                "select id, content, date, parent_id, parent_type, "
                "in_reply_to_id, "
                "frame_number, "
                "author.id, "
                "author.username, author.first_name, author.last_name, "
                "note_label_links.label.id, note_label_links.label.name, "
                "note_components.component.name "
                "from Note "
                f'where project_id is "{_esc(project_id)}" '
                f"and parent_id in ({ids})"
            )
            notes.extend(self.session.query(q).all())
        return notes

    # -----------------------
    # Attachment download
    # -----------------------
    def _download_image_attachments(self, note: dict, tmpdir: str) -> List[str]:
        """Return local paths for image attachments only."""
        try:
            server_location = self.session.query('Location where name is "ftrack.server"').one()
        except Exception:
            return []
        paths: List[str] = []
        note_components = note.get("note_components") or note.get("_note_components") or []
        for note_component in note_components:
            try:
                comp = None
                # note_component can be a NoteComponent entity or a dict-ish projection.
                if isinstance(note_component, dict) and "component" in note_component:
                    comp = note_component.get("component")
                if comp is None:
                    comp_id = None
                    try:
                        comp_id = note_component.get("component_id")  # type: ignore[union-attr]
                    except Exception:
                        comp_id = None
                    if not comp_id:
                        try:
                            comp_id = note_component.get("component.id")  # type: ignore[union-attr]
                        except Exception:
                            comp_id = None
                    if comp_id:
                        comp = self.session.get("Component", comp_id)
                if not comp:
                    continue
                name = comp.get("name") or "attachment"
                # Some image attachments may have no extension in component name.
                # Prefer a permissive approach: try download and validate header.
                url = None
                try:
                    url = server_location.get_url(comp)
                except Exception:
                    url = None
                if not url:
                    # Fallback: thumbnail URL is often accessible even when original is protected.
                    try:
                        url = server_location.get_thumbnail_url(comp, size=1024)
                    except Exception:
                        url = None
                if not url:
                    continue
                fn = safe_filename(name)
                out_path = os.path.join(tmpdir, f"{note['id']}_{fn}")
                urllib.request.urlretrieve(url, out_path)  # nosec - url comes from ftrack server location helper
                # Basic signature check: skip if not an image file header (prevents saving HTML login pages).
                try:
                    with open(out_path, "rb") as f:
                        head = f.read(32)
                    if not _looks_like_image_header(head):
                        continue
                except Exception:
                    pass
                paths.append(out_path)
            except Exception:
                continue
        return paths


def export_notes_job(
    project_id: str,
    selected_context_ids: List[str],
    selected_task_ids: List[str],
    include_descendants: bool,
    export_asset_version_notes: bool,
    export_asset_version_task_basis: bool,
    api_username: Optional[str] = None,
):
    """Background worker entry. Creates its own Session and updates a Job."""
    try:
        bg_session = ftrack_api.Session()
    except Exception as e:
        logger.error(f"Failed to create background session: {e}")
        return

    job_id = None
    try:
        user = None
        if api_username:
            try:
                user = bg_session.query(f'Select id from User where username is "{_esc(api_username)}"').first()
            except Exception:
                user = None
        try:
            job = bg_session.create(
                "Job",
                {
                    "user": user,
                    "status": "running",
                    "data": json.dumps({"description": "Notes export (processing)"}),
                },
            )
            bg_session.commit()
            job_id = job["id"]
        except Exception:
            logger.warning("Failed to create Job; continuing without Job updates.")

        exporter = NotesExporter(bg_session)
        project = exporter._get_project(project_id)
        project_name = (project.get("name") or "project").strip()

        # 1) Fetch all non-task TypedContexts for project (needed for path + descendants)
        typed_contexts = exporter._fetch_all_typed_contexts(project_id)
        ctx_by_id: Dict[str, dict] = {c["id"]: c for c in typed_contexts}
        children_map: Dict[str, List[str]] = defaultdict(list)
        for c in typed_contexts:
            pid = c.get("parent_id")
            if pid:
                children_map[pid].append(c["id"])

        # 2) Compute context scope (descendants optional) for context-based modes.
        selected_context_set: Set[str] = set([cid for cid in selected_context_ids if cid])
        all_context_ids: Set[str] = set(selected_context_set)
        if include_descendants and selected_context_set:
            stack = list(selected_context_set)
            while stack:
                nid = stack.pop()
                for ch in children_map.get(nid, []):
                    if ch not in all_context_ids:
                        all_context_ids.add(ch)
                        stack.append(ch)

        # 3) Collect tasks depending on mode
        tasks_by_id: Dict[str, dict] = {}
        if not export_asset_version_notes:
            # Export non-AssetVersion notes: tasks are eligible.
            # - If include_descendants: tasks under all_context_ids
            # - Always include explicitly selected tasks
            task_ids: Set[str] = set([tid for tid in selected_task_ids if tid])
            if include_descendants and all_context_ids:
                for t in exporter._fetch_tasks_by_parents(project_id, list(all_context_ids)):
                    tasks_by_id[t["id"]] = t
                task_ids.update(tasks_by_id.keys())
            if task_ids:
                # Ensure we have Task entities for explicitly selected tasks (to know parent_id/type.name).
                missing = [tid for tid in task_ids if tid not in tasks_by_id]
                if missing:
                    for batch in _chunked(missing, 200):
                        ids = ",".join(f'"{_esc(i)}"' for i in batch)
                        q = (
                            "select id, name, parent_id, type.name from Task "
                            f'where project_id is "{_esc(project_id)}" and id in ({ids})'
                        )
                        for t in bg_session.query(q).all():
                            tasks_by_id[t["id"]] = t

        # 4) Collect assetVersions for AssetVersion-mode
        asset_versions: List[dict] = []
        assets_by_id: Dict[str, dict] = {}
        asset_context_ids: Set[str] = set()
        if export_asset_version_notes:
            if export_asset_version_task_basis:
                # Task-basis: versions related to selected tasks.
                asset_versions = exporter._fetch_asset_versions_by_task_ids(project_id, list(dict.fromkeys(selected_task_ids)))
                asset_ids = [v.get("asset_id") for v in asset_versions if v.get("asset_id")]
                # Fetch assets by id to get name + context_id.
                for batch in _chunked(asset_ids, 200):
                    ids = ",".join(f'"{_esc(i)}"' for i in batch if i)
                    if not ids:
                        continue
                    q = f'select id, name, context_id from Asset where project_id is "{_esc(project_id)}" and id in ({ids})'
                    for a in bg_session.query(q).all():
                        if a.get("id"):
                            assets_by_id[a["id"]] = a
                asset_context_ids = set([a.get("context_id") for a in assets_by_id.values() if a.get("context_id")])
                # Build context ids for tree display: include context + ancestors
                all_context_ids = set()
                for cid in asset_context_ids:
                    cur = cid
                    while cur and cur not in all_context_ids and cur in ctx_by_id:
                        all_context_ids.add(cur)
                        cur = ctx_by_id[cur].get("parent_id")
            else:
                # Context-basis: versions under selected contexts.
                asset_context_ids = set(all_context_ids)
                if asset_context_ids:
                    assets = exporter._fetch_assets_by_context_ids(project_id, list(asset_context_ids))
                    asset_ids = [a["id"] for a in assets if a.get("id")]
                    for a in assets:
                        if a.get("id"):
                            assets_by_id[a["id"]] = a
                    if asset_ids:
                        asset_versions = exporter._fetch_asset_versions_by_asset_ids(project_id, asset_ids)

        # 4.5) Build related task path mapping for asset versions (best-effort).
        av_task_path_by_av_id: Dict[str, str] = {}
        if export_asset_version_notes and asset_versions:
            task_ids = sorted({v.get("task_id") for v in asset_versions if v.get("task_id")})
            tasks_by_id: Dict[str, dict] = {}
            try:
                for batch in _chunked(task_ids, 200):
                    ids = ",".join(f'"{_esc(i)}"' for i in batch)
                    q = f"select id, link from Task where id in ({ids})"
                    for t in bg_session.query(q).all():
                        tid = t.get("id")
                        if tid:
                            tasks_by_id[tid] = t
            except Exception:
                tasks_by_id = {}

            for av in asset_versions:
                avid = av.get("id")
                if not avid:
                    continue
                tid = av.get("task_id")
                if not tid:
                    av_task_path_by_av_id[avid] = "N/A"
                    continue
                t = tasks_by_id.get(tid)
                if not t:
                    av_task_path_by_av_id[avid] = "N/A"
                    continue
                try:
                    link = t.get("link") or []
                    if isinstance(link, list) and link:
                        av_task_path_by_av_id[avid] = " / ".join([str(x.get("name") or "") for x in link if hasattr(x, "get") and x.get("name")])
                    else:
                        av_task_path_by_av_id[avid] = "N/A"
                except Exception:
                    av_task_path_by_av_id[avid] = "N/A"

        # 5) Build parent_ids for note fetching (by id only)
        parent_ids: Set[str] = set()
        if export_asset_version_notes:
            parent_ids.update([v["id"] for v in asset_versions if v.get("id")])
        else:
            parent_ids.update(all_context_ids)
            parent_ids.update(tasks_by_id.keys())

        # 6) Fetch notes
        if job_id:
            try:
                j = bg_session.get("Job", job_id)
                j["data"] = json.dumps({"description": "Notes export (fetching notes)"})
                bg_session.commit()
            except Exception:
                pass

        notes = exporter._fetch_notes_by_parent_ids(project_id, list(parent_ids)) if parent_ids else []
        # 7) De-dup notes
        note_by_id: Dict[str, dict] = {}
        for n in notes:
            nid = n.get("id")
            if nid and nid not in note_by_id:
                note_by_id[nid] = n
        notes = list(note_by_id.values())

        # 7.05) Batch fetch authors (some projections may not populate author.username reliably)
        author_name_by_id: Dict[str, str] = {}
        try:
            author_ids = sorted({_author_id_from_note(n) for n in notes if _author_id_from_note(n)})
            for batch in _chunked(author_ids, 200):
                ids = ",".join(f'"{_esc(i)}"' for i in batch)
                q = f"select id, username, first_name, last_name from User where id in ({ids})"
                for u in bg_session.query(q).all():
                    uid = u.get("id")
                    if not uid:
                        continue
                    first = (u.get("first_name") or "").strip()
                    last = (u.get("last_name") or "").strip()
                    full = f"{first} {last}".strip()
                    author_name_by_id[uid] = full or (u.get("username") or "")
        except Exception:
            author_name_by_id = {}

        # 7.1) Batch fetch attachments for notes (so details sheet can embed images).
        try:
            nc_map = exporter._fetch_note_components_by_note_ids([n.get("id") for n in notes if n.get("id")])
            for n in notes:
                nid = n.get("id")
                if nid and nid in nc_map:
                    n["_note_components"] = nc_map[nid]
        except Exception:
            pass

        # 8) Bucket notes to their authoritative entity.
        #    NOTE: In some ftrack configurations, notes on Task may report parent_type as "TypedContext".
        #    Also, replies may not have parent_id == entity_id. We therefore:
        #      1) Prefer mapping by known parent_id sets (tasks / contexts / assetVersions)
        #      2) If note is a reply and we can find its parent note, inherit that parent's bucket
        notes_bucket: Dict[Tuple[str, str], List[dict]] = defaultdict(list)
        note_bucket_by_note_id: Dict[str, Tuple[str, str]] = {}

        known_task_ids: Set[str] = set(tasks_by_id.keys()) if not export_asset_version_notes else set()
        known_ctx_ids: Set[str] = set(all_context_ids) if not export_asset_version_notes else set(all_context_ids)
        known_av_ids: Set[str] = set([v.get("id") for v in asset_versions if v.get("id")]) if export_asset_version_notes else set()

        def bucket_key_for_note(n: dict) -> Optional[Tuple[str, str]]:
            pid = n.get("parent_id") or ""
            if not pid:
                return None
            if pid in known_task_ids:
                return ("Task", pid)
            if pid in known_av_ids:
                return ("AssetVersion", pid)
            # Treat Context notes as TypedContext nodes in our tree
            if pid in known_ctx_ids:
                return ("TypedContext", pid)
            pt = _normalize_parent_type(n.get("parent_type") or "")
            if pt:
                return (pt, pid)
            return None

        # First pass: bucket notes that directly map to known entity ids.
        pending_replies: List[dict] = []
        for n in notes:
            nid = n.get("id") or ""
            key = bucket_key_for_note(n)
            if key and key[1] in known_task_ids | known_ctx_ids | known_av_ids:
                notes_bucket[key].append(n)
                if nid:
                    note_bucket_by_note_id[nid] = key
            else:
                pending_replies.append(n)

        # Second pass: attach replies to their in_reply_to note's bucket when possible.
        for n in pending_replies:
            nid = n.get("id") or ""
            in_reply_to = n.get("in_reply_to_id") or ""
            if in_reply_to and in_reply_to in note_bucket_by_note_id:
                key = note_bucket_by_note_id[in_reply_to]
                notes_bucket[key].append(n)
                if nid:
                    note_bucket_by_note_id[nid] = key
                continue
            # Fallback: bucket by parent_type + parent_id (may not exist in tree; will be ignored)
            key = bucket_key_for_note(n)
            if key:
                notes_bucket[key].append(n)
                if nid:
                    note_bucket_by_note_id[nid] = key

        # 9) Build index tree nodes (ancestors included)
        #    Represent each node with a key: "{type}:{id}"
        node_map: Dict[str, _IndexNode] = {}

        def ensure_node(key: str, name: str, type_name: str, parent_key: Optional[str]) -> _IndexNode:
            if key in node_map:
                return node_map[key]
            n = _IndexNode(key=key, name=name, type_name=type_name, parent_key=parent_key)
            node_map[key] = n
            if parent_key and parent_key in node_map and key not in node_map[parent_key].children:
                node_map[parent_key].children.append(key)
            return n

        # Project root node
        root_key = f"Project:{project_id}"
        ensure_node(root_key, project_name, "Project", None)

        # Ensure context nodes
        for cid in all_context_ids:
            c = ctx_by_id.get(cid)
            if not c:
                continue
            p = c.get("parent_id") or project_id
            parent_key = f"TypedContext:{p}" if p != project_id else root_key
            # Ensure ancestor placeholder
            if parent_key not in node_map:
                if p != project_id and p in ctx_by_id:
                    pp = ctx_by_id[p].get("parent_id") or project_id
                    pp_key = f"TypedContext:{pp}" if pp != project_id else root_key
                    ensure_node(pp_key, (ctx_by_id.get(pp, {}).get("name") if pp != project_id else project_name) or project_name, "TypedContext" if pp != project_id else "Project", root_key if pp != project_id else None)
                    ensure_node(parent_key, ctx_by_id[p].get("name") or "", _get_display_type(ctx_by_id[p]), pp_key)
                else:
                    ensure_node(parent_key, project_name, "Project", None)
            ensure_node(f"TypedContext:{cid}", c.get("name") or "", _get_display_type(c), parent_key)

        # Task nodes only in non-AssetVersion mode (and only if task has notes)
        if not export_asset_version_notes:
            for tid, t in tasks_by_id.items():
                if ("Task", tid) not in notes_bucket:
                    continue
                pid = t.get("parent_id")
                if not pid:
                    continue
                parent_key = f"TypedContext:{pid}"
                if parent_key not in node_map:
                    pc = ctx_by_id.get(pid)
                    if pc:
                        pp = pc.get("parent_id") or project_id
                        pp_key = f"TypedContext:{pp}" if pp != project_id else root_key
                        ensure_node(pp_key, (ctx_by_id.get(pp, {}).get("name") if pp != project_id else project_name) or project_name, "TypedContext" if pp != project_id else "Project", root_key if pp != project_id else None)
                        ensure_node(parent_key, pc.get("name") or "", _get_display_type(pc), pp_key)
                # Type display for tasks should be "Task (Audio)" etc
                ttype = _get_task_type_name(t)
                tdisp = f"Task ({ttype})" if ttype else "Task"
                ensure_node(f"Task:{tid}", t.get("name") or "", tdisp, parent_key)

        # AssetVersion nodes only in AssetVersion mode (and only if has notes)
        if export_asset_version_notes:
            for av in asset_versions:
                avid = av.get("id")
                if not avid or ("AssetVersion", avid) not in notes_bucket:
                    continue
                asset_id = av.get("asset_id")
                asset = assets_by_id.get(asset_id) if asset_id else None
                ctx_id = asset.get("context_id") if asset else None
                parent_key = f"TypedContext:{ctx_id}" if ctx_id else root_key
                if ctx_id and parent_key not in node_map and ctx_id in ctx_by_id:
                    # ensure context chain node exists
                    c = ctx_by_id[ctx_id]
                    p = c.get("parent_id") or project_id
                    pp_key = f"TypedContext:{p}" if p != project_id else root_key
                    if pp_key not in node_map and p in ctx_by_id:
                        p2 = ctx_by_id[p].get("parent_id") or project_id
                        p2_key = f"TypedContext:{p2}" if p2 != project_id else root_key
                        ensure_node(p2_key, (ctx_by_id.get(p2, {}).get("name") if p2 != project_id else project_name) or project_name, "TypedContext" if p2 != project_id else "Project", root_key if p2 != project_id else None)
                        ensure_node(pp_key, ctx_by_id[p].get("name") or "", _get_display_type(ctx_by_id[p]), p2_key)
                    ensure_node(parent_key, c.get("name") or "", _get_display_type(c), pp_key)
                label = (asset.get("name") if asset else None) or "Asset"
                ver = av.get("version")
                ver_txt = f"v{int(ver):03d}" if ver is not None else ""
                name = f"{label} {ver_txt}".strip()
                ensure_node(f"AssetVersion:{avid}", name, "Asset Version", parent_key)

        # Mark nodes that have notes (detail nodes) and compute their sheet names
        for (pt, pid), bucket in notes_bucket.items():
            key = f"{pt}:{pid}"
            if key not in node_map:
                continue
            node = node_map[key]
            node.has_notes = True
            suffix = short_hash(pid, 6)
            node.sheet_name = sanitize_sheet_name(f"{project_name}_{node.name}", suffix=suffix)

        # 10) Compute path_text for each node by walking parents
        def compute_path(key: str) -> str:
            parts = []
            cur = node_map.get(key)
            while cur is not None:
                if cur.name:
                    parts.append(cur.name)
                if not cur.parent_key:
                    break
                cur = node_map.get(cur.parent_key)
            return " / ".join(reversed(parts))

        for k, n in node_map.items():
            n.path_text = compute_path(k)

        # 11) Prepare Excel
        if job_id:
            try:
                j = bg_session.get("Job", job_id)
                j["data"] = json.dumps({"description": "Notes export (building xlsx)"})
                bg_session.commit()
            except Exception:
                pass

        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, PatternFill
        from openpyxl.utils import get_column_letter

        wb = Workbook()
        # remove default sheet
        try:
            wb.remove(wb.active)
        except Exception:
            pass
        # Index sheet
        index_ws = wb.create_sheet(title=sanitize_sheet_name(project_name, suffix="", max_len=31))
        # Fill color for frozen/header area (Index row 1; Detail rows 1-3).
        header_fill = PatternFill("solid", fgColor="BF9BC9")
        if export_asset_version_notes:
            index_ws.append(["Context", "Related Task", "Type", "Notes link", "Notes count"])
            header_cols = 5
        else:
            index_ws.append(["Context", "Type", "Notes link", "Notes count"])
            header_cols = 4
        for c in range(1, header_cols + 1):
            cell = index_ws.cell(1, c)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(vertical="top")
            cell.fill = header_fill
        # Notes Count column: left aligned
        count_col = 5 if export_asset_version_notes else 4
        index_ws.cell(1, count_col).alignment = Alignment(horizontal="left", vertical="top")
        # Freeze header row on index
        index_ws.freeze_panes = "A2"

        # Create detail sheets for nodes that have notes
        detail_ws_by_key = {}
        for k, n in node_map.items():
            if not n.has_notes or not n.sheet_name:
                continue
            title = n.sheet_name
            # Ensure uniqueness
            if title in wb.sheetnames:
                title = sanitize_sheet_name(title, suffix=short_hash(k, 6))
            ws = wb.create_sheet(title=title)
            # Row 1 reserved for "Back to Index"
            ws["A1"] = "Back to Index"
            # Row 2: title (no hyperlink)
            ws["A2"] = f'Notes on "{n.path_text}"'
            ws["A2"].font = Font(bold=True)
            ws["A1"].font = Font(size=16)
            ws["A1"].alignment = Alignment(horizontal="left", vertical="top")
            ws["A1"].fill = header_fill
            ws["A2"].fill = header_fill

            if export_asset_version_notes:
                headers = ["DateTime", "Contents", "Attachments", "Frame No.", "Author", "Label", "In reply to", "Note ID"]
            else:
                headers = ["DateTime", "Contents", "Attachments", "Author", "Label", "In reply to", "Note ID"]
            ws.append(headers)  # header becomes row 3
            for c in range(1, len(headers) + 1):
                cell = ws.cell(3, c)
                cell.font = Font(bold=True)
                # Header alignment: DateTime column right-aligned, others top-aligned (wrap where needed).
                if c == 1:
                    cell.alignment = Alignment(horizontal="right", vertical="top")
                else:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.fill = header_fill
            # Also fill row 3 for visual separation.
            for c in range(1, len(headers) + 1):
                ws.cell(1, c).fill = header_fill
                ws.cell(2, c).fill = header_fill
                ws.cell(3, c).fill = header_fill
            # Freeze rows 1-3 so Back + title + header remain visible
            ws.freeze_panes = "A4"
            detail_ws_by_key[k] = ws

        # Fill detail sheets
        tmpdir = ensure_tmp_dir("notes_export_")
        # image_placements items:
        # (sheet_title, row_1based, col_1based, image_path, y_offset_px)
        image_placements: List[tuple] = []
        embed_ok = 0
        embed_failed = 0
        download_ok = 0
        download_failed = 0
        download_non_image = 0
        for (pt, pid), bucket in notes_bucket.items():
            key = f"{pt}:{pid}"
            ws = detail_ws_by_key.get(key)
            if ws is None:
                continue
            rows = []
            for note in bucket:
                label_txt = ", ".join(_label_names_from_note(note))
                au = note.get("author")
                author_txt = ""
                try:
                    if au and hasattr(au, "get"):
                        first = (au.get("first_name") or "").strip()
                        last = (au.get("last_name") or "").strip()
                        full = f"{first} {last}".strip()
                        author_txt = full or (au.get("username") or "")
                except Exception:
                    author_txt = ""
                if not author_txt:
                    author_txt = note.get("author.username") or ""
                if not author_txt:
                    author_txt = author_name_by_id.get(_author_id_from_note(note), "")
                img_paths = exporter._download_image_attachments(note, tmpdir=tmpdir)
                if img_paths:
                    download_ok += len(img_paths)
                rows.append((note, author_txt, label_txt, img_paths))

            for note, author_txt, label_txt, img_paths in rows:
                dt = utc_to_local_str(note.get("date"))
                content = note.get("content") or ""
                # In reply to: show the referenced note content (not an id) for readability.
                in_reply_to_val = ""
                try:
                    rid = note.get("in_reply_to_id")
                    if rid:
                        in_reply_to_val = (note_by_id.get(rid, {}) or {}).get("content") or ""
                        if not in_reply_to_val:
                            in_reply_to_val = "N/A"
                except Exception:
                    in_reply_to_val = ""

                if export_asset_version_notes:
                    fn = note.get("frame_number")
                    frame_no: object = "N/A"
                    if fn is not None and fn != "":
                        try:
                            frame_no = int(fn) + 1
                        except Exception:
                            frame_no = "N/A"
                    base_row = [dt, content, "", frame_no, author_txt, label_txt, in_reply_to_val, note.get("id") or ""]
                else:
                    base_row = [dt, content, "", author_txt, label_txt, in_reply_to_val, note.get("id") or ""]
                ws.append(base_row)
                row_idx = ws.max_row
                # Align all cells Top, and enable wrapping where needed.
                max_col = len(base_row)
                for col in range(1, max_col + 1):
                    if col == 1:
                        ws.cell(row_idx, col).alignment = Alignment(horizontal="right", vertical="top")
                    elif col == 2:
                        ws.cell(row_idx, col).alignment = Alignment(wrap_text=True, vertical="top")
                    else:
                        # Frame No. column should be left-aligned for readability.
                        if export_asset_version_notes and col == 4:
                            ws.cell(row_idx, col).alignment = Alignment(horizontal="left", vertical="top")
                            continue
                        # Label / In reply to should wrap to avoid overly wide columns.
                        if export_asset_version_notes and col in (6, 7):
                            ws.cell(row_idx, col).alignment = Alignment(wrap_text=True, vertical="top")
                        elif (not export_asset_version_notes) and col in (5, 6):
                            ws.cell(row_idx, col).alignment = Alignment(wrap_text=True, vertical="top")
                        else:
                            ws.cell(row_idx, col).alignment = Alignment(vertical="top")

                # Ensure Frame No. is numeric when available.
                if export_asset_version_notes:
                    c = ws.cell(row_idx, 4)
                    if isinstance(frame_no, int):
                        c.number_format = "0"

                # Embed all images vertically within the single Attachments column (col=3).
                if img_paths:
                    y_off = 0
                    total_h_px = 0
                    max_img_w_px = 0
                    for pth in img_paths:
                        w_px, h_px = _get_image_size(pth)
                        max_w, max_h = 240, 240
                        scale = min(max_w / float(w_px), max_h / float(h_px), 1.0)
                        sw = int(w_px * scale)
                        sh = int(h_px * scale)
                        max_img_w_px = max(max_img_w_px, sw)
                        # Place each image below previous one in the same cell.
                        image_placements.append((ws.title, row_idx, 3, pth, y_off))
                        y_off += sh + 10  # 10px padding
                        total_h_px = y_off
                    # Row height in points (~0.75pt per px)
                    ws.row_dimensions[row_idx].height = max(20, int(total_h_px * 0.75))

        def _autofit(ws, wrap_cols: Optional[Set[int]] = None, fixed_width: Optional[Dict[int, int]] = None):
            wrap_cols = wrap_cols or set()
            fixed_width = fixed_width or {}
            for col_idx in range(1, ws.max_column + 1):
                if col_idx in fixed_width:
                    ws.column_dimensions[get_column_letter(col_idx)].width = fixed_width[col_idx]
                    continue
                max_len = 0
                for r in range(1, ws.max_row + 1):
                    v = ws.cell(r, col_idx).value
                    max_len = max(max_len, len(str(v or "")))
                # Excel max is 255; keep a high cap so long paths become visible.
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 8), 120)
            if wrap_cols:
                for col_idx in wrap_cols:
                    for r in range(1, ws.max_row + 1):
                        ws.cell(r, col_idx).alignment = Alignment(wrap_text=True, vertical="top")

        # Auto-fit detail sheets now; index will be auto-fit after it's populated.
        for _k, ws in detail_ws_by_key.items():
            # Attachment column fixed; wide enough so images won't overlap next columns.
            # 35 chars ~= ~250px, enough to host a 240px image.
            # Match Contents (col 2) width to Attachments (col 3) for better readability.
            fixed = {2: 35, 3: 35}
            if export_asset_version_notes:
                fixed[4] = 10  # Frame No.
                _autofit(ws, wrap_cols={2, 6, 7}, fixed_width=fixed)
            else:
                _autofit(ws, wrap_cols={2, 5, 6}, fixed_width=fixed)

        # 12) Write index sheet rows via tree traversal (preorder by name)
        index_row_by_key: Dict[str, int] = {}
        def walk(key: str):
            n = node_map.get(key)
            if not n:
                return
            count = len(notes_bucket.get((n.key.split(":")[0], n.key.split(":")[1]), [])) if n.has_notes else ""
            related_task = ""
            if export_asset_version_notes and n.key.startswith("AssetVersion:"):
                avid = n.key.split(":", 1)[1]
                related_task = av_task_path_by_av_id.get(avid, "N/A")
                index_ws.append([n.path_text, related_task, n.type_name, "", count])
                link_col = 4
            else:
                if export_asset_version_notes:
                    index_ws.append([n.path_text, "", n.type_name, "", count])
                    link_col = 4
                else:
                    index_ws.append([n.path_text, n.type_name, "", count])
                    link_col = 3
            row_idx = index_ws.max_row
            index_row_by_key[n.key] = row_idx
            # Notes Count column: left aligned for readability.
            try:
                index_ws.cell(row_idx, count_col).alignment = Alignment(horizontal="left", vertical="top")
            except Exception:
                pass
            if n.has_notes and n.sheet_name:
                link_cell = index_ws.cell(row_idx, link_col)
                link_cell.value = f'Notes on "{n.path_text}"'
                link_cell.hyperlink = f"#'{n.sheet_name}'!A1"
                # Make it look like a hyperlink (blue + underline)
                link_cell.style = "Hyperlink"
                link_cell.font = Font(color="0563C1", underline="single")
            # Children
            kids = [node_map[k] for k in n.children if k in node_map]
            kids.sort(key=lambda x: (x.type_name or "", x.name or ""))
            for ch in kids:
                walk(ch.key)

        walk(root_key)

        # Update "Back to Index" links on each detail sheet to jump back to the specific index row link cell.
        for key, ws in detail_ws_by_key.items():
            idx_row = index_row_by_key.get(key)
            if not idx_row:
                continue
            back = ws["A1"]
            back.value = "Back to Index"
            link_col_letter = "D" if export_asset_version_notes else "C"
            back.hyperlink = f"#'{index_ws.title}'!{link_col_letter}{idx_row}"
            back.style = "Hyperlink"
            back.font = Font(color="0563C1", underline="single", size=16)
            back.alignment = Alignment(horizontal="left", vertical="top")
            back.fill = header_fill

        # Auto-fit index after rows are populated; wrap for long path/link text.
        if export_asset_version_notes:
            _autofit(index_ws, wrap_cols={1, 2, 4})
        else:
            _autofit(index_ws, wrap_cols={1, 3})

        # 13) Save to temp and attach to job
        outdir = tempfile.mkdtemp(prefix="notes_export_out_")
        scope = "multi" if (len(selected_context_ids) + len(selected_task_ids)) > 1 else "single"
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        out_name = _ensure_single_xlsx(f"{project_name}_notes_{scope}_{ts}")
        out_path = os.path.join(outdir, safe_filename(out_name))
        wb.save(out_path)
        # Inject images after save (no Pillow required).
        try:
            injected = _inject_images_into_xlsx(out_path, image_placements)
            embed_ok = injected
        except Exception:
            pass

        if job_id:
            try:
                server_location = bg_session.query('Location where name is "ftrack.server"').one()
                # NOTE:
                # - out_path already ends with ".xlsx"
                # - ftrack will often append the file extension on download
                # If we include ".xlsx" in the component name, the downloaded filename becomes "*.xlsx.xlsx".
                comp_name = os.path.splitext(os.path.basename(out_path))[0]
                component = bg_session.create_component(out_path, data={"name": comp_name}, location=server_location)
                bg_session.create("JobComponent", {"component_id": component["id"], "job_id": job_id})
                j = bg_session.get("Job", job_id)
                j["status"] = "done"
                j["data"] = json.dumps({
                    "description": (
                        f"Notes export done. Notes: {len(notes)}; Entities: {len(detail_ws_by_key)}; "
                        f"Images embedded: {embed_ok}; Img files: {download_ok}"
                    )
                })
                bg_session.commit()
            except Exception as e:
                logger.warning(f"Failed to attach xlsx to job: {e}")
                try:
                    j = bg_session.get("Job", job_id)
                    j["status"] = "done"
                    j["data"] = json.dumps({"description": "Notes export done (xlsx attachment failed)."})
                    bg_session.commit()
                except Exception:
                    pass
    except Exception as e:
        logger.error(f"Export failed: {e}")
        if job_id:
            try:
                j = bg_session.get("Job", job_id)
                j["status"] = "failed"
                j["data"] = json.dumps({"description": f"Notes export failed: {e}"})
                bg_session.commit()
            except Exception:
                pass
    finally:
        try:
            bg_session.close()
        except Exception:
            pass


class NotesExporterWidget(ftrack_connect.ui.application.ConnectWidget):
    name = "NotesExporter"
    try:
        icon = QtGui.QIcon(qta_icon_pixmap("mdi6:file-export", 20))  # type: ignore[arg-type]
    except Exception:
        icon = QtGui.QIcon()

    def __init__(self, session, parent=None):
        super(NotesExporterWidget, self).__init__(session, parent=parent)
        self._session = session
        self._project_by_name: Dict[str, str] = {}
        self._ctx_loaded: Dict[str, bool] = {}
        self._thumb_cache: Dict[str, QtGui.QIcon] = {}
        self._thumb_signals = _ThumbSignals()
        self._thumb_signals.icon_ready.connect(self._on_thumb_ready)
        self._thumb_pool = QtCore.QThreadPool.globalInstance()
        self._build_ui()
        self._load_projects()
        self._apply_ftrack_checkbox_assets()

    def _build_ui(self):
        layout = QtWidgets.QVBoxLayout()
        self.setLayout(layout)

        title_row = QtWidgets.QHBoxLayout()
        title_icon = QtWidgets.QLabel()
        pm = qta_icon_pixmap("mdi6:note-multiple", 24)
        if pm is not None:
            title_icon.setPixmap(pm)
        title = QtWidgets.QLabel("Export Notes to Excel")
        title.setStyleSheet("font-size:18px; font-weight:600;")
        title_row.addWidget(title_icon)
        title_row.addWidget(title)
        title_row.addStretch(1)
        layout.addLayout(title_row)

        form = QtWidgets.QFormLayout()
        self.project_combo = QtWidgets.QComboBox()
        self.project_combo.currentIndexChanged.connect(self._on_project_changed)
        form.addRow("Project", self.project_combo)
        layout.addLayout(form)

        self.project_hint = QtWidgets.QLabel("Select a project to export notes from.")
        self.project_hint.setWordWrap(True)
        self.project_hint.setStyleSheet("color:#9aa0a6; margin:2px 0 10px 2px;")
        layout.addWidget(self.project_hint)

        mode_col = QtWidgets.QVBoxLayout()
        mode_row = QtWidgets.QHBoxLayout()
        self.cb_av_mode = QtWidgets.QCheckBox("Export Asset Version Notes")
        self.cb_av_mode.setChecked(False)
        self.cb_av_mode.stateChanged.connect(self._on_mode_changed)
        mode_row.addWidget(self.cb_av_mode)
        mode_row.addStretch(1)
        mode_col.addLayout(mode_row)

        # Sub option (indented, on its own line)
        sub_row = QtWidgets.QHBoxLayout()
        sub_row.addSpacing(18)
        self.cb_av_task_basis = QtWidgets.QCheckBox("Filter by selected task(s)")
        self.cb_av_task_basis.setToolTip(
            "When enabled, only tasks can be selected. Asset Version notes will be exported based on the selected tasks."
        )
        self.cb_av_task_basis.setChecked(False)
        self.cb_av_task_basis.setEnabled(False)
        self.cb_av_task_basis.stateChanged.connect(self._on_mode_changed)
        sub_row.addWidget(self.cb_av_task_basis)
        sub_row.addStretch(1)
        mode_col.addLayout(sub_row)

        layout.addLayout(mode_col)

        self.tree = QtWidgets.QTreeWidget()
        self.tree.setHeaderLabels(["Context", "Type"])
        self.tree.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.tree.itemExpanded.connect(self._on_item_expanded)
        self.tree.setSortingEnabled(True)
        try:
            hdr = self.tree.header()
            hdr.setSortIndicatorShown(True)
            hdr.setSectionsClickable(True)
            hdr.setSortIndicator(0, QtCore.Qt.AscendingOrder)
            # Make the divider start in the middle (equal column widths),
            # so Context text is not truncated right after expanding nodes.
            hdr.setStretchLastSection(False)
            hdr.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        except Exception:
            pass
        layout.addWidget(self.tree)

        opts = QtWidgets.QHBoxLayout()
        self.cb_desc = QtWidgets.QCheckBox("Include notes from descendants")
        self.cb_desc.setChecked(True)
        self.cb_desc.stateChanged.connect(self._on_scope_changed)
        opts.addWidget(self.cb_desc)
        opts.addStretch(1)
        layout.addLayout(opts)

        btn_row = QtWidgets.QHBoxLayout()
        self.run_btn = QtWidgets.QPushButton("Export")
        set_button_icon_qta(self.run_btn, "mdi6:file-export")
        self.run_btn.clicked.connect(self._run_export)
        btn_row.addStretch(1)
        btn_row.addWidget(self.run_btn)
        layout.addLayout(btn_row)

        self.status = QtWidgets.QLabel("")
        self.status.setStyleSheet("color:#666;")
        layout.addWidget(self.status)

    def _load_projects(self):
        self.project_combo.clear()
        self._project_by_name.clear()
        # Placeholder (disabled) - user must select a project.
        self.project_combo.addItem("Select a project to export notes from.")
        try:
            self.project_combo.model().item(0).setEnabled(False)  # type: ignore[call-arg]
        except Exception:
            pass
        try:
            projects = self._session.query('select id, name from Project order by name').all()
        except Exception as e:
            show_warn(self, "Load failed", f"Failed to load projects: {e}")
            return
        for p in projects:
            name = p.get("name") or ""
            pid = p.get("id") or ""
            if name and pid:
                self.project_combo.addItem(name)
                self._project_by_name[name] = pid
        # Keep placeholder selected by default.
        self.project_combo.setCurrentIndex(0)
        self.project_hint.setVisible(True)

    def _on_project_changed(self, *_):
        name = self.project_combo.currentText().strip()
        pid = self._project_by_name.get(name)
        self.tree.clear()
        self._ctx_loaded.clear()
        if not pid:
            self.project_hint.setVisible(True)
            return
        self.project_hint.setVisible(False)
        # Ensure mode-dependent enable/disable states are up-to-date.
        self._sync_mode_controls()
        # Add root item = Project
        root = QtWidgets.QTreeWidgetItem([name, "Project"])
        root.setData(0, QtCore.Qt.UserRole, ("Project", pid))
        root.setChildIndicatorPolicy(QtWidgets.QTreeWidgetItem.ShowIndicator)
        # In AssetVersion task-basis mode, only Task nodes are selectable.
        if bool(self.cb_av_mode.isChecked()) and bool(self.cb_av_task_basis.isChecked()):
            try:
                root.setFlags(root.flags() & ~QtCore.Qt.ItemIsSelectable)
            except Exception:
                pass
        self.tree.addTopLevelItem(root)
        self.tree.expandItem(root)

    def _sync_mode_controls(self):
        """Synchronize enable/disable states without rebuilding the tree."""
        av_mode = bool(self.cb_av_mode.isChecked())
        self.cb_av_task_basis.setEnabled(av_mode)
        if not av_mode:
            self.cb_av_task_basis.setChecked(False)
        task_basis = bool(self.cb_av_task_basis.isChecked())
        self.cb_desc.setEnabled(not (av_mode and task_basis))
        if av_mode and task_basis:
            self.cb_desc.setToolTip("This option is disabled because tasks are leaf nodes in this mode.")
        else:
            self.cb_desc.setToolTip("")
        self._apply_ftrack_checkbox_assets()

    def _current_connect_theme(self) -> str:
        """Return current connect theme name ('light' or 'dark')."""
        try:
            theme = getattr(self._session, "connect_theme", None)
            if theme in ("light", "dark"):
                return theme
        except Exception:
            pass
        # Fallback: infer from palette if connect_theme not set.
        try:
            col = QtWidgets.QApplication.palette().color(QtGui.QPalette.Window)
            # crude heuristic: dark window -> dark theme
            if col.red() + col.green() + col.blue() < (255 * 3) / 2:
                return "dark"
        except Exception:
            pass
        return "light"

    def _resolve_qrc_path(self, candidates: List[str]) -> Optional[str]:
        """Return first candidate QRC path that resolves to a non-null pixmap."""
        for p in candidates:
            try:
                pm = QtGui.QPixmap(p)
                if not pm.isNull():
                    return p
            except Exception:
                continue
        return None

    def _apply_ftrack_checkbox_assets(self):
        """Match ftrack Connect checkbox visuals without applying global theme.

        We only style *our* option checkboxes by pointing their indicator images
        at ftrack Connect's built-in QRC resources. This avoids side effects like
        wiping dark theme colors or header styles.
        """
        theme = self._current_connect_theme()
        names = {
            "checked": "checkbox_checked",
            "unchecked": "checkbox_unchecked",
            "checked_disabled": "checkbox_checked_disabled",
            "unchecked_disabled": "checkbox_unchecked_disabled",
        }

        def cand(n: str) -> List[str]:
            # Try common Connect resource layouts (varies slightly across versions).
            return [
                f":/ftrack/image/{theme}/{n}",
                f":/ftrack/image/{n}",
                f":/ftrack/titlebar/connect/image/{theme}/{n}",
                f":/ftrack/titlebar/connect/image/{n}",
                f":/ftrack/connect/image/{theme}/{n}",
                f":/ftrack/connect/image/{n}",
            ]

        checked = self._resolve_qrc_path(cand(names["checked"]))
        unchecked = self._resolve_qrc_path(cand(names["unchecked"]))
        checked_dis = self._resolve_qrc_path(cand(names["checked_disabled"]))
        unchecked_dis = self._resolve_qrc_path(cand(names["unchecked_disabled"]))

        if not (checked and unchecked and checked_dis and unchecked_dis):
            # If we cannot resolve the resources, do nothing (keep default / inherited style).
            return

        qss = f"""
        QCheckBox::indicator:unchecked {{ image: url({unchecked}); }}
        QCheckBox::indicator:checked {{ image: url({checked}); }}
        QCheckBox::indicator:unchecked:disabled {{ image: url({unchecked_dis}); }}
        QCheckBox::indicator:checked:disabled {{ image: url({checked_dis}); }}
        """

        for cb in (getattr(self, "cb_av_mode", None), getattr(self, "cb_av_task_basis", None), getattr(self, "cb_desc", None)):
            if cb is None:
                continue
            try:
                cb.setStyleSheet(qss)
            except Exception:
                pass

    def _on_mode_changed(self, *_):
        self._sync_mode_controls()
        # Rebuild tree for current project (task nodes are shown/hidden depending on mode)
        self._on_project_changed()

    def _on_scope_changed(self, *_):
        # Placeholder hook for future scope estimation updates
        pass

    def _on_item_expanded(self, item: QtWidgets.QTreeWidgetItem):
        data = item.data(0, QtCore.Qt.UserRole)
        if not data:
            return
        etype, eid = data
        key = f"{etype}:{eid}"
        if self._ctx_loaded.get(key):
            return
        self._ctx_loaded[key] = True
        export_av_mode = bool(self.cb_av_mode.isChecked())
        task_basis = bool(self.cb_av_task_basis.isChecked())
        show_tasks = (not export_av_mode) or (export_av_mode and task_basis)

        try:
            q = (
                "select id, name, parent_id, object_type.name, type.name, thumbnail_id, thumbnail_url "
                "from TypedContext "
                f'where parent_id is "{_esc(eid)}" and object_type.name is_not "Task" order by name'
            )
            children = self._session.query(q).all()
        except Exception:
            children = []

        for c in children:
            cid = c.get("id") or ""
            if not cid:
                continue

            # IMPORTANT:
            # - Task inherits TypedContext. A TypedContext query can yield Task entities.
            # - Do NOT call session.get('TypedContext', id) before checking if it's a Task,
            #   otherwise we lose the ability to detect Task via entity_type.
            if _is_task_entity(c):
                continue
            # Milestone behaves like task; should not appear in AssetVersion mode.
            if export_av_mode and _is_milestone_entity(c):
                continue

            cname = c.get("name") or ""
            ctype = _get_display_type(c)

            child_item = QtWidgets.QTreeWidgetItem([cname, str(ctype)])
            child_item.setData(0, QtCore.Qt.UserRole, ("TypedContext", cid))
            # In AssetVersion task-basis mode, TypedContext nodes are for navigation only.
            if export_av_mode and task_basis:
                try:
                    child_item.setFlags(child_item.flags() & ~QtCore.Qt.ItemIsSelectable)
                except Exception:
                    pass
            # Decide expand indicator by checking if there are *visible* children in current mode.
            has_visible_children = self._has_visible_children(parent_id=cid, show_tasks=show_tasks, export_av_mode=export_av_mode)
            # If we cannot determine reliably (None), prefer showing indicator to avoid hiding valid children.
            if _is_milestone_entity(c):
                child_item.setChildIndicatorPolicy(QtWidgets.QTreeWidgetItem.DontShowIndicator)
            elif has_visible_children is False:
                child_item.setChildIndicatorPolicy(QtWidgets.QTreeWidgetItem.DontShowIndicator)
            else:
                child_item.setChildIndicatorPolicy(QtWidgets.QTreeWidgetItem.ShowIndicator)
                child_item.setToolTip(0, "By default, notes from all descendant items will be included.")
            item.addChild(child_item)

            # Lazy-load thumbnail icon.
            self._maybe_request_thumb(entity=c, item=child_item)

            # Append task children when in non-AssetVersion mode
            if show_tasks and (not _is_milestone_entity(c)):
                try:
                    tq = (
                        "select id, name, parent_id, type.name, thumbnail_id, thumbnail_url "
                        "from Task "
                        f'where parent_id is "{_esc(cid)}" order by name'
                    )
                    tasks = self._session.query(tq).all()
                except Exception:
                    tasks = []
                for t in tasks:
                    tid = t.get("id")
                    if not tid:
                        continue
                    tname = t.get("name") or ""
                    ttype = _get_task_type_name(t)
                    tdisp = f"Task ({ttype})" if ttype else "Task"
                    titem = QtWidgets.QTreeWidgetItem([tname, tdisp])
                    titem.setData(0, QtCore.Qt.UserRole, ("Task", tid))
                    # tasks are leafs
                    titem.setChildIndicatorPolicy(QtWidgets.QTreeWidgetItem.DontShowIndicator)
                    child_item.addChild(titem)
                    self._maybe_request_thumb(entity=t, item=titem)

        # Ensure sorting for this parent
        try:
            item.sortChildren(self.tree.sortColumn(), self.tree.header().sortIndicatorOrder())
        except Exception:
            pass

    def _has_visible_children(self, parent_id: str, show_tasks: bool, export_av_mode: bool) -> Optional[bool]:
        """Return whether parent has any children visible in current mode.

        Returns:
          - True: definitely has visible children
          - False: definitely has no visible children
          - None: unable to determine (query failure) -> caller should fall back to showing indicator
        """
        if not parent_id:
            return False
        had_error = False
        has_any = False
        try:
            # TypedContext children (non-Task). In AssetVersion mode, also hide Milestone.
            extra = ' and object_type.name is_not "Milestone"' if export_av_mode else ""
            q1 = (
                "select id from TypedContext "
                f'where parent_id is "{_esc(parent_id)}" and object_type.name is_not "Task"{extra} limit 1'
            )
            if self._session.query(q1).first():
                return True
        except Exception:
            had_error = True
        if show_tasks:
            try:
                q2 = f'select id from Task where parent_id is "{_esc(parent_id)}" limit 1'
                if self._session.query(q2).first():
                    return True
            except Exception:
                had_error = True
        if had_error:
            return None
        return has_any

    def _maybe_request_thumb(self, entity: dict, item: QtWidgets.QTreeWidgetItem, size: int = 18):
        eid = entity.get("id") or ""
        if not eid:
            return
        if eid in self._thumb_cache:
            item.setIcon(0, self._thumb_cache[eid])
            return
        url = _thumbnail_url_from_value(entity.get("thumbnail_url"))
        if not url:
            # Fallback: try component -> server url
            thumb_id = entity.get("thumbnail_id")
            if thumb_id:
                try:
                    server_location = self._session.query('Location where name is "ftrack.server"').one()
                    comp = self._session.get("Component", thumb_id)
                    if comp:
                        url = server_location.get_url(comp)
                except Exception:
                    url = None
        if not url:
            return
        task = _ThumbTask(self._session, eid, url, size, self._thumb_signals)
        self._thumb_pool.start(task)

    def _on_thumb_ready(self, entity_id: str, icon):
        try:
            if not isinstance(icon, QtGui.QIcon):
                return
            self._thumb_cache[entity_id] = icon
            # Update any existing items with this entity id.
            it = QtWidgets.QTreeWidgetItemIterator(self.tree)
            while it.value():
                item = it.value()
                data = item.data(0, QtCore.Qt.UserRole)
                if data and len(data) == 2 and data[1] == entity_id:
                    item.setIcon(0, icon)
                it += 1
        except Exception:
            return

    def _run_export(self):
        project_name = self.project_combo.currentText().strip()
        project_id = self._project_by_name.get(project_name)
        if not project_id:
            show_warn(self, "Missing project", "Please select a project.")
            return

        selected = self.tree.selectedItems()
        if not selected:
            show_warn(self, "Nothing selected", "Please select one or more nodes in the tree.")
            return
        # Collect selected context/task ids (ignore project root selection)
        ctx_ids: List[str] = []
        task_ids: List[str] = []
        for it in selected:
            data = it.data(0, QtCore.Qt.UserRole)
            if not data:
                continue
            etype, eid = data
            if etype == "TypedContext":
                ctx_ids.append(eid)
            elif etype == "Task":
                task_ids.append(eid)
        if not ctx_ids and not task_ids:
            show_warn(self, "Invalid selection", "Please select one or more items (not only the project root).")
            return

        include_desc = bool(self.cb_desc.isChecked())
        export_av_mode = bool(self.cb_av_mode.isChecked())
        task_basis = bool(self.cb_av_task_basis.isChecked())
        if export_av_mode and task_basis:
            include_desc = False  # disabled in UI; enforce
            if not task_ids:
                show_warn(self, "Invalid selection", "Please select one or more tasks.")
                return
            if ctx_ids:
                show_warn(self, "Invalid selection", "Only tasks can be selected in task-basis mode.")
                return
        elif export_av_mode and (not task_basis) and task_ids:
            show_warn(self, "Invalid selection", "Task nodes are not supported in this mode. Disable 'Filter by selected tasks'.")
            return

        # Resolve username for Job visibility (best-effort)
        api_user = None
        try:
            api_user = getattr(self._session, "api_user", None) or os.getenv("FTRACK_API_USER")
        except Exception:
            api_user = None

        def _bg():
            export_notes_job(
                project_id=project_id,
                selected_context_ids=list(dict.fromkeys(ctx_ids)),
                selected_task_ids=list(dict.fromkeys(task_ids)),
                include_descendants=include_desc,
                export_asset_version_notes=export_av_mode,
                export_asset_version_task_basis=task_basis,
                api_username=api_user,
            )

        threading.Thread(target=_bg, daemon=True).start()
        show_info(self, "Export started", "Job has been created. Please check Jobs in ftrack to monitor progress.")


def register(session, **kw):
    if not isinstance(session, ftrack_api.session.Session):
        return
    plugin = ftrack_connect.ui.application.ConnectWidgetPlugin(NotesExporterWidget)
    plugin.register(session, priority=20)
