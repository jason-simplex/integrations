# :coding: utf-8
import json
import logging
import os
import pathlib
import threading
import tempfile
import urllib.request
import datetime
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
    safe_filename,
    ensure_tmp_dir,
)
from notes_export_common import (
    author_id_from_note as _author_id_from_note,
    chunked as _chunked,
    ensure_single_xlsx_filename as _ensure_single_xlsx,
    esc as _esc,
    get_display_type as _get_display_type,
    get_task_type_name as _get_task_type_name,
    is_milestone_entity as _is_milestone_entity,
    is_task_entity as _is_task_entity,
    label_names_from_note as _label_names_from_note,
    normalize_note_parent_type as _normalize_parent_type,
    thumbnail_url_from_value as _thumbnail_url_from_value,
)

from xlsx_image_injector import (
    get_image_size as _get_image_size,
    inject_images_into_xlsx as _inject_images_into_xlsx,
    looks_like_image_header as _looks_like_image_header,
)

logger = logging.getLogger("connect_notes_exporter")

from ui_utils import show_info, show_warn, qta_icon_pixmap, set_button_icon_qta

from ftrack_connect.qt import QtWidgets, QtCore, QtGui
import ftrack_connect.ui.application


class _ThumbSignals(QtCore.QObject):
    icon_ready = QtCore.Signal(str, object)  # (entity_id, QIcon)


class _ThumbTask(QtCore.QRunnable):
    def __init__(self, entity_id: str, url: str, size: int, signals: _ThumbSignals):
        super(_ThumbTask, self).__init__()
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


def _update_job_status(
    session: ftrack_api.Session,
    job_id: Optional[str],
    description: str,
    status: Optional[str] = None,
):
    if not job_id:
        return
    try:
        job = session.get("Job", job_id)
        if status:
            job["status"] = status
        job["data"] = json.dumps({"description": description})
        session.commit()
    except Exception:
        pass


def _create_export_job(
    session: ftrack_api.Session,
    api_username: Optional[str],
) -> Optional[str]:
    user = None
    if api_username:
        try:
            user = session.query(
                f'Select id from User where username is "{_esc(api_username)}"'
            ).first()
        except Exception:
            user = None

    try:
        job = session.create(
            "Job",
            {
                "user": user,
                "status": "running",
                "data": json.dumps({"description": "Notes export (processing)"}),
            },
        )
        session.commit()
        return job["id"]
    except Exception:
        logger.warning("Failed to create Job; continuing without Job updates.")
        return None


def _build_context_maps(
    typed_contexts: List[dict],
) -> Tuple[Dict[str, dict], Dict[str, List[str]]]:
    ctx_by_id: Dict[str, dict] = {c["id"]: c for c in typed_contexts}
    children_map: Dict[str, List[str]] = defaultdict(list)
    for context in typed_contexts:
        parent_id = context.get("parent_id")
        if parent_id:
            children_map[parent_id].append(context["id"])
    return ctx_by_id, dict(children_map)


def _compute_context_scope(
    selected_context_ids: List[str],
    children_map: Dict[str, List[str]],
    include_descendants: bool,
) -> Set[str]:
    selected_context_set = {cid for cid in selected_context_ids if cid}
    all_context_ids = set(selected_context_set)
    if include_descendants and selected_context_set:
        stack = list(selected_context_set)
        while stack:
            node_id = stack.pop()
            for child_id in children_map.get(node_id, []):
                if child_id not in all_context_ids:
                    all_context_ids.add(child_id)
                    stack.append(child_id)
    return all_context_ids


def _collect_tasks_for_export(
    exporter: NotesExporter,
    session: ftrack_api.Session,
    project_id: str,
    all_context_ids: Set[str],
    selected_task_ids: List[str],
    include_descendants: bool,
) -> Dict[str, dict]:
    tasks_by_id: Dict[str, dict] = {}
    task_ids: Set[str] = {tid for tid in selected_task_ids if tid}

    if include_descendants and all_context_ids:
        for task in exporter._fetch_tasks_by_parents(project_id, list(all_context_ids)):
            tasks_by_id[task["id"]] = task
        task_ids.update(tasks_by_id.keys())

    missing_ids = [task_id for task_id in task_ids if task_id not in tasks_by_id]
    for batch in _chunked(missing_ids, 200):
        ids = ",".join(f'"{_esc(i)}"' for i in batch)
        q = (
            "select id, name, parent_id, type.name from Task "
            f'where project_id is "{_esc(project_id)}" and id in ({ids})'
        )
        for task in session.query(q).all():
            tasks_by_id[task["id"]] = task

    return tasks_by_id


def _collect_asset_version_scope(
    exporter: NotesExporter,
    session: ftrack_api.Session,
    project_id: str,
    selected_task_ids: List[str],
    all_context_ids: Set[str],
    ctx_by_id: Dict[str, dict],
    task_basis: bool,
) -> Tuple[List[dict], Dict[str, dict], Set[str]]:
    asset_versions: List[dict] = []
    assets_by_id: Dict[str, dict] = {}

    if task_basis:
        asset_versions = exporter._fetch_asset_versions_by_task_ids(
            project_id,
            list(dict.fromkeys(selected_task_ids)),
        )
        asset_ids = [v.get("asset_id") for v in asset_versions if v.get("asset_id")]
        for batch in _chunked(asset_ids, 200):
            ids = ",".join(f'"{_esc(i)}"' for i in batch if i)
            if not ids:
                continue
            q = (
                "select id, name, context_id from Asset "
                f'where project_id is "{_esc(project_id)}" and id in ({ids})'
            )
            for asset in session.query(q).all():
                if asset.get("id"):
                    assets_by_id[asset["id"]] = asset

        display_context_ids: Set[str] = set()
        asset_context_ids = {
            asset.get("context_id")
            for asset in assets_by_id.values()
            if asset.get("context_id")
        }
        for context_id in asset_context_ids:
            current_id = context_id
            while (
                current_id
                and current_id not in display_context_ids
                and current_id in ctx_by_id
            ):
                display_context_ids.add(current_id)
                current_id = ctx_by_id[current_id].get("parent_id")
        return asset_versions, assets_by_id, display_context_ids

    if all_context_ids:
        assets = exporter._fetch_assets_by_context_ids(project_id, list(all_context_ids))
        asset_ids = [asset["id"] for asset in assets if asset.get("id")]
        for asset in assets:
            if asset.get("id"):
                assets_by_id[asset["id"]] = asset
        if asset_ids:
            asset_versions = exporter._fetch_asset_versions_by_asset_ids(
                project_id,
                asset_ids,
            )

    return asset_versions, assets_by_id, set(all_context_ids)


def _build_related_task_path_map(
    session: ftrack_api.Session,
    asset_versions: List[dict],
) -> Dict[str, str]:
    av_task_path_by_av_id: Dict[str, str] = {}
    if not asset_versions:
        return av_task_path_by_av_id

    task_ids = sorted({v.get("task_id") for v in asset_versions if v.get("task_id")})
    tasks_by_id: Dict[str, dict] = {}
    try:
        for batch in _chunked(task_ids, 200):
            ids = ",".join(f'"{_esc(i)}"' for i in batch)
            q = f"select id, link from Task where id in ({ids})"
            for task in session.query(q).all():
                task_id = task.get("id")
                if task_id:
                    tasks_by_id[task_id] = task
    except Exception:
        tasks_by_id = {}

    for asset_version in asset_versions:
        av_id = asset_version.get("id")
        if not av_id:
            continue
        task_id = asset_version.get("task_id")
        if not task_id or task_id not in tasks_by_id:
            av_task_path_by_av_id[av_id] = "N/A"
            continue

        task = tasks_by_id[task_id]
        try:
            link = task.get("link") or []
            if isinstance(link, list) and link:
                av_task_path_by_av_id[av_id] = " / ".join(
                    [
                        str(x.get("name") or "")
                        for x in link
                        if hasattr(x, "get") and x.get("name")
                    ]
                )
            else:
                av_task_path_by_av_id[av_id] = "N/A"
        except Exception:
            av_task_path_by_av_id[av_id] = "N/A"

    return av_task_path_by_av_id


def _build_note_parent_ids(
    all_context_ids: Set[str],
    tasks_by_id: Dict[str, dict],
    asset_versions: List[dict],
    export_asset_version_notes: bool,
) -> Set[str]:
    parent_ids: Set[str] = set()
    if export_asset_version_notes:
        parent_ids.update(v["id"] for v in asset_versions if v.get("id"))
    else:
        parent_ids.update(all_context_ids)
        parent_ids.update(tasks_by_id.keys())
    return parent_ids


def _dedupe_notes(notes: List[dict]) -> Tuple[List[dict], Dict[str, dict]]:
    note_by_id: Dict[str, dict] = {}
    for note in notes:
        note_id = note.get("id")
        if note_id and note_id not in note_by_id:
            note_by_id[note_id] = note
    return list(note_by_id.values()), note_by_id


def _fetch_author_name_map(
    session: ftrack_api.Session,
    notes: List[dict],
) -> Dict[str, str]:
    author_name_by_id: Dict[str, str] = {}
    try:
        author_ids = sorted(
            {
                author_id
                for author_id in (_author_id_from_note(note) for note in notes)
                if author_id
            }
        )
        for batch in _chunked(author_ids, 200):
            ids = ",".join(f'"{_esc(i)}"' for i in batch)
            q = (
                "select id, username, first_name, last_name from User "
                f"where id in ({ids})"
            )
            for user in session.query(q).all():
                user_id = user.get("id")
                if not user_id:
                    continue
                first = (user.get("first_name") or "").strip()
                last = (user.get("last_name") or "").strip()
                full = f"{first} {last}".strip()
                author_name_by_id[user_id] = full or (user.get("username") or "")
    except Exception:
        return {}
    return author_name_by_id


def _attach_note_components_to_notes(
    exporter: NotesExporter,
    notes: List[dict],
):
    try:
        note_component_map = exporter._fetch_note_components_by_note_ids(
            [note.get("id") for note in notes if note.get("id")]
        )
        for note in notes:
            note_id = note.get("id")
            if note_id and note_id in note_component_map:
                note["_note_components"] = note_component_map[note_id]
    except Exception:
        pass


def _bucket_notes_by_entity(
    notes: List[dict],
    all_context_ids: Set[str],
    tasks_by_id: Dict[str, dict],
    asset_versions: List[dict],
    export_asset_version_notes: bool,
) -> Dict[Tuple[str, str], List[dict]]:
    notes_bucket: Dict[Tuple[str, str], List[dict]] = defaultdict(list)
    note_bucket_by_note_id: Dict[str, Tuple[str, str]] = {}

    known_task_ids: Set[str] = set(tasks_by_id.keys()) if not export_asset_version_notes else set()
    known_ctx_ids: Set[str] = set(all_context_ids)
    known_av_ids: Set[str] = (
        {v.get("id") for v in asset_versions if v.get("id")}
        if export_asset_version_notes
        else set()
    )

    def bucket_key_for_note(note: dict) -> Optional[Tuple[str, str]]:
        parent_id = note.get("parent_id") or ""
        if not parent_id:
            return None
        if parent_id in known_task_ids:
            return ("Task", parent_id)
        if parent_id in known_av_ids:
            return ("AssetVersion", parent_id)
        if parent_id in known_ctx_ids:
            return ("TypedContext", parent_id)
        parent_type = _normalize_parent_type(note.get("parent_type") or "")
        if parent_type:
            return (parent_type, parent_id)
        return None

    pending_replies: List[dict] = []
    for note in notes:
        note_id = note.get("id") or ""
        key = bucket_key_for_note(note)
        if key and key[1] in (known_task_ids | known_ctx_ids | known_av_ids):
            notes_bucket[key].append(note)
            if note_id:
                note_bucket_by_note_id[note_id] = key
        else:
            pending_replies.append(note)

    for note in pending_replies:
        note_id = note.get("id") or ""
        in_reply_to = note.get("in_reply_to_id") or ""
        if in_reply_to and in_reply_to in note_bucket_by_note_id:
            key = note_bucket_by_note_id[in_reply_to]
            notes_bucket[key].append(note)
            if note_id:
                note_bucket_by_note_id[note_id] = key
            continue
        key = bucket_key_for_note(note)
        if key:
            notes_bucket[key].append(note)
            if note_id:
                note_bucket_by_note_id[note_id] = key

    return dict(notes_bucket)


def _build_index_node_map(
    project_id: str,
    project_name: str,
    all_context_ids: Set[str],
    ctx_by_id: Dict[str, dict],
    tasks_by_id: Dict[str, dict],
    asset_versions: List[dict],
    assets_by_id: Dict[str, dict],
    notes_bucket: Dict[Tuple[str, str], List[dict]],
    export_asset_version_notes: bool,
) -> Dict[str, _IndexNode]:
    node_map: Dict[str, _IndexNode] = {}
    root_key = f"Project:{project_id}"

    def ensure_node(
        key: str,
        name: str,
        type_name: str,
        parent_key: Optional[str],
    ) -> _IndexNode:
        if key in node_map:
            return node_map[key]
        node = _IndexNode(
            key=key,
            name=name,
            type_name=type_name,
            parent_key=parent_key,
        )
        node_map[key] = node
        if parent_key and parent_key in node_map and key not in node_map[parent_key].children:
            node_map[parent_key].children.append(key)
        return node

    def ensure_context_node(context_id: Optional[str]) -> str:
        if not context_id or context_id not in ctx_by_id:
            return root_key
        context = ctx_by_id[context_id]
        parent_id = context.get("parent_id")
        if not parent_id or parent_id == project_id:
            parent_key = root_key
        else:
            parent_key = ensure_context_node(parent_id)
        return ensure_node(
            f"TypedContext:{context_id}",
            context.get("name") or "",
            _get_display_type(context),
            parent_key,
        ).key

    ensure_node(root_key, project_name, "Project", None)

    for context_id in all_context_ids:
        ensure_context_node(context_id)

    if not export_asset_version_notes:
        for task_id, task in tasks_by_id.items():
            if ("Task", task_id) not in notes_bucket:
                continue
            parent_key = ensure_context_node(task.get("parent_id"))
            task_type = _get_task_type_name(task)
            display_type = f"Task ({task_type})" if task_type else "Task"
            ensure_node(
                f"Task:{task_id}",
                task.get("name") or "",
                display_type,
                parent_key,
            )

    if export_asset_version_notes:
        for asset_version in asset_versions:
            av_id = asset_version.get("id")
            if not av_id or ("AssetVersion", av_id) not in notes_bucket:
                continue
            asset = assets_by_id.get(asset_version.get("asset_id"))
            parent_key = ensure_context_node(asset.get("context_id") if asset else None)
            label = (asset.get("name") if asset else None) or "Asset"
            version = asset_version.get("version")
            version_text = f"v{int(version):03d}" if version is not None else ""
            ensure_node(
                f"AssetVersion:{av_id}",
                f"{label} {version_text}".strip(),
                "Asset Version",
                parent_key,
            )

    for parent_type, parent_id in notes_bucket:
        key = f"{parent_type}:{parent_id}"
        if key not in node_map:
            continue
        node = node_map[key]
        node.has_notes = True
        node.sheet_name = sanitize_sheet_name(
            f"{project_name}_{node.name}",
            suffix=short_hash(parent_id, 6),
        )

    def compute_path(key: str) -> str:
        parts: List[str] = []
        current = node_map.get(key)
        while current is not None:
            if current.name:
                parts.append(current.name)
            if not current.parent_key:
                break
            current = node_map.get(current.parent_key)
        return " / ".join(reversed(parts))

    for key, node in node_map.items():
        node.path_text = compute_path(key)

    return node_map


def _build_export_workbook(
    exporter: NotesExporter,
    project_id: str,
    project_name: str,
    selected_context_ids: List[str],
    selected_task_ids: List[str],
    export_asset_version_notes: bool,
    notes_bucket: Dict[Tuple[str, str], List[dict]],
    note_by_id: Dict[str, dict],
    node_map: Dict[str, _IndexNode],
    av_task_path_by_av_id: Dict[str, str],
    author_name_by_id: Dict[str, str],
) -> Dict[str, object]:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    def _autofit(
        worksheet,
        wrap_cols: Optional[Set[int]] = None,
        fixed_width: Optional[Dict[int, int]] = None,
    ):
        wrap_cols = wrap_cols or set()
        fixed_width = fixed_width or {}
        for col_idx in range(1, worksheet.max_column + 1):
            if col_idx in fixed_width:
                worksheet.column_dimensions[get_column_letter(col_idx)].width = fixed_width[col_idx]
                continue
            max_len = 0
            for row_idx in range(1, worksheet.max_row + 1):
                value = worksheet.cell(row_idx, col_idx).value
                max_len = max(max_len, len(str(value or "")))
            worksheet.column_dimensions[get_column_letter(col_idx)].width = min(
                max(max_len + 2, 8),
                120,
            )
        for col_idx in wrap_cols:
            for row_idx in range(1, worksheet.max_row + 1):
                worksheet.cell(row_idx, col_idx).alignment = Alignment(
                    wrap_text=True,
                    vertical="top",
                )

    wb = Workbook()
    try:
        wb.remove(wb.active)
    except Exception:
        pass

    index_ws = wb.create_sheet(
        title=sanitize_sheet_name(project_name, suffix="", max_len=31)
    )
    header_fill = PatternFill("solid", fgColor="BF9BC9")
    if export_asset_version_notes:
        index_ws.append(["Context", "Related Task", "Type", "Notes link", "Notes count"])
        header_cols = 5
    else:
        index_ws.append(["Context", "Type", "Notes link", "Notes count"])
        header_cols = 4
    for col_idx in range(1, header_cols + 1):
        cell = index_ws.cell(1, col_idx)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="top")
        cell.fill = header_fill
    count_col = 5 if export_asset_version_notes else 4
    index_ws.cell(1, count_col).alignment = Alignment(
        horizontal="left",
        vertical="top",
    )
    index_ws.freeze_panes = "A2"

    detail_ws_by_key = {}
    for key, node in node_map.items():
        if not node.has_notes or not node.sheet_name:
            continue
        title = node.sheet_name
        if title in wb.sheetnames:
            title = sanitize_sheet_name(title, suffix=short_hash(key, 6))
        ws = wb.create_sheet(title=title)
        ws["A1"] = "Back to Index"
        ws["A2"] = f'Notes on "{node.path_text}"'
        ws["A2"].font = Font(bold=True)
        ws["A1"].font = Font(size=16)
        ws["A1"].alignment = Alignment(horizontal="left", vertical="top")
        ws["A1"].fill = header_fill
        ws["A2"].fill = header_fill

        if export_asset_version_notes:
            headers = [
                "DateTime",
                "Contents",
                "Attachments",
                "Frame No.",
                "Author",
                "Label",
                "In reply to",
                "Note ID",
            ]
        else:
            headers = [
                "DateTime",
                "Contents",
                "Attachments",
                "Author",
                "Label",
                "In reply to",
                "Note ID",
            ]
        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(3, col_idx)
            cell.font = Font(bold=True)
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="right", vertical="top")
            else:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.fill = header_fill
            ws.cell(1, col_idx).fill = header_fill
            ws.cell(2, col_idx).fill = header_fill
            ws.cell(3, col_idx).fill = header_fill
        ws.freeze_panes = "A4"
        detail_ws_by_key[key] = ws

    tmpdir = ensure_tmp_dir("notes_export_")
    image_placements: List[tuple] = []
    embed_ok = 0
    download_ok = 0
    for (parent_type, parent_id), bucket in notes_bucket.items():
        ws = detail_ws_by_key.get(f"{parent_type}:{parent_id}")
        if ws is None:
            continue
        rows = []
        for note in bucket:
            label_text = ", ".join(_label_names_from_note(note))
            author_text = ""
            try:
                author = note.get("author")
                if author and hasattr(author, "get"):
                    first = (author.get("first_name") or "").strip()
                    last = (author.get("last_name") or "").strip()
                    full = f"{first} {last}".strip()
                    author_text = full or (author.get("username") or "")
            except Exception:
                author_text = ""
            if not author_text:
                author_text = note.get("author.username") or ""
            if not author_text:
                author_text = author_name_by_id.get(_author_id_from_note(note), "")
            image_paths = exporter._download_image_attachments(note, tmpdir=tmpdir)
            if image_paths:
                download_ok += len(image_paths)
            rows.append((note, author_text, label_text, image_paths))

        for note, author_text, label_text, image_paths in rows:
            date_text = utc_to_local_str(note.get("date"))
            content = note.get("content") or ""
            in_reply_to_value = ""
            try:
                reply_id = note.get("in_reply_to_id")
                if reply_id:
                    in_reply_to_value = (
                        (note_by_id.get(reply_id, {}) or {}).get("content") or ""
                    )
                    if not in_reply_to_value:
                        in_reply_to_value = "N/A"
            except Exception:
                in_reply_to_value = ""

            if export_asset_version_notes:
                frame_no: object = "N/A"
                frame_value = note.get("frame_number")
                if frame_value is not None and frame_value != "":
                    try:
                        frame_no = int(frame_value) + 1
                    except Exception:
                        frame_no = "N/A"
                base_row = [
                    date_text,
                    content,
                    "",
                    frame_no,
                    author_text,
                    label_text,
                    in_reply_to_value,
                    note.get("id") or "",
                ]
            else:
                frame_no = None
                base_row = [
                    date_text,
                    content,
                    "",
                    author_text,
                    label_text,
                    in_reply_to_value,
                    note.get("id") or "",
                ]

            ws.append(base_row)
            row_idx = ws.max_row
            for col_idx in range(1, len(base_row) + 1):
                if col_idx == 1:
                    ws.cell(row_idx, col_idx).alignment = Alignment(
                        horizontal="right",
                        vertical="top",
                    )
                elif col_idx == 2:
                    ws.cell(row_idx, col_idx).alignment = Alignment(
                        wrap_text=True,
                        vertical="top",
                    )
                else:
                    if export_asset_version_notes and col_idx == 4:
                        ws.cell(row_idx, col_idx).alignment = Alignment(
                            horizontal="left",
                            vertical="top",
                        )
                        continue
                    if export_asset_version_notes and col_idx in (6, 7):
                        ws.cell(row_idx, col_idx).alignment = Alignment(
                            wrap_text=True,
                            vertical="top",
                        )
                    elif (not export_asset_version_notes) and col_idx in (5, 6):
                        ws.cell(row_idx, col_idx).alignment = Alignment(
                            wrap_text=True,
                            vertical="top",
                        )
                    else:
                        ws.cell(row_idx, col_idx).alignment = Alignment(vertical="top")

            if export_asset_version_notes and isinstance(frame_no, int):
                ws.cell(row_idx, 4).number_format = "0"

            if image_paths:
                y_off = 0
                total_height_px = 0
                for image_path in image_paths:
                    width_px, height_px = _get_image_size(image_path)
                    scale = min(240 / float(width_px), 240 / float(height_px), 1.0)
                    scaled_height = int(height_px * scale)
                    image_placements.append((ws.title, row_idx, 3, image_path, y_off))
                    y_off += scaled_height + 10
                    total_height_px = y_off
                ws.row_dimensions[row_idx].height = max(20, int(total_height_px * 0.75))

    for ws in detail_ws_by_key.values():
        fixed_width = {2: 35, 3: 30}
        if export_asset_version_notes:
            fixed_width[4] = 10
            _autofit(ws, wrap_cols={2, 6, 7}, fixed_width=fixed_width)
        else:
            _autofit(ws, wrap_cols={2, 5, 6}, fixed_width=fixed_width)

    index_row_by_key: Dict[str, int] = {}
    root_key = f"Project:{project_id}"

    def walk(key: str):
        node = node_map.get(key)
        if not node:
            return
        parent_type, _, parent_id = node.key.partition(":")
        count = len(notes_bucket.get((parent_type, parent_id), [])) if node.has_notes else ""
        if export_asset_version_notes and node.key.startswith("AssetVersion:"):
            av_id = node.key.split(":", 1)[1]
            index_ws.append(
                [node.path_text, av_task_path_by_av_id.get(av_id, "N/A"), node.type_name, "", count]
            )
            link_col = 4
        elif export_asset_version_notes:
            index_ws.append([node.path_text, "", node.type_name, "", count])
            link_col = 4
        else:
            index_ws.append([node.path_text, node.type_name, "", count])
            link_col = 3

        row_idx = index_ws.max_row
        index_row_by_key[node.key] = row_idx
        index_ws.cell(row_idx, count_col).alignment = Alignment(
            horizontal="left",
            vertical="top",
        )
        if node.has_notes and node.sheet_name:
            link_cell = index_ws.cell(row_idx, link_col)
            link_cell.value = f'Notes on "{node.path_text}"'
            link_cell.hyperlink = f"#'{node.sheet_name}'!A1"
            link_cell.style = "Hyperlink"
            link_cell.font = Font(color="0563C1", underline="single")

        children = [node_map[child_key] for child_key in node.children if child_key in node_map]
        children.sort(key=lambda child: (child.type_name or "", child.name or ""))
        for child in children:
            walk(child.key)

    walk(root_key)

    for key, ws in detail_ws_by_key.items():
        index_row = index_row_by_key.get(key)
        if not index_row:
            continue
        back = ws["A1"]
        link_col_letter = "D" if export_asset_version_notes else "C"
        back.value = "Back to Index"
        back.hyperlink = f"#'{index_ws.title}'!{link_col_letter}{index_row}"
        back.style = "Hyperlink"
        back.font = Font(color="0563C1", underline="single", size=16)
        back.alignment = Alignment(horizontal="left", vertical="top")
        back.fill = header_fill

    if export_asset_version_notes:
        _autofit(index_ws, wrap_cols={1, 2, 4})
    else:
        _autofit(index_ws, wrap_cols={1, 3})

    outdir = tempfile.mkdtemp(prefix="notes_export_out_")
    scope = "multi" if (len(selected_context_ids) + len(selected_task_ids)) > 1 else "single"
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    out_name = _ensure_single_xlsx(f"{project_name}_notes_{scope}_{timestamp}")
    out_path = os.path.join(outdir, safe_filename(out_name))
    wb.save(out_path)
    try:
        embed_ok = _inject_images_into_xlsx(out_path, image_placements)
    except Exception:
        embed_ok = 0

    return {
        "out_path": out_path,
        "entities": len(detail_ws_by_key),
        "images_embedded": embed_ok,
        "image_files": download_ok,
    }


def _attach_export_workbook_to_job(
    session: ftrack_api.Session,
    job_id: Optional[str],
    out_path: str,
    notes_count: int,
    entity_count: int,
    embed_ok: int,
    download_ok: int,
):
    if not job_id:
        return
    try:
        server_location = session.query('Location where name is "ftrack.server"').one()
        comp_name = os.path.splitext(os.path.basename(out_path))[0]
        component = session.create_component(
            out_path,
            data={"name": comp_name},
            location=server_location,
        )
        session.create(
            "JobComponent",
            {"component_id": component["id"], "job_id": job_id},
        )
        _update_job_status(
            session,
            job_id,
            (
                f"Notes export done. Notes: {notes_count}; Entities: {entity_count}; "
                f"Images embedded: {embed_ok}; Img files: {download_ok}"
            ),
            status="done",
        )
    except Exception as error:
        logger.warning(f"Failed to attach xlsx to job: {error}")
        _update_job_status(
            session,
            job_id,
            "Notes export done (xlsx attachment failed).",
            status="done",
        )


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
    except Exception as error:
        logger.error(f"Failed to create background session: {error}")
        return

    job_id = None
    try:
        job_id = _create_export_job(bg_session, api_username)

        exporter = NotesExporter(bg_session)
        project = exporter._get_project(project_id)
        project_name = (project.get("name") or "project").strip()

        typed_contexts = exporter._fetch_all_typed_contexts(project_id)
        ctx_by_id, children_map = _build_context_maps(typed_contexts)
        all_context_ids = _compute_context_scope(
            selected_context_ids,
            children_map,
            include_descendants,
        )

        tasks_by_id: Dict[str, dict] = {}
        asset_versions: List[dict] = []
        assets_by_id: Dict[str, dict] = {}
        if export_asset_version_notes:
            asset_versions, assets_by_id, all_context_ids = _collect_asset_version_scope(
                exporter,
                bg_session,
                project_id,
                selected_task_ids,
                all_context_ids,
                ctx_by_id,
                export_asset_version_task_basis,
            )
        else:
            tasks_by_id = _collect_tasks_for_export(
                exporter,
                bg_session,
                project_id,
                all_context_ids,
                selected_task_ids,
                include_descendants,
            )

        av_task_path_by_av_id = _build_related_task_path_map(
            bg_session,
            asset_versions,
        )

        _update_job_status(bg_session, job_id, "Notes export (fetching notes)")
        parent_ids = _build_note_parent_ids(
            all_context_ids,
            tasks_by_id,
            asset_versions,
            export_asset_version_notes,
        )
        notes = (
            exporter._fetch_notes_by_parent_ids(project_id, list(parent_ids))
            if parent_ids
            else []
        )
        notes, note_by_id = _dedupe_notes(notes)
        author_name_by_id = _fetch_author_name_map(bg_session, notes)
        _attach_note_components_to_notes(exporter, notes)
        notes_bucket = _bucket_notes_by_entity(
            notes,
            all_context_ids,
            tasks_by_id,
            asset_versions,
            export_asset_version_notes,
        )
        node_map = _build_index_node_map(
            project_id,
            project_name,
            all_context_ids,
            ctx_by_id,
            tasks_by_id,
            asset_versions,
            assets_by_id,
            notes_bucket,
            export_asset_version_notes,
        )

        _update_job_status(bg_session, job_id, "Notes export (building xlsx)")
        export_stats = _build_export_workbook(
            exporter,
            project_id,
            project_name,
            selected_context_ids,
            selected_task_ids,
            export_asset_version_notes,
            notes_bucket,
            note_by_id,
            node_map,
            av_task_path_by_av_id,
            author_name_by_id,
        )
        _attach_export_workbook_to_job(
            bg_session,
            job_id,
            str(export_stats["out_path"]),
            len(notes),
            int(export_stats["entities"]),
            int(export_stats["images_embedded"]),
            int(export_stats["image_files"]),
        )
    except Exception as error:
        logger.error(f"Export failed: {error}")
        _update_job_status(
            bg_session,
            job_id,
            f"Notes export failed: {error}",
            status="failed",
        )
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
        return False

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
        task = _ThumbTask(eid, url, size, self._thumb_signals)
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


def register(session, **_kw):
    if not isinstance(session, ftrack_api.session.Session):
        return
    plugin = ftrack_connect.ui.application.ConnectWidgetPlugin(NotesExporterWidget)
    plugin.register(session, priority=20)
