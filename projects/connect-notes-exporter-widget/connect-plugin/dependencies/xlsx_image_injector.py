# :coding: utf-8
"""Utilities to inject images into an existing .xlsx by editing OOXML parts.

This module is intentionally UI- and ftrack-agnostic so it can be reused by other
Connect widgets that need to embed images without Pillow.
"""

from __future__ import annotations

import os
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from typing import Dict, List, Set, Tuple


def _emu(px: int) -> int:
    # English Metric Units (EMU): 1 px ~= 9525 EMUs (Excel/OOXML convention)
    return int(px * 9525)


def looks_like_image_header(head: bytes) -> bool:
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


def get_image_size(path: str) -> Tuple[int, int]:
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
                    if marker in (
                        b"\xc0",
                        b"\xc1",
                        b"\xc2",
                        b"\xc3",
                        b"\xc5",
                        b"\xc6",
                        b"\xc7",
                        b"\xc9",
                        b"\xca",
                        b"\xcb",
                        b"\xcd",
                        b"\xce",
                        b"\xcf",
                    ):
                        _ = int.from_bytes(f.read(2), "big")  # segment length
                        _ = f.read(1)  # precision
                        h = int.from_bytes(f.read(2), "big")
                        w = int.from_bytes(f.read(2), "big")
                        return (max(w, 1), max(h, 1))
                    ln = int.from_bytes(f.read(2), "big")
                    if ln < 2:
                        break
                    f.seek(ln - 2, 1)
        except Exception:
            return (240, 240)

    return (240, 240)


def inject_images_into_xlsx(xlsx_path: str, placements: List[tuple]) -> int:
    """Inject images into an existing xlsx without Pillow/openpyxl image APIs.

    placements: list of:
      - (sheet_title, row_1based, col_1based, image_path)
      - (sheet_title, row_1based, col_1based, image_path, y_offset_px)

    Returns number of images injected.
    """
    if not placements:
        return 0

    # Group by sheet title.
    by_sheet: Dict[str, List[Tuple[int, int, str, int]]] = defaultdict(list)
    for p in placements:
        if len(p) == 4:
            sheet_title, r, c, img_path = p
            y_off = 0
        else:
            sheet_title, r, c, img_path, y_off = p
        by_sheet[str(sheet_title)].append(
            (int(r), int(c), str(img_path), int(y_off or 0))
        )

    tmp_out = xlsx_path + ".tmp"

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
        wb_root = ET.fromstring(zin.read("xl/workbook.xml"))
        sheets_el = wb_root.find("w:sheets", NS)
        if sheets_el is None:
            return 0

        title_to_sheet_path: Dict[str, str] = {}
        wb_rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {
            rel.get("Id"): rel.get("Target")
            for rel in wb_rels.findall("rel:Relationship", NS)
        }

        def _norm_target(t: str) -> str:
            t = (t or "").strip()
            if t.startswith("/"):
                t = t[1:]
            if t.startswith("xl/"):
                return t
            if t.startswith(("worksheets/", "drawings/", "media/")):
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
        while any(
            n.startswith(f"xl/media/image{next_image}.") for n in existing
        ):
            next_image += 1

        ct_root = ET.fromstring(zin.read("[Content_Types].xml"))

        def _ensure_default(ext: str, ctype: str):
            for d in ct_root.findall("ct:Default", NS):
                if (d.get("Extension") or "").lower() == ext.lower():
                    return
            ET.SubElement(
                ct_root,
                f"{{{NS['ct']}}}Default",
                {"Extension": ext, "ContentType": ctype},
            )

        def _ensure_override(part_name: str, ctype: str):
            for o in ct_root.findall("ct:Override", NS):
                if o.get("PartName") == part_name:
                    return
            ET.SubElement(
                ct_root,
                f"{{{NS['ct']}}}Override",
                {"PartName": part_name, "ContentType": ctype},
            )

        sheet_paths_to_modify: Set[str] = set()
        sheet_rels_to_modify: Set[str] = set()
        for title in by_sheet.keys():
            sp = title_to_sheet_path.get(title)
            if not sp:
                continue
            sheet_paths_to_modify.add(sp)
            sheet_rels_to_modify.add(
                sp.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"
            )

        skip_copy = (
            set(sheet_paths_to_modify)
            | set(sheet_rels_to_modify)
            | {"[Content_Types].xml"}
        )

        with zipfile.ZipFile(
            tmp_out, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
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

                sheet_rels_path = (
                    sheet_path.replace("xl/worksheets/", "xl/worksheets/_rels/")
                    + ".rels"
                )
                if sheet_rels_path in existing:
                    rel_root = ET.fromstring(zin.read(sheet_rels_path))
                else:
                    rel_root = ET.Element(f"{{{NS['rel']}}}Relationships")

                drawing_idx = next_drawing
                next_drawing += 1
                drawing_name = f"drawing{drawing_idx}.xml"
                drawing_path = f"xl/drawings/{drawing_name}"
                drawing_rels_path = f"xl/drawings/_rels/{drawing_name}.rels"

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

                if sheet_root.find("w:drawing", NS) is None:
                    ET.SubElement(
                        sheet_root,
                        f"{{{NS['w']}}}drawing",
                        {f"{{{NS['r']}}}id": rid},
                    )

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

                    w_px, h_px = get_image_size(img_file)
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
                    ET.SubElement(
                        anchor,
                        f"{{{NS['xdr']}}}ext",
                        {"cx": str(cx), "cy": str(cy)},
                    )

                    pic = ET.SubElement(anchor, f"{{{NS['xdr']}}}pic")
                    nv = ET.SubElement(pic, f"{{{NS['xdr']}}}nvPicPr")
                    ET.SubElement(
                        nv,
                        f"{{{NS['xdr']}}}cNvPr",
                        {"id": str(pic_id), "name": media_name},
                    )
                    ET.SubElement(nv, f"{{{NS['xdr']}}}cNvPicPr")
                    blipFill = ET.SubElement(pic, f"{{{NS['xdr']}}}blipFill")
                    ET.SubElement(
                        blipFill,
                        f"{{{NS['a']}}}blip",
                        {f"{{{NS['r']}}}embed": img_rid},
                    )
                    ET.SubElement(blipFill, f"{{{NS['a']}}}stretch")
                    spPr = ET.SubElement(pic, f"{{{NS['xdr']}}}spPr")
                    xfrm = ET.SubElement(spPr, f"{{{NS['a']}}}xfrm")
                    ET.SubElement(xfrm, f"{{{NS['a']}}}off", {"x": "0", "y": "0"})
                    ET.SubElement(
                        xfrm,
                        f"{{{NS['a']}}}ext",
                        {"cx": str(cx), "cy": str(cy)},
                    )
                    ET.SubElement(spPr, f"{{{NS['a']}}}prstGeom", {"prst": "rect"})
                    ET.SubElement(anchor, f"{{{NS['xdr']}}}clientData")

                    pic_id += 1
                    injected += 1

                zout.writestr(
                    drawing_path,
                    ET.tostring(wsDr, encoding="utf-8", xml_declaration=True),
                )
                zout.writestr(
                    drawing_rels_path,
                    ET.tostring(drawing_rels, encoding="utf-8", xml_declaration=True),
                )
                _ensure_override(
                    f"/xl/drawings/{drawing_name}",
                    "application/vnd.openxmlformats-officedocument.drawing+xml",
                )

                zout.writestr(
                    sheet_path,
                    ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True),
                )
                zout.writestr(
                    sheet_rels_path,
                    ET.tostring(rel_root, encoding="utf-8", xml_declaration=True),
                )

            zout.writestr(
                "[Content_Types].xml",
                ET.tostring(ct_root, encoding="utf-8", xml_declaration=True),
            )

    os.replace(tmp_out, xlsx_path)
    return injected

