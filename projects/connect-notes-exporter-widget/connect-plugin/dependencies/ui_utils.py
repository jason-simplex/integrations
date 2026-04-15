# :coding: utf-8
# :copyright: Copyright (c) 2025 bro.tiger
import os
import logging
from ftrack_connect.qt import QtWidgets, QtCore, QtGui
try:
    import qtawesome as qta
except Exception:
    qta = None

DEBUG_UI = os.getenv('FTRACK_DEBUG_UI', '0') == '1'

def ui_log(msg):
    if DEBUG_UI:
        logging.getLogger('batch_import_thumbnails').info(msg)

def get_dropped_excel_path(mime):
    if mime is None:
        return None
    try:
        if mime.hasUrls():
            for u in mime.urls():
                p = u.toLocalFile()
                if p and p.lower().endswith(('.xlsx', '.xlsm', '.xls')):
                    return p
        if mime.hasText():
            t = mime.text().strip()
            if t and t.lower().endswith(('.xlsx', '.xlsm', '.xls')):
                return t
    except Exception:
        return None
    return None

def qta_icon_pixmap(name, size=24):
    if qta is None:
        return None
    try:
        return qta.icon(name).pixmap(size, size)
    except Exception:
        return None

def set_button_icon_qta(btn, name):
    if qta is None:
        return
    try:
        btn.setIcon(qta.icon(name))
    except Exception:
        pass

def show_info(parent, title, text):
    try:
        QtWidgets.QMessageBox.information(parent, title, text)
    except Exception:
        pass

def show_warn(parent, title, text):
    try:
        QtWidgets.QMessageBox.warning(parent, title, text)
    except Exception:
        pass

