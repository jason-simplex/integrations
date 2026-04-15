# 📤 Ftrack Notes Exporter (Export Notes to Excel)

Purpose: turn ftrack **Notes / Replies / image attachments** into a clean **Excel (.xlsx)** you can share with production, clients, or teammates—for handover, reviews, audits, and archiving.

***

## Overview 👀

**Ftrack Notes Exporter** is an **ftrack Connect widget**. Select a scope from the project tree and export matching Notes into a single Excel file. The export runs as an **ftrack Job**, so Connect stays responsive.

***

## Highlights ✨

- 📦 One workbook for handover & reporting (no copy/paste from ftrack pages)
- 🧾 Rich, readable details: author, labels, replies, date/time, and more
- 🖼️ Evidence included: image attachments embedded in-cell (stacked vertically, won’t overwrite other columns)
- 🧭 Fast navigation: Index → Detail and “Back to Index” links
- 🌳 Export exactly what you need: multi-select + optional “Include notes from descendants”
- 🧬 Asset Version export that still looks like ftrack: ftrack-style paths + Related Task, plus Frame No. (when available)
- ⚡ Safe to run anytime: exports as an ftrack Job (ftrack Connect / web UI stays responsive)

***

## Usage 🧭

### Open

1. Launch **ftrack Connect**
2. Open the widget **Export Notes to Excel**

### Export

1. Select a **Project**
2. Select one or more items in the tree
3. (Optional) Toggle **Include notes from descendants**
4. Click **Export**
5. Monitor progress in **Jobs** and download the `.xlsx` from the Job attachments

***

## Installation 🧩

This widget is deployed as an **ftrack Connect plugin**.

### Option A: Install via Plugin Manager (drag & drop)

1. Launch **ftrack Connect**
2. Open **Plugins** (Plugin Manager)
3. Drag the plugin `.zip` (e.g. `connect-notes-exporter-widget-26.4.0.zip`) into the Plugin Manager window
4. Click **Install**
5. Click **Restart** when prompted

### Option B: Manual install (plugins folder)

1. Download the plugin `.zip` (e.g. `connect-notes-exporter-widget-26.4.0.zip`)
2. Unzip it into your ftrack Connect plugins folder:
   - **macOS**: `~/Library/Application Support/ftrack-connect-plugins/`
   - **Windows**: `%APPDATA%\\ftrack-connect-plugins\\`
3. Restart **ftrack Connect**

***

## Troubleshooting 🧯

### I can’t see the widget

- Restart ftrack Connect
- Ask Pipeline/TD to confirm the plugin is installed in `ftrack-connect-plugins`

### Labels or images are missing

- Confirm you can view them in the ftrack web UI (permissions may apply)
- Some notes genuinely have no labels/attachments

### Export is slow

- Reduce the scope and/or disable **Include notes from descendants**

***

## FAQ 🙋

### Does this modify anything in ftrack?

No. It only reads and exports.

### Are replies included?

Yes. Replies are exported and “In reply to” shows the referenced note content.

### Can I export multiple items at once?

Yes. The export de-duplicates notes and keeps the hierarchy in the Index sheet.

### What’s included in the exported Excel?

**Index sheet**

- **Context path** for every exported item (hierarchy preserved)
- **Notes link** to jump to the corresponding detail sheet
- **Notes count**
- (Asset Version mode) **Related Task** path when available, otherwise `N/A`

**Detail sheets**

- One sheet per entity that has notes
- Columns include:
  - **DateTime**
  - **Note Contents**
  - **Image Attachments**
  - **Author Name**
  - **Labels**
  - **In reply to** (shows the referenced note’s content)
  - (Asset Version mode) **Frame No.**, or `N/A`

