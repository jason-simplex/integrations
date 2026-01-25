<!--
Copyright (c) 2025 bro.tiger
All rights reserved.
-->

# Batch Import Entity Thumbnails From Excel 📸

Batch-import thumbnails for any ftrack entities (Shots, Tasks, Assets, Folders, etc.) — simple, visual, and fast.

## Highlights ✨

- **Generic Entity Support**: Works with any ftrack entity type, not just Tasks 🌍
- **Flexible Context Builder**: Define your own hierarchy rules (e.g., "Sequence/Shot" or "Folder/AssetBuild") 🔗
- **ftrack Connect Integration**: Runs seamlessly as a tab widget in ftrack Connect 🧩
- **Excel Native**: Extracts embedded images directly from `.xlsx` files 🖼️
- **Safe & Focused**: Only updates thumbnails; leaves other entity data untouched 🔒

## Usage 🚀

1. **Prepare your Data**: Create an Excel (.xlsx) file where each row represents an entity. Embed your thumbnail images directly into cells.
2. **Load File**: Launch the tool in ftrack-connect and drop your `.xlsx` file into the input field 📂
3. **Configure Import**:
   - Select the target **Project**.
   - Choose the **Import Sheet** containing your data.
   - Select the **Image Column** (the column containing the embedded pictures) 🖼️
4. **Build Entity Context**:
   - Use the "Entity Context Builder" to define how to find your entities.
   - Drag and drop columns from "Available Columns" to the builder list.
   - _Example_: For a Shot under a Sequence, add "Sequence" then "Shot". The last column determines the target entity. 🎯
5. **Run Import**: Click “Import” to process the file ▶️
6. **Verify**: Check the job status and see your new thumbnails in ftrack! ✅

## Notes 📝

- **File Format**: Must be `.xlsx` (Excel Workbook). CSV files do not support embedded images.
- **Image Extraction**: Images must be **embedded objects** (floating pictures) inside the Excel sheet, not file paths or formulas.
- **Hierarchy Logic**: The "Entity Context" defines the parent chain. If your context is `Folder / AssetBuild`, the tool looks for a Folder by name, then an AssetBuild inside that Folder.
- **Performance**: Very large files with high-res images may take time. Consider splitting files if necessary.

## Troubleshooting 💡

- **"Entity not found"**:
  - Check your "Entity Context" rule. Does it exactly match your project structure?
  - Ensure the names in Excel match ftrack names exactly (case-sensitive).
- **"No data found"**:
  - Ensure your Excel sheet has headers in the first row.
- **Images not importing**:
  - Verify that the images are actually embedded in the Excel file and anchored to the correct cells.
  - Ensure the correct "Image Column" is selected.

## FAQ ❓

- **Do I need Excel installed?**
  - No. The tool reads `.xlsx` files directly.
- **What entities can I update?**
  - Any entity that can be located via a parent chain (Shots, Sequences, Tasks, Folders, AssetBuilds, etc.).
- **Can I use file paths instead of embedded images?**
  - No, this tool is designed specifically for visual batching using embedded Excel images.
- **Is it safe to run multiple times?**
  - Yes. Re-importing will simply overwrite the existing thumbnails with the new ones.

Happy importing! 🎉
