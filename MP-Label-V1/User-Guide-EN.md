# MP Label-V1 User Guide

This guide covers the packaged `MP Label-V1.exe` / `MP Label-V1_中文版.exe` and the Python sources `MP Label-V1.py` / `MP Label-V1_中文版.py`.

> 中文版：[使用说明.md](使用说明.md)

## Why this tool exists

MP Label-V1 was designed around a practical microscopy workflow, not just drawing marks on images.

1. **Keep a complete photographic record of every sample.** Photograph all samples in one session and keep the images as a permanent record. The image set can be archived, re-checked and shared later without putting the slides back under the microscope.
2. **Use the microscope only when it is really needed.** Image capture is the part that requires the instrument. Annotation, classification, comparison and export can be done later on an ordinary PC, so the microscope is not occupied while the records are being organised.
3. **Use the remark field mainly for the plastic category.** Remarks are intended primarily for the microplastic category/species, plus a short note that helps later identification. Keeping this text consistent makes the exported tables much easier to group and filter.
4. **Triage unknown particles quickly.** The main purpose of the marking/screenshot record is fast classification of unknowns. If the same unknown particle appears at many positions—for example, ten occurrences of the same morphology—mark the occurrences during the scan. Later, use the exported Excel/Word record to compare them and choose one representative for FTIR/Raman or another confirmatory test. There is usually no need to test all ten.

In short: photograph once, annotate offline, record the type in the remark, and let the exported record tell you which representative particles deserve confirmatory testing.

## 0. Starting the program

This package provides 4 program files. Choose by language and by how you want to run it:

| Download | Language | Who it is for | Python required |
| --- | --- | --- | --- |
| `MP Label-V1.exe` | English | Double-click and use | No |
| `MP Label-V1_中文版.exe` | Chinese | Double-click and use | No |
| `MP Label-V1.py` | English | Run/modify the source | Yes |
| `MP Label-V1_中文版.py` | Chinese | Run/modify the source | Yes |

### Option 1: single-file exe (recommended for end users)

1. Download the exe for your language.
2. Put it in a **writable folder** (for example Desktop or Downloads; not `C:\Program Files`).
3. Double-click it.

Notes:

- The exe is a complete single-file build. Tcl/Tk and the python-docx template are included; you do **not** need Python or an `_internal` folder.
- On startup it extracts `_MEI...` next to the exe and removes it on a normal close. If the process is killed, a `_MEI...` folder may remain; delete it manually.
- If Windows shows “Unknown publisher”, choose “More info → Run anyway” (the program is not code-signed).
- If your company security policy blocks extraction in local folders, use the Python source option below.

### Option 2: Python source

1. Download the source for your language: `MP Label-V1.py` (English) or `MP Label-V1_中文版.py` (Chinese).
2. Install the dependencies and run:

```bat
python -m pip install -r requirements.txt
python "MP Label-V1.py"
:: Chinese version
python "MP Label-V1_中文版.py"
```

Requirements: Windows 10/11, Python 3.10+, with Tcl/Tk available (the python.org installer includes it by default).


## 1. Interface overview

The main window has four toolbar rows:

1. **Images and files**: Add, Delete, Prev, Next, Save, Clear, Undo, Redo, Reset View, Open Proj, Save Proj.
2. **Modes**: Annotation Mode, Rectangle Lasso, Polygon Lasso.
3. **Depth and IDs**: Depth, Renumber, Continue Num, Delete ID, Edit Remark, Edit Colour.
4. **Export**: Excel Detail, Excel Stats, Word Report(All), Word Export Test.

Status bar: current image, depth and total image count. The canvas shows the current mode and mouse controls. Annotation, batch, ID/remark/colour and export-choice popups are centered on the screen.

The Chinese build translates buttons, dialogs, status messages and export headers. Type names (`fibre`, `film`, ...), judgments (`MP` / `NoMP` / `Test`) and Test subtypes (`label` / `unlabel`) remain in English so projects and exports are compatible between both builds.

## 2. Adding images and setting depth

1. Click **Add** and choose one or more images. Supported formats: `jpg`, `jpeg`, `png`, `bmp`, `tif`, `tiff`.
2. Use **Prev / Next** to switch images.
3. Click **Depth** and enter the depth for the current image (for example `0–1 cm`, `10–11 cm`). Images at the same depth share one ID sequence.

> Keep depth text consistent (separators and spaces). Otherwise Excel/Word may treat `10cm` and `10 cm` as different groups.

## 3. Single annotation

In **Annotation Mode**:

- **Left-click** a particle to open the dialog:
  - Type: `fibre`, `film`, `fragment`, `foram`, `bead`, `pellet`, `custom`;
  - Judgment: `MP`, `NoMP`, `Test`;
  - For `Test`, choose `label` / `unlabel`;
  - **Colour**: transparent / white / black / blue / red / yellow / green / brown / grey / other, or leave it unset;
  - **Remark**: mainly for the microplastic category/species, plus any short note needed later.
- **Ctrl + left-click** near an existing mark to delete it (the hit radius follows the zoom level).
- IDs increase within the current depth; switching to another image at the same depth continues the same sequence.
- Use **Edit Colour** or **Edit Remark** to change an existing mark by ID.

### Recording repeated or unknown particles

The tool is intended for fast triage, not for identifying every particle on the spot.

1. When the same unknown morphology appears repeatedly, mark each occurrence (or at least the representative positions) while scanning.
2. Give the group consistent text in Type/Remark, for example a temporary label such as `unknown-A`, so it is easy to filter later.
3. Export **Excel Detail** / **Word Report(All)**. Use the type/remark columns to compare the occurrences and choose one representative for confirmatory identification.
4. If you need a visual record for a report or discussion, use **Save** to write an annotated PNG; the IDs, positions and remark remain in the project and exports.

This keeps the initial scan fast while preserving enough information to decide later which particles actually need instrument time.

## 4. Rectangle / polygon lasso batch annotation

### Rectangle lasso

1. Click **Rectangle Lasso**.
2. Hold the left mouse button on the canvas and drag a rectangle; release.
3. In the dialog, set count, type, judgment, Test subtype, colour and remark.
4. Click **Add Annotations**.

The program creates the requested number of marks at random positions inside the rectangle and saves the region for **Word Export Test**.

### Polygon lasso

1. Click **Polygon Lasso**.
2. Left-click the polygon vertices in order.
3. Right-click to finish.
4. In the dialog, set count, type, judgment, Test subtype, colour and remark.
5. Click **Add Annotations and Save Area**.

If the polygon is small or very narrow, the program may not find enough random positions and will report how many were actually created.

> Lasso regions are bound to “current image + depth”. Changing the depth updates the region depth; clearing annotations or deleting an image also removes the corresponding regions.

## 5. ID and remark management

- **Renumber**: renumbers all marks under the current depth as 1..N, ordered by image path then original order, and remaps remark keys.
- **Continue Num**: shows the current maximum ID and the next available ID without changing data.
- **Delete ID**: enter an ID to delete that mark from the current image.
- **Edit Remark**: enter an ID to change its remark/subtype; leave it empty to remove the remark.
- **Edit Colour**: enter an ID to change the particle colour; choose “Not set” to clear it.
- **Clear**: removes all marks on the current image (and its lasso regions).
- **Undo / Redo**: 50 snapshots covering marks, colour, depth, remarks, lasso regions and image deletion. Renumber can be undone.

## 6. Zoom and view

- **Mouse wheel** zooms around the cursor.
- **Middle-button drag** pans.
- **Reset View** fits the image to the window.
- At low zoom only IDs are shown to avoid overlapping text; at higher zoom the type/judgment, colour and remark are shown.

## 7. Saving images and projects

### Annotated PNG

Click **Save** (or `Ctrl+Shift+S`) to draw the current marks onto the original image and write `<original-name>_annotated.png` in the same folder, keeping the original resolution.

After saving, the program asks whether to add the annotated image to the current list:
- Default **No**: only the PNG is written.
- **Yes**: the derived image is added as a separate image, with new IDs and copied remarks/colour, and it will be counted again in exports.

### Project files

- **Save Proj** (`Ctrl+S`): writes a `.json` project containing image paths, depth, marks, colour, remarks, lasso regions and the current index; image binaries are not included.
- **Open Proj** (`Ctrl+O`): opens a project; if there are unsaved changes, the program asks first.
- Autosave: `%LOCALAPPDATA%\ENVer\autosave.json`; after an abnormal close, the next start asks whether to restore it.
- Error log: `%LOCALAPPDATA%\ENVer\error.log`.

> Projects store absolute image paths. After moving to another computer or folder, some images may be reported missing and will be skipped together with their marks.

## 8. Export

- **Excel Detail**: one row per mark; columns are Depth, Item ID, Type, Judgment, Subtype/Remark, Test Subtype, Colour.
- **Excel Stats**: type counts per depth plus Total.
- **Word Report(All)**: full-image report grouped by depth; images keep their full resolution, and each image is preceded by an image-name/count/ID-range table.
- **Word Export Test**: exports lasso regions by MP / NoMP / Test; Test is split into `Test-Label` / `Test-Unlabel` columns, and crops keep the original resolution.

Exports run in a background thread with a progress window. When finished, the program shows the save path. On failure, it shows the error and writes it to the log.

**Using the export for triage:** filter/group the Excel Detail by Type and Remark to see all positions of the same unknown category, then choose a representative for confirmatory testing. Word Report(All) gives the full visual context for that decision.

## 9. Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| `Ctrl+Z` / `Ctrl+Y` | Undo / redo |
| `Ctrl+S` | Save project |
| `Ctrl+O` | Open project |
| `Ctrl+Shift+S` | Save annotated PNG |
| `Ctrl+R` | Reset view |
| `Esc` | Cancel an in-progress lasso |
| `Ctrl+left-click` | Delete a nearby mark |
| Mouse wheel | Zoom |
| Middle-button drag | Pan |
| Right-click (polygon mode) | Finish the polygon |

---

Chinese version: [使用说明.md](使用说明.md)
