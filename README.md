# Plot to Docx

A local Windows desktop web application (packaged into a single `.exe`) to browse plots from local drives or network UNC paths (e.g. `\\Server\Projects`), organize and drag-and-drop them into sections and subsections without UI flicker, save/reload reusable report configurations, and export directly to Microsoft Word (`.docx`), lossless high-resolution `.pdf`, and scrollable `.html` reports.

---

## Architecture Overview

```text
plot2docx/
├── app.py              # Local web app launcher & native Windows file dialog host
├── server.py           # 127.0.0.1 HTTP API, thumbnail cache & background job runner
├── scanner.py          # Fast os.scandir() lazy folder browser & cancellable guarded scanner
├── config_store.py     # Save, load, validate & source-remap project configs (.plot2docx.json)
├── doc_generator.py    # Word (.docx) layout & report builder
├── html_generator.py   # Scrollable HTML5 report builder (network-linked or base64-embedded)
├── pdf_images.py       # High-fidelity lossless PDF image restoration engine
└── web/                # Self-contained offline HTML/CSS/JS frontend
    ├── index.html
    ├── styles.css
    └── app.js
```

---

## Key Features

### 1. Plot Selection (Nested Explorer Tree & Guarded Deep Scan)
- **Open Folder (Lazy Tree)**: Loads the target directory immediately via `os.scandir()` and expands subfolders only when clicked; safe even on huge network shares (`\\Server\...`), showing only image files.
- **Deep Scan with Guardrail**: Recursively indexes subfolders into a nested Windows Explorer-style tree with an instant **Cancel Scan** button and a 500-image safety pause (`Continue Full Scan`, `Keep Loaded`, or `Cancel`).
- **Chunked Large-Folder Rendering**: Folders with hundreds or thousands of plots render in smooth 200-item chunks with search/filter and inline image preview.
- **Smart Auto-Select**: Auto-selects folders with `<= 100` images, and leaves larger folders unchecked so huge scans never flood the organizer.

### 2. Zero-Flicker Drag-and-Drop Organizer
- **Smooth Drag-and-Drop**: Drag individual plots, multi-checked plot batches, or entire sections/subsections using the drag handles with insertion indicators.
- **Up to 2 Subsection Levels**: Create nested subsections (`Section 1` -> `1.1` -> `1.1.1`) in **Manual Grouping** while figures follow Level-1 section numbering (`Figure 1.1`, `Figure 1.2`, resetting at `Section 2`).
- **In-Place DOM Updates**: Reordering, checking, renaming, and moving plots mutate DOM elements in place with no full reload flashes and no scroll jumps.
- **Grouping Strategies**: **Manual Grouping**, **Auto-Group by Folder**, and **No Grouping (Flat)**, plus per-section **Page Break Before**, optional lazy **Thumbnails**, and a sticky **Bulk Action Bar**.
- **Non-Destructive Sync**: Going back to Step 1 to check or uncheck a few plots preserves your existing manual section names, subsections, page breaks, and ordering.

### 3. Save / Load / Remap Configurations
- **Save Config (`*.plot2docx.json`)**: Saves your source folder, section/subsection structure, plot order, page breaks, layout options, and last output path.
- **Replotting Workflow**: When plots are regenerated with the same filenames, load the saved config, click **Re-check Files** (or **Change Source Folder** to remap relative paths to a new directory), and click **Quick Regenerate to Last Output**.
- **Recent Configs**: Quick dropdown menu next to **Load Config**.

### 4. Word, Lossless PDF & HTML Report Generator
- **Output Formats**: Export to Word (`.docx`), scrollable HTML (`.html`), or both, plus optional lossless high-resolution `.pdf` generated from Word when `.docx` is enabled.
- **HTML Report Features**: Links directly to network (`\\Server\...`) or local disk images with lazy loading, collapsible Table of Contents, click-to-zoom lightbox, and Light/Dark theme toggle (plus an optional **Embed images inside HTML** mode for single-file offline sharing).
- **Layout Options**: Portrait/Landscape orientation, 1/2/3 columns, Max Width/Height (cm) with locked aspect ratio in all column modes.

---

## Running the Application

### 1. Standalone Executable (No Python Required)
Run `dist/PlotToDocx.exe` (opens in an Edge/Chrome app window or your default browser).

### 2. From Source
```bash
pip install -r requirements.txt
python app.py
```

Useful CLI flags:
- `python app.py --browser`: open in your default browser tab instead of an app window
- `python app.py --port 8765 --no-open --no-auto-exit`: run headless server for debugging

### 3. Build Single Executable (.exe)
```bash
pyinstaller PlotToDocx.spec
```
Or directly:
```bash
pyinstaller --noconsole --onefile --add-data "web;web" --hidden-import=win32com --hidden-import=win32com.client --hidden-import=pythoncom app.py -n "PlotToDocx"
```

---

## Requirements
- Windows 10/11; Microsoft Word must be installed for PDF export.
- Python 3.10+ (if running from source).
- Dependencies: `python-docx`, `pillow`, `pypdf`, `pywin32`, `pyinstaller`.
