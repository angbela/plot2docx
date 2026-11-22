# Plot to Docx

A native Windows desktop application to batch index plots (from local drives or network UNC paths, e.g. `\\WARTEG-PC\...`), organize and group them into custom sections with manual sequence control, customize page layouts, and export directly to Microsoft Word (`.docx`) and high-fidelity `.pdf`.

---

## What's New & Architecture Overview

The application is structured into a modern 3-step workflow with dedicated modules:

```text
plot2docx/
├── app.py              # Main desktop application & tab navigation
├── plot_selector.py    # Step 1: Disk/UNC recursive scan & hierarchy tree selection
├── plot_organizer.py   # Step 2: Custom section cards, sequencing & bulk operations
├── doc_generator.py    # Step 3: Page layout, document rendering & Word/PDF export
└── pdf_images.py       # High-fidelity lossless PDF image restoration engine
```

---

## Key Features

### 1. Plot Selection (`plot_selector.py`)
- **Network UNC & Local Path Support**: Recursively scan local folders or network shares (`\\Server\Projects\...`) without uploading or copying files over HTTP.
- **Interactive Checkbox Tree**: Select or deselect entire folders or individual plots with automatic cascading checks and file size indicators.
- **Fast Search & Tree Filtering**: High-performance directory walking supporting PNG, JPG, JPEG, BMP, TIFF, GIF, and WEBP.

### 2. Organize, Group & Sequence (`plot_organizer.py`)
- **Grouping Strategies**:
  - **Manual Grouping**: Full flexibility to create, rename, reorder, and delete sections.
  - **Auto-Group by Folder**: Automatically generates sections matching the source directory structure.
  - **No Grouping (Flat)**: Keeps all selected plots in a continuous flat report.
- **Section Cards & Page Breaks**:
  - Add and rename sections dynamically.
  - Toggle **"Page Break Before"** per section (ideal for engineering reports).
  - Move entire sections up (`▲`) and down (`▼`).
- **Manual Sequencing & Bulk Operations**:
  - Reorder plots individually within sections using `▲` / `▼` or transfer via `Move to...`.
  - **Bulk Toolbar**: Select All / Deselect All, bulk `▲ Move Up` / `▼ Move Down`, and bulk `Move to Section...` using multi-selection checkboxes.
  - **Live Figure Numbering Preview**: Real-time display of section-based figure numbering (e.g., `Figure 1.1`, `Figure 1.2`, `Figure 2.1`...).

### 3. Document & PDF Generator (`doc_generator.py`)
- **Layout & Column Formatting**:
  - Orientation toggle (**Landscape** / **Portrait**).
  - Multi-column side-by-side grid (**1 Column**, **2 Columns**, **3 Columns**) using borderless tables.
  - Dimension limits (**Max Width** and **Max Height** in cm) with **locked aspect ratio**.
- **Report Headings & Section Styling**:
  - Formats section headers cleanly (14pt Bold Segoe UI).
  - Keeps figure images and captions (`Figure X.Y\t{full_path}`) locked to the same page without splitting.
- **Lossless High-Resolution PDF Export**:
  - Microsoft Word renders the layout, and embedded original-resolution images (including alpha transparency) are restored losslessly in the exported PDF.

---

## Running the Application

### 1. Standalone Executable (No Python Required)
Run `dist/PlotToDocx.exe`.

### 2. From Source
```bash
# Clone the repository
git clone https://github.com/angbela/plot2docx.git
cd plot2docx

# Install dependencies
pip install -r requirements.txt

# Launch application
python app.py
```

### 3. Build Single Executable (.exe)
```bash
pyinstaller --noconsole --onefile --collect-all customtkinter --hidden-import=win32com --hidden-import=win32com.client --hidden-import=pythoncom app.py -n "PlotToDocx"
```

---

## Requirements
- Windows 10/11; Microsoft Word must be installed for PDF export.
- Python 3.10+ (if running from source).
- Dependencies: `customtkinter`, `python-docx`, `pillow`, `pypdf`, `pywin32`.

*Note: PDF export requires Word to create a tagged PDF. If its output cannot be matched unambiguously to the document's plots, the application reports an error instead of saving a potentially incorrect PDF. The Word document (`.docx`) remains available.*
