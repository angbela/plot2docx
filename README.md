# Plot to Docx

A native Windows desktop application to batch index hydrodynamic plots (from local drives or network UNC paths, e.g. `\\HK-NC01\...`), select them via an interactive folder tree hierarchy with checkboxes, customize page layouts, and export them directly to Microsoft Word (`.docx`) and high-fidelity `.pdf`.

---

## Features
- **Network UNC & Local Path Support**: Recursively scan local folders or network shares (`\\Server\Projects\...`) without uploading or copying files across HTTP.
- **Interactive Checkbox Tree**: Select/deselect entire scenario folders or individual plots with automatic cascading checks.
- **Layout & Column Formatting**:
  - Orientation toggle (**Landscape** / **Portrait**).
  - Multi-column side-by-side grid (**1 Column**, **2 Columns**, **3 Columns**) using borderless tables.
  - Sizing limits (**Max Width** and **Max Height** in cm) with **locked aspect ratio**.
- **Page-Break Protection**: Keeps image and its corresponding caption (`Figure X\t{full_path}`) on the same page.
- **Direct PDF Export**: One-click option to export matching PDF via Microsoft Word COM engine.
- **Standalone `.exe`**: Ready to run on Windows without installing Python or dependencies.

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
- Windows 10/11
- Python 3.10+ (if running from source)
- Dependencies: `customtkinter`, `python-docx`, `pillow`, `pywin32`
