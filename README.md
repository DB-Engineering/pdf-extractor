# PDF to Text & Multi-Tab Excel Extractor

A Python utility for extracting plain text and structured tabular data from PDF files (such as HVAC equipment schedules, engineering drawings, and specification sheets) into formatted, multi-tab Microsoft Excel (`.xlsx`) spreadsheets.

---

## Directory Structure

The project organizes inputs and outputs into dedicated folders:

```
root/
├── pdf_to_excel.py    # Main Python converter script
├── README.md          # Documentation
├── pdf_files/        # Input PDF files (e.g. page_1.pdf)
├── text_files/       # Extracted plain text files (e.g. page_1.txt)
└── excel_files/      # Extracted Excel spreadsheets (e.g. page_1.xlsx)
```

> **Note:** The script automatically creates `pdf_files/`, `text_files/`, and `excel_files/` if they do not already exist on startup.

---

## Features

- **Multi-Tab Excel Spreadsheets**: Parses separate tables and schedules (such as `1ST FLOOR` and `3RD FLOOR` schedules) and places each on its own styled Excel tab.
- **Smart Data Formatting**: Converts numbers and flow rates into numeric values (`#,##0` / `0.0`), centers status indicators (`X`, `-`), freezes top header rows (`A2`), applies alternating zebra striping, and auto-fits column widths.
- **Dual PDF Extraction Engines**: Uses `pdfplumber` for geometric table extraction and `pypdf` for fast text reading.
- **Text File Fallback Parsing**: Parses pre-extracted `.txt` files using regex pattern anchors when geometric grid lines are absent.
- **Batch Processing**: Scans `pdf_files/` automatically and converts all PDF files in bulk when run without arguments.

---

## Prerequisites & Installation

Required Python packages:
- `pypdf`
- `pdfplumber`
- `openpyxl`

### Installation via `pip`
```bash
pip install pypdf pdfplumber openpyxl --break-system-packages
```

### Installation via `apt` (Debian / Ubuntu)
```bash
sudo apt update && sudo apt install python3-pypdf python3-pdfplumber python3-openpyxl
```

---

## Usage

### 1. Process a Single File
Pass a file name directly. The script checks your current directory, then `pdf_files/` and `text_files/`:

```bash
python3 pdf_to_excel.py page_1.pdf
```
*Creates `text_files/page_1.txt` and `excel_files/page_1.xlsx`.*

### 2. Batch Process All PDFs
Run without arguments to convert every PDF in `pdf_files/`:

```bash
python3 pdf_to_excel.py
```

### 3. Extract Excel Only
Skip plain text output and generate only `.xlsx`:

```bash
python3 pdf_to_excel.py page_1.pdf --excel-only
```

### 4. Extract Text Only
Skip Excel generation and extract only `.txt`:

```bash
python3 pdf_to_excel.py page_1.pdf --text-only
```

### 5. Convert Pre-Extracted Text to Excel
Extract tabular data from an existing text file in `text_files/`:

```bash
python3 pdf_to_excel.py page_1.txt
```

---

## Command-Line Options

| Argument | Description |
| :--- | :--- |
| `file_path` | *(Optional)* Input file name or path. If omitted, scans `pdf_files/`. |
| `-o`, `--output` | Custom file path for extracted text output (`.txt`). |
| `--excel-output` | Custom file path for extracted Excel output (`.xlsx`). |
| `--excel-only` | Generates only the Excel file, skipping text output. |
| `--text-only` | Generates only the text file, skipping Excel output. |
| `-e`, `--engine` | PDF extraction engine (`auto`, `pypdf`, `pdfplumber`). Default: `auto`. |
| `-q`, `--quiet` | Suppress console status messages. |

---

## How Table Extraction Works

1. **Native PDF Geometry (`pdfplumber`)**: Detects vector lines and cell intersections to group text into rectangular grids.
2. **Pattern Matching Fallback**: Scans text line-by-line for equipment row patterns (`VAV-`) and splits tokens into standard schedule columns (`UNIT MARK`, `MANUFACTURER`, `MODEL`, `INLET SIZE`, `CFM MAX`, `CFM MIN`, `CFM HEAT`, `ROWS`, `GPM`, `3 WAY VALVE`, `CO2 SENSOR`, `OCC SENSOR`, `REMARKS`).
3. **Section Anchor Splitter**: Groups rows under schedule heading anchors (`SINGLE DUCT & REHEAT TERMINAL SCHEDULE`) and creates distinct worksheet tabs for each section while filtering out stamps, notes, and margin text.
