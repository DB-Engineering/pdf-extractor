#!/usr/bin/env python3
"""
PDF & Table Extractor (Interactive & Targeted Table/Column Extractor)
=====================================================================
Extracts a specific table and requested column subset from a PDF file
and exports it to a formatted Excel spreadsheet (.xlsx).

Directory Structure:
--------------------
root/
├── pdf_to_excel.py
├── pdf_files/        (Input PDFs, e.g. page_1.pdf)
├── text_files/       (Output/Input text files, e.g. page_1.txt)
└── excel_files/      (Output Excel files, e.g. page_1.xlsx)

Usage:
    python3 pdf_to_excel.py
"""

import re
import sys
import time
from pathlib import Path

# Try importing PDF extraction libraries
try:
    import pypdf
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

# Try importing openpyxl for Excel output
try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


def setup_directories(base_dir: Path) -> tuple[Path, Path, Path]:
    """Ensures pdf_files, text_files, and excel_files directories exist."""
    pdf_dir = base_dir / "pdf_files"
    text_dir = base_dir / "text_files"
    excel_dir = base_dir / "excel_files"

    pdf_dir.mkdir(parents=True, exist_ok=True)
    text_dir.mkdir(parents=True, exist_ok=True)
    excel_dir.mkdir(parents=True, exist_ok=True)

    return pdf_dir, text_dir, excel_dir


def resolve_input_file(file_name: str, base_dir: Path) -> Path:
    """Resolves input PDF/text file from current dir or subdirectories."""
    pdf_dir, text_dir, _ = setup_directories(base_dir)
    target = Path(file_name.strip())

    if target.exists() and target.is_file():
        return target.resolve()

    in_pdf = pdf_dir / target.name
    if in_pdf.exists() and in_pdf.is_file():
        return in_pdf.resolve()

    in_text = text_dir / target.name
    if in_text.exists() and in_text.is_file():
        return in_text.resolve()

    if not target.suffix:
        in_pdf_ext = pdf_dir / f"{target.name}.pdf"
        if in_pdf_ext.exists():
            return in_pdf_ext.resolve()

        in_text_ext = text_dir / f"{target.name}.txt"
        if in_text_ext.exists():
            return in_text_ext.resolve()

    raise FileNotFoundError(
        f"Could not locate file '{file_name}'. Checked:\n"
        f"  • Direct path: {target.resolve()}\n"
        f"  • PDF directory: {pdf_dir / target.name}\n"
        f"  • Text directory: {text_dir / target.name}"
    )


def extract_text_from_pdf_with_delimiters(pdf_path: Path, delimiter: str = " | ") -> str:
    """
    Extracts text from PDF while using raw bounding box coordinates and table geometry
    to insert explicit column delimiters (|) between table cells in the text file.
    """
    text_blocks = []

    if PDFPLUMBER_AVAILABLE:
        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                total_pages = len(pdf.pages)
                for idx, page in enumerate(pdf.pages, start=1):
                    # 1. Native table extraction via bounding box grid
                    tables = page.extract_tables()
                    if tables:
                        page_lines = []
                        for table in tables:
                            for row in table:
                                clean_row = [str(cell or "").strip().replace("\n", " ") for cell in row]
                                clean_row = [c for c in clean_row if c != ""]
                                if clean_row:
                                    page_lines.append(f"{delimiter}".join(clean_row))
                        if page_lines:
                            text_blocks.append(f"--- Page {idx} of {total_pages} ---\n\n" + "\n".join(page_lines))
                            continue

                    # 2. Word bounding box coordinate alignment fallback
                    words = page.extract_words()
                    if words:
                        words_sorted = sorted(words, key=lambda w: (round(w['top'], 1), w['x0']))
                        lines = []
                        current_line = []
                        last_top = None

                        for w in words_sorted:
                            if last_top is None or abs(w['top'] - last_top) < 3:
                                current_line.append(w)
                                last_top = w['top']
                            else:
                                lines.append(current_line)
                                current_line = [w]
                                last_top = w['top']
                        if current_line:
                            lines.append(current_line)

                        line_strings = []
                        for line_words in lines:
                            line_words = sorted(line_words, key=lambda w: w['x0'])
                            tokens = []
                            current_token = []
                            last_x1 = None

                            for w in line_words:
                                if last_x1 is not None and (w['x0'] - last_x1) > 8:
                                    tokens.append(" ".join(current_token))
                                    current_token = [w['text']]
                                else:
                                    current_token.append(w['text'])
                                last_x1 = w['x1']

                            if current_token:
                                tokens.append(" ".join(current_token))

                            line_strings.append(f"{delimiter}".join(tokens))

                        text_blocks.append(f"--- Page {idx} of {total_pages} ---\n\n" + "\n".join(line_strings))
                    else:
                        page_text = page.extract_text() or ""
                        text_blocks.append(f"--- Page {idx} of {total_pages} ---\n\n" + page_text.strip())

            return "\n\n".join(text_blocks)
        except Exception:
            pass

    if PYPDF_AVAILABLE:
        reader = pypdf.PdfReader(str(pdf_path))
        for idx, page in enumerate(reader.pages, start=1):
            page_text = page.extract_text() or ""
            text_blocks.append(f"--- Page {idx} of {len(reader.pages)} ---\n\n" + page_text.strip())
        return "\n\n".join(text_blocks)

    raise RuntimeError("No PDF library (pdfplumber or pypdf) is installed.")


def convert_pdf_to_text(
    pdf_path: str | Path,
    output_path: str | Path | None = None,
    base_dir: Path | None = None,
    quiet: bool = False
) -> Path:
    """Converts a PDF file to a pipe-delimited text (.txt) file in text_files/."""
    script_dir = base_dir or Path(__file__).resolve().parent
    _, text_dir, _ = setup_directories(script_dir)

    input_file = Path(pdf_path).resolve()
    if not input_file.exists():
        raise FileNotFoundError(f"Input PDF file not found at '{input_file}'")

    if output_path is None:
        output_file = text_dir / f"{input_file.stem}.txt"
    else:
        output_file = Path(output_path).resolve()

    start_time = time.time()
    extracted_text = extract_text_from_pdf_with_delimiters(input_file, delimiter=" | ")

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(extracted_text)

    elapsed = time.time() - start_time
    if not quiet:
        rel_input = input_file.relative_to(script_dir) if input_file.is_relative_to(script_dir) else input_file
        rel_output = output_file.relative_to(script_dir) if output_file.is_relative_to(script_dir) else output_file
        print(f"Text Extraction Successful!")
        print(f"  • Input PDF:    {rel_input}")
        print(f"  • Output Text:  {rel_output}")

    return output_file


def find_available_titles_in_text(text_content: str) -> list[str]:
    """Finds potential table/schedule title headings in the text."""
    titles = []
    lines = [line.strip() for line in text_content.splitlines() if line.strip()]
    for line in lines:
        if any(kw in line.upper() for kw in ["SINGLE DUCT", "TERMINAL SCEHDULE", "TERMINAL SCHEDULE", "EQUIPMENT SCHEDULE", "SCHEDULE", "TABLE"]):
            if len(line.split()) <= 15:
                titles.append(line)
    return titles


def parse_header_line(header_line: str) -> list[str]:
    """Parses column headers from a header text line."""
    if " | " in header_line:
        cols = [c.strip() for c in header_line.split(" | ") if c.strip()]
        if cols:
            return cols
    elif re.search(r'\s{2,}', header_line):
        cols = [c.strip() for c in re.split(r'\s{2,}', header_line) if c.strip()]
        if cols:
            return cols
    else:
        cols = [c.strip() for c in header_line.split() if c.strip()]
        if cols:
            return cols

    return []


def extract_single_table_with_subset(
    text_content: str,
    target_title: str,
    target_columns_csv: str
) -> tuple[dict | None, list[str]]:
    """
    Locates target_title and pulls the line directly above it as the column header line.
    Parses user column CSV from left to right, tracking position sequentially through the header line,
    and flags any remaining unconsumed column titles in the header line as missing.
    Returns (table_dict, flags_list).
    """
    flags = []
    lines = [line.strip() for line in text_content.splitlines() if line.strip()]

    requested_cols = [c.strip() for c in target_columns_csv.split(",") if c.strip()]
    if not requested_cols:
        flags.append("[FLAG/ERROR] No valid column titles provided in CSV input.")
        return None, flags

    # 1. Search for table title in text
    title_line_idx = None
    target_title_clean = target_title.strip().lower()

    for idx, line in enumerate(lines):
        line_clean = line.strip().lower()
        if target_title_clean in line_clean or line_clean in target_title_clean:
            title_line_idx = idx
            break

    if title_line_idx is None:
        avail_titles = find_available_titles_in_text(text_content)
        title_flag = f"[FLAG/ERROR] Table title '{target_title}' was NOT found in the PDF text."
        flags.append(title_flag)
        if avail_titles:
            flags.append("  • Available table titles found in PDF:\n    - " + "\n    - ".join(avail_titles))
        return None, flags

    actual_title_text = lines[title_line_idx]

    # 2. Pull line directly above title as column header line
    if title_line_idx == 0:
        flags.append(f"[FLAG/ERROR] No column header line found above table title '{actual_title_text}'.")
        return None, flags

    header_text = lines[title_line_idx - 1]
    pdf_columns = parse_header_line(header_text)

    # 3. Left-to-right sequential search & whole-word validation
    invalid_cols = []
    selected_indices = []
    selected_headers = []
    current_pos = 0

    for user_col in requested_cols:
        user_clean = user_col.strip()
        user_upper = user_clean.upper()

        pattern = r'\b' + re.escape(user_clean) + r'\b'
        match = re.search(pattern, header_text[current_pos:], re.IGNORECASE)

        if match:
            start_in_slice = match.start()
            end_in_slice = match.end()
            match_start = current_pos + start_in_slice
            match_end = current_pos + end_in_slice

            current_pos = match_end

            matched_idx = None
            for idx, col in enumerate(pdf_columns):
                if user_upper in col.upper() or col.upper() in user_upper:
                    matched_idx = idx
                    break
            if matched_idx is None and pdf_columns:
                matched_idx = len(selected_indices)

            selected_indices.append(matched_idx if matched_idx is not None else len(selected_indices))
            selected_headers.append(user_clean)
        else:
            invalid_cols.append(user_clean)

    if invalid_cols:
        col_flag = f"[FLAG/ERROR] The following requested column titles were NOT found as whole-word matches in sequential order in the header line:\n  • " + "\n  • ".join(invalid_cols)
        flags.append(col_flag)
        return None, flags

    # 4. Check remaining unconsumed text in header line for missing columns
    remaining_text = header_text[current_pos:].strip()
    remaining_text_clean = re.sub(r'^[|\s,\-:_]+', '', remaining_text).strip()

    if remaining_text_clean:
        missing_cols = parse_header_line(remaining_text_clean)
        if missing_cols:
            missing_flag = f"[FLAG/ERROR] Not all column titles from the PDF table schedule were entered.\n  • Missing required column(s): " + ", ".join(missing_cols)
            flags.append(missing_flag)
            return None, flags

    # 5. Collect table data rows (scanning both below and above title line)
    data_rows = []

    for idx in range(title_line_idx + 1, len(lines)):
        line = lines[idx]
        if any(kw in line.upper() for kw in ["SINGLE DUCT", "TERMINAL SCHEDULE", "EQUIPMENT SCHEDULE", "SCHEDULE", "TABLE"]) and idx != title_line_idx:
            break
        if " | " in line or line.startswith("VAV-") or re.match(r'^[A-Z0-9\-]+_[\w\-]+', line):
            tokens = [p.strip() for p in line.split(" | ")] if " | " in line else line.split()
            data_rows.append(tokens)

    if not data_rows:
        start_idx = max(0, title_line_idx - 120)
        for idx in range(title_line_idx - 1, start_idx, -1):
            line = lines[idx]
            if line == header_text or idx == title_line_idx:
                continue
            if any(kw in line.upper() for kw in ["SINGLE DUCT", "TERMINAL SCHEDULE", "EQUIPMENT SCHEDULE", "SCHEDULE", "TABLE"]) and idx != title_line_idx:
                break
            if " | " in line or line.startswith("VAV-") or re.match(r'^[A-Z0-9\-]+_[\w\-]+', line):
                tokens = [p.strip() for p in line.split(" | ")] if " | " in line else line.split()
                data_rows.insert(0, tokens)

    if not data_rows:
        flags.append(f"[FLAG/ERROR] No data rows found under table title '{actual_title_text}'.")
        return None, flags

    # 6. Filter data rows to selected user columns
    filtered_rows = []
    for row in data_rows:
        filtered_row = []
        for s_idx in selected_indices:
            if s_idx < len(row):
                filtered_row.append(row[s_idx])
            else:
                filtered_row.append("-")
        filtered_rows.append(filtered_row)

    table_data = {
        "title": actual_title_text,
        "headers": selected_headers,
        "rows": filtered_rows
    }

    return table_data, flags


def sanitize_sheet_title(raw_title: str) -> str:
    """Formats raw schedule title into a clean Excel sheet tab name."""
    clean = re.sub(r'[\\/*?:\[\]]', '', raw_title).strip()
    return clean or "Extracted Table"


def export_table_to_excel(table_dict: dict, output_excel_path: Path) -> Path:
    """Writes a single extracted table to an Excel spreadsheet."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to export Excel files.")

    wb = openpyxl.Workbook()
    ws = wb.active

    sheet_title = sanitize_sheet_title(table_dict["title"])
    ws.title = sheet_title[:31]
    ws.views.sheetView[0].showGridLines = True

    PRIMARY_COLOR = "1F497D"  # Navy blue
    ALT_ROW_COLOR = "F2F5F9"  # Light tint
    BORDER_COLOR = "D9D9D9"

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=PRIMARY_COLOR, end_color=PRIMARY_COLOR, fill_type="solid")
    alt_fill = PatternFill(start_color=ALT_ROW_COLOR, end_color=ALT_ROW_COLOR, fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color=BORDER_COLOR),
        right=Side(style="thin", color=BORDER_COLOR),
        top=Side(style="thin", color=BORDER_COLOR),
        bottom=Side(style="thin", color=BORDER_COLOR)
    )

    headers = table_dict.get("headers", [])
    rows = table_dict.get("rows", [])

    # Write header row
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[1].height = 28

    # Write data rows
    for row_idx, row_data in enumerate(rows, start=2):
        for col_idx, val in enumerate(row_data, start=1):
            clean_val = str(val).replace(",", "").strip() if val is not None else ""

            if clean_val.isdigit():
                cell_val = int(clean_val)
            else:
                try:
                    cell_val = float(clean_val)
                except ValueError:
                    cell_val = val

            cell = ws.cell(row=row_idx, column=col_idx, value=cell_val)
            cell.border = thin_border

            if row_idx % 2 == 1:
                cell.fill = alt_fill

            if isinstance(cell_val, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if isinstance(cell_val, float):
                    cell.number_format = "0.0"
                else:
                    cell.number_format = "#,##0"
            elif str(val).strip() in ("X", "-"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

        ws.row_dimensions[row_idx].height = 20

    ws.freeze_panes = "A2"

    for col in ws.columns:
        max_len = 0
        for cell in col:
            val_str = str(cell.value) if cell.value is not None else ""
            max_len = max(max_len, len(val_str))
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 35)

    output_excel_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_excel_path))
    return output_excel_path


def main():
    base_dir = Path(__file__).resolve().parent
    pdf_dir, text_dir, excel_dir = setup_directories(base_dir)

    print("\n=======================================================")
    print("📊 Interactive PDF & Table Extractor")
    print("=======================================================\n")

    # Step 1: Prompt for PDF / Input File
    file_input = input("1. Enter the path or name of the PDF file to extract: ").strip()

    try:
        input_file = resolve_input_file(file_input, base_dir)
    except FileNotFoundError as e:
        print(f"\n❌ [FLAG/ERROR] {e}", file=sys.stderr)
        sys.exit(1)

    # Read/extract text with explicit column delimiters (|)
    if input_file.suffix.lower() == ".pdf":
        convert_pdf_to_text(input_file, base_dir=base_dir)
        text_out = text_dir / f"{input_file.stem}.txt"
        text_content = text_out.read_text(encoding="utf-8")
    else:
        text_content = input_file.read_text(encoding="utf-8", errors="ignore")

    # Step 2: Prompt for Table Title
    table_title = input("\n2. Enter the EXACT Title of the table you want to extract: ").strip()

    # Step 3: Prompt for Column Titles
    target_columns = input("\n3. Enter the EXACT column titles to extract (CSV format) for column delimiting: ").strip()

    print("\nProcessing extraction...")

    # Step 4: Extract Table and Subset
    table_dict, flags = extract_single_table_with_subset(text_content, table_title, target_columns)

    if flags:
        print("\n--- Flags & Status Output ---")
        for flag in flags:
            print(flag)

    if table_dict is None:
        print("\n❌ Table extraction failed. Please check the flags above and try again.")
        sys.exit(1)

    # Step 5: Export to Excel
    output_excel = excel_dir / f"{input_file.stem}.xlsx"
    saved_path = export_table_to_excel(table_dict, output_excel)

    rel_excel = saved_path.relative_to(base_dir) if saved_path.is_relative_to(base_dir) else saved_path
    print(f"\n✅ Table Extraction Successful!")
    print(f"  • Output Excel: {rel_excel}")
    print(f"  • Tab Name:     '{sanitize_sheet_title(table_dict['title'])}'")
    print(f"  • Rows Exported: {len(table_dict['rows'])}\n")


if __name__ == "__main__":
    main()