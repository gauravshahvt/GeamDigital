import os
import io
import re
import base64
from typing import List, Dict, Any, Optional, Tuple
import pymupdf
import pdfplumber
from PIL import Image

from .ocr_engine import ocr_page_to_table, is_tesseract_available
from .invoice_parser import extract_invoice_metadata

def render_page_to_base64(page: pymupdf.Page, dpi: int = 150) -> str:
    """Renders a PDF page to a PNG image and returns base64 data URI."""
    zoom = dpi / 72.0
    mat = pymupdf.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img_bytes = pix.tobytes("png")
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    return f"data:image/png;base64,{b64}"

def extract_spatial_aligned_tables(page: pymupdf.Page) -> List[List[List[str]]]:
    """
    High-precision spatial layout table detector:
    - Clusters words into horizontal lines
    - Groups words within tokens
    - Detects column header boundaries
    - Maps tokens into exact column bounds based on x-coordinates
    - Supports open tables, whitespace tables, and horizontal-line tables.
    """
    words = page.get_text("words")
    if not words:
        return []

    # 1. Cluster words into lines based on vertical overlap
    lines_dict = {}
    for w in words:
        yc = (w[1] + w[3]) / 2.0
        matched = False
        for y_key in list(lines_dict.keys()):
            if abs(yc - y_key) < 5.0:
                lines_dict[y_key].append(w)
                matched = True
                break
        if not matched:
            lines_dict[yc] = [w]

    sorted_y = sorted(lines_dict.keys())

    # 2. Extract candidate multi-column lines
    candidate_lines = []
    for y in sorted_y:
        ws = sorted(lines_dict[y], key=lambda x: x[0])
        tokens = []
        cur_tok = [ws[0]]
        for w in ws[1:]:
            # If gap between words is small, merge into one token
            if w[0] - cur_tok[-1][2] < 8.0:
                cur_tok.append(w)
            else:
                tokens.append((cur_tok[0][0], cur_tok[-1][2], ' '.join(t[4] for t in cur_tok)))
                cur_tok = [w]
        tokens.append((cur_tok[0][0], cur_tok[-1][2], ' '.join(t[4] for t in cur_tok)))

        # Must have at least 2 distinct tokens across the line
        if len(tokens) >= 2:
            candidate_lines.append((y, tokens))

    if not candidate_lines:
        return []

    # 3. Find table blocks (groups of candidate lines that share column structure)
    # Search for a header line that has >= 3 columns or clear column titles
    tables = []
    i = 0
    while i < len(candidate_lines):
        y, tokens = candidate_lines[i]
        
        # Check if line looks like a table header (3+ columns, or 2+ structured columns)
        if len(tokens) >= 3 or (len(tokens) >= 2 and any(re.search(r'(?i)(date|id|code|item|desc|qty|price|amount|total|debit|credit|bal)', t[2]) for t in tokens)):
            # Potential table start
            header_tokens = tokens
            col_bounds = []
            for c_i, tok in enumerate(header_tokens):
                c_start = tok[0] - 6.0
                c_end = header_tokens[c_i + 1][0] - 6.0 if c_i + 1 < len(header_tokens) else 9999.0
                col_bounds.append((c_start, c_end, tok[2]))

            grid = [[t[2] for t in header_tokens]]
            
            # Collect following rows that fit into this table
            j = i + 1
            consecutive_misses = 0
            while j < len(candidate_lines):
                next_y, next_toks = candidate_lines[j]
                
                # If vertical gap is too large (> 40 points), table ended
                prev_y = candidate_lines[j - 1][0]
                if next_y - prev_y > 45.0:
                    break

                # Map row tokens into column boundaries
                row = [''] * len(col_bounds)
                valid_placement = False
                for tok in next_toks:
                    tx0, tx1, text = tok
                    tc = (tx0 + tx1) / 2.0
                    
                    best_col = None
                    for c_idx, (c_start, c_end, _) in enumerate(col_bounds):
                        if c_start <= tc < c_end:
                            best_col = c_idx
                            break
                    if best_col is None:
                        # Find closest column
                        best_col = min(range(len(col_bounds)), key=lambda idx: abs(tc - (col_bounds[idx][0] + min(col_bounds[idx][1], 600))/2.0))

                    if 0 <= best_col < len(row):
                        if row[best_col]:
                            row[best_col] += ' ' + text
                        else:
                            row[best_col] = text
                        valid_placement = True

                # If this line matches at least 1 column and isn't a footer text
                if valid_placement and any(row):
                    grid.append(row)
                    consecutive_misses = 0
                else:
                    consecutive_misses += 1
                    if consecutive_misses > 2:
                        break
                j += 1

            # Only accept as a table if it has at least 2 rows (header + at least 1 data row)
            if len(grid) >= 2:
                # Clean up empty columns
                col_has_data = [any(r[c].strip() for r in grid) for c in range(len(grid[0]))]
                cleaned_grid = []
                for r in grid:
                    cleaned_grid.append([r[c] for c in range(len(r)) if col_has_data[c]])
                if cleaned_grid and len(cleaned_grid[0]) >= 2:
                    tables.append(cleaned_grid)
                i = j  # Move pointer past this table
                continue

        i += 1

    return tables

def extract_tables_pymupdf_vector(page: pymupdf.Page) -> List[List[List[str]]]:
    """Uses PyMuPDF's built-in find_tables for vector-bordered tables."""
    tables = []
    try:
        tabs = page.find_tables()
        for tab in tabs:
            extracted = tab.extract()
            cleaned = []
            for row in extracted:
                clean_row = [str(c).strip() if c is not None else "" for c in row]
                if any(clean_row):
                    cleaned.append(clean_row)
            if cleaned and len(cleaned) > 1:
                tables.append(cleaned)
    except Exception:
        pass
    return tables

def extract_tables_pdfplumber(pdf_path: str, page_number: int, mode: str = "auto") -> List[List[List[str]]]:
    """Uses pdfplumber's lattice/stream table extractor."""
    tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_number <= len(pdf.pages):
                plumb_page = pdf.pages[page_number - 1]
                table_settings = {}
                if mode == "lattice":
                    table_settings = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
                elif mode == "stream":
                    table_settings = {"vertical_strategy": "text", "horizontal_strategy": "text"}
                
                extracted_raw = plumb_page.extract_tables(table_settings) if table_settings else plumb_page.extract_tables()
                for raw_tbl in extracted_raw:
                    cleaned = []
                    for row in raw_tbl:
                        clean_row = [str(c).strip() if c is not None else "" for c in row]
                        if any(clean_row):
                            cleaned.append(clean_row)
                    if cleaned and len(cleaned) > 1:
                        tables.append(cleaned)
    except Exception:
        pass
    return tables

def process_pdf(
    pdf_path: str,
    selected_pages: Optional[List[int]] = None,
    mode: str = "auto",  # 'auto', 'spatial', 'lattice', 'stream', 'ocr'
    include_preview: bool = True
) -> Dict[str, Any]:
    """
    Master extraction pipeline:
    - Analyzes document layout
    - Combines spatial column detection, vector grid detection, pdfplumber & OCR
    - Extracts invoice/receipt key-values
    - Renders visual page previews
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"File not found: {pdf_path}")

    doc = pymupdf.open(pdf_path)
    total_pages = len(doc)
    pages_to_process = selected_pages or list(range(1, total_pages + 1))

    all_tables = []
    previews = {}
    full_text = ""

    for p_num in pages_to_process:
        if p_num < 1 or p_num > total_pages:
            continue
        page = doc[p_num - 1]
        raw_text = page.get_text()
        full_text += f"\n--- Page {p_num} ---\n" + raw_text

        # 1. Preview
        if include_preview and str(p_num) not in previews:
            previews[str(p_num)] = render_page_to_base64(page)

        page_tables = []
        is_scanned = len(raw_text.strip()) < 30

        # OCR mode (scanned / user selected)
        if mode == "ocr" or (mode == "auto" and is_scanned):
            pix = page.get_pixmap(dpi=300)
            pil_img = Image.open(io.BytesIO(pix.tobytes("png")))
            df_ocr = ocr_page_to_table(pil_img)
            if not df_ocr.empty:
                headers = list(df_ocr.columns)
                rows = df_ocr.values.tolist()
                page_tables.append({
                    'page': p_num,
                    'method': 'OCR (Scanned Document)',
                    'data': [headers] + rows
                })
        else:
            # First: Try Spatial Alignment Engine (super reliable for bank statements, invoices, tables)
            spatial_tables = extract_spatial_aligned_tables(page)
            for st in spatial_tables:
                page_tables.append({
                    'page': p_num,
                    'method': 'Smart Layout Engine',
                    'data': st
                })

            # Second: PyMuPDF Vector find_tables
            if not page_tables or mode == "lattice":
                vector_tables = extract_tables_pymupdf_vector(page)
                for vt in vector_tables:
                    if not any(vt == existing['data'] for existing in page_tables):
                        page_tables.append({
                            'page': p_num,
                            'method': 'Vector Border Engine',
                            'data': vt
                        })

            # Third: pdfplumber if still empty or explicitly requested
            if not page_tables or mode in ["lattice", "stream"]:
                plumb_mode = "stream" if mode == "stream" else ("lattice" if mode == "lattice" else "auto")
                plumb_tables = extract_tables_pdfplumber(pdf_path, p_num, mode=plumb_mode)
                for pt in plumb_tables:
                    if not any(pt == existing['data'] for existing in page_tables):
                        page_tables.append({
                            'page': p_num,
                            'method': f'PDFPlumber ({plumb_mode})',
                            'data': pt
                        })

        for pt in page_tables:
            all_tables.append(pt)

    # Extract Invoice Key-Value Metadata
    invoice_metadata = extract_invoice_metadata(full_text)

    doc.close()

    return {
        'total_pages': total_pages,
        'processed_pages': pages_to_process,
        'tables_count': len(all_tables),
        'tables': all_tables,
        'invoice_metadata': invoice_metadata,
        'previews': previews,
        'is_ocr_available': is_tesseract_available()
    }
