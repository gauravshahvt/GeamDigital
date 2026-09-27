import os
import shutil
import pytesseract
from PIL import Image
import pandas as pd
import numpy as np

# Configure tesseract binary path if needed
POSSIBLE_TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    shutil.which("tesseract")
]

for p in POSSIBLE_TESSERACT_PATHS:
    if p and os.path.exists(p):
        pytesseract.pytesseract.tesseract_cmd = p
        break

def is_tesseract_available() -> bool:
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False

def ocr_page_to_table(pil_image: Image.Image) -> pd.DataFrame:
    """
    Perform OCR on a page image and reconstruct a tabular layout
    by clustering words into rows and columns based on bounding boxes.
    """
    if not is_tesseract_available():
        return pd.DataFrame([["Tesseract OCR is not installed or not found on PATH"]])

    # Run Tesseract with TSV/Dict layout output
    data = pytesseract.image_to_data(pil_image, output_type=pytesseract.Output.DICT)
    
    n_boxes = len(data['text'])
    items = []
    
    for i in range(n_boxes):
        text = data['text'][i].strip()
        conf = int(data['conf'][i]) if str(data['conf'][i]).replace('-', '').isdigit() else 0
        if text and conf > 20:  # Filter empty and ultra-low confidence noise
            x = data['left'][i]
            y = data['top'][i]
            w = data['width'][i]
            h = data['height'][i]
            items.append({
                'text': text,
                'x0': x,
                'x1': x + w,
                'y0': y,
                'y1': y + h,
                'yc': y + h / 2.0,
                'xc': x + w / 2.0,
                'h': h
            })

    if not items:
        return pd.DataFrame()

    # Sort primarily by vertical coordinate (top to bottom)
    items.sort(key=lambda item: item['y0'])

    # Group items into lines (rows) based on vertical overlap / proximity
    rows = []
    current_row = []
    current_y = None
    row_height_threshold = 12  # Adaptive threshold

    for item in items:
        if current_y is None:
            current_row.append(item)
            current_y = item['yc']
            row_height_threshold = max(10, item['h'] * 0.7)
        else:
            if abs(item['yc'] - current_y) <= row_height_threshold:
                current_row.append(item)
                # Update moving average
                current_y = sum(it['yc'] for it in current_row) / len(current_row)
            else:
                # Finish previous row
                current_row.sort(key=lambda it: it['x0'])
                rows.append(current_row)
                current_row = [item]
                current_y = item['yc']
                row_height_threshold = max(10, item['h'] * 0.7)

    if current_row:
        current_row.sort(key=lambda it: it['x0'])
        rows.append(current_row)

    if not rows:
        return pd.DataFrame()

    # Collect column x-coordinates across all rows to determine column boundaries
    all_x0s = [item['x0'] for row in rows for item in row]
    if not all_x0s:
        return pd.DataFrame()

    # Cluster column positions using a histogram / gap threshold
    all_x0s.sort()
    col_clusters = []
    cluster_gap = 45  # Pixel gap to define a distinct column

    for x in all_x0s:
        matched = False
        for c in col_clusters:
            if abs(c['mean'] - x) <= cluster_gap:
                c['points'].append(x)
                c['mean'] = sum(c['points']) / len(c['points'])
                matched = True
                break
        if not matched:
            col_clusters.append({'mean': float(x), 'points': [x]})

    # Keep clusters that appear multiple times or span the page
    col_clusters.sort(key=lambda c: c['mean'])
    col_centers = [c['mean'] for c in col_clusters if len(c['points']) >= max(1, len(rows) * 0.15)]
    if not col_centers:
        col_centers = [c['mean'] for c in col_clusters]

    # Map words into the nearest column center
    grid = []
    for row in rows:
        row_cells = [""] * len(col_centers)
        for item in row:
            # find closest column index
            closest_idx = min(range(len(col_centers)), key=lambda idx: abs(item['x0'] - col_centers[idx]))
            if row_cells[closest_idx]:
                row_cells[closest_idx] += " " + item['text']
            else:
                row_cells[closest_idx] = item['text']
        # If row has any non-empty cell
        if any(cell.strip() for cell in row_cells):
            grid.append(row_cells)

    if not grid:
        return pd.DataFrame()

    df = pd.DataFrame(grid)
    # Promote first row as header if it contains string headers
    if len(df) > 1:
        first_row = df.iloc[0].tolist()
        if all(isinstance(c, str) and c.strip() for c in first_row):
            df.columns = [c.strip() if c.strip() else f"Col_{i+1}" for i, c in enumerate(first_row)]
            df = df.iloc[1:].reset_index(drop=True)

    return df
