import sys
import os
import re
import pymupdf
import pytesseract
from PIL import Image
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

def extract_voters_from_pdf(pdf_path: str):
    doc = pymupdf.open(pdf_path)
    total_pages = len(doc)
    print(f"Opened PDF: {pdf_path}, Total pages: {total_pages}")

    # 1. Metadata Extraction from Page 1
    p1 = doc[0]
    p1_pix = p1.get_pixmap(dpi=150)
    p1_img = Image.frombytes("RGB", [p1_pix.width, p1_pix.height], p1_pix.samples)
    p1_ocr = pytesseract.image_to_string(p1_img, lang="hin+eng")

    # Extract Ward / Part Number
    part_no = "1"
    ward_m = re.search(r'(?:वार्ड\s*क्रमांक|भाग\s*संख्या)\s*[:\-\s]*([0-9]+)', p1_ocr)
    if ward_m:
        part_no = ward_m.group(1).strip()

    # Extract Booth Address
    booth_address = ""
    booth_m = re.search(r'मतदान\s*बूथ\s*की\s*संख्या\s*एवं\s*पता\s*[:\-\s]*([^\n]+)', p1_ocr)
    if booth_m:
        booth_address = booth_m.group(1).strip()
    else:
        # Fallback search
        for line in p1_ocr.split('\n'):
            if 'विद्यालय' in line or 'कमरा' in line or 'राजकीय' in line:
                booth_address = line.strip()
                break

    # Extract Village / Mohalla Address from Page 3
    address = "बागमाली"
    if total_pages >= 3:
        p3_text = doc[2].get_text()
        addr_m = re.search(r'([^\n]+(?:बागमाली|वार्ड|ग्राम)[^\n]*)', p3_text)
        # Search page 3 header
        for b in doc[2].get_text('blocks'):
            txt = b[4].strip()
            if 'बागमाली' in txt and len(txt) < 40 and not 'ग्रामपंचायत' in txt:
                address = txt.replace('\n', ', ')
                break

    print(f"Extracted Metadata:\n  भाग संख्या: {part_no}\n  बूथ का पता: {booth_address}\n  एड्रेस: {address}")

    # 2. Extract Deletion List from Page 14 (if present)
    deleted_serials = set()
    deleted_epics = set()

    for pno in range(total_pages):
        p_txt = doc[pno].get_text()
        if "विलोपन सूची" in p_txt or "घटक 2" in p_txt:
            print(f"Found Deletion List on Page {pno + 1}")
            blocks = doc[pno].get_text("blocks")
            for b in blocks:
                txt = b[4].strip()
                epic_m = re.search(r'([A-Z]{3}\d{7}|[A-Z0-9/]{8,18})', txt)
                if epic_m and ('S' in txt or 'E' in txt or 'R' in txt):
                    deleted_epics.add(epic_m.group(1))

    print(f"Deleted EPICs identified from Deletion List: {len(deleted_epics)}")

    # 3. Process all voter pages (typically page 3 onwards, skipping Nazri Naksha / Deletion / Summary)
    voters = []
    zoom = 300 / 72.0
    mat = pymupdf.Matrix(zoom, zoom)

    for pno in range(2, total_pages):
        p = doc[pno]
        p_txt = p.get_text()

        # Skip non-voter pages (Map, Deletion List, Summary)
        if "वार्ड का नक्शा" in p_txt or "घटक 2: विलोपन" in p_txt or "(क) निर्वाचकों की संख्या" in p_txt or "विशेष संक्षिप्त पुनरीक्षण" in p_txt:
            print(f"Skipping Page {pno + 1} (Not an active voter list page)")
            continue

        # Check for watermark images (DELETED stamps)
        deleted_rects = []
        for img in p.get_images():
            if img[2] == 106 and img[3] == 76:  # standard DELETED stamp dimensions
                deleted_rects.extend(p.get_image_rects(img[0]))

        # Find voter card bounding boxes on this page
        # Group drawings into horizontal and vertical lines
        drawings = p.get_drawings()
        lines = []
        for d in drawings:
            for it in d.get('items', []):
                if it[0] == 'l':
                    lines.append((it[1], it[2]))

        h_lines = sorted(list(set(round(it[0].y, 1) for it in lines if abs(it[0].y - it[1].y) < 1 and 130 < it[0].y < 750)))
        
        # Build row intervals
        row_intervals = []
        if len(h_lines) >= 2:
            for i in range(len(h_lines) - 1):
                y0 = h_lines[i]
                y1 = h_lines[i+1]
                if 60 < (y1 - y0) < 85:
                    row_intervals.append((y0, y1))
        
        # If no h-lines (like page 13 which uses rects), extract from rects
        if not row_intervals:
            rects = [d['rect'] for d in drawings if d.get('rect')]
            card_rects = [r for r in rects if 150 < (r.x1 - r.x0) < 200 and 60 < (r.y1 - r.y0) < 85]
            y_starts = sorted(list(set(round(r.y0, 1) for r in card_rects)))
            for y0 in y_starts:
                matching = [r for r in card_rects if abs(r.y0 - y0) < 2]
                if matching:
                    row_intervals.append((y0, matching[0].y1))

        if not row_intervals:
            print(f"Page {pno + 1}: No voter card grid detected, skipping.")
            continue

        x_cols = [(34.4, 207.9), (208.4, 381.4), (381.9, 555.4)]

        print(f"Processing Page {pno + 1}: {len(row_intervals)} rows x 3 cols = up to {len(row_intervals)*3} cards...")

        for y0, y1 in row_intervals:
            for x0, x1 in x_cols:
                card_rect = pymupdf.Rect(x0, y0, x1, y1)
                card_words = p.get_text("words", clip=card_rect)
                if not card_words:
                    continue

                card_text = " ".join(w[4] for w in sorted(card_words, key=lambda w: (round(w[1]/10)*10, w[0])))
                
                # Check for DELETED
                is_deleted = False
                if any(card_rect.intersects(dr) for dr in deleted_rects):
                    is_deleted = True

                first_words = [w[4] for w in card_words[:4]]
                if 'S' in first_words or 'E' in first_words or 'R' in first_words:
                    is_deleted = True

                # Extract Serial Number
                serial = None
                for w in card_words:
                    if w[4].isdigit() and int(w[4]) < 1000 and (w[1] - y0) < 28 and (w[0] - x0) < 50:
                        serial = int(w[4])
                        break

                # Extract Voter ID (EPIC)
                epic = ""
                for w in card_words:
                    t = w[4]
                    if re.match(r'^[A-Z]{3}\d{7}$', t) or re.match(r'^[A-Z0-9/]{8,18}$', t):
                        epic = t
                        break

                if epic and epic in deleted_epics:
                    is_deleted = True

                # Extract Age
                age = ""
                age_m = re.search(r'(?:आजच|आयु)\s*:\s*(\d{1,3})', card_text)
                if age_m:
                    age = int(age_m.group(1))

                # Extract House Number
                house = ""
                house_m = re.search(r'(?:मकरन|मकान)\s*(?:सपखजर|संख्या)\s*:\s*([^\s]+)', card_text)
                if house_m:
                    house = house_m.group(1).strip()
                    # Clean trailing artifacts
                    house = re.sub(r'[^0-9a-zA-Z\u0900-\u097F]', '', house)

                # Skip if deleted or if no serial found
                if is_deleted:
                    print(f"  [DELETED SKIPPED] Serial: {serial}, EPIC: {epic}")
                    continue

                if serial is None:
                    continue

                # Run OCR on the name block of the card
                # Left 65% of the card, excluding serial and photo
                name_crop_rect = pymupdf.Rect(x0 + 1, y0 + 13, x0 + 106, y0 + 52)
                pix = p.get_pixmap(matrix=mat, clip=name_crop_rect)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                
                ocr_txt = pytesseract.image_to_string(img, lang="hin", config="--psm 6").strip()
                ocr_lines = [l.strip() for l in ocr_txt.split('\n') if l.strip()]

                voter_name = ""
                relative_name = ""

                for line in ocr_lines:
                    if "नाम" in line and not ("पिता" in line or "पति" in line):
                        raw_name = re.sub(r'^.*?नाम\s*[:\-\s,]*', '', line).strip()
                        voter_name = clean_hindi_name(raw_name)
                    elif "पिता" in line or "पति" in line:
                        raw_rel = re.sub(r'^.*?(?:पिता|पति)\s*का\s*नाम\s*[:\-\s,]*', '', line).strip()
                        relative_name = clean_hindi_name(raw_rel)

                voter_record = {
                    'भाग संख्या': part_no,
                    'क्रम संख्या': serial,
                    'नाम': voter_name,
                    'पिता/पति का नाम': relative_name,
                    'आयु': age,
                    'मोबाइल नो': '',
                    'वोटर ID': epic,
                    'हाउस नंबर': house,
                    'एड्रेस': address,
                    'बूथ का पता': booth_address
                }
                voters.append(voter_record)

    doc.close()
    return voters

def clean_hindi_name(text: str) -> str:
    """Cleans OCR noise from Hindi names."""
    if not text:
        return ""
    # Remove English characters, trailing symbols, bars, quotes
    cleaned = re.sub(r'[a-zA-Z0-9|!@#$%^&*()_+=\[\]{};\'"\\<>\/?~`]', '', text).strip()
    # Remove leading commas or colons
    cleaned = re.sub(r'^[,\-.:\s]+', '', cleaned).strip()
    cleaned = re.sub(r'[,\-.:\s]+$', '', cleaned).strip()
    # Fix common OCR artifact where 'म' at beginning is seen as ','
    if cleaned.startswith('ो'):
        cleaned = 'मो' + cleaned[1:]
    return cleaned

if __name__ == '__main__':
    voters = extract_voters_from_pdf('uploads/b9707b19-9ff5-425e-bd9d-f0973deb1ea6_BAGMALI-Ward_No-001.pdf')
    print(f"\nSuccessfully extracted {len(voters)} active voters!")
    df = pd.DataFrame(voters)
    df.to_excel('exports/voter_list_bagmali_ward_001.xlsx', index=False)
    print(f"Saved to: exports/voter_list_bagmali_ward_001.xlsx")
    print("\nFirst 10 voters:")
    print(df.head(10)[['क्रम संख्या', 'नाम', 'पिता/पति का नाम', 'आयु', 'वोटर ID', 'हाउस नंबर']])
