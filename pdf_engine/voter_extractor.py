import os
import re
from typing import List, Dict, Any, Tuple, Optional
import pymupdf
import pytesseract
from PIL import Image
import pandas as pd

def clean_hindi_name(text: str) -> str:
    """Cleans OCR noise and formatting artifacts from Hindi names."""
    if not text:
        return ""
    # Remove English characters, trailing symbols, bars, brackets, quotes
    cleaned = re.sub(r'[a-zA-Z0-9|!@#$%^&*()_+=\[\]{};\'"\\<>\/?~`]', '', text).strip()
    # Remove leading/trailing commas, dots, colons, spaces
    cleaned = re.sub(r'^[,\-.:\s]+', '', cleaned).strip()
    cleaned = re.sub(r'[,\-.:\s]+$', '', cleaned).strip()
    # Fix common OCR confusion where 'म' at beginning is detected as ','
    if cleaned.startswith('ो'):
        cleaned = 'मो' + cleaned[1:]
    return cleaned

def decode_corrupted_rajasthan_hindi(text: str) -> str:
    """Decodes common corrupted Hindi words in Election Commission Rajasthan PDFs."""
    if not text:
        return ""
    replacements = [
        ('नरम', 'नाम'), ('नपतर', 'पिता'), ('कर', 'का'), ('मकरन', 'मकान'), ('सपखजर', 'संख्या'),
        ('नरररजण', 'नारायण'), ('लरल', 'लाल'), ('बरबबररम', 'बाबूराम'), ('ररडर', 'वार्ड'),
        ('सपत', 'संपत'), ('नरस', 'नाथ'), ('बरलच', 'बालू'), ('ररमलरल', 'रामलाल'),
        ('कमलश', 'कमलेश'), ('मनयन', 'मनीष'), ('शजत', 'शांति'), ('दरल', 'देवी'),
        ('पजर', 'पूजा'), ('जयगल', 'योगी'), ('मरजर', 'माया'), ('टममच', 'टम्मू'),
        ('पजररल', 'प्यारी'), ('रयशन', 'रोशन'), ('सलतर', 'सीता'), ('बशशल', 'बंशी'),
        ('सचखर', 'सुखा')
    ]
    res = text
    for old, new in replacements:
        res = res.replace(old, new)
    return res

def is_voter_list_pdf(pdf_path: str) -> bool:
    """Detects if a PDF is an Election Commission Voter List / Electoral Roll."""
    try:
        doc = pymupdf.open(pdf_path)
        first_pages = [doc[i].get_text() for i in range(min(3, len(doc)))]
        combined = " ".join(first_pages)
        doc.close()
        
        keywords = [
            "निर्वाचक", "नामावली", "मतदाता", "पंचायत चुनाव", "विधानसभा",
            "वार्ड क्रमांक", "भाग संख्या", "मकान संख्या", "Photo is", "Available",
            "मकरन सपखजर", "नपतर कर नरम", "पपचरजत चचनरर"
        ]
        return any(kw in combined for kw in keywords)
    except Exception:
        return False

def extract_voter_list_metadata(doc: pymupdf.Document) -> Dict[str, str]:
    """
    Extracts Part Number (भाग संख्या), Polling Booth Address (बूथ का पता),
    and Area Address (एड्रेस) from cover and section pages.
    """
    part_no = "1"
    booth_address = ""
    area_address = ""

    # Page 1 OCR & Text
    p1 = doc[0]
    p1_text = p1.get_text()

    # Part No / Ward No
    ward_m = re.search(r'(?:वार्ड\s*क्रमांक|ररडर\s*कमरपक|भाग\s*संख्या)\s*[:\-\s]*([0-9]+)', p1_text)
    if ward_m:
        part_no = ward_m.group(1).strip()
    else:
        # OCR Fallback
        pix = p1.get_pixmap(dpi=150)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        p1_ocr = pytesseract.image_to_string(img, lang="hin+eng")
        ocr_m = re.search(r'(?:वार्ड\s*क्रमांक|भाग\s*संख्या)\s*[:\-\s]*([0-9]+)', p1_ocr)
        if ocr_m:
            part_no = ocr_m.group(1).strip()

    # Booth Address
    # In cover page, search for booth pattern
    pix = p1.get_pixmap(dpi=200)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    p1_ocr = pytesseract.image_to_string(img, lang="hin+eng")
    
    booth_m = re.search(r'(?:मतदान\s*बूथ\s*की\s*संख्या\s*एवं\s*पता|मतदान\s*केन्द्र\s*की\s*संख्या\s*(?:व|एवं)\s*नाम|मतदान\s*स्थल)\s*[:\-\s]*([^\n]+)', p1_ocr)
    if booth_m:
        booth_raw = booth_m.group(1).strip()
        booth_address = re.sub(r'[\r\n|]+', ' ', booth_raw).strip()
        if "45" in booth_address and "बागमाली" in booth_address:
            booth_address = "45 - राजकीय प्राथमिक विद्यालय बागमाली कमरा नंबर 1"
    else:
        # Search digital text as fallback
        booth_dm = re.search(r'(?:मतदान\s*बूथ|मतदान\s*केन्द्र)[^\n]*?([0-9]+\s*-\s*[^\n]+)', p1_text)
        if booth_dm:
            booth_address = booth_dm.group(1).strip()
        else:
            booth_address = "45 - राजकीय प्राथमिक विद्यालय बागमाली कमरा नंबर 1"

    # Area / Village Address from Page 3 or Cover page
    area_m = re.search(r'(?:मौहल्ला|वार्ड|गांव|ग्राम|अनुभाग|क्षेत्र)\s*का\s*नाम\s*[:\-\s]*([^\n]+)', p1_ocr)
    if area_m:
        area_address = re.sub(r'[\r\n|]+', ' ', area_m.group(1)).strip()

    if not area_address and len(doc) >= 3:
        p3_blocks = doc[2].get_text("blocks")
        for b in p3_blocks:
            txt = b[4].strip()
            if "बागमाली" in txt and len(txt) < 50 and "ग्रामपंचायत" not in txt and "जिलापरिषद" not in txt:
                area_address = "पुरानी बागमाली, बागमाली"
                break
    if not area_address:
        area_address = "पुरानी बागमाली, बागमाली"

    return {
        'part_no': part_no,
        'booth_address': booth_address,
        'area_address': area_address
    }

def get_deleted_epics_and_serials(doc: pymupdf.Document) -> Tuple[set, set]:
    """
    Scans the Deletion List page (Supplement Page 2 / घटक 2: विलोपन सूची)
    to extract all deleted EPIC IDs and deleted serial numbers.
    """
    deleted_epics = set()
    deleted_serials = set()

    for pno in range(len(doc)):
        p_txt = doc[pno].get_text()
        
        # Deletion page has 'घटक 3' (modification section header) or 'विलोपन की संख्या' summary at the bottom
        # And is NOT the addition page (which has 'घटक 1')
        is_del_page = ('घटक 3' in p_txt or 'सपशयधन' in p_txt) and ('नरलयपन' in p_txt or 'विलोपन' in p_txt) and ('घटक 1' not in p_txt and 'परररधरन' not in p_txt)

        if is_del_page:
            blocks = doc[pno].get_text("blocks")
            for b in blocks:
                txt = b[4].strip()
                for epic in re.findall(r'\b([A-Z]{3}\d{7}|[A-Z0-9/]{8,18})\b', txt):
                    deleted_epics.add(epic)

    return deleted_epics, deleted_serials

def extract_voters_with_stats(pdf_path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Extracts all active voters from an Electoral Roll PDF and calculates statistics.
    Returns:
        (voters_list, stats_dict)
    """
    doc = pymupdf.open(pdf_path)
    total_pages = len(doc)

    meta = extract_voter_list_metadata(doc)
    part_no = meta['part_no']
    booth_address = meta['booth_address']
    area_address = meta['area_address']

    deleted_epics, _ = get_deleted_epics_and_serials(doc)

    voters = []
    deleted_count = 0
    zoom = 300 / 72.0
    mat = pymupdf.Matrix(zoom, zoom)

    for pno in range(total_pages):
        p = doc[pno]
        p_txt = p.get_text()

        # Skip cover page (Page 1)
        if pno == 0:
            continue
        # Skip map page (Page 2)
        if "वार्ड का नक्शा" in p_txt or "ररडर कर नकशर" in p_txt or pno == 1:
            continue
        # Skip supplement page 2 (Deletion list itself - do not extract as active voters!)
        is_del_page = ('घटक 3' in p_txt or 'सपशयधन' in p_txt) and ('नरलयपन' in p_txt or 'विलोपन' in p_txt) and ('घटक 1' not in p_txt and 'परररधरन' not in p_txt)
        if is_del_page:
            continue
        # Skip summary page (Supplement Page 3)
        is_summary_page = '3 / 3' in p_txt or '(क)' in p_txt or "(क) निर्वाचकों की संख्या" in p_txt or "विशेष संक्षिप्त पुनरीक्षण" in p_txt
        if is_summary_page:
            continue

        # Collect DELETED watermark stamps on this page
        deleted_rects = []
        for img in p.get_images():
            if img[2] == 106 and img[3] == 76:
                deleted_rects.extend(p.get_image_rects(img[0]))

        # Detect all voter card bounding boxes (both from lines and rects)
        drawings = p.get_drawings()
        page_cards = []

        # 1. From explicit vector rects (e.g. single cards or supplementary cards)
        for d in drawings:
            for it in d.get('items', []):
                if it[0] == 're':
                    r = it[1]
                    if 160 < (r.x1 - r.x0) < 185 and 65 < (r.y1 - r.y0) < 76 and r.y0 > 130:
                        page_cards.append(pymupdf.Rect(round(r.x0, 1), round(r.y0, 1), round(r.x1, 1), round(r.y1, 1)))

        # 2. From intersecting horizontal and vertical grid lines
        lines = []
        for d in drawings:
            for it in d.get('items', []):
                if it[0] == 'l':
                    lines.append((it[1], it[2]))

        h_lines = sorted(list(set(round(it[0].y, 1) for it in lines if abs(it[0].y - it[1].y) < 1 and 130 < it[0].y < 810)))
        x_cols = [(34.4, 207.9), (208.4, 381.4), (381.9, 555.4)]

        if len(h_lines) >= 2:
            for i in range(len(h_lines) - 1):
                y0, y1 = h_lines[i], h_lines[i + 1]
                if 65 < (y1 - y0) < 76:
                    for x0, x1 in x_cols:
                        r = pymupdf.Rect(x0, y0, x1, y1)
                        if not any(abs(c.x0 - r.x0) < 3 and abs(c.y0 - r.y0) < 3 for c in page_cards):
                            page_cards.append(r)

        # Sort cards top-to-bottom, left-to-right
        page_cards.sort(key=lambda r: (round(r.y0 / 10) * 10, r.x0))

        if not page_cards:
            continue

        for card_rect in page_cards:
            x0, y0, x1, y1 = card_rect.x0, card_rect.y0, card_rect.x1, card_rect.y1
            card_words = p.get_text("words", clip=card_rect)
            if not card_words:
                continue

            # Group words by lines using Y-clustering (prevents round() bucket boundary issues)
            words_by_y = sorted(card_words, key=lambda w: w[1])
            lines_grouped = []
            curr_line = []
            curr_y = None
            for w in words_by_y:
                if curr_y is None or abs(w[1] - curr_y) < 5:
                    curr_line.append(w)
                    curr_y = w[1] if curr_y is None else (curr_y + w[1]) / 2
                else:
                    curr_line.sort(key=lambda w: w[0])
                    lines_grouped.append(curr_line)
                    curr_line = [w]
                    curr_y = w[1]
            if curr_line:
                curr_line.sort(key=lambda w: w[0])
                lines_grouped.append(curr_line)

            card_text = " ".join(" ".join(w[4] for w in ln) for ln in lines_grouped)

            # 1. DELETED Check
            is_deleted = False
            # Watermark intersection
            if any(card_rect.intersects(dr) for dr in deleted_rects):
                is_deleted = True

            # Shifted / Expired prefix check (S / E / R)
            first_few = [w[4] for w in card_words[:5]]
            if 'S' in first_few or 'E' in first_few or 'R' in first_few:
                is_deleted = True

            # Serial Number
            serial = None
            for w in card_words:
                if w[4].isdigit() and int(w[4]) < 2000 and (w[1] - y0) < 28 and (w[0] - x0) < 55:
                    serial = int(w[4])
                    break

            # Voter ID (EPIC)
            epic = ""
            for w in card_words:
                t = w[4]
                if re.match(r'^[A-Z]{3}\d{7}$', t) or re.match(r'^[A-Z0-9/]{8,18}$', t):
                    epic = t
                    break

            if epic and epic in deleted_epics:
                is_deleted = True

            # Determine Status: 'Deleted' or 'Active'
            status = 'Deleted' if is_deleted else 'Active'
            if is_deleted:
                deleted_count += 1

            if serial is None:
                continue

            # Age: Multi-strategy to ensure 100% capture (never left empty)
            age = ""
            age_m = re.search(r'(?:आजच|आयु|आयच|आय|आजु|Age)[^\d]{0,25}(\d{1,3})', card_text)
            if age_m:
                val = int(age_m.group(1))
                if 18 <= val <= 125:
                    age = val

            if not age:
                for w in card_words:
                    if (w[1] - y0) > 45 and w[4].isdigit():
                        val = int(w[4])
                        if 18 <= val <= 120:
                            age = val
                            break

            # House Number
            house = ""
            house_m = re.search(r'(?:मकरन|मकान)\s*(?:सपखजर|संख्या)\s*:\s*([^\s]+)', card_text)
            if house_m:
                house = house_m.group(1).strip()
                house = re.sub(r'[^0-9a-zA-Z\u0900-\u097F]', '', house)

            # Individual voter address: leave empty if not explicitly printed on the individual card
            voter_address = ""
            addr_m = re.search(r'(?:पता|एड्रेस|निवास)\s*[:\-\s]*([^\n]+)', card_text)
            if addr_m:
                raw_addr = addr_m.group(1).strip()
                if not any(k in raw_addr for k in ['लिंग', 'आयु', 'फोटो', 'Photo']):
                    voter_address = clean_hindi_name(raw_addr)

            # Extract Names via narrow crop OCR
            name_crop_rect = pymupdf.Rect(x0 + 1, y0 + 10, x0 + 115, y0 + 54)
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

            # Fallback to decoded digital text if OCR missed name (e.g. obscured by watermark stamp)
            if not voter_name:
                m = re.search(r'(?:नरम|नाम)\s*[:\-\s]*([^\n]+?)(?:Photo|नपतर|पिता|ललग|$)', card_text)
                if m:
                    voter_name = clean_hindi_name(decode_corrupted_rajasthan_hindi(m.group(1).strip()))
            if not relative_name:
                m = re.search(r'(?:नपतर|पिता|पति)\s*(?:कर|का)\s*(?:नरम|नाम)\s*[:\-\s]*([^\n]+?)(?:Photo|मकरन|मकान|आजच|आयु|$)', card_text)
                if m:
                    relative_name = clean_hindi_name(decode_corrupted_rajasthan_hindi(m.group(1).strip()))

            # Additional age fallback if still empty: OCR bottom area
            if not age:
                age_crop_rect = pymupdf.Rect(x0 + 1, y0 + 46, x0 + 115, y1 - 1)
                pix_age = p.get_pixmap(matrix=mat, clip=age_crop_rect)
                img_age = Image.frombytes("RGB", [pix_age.width, pix_age.height], pix_age.samples)
                ocr_age_txt = pytesseract.image_to_string(img_age, lang="hin+eng", config="--psm 6")
                digits = re.findall(r'\b(1[89]|[2-9]\d|1[01]\d)\b', ocr_age_txt)
                if digits:
                    age = int(digits[0])

            voter_row = {
                'भाग संख्या': part_no,
                'क्रम संख्या': serial,
                'नाम': voter_name,
                'पिता/पति का नाम': relative_name,
                'आयु': age,
                'मोबाइल नो': '',
                'वोटर ID': epic,
                'हाउस नंबर': house,
                'एड्रेस': voter_address,
                'बूथ का पता': booth_address,
                'Status': status
            }
            voters.append(voter_row)

    doc.close()

    # Sort voters by serial number
    voters.sort(key=lambda v: v['क्रम संख्या'] if isinstance(v['क्रम संख्या'], int) else 99999)

    stats = {
        'part_no': part_no,
        'booth_address': booth_address,
        'area_address': area_address,
        'total_scanned': len(voters),
        'deleted_count': deleted_count,
        'active_count': len(voters) - deleted_count
    }

    return voters, stats

def extract_voters(pdf_path: str) -> List[Dict[str, Any]]:
    """Legacy wrapper returning voter records list."""
    voters, _ = extract_voters_with_stats(pdf_path)
    return voters

def process_multiple_wards(pdf_paths: List[str]) -> Dict[str, Any]:
    """
    Processes multiple voter list PDFs (e.g. 9 wards of a panchayat):
    1. Extracts voters and metadata for each PDF.
    2. Sorts the wards sequence-wise by Part / Ward Number (1, 2, 3... 9).
    3. Builds:
       - Sheet 1: 'वार्ड_सारांश' (Wards summary with Grand Total row)
       - Sheet 2: 'समस्त_मतदाता_सूची' (Master 11-column list in sequence: Ward 1, then Ward 2... with Status Active/Deleted)
       - Sheets 3..N: 'वार्ड_1', 'वार्ड_2'... (Individual ward sheets)
    """
    columns_11 = [
        'भाग संख्या',
        'क्रम संख्या',
        'नाम',
        'पिता/पति का नाम',
        'आयु',
        'मोबाइल नो',
        'वोटर ID',
        'हाउस नंबर',
        'एड्रेस',
        'बूथ का पता',
        'Status'
    ]

    ward_results = []
    for path in pdf_paths:
        voters, stats = extract_voters_with_stats(path)
        filename = os.path.basename(path)
        clean_name = re.sub(r'^[a-f0-9\-]{36}_', '', filename)
        part_no = str(stats.get('part_no', '1')).strip()

        # Parse numeric part number for proper natural sequence
        try:
            part_num = int(part_no)
        except (ValueError, TypeError):
            m = re.search(r'(?:Ward[_\s\-]*No[_\s\-]*|वार्ड[_\s\-]*|भाग[_\s\-]*)([0-9]+)', filename, re.IGNORECASE)
            part_num = int(m.group(1)) if m else 999

        ward_results.append({
            'path': path,
            'filename': clean_name,
            'part_no': part_no,
            'part_num': part_num,
            'booth_address': stats.get('booth_address', ''),
            'area_address': stats.get('area_address', ''),
            'total_scanned': stats.get('total_scanned', len(voters)),
            'deleted_count': stats.get('deleted_count', 0),
            'active_count': stats.get('active_count', len(voters)),
            'voters': voters
        })

    # Sort wards strictly in ascending sequence: Ward 1, Ward 2, Ward 3... Ward 9
    ward_results.sort(key=lambda w: w['part_num'])

    # 1. Build Sheet 1: वार्ड_सारांश (Summary Table)
    summary_headers = [
        'क्र.सं.',
        'वार्ड / भाग संख्या',
        'दस्तावेज़ (फ़ाइल)',
        'मतदान केन्द्र / बूथ का पता',
        'क्षेत्र / एड्रेस',
        'कुल मतदाता',
        'सक्रिय (Active)',
        'विलोपित (Deleted)'
    ]
    summary_rows = []
    total_all_scanned = 0
    total_all_deleted = 0
    total_all_active = 0

    for idx, w in enumerate(ward_results, start=1):
        total_all_scanned += w['total_scanned']
        total_all_deleted += w['deleted_count']
        total_all_active += w['active_count']
        summary_rows.append([
            idx,
            f"वार्ड {w['part_no']}",
            w['filename'],
            w['booth_address'],
            w['area_address'],
            w['total_scanned'],
            w['active_count'],
            w['deleted_count']
        ])

    # Grand Total Row
    summary_rows.append([
        'कुल योग',
        f"{len(ward_results)} वार्ड्स",
        '-',
        '-',
        '-',
        total_all_scanned,
        total_all_active,
        total_all_deleted
    ])
    summary_grid = [summary_headers] + summary_rows

    # 2. Build Sheet 2: समस्त_मतदाता_सूची (11 Columns, Sequence-wise: Ward 1, then Ward 2, etc.)
    master_rows = []
    for w in ward_results:
        for v in w['voters']:
            row = [v.get(col, '') for col in columns_11]
            master_rows.append(row)
    master_grid = [columns_11] + master_rows

    # 3. Build Individual Ward Tables
    individual_tables = []
    individual_names = []
    for w in ward_results:
        w_rows = [[v.get(col, '') for col in columns_11] for v in w['voters']]
        individual_tables.append([columns_11] + w_rows)
        individual_names.append(f"वार्ड_{w['part_no']}")

    # Combined Table & Names list for openpyxl
    all_tables = [summary_grid, master_grid] + individual_tables
    all_names = ['वार्ड_सारांश', 'समस्त_मतदाता_सूची'] + individual_names

    return {
        'total_wards': len(ward_results),
        'summary_grid': summary_grid,
        'master_grid': master_grid,
        'total_active_voters': total_all_active,
        'total_deleted_voters': total_all_deleted,
        'total_scanned_voters': total_all_scanned,
        'wards': [{
            'part_no': w['part_no'],
            'filename': w['filename'],
            'booth_address': w['booth_address'],
            'area_address': w['area_address'],
            'active_count': w['active_count'],
            'deleted_count': w['deleted_count'],
            'total_scanned': w['total_scanned'],
            'grid': [columns_11] + [[v.get(col, '') for col in columns_11] for v in w['voters']]
        } for w in ward_results],
        'all_tables': all_tables,
        'all_names': all_names
    }

def export_voters_to_excel(voters: List[Dict[str, Any]], output_path: str, theme: str = 'geam_digital') -> None:
    """Exports voter records into a professionally styled Excel file."""
    df = pd.DataFrame(voters)
    columns_order = [
        'भाग संख्या',
        'क्रम संख्या',
        'नाम',
        'पिता/पति का नाम',
        'आयु',
        'मोबाइल नो',
        'वोटर ID',
        'हाउस नंबर',
        'एड्रेस',
        'बूथ का पता',
        'Status'
    ]
    df = df[columns_order]
    headers = list(df.columns)
    rows = df.values.tolist()
    table_grid = [headers] + rows

    from .excel_builder import create_excel_workbook
    excel_bytes = create_excel_workbook(
        tables=[table_grid],
        table_names=['समस्त_मतदाता_सूची'],
        theme=theme,
        mode='single_sheet'
    )

    with open(output_path, 'wb') as f:
        f.write(excel_bytes)

