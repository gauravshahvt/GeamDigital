import sys
import re
import pymupdf
import pytesseract
from PIL import Image

sys.stdout.reconfigure(encoding='utf-8')

doc = pymupdf.open('uploads/b9707b19-9ff5-425e-bd9d-f0973deb1ea6_BAGMALI-Ward_No-001.pdf')
p = doc[2] # page 3

# Deletion markers
deleted_img_xrefs = [img[0] for img in p.get_images() if img[2] == 106 and img[3] == 76]
deleted_rects = []
for xref in deleted_img_xrefs:
    deleted_rects.extend(p.get_image_rects(xref))

print(f"Deleted image rects on page 3: {len(deleted_rects)}")

# Grid definition on page 3
# y: 159.8 to 738.6, step 72.35
# x: 34.4, 208.4, 381.9, 555.4
x_cols = [(34.4, 207.9), (208.4, 381.4), (381.9, 555.4)]
y_rows = [
    (159.8, 232.1), (232.1, 304.5), (304.5, 376.8), (376.8, 449.2),
    (449.2, 521.5), (521.5, 593.9), (593.9, 666.2), (666.2, 738.6)
]

zoom = 300 / 72.0
mat = pymupdf.Matrix(zoom, zoom)

for r_idx, (y0, y1) in enumerate(y_rows):
    for c_idx, (x0, x1) in enumerate(x_cols):
        card_rect = pymupdf.Rect(x0, y0, x1, y1)
        
        # Check if DELETED
        is_deleted = any(card_rect.intersects(dr) for dr in deleted_rects)
        
        # PyMuPDF words in this card
        card_words = p.get_text("words", clip=card_rect)
        card_text = " ".join(w[4] for w in sorted(card_words, key=lambda w: (round(w[1]/10)*10, w[0])))
        
        # Check if serial starts with S, E, R
        first_few = [w[4] for w in card_words[:4]]
        if 'S' in first_few or 'E' in first_few or 'R' in first_few:
            is_deleted = True

        # Extract Serial Number
        serial = None
        for w in card_words:
            if w[4].isdigit() and int(w[4]) < 500 and w[1] - y0 < 25 and w[0] - x0 < 45:
                serial = int(w[4])
                break

        # Extract Voter ID (EPIC)
        epic = ""
        for w in card_words:
            t = w[4]
            if re.match(r'^[A-Z]{3}\d{7}$', t) or re.match(r'^[A-Z0-9/]{8,16}$', t):
                epic = t
                break

        # Extract Age
        age = ""
        age_m = re.search(r'(?:आजच|आयु)\s*:\s*(\d{1,3})', card_text)
        if age_m:
            age = age_m.group(1)

        # Extract House Number
        house = ""
        house_m = re.search(r'(?:मकरन|मकान)\s*(?:सपखजर|संख्या)\s*:\s*([^\s]+)', card_text)
        if house_m:
            house = house_m.group(1)

        # If deleted, skip
        if is_deleted:
            print(f"Skipping DELETED Serial: {serial} (EPIC: {epic})")
            continue

        # Run OCR on the name block of the card (left 65% of the card, excluding serial and photo)
        # y from y0 + 14 to y0 + 52, x from x0 + 2 to x0 + 115
        name_crop_rect = pymupdf.Rect(x0 + 1, y0 + 14, x0 + 120, y0 + 52)
        pix = p.get_pixmap(matrix=mat, clip=name_crop_rect)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        ocr_lines = pytesseract.image_to_string(img, lang="hin", config="--psm 6").strip().split('\n')
        ocr_lines = [l.strip() for l in ocr_lines if l.strip()]

        voter_name = ""
        relative_name = ""
        rel_type = "पिता/पति का नाम"

        for line in ocr_lines:
            if "नाम" in line and not ("पिता" in line or "पति" in line):
                voter_name = re.sub(r'^.*?नाम\s*[:\-\s]*', '', line).strip()
            elif "पिता" in line or "पति" in line:
                relative_name = re.sub(r'^.*?(?:पिता|पति)\s*का\s*नाम\s*[:\-\s]*', '', line).strip()

        print(f"Serial {serial:3d} | EPIC: {epic:10s} | Name: {voter_name:18s} | Relative: {relative_name:18s} | Age: {age:2s} | House: {house:4s}")
