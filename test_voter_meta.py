import sys
import re
import pymupdf
import pytesseract
from PIL import Image

sys.stdout.reconfigure(encoding='utf-8')

pdf_path = 'uploads/b9707b19-9ff5-425e-bd9d-f0973deb1ea6_BAGMALI-Ward_No-001.pdf'
doc = pymupdf.open(pdf_path)

print(f"Total pages in doc: {len(doc)}")

# 1. Extract cover page metadata from Page 1
p1 = doc[0]
p1_text = p1.get_text()

# Extract booth address
booth_match = re.search(r'मतदान\s*बूथ\s*की\s*संख्या\s*एवं\s*पता\s*:\s*([^\n]+)', p1_text)
# or look at OCR of page 1 if font is scrambled
print("Page 1 text sample:", p1_text[:300].replace('\n', ' | '))

# Extract Deleted serial numbers from Page 14 (Deletion list)
p14 = doc[13] if len(doc) >= 14 else None
deleted_serials = set()
if p14:
    p14_words = p14.get_text("words")
    # In page 14, boxes contain serial numbers like '1', '18', '19', '20', '23', '24', '25', '26', '27', '40', '180', '243'
    # with 'S' or 'E'
    for i, w in enumerate(p14_words):
        if w[4] in ['S', 'E', 'R']:
            # look at next word or surrounding words for digits
            for nw in p14_words[max(0, i-2):min(len(p14_words), i+3)]:
                if nw[4].isdigit():
                    deleted_serials.add(int(nw[4]))

print("Deleted serials found on page 14:", sorted(list(deleted_serials)))
