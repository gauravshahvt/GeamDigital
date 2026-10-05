import os
import re
import subprocess
import tempfile
from io import BytesIO
from typing import List, Dict, Any, Optional, Tuple
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from pdf_engine.parchi_generator import get_browser_executable, clean_val_str

# ==============================================================
# HINDI TO ENGLISH PHONETIC TRANSLITERATION ENGINE
# ==============================================================

VOWELS = {
    'अ': 'a', 'आ': 'a', 'इ': 'i', 'ई': 'i', 'उ': 'u', 'ऊ': 'u',
    'ऋ': 'ri', 'ए': 'e', 'ऐ': 'ai', 'ओ': 'o', 'औ': 'au',
    'अं': 'an', 'अः': 'ah'
}

MATRAS = {
    'ा': 'a', 'ि': 'i', 'ी': 'i', 'ु': 'u', 'ू': 'u',
    'ृ': 'ri', 'े': 'e', 'ै': 'ai', 'ो': 'o', 'ौ': 'au',
    'ं': 'n', 'ँ': 'n', 'ः': 'h', '्': ''
}

CONSONANTS = {
    'क': 'k', 'ख': 'kh', 'ग': 'g', 'घ': 'gh', 'ङ': 'ng',
    'च': 'ch', 'छ': 'chh', 'ज': 'j', 'झ': 'jh', 'ञ': 'ny',
    'ट': 't', 'ठ': 'th', 'ड': 'd', 'ढ': 'dh', 'ण': 'n',
    'त': 't', 'थ': 'th', 'द': 'd', 'ध': 'dh', 'न': 'n',
    'प': 'p', 'फ': 'ph', 'ब': 'b', 'भ': 'bh', 'म': 'm',
    'य': 'y', 'र': 'r', 'ल': 'l', 'व': 'v',
    'श': 'sh', 'ष': 'sh', 'स': 's', 'ह': 'h',
    'क्ष': 'ksh', 'त्र': 'tr', 'ज्ञ': 'gy', 'श्र': 'shr',
    'क़': 'k', 'ख़': 'kh', 'ग़': 'gh', 'ज़': 'z', 'फ़': 'f', 'ड़': 'r', 'ढ़': 'rh'
}

def hindi_to_english_sort_key(text: str) -> str:
    """
    Converts Hindi (Devanagari) names into Latin / English phonetic sort keys
    so that Hindi names sort in exact English A, B, C, D order.
    Example:
      अब्दुल -> abdul
      अभिजीत -> abhijit
      अभिषेक -> abhishek
      आदर्श -> adarsh
      अदिति -> aditi
      आदित्य -> aditya
      अजय -> ajay
      अजीत -> ajit
      आकांक्षा -> akanksha
      आकाश -> akash
      अक्षय -> akshay
    """
    if not text:
        return ""
    text = str(text).strip()
    
    # If already Latin text, return lowercase normalized
    if re.match(r'^[a-zA-Z0-9\s.,-]+$', text):
        return re.sub(r'[^a-z0-9]', '', text.lower())

    # Pre-cleaning of Devanagari nuktas
    text = text.replace('क़', 'क').replace('ख़', 'ख').replace('ग़', 'ग')
    
    result = []
    i = 0
    n = len(text)
    while i < n:
        char = text[i]
        
        # Check conjunct / 2-char combinations
        if i + 1 < n and text[i:i+2] in CONSONANTS:
            result.append(CONSONANTS[text[i:i+2]])
            i += 2
            continue
            
        if char in VOWELS:
            result.append(VOWELS[char])
            i += 1
        elif char in CONSONANTS:
            base = CONSONANTS[char]
            # Check next character for matra or halant
            if i + 1 < n:
                next_char = text[i+1]
                if next_char in MATRAS:
                    if next_char == '्':  # Halant (half letter)
                        result.append(base)
                    else:
                        result.append(base + MATRAS[next_char])
                    i += 2
                else:
                    # Inherent vowel 'a'
                    result.append(base + 'a')
                    i += 1
            else:
                result.append(base)
                i += 1
        elif char in MATRAS:
            result.append(MATRAS[char])
            i += 1
        elif char in ' \t\n-./':
            result.append(' ')
            i += 1
        else:
            result.append(char.lower())
            i += 1
            
    key = "".join(result).lower()
    # Normalize double vowels and clean whitespace
    key = re.sub(r'\s+', ' ', key).strip()
    key = key.replace('ee', 'i').replace('oo', 'u')
    return key

def normalize_gender(val: Any) -> str:
    s = str(val or '').strip().lower()
    if not s:
        return 'पुरूष'
    if s in ['f', 'female', 'woman', 'f.'] or any(k in s for k in ['स्त्री', 'महिला', 'female', 'stree', 'mahila', 'सल']):
        return 'स्त्री'
    if s in ['m', 'male', 'man', 'm.'] or any(k in s for k in ['पुरूष', 'पुरुष', 'male', 'purush', 'पचरष', 'पु.']):
        return 'पुरूष'
    if s == 'पु':
        return 'पुरूष'
    if s == 'म':
        return 'स्त्री'
    return 'पुरूष'

def sort_voters_alphabetically(voters: List[Dict[str, Any]], group_by_ward: bool = False) -> List[Dict[str, Any]]:
    """
    Sorts a voter list by Devanagari transliteration key (A, B, C, D order).
    """
    for v in voters:
        if '_sort_key' not in v:
            v['_sort_key'] = hindi_to_english_sort_key(v.get('name', ''))
            v['_rel_sort_key'] = hindi_to_english_sort_key(v.get('relative_name', ''))

    if group_by_ward:
        # Group by part_no/ward first, then by ABCD name
        def ward_int_key(v):
            p = str(v.get('part_no', v.get('ward', '1'))).strip()
            digits = re.findall(r'\d+', p)
            return int(digits[0]) if digits else 999999

        return sorted(voters, key=lambda v: (ward_int_key(v), v.get('_sort_key', ''), v.get('_rel_sort_key', ''), v.get('serial', 0)))
    else:
        return sorted(voters, key=lambda v: (v.get('_sort_key', ''), v.get('_rel_sort_key', ''), v.get('serial', 0)))

# ==============================================================
# HTML & PDF GENERATOR ENGINE FOR ALPHABETICAL VOTER LIST
# ==============================================================

def generate_alphabetical_html(
    voters: List[Dict[str, Any]],
    ward_title: str = "वार्ड नं.- 8",
    part_no: str = "1",
    booth_address: str = "28 - राजकीय बालिका उच्च माध्यमिक विद्यालय बापूनगर भीलवाडा कमरा न. 7",
    list_title: str = "अल्फाबेटिक ABCD से वोटर लिस्ट",
    filter_active_only: bool = True,
    selected_ward: Optional[str] = None,
    for_preview: bool = False,
    preview_page: int = 1,
    rows_per_page: int = 31,
    jila_parishad: Optional[str] = None,
    panchayat_samiti: Optional[str] = None
) -> Tuple[str, int, int]:
    """
    Renders high-quality A4 printable HTML for Alphabetical Voter List.
    Matching the reference layout:
      - Top Header: Ward No (Left), Title (Center: list_title), Part No (Right)
      - Subheader: Booth Address / Polling Station
      - Table Header (Grey background)
      - 8 Columns: क्रम संख्या, मकान संख्या, मतदाता का नाम, पिता/पति का नाम, उम्र, लिंग, मतदाता पहचान पत्र, मोबाइल नं.
      - Footer: Page X of Y
    """
    # Filter active only if enabled
    filtered = voters
    if filter_active_only:
        filtered = [v for v in filtered if v.get('status') != 'Deleted']

    def extract_ward_key(val: Any) -> Tuple[int, str]:
        s = str(val or '').strip()
        digits = re.findall(r'\d+', s)
        if digits:
            return (int(digits[0]), s)
        return (999999, s)

    def extract_ward_num(val: Any) -> str:
        s = str(val or '').strip()
        digits = re.findall(r'\d+', s)
        if digits:
            return str(int(digits[0]))
        return s if s else '1'

    # Filter by specific ward if selected
    if selected_ward and str(selected_ward).strip() not in ['all', '', 'समस्त', 'सभी']:
        target_ward = extract_ward_num(selected_ward)
        filtered = [v for v in filtered if extract_ward_num(v.get('part_no', v.get('ward', '1'))) == target_ward]

    total_voters = len(filtered)

    # Group voters by ward
    ward_map: Dict[str, List[Dict[str, Any]]] = {}
    for v in filtered:
        w_num = extract_ward_num(v.get('part_no', v.get('ward', '1')))
        if w_num not in ward_map:
            ward_map[w_num] = []
        ward_map[w_num].append(v)

    sorted_wards = sorted(ward_map.keys(), key=extract_ward_key)

    all_pages = []
    for w_num in sorted_wards:
        w_voters = ward_map[w_num]
        sorted_w_voters = sort_voters_alphabetically(w_voters, group_by_ward=False)

        # Directly take JP & PS from voter records of this ward
        w_jp = next((str(v.get('jila_parishad')).strip() for v in w_voters if v.get('jila_parishad')), str(jila_parishad or '').strip())
        w_ps = next((str(v.get('panchayat_samiti')).strip() for v in w_voters if v.get('panchayat_samiti')), str(panchayat_samiti or '').strip())
        w_booth = next((str(v.get('booth_address')).strip() for v in w_voters if v.get('booth_address')), str(booth_address or '').strip())

        w_disp = f"वार्ड संख्या : {w_num}"
        right_parts = []
        if w_jp:
            right_parts.append(f"जि. प. : <b>{w_jp}</b>")
        if w_ps:
            right_parts.append(f"पं. स. : <b>{w_ps}</b>")
        right_parts.append(f"<b>{w_disp}</b>")
        w_right_header_html = " &nbsp;|&nbsp; ".join(right_parts)

        w_total_pages = max(1, (len(sorted_w_voters) + rows_per_page - 1) // rows_per_page)
        chunks = [sorted_w_voters[i * rows_per_page : (i + 1) * rows_per_page] for i in range(w_total_pages)]

        for p_idx, chunk in enumerate(chunks, start=1):
            all_pages.append({
                'ward_num': w_num,
                'ward_title': ward_title,
                'list_title': list_title,
                'right_header_html': w_right_header_html,
                'booth_address': w_booth,
                'chunk': chunk,
                'page_in_ward': p_idx,
                'total_in_ward': w_total_pages,
                'ward_voters': len(w_voters)
            })

    if not all_pages:
        w_disp = f"वार्ड संख्या : {part_no}" if not str(part_no).strip().startswith("वार्ड") else str(part_no).strip()
        right_parts = []
        if jila_parishad:
            right_parts.append(f"जि. प. : <b>{jila_parishad}</b>")
        if panchayat_samiti:
            right_parts.append(f"पं. स. : <b>{panchayat_samiti}</b>")
        right_parts.append(f"<b>{w_disp}</b>")
        all_pages.append({
            'ward_num': str(part_no),
            'ward_title': ward_title,
            'list_title': list_title,
            'right_header_html': " &nbsp;|&nbsp; ".join(right_parts),
            'booth_address': booth_address,
            'chunk': [],
            'page_in_ward': 1,
            'total_in_ward': 1,
            'ward_voters': 0
        })

    total_pages = len(all_pages)

    if for_preview:
        p_idx = max(1, min(preview_page, total_pages))
        render_pages = [(p_idx, all_pages[p_idx - 1])]
    else:
        render_pages = list(enumerate(all_pages, start=1))

    html = f"""<!DOCTYPE html>
<html lang="hi">
<head>
<meta charset="utf-8">
<title>Geam Digital - {ward_title} - {list_title}</title>
<style>
@page {{
    size: A4 portrait;
    margin: 8mm 6mm 8mm 6mm;
}}
* {{
    box-sizing: border-box;
    margin: 0;
    padding: 0;
}}
body {{
    font-family: 'Nirmala UI', 'Segoe UI', Arial, sans-serif;
    background: #fff;
    color: #000;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
}}
@media print {{
    body {{
        background: #fff;
    }}
    .alpha-sheet-page {{
        margin: 0 auto;
        page-break-after: always;
        break-after: page;
        height: 281mm !important;
    }}
}}
.alpha-sheet-page {{
    width: 198mm;
    min-height: 281mm;
    margin: 0 auto 10mm auto;
    padding: 2mm 0;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    background: #fff;
    box-sizing: border-box;
    page-break-after: always;
    break-after: page;
}}
.alpha-header-container {{
    margin-bottom: 2mm;
    border-bottom: 1px solid #111;
    padding-bottom: 2mm;
}}
.alpha-header-row {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    line-height: 1.2;
    margin-bottom: 1.5mm;
}}
.alpha-header-left {{
    font-size: 11.5pt;
    font-weight: 800;
    color: #000;
    text-align: left;
    width: 30%;
}}
.alpha-header-center {{
    font-size: 13.5pt;
    font-weight: 900;
    color: #000;
    text-align: center;
    width: 34%;
    letter-spacing: -0.2px;
}}
.alpha-header-right {{
    font-size: 10.5pt;
    font-weight: 800;
    color: #000;
    text-align: right;
    width: 36%;
    white-space: nowrap;
}}
.alpha-subheader {{
    font-size: 9.8pt;
    font-weight: 800;
    color: #111;
    line-height: 1.25;
    text-align: left;
}}
.alpha-table {{
    width: 100%;
    border-collapse: collapse;
    border: 1.2px solid #000;
    font-size: 8.5pt;
    line-height: 1.15;
    table-layout: fixed;
}}
.alpha-table thead th {{
    background-color: #cbd5e1;
    color: #000;
    font-weight: 800;
    font-size: 8.2pt;
    padding: 2mm 1mm;
    border: 1px solid #000;
    text-align: center;
    white-space: nowrap;
}}
.alpha-table tbody td {{
    padding: 1.6mm 1.2mm;
    border: 1px solid #64748b;
    vertical-align: middle;
    color: #000;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}
.alpha-table tbody tr:nth-child(even) {{
    background-color: #f8fafc;
}}
.col-serial {{
    width: 7.5%;
    text-align: center;
    font-weight: 800;
    font-family: Arial, sans-serif;
}}
.col-house {{
    width: 13%;
    text-align: center;
    font-weight: 700;
}}
.col-name {{
    width: 21%;
    text-align: left;
    font-weight: 800;
    padding-left: 2mm !important;
}}
.col-rel {{
    width: 21%;
    text-align: left;
    padding-left: 2mm !important;
}}
.col-age {{
    width: 5.5%;
    text-align: center;
    font-weight: 700;
    font-family: Arial, sans-serif;
}}
.col-gender {{
    width: 7.5%;
    text-align: center;
    font-weight: 700;
}}
.col-epic {{
    width: 16.5%;
    text-align: center;
    font-weight: 800;
    font-family: Arial, sans-serif;
    letter-spacing: 0.2px;
}}
.col-mobile {{
    width: 8%;
    text-align: center;
}}
.alpha-footer {{
    display: flex;
    justify-content: flex-end;
    align-items: center;
    padding-top: 1.5mm;
    font-size: 8.5pt;
    font-family: Arial, sans-serif;
    color: #334155;
    font-weight: 600;
}}
</style>
</head>
<body>
"""

    for global_p_idx, p_data in render_pages:
        html += f"""
<div class="alpha-sheet-page" data-page="{global_p_idx}">
    <div>
        <!-- Top Header as per reference image -->
        <div class="alpha-header-container">
            <div class="alpha-header-row">
                <div class="alpha-header-left">{p_data['ward_title']}</div>
                <div class="alpha-header-center">{p_data['list_title']}</div>
                <div class="alpha-header-right">{p_data['right_header_html']}</div>
            </div>
            <div class="alpha-subheader">
                {p_data['booth_address']}
            </div>
        </div>

        <!-- 8-Column Data Table -->
        <table class="alpha-table">
            <thead>
                <tr>
                    <th class="col-serial">क्रम संख्या</th>
                    <th class="col-house">मकान संख्या</th>
                    <th class="col-name">मतदाता का नाम</th>
                    <th class="col-rel">पिता / पति का नाम</th>
                    <th class="col-age">उम्र</th>
                    <th class="col-gender">लिंग</th>
                    <th class="col-epic">मतदाता पहचान पत्र</th>
                    <th class="col-mobile">मोबाइल नं.</th>
                </tr>
            </thead>
            <tbody>
"""
        for v in p_data['chunk']:
            s_val = v.get('serial', '')
            h_val = v.get('house', '')
            if not h_val or str(h_val).strip() in ['00', '0', '-']:
                h_val = ''
            name_val = v.get('name', '')
            rel_val = v.get('relative_name', '')
            age_val = v.get('age', '')
            gender_val = normalize_gender(v.get('gender', ''))
            epic_val = v.get('epic', '')
            mobile_val = ''

            html += f"""
                <tr>
                    <td class="col-serial">{s_val}</td>
                    <td class="col-house">{h_val}</td>
                    <td class="col-name">{name_val}</td>
                    <td class="col-rel">{rel_val}</td>
                    <td class="col-age">{age_val}</td>
                    <td class="col-gender">{gender_val}</td>
                    <td class="col-epic">{epic_val}</td>
                    <td class="col-mobile">{mobile_val}</td>
                </tr>
"""

        html += f"""
            </tbody>
        </table>
    </div>

    <!-- Bottom Footer (Page X of Y per Ward) -->
    <div class="alpha-footer">
        Page {p_data['page_in_ward']} of {p_data['total_in_ward']}
    </div>
</div>
"""

    html += """</body></html>"""
    cur_p = render_pages[0][1] if render_pages else all_pages[0]
    page_meta = {
        'current_ward': str(cur_p.get('ward_num', '1')),
        'page_in_ward': int(cur_p.get('page_in_ward', 1)),
        'total_in_ward': int(cur_p.get('total_in_ward', 1)),
        'ward_voters': int(cur_p.get('ward_voters', total_voters)),
        'total_wards': len(sorted_wards)
    }
    return html, total_pages, total_voters, page_meta

def generate_alphabetical_pdf(
    voters: List[Dict[str, Any]],
    ward_title: str = "वार्ड नं.- 8",
    part_no: str = "1",
    booth_address: str = "28 - राजकीय बालिका उच्च माध्यमिक विद्यालय बापूनगर भीलवाडा कमरा न. 7",
    list_title: str = "अल्फाबेटिक ABCD से वोटर लिस्ट",
    filter_active_only: bool = True,
    selected_ward: Optional[str] = None,
    jila_parishad: Optional[str] = None,
    panchayat_samiti: Optional[str] = None
) -> bytes:
    """
    Generates a production A4 PDF for Alphabetical Voter list using headless Chrome/Edge.
    """
    html_content, _, _, _ = generate_alphabetical_html(
        voters=voters,
        ward_title=ward_title,
        part_no=part_no,
        booth_address=booth_address,
        list_title=list_title,
        filter_active_only=filter_active_only,
        selected_ward=selected_ward,
        for_preview=False,
        rows_per_page=31,
        jila_parishad=jila_parishad,
        panchayat_samiti=panchayat_samiti
    )

    browser_bin = get_browser_executable()
    if not browser_bin:
        raise RuntimeError("No headless Chrome or Edge browser found on system to generate PDF.")

    with tempfile.TemporaryDirectory() as tmpdir:
        html_file = os.path.join(tmpdir, "alpha_voters.html")
        pdf_file = os.path.join(tmpdir, "alpha_voters.pdf")

        with open(html_file, "w", encoding="utf-8") as f:
            f.write(html_content)

        cmd = [
            browser_bin,
            "--headless",
            "--disable-gpu",
            f"--print-to-pdf={pdf_file}",
            "--no-pdf-header-footer",
            html_file
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if res.returncode != 0 or not os.path.exists(pdf_file):
            raise RuntimeError(f"Browser PDF generation failed: {res.stderr}")

        with open(pdf_file, "rb") as f:
            pdf_bytes = f.read()

    return pdf_bytes

def generate_alphabetical_excel(
    voters: List[Dict[str, Any]],
    ward_title: str = "वार्ड नं.- 8",
    part_no: str = "1",
    booth_address: str = "28 - राजकीय बालिका उच्च माध्यमिक विद्यालय बापूनगर भीलवाडा कमरा न. 7",
    list_title: str = "अल्फाबेटिक ABCD से वोटर लिस्ट",
    filter_active_only: bool = True,
    selected_ward: Optional[str] = None,
    jila_parishad: Optional[str] = None,
    panchayat_samiti: Optional[str] = None
) -> bytes:
    """
    Exports a styled Excel (.xlsx) file with voters sorted by English ABCD letters, organized ward-wise.
    """
    filtered = voters
    if filter_active_only:
        filtered = [v for v in filtered if v.get('status') != 'Deleted']

    def extract_ward_key(val: Any) -> Tuple[int, str]:
        s = str(val or '').strip()
        digits = re.findall(r'\d+', s)
        if digits:
            return (int(digits[0]), s)
        return (999999, s)

    def extract_ward_num(val: Any) -> str:
        s = str(val or '').strip()
        digits = re.findall(r'\d+', s)
        if digits:
            return str(int(digits[0]))
        return s if s else '1'

    if selected_ward and str(selected_ward).strip() not in ['all', '', 'समस्त', 'सभी']:
        target_ward = extract_ward_num(selected_ward)
        filtered = [v for v in filtered if extract_ward_num(v.get('part_no', v.get('ward', '1'))) == target_ward]

    ward_map: Dict[str, List[Dict[str, Any]]] = {}
    for v in filtered:
        w_num = extract_ward_num(v.get('part_no', v.get('ward', '1')))
        if w_num not in ward_map:
            ward_map[w_num] = []
        ward_map[w_num].append(v)

    sorted_wards = sorted(ward_map.keys(), key=extract_ward_key)
    if not sorted_wards:
        sorted_wards = [str(part_no or '1')]
        ward_map[sorted_wards[0]] = []

    wb = openpyxl.Workbook()

    # Styling definitions
    font_title = Font(name="Nirmala UI", size=14, bold=True, color="FFFFFF")
    font_sub = Font(name="Nirmala UI", size=10, bold=True, color="1E293B")
    font_header = Font(name="Nirmala UI", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Nirmala UI", size=10)
    font_data_bold = Font(name="Nirmala UI", size=10, bold=True)

    fill_title = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_sub = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_header = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    thin_border_side = Side(style='thin', color='CBD5E1')
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    headers = [
        "क्रम संख्या",
        "मकान संख्या",
        "मतदाता का नाम",
        "पिता / पति का नाम",
        "उम्र",
        "लिंग",
        "मतदाता पहचान पत्र",
        "मोबाइल नं."
    ]

    for s_idx, w_num in enumerate(sorted_wards):
        if s_idx == 0:
            ws = wb.active
            ws.title = f"वार्ड_{w_num}"
        else:
            ws = wb.create_sheet(title=f"वार्ड_{w_num}")

        w_voters = ward_map[w_num]
        sorted_w_voters = sort_voters_alphabetically(w_voters, group_by_ward=False)

        w_jp = next((str(v.get('jila_parishad')).strip() for v in w_voters if v.get('jila_parishad')), str(jila_parishad or '').strip())
        w_ps = next((str(v.get('panchayat_samiti')).strip() for v in w_voters if v.get('panchayat_samiti')), str(panchayat_samiti or '').strip())
        w_booth = next((str(v.get('booth_address')).strip() for v in w_voters if v.get('booth_address')), str(booth_address or '').strip())

        # Row 1: Title Banner
        ws.merge_cells('A1:H1')
        c1 = ws['A1']
        title_parts = [ward_title, list_title]
        if w_jp:
            title_parts.append(f"जि. प. : {w_jp}")
        if w_ps:
            title_parts.append(f"पं. स. : {w_ps}")
        title_parts.append(f"वार्ड संख्या : {w_num}")
        c1.value = " | ".join(title_parts)
        c1.font = font_title
        c1.fill = fill_title
        c1.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        # Row 2: Booth Address
        ws.merge_cells('A2:H2')
        c2 = ws['A2']
        c2.value = f"मतदान केंद्र: {w_booth}"
        c2.font = font_sub
        c2.fill = fill_sub
        c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[2].height = 22

        # Row 3: Headers
        ws.append(headers)
        ws.row_dimensions[3].height = 24
        for col_idx in range(1, 9):
            cell = ws.cell(row=3, column=col_idx)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_cell

        # Rows 4+: Data
        for r_idx, v in enumerate(sorted_w_voters, start=4):
            s_val = v.get('serial', '')
            h_val = v.get('house', '')
            if not h_val or str(h_val).strip() in ['00', '0', '-']:
                h_val = ''
            name_val = v.get('name', '')
            rel_val = v.get('relative_name', '')
            age_val = v.get('age', '')
            gender_val = normalize_gender(v.get('gender', ''))
            epic_val = v.get('epic', '')
            mobile_val = ''

            row_vals = [s_val, h_val, name_val, rel_val, age_val, gender_val, epic_val, mobile_val]
            ws.append(row_vals)
            ws.row_dimensions[r_idx].height = 20

            is_even = (r_idx % 2 == 0)
            for c_idx in range(1, 9):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = font_data_bold if c_idx in [1, 3, 7] else font_data
                cell.border = border_cell
                if is_even:
                    cell.fill = fill_zebra
                if c_idx in [1, 2, 5, 6, 7]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # Column widths
        col_widths = {1: 12, 2: 15, 3: 26, 4: 26, 5: 10, 6: 12, 7: 22, 8: 15}
        for c_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(c_idx)].width = width

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
