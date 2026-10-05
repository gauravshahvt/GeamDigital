"""
pdf_engine/family_voter_list.py
================================
Feature 5: बड़े-घर / बड़े परिवारों की लिस्ट (Large Families / Joint Families Voter List Generator)

Key Features:
1. Groups voters by House Number ('मकान सं' / 'हाउस नंबर').
2. Counts total members in each house ('Total Members :: X').
3. Sorts families in Descending Order of total members (largest families first).
4. Preserves voter roll serial number ('क्रम सं.') within each family.
5. 7 Columns matching reference layout:
   - क्रम सं.
   - मकान सं
   - मतदाता का नाम
   - पिता / पति का नाम
   - उम्र
   - लिंग
   - मतदाता पहचान पत्र
6. Top Header:
   - Left: Ward No (उदा. 'वार्ड न.- 3')
   - Center: Title ('बड़े-घर की लिस्ट')
   - Right: Part No (उदा. 'भाग संख्या : 1')
   - Sub-header: Polling Station / Booth Address
7. Clean A4 Print Layout with 'Total Members :: X' header per family and separator lines.
8. Vector PDF generation via headless Chrome/Edge.
9. Styled Excel (.xlsx) export via openpyxl.
"""

import os
import re
import tempfile
import subprocess
from typing import List, Dict, Any, Tuple, Optional
from collections import defaultdict
from io import BytesIO

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from pdf_engine.parchi_generator import get_browser_executable, clean_val_str

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

# ==============================================================
# 1. GROUPING & SORTING ENGINE FOR LARGE FAMILIES
# ==============================================================

def clean_house_number(val: Any) -> str:
    """
    Cleans and standardizes house numbers:
    strips whitespace, removes trailing .0 for numeric floats.
    """
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith('.0') and s[:-2].isdigit():
        s = s[:-2]
    return s

def is_zero_or_blank_house(h: str) -> bool:
    """
    Checks if house number is a placeholder for homeless or missing house number
    (e.g. '', '0', '00', '000', '-', 'na', 'null', 'none').
    """
    norm = h.lower().strip()
    return norm in ['', '0', '00', '000', '0000', '-', '--', 'na', 'n/a', 'null', 'none', 'nil']

def group_and_sort_families(
    voters: List[Dict[str, Any]],
    min_family_size: int = 2,
    ignore_zero_houses: bool = True,
    selected_ward: Optional[str] = None,
    filter_active_only: bool = True
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Groups voters by house number and sorts families descending by member count.
    
    Returns:
      (families_list, summary_stats)
      where each family item is:
      {
          'house': str,
          'total_members': int,
          'members': List[Dict[str, Any]]
      }
    """
    filtered = voters
    if filter_active_only:
        filtered = [v for v in filtered if v.get('status') != 'Deleted']

    if selected_ward and selected_ward not in ['all', '', 'समस्त', 'सभी']:
        filtered = [v for v in filtered if str(v.get('part_no', v.get('ward', ''))).strip() == str(selected_ward).strip()]

    family_map = defaultdict(list)
    zero_house_voters = []

    for v in filtered:
        h = clean_house_number(v.get('house', ''))
        if is_zero_or_blank_house(h):
            if not ignore_zero_houses:
                zero_house_voters.append(v)
        else:
            family_map[h].append(v)

    families = []
    for h, members in family_map.items():
        if len(members) >= min_family_size:
            # Sort members within family by original serial number
            def sort_serial(item):
                s = str(item.get('serial', '0')).strip()
                digits = re.findall(r'\d+', s)
                return int(digits[0]) if digits else 999999
            
            sorted_members = sorted(members, key=sort_serial)
            families.append({
                'house': h,
                'total_members': len(sorted_members),
                'members': sorted_members
            })

    # Sort families: Largest families first (descending), then numeric/alphabetical house number
    def family_sort_key(f):
        count = f['total_members']
        h_str = f['house']
        digits = re.findall(r'\d+', h_str)
        num_val = int(digits[0]) if digits else 999999
        return (-count, num_val, h_str)

    families.sort(key=family_sort_key)

    # If zero house voters included and >= min_family_size, append at the end
    if not ignore_zero_houses and len(zero_house_voters) >= min_family_size:
        families.append({
            'house': '00 / बिना मकान',
            'total_members': len(zero_house_voters),
            'members': zero_house_voters
        })

    total_family_voters = sum(f['total_members'] for f in families)
    max_family_size = families[0]['total_members'] if families else 0

    stats = {
        'total_families': len(families),
        'total_voters': total_family_voters,
        'max_family_size': max_family_size,
        'min_family_size_applied': min_family_size
    }

    return families, stats

# ==============================================================
# 2. A4 PAGINATION ALGORITHM FOR LARGE FAMILIES
# ==============================================================

def paginate_families(
    families: List[Dict[str, Any]],
    max_lines_per_page: int = 27
) -> List[List[Dict[str, Any]]]:
    """
    Paginates families into A4 pages with ~27 lines per page.
    Prevents awkward single-line orphan family headers.
    Returns a list of pages, where each page is a list of display items:
      - {'type': 'header', 'house': str, 'total_members': int, 'is_contd': bool}
      - {'type': 'row', 'voter': dict, 'house': str, 'is_last_in_family': bool}
    """
    pages = []
    current_page = []
    current_lines = 0

    for f in families:
        h = f['house']
        total_m = f['total_members']
        members = f['members']
        
        # If remaining lines on current page are too few (<= 2 lines), start family on a fresh page
        lines_left = max_lines_per_page - current_lines
        if current_lines > 0 and lines_left <= 2 and total_m >= 2:
            pages.append(current_page)
            current_page = []
            current_lines = 0

        # Start family with header
        current_page.append({
            'type': 'family_header',
            'house': h,
            'total_members': total_m,
            'is_contd': False
        })
        current_lines += 1

        for idx, voter in enumerate(members):
            if current_lines >= max_lines_per_page:
                # Page is full, push to next page
                pages.append(current_page)
                current_page = []
                current_lines = 0
                
                # Show continuation header for this family on the new page
                current_page.append({
                    'type': 'family_header',
                    'house': h,
                    'total_members': total_m,
                    'is_contd': True
                })
                current_lines += 1

            is_last = (idx == len(members) - 1)
            current_page.append({
                'type': 'voter_row',
                'voter': voter,
                'house': h,
                'is_last_in_family': is_last
            })
            current_lines += 1

    if current_page:
        pages.append(current_page)

    return pages if pages else [[]]

# ==============================================================
# 3. HTML & PDF GENERATOR ENGINE FOR FEATURE 5
# ==============================================================

def generate_family_html(
    voters: List[Dict[str, Any]],
    ward_title: str = "वार्ड न.- 3",
    part_no: str = "1",
    booth_address: str = "3 - महात्मा गाँधी राजकीय विद्यालय आसींद कमरा नंबर 1",
    list_title: str = "बड़े-घर की लिस्ट",
    min_family_size: int = 2,
    ignore_zero_houses: bool = True,
    selected_ward: Optional[str] = None,
    filter_active_only: bool = True,
    for_preview: bool = False,
    preview_page: int = 1,
    max_lines_per_page: int = 27,
    jila_parishad: Optional[str] = None,
    panchayat_samiti: Optional[str] = None
) -> Tuple[str, int, int, int]:
    """
    Renders high-quality A4 printable HTML for Feature 5 (बड़े-घर की लिस्ट).
    Matching the reference layout:
      - Top Header: Ward No (Left), Title (Center: 'बड़े-घर की लिस्ट'), Part No (Right)
      - Subheader: Booth Address / Polling Station
      - Table Header (Grey background): 7 Columns
      - Per-family: 'Total Members :: X' header row
      - Member rows: क्रम सं., मकान सं, नाम, पिता/पति का नाम, उम्र, लिंग, मतदाता पहचान पत्र
      - Solid divider border after each family
      - Footer: Page X of Y
    Returns: (html, total_pages, total_voters, total_families)
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

    all_pages = []
    total_voters = 0
    total_families = 0

    for w_num in sorted_wards:
        w_voters = ward_map[w_num]
        families, stats = group_and_sort_families(
            voters=w_voters,
            min_family_size=min_family_size,
            ignore_zero_houses=ignore_zero_houses,
            selected_ward=None,
            filter_active_only=False
        )
        total_voters += stats['total_voters']
        total_families += stats['total_families']

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

        w_sub_parts = []
        if w_booth:
            w_sub_parts.append(f"मतदान केंद्र: {w_booth}")
        w_sub_parts.append(f"कुल बड़े परिवार: <b>{stats['total_families']}</b>")
        w_sub_parts.append(f"कुल मतदाता: <b>{stats['total_voters']}</b>")
        w_subheader_html = " &nbsp;|&nbsp; ".join(w_sub_parts)

        w_pages = paginate_families(families, max_lines_per_page=max_lines_per_page)
        w_total_pages = max(1, len(w_pages))

        for p_idx, p_items in enumerate(w_pages, start=1):
            all_pages.append({
                'ward_num': w_num,
                'ward_title': ward_title,
                'list_title': list_title,
                'right_header_html': w_right_header_html,
                'subheader_html': w_subheader_html,
                'items': p_items,
                'page_in_ward': p_idx,
                'total_in_ward': w_total_pages,
                'ward_voters': stats['total_voters'],
                'ward_families': stats['total_families']
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
            'subheader_html': f"मतदान केंद्र: {booth_address} &nbsp;|&nbsp; कुल बड़े परिवार: <b>0</b> &nbsp;|&nbsp; कुल मतदाता: <b>0</b>",
            'items': [],
            'page_in_ward': 1,
            'total_in_ward': 1,
            'ward_voters': 0,
            'ward_families': 0
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
    .family-sheet-page {{
        margin: 0 auto;
        page-break-after: always;
        break-after: page;
        height: 281mm !important;
    }}
}}
.family-sheet-page {{
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
.family-header-container {{
    margin-bottom: 2mm;
    border-bottom: 1px solid #111;
    padding-bottom: 2mm;
}}
.family-header-row {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    line-height: 1.2;
    margin-bottom: 1.5mm;
}}
.family-header-left {{
    font-size: 11.5pt;
    font-weight: 800;
    color: #000;
    text-align: left;
    width: 30%;
}}
.family-header-center {{
    font-size: 13.5pt;
    font-weight: 900;
    color: #000;
    text-align: center;
    width: 34%;
    letter-spacing: -0.2px;
}}
.family-header-right {{
    font-size: 10.5pt;
    font-weight: 800;
    color: #000;
    text-align: right;
    width: 36%;
    white-space: nowrap;
}}
.family-subheader {{
    font-size: 10pt;
    font-weight: 800;
    color: #111;
    line-height: 1.25;
    text-align: left;
}}
.family-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 8.8pt;
    line-height: 1.2;
    table-layout: fixed;
}}
.family-table thead th {{
    background-color: #cbd5e1;
    color: #000;
    font-weight: 800;
    font-size: 8.5pt;
    padding: 2mm 1mm;
    border: 1px solid #334155;
    text-align: center;
    white-space: nowrap;
}}
.family-table tbody td {{
    padding: 1.6mm 1.5mm;
    vertical-align: middle;
    color: #000;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    border-top: none;
    border-bottom: none;
}}
.family-total-banner {{
    text-align: right;
    font-weight: 800;
    font-size: 9.2pt;
    color: #000;
    padding: 2.2mm 2mm 1mm 2mm !important;
    border-top: none !important;
    border-bottom: none !important;
}}
.family-last-row td {{
    border-bottom: 2px solid #000 !important;
    padding-bottom: 2.5mm !important;
}}
.col-serial {{
    width: 8%;
    text-align: center;
    font-weight: 800;
    font-family: Arial, sans-serif;
}}
.col-house {{
    width: 10%;
    text-align: center;
    font-weight: 800;
}}
.col-name {{
    width: 24%;
    text-align: left;
    font-weight: 800;
    padding-left: 2mm !important;
}}
.col-rel {{
    width: 24%;
    text-align: left;
    padding-left: 2mm !important;
}}
.col-age {{
    width: 6%;
    text-align: center;
    font-weight: 700;
    font-family: Arial, sans-serif;
}}
.col-gender {{
    width: 8%;
    text-align: center;
    font-weight: 700;
}}
.col-epic {{
    width: 20%;
    text-align: left;
    font-weight: 800;
    font-family: Arial, sans-serif;
    letter-spacing: 0.2px;
    padding-left: 2mm !important;
}}
.family-footer {{
    display: flex;
    justify-content: flex-end;
    align-items: center;
    padding-top: 2mm;
    font-size: 8.8pt;
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
<div class="family-sheet-page" data-page="{global_p_idx}">
    <div>
        <!-- Top Header as per reference image -->
        <div class="family-header-container">
            <div class="family-header-row">
                <div class="family-header-left">{p_data['ward_title']}</div>
                <div class="family-header-center">{p_data['list_title']}</div>
                <div class="family-header-right">{p_data['right_header_html']}</div>
            </div>
            <div class="family-subheader">
                {p_data['subheader_html']}
            </div>
        </div>

        <!-- 7-Column Data Table -->
        <table class="family-table">
            <thead>
                <tr>
                    <th class="col-serial">क्रम सं.</th>
                    <th class="col-house">मकान सं</th>
                    <th class="col-name">मतदाता का नाम</th>
                    <th class="col-rel">पिता / पति का नाम</th>
                    <th class="col-age">उम्र</th>
                    <th class="col-gender">लिंग</th>
                    <th class="col-epic">मतदाता पहचान पत्र</th>
                </tr>
            </thead>
            <tbody>
"""
        for item in p_data['items']:
            if item['type'] == 'family_header':
                contd_text = " (जारी)" if item.get('is_contd') else ""
                html += f"""
                <tr>
                    <td colspan="7" class="family-total-banner">
                        Total Members :: &nbsp;{item['total_members']}{contd_text}
                    </td>
                </tr>
"""
            elif item['type'] == 'voter_row':
                v = item['voter']
                s_val = v.get('serial', '')
                h_val = item['house']
                name_val = v.get('name', '')
                rel_val = v.get('relative_name', '')
                age_val = v.get('age', '')
                gender_val = normalize_gender(v.get('gender', ''))
                epic_val = v.get('epic', '')
                
                row_cls = "family-last-row" if item.get('is_last_in_family') else ""

                html += f"""
                <tr class="{row_cls}">
                    <td class="col-serial">{s_val}</td>
                    <td class="col-house">{h_val}</td>
                    <td class="col-name">{name_val}</td>
                    <td class="col-rel">{rel_val}</td>
                    <td class="col-age">{age_val}</td>
                    <td class="col-gender">{gender_val}</td>
                    <td class="col-epic">{epic_val}</td>
                </tr>
"""

        html += f"""
            </tbody>
        </table>
    </div>

    <!-- Bottom Footer (Page X of Y per Ward) -->
    <div class="family-footer">
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
        'ward_families': int(cur_p.get('ward_families', total_families)),
        'total_wards': len(sorted_wards)
    }
    return html, total_pages, total_voters, total_families, page_meta

# ==============================================================
# 4. VECTOR PDF EXPORT ENGINE
# ==============================================================

def generate_family_pdf(
    voters: List[Dict[str, Any]],
    ward_title: str = "वार्ड न.- 3",
    part_no: str = "1",
    booth_address: str = "3 - महात्मा गाँधी राजकीय विद्यालय आसींद कमरा नंबर 1",
    list_title: str = "बड़े-घर की लिस्ट",
    min_family_size: int = 2,
    ignore_zero_houses: bool = True,
    selected_ward: Optional[str] = None,
    filter_active_only: bool = True,
    jila_parishad: Optional[str] = None,
    panchayat_samiti: Optional[str] = None
) -> bytes:
    """
    Generates a production A4 PDF for Feature 5 using headless Chrome/Edge.
    """
    html_content, _, _, _, _ = generate_family_html(
        voters=voters,
        ward_title=ward_title,
        part_no=part_no,
        booth_address=booth_address,
        list_title=list_title,
        min_family_size=min_family_size,
        ignore_zero_houses=ignore_zero_houses,
        selected_ward=selected_ward,
        filter_active_only=filter_active_only,
        for_preview=False,
        max_lines_per_page=27,
        jila_parishad=jila_parishad,
        panchayat_samiti=panchayat_samiti
    )

    browser_bin = get_browser_executable()
    if not browser_bin:
        raise RuntimeError("No headless Chrome or Edge browser found on system to generate PDF.")

    with tempfile.TemporaryDirectory() as tmpdir:
        html_file = os.path.join(tmpdir, "family_voters.html")
        pdf_file = os.path.join(tmpdir, "family_voters.pdf")

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

# ==============================================================
# 5. STYLED EXCEL EXPORT ENGINE
# ==============================================================

def generate_family_excel(
    voters: List[Dict[str, Any]],
    ward_title: str = "वार्ड न.- 3",
    part_no: str = "1",
    booth_address: str = "3 - महात्मा गाँधी राजकीय विद्यालय आसींद कमरा नंबर 1",
    list_title: str = "बड़े-घर की लिस्ट",
    min_family_size: int = 2,
    ignore_zero_houses: bool = True,
    selected_ward: Optional[str] = None,
    filter_active_only: bool = True,
    jila_parishad: Optional[str] = None,
    panchayat_samiti: Optional[str] = None
) -> bytes:
    """
    Exports a styled Excel (.xlsx) file with families arranged in descending order, organized ward-wise.
    Includes 'Total Members :: X' header banners and styled borders.
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
    font_family_banner = Font(name="Nirmala UI", size=10, bold=True, color="0F172A")
    font_data = Font(name="Nirmala UI", size=10)
    font_data_bold = Font(name="Nirmala UI", size=10, bold=True)

    fill_title = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_sub = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_header = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    fill_family = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    thin_border = Side(style='thin', color='CBD5E1')
    thick_bottom = Side(style='medium', color='0F172A')
    border_cell = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    border_last_row = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thick_bottom)

    headers = [
        "क्रम सं.",
        "मकान सं",
        "मतदाता का नाम",
        "पिता / पति का नाम",
        "उम्र",
        "लिंग",
        "मतदाता पहचान पत्र"
    ]

    for s_idx, w_num in enumerate(sorted_wards):
        if s_idx == 0:
            ws = wb.active
            ws.title = f"वार्ड_{w_num}"
        else:
            ws = wb.create_sheet(title=f"वार्ड_{w_num}")

        w_voters = ward_map[w_num]
        families, stats = group_and_sort_families(
            voters=w_voters,
            min_family_size=min_family_size,
            ignore_zero_houses=ignore_zero_houses,
            selected_ward=None,
            filter_active_only=False
        )

        w_jp = next((str(v.get('jila_parishad')).strip() for v in w_voters if v.get('jila_parishad')), str(jila_parishad or '').strip())
        w_ps = next((str(v.get('panchayat_samiti')).strip() for v in w_voters if v.get('panchayat_samiti')), str(panchayat_samiti or '').strip())
        w_booth = next((str(v.get('booth_address')).strip() for v in w_voters if v.get('booth_address')), str(booth_address or '').strip())

        # Row 1: Title Banner
        ws.merge_cells('A1:G1')
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
        ws.row_dimensions[1].height = 28

        # Row 2: Subheader (Booth Address & Summary Stats)
        ws.merge_cells('A2:G2')
        c2 = ws['A2']
        c2.value = f"मतदान केंद्र: {w_booth}  |  कुल बड़े परिवार: {stats['total_families']}  |  कुल मतदाता: {stats['total_voters']}"
        c2.font = font_sub
        c2.fill = fill_sub
        c2.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 22

        # Row 3: Column Headers
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_cell
        ws.row_dimensions[3].height = 24

        current_r = 4
        for f in families:
            h_no = f['house']
            total_m = f['total_members']
            members = f['members']

            # Family Banner Row
            ws.merge_cells(start_row=current_r, start_column=1, end_row=current_r, end_column=7)
            b_cell = ws.cell(row=current_r, column=1)
            b_cell.value = f"मकान नं.: {h_no}   >>>   Total Members :: {total_m}"
            b_cell.font = font_family_banner
            b_cell.fill = fill_family
            b_cell.alignment = Alignment(horizontal="right", vertical="center")
            ws.row_dimensions[current_r].height = 20
            current_r += 1

            for idx, m in enumerate(members):
                is_last = (idx == len(members) - 1)
                row_border = border_last_row if is_last else border_cell

                ws.cell(row=current_r, column=1, value=m.get('serial', '')).alignment = Alignment(horizontal="center")
                ws.cell(row=current_r, column=2, value=h_no).alignment = Alignment(horizontal="center")
                ws.cell(row=current_r, column=3, value=m.get('name', '')).alignment = Alignment(horizontal="left")
                ws.cell(row=current_r, column=4, value=m.get('relative_name', '')).alignment = Alignment(horizontal="left")
                ws.cell(row=current_r, column=5, value=m.get('age', '')).alignment = Alignment(horizontal="center")
                ws.cell(row=current_r, column=6, value=normalize_gender(m.get('gender', ''))).alignment = Alignment(horizontal="center")
                ws.cell(row=current_r, column=7, value=m.get('epic', '')).alignment = Alignment(horizontal="left")

                for col_idx in range(1, 8):
                    c = ws.cell(row=current_r, column=col_idx)
                    c.font = font_data
                    c.border = row_border
                    if idx % 2 == 1:
                        c.fill = fill_zebra

                ws.row_dimensions[current_r].height = 19
                current_r += 1

        # Column Widths
        col_widths = {1: 10, 2: 12, 3: 26, 4: 26, 5: 8, 6: 10, 7: 20}
        for col_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

    output = BytesIO()
    wb.save(output)
    return output.getvalue()
