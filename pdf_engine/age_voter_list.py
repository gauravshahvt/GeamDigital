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

def normalize_gender(val: Any) -> str:
    s = str(val or '').strip().lower()
    if not s:
        return 'पुरूष'
    if any(m in s for m in ['m', 'male', 'purush', 'पु', 'पुरूष', 'पचरष']):
        return 'पुरूष'
    if any(f in s for f in ['f', 'female', 'mahila', 'stree', 'स्त्री', 'महिला', 'म.', 'सल']):
        return 'स्त्री'
    return 'पुरूष'

def parse_age_num(val: Any) -> Optional[int]:
    """
    Extracts an integer age value safely from integer, float, string, or Hindi formats.
    Returns None if age is unparseable or empty.
    """
    if val is None or val == '':
        return None
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        m = re.search(r'\d+', str(val))
        if m:
            return int(m.group(0))
        return None

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

def filter_and_sort_by_age(
    voters: List[Dict[str, Any]],
    list_type: str = "young",  # "young" (18-26 asc), "senior" (120-70 desc), or "custom"
    min_age: Optional[int] = None,
    max_age: Optional[int] = None,
    sort_order: Optional[str] = None,  # "asc" or "desc"
    filter_active_only: bool = True,
    selected_ward: Optional[str] = None
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Filters voters by age range and sorts them accordingly.
    (A) Young: 18 - 26 बढ़ते क्रम में (Ascending: 18 -> 26)
    (B) Bujurg / Senior: 120 - 70 घटते क्रम में (Descending: 120 -> 70)
    """
    # Set default ranges based on list_type
    if list_type == "young":
        eff_min = 18 if min_age is None else min_age
        eff_max = 26 if max_age is None else max_age
        eff_sort = "asc" if sort_order is None else sort_order
        default_title = "युवा मतदाता सूची (आयु 18 से 26 वर्ष)"
    elif list_type == "senior":
        eff_min = 70 if min_age is None else min_age
        eff_max = 120 if max_age is None else max_age
        eff_sort = "desc" if sort_order is None else sort_order
        default_title = "वरिष्ठ / बुजुर्ग मतदाता सूची (आयु 120 से 70 वर्ष)"
    else:  # custom
        eff_min = 18 if min_age is None else min_age
        eff_max = 120 if max_age is None else max_age
        eff_sort = "asc" if sort_order is None else sort_order
        default_title = f"मतदाता सूची (आयु {eff_min} से {eff_max} वर्ष)"

    # Filter active only if enabled
    filtered = voters
    if filter_active_only:
        filtered = [v for v in filtered if v.get('status') != 'Deleted']

    # Filter by selected ward if specific ward
    if selected_ward and str(selected_ward).strip() not in ['all', 'all_wardwise', '', 'समस्त', 'सभी']:
        target_ward = extract_ward_num(selected_ward)
        filtered = [v for v in filtered if extract_ward_num(v.get('part_no', v.get('ward', '1'))) == target_ward]

    # Filter by Age Range
    valid_voters = []
    for v in filtered:
        age_int = parse_age_num(v.get('age'))
        if age_int is not None and eff_min <= age_int <= eff_max:
            v_copy = dict(v)
            v_copy['_parsed_age'] = age_int
            valid_voters.append(v_copy)

    # Sort
    is_desc = (eff_sort == 'desc')

    def get_serial_key(v):
        s = v.get('serial', 0)
        try:
            return int(float(str(s).strip()))
        except:
            return 999999

    valid_voters.sort(key=lambda v: (
        -(v['_parsed_age']) if is_desc else v['_parsed_age'],
        extract_ward_key(v.get('part_no', v.get('ward', '1'))),
        get_serial_key(v)
    ))

    meta = {
        'list_type': list_type,
        'min_age': eff_min,
        'max_age': eff_max,
        'sort_order': eff_sort,
        'default_title': default_title,
        'total_matched': len(valid_voters)
    }

    return valid_voters, meta

# ==============================================================
# HTML GENERATOR FOR FEATURE 6: AGE-WISE VOTER LIST
# ==============================================================

def generate_age_html(
    voters: List[Dict[str, Any]],
    ward_title: str = "ग्राम पंचायत",
    booth_address: str = "राजकीय उच्च माध्यमिक विद्यालय",
    list_title: Optional[str] = None,
    list_type: str = "young",  # 'young' or 'senior' or 'custom'
    min_age: Optional[int] = None,
    max_age: Optional[int] = None,
    sort_order: Optional[str] = None,
    filter_active_only: bool = True,
    selected_ward: Optional[str] = "all",
    for_preview: bool = False,
    preview_page: int = 1,
    rows_per_page: int = 31,
    jila_parishad: Optional[str] = None,
    panchayat_samiti: Optional[str] = None
) -> Tuple[str, int, int, Dict[str, Any]]:
    """
    Renders high-quality A4 printable HTML for Feature 6: Age-Wise Voter List.
    Header:
      - Left: Panchayat Name (e.g. ग्राम पंचायत)
      - Center: List Title (युवा मतदाता सूची 18-26 बढ़ते क्रम में / बुजुर्ग मतदाता सूची 120-70 घटते क्रम में)
      - Right: जि. प. : [संख्या] | पं. स. : [संख्या] (and Ward No if specific)
      - Subheader: Booth Address
    Row Columns (7 columns exactly):
      1. वार्ड संख्या
      2. क्रम संख्या
      3. मतदाता का नाम
      4. पिता / पति का नाम
      5. आयु
      6. लिंग
      7. मतदाता पहचान पत्र क्रमांक
    """
    filtered_voters, meta = filter_and_sort_by_age(
        voters=voters,
        list_type=list_type,
        min_age=min_age,
        max_age=max_age,
        sort_order=sort_order,
        filter_active_only=filter_active_only,
        selected_ward=selected_ward
    )

    if not list_title or list_title.strip() == '':
        list_title = meta['default_title']

    is_desc = (meta['sort_order'] == 'desc')

    # Determine grouping strategy:
    # 1. 'all_wardwise': group ward-by-ward (Ward 1 page 1-N, Ward 2 page 1-M...)
    # 2. specific ward (e.g. '1', '2'): single ward
    # 3. 'all': global panchayat list sorted purely by age (18->26 or 120->70) with ward column in each row
    all_pages = []

    # Helper for serial sorting
    def get_serial_key(v):
        s = v.get('serial', 0)
        try:
            return int(float(str(s).strip()))
        except:
            return 999999

    if selected_ward == 'all_wardwise':
        # Ward-by-ward grouping
        ward_map: Dict[str, List[Dict[str, Any]]] = {}
        for v in filtered_voters:
            w_num = extract_ward_num(v.get('part_no', v.get('ward', '1')))
            if w_num not in ward_map:
                ward_map[w_num] = []
            ward_map[w_num].append(v)

        sorted_wards = sorted(ward_map.keys(), key=extract_ward_key)

        for w_num in sorted_wards:
            w_voters = ward_map[w_num]
            w_voters.sort(key=lambda v: (-(v['_parsed_age']) if is_desc else v['_parsed_age'], get_serial_key(v)))

            w_jp = next((str(v.get('jila_parishad')).strip() for v in w_voters if v.get('jila_parishad')), str(jila_parishad or '').strip())
            w_ps = next((str(v.get('panchayat_samiti')).strip() for v in w_voters if v.get('panchayat_samiti')), str(panchayat_samiti or '').strip())
            w_booth = next((str(v.get('booth_address')).strip() for v in w_voters if v.get('booth_address')), str(booth_address or '').strip())

            right_parts = []
            if w_jp:
                right_parts.append(f"जि. प. : <b>{w_jp}</b>")
            if w_ps:
                right_parts.append(f"पं. स. : <b>{w_ps}</b>")
            right_parts.append(f"<b>वार्ड संख्या : {w_num}</b>")
            w_right_header_html = " &nbsp;|&nbsp; ".join(right_parts)

            w_total_pages = max(1, (len(w_voters) + rows_per_page - 1) // rows_per_page)
            chunks = [w_voters[i * rows_per_page : (i + 1) * rows_per_page] for i in range(w_total_pages)]

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

    elif selected_ward and selected_ward not in ['all', '', 'समस्त', 'सभी']:
        # Single specific ward
        w_num = extract_ward_num(selected_ward)
        filtered_voters.sort(key=lambda v: (-(v['_parsed_age']) if is_desc else v['_parsed_age'], get_serial_key(v)))

        w_jp = next((str(v.get('jila_parishad')).strip() for v in filtered_voters if v.get('jila_parishad')), str(jila_parishad or '').strip())
        w_ps = next((str(v.get('panchayat_samiti')).strip() for v in filtered_voters if v.get('panchayat_samiti')), str(panchayat_samiti or '').strip())
        w_booth = next((str(v.get('booth_address')).strip() for v in filtered_voters if v.get('booth_address')), str(booth_address or '').strip())

        right_parts = []
        if w_jp:
            right_parts.append(f"जि. प. : <b>{w_jp}</b>")
        if w_ps:
            right_parts.append(f"पं. स. : <b>{w_ps}</b>")
        right_parts.append(f"<b>वार्ड संख्या : {w_num}</b>")
        w_right_header_html = " &nbsp;|&nbsp; ".join(right_parts)

        w_total_pages = max(1, (len(filtered_voters) + rows_per_page - 1) // rows_per_page)
        chunks = [filtered_voters[i * rows_per_page : (i + 1) * rows_per_page] for i in range(w_total_pages)]

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
                'ward_voters': len(filtered_voters)
            })

    else:
        # 'all': Global panchayat list sorted primarily by age (18->26 or 120->70), then by ward, then serial
        filtered_voters.sort(key=lambda v: (
            -(v['_parsed_age']) if is_desc else v['_parsed_age'],
            extract_ward_key(v.get('part_no', v.get('ward', '1'))),
            get_serial_key(v)
        ))

        w_jp = next((str(v.get('jila_parishad')).strip() for v in filtered_voters if v.get('jila_parishad')), str(jila_parishad or '').strip())
        w_ps = next((str(v.get('panchayat_samiti')).strip() for v in filtered_voters if v.get('panchayat_samiti')), str(panchayat_samiti or '').strip())
        w_booth = next((str(v.get('booth_address')).strip() for v in filtered_voters if v.get('booth_address')), str(booth_address or '').strip())

        right_parts = []
        if w_jp:
            right_parts.append(f"जि. प. : <b>{w_jp}</b>")
        if w_ps:
            right_parts.append(f"पं. स. : <b>{w_ps}</b>")
        w_right_header_html = " &nbsp;|&nbsp; ".join(right_parts) if right_parts else "समस्त वार्ड"

        total_p = max(1, (len(filtered_voters) + rows_per_page - 1) // rows_per_page)
        chunks = [filtered_voters[i * rows_per_page : (i + 1) * rows_per_page] for i in range(total_p)]

        for p_idx, chunk in enumerate(chunks, start=1):
            all_pages.append({
                'ward_num': 'समस्त',
                'ward_title': ward_title,
                'list_title': list_title,
                'right_header_html': w_right_header_html,
                'booth_address': w_booth,
                'chunk': chunk,
                'page_in_ward': p_idx,
                'total_in_ward': total_p,
                'ward_voters': len(filtered_voters)
            })

    if not all_pages:
        right_parts = []
        if jila_parishad:
            right_parts.append(f"जि. प. : <b>{jila_parishad}</b>")
        if panchayat_samiti:
            right_parts.append(f"पं. स. : <b>{panchayat_samiti}</b>")
        all_pages.append({
            'ward_num': '1',
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
    total_matched = len(filtered_voters)

    if for_preview:
        p_idx = max(1, min(preview_page, total_pages))
        render_pages = [(p_idx, all_pages[p_idx - 1])]
    else:
        render_pages = list(enumerate(all_pages, start=1))

    # Age sorting badge text
    if list_type == 'young':
        sort_badge_text = "युवा मतदाता (आयु 18 - 26 वर्ष • बढ़ते क्रम में)"
        age_badge_class = "age-pill-young"
    elif list_type == 'senior':
        sort_badge_text = "वरिष्ठ / बुजुर्ग मतदाता (आयु 120 - 70 वर्ष • घटते क्रम में)"
        age_badge_class = "age-pill-senior"
    else:
        sort_badge_text = f"आयु {meta['min_age']} - {meta['max_age']} वर्ष ({'घटते क्रम' if is_desc else 'बढ़ते क्रम'})"
        age_badge_class = "age-pill-custom"

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
    .age-sheet-page {{
        margin: 0 auto;
        page-break-after: always;
        break-after: page;
        height: 281mm !important;
    }}
}}
.age-sheet-page {{
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
.age-header-container {{
    margin-bottom: 2mm;
    border-bottom: 1.2px solid #0f172a;
    padding-bottom: 1.5mm;
}}
.age-header-row {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    line-height: 1.2;
    margin-bottom: 1.5mm;
}}
.age-header-left {{
    font-size: 11.5pt;
    font-weight: 800;
    color: #0f172a;
    text-align: left;
    width: 30%;
}}
.age-header-center {{
    font-size: 13pt;
    font-weight: 900;
    color: #0f172a;
    text-align: center;
    width: 38%;
    letter-spacing: -0.2px;
}}
.age-header-right {{
    font-size: 10.5pt;
    font-weight: 800;
    color: #0f172a;
    text-align: right;
    width: 32%;
    white-space: nowrap;
}}
.age-subheader-row {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-size: 9.2pt;
    font-weight: 700;
    color: #334155;
    line-height: 1.25;
}}
.age-booth-text {{
    text-align: left;
    flex: 1;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}
.age-badge-text {{
    font-size: 8.5pt;
    font-weight: 800;
    padding: 1px 6px;
    border-radius: 4px;
    white-space: nowrap;
    border: 1px solid #0d9488;
    background-color: #f0fdfa;
    color: #0f766e;
}}
.age-table {{
    width: 100%;
    border-collapse: collapse;
    border: 1.2px solid #000;
    font-size: 8.5pt;
    line-height: 1.15;
    table-layout: fixed;
}}
.age-table thead th {{
    background-color: #e2e8f0;
    color: #0f172a;
    font-weight: 800;
    font-size: 8.2pt;
    padding: 2mm 1mm;
    border: 1px solid #000;
    text-align: center;
    white-space: nowrap;
}}
.age-table tbody td {{
    padding: 1.6mm 1.2mm;
    border: 1px solid #94a3b8;
    vertical-align: middle;
    color: #000;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}
.age-table tbody tr:nth-child(even) {{
    background-color: #f8fafc;
}}
.col-ward {{
    width: 9%;
    text-align: center;
    font-weight: 800;
}}
.col-serial {{
    width: 8%;
    text-align: center;
    font-weight: 800;
    font-family: Arial, sans-serif;
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
    width: 7.5%;
    text-align: center;
    font-weight: 900;
    font-size: 9pt;
    font-family: Arial, sans-serif;
    color: #0f172a;
}}
.col-gender {{
    width: 8%;
    text-align: center;
    font-weight: 700;
}}
.col-epic {{
    width: 19.5%;
    text-align: center;
    font-weight: 800;
    font-family: Arial, sans-serif;
    letter-spacing: 0.3px;
}}
.age-footer {{
    display: flex;
    justify-content: space-between;
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
<div class="age-sheet-page" data-page="{global_p_idx}">
    <div>
        <!-- Top Header -->
        <div class="age-header-container">
            <div class="age-header-row">
                <div class="age-header-left">{p_data['ward_title']}</div>
                <div class="age-header-center">{p_data['list_title']}</div>
                <div class="age-header-right">{p_data['right_header_html']}</div>
            </div>
            <div class="age-subheader-row">
                <div class="age-booth-text"><b>मतदान केंद्र / बूथ:</b> {p_data['booth_address']}</div>
                <div class="age-badge-text">{sort_badge_text}</div>
            </div>
        </div>

        <!-- 7-Column Data Table -->
        <table class="age-table">
            <thead>
                <tr>
                    <th class="col-ward">वार्ड संख्या</th>
                    <th class="col-serial">क्रम संख्या</th>
                    <th class="col-name">मतदाता का नाम</th>
                    <th class="col-rel">पिता / पति का नाम</th>
                    <th class="col-age">आयु</th>
                    <th class="col-gender">लिंग</th>
                    <th class="col-epic">मतदाता पहचान पत्र क्रमांक</th>
                </tr>
            </thead>
            <tbody>
"""
        for v in p_data['chunk']:
            w_val = extract_ward_num(v.get('part_no', v.get('ward', '1')))
            s_val = v.get('serial', '')
            name_val = v.get('name', '')
            rel_val = v.get('relative_name', '')
            age_val = v.get('_parsed_age', v.get('age', ''))
            gender_val = normalize_gender(v.get('gender', ''))
            epic_val = v.get('epic', '')

            html += f"""
                <tr>
                    <td class="col-ward">वार्ड {w_val}</td>
                    <td class="col-serial">{s_val}</td>
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

    <!-- Bottom Footer -->
    <div class="age-footer">
        <div>Geam Digital &bull; कुल मतदाता: <b>{total_matched}</b></div>
        <div>Page {p_data['page_in_ward']} of {p_data['total_in_ward']} (कुल पेज {total_pages})</div>
    </div>
</div>
"""

    html += """</body></html>"""
    cur_p = render_pages[0][1] if render_pages else all_pages[0]
    page_meta = {
        'current_ward': str(cur_p.get('ward_num', 'समस्त')),
        'page_in_ward': int(cur_p.get('page_in_ward', 1)),
        'total_in_ward': int(cur_p.get('total_in_ward', 1)),
        'ward_voters': int(cur_p.get('ward_voters', total_matched)),
        'total_matched': total_matched,
        'list_type': list_type,
        'min_age': meta['min_age'],
        'max_age': meta['max_age'],
        'sort_order': meta['sort_order']
    }
    return html, total_pages, total_matched, page_meta

# ==============================================================
# VECTOR PDF GENERATOR ENGINE FOR FEATURE 6
# ==============================================================

def generate_age_pdf(
    voters: List[Dict[str, Any]],
    ward_title: str = "ग्राम पंचायत",
    booth_address: str = "राजकीय उच्च माध्यमिक विद्यालय",
    list_title: Optional[str] = None,
    list_type: str = "young",
    min_age: Optional[int] = None,
    max_age: Optional[int] = None,
    sort_order: Optional[str] = None,
    filter_active_only: bool = True,
    selected_ward: Optional[str] = "all",
    jila_parishad: Optional[str] = None,
    panchayat_samiti: Optional[str] = None
) -> bytes:
    """
    Generates a production A4 PDF for Feature 6 using headless Chrome/Edge.
    """
    html_content, _, _, _ = generate_age_html(
        voters=voters,
        ward_title=ward_title,
        booth_address=booth_address,
        list_title=list_title,
        list_type=list_type,
        min_age=min_age,
        max_age=max_age,
        sort_order=sort_order,
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
        html_file = os.path.join(tmpdir, "age_voters.html")
        pdf_file = os.path.join(tmpdir, "age_voters.pdf")

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
# EXCEL EXPORT ENGINE FOR FEATURE 6
# ==============================================================

def generate_age_excel(
    voters: List[Dict[str, Any]],
    ward_title: str = "ग्राम पंचायत",
    booth_address: str = "राजकीय उच्च माध्यमिक विद्यालय",
    list_title: Optional[str] = None,
    list_type: str = "young",
    min_age: Optional[int] = None,
    max_age: Optional[int] = None,
    sort_order: Optional[str] = None,
    filter_active_only: bool = True,
    selected_ward: Optional[str] = "all",
    jila_parishad: Optional[str] = None,
    panchayat_samiti: Optional[str] = None
) -> bytes:
    """
    Exports a styled Excel (.xlsx) file with voters filtered by age and sorted.
    """
    filtered_voters, meta = filter_and_sort_by_age(
        voters=voters,
        list_type=list_type,
        min_age=min_age,
        max_age=max_age,
        sort_order=sort_order,
        filter_active_only=filter_active_only,
        selected_ward=selected_ward
    )

    if not list_title or list_title.strip() == '':
        list_title = meta['default_title']

    is_desc = (meta['sort_order'] == 'desc')

    def get_serial_key(v):
        s = v.get('serial', 0)
        try:
            return int(float(str(s).strip()))
        except:
            return 999999

    if selected_ward == 'all_wardwise':
        filtered_voters.sort(key=lambda v: (
            extract_ward_key(v.get('part_no', v.get('ward', '1'))),
            -(v['_parsed_age']) if is_desc else v['_parsed_age'],
            get_serial_key(v)
        ))
    else:
        filtered_voters.sort(key=lambda v: (
            -(v['_parsed_age']) if is_desc else v['_parsed_age'],
            extract_ward_key(v.get('part_no', v.get('ward', '1'))),
            get_serial_key(v)
        ))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "युवा_एवं_बुजुर्ग_मतदाता_सूची"
    ws.views.sheetView[0].showGridLines = True

    # Styles
    font_title = Font(name="Nirmala UI", size=13, bold=True, color="FFFFFF")
    font_sub = Font(name="Nirmala UI", size=10, bold=True, color="0F172A")
    font_header = Font(name="Nirmala UI", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Nirmala UI", size=10)
    font_data_bold = Font(name="Nirmala UI", size=10, bold=True)

    fill_title = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    fill_sub = PatternFill(start_color="F0FDFA", end_color="F0FDFA", fill_type="solid")
    fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    thin_border = Side(style='thin', color='CBD5E1')
    border_cell = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # Row 1: Title
    ws.merge_cells('A1:J1')
    c1 = ws['A1']
    c1.value = f"{ward_title} • {list_title}"
    c1.font = font_title
    c1.fill = fill_title
    c1.alignment = align_center
    ws.row_dimensions[1].height = 32

    # Row 2: Subheader (JP, PS, Booth)
    ws.merge_cells('A2:J2')
    c2 = ws['A2']
    c2.value = f"जि. प.: {jila_parishad or '-'}   |   पं. स.: {panchayat_samiti or '-'}   |   मतदान केंद्र: {booth_address or '-'}   |   कुल मतदाता: {len(filtered_voters)}"
    c2.font = font_sub
    c2.fill = fill_sub
    c2.alignment = align_center
    ws.row_dimensions[2].height = 24

    # Row 3: Headers
    headers = [
        "जि. प.", "पं. स.", "वार्ड संख्या", "क्रम संख्या",
        "मतदाता का नाम", "पिता / पति का नाम", "आयु", "लिंग",
        "मतदाता पहचान पत्र क्रमांक", "मतदान केंद्र का पता"
    ]
    ws.row_dimensions[3].height = 24
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_cell

    # Rows 4+: Data
    for row_idx, v in enumerate(filtered_voters, start=4):
        w_num = extract_ward_num(v.get('part_no', v.get('ward', '1')))
        jp_val = v.get('jila_parishad') or jila_parishad or ''
        ps_val = v.get('panchayat_samiti') or panchayat_samiti or ''
        booth_val = v.get('booth_address') or booth_address or ''

        row_data = [
            jp_val,
            ps_val,
            f"वार्ड {w_num}",
            v.get('serial', ''),
            v.get('name', ''),
            v.get('relative_name', ''),
            v.get('_parsed_age', v.get('age', '')),
            normalize_gender(v.get('gender', '')),
            v.get('epic', ''),
            booth_val
        ]

        ws.row_dimensions[row_idx].height = 20
        is_even = (row_idx % 2 == 0)

        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_data_bold if col_idx in [3, 4, 7, 9] else font_data
            if is_even:
                cell.fill = fill_zebra
            cell.border = border_cell
            cell.alignment = align_center if col_idx in [1, 2, 3, 4, 7, 8, 9] else align_left

    # Auto-adjust column widths
    col_widths = {
        1: 10,  # जि. प.
        2: 12,  # पं. स.
        3: 14,  # वार्ड
        4: 12,  # क्रम संख्या
        5: 26,  # नाम
        6: 26,  # पिता/पति
        7: 10,  # आयु
        8: 12,  # लिंग
        9: 20,  # EPIC
        10: 35  # Booth
    }
    for col_idx, width in col_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
