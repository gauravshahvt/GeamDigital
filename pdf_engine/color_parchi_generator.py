import os
import re
import subprocess
import tempfile
import base64
from io import BytesIO
from typing import List, Dict, Any, Optional, Tuple
from pdf_engine.parchi_generator import get_browser_executable

# ==========================================
# DEFAULT VECTOR ASSETS (SVGs as Data URLs)
# ==========================================

# 1. Hand Symbol (हाथ - तिरंगा पृष्ठभूमि)
HAND_SYMBOL_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 130" width="100" height="130">
  <rect x="2" y="2" width="96" height="42" fill="#FF9933" rx="4"/>
  <rect x="2" y="44" width="96" height="42" fill="#FFFFFF"/>
  <rect x="2" y="86" width="96" height="42" fill="#138808" rx="4"/>
  <rect x="2" y="2" width="96" height="126" fill="none" stroke="#333" stroke-width="2" rx="4"/>
  <g fill="#ffffff" stroke="#111" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round">
    <path d="M 32 85 C 32 95 38 108 50 110 C 62 108 72 95 72 82 L 72 45 C 72 41 66 41 66 45 L 66 30 C 66 26 60 26 60 30 L 60 25 C 60 21 54 21 54 25 L 54 32 C 54 28 48 28 48 32 L 48 65 L 43 55 C 39 48 32 52 35 60 L 40 72 L 32 85 Z"/>
    <path d="M 45 78 Q 55 85 65 76" fill="none" stroke="#222" stroke-width="1.5"/>
    <path d="M 42 88 Q 54 94 66 86" fill="none" stroke="#222" stroke-width="1.5"/>
  </g>
</svg>"""

# 2. Lotus Symbol (कमल)
LOTUS_SYMBOL_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 130" width="100" height="130">
  <rect x="2" y="2" width="96" height="126" fill="#f0fdf4" stroke="#333" stroke-width="2" rx="4"/>
  <g fill="#f43f5e" stroke="#881337" stroke-width="1.5">
    <path d="M 50 30 C 40 50 42 70 50 90 C 58 70 60 50 50 30 Z" fill="#fb7185"/>
    <path d="M 50 90 C 35 75 25 55 35 45 C 42 55 45 72 50 90 Z"/>
    <path d="M 50 90 C 65 75 75 55 65 45 C 58 55 55 72 50 90 Z"/>
    <path d="M 50 90 C 30 85 15 70 20 60 C 28 65 38 78 50 90 Z"/>
    <path d="M 50 90 C 70 85 85 70 80 60 C 72 65 62 78 50 90 Z"/>
  </g>
  <path d="M 25 96 Q 50 88 75 96 L 70 102 Q 50 94 30 102 Z" fill="#15803d"/>
  <path d="M 48 98 L 48 115 L 52 115 L 52 98 Z" fill="#166534"/>
</svg>"""

# 3. Cycle Symbol (साइकिल)
CYCLE_SYMBOL_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 130" width="100" height="130">
  <rect x="2" y="2" width="96" height="126" fill="#fef2f2" stroke="#333" stroke-width="2" rx="4"/>
  <g stroke="#dc2626" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round">
    <circle cx="28" cy="85" r="18"/>
    <circle cx="72" cy="85" r="18"/>
    <circle cx="28" cy="85" r="4" fill="#dc2626"/>
    <circle cx="72" cy="85" r="4" fill="#dc2626"/>
    <polygon points="28,85 50,85 62,60 40,60" fill="none"/>
    <line x1="50" y1="85" x2="43" y2="52"/>
    <line x1="72" y1="85" x2="62" y2="60"/>
    <line x1="62" y1="60" x2="60" y2="48"/>
    <line x1="52" y1="48" x2="68" y2="48"/>
    <line x1="43" y1="52" x2="35" y2="52"/>
  </g>
</svg>"""

# 4. Default Candidate Photo (Rajasthani Pagri & White Kurta with Namaste)
CANDIDATE_PHOTO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 150" width="120" height="150">
  <rect width="120" height="150" fill="#fdf2e9"/>
  <!-- Turban / Pagri (Rajasthani style multicolor) -->
  <path d="M 20 52 C 20 20 40 10 60 10 C 80 10 100 20 100 52 Z" fill="#dc2626"/>
  <path d="M 22 45 Q 60 25 98 45" fill="none" stroke="#fbbf24" stroke-width="6"/>
  <path d="M 25 35 Q 60 18 95 35" fill="none" stroke="#16a34a" stroke-width="5"/>
  <path d="M 28 27 Q 60 12 92 27" fill="none" stroke="#2563eb" stroke-width="4"/>
  <circle cx="60" cy="18" r="5" fill="#f59e0b"/>
  <!-- Face -->
  <ellipse cx="60" cy="65" rx="24" ry="26" fill="#fbcfe8"/>
  <ellipse cx="60" cy="65" rx="22" ry="24" fill="#fed7aa"/>
  <!-- Tilak -->
  <path d="M 59 48 L 61 48 L 60 56 Z" fill="#dc2626"/>
  <circle cx="60" cy="58" r="1.5" fill="#fbbf24"/>
  <!-- Eyes -->
  <ellipse cx="50" cy="62" rx="3" ry="1.8" fill="#1e293b"/>
  <ellipse cx="70" cy="62" rx="3" ry="1.8" fill="#1e293b"/>
  <!-- Moustache / Smile -->
  <path d="M 46 76 Q 60 84 74 76 Q 60 78 46 76" fill="#1e293b"/>
  <!-- White Kurta / Shirt with folded hands -->
  <path d="M 20 150 L 30 100 C 35 92 50 90 60 90 C 70 90 85 92 90 100 L 100 150 Z" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5"/>
  <!-- Namaste hands -->
  <path d="M 52 135 L 56 105 C 57 100 63 100 64 105 L 68 135 C 64 140 56 140 52 135 Z" fill="#fed7aa" stroke="#ea580c" stroke-width="1.2"/>
</svg>"""

def svg_to_data_url(svg_str: str) -> str:
    return "data:image/svg+xml;base64," + base64.b64encode(svg_str.strip().encode('utf-8')).decode('ascii')

DEFAULT_HAND_DATA = svg_to_data_url(HAND_SYMBOL_SVG)
DEFAULT_LOTUS_DATA = svg_to_data_url(LOTUS_SYMBOL_SVG)
DEFAULT_CYCLE_DATA = svg_to_data_url(CYCLE_SYMBOL_SVG)
DEFAULT_CANDIDATE_DATA = svg_to_data_url(CANDIDATE_PHOTO_SVG)

SYMBOL_PRESETS = {
    'hand': DEFAULT_HAND_DATA,
    'lotus': DEFAULT_LOTUS_DATA,
    'cycle': DEFAULT_CYCLE_DATA
}

# ==========================================
# COLOR PARCHI GENERATOR ENGINE
# ==========================================

def generate_color_parchi_html(
    voters: List[Dict[str, Any]],
    candidate_info: Optional[Dict[str, Any]] = None,
    layout_type: str = "8",
    show_page_number: bool = True,
    filter_active_only: bool = True,
    range_from: Optional[int] = None,
    range_to: Optional[int] = None,
    for_preview: bool = False,
    preview_page: int = 1
) -> Tuple[str, int, int]:
    """
    Renders high-quality HTML for Feature 3: Color Voter Slip with Candidate Poster.
    Layout Types:
      - '8': A4 - 8 (2 columns x 4 rows) [Default & Recommended as per reference]
      - '12': A4 - 12 (3 columns x 4 rows)
    """
    info = candidate_info or {}
    
    # Candidate info defaults
    candidate_name = info.get('candidate_name', 'संजय कुमार मेवाड़ा')
    header_prefix = info.get('header_prefix', 'पार्षद पद हेतु वार्ड नं')
    ward_badge = info.get('ward_badge', '03')
    header_suffix = info.get('header_suffix', 'से लोकप्रिय प्रत्याशी')
    button_no = info.get('button_no', '2')
    symbol_name = info.get('symbol_name', 'हाथ')
    slogan = info.get('slogan', 'को हाथ के निशान पर बटन दबा कर विजयी बनावें।')
    voting_time = info.get('voting_time', 'प्रातः 7 से सायं 6 बजे तक')
    time_color = info.get('time_color', '#e11d48')
    poster_bg_color = info.get('poster_bg_color', '#fff033')
    candidate_name_color = info.get('candidate_name_color', '#b91c1c')
    btn_box_color = info.get('btn_box_color', '#0284c7')
    
    # Symbols & Photo
    poster_img = info.get('poster_img_data', '')
    candidate_photo = info.get('candidate_photo_data', '') or DEFAULT_CANDIDATE_DATA
    symbol_key = info.get('symbol_preset', 'hand')
    symbol_img = info.get('custom_symbol_data', '') or info.get('symbol_img_data', '') or SYMBOL_PRESETS.get(symbol_key, DEFAULT_HAND_DATA)

    # Filter voters
    filtered = voters
    if filter_active_only:
        filtered = [v for v in voters if v.get('status') != 'Deleted']

    if range_from is not None or range_to is not None:
        r_from = range_from if range_from is not None else 1
        r_to = range_to if range_to is not None else 999999
        filtered = [v for v in filtered if r_from <= (v.get('serial') or 0) <= r_to]

    layout_type = str(layout_type).strip()
    if layout_type == "12":
        cols = 3
        rows = 4
        col_w = "66mm"
        cards_per_page = 12
        gap_w = "1.2mm"
        poster_h = "29mm"
        f_cand_name = "11pt"
        f_slogan = "5.5pt"
        f_ribbon = "7pt"
        f_data = "7.8pt"
        f_booth = "6.5pt"
    else:  # default "8" (2x4)
        cols = 2
        rows = 4
        col_w = "99.5mm"
        cards_per_page = 8
        gap_w = "1.5mm"
        poster_h = "31mm"
        f_cand_name = "13.5pt"
        f_slogan = "6.8pt"
        f_ribbon = "8.8pt"
        f_data = "8.5pt"
        f_booth = "7.2pt"

    card_height = "67.5mm" if show_page_number else "70.2mm"
    total_voters = len(filtered)
    total_pages = max(1, (total_voters + cards_per_page - 1) // cards_per_page)

    if for_preview:
        p_idx = max(1, min(preview_page, total_pages))
        page_chunks = [filtered[(p_idx - 1) * cards_per_page : p_idx * cards_per_page]]
        display_pages = [p_idx]
    else:
        page_chunks = [filtered[i * cards_per_page : (i + 1) * cards_per_page] for i in range(total_pages)]
        display_pages = list(range(1, total_pages + 1))

    html = f"""<!DOCTYPE html>
<html lang="hi">
<head>
<meta charset="utf-8">
<title>Geam Digital - Color Voter Slip with Candidate Poster</title>
<style>
@page {{
    size: A4 portrait;
    margin: 4mm;
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
    margin: 0;
    padding: 0;
}}
@media print {{
    body {{
        background: #fff;
        margin: 0;
        padding: 0;
    }}
    .sheet-page {{
        margin: 0 auto;
        page-break-after: always;
        break-after: page;
    }}
}}
.sheet-page {{
    width: 202mm;
    height: 288mm;
    display: grid;
    grid-template-columns: repeat({cols}, {col_w});
    grid-template-rows: repeat(4, 70.8mm);
    gap: {gap_w};
    margin: 0 auto;
    overflow: hidden;
    background: #fff;
    page-break-after: always;
    break-after: page;
}}
.parchi-cell {{
    display: flex;
    flex-direction: column;
    height: 70.8mm;
    justify-content: flex-start;
}}
.color-parchi-card {{
    border: 1.2px solid #000;
    height: {card_height};
    background: #fff;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    overflow: hidden;
    box-sizing: border-box;
}}
.parchi-page-num {{
    text-align: right;
    font-size: 7.5pt;
    font-family: Arial, sans-serif;
    color: #000;
    padding-right: 1.5mm;
    line-height: 1.2;
    margin-top: 0.4mm;
}}

/* Top Poster Area */
.poster-box {{
    background: {poster_bg_color};
    border-bottom: 1.2px solid #000;
    display: flex;
    flex-direction: column;
    padding: 1mm 1.5mm;
    height: {poster_h};
    justify-content: space-between;
    box-sizing: border-box;
}}
.poster-full-box {{
    height: {poster_h};
    border-bottom: 1.2px solid #000;
    overflow: hidden;
    display: flex;
    align-items: center;
    justify-content: center;
    background: #fff;
    margin: 0;
    padding: 0;
}}
.poster-full-img {{
    width: 100%;
    height: 100%;
    object-fit: fill;
    display: block;
}}
.poster-ribbon {{
    text-align: center;
    font-size: {f_ribbon};
    font-weight: 800;
    color: #003b73;
    letter-spacing: 0.2px;
    line-height: 1.15;
    padding-bottom: 1px;
}}
.ward-badge {{
    background: #15803d;
    color: #fff;
    padding: 0 4px;
    border-radius: 9999px;
    font-size: 7.5pt;
    font-weight: 800;
    display: inline-block;
}}
.poster-body {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 1.5mm;
    flex: 1;
}}
.symbol-col {{
    display: flex;
    flex-direction: column;
    align-items: center;
    width: 18mm;
    flex-shrink: 0;
}}
.symbol-img {{
    height: 16mm;
    width: 17mm;
    object-fit: contain;
}}
.symbol-lbl {{
    font-size: 5.5pt;
    font-weight: 800;
    color: #111;
    margin-top: 0.5px;
}}
.center-col {{
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    text-align: center;
    flex: 1;
}}
.candidate-name {{
    font-size: {f_cand_name};
    font-weight: 900;
    color: {candidate_name_color};
    letter-spacing: -0.2px;
    line-height: 1.05;
    text-shadow: 0.5px 0.5px 0 #fff;
}}
.btn-box {{
    background: {btn_box_color};
    color: #fff;
    font-size: 7.8pt;
    font-weight: 800;
    padding: 0 8px;
    border-radius: 9999px;
    margin: 1.5px 0;
    display: inline-flex;
    align-items: center;
    gap: 3px;
    box-shadow: 0 1px 2px rgba(0,0,0,0.15);
}}
.btn-num {{
    background: #fff;
    color: {btn_box_color};
    border-radius: 9999px;
    width: 14px;
    height: 14px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-size: 7.8pt;
    font-weight: 900;
}}
.slogan-txt {{
    font-size: {f_slogan};
    font-weight: 800;
    color: #92400e;
    line-height: 1.15;
    letter-spacing: -0.1px;
}}
.photo-col {{
    width: 19mm;
    height: 22mm;
    flex-shrink: 0;
    display: flex;
    align-items: center;
    justify-content: center;
}}
.candidate-photo {{
    width: 19mm;
    height: 22mm;
    object-fit: cover;
    border-radius: 3px;
}}

/* Middle Perforated Cut Line */
.cut-divider {{
    border-top: 1.2px dashed #111;
    border-bottom: 1.2px dashed #111;
    padding: 1px 2mm;
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #fff;
    line-height: 1;
}}
.scissor {{
    font-size: 9pt;
    color: #000;
    line-height: 1;
}}
.cut-label {{
    font-size: 5.6pt;
    font-weight: 800;
    color: #000;
    letter-spacing: -0.1px;
    text-align: center;
}}

/* Voter Details Area */
.voter-details-area {{
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    flex: 1;
    padding: 1.2mm 2.5mm 0 2.5mm;
    background: #fff;
}}
.v-row {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-size: {f_data};
    line-height: 1.2;
}}
.v-row-top {{
    font-size: 9.5pt;
}}
.v-col {{
    white-space: nowrap;
}}
.v-epic {{
    font-size: 8pt;
}}
.font-bold {{
    font-weight: 800;
}}
.font-mono {{
    font-family: Arial, monospace;
    letter-spacing: 0.2px;
}}
.serial-square {{
    border: 1.5px solid #000;
    padding: 0 5px;
    font-weight: 900;
    font-size: 9.5pt;
    display: inline-block;
    min-width: 18px;
    text-align: center;
}}
.v-row-booth {{
    font-size: {f_booth};
    line-height: 1.2;
    color: #111;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    padding-top: 0.5px;
}}
.voting-time-bar {{
    background: {time_color};
    color: #fff;
    text-align: center;
    font-size: 7.5pt;
    font-weight: 800;
    padding: 1.5px 0;
    margin: 1mm -2.5mm 0 -2.5mm;
    letter-spacing: 0.2px;
}}
.empty-card {{
    border: 1px dashed #ccc;
    background: #fafafa;
}}
</style>
</head>
<body>
"""

    for page_num, chunk in zip(display_pages, page_chunks):
        html += f"""<div class="sheet-page" data-page="{page_num}">\n"""
        
        for card_idx in range(cards_per_page):
            page_num_html = f"""<div class="parchi-page-num">{page_num}</div>""" if show_page_number else ""
            if card_idx < len(chunk):
                v = chunk[card_idx]
                serial = v.get('serial', '')
                part_no = v.get('part_no', '1')
                ward_val = v.get('ward', v.get('part_no', ward_badge))
                epic = v.get('epic', '')
                name = v.get('name', '')
                rel = v.get('relative_name', '')
                age = v.get('age', '')
                gender = v.get('gender', '')
                house = v.get('house', '1')
                booth = v.get('booth_address', '')
                jp = v.get('jila_parishad') or info.get('jila_parishad', '')
                ps = v.get('panchayat_samiti') or info.get('panchayat_samiti', '')
                gender_age_str = f"{gender} {age}".strip()

                if poster_img and str(poster_img).strip() not in ['', 'null', 'undefined', 'None']:
                    poster_html = f"""<div class="poster-full-box"><img src="{poster_img}" class="poster-full-img" alt="Candidate Poster"></div>"""
                else:
                    poster_html = f"""
                    <div class="poster-box">
                        <div class="poster-ribbon">
                            {header_prefix} <span class="ward-badge">{ward_badge}</span> {header_suffix}
                        </div>
                        <div class="poster-body">
                            <div class="symbol-col">
                                <img src="{symbol_img}" class="symbol-img">
                                <div class="symbol-lbl">चुनाव चिन्ह</div>
                            </div>
                            <div class="center-col">
                                <div class="candidate-name">{candidate_name}</div>
                                <div class="btn-box">
                                    बटन नं. <span class="btn-num">{button_no}</span>
                                </div>
                                <div class="slogan-txt">
                                    {slogan}
                                </div>
                            </div>
                            <div class="photo-col">
                                <img src="{candidate_photo}" class="candidate-photo">
                            </div>
                        </div>
                    </div>
                    """

                card_html = f"""
    <div class="parchi-cell">
        <div class="color-parchi-card">
            {poster_html}
            
            <div class="cut-divider">
                <span class="scissor">✂</span>
                <span class="cut-label">मतदान तिथि से पूर्व बांटी गयी पर्ची , मतदान केंद्र पर जाने से पहले ऊपर का हिस्सा फाड़ देवे</span>
                <span class="scissor">✂</span>
            </div>

            <div class="voter-details-area">
                <div class="v-row v-row-top">
                    <div class="v-col">
                        क्रम सं. : <span class="serial-square">{serial}</span>
                    </div>
                    {f'<div class="v-col" style="font-size: 7.8pt; font-weight: 700; color: #1e293b;">जि. प. : <b>{jp}</b> &nbsp;|&nbsp; पं. स. : <b>{ps}</b></div>' if (jp or ps) else ''}
                    <div class="v-col font-bold">
                        वार्ड नं.- {ward_val}
                    </div>
                </div>

                <div class="v-row">
                    <div class="v-col">
                        नाम : <span class="font-bold">{name}</span>
                    </div>
                    <div class="v-col">
                        भाग सं <span class="font-bold">{part_no}</span>
                    </div>
                    <div class="v-col v-epic">
                        Voter ID <span class="font-bold font-mono">{epic}</span>
                    </div>
                </div>

                <div class="v-row">
                    <div class="v-col">
                        पिता/पति : <span class="font-bold">{rel}</span>
                    </div>
                    <div class="v-col">
                        मकान सं : <span class="font-bold">{house}</span>
                    </div>
                    <div class="v-col font-bold">
                        {gender_age_str}
                    </div>
                </div>

                <div class="v-row-booth">
                    मतदान केंद्र - {booth}
                </div>

                <div class="voting-time-bar">
                    मतदान समय : {voting_time}
                </div>
            </div>
        </div>
        {page_num_html}
    </div>
"""
            else:
                card_html = f"""
    <div class="parchi-cell">
        <div class="color-parchi-card empty-card"></div>
        {page_num_html}
    </div>
"""
            html += card_html

        html += """</div>\n"""

    html += """</body></html>"""
    return html, total_pages, total_voters

def generate_color_parchi_pdf(
    voters: List[Dict[str, Any]],
    candidate_info: Optional[Dict[str, Any]] = None,
    layout_type: str = "8",
    show_page_number: bool = True,
    filter_active_only: bool = True,
    range_from: Optional[int] = None,
    range_to: Optional[int] = None
) -> bytes:
    """
    Generates a production-ready A4 PDF with Color Voter Slips per sheet using headless Chrome/Edge.
    """
    html_content, total_pages, total_voters = generate_color_parchi_html(
        voters=voters,
        candidate_info=candidate_info,
        layout_type=layout_type,
        show_page_number=show_page_number,
        filter_active_only=filter_active_only,
        range_from=range_from,
        range_to=range_to,
        for_preview=False
    )

    browser_bin = get_browser_executable()
    if not browser_bin:
        raise RuntimeError("No headless Chrome or Edge browser found on system to generate PDF.")

    with tempfile.TemporaryDirectory() as tmpdir:
        html_file = os.path.join(tmpdir, "color_parchi.html")
        pdf_file = os.path.join(tmpdir, "color_parchi.pdf")

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
