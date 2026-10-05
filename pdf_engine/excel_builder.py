import re
from datetime import datetime
from io import BytesIO
from typing import List, Dict, Any, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

THEMES = {
    'geam_digital': {
        'name': 'Geam Digital Pro',
        'header_fill': '0F2942',       # Premium Deep Navy / Slate
        'header_text': 'FFFFFF',
        'sub_fill': 'E0F2FE',          # Light Sky Blue Accent
        'zebra_fill': 'F8FAFC',        # Off-white / Very faint slate
        'border_color': 'CBD5E1',      # Light slate border
    },
    'navy': {
        'name': 'Corporate Navy',
        'header_fill': '1E3A8A',       # Deep Blue
        'header_text': 'FFFFFF',
        'sub_fill': 'DBEAFE',          # Light Blue Accent
        'zebra_fill': 'F8FAFC',        # Off-white / Very faint slate
        'border_color': 'CBD5E1',      # Light slate border
    },
    'emerald': {
        'name': 'Emerald Finance',
        'header_fill': '065F46',       # Deep Emerald Green
        'header_text': 'FFFFFF',
        'sub_fill': 'D1FAE5',
        'zebra_fill': 'F9FAFB',
        'border_color': 'D1D5DB',
    },
    'charcoal': {
        'name': 'Modern Slate',
        'header_fill': '1E293B',       # Slate 800
        'header_text': 'FFFFFF',
        'sub_fill': 'E2E8F0',
        'zebra_fill': 'F8FAFC',
        'border_color': 'CBD5E1',
    },
    'classic': {
        'name': 'Classic Office',
        'header_fill': '2C3E50',
        'header_text': 'FFFFFF',
        'sub_fill': 'ECF0F1',
        'zebra_fill': 'FAFAFA',
        'border_color': 'BDC3C7',
    }
}

def clean_cell_value(val: Any) -> Any:
    """
    Cleans cell content and converts strings into typed Python values
    (float, int, datetime, or clean string).
    """
    if val is None:
        return ""
    if not isinstance(val, str):
        return val

    s = val.strip()
    if not s:
        return ""

    # Clean redundant whitespace and invisible characters
    s_clean = re.sub(r'[\r\n]+', ' ', s).strip()
    s_clean = re.sub(r'\s{2,}', ' ', s_clean)

    # 1. Check for Percentage: e.g. "12.5%", "10%"
    pct_match = re.fullmatch(r'([+-]?\d+(?:\.\d+)?)\s*%', s_clean)
    if pct_match:
        try:
            return float(pct_match.group(1)) / 100.0
        except ValueError:
            pass

    # 2. Check for Currency / Numbers: e.g. "$1,234.56", "-123.45", "(500.00)"
    num_str = s_clean
    # Strip common currency symbols
    num_str = re.sub(r'[$€£₹¥]', '', num_str).strip()
    
    # Handle accounting negative in parentheses: (123.45) -> -123.45
    is_negative = False
    if num_str.startswith('(') and num_str.endswith(')'):
        is_negative = True
        num_str = num_str[1:-1].strip()

    # Match integer or float format with commas: 1,234,567.89 or 1234.56
    if re.fullmatch(r'^-?[\d,]+(\.\d+)?$', num_str):
        clean_num = num_str.replace(',', '')
        try:
            if '.' in clean_num:
                val_float = float(clean_num)
                return -val_float if is_negative else val_float
            else:
                val_int = int(clean_num)
                return -val_int if is_negative else val_int
        except ValueError:
            pass

    # 3. Check for standard Dates: YYYY-MM-DD or DD/MM/YYYY or MM/DD/YYYY
    date_formats = [
        '%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%m/%d/%Y',
        '%Y/%m/%d', '%b %d, %Y', '%d %b %Y', '%B %d, %Y'
    ]
    for fmt in date_formats:
        try:
            dt = datetime.strptime(s_clean, fmt)
            return dt.date()
        except ValueError:
            pass

    return s_clean

def format_excel_sheet(ws, theme_name: str = 'navy', is_invoice: bool = False):
    """
    Applies professional styling, auto-filter, freeze panes, borders,
    alignment, number formatting, and auto column widths to an openpyxl worksheet.
    """
    theme = THEMES.get(theme_name, THEMES['navy'])
    
    font_family = "Segoe UI"
    header_font = Font(name=font_family, size=11, bold=True, color=theme['header_text'])
    header_fill = PatternFill(start_color=theme['header_fill'], end_color=theme['header_fill'], fill_type="solid")
    
    regular_font = Font(name=font_family, size=10, color="1E293B")
    bold_font = Font(name=font_family, size=10, bold=True, color="1E293B")
    
    zebra_fill = PatternFill(start_color=theme['zebra_fill'], end_color=theme['zebra_fill'], fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color=theme['border_color']),
        right=Side(style='thin', color=theme['border_color']),
        top=Side(style='thin', color=theme['border_color']),
        bottom=Side(style='thin', color=theme['border_color'])
    )

    max_row = ws.max_row
    max_col = ws.max_column

    if max_row < 1 or max_col < 1:
        return

    # Freeze top row
    ws.freeze_panes = "A2"

    # Style Header (Row 1)
    ws.row_dimensions[1].height = 26
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Enable AutoFilter on header row
    if max_col >= 1 and max_row >= 1:
        start_letter = get_column_letter(1)
        end_letter = get_column_letter(max_col)
        ws.auto_filter.ref = f"{start_letter}1:{end_letter}{max_row}"

    # Style Data Rows
    col_widths = {c: len(str(ws.cell(row=1, column=c).value or '')) for c in range(1, max_col + 1)}

    for row_idx in range(2, max_row + 1):
        ws.row_dimensions[row_idx].height = 22 if (row_idx == max_row and 'कुल' in str(ws.cell(row=row_idx, column=1).value or '')) else 20
        is_even = (row_idx % 2 == 0)
        is_total_row = (row_idx == max_row and 'कुल' in str(ws.cell(row=row_idx, column=1).value or ''))
        row_fill = PatternFill(start_color=theme['sub_fill'], end_color=theme['sub_fill'], fill_type="solid") if is_total_row else (zebra_fill if is_even else white_fill)
        row_font = bold_font if is_total_row else regular_font

        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val = cell.value
            cell.font = row_font
            cell.border = thin_border
            cell.fill = row_fill

            header_name = str(ws.cell(row=1, column=col_idx).value or '').strip()
            
            # Special alignment and conditional styling for voter list and summary columns
            if header_name == 'Status':
                cell.alignment = Alignment(horizontal="center", vertical="center")
                val_str = str(val or '').strip().lower()
                if val_str == 'deleted':
                    cell.font = Font(name=font_family, size=10, bold=True, color="991B1B")
                    cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
                elif val_str == 'active':
                    cell.font = Font(name=font_family, size=10, bold=True, color="065F46")
                    cell.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
            elif header_name in ['वार्ड नं.', 'वार्ड संख्या', 'वार्ड नं', 'भाग संख्या', 'क्रम संख्या', 'आयु', 'लिंग', 'मोबाइल नो', 'वोटर ID', 'हाउस नंबर', 'क्र.सं.', 'वार्ड / भाग संख्या', 'कुल मतदाता', 'सक्रिय (Active)', 'विलोपित (Deleted)', 'वैध सक्रिय मतदाता']:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                # If decimal, format with 2 decimals
                if isinstance(val, float):
                    cell.number_format = '#,##0.00'
                else:
                    cell.number_format = '#,##0'
            elif hasattr(val, 'strftime'):  # Date or Datetime
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = 'yyyy-mm-dd'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

            # Track column width
            val_len = len(str(val or ''))
            if val_len > col_widths.get(col_idx, 0):
                col_widths[col_idx] = val_len

    # Set column widths with safety margins
    for col_idx, width in col_widths.items():
        col_letter = get_column_letter(col_idx)
        # Add padding of 4 chars; minimum width 12, max 60
        adjusted_width = max(12, min(width + 4, 60))
        ws.column_dimensions[col_letter].width = adjusted_width

def create_excel_workbook(
    tables: List[List[List[Any]]],
    table_names: Optional[List[str]] = None,
    theme: str = 'geam_digital',
    mode: str = 'multi_sheet',
    metadata: Optional[Dict[str, Any]] = None
) -> bytes:
    """
    Takes extracted table grids and builds a polished Excel file (.xlsx) in memory.
    Supports:
    - 'multi_sheet': Each table is on a separate named sheet.
    - 'single_sheet': All tables combined into one sheet separated by clean headers.
    - 'metadata': Adds an Executive Summary sheet for invoice / form key-value fields.
    """
    wb = openpyxl.Workbook()
    # Brand properties
    wb.properties.creator = "Geam Digital"
    wb.properties.lastModifiedBy = "Geam Digital"
    wb.properties.title = "Geam Digital Electoral Voter List"
    wb.properties.company = "Geam Digital"

    # Remove default sheet
    wb.remove(wb.active)

    theme_info = THEMES.get(theme, THEMES['geam_digital'])

    # 1. If metadata exists (Invoice / Form mode), create Executive Summary Sheet
    if metadata and any(metadata.values()):
        ws_meta = wb.create_sheet(title="Document Summary")
        ws_meta.views.sheetView[0].showGridLines = True
        
        # Title Banner
        ws_meta.row_dimensions[1].height = 34
        ws_meta.merge_cells("A1:C1")
        title_cell = ws_meta["A1"]
        title_cell.value = "DOCUMENT METADATA & SUMMARY"
        title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color=theme_info['header_fill'], end_color=theme_info['header_fill'], fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws_meta.row_dimensions[3].height = 24
        ws_meta["A3"] = "Field Name"
        ws_meta["B3"] = "Extracted Value"
        for c in ["A3", "B3"]:
            ws_meta[c].font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            ws_meta[c].fill = PatternFill(start_color=theme_info['header_fill'], end_color=theme_info['header_fill'], fill_type="solid")
            ws_meta[c].alignment = Alignment(horizontal="left", vertical="center")

        curr_row = 4
        thin_border = Border(
            left=Side(style='thin', color=theme_info['border_color']),
            right=Side(style='thin', color=theme_info['border_color']),
            top=Side(style='thin', color=theme_info['border_color']),
            bottom=Side(style='thin', color=theme_info['border_color'])
        )

        for k, v in metadata.items():
            if v:
                ws_meta.row_dimensions[curr_row].height = 20
                c1 = ws_meta.cell(row=curr_row, column=1, value=k)
                c2 = ws_meta.cell(row=curr_row, column=2, value=clean_cell_value(v))
                
                c1.font = Font(name="Segoe UI", size=10, bold=True, color="334155")
                c2.font = Font(name="Segoe UI", size=10, color="0F172A")
                c1.border = thin_border
                c2.border = thin_border
                
                # Highlight total amount if present
                if 'total' in k.lower():
                    c1.fill = PatternFill(start_color=theme_info['sub_fill'], end_color=theme_info['sub_fill'], fill_type="solid")
                    c2.fill = PatternFill(start_color=theme_info['sub_fill'], end_color=theme_info['sub_fill'], fill_type="solid")
                    c2.font = Font(name="Segoe UI", size=11, bold=True, color="0F172A")
                
                curr_row += 1

        ws_meta.column_dimensions['A'].width = 28
        ws_meta.column_dimensions['B'].width = 38
        ws_meta.column_dimensions['C'].width = 15

    # 2. Add Tables
    if not tables:
        # If no tables were detected, create a blank informative sheet
        ws = wb.create_sheet(title="Extracted Content")
        ws.cell(row=1, column=1, value="No tabular data detected in PDF.")
    elif mode == 'single_sheet':
        ws = wb.create_sheet(title="All Tables Combined")
        ws.views.sheetView[0].showGridLines = True
        current_row = 1

        for idx, table in enumerate(tables):
            if not table:
                continue
            table_name = table_names[idx] if table_names and idx < len(table_names) else f"Table {idx + 1}"
            
            # Write Section Header if multiple tables
            if len(tables) > 1:
                ws.row_dimensions[current_row].height = 26
                header_cell = ws.cell(row=current_row, column=1, value=f"### {table_name}")
                header_cell.font = Font(name="Segoe UI", size=11, bold=True, color=theme_info['header_text'])
                header_cell.fill = PatternFill(start_color=theme_info['header_fill'], end_color=theme_info['header_fill'], fill_type="solid")
                current_row += 1

            for row_data in table:
                ws.row_dimensions[current_row].height = 20
                for col_idx, val in enumerate(row_data, start=1):
                    cleaned_val = clean_cell_value(val)
                    cell = ws.cell(row=current_row, column=col_idx, value=cleaned_val)
                current_row += 1
            current_row += 2  # Blank spacing between tables

        format_excel_sheet(ws, theme_name=theme)
    else:
        # Multi-sheet: each table in its own worksheet
        for idx, table in enumerate(tables):
            if not table:
                continue
            name = table_names[idx] if table_names and idx < len(table_names) else f"Table_{idx + 1}"
            # Excel sheet names max 31 chars and no invalid chars: \ / ? * : [ ]
            clean_name = re.sub(r'[\/\\?*:[\]]', '_', name)[:31].strip() or f"Sheet_{idx+1}"
            
            # Avoid name collision
            count = 1
            base_name = clean_name
            while clean_name in wb.sheetnames:
                count += 1
                clean_name = f"{base_name[:27]}_{count}"

            ws = wb.create_sheet(title=clean_name)
            ws.views.sheetView[0].showGridLines = True

            for row_idx, row_data in enumerate(table, start=1):
                for col_idx, val in enumerate(row_data, start=1):
                    cleaned_val = clean_cell_value(val)
                    ws.cell(row=row_idx, column=col_idx, value=cleaned_val)

            format_excel_sheet(ws, theme_name=theme)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()
