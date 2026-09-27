import re
import pandas as pd
from typing import Dict, Any, List

# Common Regex Patterns for Invoices & Receipts
INVOICE_PATTERNS = {
    'invoice_number': [
        r'(?i)(?:invoice|inv|bill|receipt|tax\s*invoice|reference)[\s#.:-]*([a-zA-Z0-9\-_/]{3,30})',
        r'(?i)(?:invoice\s*no|inv\s*no|bill\s*no)[\s.:]*([a-zA-Z0-9\-_/]+)',
    ],
    'invoice_date': [
        r'(?i)(?:invoice\s*date|bill\s*date|date\s*of\s*issue|issue\s*date|dated)[\s.:]*([0-9]{1,2}[-/.][0-9]{1,2}[-/.][0-9]{2,4})',
        r'(?i)(?:invoice\s*date|bill\s*date|date)[\s.:]*([a-zA-Z]+\s+\d{1,2},?\s+\d{4})',
        r'(?i)(?:invoice\s*date|bill\s*date|date)[\s.:]*(\d{1,2}\s+[a-zA-Z]+\s+\d{4})',
        r'(\b\d{4}[-/]\d{2}[-/]\d{2}\b)',
    ],
    'due_date': [
        r'(?i)(?:due\s*date|payment\s*due)[\s.:]*([0-9]{1,2}[-/.][0-9]{1,2}[-/.][0-9]{2,4})',
        r'(?i)(?:due\s*date|payment\s*due)[\s.:]*([a-zA-Z]+\s+\d{1,2},?\s+\d{4})',
    ],
    'po_number': [
        r'(?i)(?:p\.?o\.?\s*#?|purchase\s*order|order\s*no)[\s.:]*([a-zA-Z0-9\-_/]+)',
    ],
    'total_amount': [
        r'(?i)(?:grand\s*total|total\s*amount|balance\s*due|amount\s*due|net\s*payable|total)[\s.:$€£₹]*([\d,]+\.\d{2})',
        r'(?i)(?:total)[\s.:$€£₹]*([\d,]+(?:\.\d{2})?)',
    ],
    'subtotal': [
        r'(?i)(?:sub\s*total|subtotal|net\s*amount)[\s.:$€£₹]*([\d,]+\.\d{2})',
    ],
    'tax_amount': [
        r'(?i)(?:tax|vat|gst|sales\s*tax|cgst|sgst|igst)[\s.:$€£₹]*([\d,]+\.\d{2})',
    ],
    'currency': [
        r'(\$|€|£|₹|USD|EUR|GBP|INR|CAD|AUD)',
    ]
}

def extract_invoice_metadata(text: str) -> Dict[str, Any]:
    """
    Scans document text to extract standard key-value invoice fields.
    """
    metadata = {
        'Invoice Number': '',
        'Invoice Date': '',
        'Due Date': '',
        'PO Number': '',
        'Subtotal': '',
        'Tax / VAT': '',
        'Total Amount': '',
        'Currency': '$',
        'Vendor Name': '',
        'Customer / Bill To': '',
    }
    
    # Currency search
    curr_match = re.search(INVOICE_PATTERNS['currency'][0], text)
    if curr_match:
        metadata['Currency'] = curr_match.group(1)

    for field, patterns in INVOICE_PATTERNS.items():
        if field == 'currency':
            continue
        display_name = {
            'invoice_number': 'Invoice Number',
            'invoice_date': 'Invoice Date',
            'due_date': 'Due Date',
            'po_number': 'PO Number',
            'total_amount': 'Total Amount',
            'subtotal': 'Subtotal',
            'tax_amount': 'Tax / VAT'
        }.get(field, field)

        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                val = match.group(1).strip()
                if val:
                    metadata[display_name] = val
                    break

    # Look for "Bill To" or "Customer" block
    bill_to_match = re.search(r'(?i)(?:bill\s*to|invoiced\s*to|customer|client)[\s.:\n]+([^\n]+(?:\n[^\n]+){1,3})', text)
    if bill_to_match:
        raw_cust = bill_to_match.group(1).strip()
        # Take first clean line as customer name
        cust_lines = [l.strip() for l in raw_cust.split('\n') if l.strip() and not re.search(r'(?i)(invoice|date|phone|email)', l)]
        if cust_lines:
            metadata['Customer / Bill To'] = cust_lines[0]

    # Look for Vendor / Top company name
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if lines:
        for line in lines[:5]:
            if not re.search(r'(?i)(invoice|bill|tax|page|\d{3,})', line) and len(line) > 2:
                metadata['Vendor Name'] = line
                break

    return metadata

def metadata_to_dataframe(metadata: Dict[str, Any]) -> pd.DataFrame:
    """
    Convert metadata dictionary to a 2-column DataFrame suitable for Excel summary.
    """
    rows = []
    for k, v in metadata.items():
        if v:  # Only include populated fields or non-empty
            rows.append({'Property': k, 'Value': str(v)})
    return pd.DataFrame(rows)
