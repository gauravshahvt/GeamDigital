import os
import pymupdf

def create_bank_statement_pdf(filepath: str):
    """Creates a sample 2-page Bank Statement PDF with multi-row transaction tables."""
    doc = pymupdf.open()
    
    # Page 1
    page1 = doc.new_page(width=612, height=792)  # Standard Letter
    
    # Header
    page1.insert_text((50, 50), "GLOBAL TRUST BANK", fontsize=18, color=(0.1, 0.2, 0.5))
    page1.insert_text((50, 70), "Monthly Account Statement", fontsize=12, color=(0.3, 0.3, 0.3))
    page1.insert_text((400, 50), "Account: #9876-5432-10", fontsize=10, color=(0.2, 0.2, 0.2))
    page1.insert_text((400, 65), "Period: Oct 01 - Oct 31, 2026", fontsize=10, color=(0.2, 0.2, 0.2))
    page1.insert_text((400, 80), "Currency: USD ($)", fontsize=10, color=(0.2, 0.2, 0.2))
    
    # Customer Info
    page1.insert_text((50, 110), "Account Holder: Johnathan Miller", fontsize=10, color=(0.1, 0.1, 0.1))
    page1.insert_text((50, 125), "Address: 742 Evergreen Terrace, Springfield", fontsize=9, color=(0.3, 0.3, 0.3))
    page1.insert_text((50, 140), "Starting Balance: $14,250.00", fontsize=10, color=(0.1, 0.5, 0.2))
    
    # Draw Table
    headers = ["Date", "Description", "Reference ID", "Withdrawals", "Deposits", "Balance"]
    col_x = [50, 120, 270, 370, 450, 530, 570]
    
    y = 175
    # Header Row Box
    page1.draw_rect(pymupdf.Rect(50, y - 14, 570, y + 8), color=(0.1, 0.2, 0.5), fill=(0.9, 0.93, 0.98))
    for i, h in enumerate(headers):
        page1.insert_text((col_x[i] + 4, y), h, fontsize=9, color=(0.1, 0.2, 0.5))
        
    transactions = [
        ("2026-10-01", "Opening Balance", "REF-001", "", "", "$14,250.00"),
        ("2026-10-03", "Payroll Direct Deposit - Tech Corp", "REF-8921", "", "$5,200.00", "$19,450.00"),
        ("2026-10-05", "Apex Supermarket Groceries", "TXN-4412", "$145.20", "", "$19,304.80"),
        ("2026-10-07", "Electric Utility AutoPay", "UTIL-993", "$85.50", "", "$19,219.30"),
        ("2026-10-12", "Office Supplies Depot", "ORD-6712", "$210.00", "", "$19,009.30"),
        ("2026-10-15", "Client Consulting Retainer", "INV-1029", "", "$3,500.00", "$22,509.30"),
        ("2026-10-18", "Cloud Hosting Infrastructure", "AWS-3319", "$320.40", "", "$22,188.90"),
        ("2026-10-22", "Health Insurance Premium", "INS-5510", "$450.00", "", "$21,738.90"),
        ("2026-10-25", "Coffee Shop Meeting", "POS-782", "$14.75", "", "$21,724.15"),
        ("2026-10-28", "High-Yield Interest Earned", "INT-2026", "", "$42.15", "$21,766.30"),
        ("2026-10-31", "Apartment Rental Payment", "RENT-110", "$1,850.00", "", "$19,916.30"),
    ]
    
    y += 24
    for row in transactions:
        page1.draw_line(pymupdf.Point(50, y + 6), pymupdf.Point(570, y + 6), color=(0.85, 0.85, 0.85))
        for i, val in enumerate(row):
            page1.insert_text((col_x[i] + 4, y), val, fontsize=8.5, color=(0.15, 0.15, 0.15))
        y += 20
        
    doc.save(filepath)
    doc.close()
    print(f"Created: {filepath}")

def create_invoice_pdf(filepath: str):
    """Creates a sample Corporate Invoice PDF with Key-Values and Line Items."""
    doc = pymupdf.open()
    page = doc.new_page(width=612, height=792)
    
    # Header Banner
    page.draw_rect(pymupdf.Rect(40, 40, 572, 90), color=(0.12, 0.23, 0.54), fill=(0.12, 0.23, 0.54))
    page.insert_text((55, 72), "NEXUS INNOVATIONS INC.", fontsize=18, color=(1, 1, 1))
    page.insert_text((420, 72), "TAX INVOICE", fontsize=16, color=(1, 1, 1))
    
    # Metadata Block
    y = 120
    page.insert_text((50, y), "Vendor Details:", fontsize=10, color=(0.1, 0.2, 0.5))
    page.insert_text((50, y+16), "Nexus Innovations Inc.", fontsize=9, color=(0.2, 0.2, 0.2))
    page.insert_text((50, y+30), "100 Innovation Way, Suite 400", fontsize=9, color=(0.4, 0.4, 0.4))
    page.insert_text((50, y+44), "San Francisco, CA 94105", fontsize=9, color=(0.4, 0.4, 0.4))
    page.insert_text((50, y+58), "Email: billing@nexusinno.com", fontsize=9, color=(0.4, 0.4, 0.4))
    
    page.insert_text((350, y), "Invoice Details:", fontsize=10, color=(0.1, 0.2, 0.5))
    page.insert_text((350, y+16), "Invoice Number: INV-2026-8849", fontsize=9, color=(0.2, 0.2, 0.2))
    page.insert_text((350, y+30), "Invoice Date: 2026-10-15", fontsize=9, color=(0.2, 0.2, 0.2))
    page.insert_text((350, y+44), "Due Date: 2026-11-15", fontsize=9, color=(0.2, 0.2, 0.2))
    page.insert_text((350, y+58), "PO Number: PO-99412", fontsize=9, color=(0.2, 0.2, 0.2))
    
    y = 210
    page.insert_text((50, y), "Bill To:", fontsize=10, color=(0.1, 0.2, 0.5))
    page.insert_text((50, y+16), "Acme Global Enterprises", fontsize=9, color=(0.2, 0.2, 0.2))
    page.insert_text((50, y+30), "Attn: Accounts Payable", fontsize=9, color=(0.4, 0.4, 0.4))
    page.insert_text((50, y+44), "500 Corporate Blvd, New York, NY 10001", fontsize=9, color=(0.4, 0.4, 0.4))
    
    # Items Table
    y = 285
    headers = ["Item #", "Description", "Qty", "Unit Price", "Tax %", "Total Amount"]
    col_x = [50, 110, 310, 370, 440, 500, 560]
    
    page.draw_rect(pymupdf.Rect(50, y - 14, 560, y + 8), color=(0.12, 0.23, 0.54), fill=(0.92, 0.95, 0.99))
    for i, h in enumerate(headers):
        page.insert_text((col_x[i] + 4, y), h, fontsize=9, color=(0.12, 0.23, 0.54))
        
    line_items = [
        ("1", "Enterprise Cloud AI Platform License", "2", "$2,400.00", "8.5%", "$4,800.00"),
        ("2", "Custom Integration & API Setup", "15", "$150.00", "8.5%", "$2,250.00"),
        ("3", "Staff Training Workshop (4 Sessions)", "1", "$1,200.00", "0.0%", "$1,200.00"),
        ("4", "24/7 SLA Priority Technical Support", "1", "$850.00", "8.5%", "$850.00"),
        ("5", "Data Migration and Schema Mapping", "1", "$1,500.00", "8.5%", "$1,500.00"),
    ]
    
    y += 24
    for row in line_items:
        page.draw_line(pymupdf.Point(50, y + 6), pymupdf.Point(560, y + 6), color=(0.88, 0.88, 0.88))
        for i, val in enumerate(row):
            page.insert_text((col_x[i] + 4, y), val, fontsize=8.5, color=(0.2, 0.2, 0.2))
        y += 22
        
    # Summary Totals Box
    y += 20
    page.draw_rect(pymupdf.Rect(350, y, 560, y + 90), color=(0.8, 0.8, 0.8), fill=(0.97, 0.98, 1.0))
    page.insert_text((365, y + 20), "Subtotal:", fontsize=9.5, color=(0.3, 0.3, 0.3))
    page.insert_text((485, y + 20), "$10,600.00", fontsize=9.5, color=(0.1, 0.1, 0.1))
    
    page.insert_text((365, y + 42), "Tax / VAT (8.5%):", fontsize=9.5, color=(0.3, 0.3, 0.3))
    page.insert_text((485, y + 42), "$799.00", fontsize=9.5, color=(0.1, 0.1, 0.1))
    
    page.draw_line(pymupdf.Point(360, y + 54), pymupdf.Point(550, y + 54), color=(0.7, 0.7, 0.7))
    
    page.insert_text((365, y + 74), "Total Amount:", fontsize=11, color=(0.12, 0.23, 0.54))
    page.insert_text((475, y + 74), "$11,399.00", fontsize=12, color=(0.12, 0.23, 0.54))
    
    doc.save(filepath)
    doc.close()
    print(f"Created: {filepath}")

def create_product_catalog_pdf(filepath: str):
    """Creates a sample multi-column Product Inventory table."""
    doc = pymupdf.open()
    page = doc.new_page(width=612, height=792)
    
    page.insert_text((50, 45), "Q3 PRODUCT INVENTORY & SALES REPORT", fontsize=15, color=(0.02, 0.37, 0.27))
    page.insert_text((50, 62), "Department of Commercial Logistics", fontsize=9, color=(0.4, 0.4, 0.4))
    
    headers = ["SKU Code", "Product Name", "Category", "Stock Qty", "Unit Cost", "Retail Price", "Margin %"]
    col_x = [45, 115, 235, 315, 380, 445, 515, 565]
    
    y = 95
    page.draw_rect(pymupdf.Rect(45, y - 12, 565, y + 8), color=(0.02, 0.37, 0.27), fill=(0.9, 0.96, 0.93))
    for i, h in enumerate(headers):
        page.insert_text((col_x[i] + 3, y), h, fontsize=8.5, color=(0.02, 0.37, 0.27))
        
    products = [
        ("PRD-101", "Ultra Wireless Mouse M3", "Peripherals", "450", "$14.50", "$29.99", "51.6%"),
        ("PRD-102", "Mechanical RGB Keyboard K9", "Peripherals", "180", "$38.00", "$79.99", "52.5%"),
        ("PRD-103", "4K HDR IPS Monitor 27-inch", "Displays", "75", "$165.00", "$299.99", "45.0%"),
        ("PRD-104", "USB-C Multiport Dock 10-in-1", "Accessories", "320", "$22.00", "$49.99", "56.0%"),
        ("PRD-105", "Noise Cancelling Headset Pro", "Audio", "140", "$45.00", "$99.99", "55.0%"),
        ("PRD-106", "Ergonomic Mesh Task Chair", "Furniture", "60", "$95.00", "$189.99", "50.0%"),
        ("PRD-107", "Dual Monitor Gas Spring Arm", "Accessories", "210", "$28.00", "$59.99", "53.3%"),
        ("PRD-108", "Portable SSD Drive 1TB NVMe", "Storage", "190", "$48.00", "$89.99", "46.7%"),
    ]
    
    y += 22
    for p in products:
        page.draw_line(pymupdf.Point(45, y + 5), pymupdf.Point(565, y + 5), color=(0.88, 0.88, 0.88))
        for i, val in enumerate(p):
            page.insert_text((col_x[i] + 3, y), val, fontsize=8, color=(0.15, 0.15, 0.15))
        y += 20
        
    doc.save(filepath)
    doc.close()
    print(f"Created: {filepath}")

if __name__ == '__main__':
    out_dir = os.path.join(os.path.dirname(__file__), "static", "samples")
    os.makedirs(out_dir, exist_ok=True)
    create_bank_statement_pdf(os.path.join(out_dir, "sample_bank_statement.pdf"))
    create_invoice_pdf(os.path.join(out_dir, "sample_invoice.pdf"))
    create_product_catalog_pdf(os.path.join(out_dir, "sample_inventory_report.pdf"))
