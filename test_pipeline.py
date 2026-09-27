import os
from pdf_engine.extractor import process_pdf
from pdf_engine.excel_builder import create_excel_workbook

samples = ['sample_bank_statement.pdf', 'sample_invoice.pdf', 'sample_inventory_report.pdf']
os.makedirs('exports', exist_ok=True)

for s in samples:
    path = os.path.join('static', 'samples', s)
    res = process_pdf(path)
    print(f"=== {s} ===")
    print(f"Tables found: {res['tables_count']}")
    for idx, t in enumerate(res['tables']):
        rows = len(t['data'])
        cols = len(t['data'][0]) if rows > 0 else 0
        print(f"  Table {idx+1}: {rows} rows x {cols} cols (Method: {t['method']})")
        print(f"  Header: {t['data'][0]}")
    
    # Test building Excel
    table_grids = [t['data'] for t in res['tables']]
    table_names = [f"Table_{i+1}" for i in range(len(table_grids))]
    excel_bytes = create_excel_workbook(table_grids, table_names=table_names, metadata=res['invoice_metadata'])
    out_xlsx = os.path.join('exports', s.replace('.pdf', '.xlsx'))
    with open(out_xlsx, 'wb') as f:
        f.write(excel_bytes)
    print(f"  Saved Excel: {out_xlsx} ({len(excel_bytes)} bytes)\n")

print("All tests passed successfully!")
