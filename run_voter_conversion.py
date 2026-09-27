import sys
import pandas as pd
from pdf_engine.voter_extractor import extract_voters, export_voters_to_excel

sys.stdout.reconfigure(encoding='utf-8')

pdf_path = 'uploads/b9707b19-9ff5-425e-bd9d-f0973deb1ea6_BAGMALI-Ward_No-001.pdf'
out_xlsx = 'exports/voter_list_proper.xlsx'

print("Starting Voter List Extraction...")
voters = extract_voters(pdf_path)
print(f"Extraction complete! Total voters extracted: {len(voters)}")

# Verify against official figures: 255 active voters
print(f"Target count: 255 voters. Extracted count: {len(voters)} voters.")

export_voters_to_excel(voters, out_xlsx, theme='navy')
print(f"Excel file successfully generated at: {out_xlsx}")

# Inspect first 10 rows and last 5 rows
df = pd.DataFrame(voters)
print("\n--- Columns in Excel Sheet ---")
print(list(df.columns))

print("\n--- First 5 Active Voters ---")
print(df.head(5)[['भाग संख्या', 'क्रम संख्या', 'नाम', 'पिता/पति का नाम', 'आयु', 'मोबाइल नो', 'वोटर ID', 'हाउस नंबर']])

print("\n--- Last 5 Active Voters (from Addition List) ---")
print(df.tail(5)[['भाग संख्या', 'क्रम संख्या', 'नाम', 'पिता/पति का नाम', 'आयु', 'मोबाइल नो', 'वोटर ID', 'हाउस नंबर']])

# Check that deleted serials are NOT present
deleted_test_serials = [1, 18, 19, 20, 23, 24, 25, 26, 27, 40, 180, 243]
extracted_serials = set(df['क्रम संख्या'])
found_deleted = [s for s in deleted_test_serials if s in extracted_serials]
print(f"\nChecking for Deleted Serials (should be empty): {found_deleted}")
assert len(found_deleted) == 0, f"Error: Deleted serials found in active list: {found_deleted}"
print("VERIFICATION PASSED: No deleted voters are included in the Excel sheet!")
