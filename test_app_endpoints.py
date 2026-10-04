from app import app
import json

client = app.test_client()

print("1. Testing GET /")
res = client.get('/')
assert res.status_code == 200, f"Expected 200, got {res.status_code}"
assert b"Geam Digital" in res.data
print("   Passed!")

print("2. Testing GET /api/status")
res = client.get('/api/status')
assert res.status_code == 200
data = res.get_json()
print("   Status:", data)
assert data['status'] == 'online'
print("   Passed!")

print("3. Testing POST /api/load_sample (Bank Statement)")
res = client.post('/api/load_sample', json={'sample_name': 'sample_bank_statement.pdf', 'mode': 'auto'})
assert res.status_code == 200, f"Expected 200, got {res.status_code}: {res.data}"
data = res.get_json()
assert 'data' in data
assert data['data']['tables_count'] > 0
print(f"   Loaded sample successfully! Found {data['data']['tables_count']} tables.")
sample_tables = [t['data'] for t in data['data']['tables']]

print("4. Testing POST /api/export_excel")
export_res = client.post('/api/export_excel', json={
    'tables': sample_tables,
    'table_names': ['Summary', 'Transactions'],
    'theme': 'navy',
    'layout_mode': 'multi_sheet',
    'metadata': data['data']['invoice_metadata'],
    'file_name': 'test_bank_statement.xlsx'
})
assert export_res.status_code == 200
assert len(export_res.data) > 1000
print(f"   Generated Excel file: {len(export_res.data)} bytes!")

print("5. Testing POST /api/export_csv")
csv_res = client.post('/api/export_csv', json={
    'table': sample_tables[0],
    'file_name': 'test.csv'
})
assert csv_res.status_code == 200
assert len(csv_res.data) > 0
print(f"   Generated CSV file: {len(csv_res.data)} bytes!")

print("6. Testing POST /api/load_sample (Voter List Rajasthan - Ward 001)")
voter_res = client.post('/api/load_sample', json={
    'sample_name': 'sample_voter_list_rajasthan.pdf',
    'mode': 'voter_list'
})
assert voter_res.status_code == 200, f"Expected 200, got {voter_res.status_code}: {voter_res.data}"
voter_data = voter_res.get_json()
assert 'data' in voter_data
v_table = voter_data['data']['tables'][0]['data']
# Header row + 255 voter rows = 256 rows total
print(f"   Extracted Voter Grid: {len(v_table)} rows x {len(v_table[0])} columns!")
assert len(v_table) == 256, f"Expected 256 rows (header + 255 voters), got {len(v_table)}"
assert v_table[0] == ['वार्ड नं.', 'क्रम संख्या', 'नाम', 'पिता/पति का नाम', 'आयु', 'मोबाइल नो', 'वोटर ID', 'हाउस नंबर', 'एड्रेस', 'बूथ का पता', 'Status']
print("   Columns match requested format 100%!")

# Export voter excel
voter_export = client.post('/api/export_excel', json={
    'tables': [v_table],
    'table_names': ['Voter_List'],
    'theme': 'navy',
    'layout_mode': 'single_sheet',
    'file_name': 'voter_list_ward_001.xlsx'
})
assert voter_export.status_code == 200
assert len(voter_export.data) > 5000
print(f"   Voter List Excel file generated: {len(voter_export.data)} bytes!")

print("\nALL FLASK API ENDPOINTS AND VOTER LIST WORKFLOWS VERIFIED 100% SUCCESSFULLY!")
